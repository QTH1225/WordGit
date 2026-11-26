using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Security.Authentication;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace WordGit
{
    internal enum UpdateState { Available, Current, LocalNewer, NoRelease, RateLimited, Unavailable, Unsupported }

    internal sealed class UpdateResult
    {
        internal readonly UpdateState State;
        internal readonly string Message, Version, Notes, DownloadUrl;
        internal UpdateResult(UpdateState state, string message, string version = null, string notes = null, string downloadUrl = null)
        {
            State = state;
            Message = message;
            Version = version;
            Notes = notes;
            DownloadUrl = downloadUrl;
        }
    }

    internal sealed class ReleaseVersion : IComparable<ReleaseVersion>
    {
        private static readonly Regex Pattern = new Regex(@"\Av?(0|[1-9][0-9]{0,8})\.(0|[1-9][0-9]{0,8})\.(0|[1-9][0-9]{0,8})(?:\.(0|[1-9][0-9]{0,8})|-beta\.([1-9][0-9]{0,8}))?\z");
        private readonly Version number;
        private readonly int beta;
        internal readonly string Text;
        internal bool IsBeta { get { return beta != 0; } }

        private ReleaseVersion(Match match)
        {
            number = new Version(Int32.Parse(match.Groups[1].Value), Int32.Parse(match.Groups[2].Value),
                Int32.Parse(match.Groups[3].Value), match.Groups[4].Success ? Int32.Parse(match.Groups[4].Value) : 0);
            beta = match.Groups[5].Success ? Int32.Parse(match.Groups[5].Value) : 0;
            Text = match.Value.TrimStart('v');
        }

        internal static ReleaseVersion Parse(string text)
        {
            if (text == null) return null;
            var match = Pattern.Match(text);
            return match.Success ? new ReleaseVersion(match) : null;
        }

        public int CompareTo(ReleaseVersion other)
        {
            var order = number.CompareTo(other.number);
            if (order != 0 || beta == other.beta) return order;
            if (!IsBeta) return 1;
            if (!other.IsBeta) return -1;
            return beta.CompareTo(other.beta);
        }
    }

    internal static class UpdateChecker
    {
        internal static async Task<UpdateResult> CheckAsync(string current, bool beta, CancellationToken cancellation)
        {
            using (var handler = new HttpClientHandler { AllowAutoRedirect = false, SslProtocols = SslProtocols.Tls12 })
            using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(12), MaxResponseContentBufferSize = 1024 * 1024 })
                return await CheckAsync(current, beta, client, cancellation).ConfigureAwait(false);
        }

        internal static async Task<UpdateResult> CheckAsync(string current, bool beta, HttpClient client, CancellationToken cancellation)
        {
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, AppInfo.ReleasesApi))
                {
                    request.Headers.UserAgent.ParseAdd("WordGit/" + AppInfo.DisplayVersion);
                    request.Headers.Accept.ParseAdd("application/vnd.github+json");
                    request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");
                    using (var response = await client.SendAsync(request, cancellation).ConfigureAwait(false))
                    {
                        if (response.StatusCode == HttpStatusCode.NotFound) return NoRelease();
                        if (response.StatusCode == HttpStatusCode.Forbidden || (int)response.StatusCode == 429)
                            return new UpdateResult(UpdateState.RateLimited, "GitHub 暂时限制了更新查询，请稍后重试。");
                        if (!response.IsSuccessStatusCode)
                            return new UpdateResult(UpdateState.Unavailable, "检查更新失败（HTTP " + (int)response.StatusCode + "）。");
                        return Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false), current, beta);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                if (cancellation.IsCancellationRequested) throw;
                return new UpdateResult(UpdateState.Unavailable, "检查更新超时，请稍后重试。");
            }
            catch (HttpRequestException)
            {
                return new UpdateResult(UpdateState.Unavailable, "无法连接 GitHub，请检查网络后重试。");
            }
        }

        internal static UpdateResult Parse(string json, string current, bool beta)
        {
            var installed = ReleaseVersion.Parse(current);
            if (installed == null) return new UpdateResult(UpdateState.Unsupported, "当前版本号无法识别。");
            try
            {
                var releases = new JavaScriptSerializer { MaxJsonLength = 1024 * 1024 }.DeserializeObject(json) as object[];
                if (releases == null) return InvalidResponse();
                ReleaseVersion latest = null;
                Dictionary<string, object> selected = null;
                foreach (var item in releases)
                {
                    var release = item as Dictionary<string, object>;
                    object draft, prerelease;
                    if (release == null || !release.TryGetValue("draft", out draft) || !(draft is bool) ||
                        !release.TryGetValue("prerelease", out prerelease) || !(prerelease is bool)) return InvalidResponse();
                    var version = ReleaseVersion.Parse(Field(release, "tag_name") as string);
                    if (version == null || (bool)draft ||
                        (version.IsBeta && !beta) || ((bool)prerelease && !version.IsBeta)) continue;
                    if (latest != null && version.CompareTo(latest) <= 0) continue;
                    latest = version;
                    selected = release;
                }
                if (latest == null) return NoRelease();
                var order = latest.CompareTo(installed);
                if (order < 0) return new UpdateResult(UpdateState.LocalNewer, "当前版本高于此渠道的最新版本（" + latest.Text + "）。");
                if (order == 0) return new UpdateResult(UpdateState.Current, "当前已是最新版（" + latest.Text + "）。");

                var name = "WordGit-" + latest.Text + "-Setup.exe";
                var tag = (string)Field(selected, "tag_name");
                var link = AppInfo.ReleasesUrl + "/tag/" + tag;
                var assets = Field(selected, "assets") as object[];
                if (assets != null && Array.Exists(assets, item =>
                {
                    var asset = item as Dictionary<string, object>;
                    return asset != null && (Field(asset, "name") as string) == name && (Field(asset, "state") as string) == "uploaded";
                })) link = AppInfo.ReleasesUrl + "/download/" + tag + "/" + name;
                var body = Field(selected, "body") as string;
                var notes = String.IsNullOrWhiteSpace(body) ? "此版本未填写更新日志。" : body.Trim();
                if (notes.Length > 24000) notes = notes.Substring(0, 24000) + "\n…";
                return new UpdateResult(UpdateState.Available, "当前版本 " + installed.Text + "  →  新版本 " + latest.Text,
                    latest.Text, notes, link);
            }
            catch (ArgumentException) { return InvalidResponse(); }
            catch (InvalidOperationException) { return InvalidResponse(); }
        }

        private static object Field(Dictionary<string, object> entry, string name)
        {
            object value;
            return entry.TryGetValue(name, out value) ? value : null;
        }

        private static UpdateResult NoRelease()
        {
            return new UpdateResult(UpdateState.NoRelease, "当前渠道暂无可用版本。");
        }

        private static UpdateResult InvalidResponse()
        {
            return new UpdateResult(UpdateState.Unavailable, "GitHub 返回的数据无法识别，请稍后重试。");
        }
    }
}
