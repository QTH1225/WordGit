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
        internal readonly string Message;
        internal UpdateResult(UpdateState state, string message) { State = state; Message = message; }
    }

    internal static class UpdateChecker
    {
        internal static async Task<UpdateResult> CheckAsync(Version current, CancellationToken cancellation)
        {
            using (var handler = new HttpClientHandler { AllowAutoRedirect = false, SslProtocols = SslProtocols.Tls12 })
            using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(12), MaxResponseContentBufferSize = 256 * 1024 })
                return await CheckAsync(current, client, cancellation).ConfigureAwait(false);
        }

        internal static async Task<UpdateResult> CheckAsync(Version current, HttpClient client, CancellationToken cancellation)
        {
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, AppInfo.LatestReleaseApi))
                {
                    request.Headers.UserAgent.ParseAdd("WordGit/" + AppInfo.DisplayVersion);
                    request.Headers.Accept.ParseAdd("application/vnd.github+json");
                    request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");
                    using (var response = await client.SendAsync(request, cancellation).ConfigureAwait(false))
                    {
                        if (response.StatusCode == HttpStatusCode.NotFound)
                            return new UpdateResult(UpdateState.NoRelease, "未找到可用的正式发布版本。请打开发布页确认；源码提交不等同于软件发布。");
                        if (response.StatusCode == HttpStatusCode.Forbidden || (int)response.StatusCode == 429)
                            return new UpdateResult(UpdateState.RateLimited, "GitHub 暂时拒绝或限制了更新查询，请稍后重试，或直接打开发布页。");
                        if (!response.IsSuccessStatusCode)
                            return new UpdateResult(UpdateState.Unavailable, "检查更新失败（HTTP " + (int)response.StatusCode + "）。可直接打开发布页。");
                        return Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false), current);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                if (cancellation.IsCancellationRequested) throw;
                return new UpdateResult(UpdateState.Unavailable, "检查更新超时，请确认网络后重试，或打开发布页。");
            }
            catch (HttpRequestException)
            {
                return new UpdateResult(UpdateState.Unavailable, "无法连接 GitHub，请确认网络后重试，或打开发布页。");
            }
        }

        internal static UpdateResult Parse(string json, Version current)
        {
            try
            {
                var payload = new JavaScriptSerializer { MaxJsonLength = 256 * 1024 }.Deserialize<Dictionary<string, object>>(json);
                object tag, draft, prerelease;
                if (payload == null || !payload.TryGetValue("draft", out draft) || !(draft is bool) ||
                    !payload.TryGetValue("prerelease", out prerelease) || !(prerelease is bool))
                    throw new ArgumentException();
                if ((bool)draft || (bool)prerelease)
                    return new UpdateResult(UpdateState.NoRelease, "该版本为草稿或预发布版本，请前往发布页查看。");
                if (!payload.TryGetValue("tag_name", out tag) || !(tag is string)) throw new ArgumentException();
                var match = Regex.Match((string)tag, @"\Av?(\d+)\.(\d+)\.(\d+)(?:\.(\d+))?\z", RegexOptions.IgnoreCase);
                Version latest;
                if (!match.Success || !System.Version.TryParse(match.Value.TrimStart('v', 'V'), out latest))
                    return new UpdateResult(UpdateState.Unsupported, "发布版本号无法识别，请前往发布页确认。");
                latest = Normalize(latest);
                var difference = latest.CompareTo(Normalize(current));
                if (difference > 0)
                    return new UpdateResult(UpdateState.Available, "发现新版本 " + latest.ToString(3) + "。请打开发布页下载安装；安装前保存并关闭 Word，现有历史会保留。");
                if (difference < 0)
                    return new UpdateResult(UpdateState.LocalNewer, "当前本地版本高于最新正式发布版本（" + latest.ToString(3) + "），无需降级。");
                return new UpdateResult(UpdateState.Current, "当前已是最新正式发布版本（" + latest.ToString(3) + "）。");
            }
            catch (ArgumentException) { return InvalidResponse(); }
            catch (InvalidOperationException) { return InvalidResponse(); }
        }

        private static Version Normalize(Version version)
        {
            return new Version(version.Major, version.Minor, Math.Max(0, version.Build), Math.Max(0, version.Revision));
        }

        private static UpdateResult InvalidResponse()
        {
            return new UpdateResult(UpdateState.Unavailable, "GitHub 返回的数据无法识别，请稍后重试，或打开发布页。");
        }
    }
}
