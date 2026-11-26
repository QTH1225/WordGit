using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace WordGit
{
    internal static class UpdatesSmoke
    {
        private static string Release(string tag, bool draft = false, bool prerelease = false, string asset = null)
        {
            return new JavaScriptSerializer().Serialize(new { tag_name = tag, draft = draft, prerelease = prerelease,
                html_url = "https://example.invalid/untrusted-link", body = "更新日志",
                assets = new[] { new { name = asset, state = "uploaded", browser_download_url = "https://example.invalid/untrusted-link" } } });
        }

        private static void Require(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static int Main(string[] args)
        {
            try
            {
                if (args.Length == 1 && args[0] == "--live")
                {
                    var result = UpdateChecker.CheckAsync(AppInfo.DisplayVersion, false, CancellationToken.None).GetAwaiter().GetResult();
                    Console.WriteLine(result.State + ": " + result.Message);
                    return result.State == UpdateState.Unavailable ? 1 : 0;
                }
                Run().GetAwaiter().GetResult();
                Console.WriteLine("PASS: stable and Beta ordering, trusted download, malformed/draft releases, HTTP failures, cancellation and request destination.");
                return 0;
            }
            catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        }

        private static async Task Run()
        {
            const string current = "0.12.0";
            Func<string[], string> list = items => "[" + String.Join(",", items) + "]";
            var setup = "WordGit-0.12.1-beta.10-Setup.exe";
            var releases = list(new[] { Release("v0.12.1-beta.2", prerelease: true),
                Release("v0.12.1-beta.10", asset: setup),
                Release("v0.12.3", draft: true), Release("0.12.0") });
            Require(UpdateChecker.Parse(releases, current, false).State == UpdateState.Current, "Beta was offered to stable users");
            var beta = UpdateChecker.Parse(releases, "0.12.1-beta.2", true);
            Require(beta.State == UpdateState.Available && beta.Version == "0.12.1-beta.10", "Beta ordering is lexical");
            Require(beta.Notes == "更新日志" && beta.DownloadUrl == AppInfo.ReleasesUrl + "/download/v0.12.1-beta.10/" + setup &&
                AppInfo.IsAllowed(beta.DownloadUrl), "Update did not use its trusted installer and notes");
            Require(!AppInfo.IsAllowed("https://example.invalid/") &&
                !AppInfo.IsAllowed(AppInfo.ReleasesUrl + "/download/v0.12.1-beta.10/setup.exe?url=other"),
                "Untrusted download was accepted");
            Require(UpdateChecker.Parse(list(new[] { Release("v0.12.1") }), "0.12.1-beta.11", true).State == UpdateState.Available,
                "Stable should supersede its own beta");
            Require(UpdateChecker.Parse(releases, "0.12.1-beta.11", true).State == UpdateState.LocalNewer, "Newer local beta would downgrade");
            Require(UpdateChecker.Parse(list(new[] { Release("v0.12.1", prerelease: true) }), current, true).State == UpdateState.NoRelease,
                "Numeric preview was offered as a regular release");
            foreach (var json in new[] { "{", "null", "{}", "[{}]", "[null]" })
                Require(UpdateChecker.Parse(json, current, false).State == UpdateState.Unavailable, "Malformed result reported as current");
            foreach (var status in new[] { 200, 404, 403, 429, 500 })
            using (var handler = new StubHandler { Code = status })
            using (var client = new HttpClient(handler))
            {
                var result = await UpdateChecker.CheckAsync(current, false, client, CancellationToken.None);
                var expected = status == 200 ? UpdateState.Available : status == 404 ? UpdateState.NoRelease :
                    status == 403 || status == 429 ? UpdateState.RateLimited : UpdateState.Unavailable;
                Require(result.State == expected && handler.ExpectedRequest, "Wrong status or unexpected request destination");
            }
            foreach (var failure in new[] { 1, 2 })
            using (var handler = new StubHandler { Failure = failure })
            using (var client = new HttpClient(handler))
                Require((await UpdateChecker.CheckAsync(current, false, client, CancellationToken.None)).State == UpdateState.Unavailable, "Network failure escaped");
            using (var cancellation = new CancellationTokenSource())
            using (var client = new HttpClient(new StubHandler()))
            {
                cancellation.Cancel();
                var cancelled = false;
                try { await UpdateChecker.CheckAsync(current, false, client, cancellation.Token); }
                catch (OperationCanceledException) { cancelled = true; }
                Require(cancelled, "Closing update window does not cancel the request");
            }
        }

        private sealed class StubHandler : HttpMessageHandler
        {
            internal int Code = 200, Failure;
            internal bool ExpectedRequest;
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellation)
            {
                cancellation.ThrowIfCancellationRequested();
                ExpectedRequest = request.Method == HttpMethod.Get && request.RequestUri.AbsoluteUri == AppInfo.ReleasesApi &&
                    request.Headers.UserAgent.ToString().StartsWith("WordGit/") && request.Headers.Authorization == null;
                if (Failure == 1) throw new HttpRequestException("offline");
                if (Failure == 2) throw new TaskCanceledException("timeout");
                return Task.FromResult(new HttpResponseMessage((HttpStatusCode)Code) { Content = new StringContent("[" + Release("v0.12.1") + "]") });
            }
        }
    }
}
