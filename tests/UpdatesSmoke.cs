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
        private static string Release(string tag, bool draft = false, bool prerelease = false)
        {
            return new JavaScriptSerializer().Serialize(new { tag_name = tag, draft = draft, prerelease = prerelease,
                html_url = "https://example.invalid/untrusted-link" });
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
                    var result = UpdateChecker.CheckAsync(AppInfo.Version, CancellationToken.None).GetAwaiter().GetResult();
                    Console.WriteLine(result.State + ": " + result.Message);
                    return result.State == UpdateState.Unavailable ? 1 : 0;
                }
                Run().GetAwaiter().GetResult();
                Console.WriteLine("PASS: numeric release ordering, malformed/draft releases, HTTP failures, cancellation and request destination.");
                return 0;
            }
            catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        }

        private static async Task Run()
        {
            var current = new Version(0, 9, 7, 0);
            Require(UpdateChecker.Parse(Release("v0.9.10"), current).State == UpdateState.Available, "Version ordering is lexical");
            Require(UpdateChecker.Parse(Release("0.9.7"), current).State == UpdateState.Current, "Three-component version differs");
            Require(UpdateChecker.Parse(Release("v0.9.6"), current).State == UpdateState.LocalNewer, "Newer local build should not downgrade");
            Require(UpdateChecker.Parse(Release("v0.9.8", draft: true), current).State == UpdateState.NoRelease, "Draft is offered");
            Require(UpdateChecker.Parse(Release("v0.9.8", prerelease: true), current).State == UpdateState.NoRelease, "Prerelease is offered");
            Require(UpdateChecker.Parse(Release("v0.9.8-preview"), current).State == UpdateState.Unsupported, "Unknown version is offered");
            foreach (var json in new[] { "{", "null", "[]", "{}", "{\"tag_name\":\"v0.9.8\"}" })
                Require(UpdateChecker.Parse(json, current).State == UpdateState.Unavailable, "Malformed result reported as current");
            foreach (var status in new[] { 200, 404, 403, 429, 500 })
            using (var handler = new StubHandler { Code = status })
            using (var client = new HttpClient(handler))
            {
                var result = await UpdateChecker.CheckAsync(current, client, CancellationToken.None);
                var expected = status == 200 ? UpdateState.Available : status == 404 ? UpdateState.NoRelease :
                    status == 403 || status == 429 ? UpdateState.RateLimited : UpdateState.Unavailable;
                Require(result.State == expected && handler.ExpectedRequest, "Wrong status or unexpected request destination");
            }
            foreach (var failure in new[] { 1, 2 })
            using (var handler = new StubHandler { Failure = failure })
            using (var client = new HttpClient(handler))
                Require((await UpdateChecker.CheckAsync(current, client, CancellationToken.None)).State == UpdateState.Unavailable, "Network failure escaped");
            using (var cancellation = new CancellationTokenSource())
            using (var client = new HttpClient(new StubHandler()))
            {
                cancellation.Cancel();
                var cancelled = false;
                try { await UpdateChecker.CheckAsync(current, client, cancellation.Token); }
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
                ExpectedRequest = request.Method == HttpMethod.Get && request.RequestUri.AbsoluteUri == AppInfo.LatestReleaseApi &&
                    request.Headers.UserAgent.ToString().StartsWith("WordGit/") && request.Headers.Authorization == null;
                if (Failure == 1) throw new HttpRequestException("offline");
                if (Failure == 2) throw new TaskCanceledException("timeout");
                return Task.FromResult(new HttpResponseMessage((HttpStatusCode)Code) { Content = new StringContent(Release("v0.9.10")) });
            }
        }
    }
}
