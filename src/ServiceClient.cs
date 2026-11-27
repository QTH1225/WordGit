using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WordGit
{
    // A failure reported by the local core. Kind mirrors the service error kinds:
    // invalid, data, timeout, cancelled, io.
    internal sealed class ServiceException : Exception
    {
        internal readonly string Kind;
        internal readonly bool Retryable;

        internal ServiceException(string message, string kind, bool retryable = false) : base(message)
        {
            Kind = String.IsNullOrEmpty(kind) ? "io" : kind;
            Retryable = retryable;
        }

        internal bool IsCancelled { get { return Kind == "cancelled"; } }
    }

    // The local WordGit core: a 127.0.0.1 HTTP service. The add-in starts it on
    // demand, reads the shared service.json, and retries a call once after a
    // restart (the service rotates its token every time it starts).
    internal sealed class ServiceClient
    {
        internal const int DefaultPort = 47821;
        private static readonly object Gate = new object();
        private static readonly string DefaultRoot =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordGit");
        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };
        private static readonly HttpClient Quick = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
        private static ServiceClient current;

        private readonly int port;
        private readonly string address;
        private readonly string token;

        private ServiceClient(int port, string token)
        {
            this.port = port;
            address = "http://127.0.0.1:" + port.ToString(CultureInfo.InvariantCulture);
            this.token = token;
        }

        internal static string Root
        {
            get
            {
                var home = Environment.GetEnvironmentVariable("WORDGIT_HOME");
                return String.IsNullOrEmpty(home) ? DefaultRoot : home;
            }
        }

        private static string ServiceFile { get { return Path.Combine(Root, "service.json"); } }

        // Returns the shared client, discovering or starting the core when needed.
        // Callers never dispose it; a restart is picked up by the retry in Call.
        internal static ServiceClient Connect()
        {
            lock (Gate)
            {
                if (current != null) return current;
                current = Discover();
                return current;
            }
        }

        internal static void Forget()
        {
            lock (Gate)
            {
                current = null;
            }
        }

        // Convenience for building a request body: Args("document", path, "revision", id).
        internal static Dictionary<string, object> Args(params object[] pairs)
        {
            var result = new Dictionary<string, object>();
            for (var index = 0; index + 1 < pairs.Length; index += 2) result[(string)pairs[index]] = pairs[index + 1];
            return result;
        }

        internal DocumentStatus Status(string document)
        {
            return DocumentStatus.From(Call("status", Args("document", document)));
        }

        internal string WorkingState(string document)
        {
            return Json.String(Call("status", Args("document", document, "workingOnly", true)), "workingState");
        }

        // Swaps in the file the kernel prepared and publishes the matching ref change in one step.
        internal string ApplyPlan(string plan)
        {
            var value = Call("plan/apply", Args("plan", plan));
            var message = Json.String(value, "message");
            return String.IsNullOrEmpty(message) ? "已更新当前稿。" : message;
        }

        // A pending plan keeps the document lock; releasing it makes the next attempt possible.
        internal void CancelPlan(string plan)
        {
            if (String.IsNullOrEmpty(plan)) return;
            try { Call("plan/cancel", Args("plan", plan)); }
            catch (ServiceException) { }
            catch (InvalidOperationException) { }
        }

        internal Dictionary<string, object> Call(string endpoint, IDictionary<string, object> body)
        {
            var payload = Json.Send(body);
            var client = Connect();
            for (var attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    return client.CallOnce(endpoint, payload);
                }
                catch (ServiceException error)
                {
                    if (error.IsCancelled) throw new OperationCanceledException(error.Message, error);
                    if (attempt == 0 && error.Retryable) { client = client.Refresh(); continue; }
                    throw;
                }
                catch (HttpRequestException)
                {
                    if (attempt == 0) { client = client.Refresh(); continue; }
                    throw new ServiceException("无法连接 WordGit 内核，请重新打开编辑器后再试。", "io");
                }
                catch (TaskCanceledException)
                {
                    if (attempt == 0) { client = client.Refresh(); continue; }
                    throw new ServiceException("WordGit 内核响应超时。", "timeout");
                }
            }
            throw new ServiceException("无法连接 WordGit 内核。", "io");
        }

        // Best effort: the operation reports the cancellation in its own response.
        internal void CancelOperation(string operation)
        {
            if (String.IsNullOrEmpty(operation)) return;
            using (var request = CreateRequest("cancel", Json.Send(Args("operation", operation))))
            using (var response = Quick.SendAsync(request).GetAwaiter().GetResult())
            {
                response.EnsureSuccessStatusCode();
            }
        }

        private Dictionary<string, object> CallOnce(string endpoint, string payload)
        {
            using (var request = CreateRequest(endpoint, payload))
            using (var response = Http.SendAsync(request).GetAwaiter().GetResult())
            {
                var text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                Dictionary<string, object> parsed = null;
                try { parsed = Json.Parse(text); }
                catch (InvalidOperationException) { parsed = null; }
                if (!response.IsSuccessStatusCode)
                {
                    var message = parsed == null ? null : Json.String(parsed, "error");
                    var kind = parsed == null ? null : Json.String(parsed, "kind");
                    throw new ServiceException(
                        String.IsNullOrEmpty(message) ? "WordGit 内核返回错误（" + (int)response.StatusCode + "）。" : message,
                        kind, response.StatusCode == HttpStatusCode.Unauthorized);
                }
                if (parsed == null) throw new ServiceException("WordGit 内核返回的数据无法识别。", "io");
                return parsed;
            }
        }

        private HttpRequestMessage CreateRequest(string endpoint, string payload = null)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, address + "/api/v1/" + endpoint);
            request.Headers.Add("x-wordgit-token", token);
            if (payload != null) request.Content = new StringContent(payload, Encoding.UTF8, "application/json");
            return request;
        }

        private ServiceClient Refresh()
        {
            lock (Gate)
            {
                if (current != null && current != this) return current;
                current = null;
                current = Discover(port);
                return current;
            }
        }

        private static ServiceClient Discover(int port = DefaultPort)
        {
            var running = FromServiceFile();
            if (running != null) return running;
            Start(port);
            var started = WaitForStartup();
            if (started != null) return started;
            throw new InvalidOperationException(
                "WordGit 内核没有运行，也无法自动启动。请重新安装 WordGit，或手动运行 wordgit.exe serve。");
        }

        private static ServiceClient FromServiceFile()
        {
            try
            {
                if (!File.Exists(ServiceFile)) return null;
                var info = Json.Parse(File.ReadAllText(ServiceFile));
                var port = Json.Int(info, "port");
                var token = Json.String(info, "token");
                if (port <= 0 || port > UInt16.MaxValue || String.IsNullOrEmpty(token)) return null;
                var client = new ServiceClient(port, token);
                return client.IsReady() ? client : null;
            }
            catch (IOException) { return null; }
            catch (UnauthorizedAccessException) { return null; }
            catch (InvalidOperationException) { return null; }
        }

        private static void Start(int port)
        {
            var executable = ExecutablePath();
            if (executable == null) return;
            try
            {
                var info = new ProcessStartInfo(executable, "serve --port " + port.ToString(CultureInfo.InvariantCulture));
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;
                info.WorkingDirectory = Path.GetDirectoryName(executable);
                Process.Start(info);
            }
            catch (System.ComponentModel.Win32Exception) { }
            catch (InvalidOperationException) { }
        }

        private static string ExecutablePath()
        {
            var configured = Environment.GetEnvironmentVariable("WORDGIT_SERVICE_EXE");
            if (!String.IsNullOrEmpty(configured) && File.Exists(configured)) return configured;
            // The COM add-in is loaded into Word's application domain, whose base is the host.
            var assemblyDirectory = Path.GetDirectoryName(typeof(ServiceClient).Assembly.Location);
            if (!String.IsNullOrEmpty(assemblyDirectory))
            {
                var alongside = Path.Combine(assemblyDirectory, "wordgit.exe");
                if (File.Exists(alongside)) return alongside;
            }

            var installed = Path.Combine(DefaultRoot, "app", "wordgit.exe");
            if (File.Exists(installed)) return installed;
            var local = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wordgit.exe");
            return File.Exists(local) ? local : null;
        }

        private static ServiceClient WaitForStartup()
        {
            var deadline = DateTime.UtcNow.AddSeconds(20);
            while (DateTime.UtcNow < deadline)
            {
                Thread.Sleep(120);
                var client = FromServiceFile();
                if (client != null) return client;
            }
            return null;
        }

        private bool IsReady()
        {
            try
            {
                // The public ping succeeds even when service.json still contains the previous token.
                using (var request = CreateRequest("info"))
                using (var response = Quick.SendAsync(request).GetAwaiter().GetResult())
                {
                    if (!response.IsSuccessStatusCode) return false;
                    var text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    var info = Json.Parse(text);
                    if (!Json.Bool(info, "ok")) return false;
                    if (Json.Int(info, "api") != 1)
                        throw new ServiceException("WordGit 内核与本插件的版本不匹配，请重新安装 WordGit。", "version");
                    return true;
                }
            }
            catch (HttpRequestException) { return false; }
            catch (TaskCanceledException) { return false; }
            catch (InvalidOperationException) { return false; }
        }
    }
}
