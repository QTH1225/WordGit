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
        private static ServiceClient current;

        private readonly string address;
        private readonly string token;
        private readonly HttpClient http;
        private readonly HttpClient quick;

        private ServiceClient(string address, string token)
        {
            this.address = address;
            this.token = token ?? String.Empty;
            http = new HttpClient();
            http.Timeout = TimeSpan.FromMinutes(10);
            http.DefaultRequestHeaders.Add("x-wordgit-token", this.token);
            quick = new HttpClient();
            quick.Timeout = TimeSpan.FromSeconds(4);
        }

        internal static string Root
        {
            get
            {
                var home = Environment.GetEnvironmentVariable("WORDGIT_HOME");
                return String.IsNullOrEmpty(home)
                    ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordGit")
                    : home;
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
                if (current != null) current.Dispose();
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
            var client = this;
            for (var attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    return client.CallOnce(endpoint, payload);
                }
                catch (ServiceException error)
                {
                    if (attempt == 0 && error.Retryable) { client = Refresh(); continue; }
                    throw;
                }
                catch (HttpRequestException)
                {
                    if (attempt == 0) { client = Refresh(); continue; }
                    throw new ServiceException("无法连接 WordGit 内核，请重新打开编辑器后再试。", "io");
                }
                catch (TaskCanceledException)
                {
                    if (attempt == 0) { client = Refresh(); continue; }
                    throw new ServiceException("WordGit 内核响应超时。", "timeout");
                }
            }
            throw new ServiceException("无法连接 WordGit 内核。", "io");
        }

        // Best effort: the operation reports the cancellation in its own response.
        internal void CancelOperation(string operation)
        {
            if (String.IsNullOrEmpty(operation)) return;
            try
            {
                using (var content = new StringContent(Json.Send(Args("operation", operation)), Encoding.UTF8, "application/json"))
                {
                    var response = quick.PostAsync(address + "/api/v1/cancel", content).GetAwaiter().GetResult();
                    if (response != null) response.Dispose();
                }
            }
            catch (Exception)
            {
                // Cancelling is advisory; an unreachable service has nothing to cancel.
            }
        }

        internal void Dispose()
        {
            http.Dispose();
            quick.Dispose();
        }

        private Dictionary<string, object> CallOnce(string endpoint, string payload)
        {
            using (var content = new StringContent(payload, Encoding.UTF8, "application/json"))
            {
                var response = http.PostAsync(address + "/api/v1/" + endpoint, content).GetAwaiter().GetResult();
                using (response)
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
        }

        private static ServiceClient Refresh()
        {
            lock (Gate)
            {
                if (current != null) current.Dispose();
                current = Discover();
                return current;
            }
        }

        private static ServiceClient Discover()
        {
            var running = FromServiceFile();
            if (running != null) return running;
            Start();
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
                if (port <= 0 || String.IsNullOrEmpty(token)) return null;
                var client = new ServiceClient("http://127.0.0.1:" + port.ToString(CultureInfo.InvariantCulture), token);
                if (client.Ping()) return client;
                client.Dispose();
                return null;
            }
            catch (IOException) { return null; }
            catch (UnauthorizedAccessException) { return null; }
            catch (InvalidOperationException) { return null; }
        }

        private static void Start()
        {
            var executable = ExecutablePath();
            if (executable == null) return;
            try
            {
                var info = new ProcessStartInfo(executable, "serve");
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

        private bool Ping()
        {
            try
            {
                var response = quick.GetAsync(address + "/api/v1/ping").GetAwaiter().GetResult();
                using (response)
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