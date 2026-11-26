using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace WordGit
{
    // Shared by every test program. The git layer now lives in the Rust core, so the helpers
    // below only cover files and the service handle.
    internal static class TestKernel
    {
        public static string Home { get; private set; }

        private static System.Diagnostics.Process owned;

        // Starts an isolated core in its own root and on its own port so tests never touch real
        // history and never collide with a service the user is already running.
        public static ServiceClient Start()
        {
            var executable = Environment.GetEnvironmentVariable("WORDGIT_SERVICE_EXE");
            if (String.IsNullOrEmpty(executable) || !File.Exists(executable))
                throw new InvalidOperationException("请先构建 Rust 内核，并用 WORDGIT_SERVICE_EXE 指向 wordgit.exe。");
            var home = Path.Combine(Path.GetTempPath(), "wordgit-kernel-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(home);
            Home = home;
            Environment.SetEnvironmentVariable("WORDGIT_HOME", home);
            var port = FreePort();
            var start = new System.Diagnostics.ProcessStartInfo(executable, "serve --port " + port)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(executable)
            };
            owned = System.Diagnostics.Process.Start(start);
            var file = Path.Combine(home, "service.json");
            var limit = DateTime.UtcNow.AddSeconds(25);
            while (!File.Exists(file))
            {
                if (owned.HasExited)
                    throw new InvalidOperationException("WordGit 内核启动后立即退出（代码 " + owned.ExitCode + "）。");
                if (DateTime.UtcNow > limit)
                    throw new InvalidOperationException("WordGit 内核未在 25 秒内就绪。");
                System.Threading.Thread.Sleep(100);
            }
            ServiceClient.Forget();
            var client = ServiceClient.Connect();
            client.Call("notices", ServiceClient.Args());
            return client;
        }

        public static void VerifyRestart(ServiceClient client)
        {
            var file = Path.Combine(Home, "service.json");
            var previous = Json.Parse(File.ReadAllText(file));
            var executable = owned.MainModule.FileName;
            owned.Kill();
            if (!owned.WaitForExit(10000)) throw new InvalidOperationException("测试内核没有退出。");
            owned.Dispose();
            owned = null;

            var configured = Environment.GetEnvironmentVariable("WORDGIT_SERVICE_EXE");
            Environment.SetEnvironmentVariable("WORDGIT_SERVICE_EXE", null);
            try
            {
                client.Call("notices", ServiceClient.Args());
                var next = Json.Parse(File.ReadAllText(file));
                if (Json.String(next, "token") == Json.String(previous, "token"))
                    throw new InvalidOperationException("WordGit 没有重新启动测试内核。");
                var process = System.Diagnostics.Process.GetProcessById(Json.Int(next, "pid"));
                if (!String.Equals(process.MainModule.FileName, executable, StringComparison.OrdinalIgnoreCase))
                {
                    process.Dispose();
                    throw new InvalidOperationException("WordGit 启动的内核与测试包不一致。");
                }
                owned = process;
            }
            finally { Environment.SetEnvironmentVariable("WORDGIT_SERVICE_EXE", configured); }
        }

        private static int FreePort()
        {
            var listener = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0);
            listener.Start();
            var port = ((System.Net.IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }

        public static void Stop(ServiceClient client)
        {
            try { ServiceClient.Forget(); } catch (Exception) { }
            if (owned != null)
            {
                try { owned.Kill(); owned.WaitForExit(10000); } catch (Exception) { }
                owned = null;
            }
            if (Home == null) return;
            try { Directory.Delete(Home, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            Home = null;
            Environment.SetEnvironmentVariable("WORDGIT_HOME", null);
        }
    }

    internal static class TestConsole
    {
        // Behave like a UTF-8 console code page (65001, as on CI runners), whose encoding has a BOM
        // preamble, so stream handling is exercised on every machine. Restored on dispose.
        public static IDisposable Utf8InputWithPreamble()
        {
            var previous = Console.InputEncoding;
            try { Console.InputEncoding = new UTF8Encoding(true); }
            catch (IOException) { return new Restore(null); }
            if (Console.InputEncoding.GetPreamble().Length == 0)
                throw new InvalidOperationException("Test console did not adopt a UTF-8 encoding with a preamble");
            return new Restore(previous);
        }

        private sealed class Restore : IDisposable
        {
            private readonly Encoding previous;
            public Restore(Encoding previous) { this.previous = previous; }
            public void Dispose()
            {
                if (previous == null) return;
                try { Console.InputEncoding = previous; } catch (IOException) { }
            }
        }
    }

    internal static class TestFiles
    {
        public static string Temporary(byte[] bytes)
        {
            var path = Path.Combine(Path.GetTempPath(), "wordgit-test-" + Guid.NewGuid().ToString("N") + ".docx");
            File.WriteAllBytes(path, bytes);
            return path;
        }

        public static void WriteNew(string path, byte[] bytes)
        {
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                output.Write(bytes, 0, bytes.Length);
        }

        public static byte[] ReadDocx(string path)
        {
            Validate(path);
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var buffer = new MemoryStream())
            {
                input.CopyTo(buffer);
                return buffer.ToArray();
            }
        }

        public static byte[] Snapshot(object document) { return ReadDocx(WordDocuments.SaveDocument(document)); }

        public static string Sha(byte[] bytes)
        {
            using (var sha = SHA256.Create())
            {
                var builder = new StringBuilder(64);
                foreach (var value in sha.ComputeHash(bytes)) builder.Append(value.ToString("x2"));
                return builder.ToString();
            }
        }

        // A local check that the file is a readable .docx, independent of the removed C# layer.
        public static void Validate(string path)
        {
            try
            {
                using (var archive = new ZipArchive(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), ZipArchiveMode.Read))
                {
                    if (archive.GetEntry("[Content_Types].xml") == null || archive.GetEntry("word/document.xml") == null)
                        throw new InvalidOperationException("不是有效的 .docx 文件：" + path);
                }
            }
            catch (InvalidDataException) { throw new InvalidOperationException("不是有效的 .docx 文件：" + path); }
        }

        public static bool SameDocx(string left, string right)
        {
            return ReadDocx(left).SequenceEqual(ReadDocx(right));
        }
    }
}