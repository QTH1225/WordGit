using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WordGit
{
    internal sealed class GitResult
    {
        public int Code;
        public byte[] Bytes;
        public string Error;
    }

    // Owns every git.exe invocation: argument quoting, environment isolation, durability settings,
    // timeouts and process-tree cleanup. Repository semantics live in GitHistory.
    internal sealed class GitProcess
    {
        internal const int DefaultTimeout = 45000;
        internal const int NetworkIdleTimeout = 180000;
        private static readonly Regex ProgressLine = new Regex(
            @"\A(?:remote: )?(?:Enumerating|Counting|Compressing|Receiving|Resolving|Writing|Total|Delta compression|Unpacking|Checking)\b",
            RegexOptions.CultureInvariant);

        // core.fsync=committed flushes loose objects, packs and refs before git reports success,
        // so a power loss right after a commit cannot leave a branch pointing at an empty object.
        private static readonly string[] LocalOptions = {
            "-c", "core.longpaths=true", "-c", "core.fsync=committed", "-c", "core.fsyncMethod=batch"
        };

        private readonly string executable, gitDirectory;

        internal GitProcess(string executable, string gitDirectory)
        {
            this.executable = executable;
            this.gitDirectory = gitDirectory;
        }

        internal string Executable { get { return executable; } }

        internal static string FindGit()
        {
            var explicitPath = Environment.GetEnvironmentVariable("WORDGIT_GIT");
            if (String.IsNullOrEmpty(explicitPath)) explicitPath = Environment.GetEnvironmentVariable("WORDGIT_GIT", EnvironmentVariableTarget.User);
            if (!String.IsNullOrEmpty(explicitPath))
            {
                if (!Path.IsPathRooted(explicitPath) || !File.Exists(explicitPath))
                    throw new InvalidOperationException("WORDGIT_GIT 必须指向有效的 git.exe 绝对路径。");
                return explicitPath;
            }
            foreach (var folder in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator)
                .Concat(new[] { Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Git", "cmd"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Git", "cmd"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Git", "cmd") }))
            {
                var entry = folder.Trim().Trim('"');
                if (Path.IsPathRooted(entry) && File.Exists(Path.Combine(entry, "git.exe"))) return Path.Combine(entry, "git.exe");
            }
            throw new InvalidOperationException("找不到 Git for Windows。请安装 Git，或用 WORDGIT_GIT 指定 git.exe。");
        }

        internal static string Quote(string arg)
        {
            var result = new StringBuilder("\"");
            var slashes = 0;
            foreach (var character in arg)
            {
                if (character == '\\') { slashes++; continue; }
                result.Append('\\', character == '"' ? slashes * 2 + 1 : slashes);
                result.Append(character);
                slashes = 0;
            }
            return result.Append('\\', slashes * 2).Append('"').ToString();
        }

        internal ProcessStartInfo CreateInfo(string identity, bool redirectInput, IEnumerable<string> args)
        {
            var command = new List<string>(LocalOptions) { "--git-dir=." };
            command.AddRange(args);
            var info = new ProcessStartInfo(executable, String.Join(" ", command.Select(Quote)))
            {
                WorkingDirectory = gitDirectory, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                RedirectStandardInput = redirectInput, StandardErrorEncoding = Encoding.UTF8
            };
            foreach (var key in info.EnvironmentVariables.Keys.Cast<string>().Where(key => key.StartsWith("GIT_", StringComparison.OrdinalIgnoreCase)).ToList())
                info.EnvironmentVariables.Remove(key);
            info.EnvironmentVariables["GIT_CONFIG_NOSYSTEM"] = "1";
            info.EnvironmentVariables["GIT_CONFIG_GLOBAL"] = "NUL";
            info.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            if (identity != null)
            {
                info.EnvironmentVariables["GIT_AUTHOR_NAME"] = identity;
                info.EnvironmentVariables["GIT_COMMITTER_NAME"] = identity;
                info.EnvironmentVariables["GIT_AUTHOR_EMAIL"] = "wordgit@local.invalid";
                info.EnvironmentVariables["GIT_COMMITTER_EMAIL"] = "wordgit@local.invalid";
            }
            return info;
        }

        internal GitResult Invoke(byte[] input, string identity, params string[] args)
        {
            return Invoke(input == null ? null : (Action<Stream>)(stream => stream.Write(input, 0, input.Length)),
                identity, DefaultTimeout, false, args);
        }

        internal GitResult Invoke(Action<Stream> input, string identity, int timeout, bool lowPriority, params string[] args)
        {
            return Run(CreateInfo(identity, input != null, args), input, timeout, lowPriority);
        }

        // Reads the user's own Git configuration (global and system), e.g. their SSH command.
        internal GitResult InvokeWithUserConfig(params string[] args)
        {
            var info = CreateInfo(null, false, args);
            info.EnvironmentVariables.Remove("GIT_CONFIG_NOSYSTEM");
            info.EnvironmentVariables.Remove("GIT_CONFIG_GLOBAL");
            return Run(info, null, DefaultTimeout, false);
        }

        private static GitResult Run(ProcessStartInfo info, Action<Stream> input, int timeout, bool lowPriority)
        {
            using (var running = ProcessTree.Start(info))
            using (var stdout = new MemoryStream())
            {
                var process = running.Process;
                if (lowPriority) running.LowerPriority();
                var read = process.StandardOutput.BaseStream.CopyToAsync(stdout);
                var errors = process.StandardError.ReadToEndAsync();
                var write = input == null ? Task.FromResult(true) : Task.Run(() =>
                {
                    using (var stream = process.StandardInput.BaseStream) input(stream);
                    return true;
                });
                if (!process.WaitForExit(timeout))
                {
                    running.Kill();
                    throw new TimeoutException("Git 操作超时。");
                }
                running.Drain(read, errors);
                if (process.ExitCode == 0) write.GetAwaiter().GetResult();
                // Surface writer errors that are not a consequence of git failing (a broken pipe
                // is just the downstream effect of git exiting; our own errors must not be lost).
                else if (write.IsFaulted && !(write.Exception.InnerException is IOException)) write.GetAwaiter().GetResult();
                else if (write.IsFaulted) GC.KeepAlive(write.Exception);
                return new GitResult { Code = process.ExitCode, Bytes = stdout.ToArray(), Error = errors.Result };
            }
        }

        // Like Invoke, but stdout is streamed to a reader callback instead of being buffered into
        // memory. If the reader reports a data problem, git is stopped and the reader's error is
        // thrown (it is more informative than git's exit code).
        internal void Stream(byte[] input, int timeout, Action<Stream> reader, params string[] args)
        {
            var info = CreateInfo(null, input != null, args);
            using (var running = ProcessTree.Start(info))
            {
                var process = running.Process;
                var errors = process.StandardError.ReadToEndAsync();
                var write = input == null ? Task.FromResult(true) : Task.Run(() =>
                {
                    using (var stream = process.StandardInput.BaseStream) stream.Write(input, 0, input.Length);
                    return true;
                });
                Exception readerError = null;
                var read = Task.Run(() =>
                {
                    try { reader(process.StandardOutput.BaseStream); }
                    catch (Exception error) { readerError = error; }
                });
                var exited = Task.Run(() => { process.WaitForExit(); return true; });
                if (Task.WaitAny(new Task[] { read, exited }, timeout) < 0)
                {
                    running.Kill();
                    throw new TimeoutException("Git 操作超时。");
                }
                if (readerError != null)
                {
                    running.Kill();
                    running.Drain(errors);
                    throw readerError;
                }
                if (!process.WaitForExit(timeout))
                {
                    running.Kill();
                    throw new TimeoutException("Git 操作超时。");
                }
                running.Drain(read, errors);
                if (readerError != null) throw readerError;
                if (process.ExitCode != 0) throw new InvalidOperationException("Git 操作失败：" + errors.Result.Trim());
                write.GetAwaiter().GetResult();
            }
        }

        // Network operations: there is no total deadline, only a deadline for silence, because a
        // first clone over a slow link can legitimately take far longer than any fixed limit.
        // Progress output (--progress) counts as activity. Cancellation and timeouts stop the
        // whole process tree, including git-remote-https and ssh.
        internal void RunNetwork(ProcessStartInfo info, CancellationToken cancellation, int idleTimeout)
        {
            cancellation.ThrowIfCancellationRequested();
            using (var running = ProcessTree.Start(info))
            {
                var process = running.Process;
                var lastActivity = Environment.TickCount;
                var detail = new StringBuilder();
                var pending = new StringBuilder();
                Action<string> collect = chunk =>
                {
                    Interlocked.Exchange(ref lastActivity, Environment.TickCount);
                    lock (detail)
                    {
                        pending.Append(chunk);
                        var text = pending.ToString();
                        var cut = Math.Max(text.LastIndexOf('\n'), text.LastIndexOf('\r'));
                        if (cut < 0) return;
                        foreach (var line in text.Substring(0, cut).Split('\r', '\n'))
                            if (line.Trim().Length > 0 && !ProgressLine.IsMatch(line.Trim())) detail.AppendLine(line.TrimEnd());
                        pending.Clear().Append(text.Substring(cut + 1));
                        if (detail.Length > 8000) detail.Remove(0, detail.Length - 4000);
                    }
                };
                var errors = Pump(process.StandardError, collect);
                var output = Pump(process.StandardOutput, chunk => Interlocked.Exchange(ref lastActivity, Environment.TickCount));
                var cancelled = false;
                while (!process.WaitForExit(250))
                {
                    if (cancellation.IsCancellationRequested) { cancelled = true; break; }
                    if (unchecked(Environment.TickCount - Interlocked.CompareExchange(ref lastActivity, 0, 0)) > idleTimeout)
                    {
                        running.Kill();
                        throw new TimeoutException("远程 Git 操作长时间没有进展，已停止。请检查网络，或完成 Git for Windows 登录后重试。");
                    }
                }
                if (cancelled)
                {
                    running.Kill();
                    throw new OperationCanceledException(cancellation);
                }
                running.Drain(output, errors);
                cancellation.ThrowIfCancellationRequested();
                if (process.ExitCode != 0)
                {
                    string message;
                    lock (detail) message = (detail.ToString() + pending.ToString()).Trim();
                    if (message.Length > 1600) message = message.Substring(message.Length - 1600);
                    throw new InvalidOperationException("远程 Git 操作未完成。推送被拒绝时请先获取并拉取；认证失败时请完成 Git for Windows 登录。\n" + message);
                }
            }
        }

        private static Task Pump(StreamReader reader, Action<string> chunk)
        {
            return Task.Run(() =>
            {
                var buffer = new char[4096];
                int count;
                while ((count = reader.Read(buffer, 0, buffer.Length)) > 0) chunk(new string(buffer, 0, count));
            });
        }

        // A started git process plus a Windows job object containing it and every helper it spawns.
        // Only an explicit Kill terminates the tree; a normal exit leaves intentionally detached
        // helpers (such as a credential cache daemon) alone.
        private sealed class ProcessTree : IDisposable
        {
            private const int DrainMilliseconds = 10000;
            internal readonly Process Process;
            private readonly IntPtr job;

            private ProcessTree(Process process, IntPtr job)
            {
                Process = process;
                this.job = job;
            }

            internal static ProcessTree Start(ProcessStartInfo info)
            {
                var process = Process.Start(info);
                var job = CreateJobObject(IntPtr.Zero, null);
                if (job != IntPtr.Zero && !AssignProcessToJobObject(job, process.Handle))
                {
                    CloseHandle(job);
                    job = IntPtr.Zero;
                }
                return new ProcessTree(process, job);
            }

            internal void LowerPriority()
            {
                try { Process.PriorityClass = ProcessPriorityClass.BelowNormal; }
                catch (System.ComponentModel.Win32Exception) { }
                catch (InvalidOperationException) { }
            }

            internal void Kill()
            {
                if (job != IntPtr.Zero) TerminateJobObject(job, 1);
                try { if (!Process.HasExited) Process.Kill(); }
                catch (InvalidOperationException) { }
                catch (System.ComponentModel.Win32Exception) { }
                Process.WaitForExit(DrainMilliseconds);
            }

            // Waits for the pipe readers; a helper that inherited a pipe and outlived git must not
            // hang the caller, so the tree is stopped if the pipes stay open too long.
            internal void Drain(params Task[] readers)
            {
                if (Task.WaitAll(readers, DrainMilliseconds)) return;
                Kill();
                if (!Task.WaitAll(readers, DrainMilliseconds)) throw new TimeoutException("Git 进程没有正常结束。");
            }

            public void Dispose()
            {
                if (job != IntPtr.Zero) CloseHandle(job);
                Process.Dispose();
            }

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            private static extern IntPtr CreateJobObject(IntPtr attributes, string name);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool TerminateJobObject(IntPtr job, uint exitCode);

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool CloseHandle(IntPtr handle);
        }
    }
}
