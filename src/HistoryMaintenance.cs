using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace WordGit
{
    internal sealed class HistoryMaintenance : IDisposable
    {
        internal const int CommitThreshold = 10;
        private const int IdleMilliseconds = 20000;
        private readonly object gate = new object();
        private readonly HashSet<string> pending = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Timer timer;
        private readonly string repositoryDirectory;
        private readonly Func<bool> isIdle;
        private bool disposed;
        private bool scanOnStart = true;
        private int running;

        public HistoryMaintenance(string repositoryDirectory, Func<bool> isIdle = null)
        {
            this.repositoryDirectory = repositoryDirectory;
            this.isIdle = isIdle ?? IsIdle;
            timer = new Timer(Check, null, 5000, 5000);
        }

        public void Schedule(string repositoryPath)
        {
            lock (gate)
                if (!disposed) pending.Add(repositoryPath);
        }

        public void Dispose()
        {
            lock (gate)
            {
                disposed = true;
                pending.Clear();
            }
            timer.Dispose();
        }

        private void Check(object state)
        {
            if (!isIdle() || Interlocked.Exchange(ref running, 1) != 0) return;
            try
            {
                if (scanOnStart)
                {
                    scanOnStart = false;
                    // Pick up commits that reached the threshold shortly before Word was closed.
                    if (Directory.Exists(repositoryDirectory))
                        foreach (var repository in Directory.EnumerateDirectories(repositoryDirectory, "*.git"))
                            if (Path.GetFileNameWithoutExtension(repository).Length == 64) Schedule(repository);
                }

                while (isIdle())
                {
                    var path = TakeNext();
                    if (path == null) return;

                    try
                    {
                        var repo = GitHistory.OpenRepository(path);
                        if (!repo.TryAutoRepack(CommitThreshold, CanStart))
                        {
                            Schedule(path);
                            return;
                        }
                    }
                    catch (Exception error)
                    {
                        AppPaths.LogError("background repack check: " + error);
                    }
                }
            }
            catch (Exception error)
            {
                AppPaths.LogError("background maintenance: " + error);
            }
            finally { Interlocked.Exchange(ref running, 0); }
        }

        private string TakeNext()
        {
            lock (gate)
            {
                if (disposed || pending.Count == 0) return null;
                var path = pending.First();
                pending.Remove(path);
                return path;
            }
        }

        private bool CanStart()
        {
            lock (gate) return !disposed && isIdle();
        }

        private static bool IsIdle()
        {
            var input = new LastInput { Size = (uint)Marshal.SizeOf(typeof(LastInput)) };
            if (!GetLastInputInfo(ref input)) return false;
            var elapsed = unchecked((uint)Environment.TickCount - input.Tick);
            return elapsed >= IdleMilliseconds && elapsed < Int32.MaxValue;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LastInput
        {
            public uint Size, Tick;
        }

        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LastInput input);
    }
}
