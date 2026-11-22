using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WordGit
{
    internal enum SyncOperation { Fetch, Push, Pull }

    internal sealed class RemoteState
    {
        internal string Url, LocalBranch, RemoteBranch, RemoteReference, LocalTip, RemoteTip;
        internal int Ahead, Behind;
        internal bool Unrelated;
        internal bool Linked { get { return !String.IsNullOrEmpty(Url); } }
        internal string Caption
        {
            get {
                if (!Linked) return "尚未关联 GitHub 仓库";
                if (RemoteTip == null) return "origin/" + RemoteBranch + " · 尚未获取到该分支";
                if (Unrelated) return "origin/" + RemoteBranch + " · 没有共同历史";
                return "origin/" + RemoteBranch + " · 领先 " + Ahead + " / 落后 " + Behind;
            }
        }
    }

    internal sealed partial class GitHistory
    {
        internal const string RemotePrefix = "refs/remotes/origin/";
        private const string BackupRoot = "refs/wordgit/remote-backup/";
        // A run of timestamped backups protects against remote history rewrites; the cap prevents unbounded growth.
        private const int BackupRetention = 20;

        internal static string ReferenceLabel(string reference)
        {
            return reference != null && reference.StartsWith(RemotePrefix, StringComparison.Ordinal)
                ? "origin/" + reference.Substring(RemotePrefix.Length) : reference;
        }

        private static string HistoryReference(string branch)
        {
            if (branch != null && branch.StartsWith(RemotePrefix, StringComparison.Ordinal))
                return RemotePrefix + BranchRef(branch.Substring(RemotePrefix.Length)).Substring(Heads.Length);
            return BranchRef(branch);
        }

        internal static string NormalizeRemoteUrl(string url, bool allowLocal = false)
        {
            url = (url ?? "").Trim();
            if (allowLocal && Path.IsPathRooted(url) && Directory.Exists(url)) return Path.GetFullPath(url);
            var match = Regex.Match(url, @"\A(?:https://github\.com/|git@github\.com:|ssh://git@github\.com/)([A-Za-z0-9_.-]+)/([A-Za-z0-9_.-]+?)(?:\.git)?/?\z", RegexOptions.IgnoreCase);
            if (!match.Success || new[] { ".", ".." }.Contains(match.Groups[1].Value) || new[] { ".", ".." }.Contains(match.Groups[2].Value))
                throw new ArgumentException("请输入 GitHub 仓库的 HTTPS 或 SSH 地址，不包含账号口令、查询参数或其他命令。");
            var path = match.Groups[1].Value + "/" + match.Groups[2].Value + ".git";
            return url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ? "https://github.com/" + path : "git@github.com:" + path;
        }

        private string ConfigValue(string key)
        {
            if (!File.Exists(Path.Combine(root, "HEAD"))) return null;
            var result = Invoke(null, null, "config", "--local", "--get", key);
            if (result.Code == 1) return null;
            return Text(result);
        }

        internal string RemoteUrl { get { return ConfigValue("remote.origin.url"); } }

        internal void ConfigureRemote(string url, string remoteBranch, bool allowLocal = false)
        {
            url = NormalizeRemoteUrl(url, allowLocal);
            var reference = BranchRef(remoteBranch);
            EnsureRepository();
            Run(null, null, "check-ref-format", reference);
            var old = RemoteUrl;
            if (old != null && !String.Equals(NormalizeRemoteUrl(old, allowLocal), url, StringComparison.OrdinalIgnoreCase))
            {
                ArchiveRemoteTips();
                CleanRemoteBackups();
                var previous = RemoteBranches();
                if (previous.Length > 0)
                    RefTransaction(null, previous.Select(oldReference => "delete " + oldReference + " " + BranchTip(oldReference)).ToArray());
                Run(null, null, "config", "--local", "--add", "wordgit.previousRemote", old);
                Run(null, null, "remote", "set-url", "origin", url);
            }
            if (old == null) Run(null, null, "remote", "add", "origin", url);
            Run(null, null, "config", "--local", "remote.origin.fetch", "+refs/heads/*:refs/remotes/origin/*");
            Run(null, null, "config", "--local", "branch." + ActiveBranch() + ".remote", "origin");
            Run(null, null, "config", "--local", "branch." + ActiveBranch() + ".merge", reference);
        }

        internal string[] RemoteBranches()
        {
            if (RemoteUrl == null) return new string[0];
            return Text(Run(null, null, "for-each-ref", "--format=%(refname)", RemotePrefix))
                .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(reference => reference != RemotePrefix + "HEAD").ToArray();
        }

        internal RemoteState SyncState()
        {
            var local = ActiveBranch();
            var merge = ConfigValue("branch." + local + ".merge");
            var branch = merge != null && merge.StartsWith(Heads, StringComparison.Ordinal) ? merge.Substring(Heads.Length) : local;
            var state = new RemoteState { Url = RemoteUrl, LocalBranch = local, RemoteBranch = branch,
                LocalTip = BranchTip(local), RemoteReference = RemotePrefix + branch };
            if (!state.Linked) return state;
            state.RemoteTip = BranchTip(state.RemoteReference);
            if (state.RemoteTip == null || state.LocalTip == null) return state;
            var counts = Text(Run(null, null, "rev-list", "--left-right", "--count", state.LocalTip + "..." + state.RemoteTip))
                .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            state.Ahead = Int32.Parse(counts[0], CultureInfo.InvariantCulture);
            state.Behind = Int32.Parse(counts[1], CultureInfo.InvariantCulture);
            var common = Invoke(null, null, "merge-base", state.LocalTip, state.RemoteTip);
            if (common.Code == 1) state.Unrelated = true;
            else Id(Text(common));
            return state;
        }

        private void ArchiveRemoteTips()
        {
            var refs = RemoteBranches();
            if (refs.Length == 0) return;
            ArchiveTips(refs.ToDictionary(reference => reference, reference => BranchTip(reference)));
        }

        private void ArchiveTips(IDictionary<string, string> tips)
        {
            if (tips.Count == 0) return;
            var prefix = BackupRoot + DateTime.UtcNow.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) +
                "-" + Guid.NewGuid().ToString("N").Substring(0, 6) + "/";
            RefTransaction(null, tips.Select(pair =>
                "create " + prefix + pair.Key.Substring(RemotePrefix.Length) + " " + Id(pair.Value)).ToArray());
        }

        private string[] BackupRefs()
        {
            return Text(Run(null, null, "for-each-ref", "--format=%(refname)", BackupRoot))
                .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        }

        internal string[] RemoteBackupNamespaces()
        {
            return BackupRefs()
                .Select(reference => reference.Substring(BackupRoot.Length).Split('/')[0])
                .Distinct().ToArray();
        }

        // Keeps only the newest timestamped backups; GUID namespaces left by older versions always count as oldest.
        private void CleanRemoteBackups()
        {
            var stale = BackupRefs()
                .GroupBy(reference => reference.Substring(BackupRoot.Length).Split('/')[0])
                .OrderByDescending(group => BackupSortKey(group.Key), StringComparer.Ordinal)
                .Skip(BackupRetention)
                .SelectMany(group => group)
                .Select(reference => new { Reference = reference, Tip = RefTip(reference) })
                .Where(item => item.Tip != null)
                .ToArray();
            if (stale.Length > 0)
                RefTransaction(null, stale.Select(item => "delete " + item.Reference + " " + item.Tip).ToArray());
        }

        private static string BackupSortKey(string name)
        {
            return Regex.IsMatch(name, @"\A\d{8}") ? "1" + name : "0" + name;
        }

        private string RefTip(string reference)
        {
            var result = Invoke(null, null, "rev-parse", "--verify", "-q", reference);
            if (result.Code == 1) return null;
            return Text(result);
        }

        internal RemoteState FetchRemote(bool allowLocal = false, CancellationToken cancellation = default(CancellationToken))
        {
            NormalizeRemoteUrl(RemoteUrl, allowLocal);
            var before = RemoteBranches().ToDictionary(reference => reference, reference => BranchTip(reference));
            NetworkRun(allowLocal, cancellation, "-c", "transfer.fsckObjects=true", "-c", "fetch.fsckObjects=true",
                "fetch", "--progress", "--no-prune", "--no-tags", "--no-recurse-submodules", "origin",
                "+refs/heads/*:refs/remotes/origin/*");
            // Only a forced (non-fast-forward) remote update can hide commits we already fetched;
            // those old tips are kept under refs/wordgit/remote-backup so they stay in the history.
            var rewritten = before.Where(pair => pair.Value != null)
                .Where(pair => { var now = BranchTip(pair.Key); return now != pair.Value && (now == null || !IsAncestor(pair.Value, now)); })
                .ToDictionary(pair => pair.Key, pair => pair.Value);
            ArchiveTips(rewritten);
            CleanRemoteBackups();
            return SyncState();
        }

        internal RemoteState PushRemote(bool allowLocal = false, CancellationToken cancellation = default(CancellationToken))
        {
            var state = SyncState();
            if (!state.Linked || state.LocalTip == null) throw new InvalidOperationException("请先关联仓库并提交当前稿。");
            NormalizeRemoteUrl(state.Url, allowLocal);
            // A normal push refuses a non-fast-forward update on the server.
            NetworkRun(allowLocal, cancellation, "push", "--progress", "--porcelain", "origin",
                BranchRef(state.LocalBranch) + ":" + BranchRef(state.RemoteBranch));
            return SyncState();
        }

        internal void FastForwardRemote(string branch, string oldTip, string reference, string remoteTip)
        {
            if (reference == null || !reference.StartsWith(RemotePrefix, StringComparison.Ordinal) ||
                ActiveBranch() != branch || BranchTip(branch) != oldTip || BranchTip(reference) != remoteTip)
                throw new InvalidOperationException("本地或远程跟踪分支已变化，请重新拉取。");
            var ancestor = Invoke(null, null, "merge-base", "--is-ancestor", Id(oldTip), Id(remoteTip));
            if (ancestor.Code != 0) throw new InvalidOperationException("双方已有不同修改，需要通过 Word 修订合并。");
            RefTransaction(null, "update " + BranchRef(branch) + " " + Id(remoteTip) + " " + Id(oldTip),
                "verify " + HistoryReference(reference) + " " + Id(remoteTip));
        }

        internal static GitHistory CloneRemote(string url, string remoteBranch, string documentPath,
            string storage = null, bool allowLocal = false, CancellationToken cancellation = default(CancellationToken))
        {
            url = NormalizeRemoteUrl(url, allowLocal);
            documentPath = Path.GetFullPath(documentPath);
            if (Path.GetExtension(documentPath).ToLowerInvariant() != ".docx" || File.Exists(documentPath) || Directory.Exists(documentPath))
                throw new IOException("请选择一个尚不存在的本地 .docx 文件名。");
            var destination = new GitHistory(documentPath, storage);
            if (Directory.Exists(destination.root)) throw new IOException("该路径已有历史仓库，请选择新的文件名；原历史不会被替换。");
            var parent = Path.GetDirectoryName(destination.root);
            Directory.CreateDirectory(parent);
            var temporary = Path.Combine(parent, "clone-" + Guid.NewGuid().ToString("N").Substring(0, 12) + ".git");
            var stagedDocument = Path.Combine(Path.GetDirectoryName(documentPath), "wordgit-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                var staged = new GitHistory(temporary, destination.git.Executable, null);
                staged.EnsureRepository();
                RunBranch(staged, remoteBranch);
                staged.ConfigureRemote(url, remoteBranch, allowLocal);
                var state = staged.FetchRemote(allowLocal, cancellation);
                if (state.RemoteTip == null) throw new InvalidOperationException("远程没有所选分支，请确认分支名或先由另一位作者推送版本。");
                using (var output = new FileStream(stagedDocument, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    staged.ExtractTo(state.RemoteTip, output);
                DocxFormat.ValidateFile(stagedDocument);
                staged.RefTransaction(null, "create " + BranchRef(remoteBranch) + " " + state.RemoteTip);
                Directory.Move(temporary, destination.root);
                File.Move(stagedDocument, documentPath);
                // Content written by collaborators is treated like a download, so Word opens it
                // in Protected View when the user later opens the file directly.
                ZoneMarker.MarkFromInternet(documentPath);
                return destination;
            }
            finally
            {
                if (File.Exists(stagedDocument)) File.Delete(stagedDocument);
                var full = Path.GetFullPath(temporary);
                if (Path.GetDirectoryName(full) == Path.GetFullPath(parent) &&
                    Regex.IsMatch(Path.GetFileName(full), @"\Aclone-[a-f0-9]{12}\.git\z") && Directory.Exists(full))
                {
                    try
                    {
                        foreach (var file in Directory.GetFiles(full, "*", SearchOption.AllDirectories))
                            File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
                        Directory.Delete(full, true);
                    }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
        }

        private static void RunBranch(GitHistory repository, string branch)
        {
            var reference = BranchRef(branch);
            repository.Run(null, null, "check-ref-format", reference);
            repository.Run(null, null, "symbolic-ref", "HEAD", reference);
        }

        private void NetworkRun(bool allowLocal, CancellationToken cancellation, params string[] command)
        {
            NormalizeRemoteUrl(RemoteUrl, allowLocal);
            var options = new[] { "-c", "core.hooksPath=NUL", "-c", "gc.auto=0", "-c", "maintenance.auto=false",
                "-c", "push.followTags=false", "-c", "push.recurseSubmodules=no", "-c", "remote.origin.mirror=false",
                "-c", "protocol.ext.allow=never", "-c", allowLocal ? "protocol.file.allow=always" : "protocol.file.allow=never" };
            var info = git.CreateInfo(null, false, options.Concat(command));
            // Restore the user's Git credential/proxy configuration only for network operations.
            info.EnvironmentVariables.Remove("GIT_CONFIG_NOSYSTEM");
            info.EnvironmentVariables.Remove("GIT_CONFIG_GLOBAL");
            info.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            ConfigureSsh(info);
            git.RunNetwork(info, cancellation, GitProcess.NetworkIdleTimeout);
        }

        // Keeps a user's own SSH setup (plink, a specific key, a custom known_hosts). Otherwise ssh
        // must never wait for terminal input, and a first connection records GitHub's host key.
        private void ConfigureSsh(System.Diagnostics.ProcessStartInfo info)
        {
            foreach (var name in new[] { "GIT_SSH_COMMAND", "GIT_SSH", "GIT_SSH_VARIANT" })
            {
                var value = Environment.GetEnvironmentVariable(name);
                if (!String.IsNullOrEmpty(value)) info.EnvironmentVariables[name] = value;
            }
            if (info.EnvironmentVariables.ContainsKey("GIT_SSH_COMMAND") || info.EnvironmentVariables.ContainsKey("GIT_SSH")) return;
            var configured = git.InvokeWithUserConfig("config", "--get", "core.sshCommand");
            if (configured.Code == 0 && configured.Bytes.Length > 0) return;
            info.EnvironmentVariables["GIT_SSH_COMMAND"] = "ssh -o BatchMode=yes -o StrictHostKeyChecking=accept-new";
        }
    }

    // Writes the Windows "mark of the web" so Office treats a file like a download.
    internal static class ZoneMarker
    {
        internal static void MarkFromInternet(string path)
        {
            try
            {
                using (var handle = CreateFile(path + ":Zone.Identifier", 0x40000000, 0, IntPtr.Zero, 2, 0x80, IntPtr.Zero))
                {
                    if (handle.IsInvalid) return;
                    using (var stream = new FileStream(handle, FileAccess.Write))
                    {
                        var bytes = System.Text.Encoding.ASCII.GetBytes("[ZoneTransfer]\r\nZoneId=3\r\n");
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        internal static bool IsMarked(string path)
        {
            using (var handle = CreateFile(path + ":Zone.Identifier", 0x80000000, 1, IntPtr.Zero, 3, 0x80, IntPtr.Zero))
                return !handle.IsInvalid;
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
        private static extern Microsoft.Win32.SafeHandles.SafeFileHandle CreateFile(string name, uint access, uint share,
            IntPtr security, uint disposition, uint flags, IntPtr template);
    }
}
