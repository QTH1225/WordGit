using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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
                LocalTip = Head(), RemoteReference = RemotePrefix + branch };
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
            var prefix = "refs/wordgit/remote-backup/" + Guid.NewGuid().ToString("N") + "/";
            RefTransaction(null, refs.Select(reference =>
                "create " + prefix + reference.Substring(RemotePrefix.Length) + " " + BranchTip(reference)).ToArray());
        }

        internal RemoteState FetchRemote(bool allowLocal = false)
        {
            NormalizeRemoteUrl(RemoteUrl, allowLocal);
            ArchiveRemoteTips();
            NetworkRun(allowLocal, "fetch", "--no-prune", "--no-tags", "--no-recurse-submodules", "origin",
                "+refs/heads/*:refs/remotes/origin/*");
            return SyncState();
        }

        internal RemoteState PushRemote(bool allowLocal = false)
        {
            var state = SyncState();
            if (!state.Linked || state.LocalTip == null) throw new InvalidOperationException("请先关联仓库并提交当前稿。");
            NormalizeRemoteUrl(state.Url, allowLocal);
            // A normal push refuses a non-fast-forward update on the server.
            NetworkRun(allowLocal, "push", "--porcelain", "origin", BranchRef(state.LocalBranch) + ":" + BranchRef(state.RemoteBranch));
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
            string storage = null, bool allowLocal = false)
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
                var staged = new GitHistory(temporary, destination.git, null);
                staged.EnsureRepository();
                RunBranch(staged, remoteBranch);
                staged.ConfigureRemote(url, remoteBranch, allowLocal);
                var state = staged.FetchRemote(allowLocal);
                if (state.RemoteTip == null) throw new InvalidOperationException("远程没有所选分支，请确认分支名或先由另一位作者推送版本。");
                WordDocuments.ExtractToFile(staged, state.RemoteTip, stagedDocument);
                WordDocuments.ValidateDocx(stagedDocument);
                staged.RefTransaction(null, "create " + BranchRef(remoteBranch) + " " + state.RemoteTip);
                Directory.Move(temporary, destination.root);
                File.Move(stagedDocument, documentPath);
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

        private void NetworkRun(bool allowLocal, params string[] command)
        {
            NormalizeRemoteUrl(RemoteUrl, allowLocal);
            var options = new[] { "-c", "core.hooksPath=NUL", "-c", "gc.auto=0", "-c", "maintenance.auto=false",
                "-c", "push.followTags=false", "-c", "push.recurseSubmodules=no", "-c", "remote.origin.mirror=false",
                "-c", "protocol.ext.allow=never", "-c", allowLocal ? "protocol.file.allow=always" : "protocol.file.allow=never" };
            var info = CreateInfo(null, false, options.Concat(command).ToArray());
            // Restore the user's Git credential/proxy configuration only for network operations.
            info.EnvironmentVariables.Remove("GIT_CONFIG_NOSYSTEM");
            info.EnvironmentVariables.Remove("GIT_CONFIG_GLOBAL");
            info.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            info.EnvironmentVariables["GIT_SSH_COMMAND"] = "ssh -o BatchMode=yes";
            using (var process = Process.Start(info))
            {
                var output = process.StandardOutput.ReadToEndAsync();
                var errors = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(180000))
                {
                    process.Kill();
                    process.WaitForExit();
                    throw new TimeoutException("远程 Git 操作超时，请检查网络或完成 Git for Windows 登录后重试。");
                }
                Task.WaitAll(output, errors);
                if (process.ExitCode != 0)
                {
                    var detail = errors.Result.Trim();
                    if (detail.Length > 1600) detail = detail.Substring(detail.Length - 1600);
                    throw new InvalidOperationException("远程 Git 操作未完成。推送被拒绝时请先获取并拉取；认证失败时请完成 Git for Windows 登录。\n" + detail);
                }
            }
        }
    }
}
