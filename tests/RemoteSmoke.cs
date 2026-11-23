using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordGit
{
    internal static class RemoteSmoke
    {
        private static void Require(bool condition, string message)
        { if (!condition) throw new InvalidOperationException(message); }

        private static string Git(string folder, params string[] args)
        {
            var info = new ProcessStartInfo(GitProcess.FindGit(), String.Join(" ", args.Select(arg => "\"" + arg.Replace("\"", "\\\"") + "\"")))
            { WorkingDirectory = folder, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8 };
            info.EnvironmentVariables["GIT_CONFIG_NOSYSTEM"] = "1";
            info.EnvironmentVariables["GIT_CONFIG_GLOBAL"] = "NUL";
            using (var process = Process.Start(info))
            {
                var output = process.StandardOutput.ReadToEndAsync();
                var error = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(30000)) { process.Kill(); throw new TimeoutException(); }
                Task.WaitAll(output, error);
                if (process.ExitCode != 0) throw new Exception(error.Result);
                return output.Result.Trim();
            }
        }

        private static byte[] Fixture(string text)
        {
            using (var output = new MemoryStream())
            {
                using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
                {
                    using (var writer = new StreamWriter(zip.CreateEntry("[Content_Types].xml").Open())) writer.Write("<Types/>");
                    using (var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open())) writer.Write("<document>" + text + "</document>");
                }
                return output.ToArray();
            }
        }

        private static void Main()
        {
            Console.OutputEncoding = Encoding.UTF8;
            var root = Path.Combine(Path.GetTempPath(), "wg-r-" + Guid.NewGuid().ToString("N").Substring(0,12));
            Directory.CreateDirectory(root);
            try { using (TestConsole.Utf8InputWithPreamble()) Run(root); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.ExitCode = 1; }
            finally {
                foreach (var file in Directory.GetFiles(root, "*", SearchOption.AllDirectories))
                    File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
                Directory.Delete(root, true);
            }
        }

        private static void VerifyView(GitHistory repo, string phase)
        {
            var expected = repo.SyncState();
            var actual = repo.ReadView();
            var state = actual.Remote;
            Require(state.Url == expected.Url && state.LocalBranch == expected.LocalBranch &&
                state.RemoteBranch == expected.RemoteBranch && state.RemoteReference == expected.RemoteReference &&
                state.LocalTip == expected.LocalTip && state.RemoteTip == expected.RemoteTip &&
                state.Ahead == expected.Ahead && state.Behind == expected.Behind && state.Unrelated == expected.Unrelated,
                "Batched pane state differs: " + phase);
            Require(actual.Branches.SequenceEqual(repo.Branches().Concat(repo.RemoteBranches())),
                "Batched branch list differs: " + phase);
            var history = repo.History();
            Require(actual.History.Select(row => row.Id).SequenceEqual(history.Select(row => row.Id)) &&
                actual.History.Zip(history, (left, right) => left.Parents.SequenceEqual(right.Parents) &&
                    left.References.SequenceEqual(right.References) && left.Message == right.Message).All(same => same),
                "Batched history differs: " + phase);
        }

        private static void Run(string root)
        {
            Require(GitHistory.NormalizeRemoteUrl("https://github.com/owner/paper") == "https://github.com/owner/paper.git", "HTTPS normalization");
            Require(GitHistory.NormalizeRemoteUrl("ssh://git@github.com/owner/paper.git") == "git@github.com:owner/paper.git", "SSH normalization");
            foreach (var url in new[] { "https://user:secret@github.com/a/b", "ext::run", "file:///x", "https://example.com/a/b", "https://github.com/a/b?token=x" })
            {
                var blocked = false;
                try { GitHistory.NormalizeRemoteUrl(url); } catch (ArgumentException) { blocked = true; }
                Require(blocked, "Unexpected remote protocol accepted");
            }
            var server = Path.Combine(root, "server.git");
            Git(root, "init", "--bare", "--initial-branch=main", server);
            var aPath = Path.Combine(root, "作者甲.docx");
            var bPath = Path.Combine(root, "另一台电脑.docx");
            var aBytes = Fixture("送审 A");
            File.WriteAllBytes(aPath, aBytes);
            var a = new GitHistory(aPath, Path.Combine(root, "a"));
            VerifyView(a, "empty");
            var first = a.CommitFile(aPath, "初稿", "甲", null);
            VerifyView(a, "local");
            a.CreateBranch("保留的本地分支", first);
            a.ConfigureRemote(server, "main", true);
            Require(a.Head() == first && a.Extract(first).SequenceEqual(aBytes), "Link changed existing history");
            var state = a.PushRemote(true);
            Require(state.Ahead == 0 && state.Behind == 0 && state.RemoteTip == first, "Initial push did not update tracking state");
            VerifyView(a, "equal remote tips");
            var b = GitHistory.CloneRemote(server, "main", bPath, Path.Combine(root, "b"), true);
            Require(b.Head() == first && File.ReadAllBytes(bPath).SequenceEqual(aBytes), "Clone changed commit IDs or document bytes");
            Require(ZoneMarker.IsMarked(bPath) && !ZoneMarker.IsMarked(aPath), "Cloned document was not marked as downloaded content");
            var bBytes = Fixture("乙新增");
            File.WriteAllBytes(bPath, bBytes);
            var second = b.CommitFile(bPath, "乙修改", "乙", first);
            b.PushRemote(true);
            var backupsBefore = a.RemoteBackupNamespaces().Length;
            state = a.FetchRemote(true);
            Require(a.RemoteBackupNamespaces().Length == backupsBefore, "A fast-forward fetch created a remote backup");
            Require(a.Head() == first && File.ReadAllBytes(aPath).SequenceEqual(aBytes) && state.Ahead == 0 && state.Behind == 1,
                "Fetch overwrote the current document or computed wrong behind count");
            VerifyView(a, "behind");
            Require(a.History(state.RemoteReference).First().Id == second &&
                a.History().Any(row => row.References.Contains("origin/main")), "Remote branches are missing from history");
            File.WriteAllBytes(aPath, bBytes);
            a.FastForwardRemote("main", first, state.RemoteReference, second);
            Require(a.Head() == second && a.Extract(first).SequenceEqual(aBytes), "Fast-forward lost the previous version");
            var localBytes = Fixture("甲继续修改");
            File.WriteAllBytes(aPath, localBytes);
            var local = a.CommitFile(aPath, "甲独立修改", "甲", second);
            var remoteBytes = Fixture("乙继续修改");
            File.WriteAllBytes(bPath, remoteBytes);
            var remote = b.CommitFile(bPath, "乙独立修改", "乙", second);
            b.PushRemote(true);
            var refused = false;
            try { a.PushRemote(true); } catch (InvalidOperationException) { refused = true; }
            Require(refused && Git(root, "--git-dir=" + server, "rev-parse", "refs/heads/main") == remote &&
                a.Head() == local && File.ReadAllBytes(aPath).SequenceEqual(localBytes), "Rejected push overwrote either side");
            state = a.FetchRemote(true);
            Require(state.Ahead == 1 && state.Behind == 1 && !state.Unrelated &&
                a.MergeBase(local, remote) == second, "Divergence or merge base is wrong");
            VerifyView(a, "divergence");
            var mergedBytes = Fixture("甲乙确认的合并稿");
            File.WriteAllBytes(aPath, mergedBytes);
            var merged = a.CommitFile(aPath, "确认远程合并", "甲", local, "main", state.RemoteReference, remote);
            Require(a.History().Single(row => row.Id == merged).Parents.SequenceEqual(new[] { local, remote }), "Remote merge did not keep two parents");
            a.PushRemote(true);
            state = b.FetchRemote(true);
            File.WriteAllBytes(bPath, mergedBytes);
            b.FastForwardRemote("main", remote, state.RemoteReference, merged);
            Require(b.Head() == merged, "Peer could not download the merge commit");

            b.CreateBranch("review/remote-only", first);
            b.Switch("main", merged, "review/remote-only", first);
            var sideBytes = Fixture("尚未合并的远程版本");
            File.WriteAllBytes(bPath, sideBytes);
            var side = b.CommitFile(bPath, "仅远程分支", "乙", first);
            b.PushRemote(true);
            a.FetchRemote(true);
            Git(root, "--git-dir=" + server, "update-ref", "refs/heads/review/remote-only", first, side);
            var beforeRewrite = a.RemoteBackupNamespaces().Length;
            a.FetchRemote(true);
            Require(a.RemoteBackupNamespaces().Length == beforeRewrite + 1, "A forced remote update was not archived");
            Require(a.History().Any(row => row.Id == side) && a.Extract(side).SequenceEqual(sideBytes) &&
                a.History().All(row => row.References.All(reference => !reference.StartsWith("refs/wordgit/"))),
                "Remote rewrite erased previously fetched history or leaked archive labels");
            Git(root, "--git-dir=" + server, "update-ref", "-d", "refs/heads/review/remote-only", first);
            a.FetchRemote(true);
            Require(a.BranchTip(GitHistory.RemotePrefix + "review/remote-only") == first, "Fetch pruned old remote history");

            b.Switch("review/remote-only", side, "main", merged);
            File.WriteAllBytes(bPath, Fixture("远程最新 N"));
            var next = b.CommitFile(bPath, "远程 N", "乙", merged);
            b.PushRemote(true);
            state = a.FetchRemote(true);
            File.WriteAllBytes(aPath, Fixture("本地并发 Z"));
            var concurrent = a.CommitFile(aPath, "本地并发", "甲", merged);
            refused = false;
            try { a.FastForwardRemote("main", merged, state.RemoteReference, next); } catch (InvalidOperationException) { refused = true; }
            Require(refused && a.Head() == concurrent, "Stale fast-forward overwrote a new local commit");
            File.WriteAllBytes(bPath, Fixture("远程又更新 N2"));
            var next2 = b.CommitFile(bPath, "远程 N2", "乙", next);
            b.PushRemote(true);
            a.FetchRemote(true);
            refused = false;
            try { a.Commit(Fixture("过期合并"), "过期", "甲", concurrent, "main", GitHistory.RemotePrefix + "main", next); }
            catch (InvalidOperationException) { refused = true; }
            Require(refused && a.Head() == concurrent && a.BranchTip(GitHistory.RemotePrefix + "main") == next2,
                "Stale remote review was committed");

            var unrelatedPath = Path.Combine(root, "无关稿件.docx");
            File.WriteAllBytes(unrelatedPath, Fixture("另一份稿件"));
            var unrelated = new GitHistory(unrelatedPath, Path.Combine(root, "u"));
            unrelated.CommitFile(unrelatedPath, "无关初稿", "丙", null);
            unrelated.ConfigureRemote(server, "unrelated", true);
            unrelated.PushRemote(true);
            a.ConfigureRemote(server, "unrelated", true);
            Require(a.FetchRemote(true).Unrelated && a.Head() == concurrent, "Unrelated history was treated as mergeable");
            VerifyView(a, "unrelated tracking branch");
            a.ConfigureRemote(server, "main", true);

            var before = File.ReadAllBytes(aPath);
            refused = false;
            try { GitHistory.CloneRemote(server, "main", aPath, Path.Combine(root,"a"), true); } catch (IOException) { refused = true; }
            Require(refused && File.ReadAllBytes(aPath).SequenceEqual(before), "Clone overwrote an existing document");
            var reserved = Path.Combine(root, "已有历史的新文件名.docx");
            var reservedRepo = new GitHistory(reserved, Path.Combine(root,"reserved"));
            Directory.CreateDirectory(reservedRepo.RepositoryPath);
            File.WriteAllText(Path.Combine(reservedRepo.RepositoryPath, "sentinel"), "keep");
            refused = false;
            try { GitHistory.CloneRemote(server, "main", reserved, Path.Combine(root,"reserved"), true); } catch (IOException) { refused = true; }
            Require(refused && File.Exists(Path.Combine(reservedRepo.RepositoryPath, "sentinel")), "Clone replaced an existing repository");
            var missing = Path.Combine(root, "不存在的远程分支.docx");
            refused = false;
            try { GitHistory.CloneRemote(server, "missing", missing, Path.Combine(root,"missing"), true); } catch (InvalidOperationException) { refused = true; }
            Require(refused && !File.Exists(missing), "Missing branch published a document");
            var secondServer = Path.Combine(root, "second-server.git");
            Git(root, "init", "--bare", "--initial-branch=main", secondServer);
            a.ConfigureRemote(secondServer, "main", true);
            Require(a.RemoteBranches().Length == 0 && a.Head() == concurrent &&
                a.History().Any(row => row.Id == side) && a.Extract(side).SequenceEqual(sideBytes),
                "Changing remote address lost old history or retained misleading tracking branches");
            Require(a.RemoteBackupNamespaces().Length >= 1,
                "URL change did not archive the previous remote tips");
            a.ConfigureRemote(server, "main", true);
            var cancelled = new System.Threading.CancellationTokenSource();
            cancelled.Cancel();
            var cancelRejected = false;
            try { a.FetchRemote(true, cancelled.Token); } catch (OperationCanceledException) { cancelRejected = true; }
            Require(cancelRejected && a.Head() == concurrent, "Cancelled fetch was not reported or changed the local branch");
            a.FetchRemote(true);
            for (var index = 0; index < 25; index++)
                Git(root, "--git-dir=" + a.RepositoryPath, "update-ref",
                    "refs/wordgit/remote-backup/legacy" + index.ToString("00") + "/main", first);
            Require(a.RemoteBackupNamespaces().Contains("legacy00"), "Test failed to plant legacy backup refs");
            a.FetchRemote(true);
            Require(!a.RemoteBackupNamespaces().Contains("legacy00") && a.RemoteBackupNamespaces().Length <= 20,
                "Legacy remote backups were not cleaned");
            a.Repack();
            VerifyView(a, "packed refs");
            Require(a.Extract(side).SequenceEqual(sideBytes) && a.Extract(first).SequenceEqual(aBytes) &&
                a.Branches().Contains("保留的本地分支"), "Remote operations/repack lost existing versions or branches");
            Console.WriteLine("PASS: two clients, link/clone/push/fetch/fast-forward/divergent merge, remote graph/counts, refused push, concurrent changes, unrelated history, fetch cancellation, archive retention and existing-history protection.");
        }
    }
}
