using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace WordGit
{
    internal static class Smoke
    {
        private static void Require(bool condition, string detail)
        {
            if (!condition) throw new InvalidOperationException(detail);
        }

        [STAThread]
        private static void Main()
        {
            Console.OutputEncoding = Encoding.UTF8;
            try { Run(); }
            catch (Exception error)
            {
                Console.Error.WriteLine(error.GetType().Name + ": " + error.Message);
                Console.Error.WriteLine(error.StackTrace);
                Environment.ExitCode = 1;
            }
        }

        private static void Run()
        {
            var root = Path.Combine(Path.GetTempPath(), "WordGitTest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            var first = Path.Combine(root, "测试 文档.docx");
            var second = Path.Combine(root, "第二份.docx");
            var data1 = Fixture("原版：中英文混排");
            var data2 = Fixture("新版：修改了试验条件");
            try
            {
                WordDocuments.WriteNew(first, data1);
                var repo = new GitHistory(first, root);
                Require(repo.Head() == null && repo.History().Count == 0, "Empty repository");
                var initial = repo.Commit(WordDocuments.ReadDocx(first), "初稿", "张三", null);
                var updated = repo.Commit(data2, "导师审阅版", "张三", initial);
                var history = repo.History();
                Require(history.Count == 2 && history[0].Id == updated && history[0].Subject == "导师审阅版" && history[0].Author == "张三", "History or UTF-8 mismatch");
                Require(history[0].Parents.SequenceEqual(new[] { initial }) && history[0].References.Contains("HEAD -> main"), "Git ancestry or refs missing");
                var visualized = CommitGraph.Layout(history, updated);
                Require(visualized[0].IsHead && visualized[0].Edges.Count == 1 && visualized[1].Edges.Count == 1 &&
                    visualized[0].Edges[0].To == visualized[1].Edges[0].From, "Linear graph is disconnected");
                var branched = CommitGraph.Layout(new[] {
                    new Revision { Id = "merge", Parents = new[] { "left", "right" } },
                    new Revision { Id = "left", Parents = new[] { "root" } },
                    new Revision { Id = "right", Parents = new[] { "root" } },
                    new Revision { Id = "root", Parents = new string[0] }
                }, "merge");
                Require(branched[0].Edges.Count == 2 && branched[0].Edges.Select(edge => edge.To).Distinct().Count() == 2 &&
                    branched[2].Edges.Any(edge => edge.Start == 1 && edge.To == branched[3].Node),
                    "Forked graph does not reconnect at the common ancestor");
                Require(repo.Extract(initial).SequenceEqual(data1) && repo.Extract(updated).SequenceEqual(data2), "Git changed the DOCX bytes");
                Require(repo.Commit(data2, "重复提交", "张三", updated) == null, "Duplicate revision");
                try { repo.Commit(data2, "过期写入", "张三", initial); throw new Exception("Stale head accepted"); }
                catch (InvalidOperationException error) { Require(error.Message.Contains("已变化"), "Unexpected stale-head error"); }
                var restored = repo.Commit(repo.Extract(initial), "恢复原版", "张三", updated);
                Require(repo.Extract(restored).SequenceEqual(data1) && repo.History().Count == 3, "Restore did not retain history");
                repo.CreateBranch("review/导师", initial);
                Require(repo.ActiveBranch() == "main" && repo.BranchTip("review/导师") == initial, "Branch start");
                repo.Switch("main", restored, "review/导师", initial);
                Require(repo.ActiveBranch() == "review/导师", "Switch did not move HEAD");
                var side = repo.Commit(data2, "分支修改", "张三", initial, "review/导师");
                Require(repo.BranchTip("main") == restored && repo.MergeBase(restored, side) == initial, "Branch separation or merge base");
                Require(repo.History("review/导师").Select(revision => revision.Id).SequenceEqual(new[] { side, initial }), "Branch history");
                repo.Switch("review/导师", side, "main", restored);
                Require(!repo.Branches(restored).Contains("review/导师"), "Unmerged branch listed as deletable");
                try { repo.DeleteMergedBranch("main", restored, "review/导师", side); throw new Exception("Unmerged branch deleted"); }
                catch (InvalidOperationException error) { Require(error.Message.Contains("未合入"), "Unexpected unmerged branch error"); }
                try { repo.DeleteMergedBranch("main", restored, "main", restored); throw new Exception("Active branch deleted"); }
                catch (InvalidOperationException error) { Require(error.Message.Contains("当前分支"), "Unexpected active branch error"); }
                var merge = repo.Commit(data1, "保留原排版并记录合并", "张三", restored, "main", "review/导师", side);
                Require(repo.History().First().Parents.SequenceEqual(new[] { restored, side }) &&
                    repo.Extract(merge).SequenceEqual(data1) && repo.BranchTip("review/导师") == side, "Merge must retain both parents");
                Require(repo.Branches(merge).Contains("review/导师"), "Merged branch missing from deletion list");
                try { repo.DeleteMergedBranch("main", merge, "review/导师", initial); throw new Exception("Deleted branch after its tip changed"); }
                catch (InvalidOperationException error) { Require(error.Message.Contains("已变化"), "Unexpected stale branch error"); }
                using (repo.Lock()) repo.DeleteMergedBranch("main", merge, "review/导师", side);
                Require(repo.ActiveBranch() == "main" && repo.Head() == merge && repo.BranchTip("review/导师") == null &&
                    repo.History().Any(revision => revision.Id == side) && repo.Extract(side).SequenceEqual(data2),
                    "Deleting a merged branch changed the active branch or lost its commits");

                var saved = true;
                WordDocuments.WriteNew(second, data1);
                var working = new WorkingCopy(second, () => saved, () => repo.SnapshotChecksum(merge));
                Require(working.State() == WorkingCopyState.Clean, "Clean working copy shown as modified");
                saved = false;
                Require(working.State() == WorkingCopyState.Unsaved, "Unsaved edits were not recognized");
                File.WriteAllBytes(second, data2);
                saved = true;
                Require(working.State() == WorkingCopyState.Changed, "Saved but uncommitted edits were not recognized");
                File.WriteAllBytes(second, data1);
                File.SetLastWriteTimeUtc(second, DateTime.UtcNow.AddSeconds(2));
                Require(working.State() == WorkingCopyState.Clean, "Working status did not recover after restoring the committed file");
                Require(new WorkingCopy(second, () => true, null).State() == WorkingCopyState.NoCommit, "Missing first commit was not recognized");
                Require(new GitHistory(second, root).History().Count == 0, "Unrelated file shares history");
                try { WordDocuments.WriteNew(first, data2); throw new Exception("Existing file overwritten"); }
                catch (IOException) { Require(File.ReadAllBytes(first).SequenceEqual(data1), "Existing file modified"); }
                VerifyLegacyFormats(root);
                VerifyCompression(root);
                VerifyStreaming(root);
                Console.WriteLine("PASS: Git history, exact DOCX restoration, branches, merge, working-copy state and lossless history compression.");
            }
            finally { try { Directory.Delete(root, true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
        }

        private static void VerifyStreaming(string root)
        {
            // Streaming chunk planning must produce output byte-identical to the in-memory path.
            var image = new byte[192 * 1024];
            new Random(919).NextBytes(image);
            var samples = new[] { Fixture("流式分块"), Fixture("带图片", image), new byte[] { 1, 2, 3, 4, 5 } };
            foreach (var bytes in samples)
            {
                var chunks = DocumentSnapshot.Split(bytes);
                var objects = DocumentSnapshot.BlobObjects(chunks, "sha256");
                using (var stream = new MemoryStream(bytes, false))
                {
                    var plan = DocumentSnapshot.Plan(stream, "sha256");
                    Require(plan.Length == bytes.Length && plan.Checksum == GitHistory.Sha(bytes) &&
                        plan.Objects.SequenceEqual(objects) &&
                        plan.Sizes.SequenceEqual(chunks.Select(chunk => chunk.Count)) &&
                        plan.Offsets.SequenceEqual(chunks.Select(chunk => (long)chunk.Offset)),
                        "Streaming plan differs from in-memory split");
                    Require(DocumentSnapshot.Manifest(plan.Length, plan.Checksum, plan.Sizes, plan.Objects)
                        .SequenceEqual(DocumentSnapshot.Manifest(bytes, chunks, objects)), "Streaming manifest differs");
                }
            }

            // File-based commits produce the same git trees as memory-based commits.
            var path = Path.Combine(root, "流式提交.docx");
            var bytes1 = Fixture("流式第一版", image);
            WordDocuments.WriteNew(path, bytes1);
            var repo = new GitHistory(path, root);
            var first = repo.CommitFile(path, "流式第一版", "测试", null);
            Require(first != null && repo.Extract(first).SequenceEqual(bytes1), "File commit lost bytes");
            var memoryRepo = new GitHistory(Path.Combine(root, "内存对照.docx"), root);
            var memoryFirst = memoryRepo.Commit(bytes1, "内存版", "测试", null);
            var treeOf = new Func<GitHistory, string, string>((repository, commit) => Encoding.UTF8.GetString(
                CompatibilityGit(repository.RepositoryPath, null, "rev-parse", commit + "^{tree}")).Trim());
            Require(treeOf(repo, first) == treeOf(memoryRepo, memoryFirst),
                "File and memory commits produced different trees");
            Require(repo.CommitFile(path, "无变化", "测试", first) == null, "Unchanged file committed");

            // Streamed extraction restores exact bytes for every committed version.
            var bytes2 = Fixture("流式第二版：修改了试验条件", image);
            File.WriteAllBytes(path, bytes2);
            var second = repo.CommitFile(path, "第二版", "测试", first);
            foreach (var pair in new[] {
                new KeyValuePair<string, byte[]>(first, bytes1), new KeyValuePair<string, byte[]>(second, bytes2) })
                using (var output = new MemoryStream())
                {
                    repo.ExtractTo(pair.Key, output);
                    Require(output.ToArray().SequenceEqual(pair.Value), "Streamed extract changed bytes");
                }

            // A file changed between planning and import must abort with a clear error and no ref change.
            // The plan must contain chunks the repository does not have yet, otherwise import is a no-op.
            var image2 = new byte[192 * 1024];
            new Random(777).NextBytes(image2);
            var bytes3 = Fixture("第三版：又修改了试验方案", image2);
            File.WriteAllBytes(path, bytes3);
            var planned = repo.PlanFile(path);
            var forged = new FileInfo(path).LastWriteTimeUtc;
            File.WriteAllBytes(path, Fixture("被篡改，长度明显不同的另一个版本……", image));
            var rejected = false;
            try { repo.ImportPlan(path, planned); }
            catch (InvalidOperationException error) { Require(error.Message.Contains("被修改"), "Unexpected change error"); rejected = true; }
            Require(rejected, "Modified file was imported");

            // Same-length tampering hidden behind a forged timestamp is caught by the get-mark verification.
            var tampered = (byte[])bytes3.Clone();
            tampered[tampered.Length - 10] ^= 0xFF;
            File.WriteAllBytes(path, tampered);
            File.SetLastWriteTimeUtc(path, forged);
            rejected = false;
            try { repo.ImportPlan(path, planned); }
            catch (InvalidDataException error) { Require(error.Message.Contains("校验失败"), "Unexpected verify error"); rejected = true; }
            Require(rejected, "Forged-timestamp tamper was imported");
            File.WriteAllBytes(path, bytes2);
            Require(repo.Head() == second, "Rejected import changed the branch");

            // Streaming restore rejects truncated or corrupt batch data.
            var manifestBytes = CompatibilityGit(repo.RepositoryPath, null, "cat-file", "blob",
                second + ":" + DocumentSnapshot.ManifestFile);
            var snapshot = DocumentSnapshot.Read(manifestBytes);
            var request = Encoding.UTF8.GetBytes(String.Join("\n",
                snapshot.Objects.Select(id => second + ":" + DocumentSnapshot.ChunksFolder + "/" + id)) + "\n");
            var batch = CompatibilityGit(repo.RepositoryPath, request, "cat-file", "--batch");
            using (var output = new MemoryStream())
            {
                snapshot.Restore(new MemoryStream(batch, false), output);
                Require(output.ToArray().SequenceEqual(bytes2), "Streaming restore changed bytes");
            }
            rejected = false;
            try { snapshot.Restore(new MemoryStream(batch.Take(batch.Length - 1).ToArray(), false), new MemoryStream()); }
            catch (InvalidDataException) { rejected = true; }
            Require(rejected, "Truncated batch accepted");
            var corrupt = (byte[])batch.Clone();
            corrupt[0] = (byte)(corrupt[0] == '0' ? '1' : '0');
            rejected = false;
            try { snapshot.Restore(new MemoryStream(corrupt, false), new MemoryStream()); }
            catch (InvalidDataException) { rejected = true; }
            Require(rejected, "Corrupt batch accepted");

            // A multi-chunk document exercises bounded-memory streaming end to end.
            var bigImage = new byte[3 * 1024 * 1024];
            new Random(313).NextBytes(bigImage);
            var bigBytes = Fixture("大文档流式提交", bigImage);
            var bigPath = Path.Combine(root, "大文档.docx");
            WordDocuments.WriteNew(bigPath, bigBytes);
            var bigRepo = new GitHistory(bigPath, root);
            var bigCommit = bigRepo.CommitFile(bigPath, "大文档", "测试", null);
            var bigOut = Path.Combine(root, "大文档-取回.docx");
            using (var output = new FileStream(bigOut, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                bigRepo.ExtractTo(bigCommit, output);
            Require(File.ReadAllBytes(bigOut).SequenceEqual(bigBytes), "Large document changed through streaming roundtrip");
            Console.WriteLine("PASS: streaming plan, file commit, streamed extract, change detection and corruption rejection.");
        }

        private static void VerifyCompression(string root)
        {
            var path = Path.Combine(root, "图片测试.docx");
            var repo = new GitHistory(path, root);
            var image = new byte[384 * 1024];
            new Random(417).NextBytes(image);
            var versions = new System.Collections.Generic.Dictionary<string, byte[]>();
            string head = null;
            for (var index = 0; index < 8; index++)
            {
                var bytes = Fixture("修改正文 " + index, image);
                head = repo.Commit(bytes, "第 " + (index + 1) + " 版", "WordGit 测试", head);
                versions.Add(head, bytes);
                if (index == 0) repo.CreateBranch("review", head);
            }

            var mainTip = head;
            var first = repo.BranchTip("review");
            repo.Switch("main", mainTip, "review", first);
            var reviewBytes = Fixture("审阅分支独立修改", image);
            var reviewTip = repo.Commit(reviewBytes, "导师审阅", "WordGit 测试", first, "review");
            versions.Add(reviewTip, reviewBytes);
            var working = Fixture("尚未提交的编辑", image);
            WordDocuments.WriteNew(path, working);

            var result = repo.Repack();
            Console.WriteLine("Chunked storage before/after maintenance: " + result.BeforeBytes + " / " + result.AfterBytes);
            Require(result.BeforeBytes < image.Length * 3 && result.AfterBytes < image.Length * 3,
                "Repeated image data was duplicated before maintenance");
            Require(repo.ActiveBranch() == "review" && repo.BranchTip("main") == mainTip && repo.Head() == reviewTip,
                "Repack changed the active branch or branch tips");
            Require(repo.History().Count == versions.Count && versions.All(version => repo.Extract(version.Key).SequenceEqual(version.Value)),
                "A packed DOCX did not restore byte-for-byte");
            Require(File.ReadAllBytes(path).SequenceEqual(working), "Repack changed the working document");
            Console.WriteLine("PASS: all " + versions.Count + " packed versions restored byte-for-byte; Git objects " +
                result.BeforeBytes + " -> " + result.AfterBytes + " bytes.");

            VerifyAutomaticCompression(root, path, repo, image, versions, working);
        }

        private static byte[] CompatibilityGit(string repository, byte[] input, params string[] arguments)
        {
            var quote = (Func<string, string>)Delegate.CreateDelegate(typeof(Func<string, string>),
                typeof(GitHistory).GetMethod("Quote", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static));
            var info = new System.Diagnostics.ProcessStartInfo("git", "-c core.longpaths=true --git-dir=. " + String.Join(" ", arguments.Select(quote))) {
                WorkingDirectory = repository, UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
            };
            info.EnvironmentVariables["GIT_INDEX_FILE"] = "compatibility.index";
            info.EnvironmentVariables["GIT_CONFIG_NOSYSTEM"] = "1";
            info.EnvironmentVariables["GIT_CONFIG_GLOBAL"] = "NUL";
            info.EnvironmentVariables["GIT_AUTHOR_NAME"] = "Compatibility test";
            info.EnvironmentVariables["GIT_AUTHOR_EMAIL"] = "test@wordgit.local";
            info.EnvironmentVariables["GIT_COMMITTER_NAME"] = "Compatibility test";
            info.EnvironmentVariables["GIT_COMMITTER_EMAIL"] = "test@wordgit.local";
            using (var process = System.Diagnostics.Process.Start(info))
            using (var output = new MemoryStream())
            {
                var reading = process.StandardOutput.BaseStream.CopyToAsync(output);
                var error = process.StandardError.ReadToEndAsync();
                if (input != null) process.StandardInput.BaseStream.Write(input, 0, input.Length);
                process.StandardInput.Close();
                process.WaitForExit();
                System.Threading.Tasks.Task.WaitAll(reading, error);
                Require(process.ExitCode == 0, "Compatibility fixture Git failed: " + error.Result);
                return output.ToArray();
            }
        }

        private static string CompatibilityCommit(GitHistory repository, Dictionary<string, byte[]> files, string parent, bool advance)
        {
            var path = repository.RepositoryPath;
            CompatibilityGit(path, null, "read-tree", "--empty");
            foreach (var file in files)
            {
                var blob = Encoding.UTF8.GetString(CompatibilityGit(path, file.Value, "hash-object", "-w", "--stdin")).Trim();
                CompatibilityGit(path, null, "update-index", "--add", "--cacheinfo", "100644", blob, file.Key);
            }
            var tree = Encoding.UTF8.GetString(CompatibilityGit(path, null, "write-tree")).Trim();
            var id = Encoding.UTF8.GetString(CompatibilityGit(path, null, "commit-tree", tree, "-p", parent, "-m", "historical format")).Trim();
            if (advance) CompatibilityGit(path, null, "update-ref", "refs/heads/main", id, parent);
            return id;
        }

        private static void VerifyLegacyFormats(string root)
        {
            var repository = new GitHistory(Path.Combine(root, "跨格式兼容.docx"), root);
            var firstBytes = Fixture("整包旧版");
            var seed = repository.Commit(firstBytes, "初始化", "测试", null);
            var whole = CompatibilityCommit(repository, new Dictionary<string, byte[]> { { "document.docx", firstBytes } }, seed, true);
            Require(repository.Extract(whole).SequenceEqual(firstBytes), "Whole DOCX compatibility");
            using (var output = new MemoryStream())
            {
                repository.ExtractTo(whole, output);
                Require(output.ToArray().SequenceEqual(firstBytes), "Whole DOCX changed through streamed extract");
            }
            var original = Fixture("旧拆包版：中英文");
            var parts = DocxFormat.Read(original);
            parts.Add(new DocxPart { Name = "word/media/图片 数据.bin", Bytes = new byte[] { 0, 10, 13, 255, 3, 7 } });
            var meta = DocxFormat.MetaText(DocxFormat.Fingerprint(parts), parts);
            var files = parts.ToDictionary(part => part.Name, part => part.Bytes);
            files.Add(DocxFormat.MetaName, Encoding.UTF8.GetBytes(meta));
            var expanded = CompatibilityCommit(repository, files, whole, true);
            var extracted = repository.Extract(expanded);
            var actual = DocxFormat.Read(extracted).ToDictionary(part => part.Name);
            Require(actual.Count == parts.Count && parts.All(part => actual[part.Name].Bytes.SequenceEqual(part.Bytes)),
                "Expanded historical parts changed");
            using (var output = new MemoryStream())
            {
                repository.ExtractTo(expanded, output);
                Require(output.ToArray().SequenceEqual(extracted), "Expanded DOCX changed through streamed extract");
            }
            var nextBytes = Fixture("新分块版");
            var next = repository.Commit(nextBytes, "继续提交", "测试", expanded);
            Require(repository.Extract(next).SequenceEqual(nextBytes) && repository.History().Any(item => item.Id == expanded),
                "A new commit lost expanded history");
            var corruptFiles = new Dictionary<string, byte[]>(files);
            corruptFiles[DocxFormat.MetaName] = Encoding.UTF8.GetBytes(meta.Replace(GitHistory.Sha(parts[0].Bytes), new string('0', 64)));
            var corrupt = CompatibilityCommit(repository, corruptFiles, next, false);
            var rejected = false;
            try { repository.Extract(corrupt); } catch (InvalidDataException) { rejected = true; }
            Require(rejected && repository.Head() == next, "Corrupt legacy parts were accepted or refs changed");
            var missingFiles = new Dictionary<string, byte[]>(files);
            missingFiles.Remove(parts[0].Name);
            var missing = CompatibilityCommit(repository, missingFiles, next, false);
            rejected = false;
            try { repository.Extract(missing); } catch (InvalidDataException) { rejected = true; }
            Require(rejected, "Missing legacy part was ignored");
            repository.Repack();
            Require(repository.Extract(whole).SequenceEqual(firstBytes) &&
                DocxFormat.Fingerprint(repository.Extract(expanded)).SequenceEqual(DocxFormat.Fingerprint(parts)) &&
                repository.Extract(next).SequenceEqual(nextBytes), "History maintenance changed mixed-format revisions");
            Console.WriteLine("PASS: whole DOCX, legacy expanded parts and chunked revisions in one history, migration, corruption detection and maintenance.");
        }

        private static void VerifyAutomaticCompression(string root, string path, GitHistory repo, byte[] image,
            System.Collections.Generic.Dictionary<string, byte[]> versions, byte[] working)
        {
            var threshold = HistoryMaintenance.CommitThreshold;
            var savedTips = Path.Combine(repo.RepositoryPath, "pack-tips");
            var previousTips = File.ReadAllText(savedTips);
            var branch = repo.ActiveBranch();
            var head = repo.Head();
            using (var packing = new FileStream(Path.Combine(repo.RepositoryPath, "wordgit-repack.lock"), FileMode.OpenOrCreate,
                FileAccess.ReadWrite, FileShare.None))
            {
                var saved = Fixture("打包锁持有期间提交", image);
                using (repo.Lock()) head = repo.Commit(saved, "打包期间修改", "WordGit 测试", head, branch);
                versions.Add(head, saved);
                Require(!repo.TryAutoRepack(threshold, () => true), "Two repacks can acquire the same repository");
            }

            // Reopen the repository to check that the threshold survives a Word restart.
            repo = new GitHistory(path, root);
            for (var i = 1; i < threshold; i++)
            {
                var saved = Fixture("自动打包测试 " + i, image);
                using (repo.Lock()) head = repo.Commit(saved, "自动打包测试", "WordGit 测试", head, branch);
                versions.Add(head, saved);
                if (i < threshold - 1)
                    Require(repo.TryAutoRepack(threshold, () => { throw new InvalidOperationException("Repack started below the threshold"); }),
                        "A check below the threshold deferred unexpectedly");
            }
            Require(File.ReadAllText(savedTips) == previousTips && !repo.TryAutoRepack(threshold, () => false),
                "The threshold check did not defer until idle");

            var packedHead = head;
            var late = Fixture("打包开始时的新提交", image);
            Require(repo.TryAutoRepack(threshold, () =>
            {
                using (repo.Lock()) head = repo.Commit(late, "打包启动时的新提交", "WordGit 测试", head, branch);
                versions.Add(head, late);
                return true;
            }), "Automatic repack did not run once idle");
            var tips = File.ReadAllLines(savedTips);
            Require(tips.Contains(packedHead) && !tips.Contains(head) && repo.Head() == head,
                "A commit made during packing was lost or incorrectly counted as covered");
            Require(repo.TryAutoRepack(threshold, () => { throw new InvalidOperationException("Repack did not reset its count"); }),
                "New count after repack was incorrect");
            Require(versions.All(version => repo.Extract(version.Key).SequenceEqual(version.Value)) &&
                File.ReadAllBytes(path).SequenceEqual(working), "Automatic repack changed the work copy or a historical version");
            // Simulate reopening Word before the pending versions had been packed.
            File.WriteAllText(savedTips, previousTips);
            var previousTime = DateTime.UtcNow.AddMinutes(-1);
            File.SetLastWriteTimeUtc(savedTips, previousTime);
            using (var maintenance = new HistoryMaintenance(root, () => true))
                Require(System.Threading.SpinWait.SpinUntil(() => File.GetLastWriteTimeUtc(savedTips) > previousTime.AddSeconds(1), 20000),
                    "The background timer did not pick up the pending repository");
            Require(File.ReadAllLines(savedTips).Contains(head) &&
                versions.All(version => repo.Extract(version.Key).SequenceEqual(version.Value)),
                "The background task changed or lost a DOCX version");
            Console.WriteLine("PASS: automatic repack threshold, idle deferral, startup timer, concurrent commits and " + versions.Count + " exact DOCX versions.");
        }

        private static byte[] Fixture(string text, byte[] image = null)
        {
            using (var buffer = new MemoryStream())
            {
                using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, true))
                {
                    using (var writer = new StreamWriter(archive.CreateEntry("[Content_Types].xml").Open())) writer.Write("<Types/>");
                    using (var writer = new StreamWriter(archive.CreateEntry("word/document.xml").Open())) writer.Write("<document>" + text + "</document>");
                    if (image != null)
                        using (var picture = archive.CreateEntry("word/media/image1.png").Open()) picture.Write(image, 0, image.Length);
                }
                return buffer.ToArray();
            }
        }
    }
}
