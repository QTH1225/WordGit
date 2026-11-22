using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordGit
{
    internal sealed class Revision
    {
        public string Id, Author, Subject, Message;
        public string[] Parents, References;
        public DateTimeOffset Date;
        public override string ToString() { return Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm") + "   " + Id.Substring(0, 8) + "   " + Subject + "   · " + Author; }
    }

    internal sealed class RepackResult
    {
        public readonly long BeforeBytes, AfterBytes;

        public RepackResult(long beforeBytes, long afterBytes)
        {
            BeforeBytes = beforeBytes;
            AfterBytes = afterBytes;
        }
    }

    internal sealed partial class GitHistory
    {
        private const string FileName = "document.docx";
        private const string Heads = "refs/heads/";
        private static readonly Encoding Utf8 = new UTF8Encoding(false, true);
        private readonly string root, git;
        private readonly Action<string> commitCreated;
        public string RepositoryPath { get { return root; } }

        public GitHistory(string documentPath, string storage = null, string executable = null, Action<string> commitCreated = null)
            : this(Path.Combine(storage ?? AppPaths.Repositories, Sha(Utf8.GetBytes(Path.GetFullPath(documentPath).ToUpperInvariant())) + ".git"),
                executable, commitCreated)
        {
        }

        private GitHistory(string repositoryPath, string executable, Action<string> commitCreated)
        {
            root = repositoryPath;
            git = executable ?? FindGit();
            this.commitCreated = commitCreated;
        }

        internal static GitHistory OpenRepository(string repositoryPath)
        {
            return new GitHistory(repositoryPath, null, (Action<string>)null);
        }

        public static string Sha(byte[] bytes)
        {
            using (var algorithm = SHA256.Create()) return BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }

        private static string FindGit()
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

        public string ActiveBranch()
        {
            if (!File.Exists(Path.Combine(root, "HEAD"))) return "main";
            var reference = Text(Run(null, null, "symbolic-ref", "HEAD"));
            if (!reference.StartsWith(Heads, StringComparison.Ordinal)) throw new InvalidOperationException("HEAD 未指向本地分支。");
            return reference.Substring(Heads.Length);
        }

        public string Head() { return BranchTip(ActiveBranch()); }

        public string BranchTip(string branch)
        {
            if (!File.Exists(Path.Combine(root, "HEAD"))) return null;
            var result = Invoke(null, null, "rev-parse", "--verify", "-q", HistoryReference(branch));
            if (result.Code == 1) return null;
            return Id(Text(result));
        }

        public string[] Branches(string mergedInto = null)
        {
            if (!File.Exists(Path.Combine(root, "HEAD"))) return new string[0];
            var args = new List<string> { "for-each-ref", "--sort=refname", "--format=%(refname:strip=2)" };
            if (mergedInto != null) args.Add("--merged=" + Id(mergedInto));
            args.Add(Heads);
            return Text(Run(null, null, args.ToArray())).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public void CreateBranch(string name, string revision)
        {
            var reference = BranchRef(name);
            Run(null, null, "check-ref-format", reference);
            if (Branches().Any(branch => String.Equals(branch, name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("同名分支已存在。");
            revision = Id(revision);
            Extract(revision);
            Run(null, null, "update-ref", reference, revision, new string('0', revision.Length));
        }

        public string ImportReviewFile(string path, string basis, string newBranch, string reviewer,
            string targetBranch, string expectedTarget)
        {
            if (String.IsNullOrWhiteSpace(reviewer) || reviewer.IndexOf('\0') >= 0)
                throw new ArgumentException("请填写返稿来源。");
            EnsureRepository();
            var reference = BranchRef(newBranch);
            Run(null, null, "check-ref-format", reference);
            if (Branches().Any(branch => String.Equals(branch, newBranch, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("同名分支已存在，请换一个返稿分支名。");
            basis = Id(basis);
            if (expectedTarget == null || ActiveBranch() != targetBranch || BranchTip(targetBranch) != expectedTarget)
                throw new InvalidOperationException("当前分支或提交已变化，请刷新后重新导入。");
            ExtractTo(basis, Stream.Null);
            WordDocuments.ValidateDocx(path);
            var prepared = PrepareSnapshotFile(path);
            if (ActiveBranch() != targetBranch || BranchTip(targetBranch) != expectedTarget)
                throw new InvalidOperationException("导入期间当前分支已变化，返稿分支尚未创建，请重新导入。");
            var identity = Regex.Replace(reviewer, @"[<>\x00-\x1f]", "").Trim();
            if (identity.Length == 0) throw new ArgumentException("请填写有效的返稿来源。");
            var message = "导入审阅稿：" + Path.GetFileName(path) + "\n\nWordGit-Review-Base: " + basis +
                "\nWordGit-Review-SHA256: " + prepared.Checksum + "\n";
            var commit = Id(Text(Run(Utf8.GetBytes(message), identity, "commit-tree", prepared.Tree, "-p", basis, "-F", "-")));
            RefTransaction(identity, "create " + reference + " " + commit,
                "verify " + BranchRef(targetBranch) + " " + Id(expectedTarget));
            if (commitCreated != null) commitCreated(root);
            return commit;
        }

        internal string ImportedReviewBase(string revision)
        {
            var text = Text(Run(null, null, "show", "-s", "--format=%P%n%B", Id(revision)));
            var lines = text.Split('\n');
            var marker = lines.Where(line => line.StartsWith("WordGit-Review-Base: ", StringComparison.Ordinal)).ToArray();
            if (marker.Length != 1) return null;
            var parent = marker[0].Substring("WordGit-Review-Base: ".Length).Trim();
            return Regex.IsMatch(parent, @"\A(?:[a-f0-9]{40}|[a-f0-9]{64})\z") && lines[0].Trim() == parent ? parent : null;
        }

        public void Switch(string from, string oldTip, string to, string newTip)
        {
            var currentRef = BranchRef(from);
            var targetRef = BranchRef(to);
            if (ActiveBranch() != from || BranchTip(from) != oldTip)
                throw new InvalidOperationException("当前分支已变化，请刷新后重试。");
            RefTransaction(null, "verify " + targetRef + " " + Id(newTip),
                "option no-deref", "symref-update HEAD " + targetRef + " ref " + currentRef);
        }

        public void DeleteMergedBranch(string current, string head, string target, string targetTip)
        {
            var active = ActiveBranch();
            if (String.Equals(current, target, StringComparison.OrdinalIgnoreCase) ||
                String.Equals(active, target, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("不能删除当前分支。");
            if (active != current || BranchTip(current) != head || BranchTip(target) != targetTip)
                throw new InvalidOperationException("分支或版本历史已变化，请刷新后重试。");

            var result = Invoke(null, null, "merge-base", "--is-ancestor", Id(targetTip), Id(head));
            if (result.Code == 1)
                throw new InvalidOperationException("该分支还有提交未合入当前分支。");
            Checked(result);

            // Git locks HEAD and its referent. If HEAD now points to the target, the two operations conflict.
            RefTransaction(null, "verify HEAD " + head,
                "option no-deref", "delete " + BranchRef(target) + " " + targetTip);
        }

        public IDisposable Lock()
        {
            Directory.CreateDirectory(root);
            try { return new FileStream(Path.Combine(root, "wordgit.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) { throw new InvalidOperationException("另一处 WordGit 正在处理这个文档，请稍后重试。"); }
        }

        public RepackResult Repack()
        {
            using (var packing = TryRepackLock())
            {
                if (packing == null) throw new InvalidOperationException("此文档的历史正在压缩。");
                var tips = RepackTips();
                if (tips.Length == 0) throw new InvalidOperationException("尚无提交可压缩。");

                var before = ObjectBytes();
                Repack(tips, false);
                return new RepackResult(before, ObjectBytes());
            }
        }

        internal bool TryAutoRepack(int threshold, Func<bool> canStart)
        {
            if (!File.Exists(Path.Combine(root, "HEAD"))) return true;
            using (var packing = TryRepackLock())
            {
                if (packing == null) return false;
                var tips = RepackTips();
                if (tips.Length == 0) return true;

                var excluded = File.Exists(PackedTipsPath)
                    ? File.ReadAllLines(PackedTipsPath, Utf8).Select(tip => "^" + Id(tip))
                    : Enumerable.Empty<string>();
                var input = Utf8.GetBytes(String.Join("\n", tips.Concat(excluded)) + "\n");
                var count = Text(Run(input, null, "rev-list", "--count", "--ignore-missing",
                    "--max-count=" + threshold.ToString(System.Globalization.CultureInfo.InvariantCulture), "--stdin"));
                if (Int32.Parse(count, System.Globalization.CultureInfo.InvariantCulture) < threshold) return true;
                if (!canStart()) return false;

                Repack(tips, true);
                return true;
            }
        }

        private string PackedTipsPath { get { return Path.Combine(root, "pack-tips"); } }

        private FileStream TryRepackLock()
        {
            try { return new FileStream(Path.Combine(root, "wordgit-repack.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException error)
            {
                var code = error.HResult & 0xFFFF;
                if (code == 32 || code == 33) return null;
                throw;
            }
        }

        private string[] RepackTips()
        {
            if (!File.Exists(Path.Combine(root, "HEAD"))) return new string[0];
            return Text(Run(null, null, "for-each-ref", "--format=%(objectname)", Heads))
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(Id).Distinct().ToArray();
        }

        private void Repack(string[] tips, bool background)
        {
            // Keep even unreferenced objects, including objects created concurrently by a commit.
            var args = new List<string> { "repack", "-a", "-d", "--keep-unreachable", "--no-cruft", "-n", "-q" };
            if (background) args.AddRange(new[] { "--threads=1", "--window-memory=64m" });
            Checked(Invoke(null, null, System.Threading.Timeout.Infinite, background, args.ToArray()));

            // Record the starting tips so a commit made during packing is counted next time.
            var temporary = PackedTipsPath + ".tmp";
            try
            {
                File.WriteAllLines(temporary, tips, Utf8);
                if (File.Exists(PackedTipsPath)) File.Replace(temporary, PackedTipsPath, null);
                else File.Move(temporary, PackedTipsPath);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private long ObjectBytes()
        {
            var sizes = Text(Run(null, null, "count-objects", "-v")).Split('\n')
                .Where(line => line.StartsWith("size: ", StringComparison.Ordinal) ||
                    line.StartsWith("size-pack: ", StringComparison.Ordinal));
            return sizes.Sum(line => Int64.Parse(line.Substring(line.IndexOf(':') + 1),
                System.Globalization.CultureInfo.InvariantCulture)) * 1024;
        }

        public string MergeBase(string left, string right)
        {
            var result = Invoke(null, null, "merge-base", "--all", Id(left), Id(right));
            if (result.Code == 1) throw new InvalidOperationException("两个分支没有共同祖先，无法创建可靠的修订稿。");
            var bases = Text(result).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (bases.Length != 1) throw new InvalidOperationException("两个分支有多个共同合并基准，当前版本不支持这类历史。");
            return Id(bases[0]);
        }

        public List<Revision> History(string branch = null)
        {
            if (!File.Exists(Path.Combine(root, "HEAD"))) return new List<Revision>();
            var raw = Text(Run(null, null, "log", branch == null ? "--all" : HistoryReference(branch), "--topo-order",
                "--decorate=full", "-z", "--format=%H%x00%P%x00%aI%x00%an%x00%D%x00%s%x00%B", "--"));
            if (raw.Length == 0) return new List<Revision>();
            var fields = (raw.EndsWith("\0", StringComparison.Ordinal) ? raw.Substring(0, raw.Length - 1) : raw).Split('\0');
            if (fields.Length % 7 != 0) throw new InvalidOperationException("无法读取此仓库的版本历史。");
            var revisions = new List<Revision>();
            for (var index = 0; index < fields.Length; index += 7)
                revisions.Add(new Revision {
                    Id = Id(fields[index]),
                    Parents = fields[index + 1].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(Id).ToArray(),
                    Date = DateTimeOffset.Parse(fields[index + 2], System.Globalization.CultureInfo.InvariantCulture),
                    Author = fields[index + 3],
                    References = fields[index + 4].Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries)
                        .Where(reference => !reference.StartsWith("refs/wordgit/", StringComparison.Ordinal))
                        .Select(reference => reference.Replace("refs/heads/", "").Replace("refs/tags/", "").Replace("refs/remotes/", "")).ToArray(),
                    Subject = fields[index + 5],
                    Message = fields[index + 6].TrimEnd()
                });
            return revisions;
        }

        public byte[] Extract(string revision)
        {
            revision = Id(revision);
            var entry = Invoke(null, null, "cat-file", "blob", revision + ":" + DocumentSnapshot.ManifestFile);
            if (entry.Code == 0)
            {
                var snapshot = DocumentSnapshot.Read(entry.Bytes);
                var request = Utf8.GetBytes(String.Join("\n", snapshot.Objects.Select(id =>
                    revision + ":" + DocumentSnapshot.ChunksFolder + "/" + id)) + "\n");
                return snapshot.Restore(Checked(Invoke(request, null, System.Threading.Timeout.Infinite, false, "cat-file", "--batch")).Bytes);
            }
            var expanded = Invoke(null, null, "cat-file", "blob", revision + ":" + DocxFormat.MetaName);
            if (expanded.Code == 0) return ExtractExpanded(revision, DocxFormat.ParseMeta(Utf8.GetString(expanded.Bytes)));
            var whole = Invoke(null, null, "cat-file", "blob", revision + ":" + FileName);
            if (whole.Code == 0) return whole.Bytes;
            throw new InvalidDataException("版本 " + revision.Substring(0, 8) +
                " 不包含可识别的文档快照。请保留仓库并检查版本格式；此操作不会删除历史。");
        }

        // Streaming counterpart of Extract: chunk data flows from git straight into the output
        // stream (a temp file), so restoring a large document needs only a few megabytes of memory.
        public void ExtractTo(string revision, Stream output)
        {
            revision = Id(revision);
            var entry = Invoke(null, null, "cat-file", "blob", revision + ":" + DocumentSnapshot.ManifestFile);
            if (entry.Code != 0)
            {
                var bytes = Extract(revision);
                output.Write(bytes, 0, bytes.Length);
                return;
            }
            var snapshot = DocumentSnapshot.Read(entry.Bytes);
            var request = Utf8.GetBytes(String.Join("\n", snapshot.Objects.Select(id =>
                revision + ":" + DocumentSnapshot.ChunksFolder + "/" + id)) + "\n");
            InvokeWithOutput(request, null, System.Threading.Timeout.Infinite, false,
                stream => snapshot.Restore(stream, output), "cat-file", "--batch");
        }

        private byte[] ExtractExpanded(string revision, DocxMeta meta)
        {
            if (meta.Entries.Count == 0 || meta.Entries.Select(item => item.Name).Distinct(StringComparer.Ordinal).Count() != meta.Entries.Count)
                throw new InvalidDataException("旧版文档清单为空或包含重复部件。");
            var request = Utf8.GetBytes(String.Join("\n", meta.Entries.Select(item => revision + ":" + item.Name)) + "\n");
            var batch = Checked(Invoke(request, null, System.Threading.Timeout.Infinite, false, "cat-file", "--batch")).Bytes;
            var position = 0;
            var parts = new List<DocxPart>();
            foreach (var item in meta.Entries)
            {
                var end = Array.IndexOf(batch, (byte)'\n', position);
                if (end < 0 || end - position > 100) throw new InvalidDataException("旧版文档部件读取不完整。");
                var header = Encoding.ASCII.GetString(batch, position, end - position);
                var fields = header.Split(' ');
                long size;
                if (fields.Length != 3 || fields[1] != "blob" || !Regex.IsMatch(fields[0], @"\A(?:[a-f0-9]{40}|[a-f0-9]{64})\z") ||
                    !Int64.TryParse(fields[2], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out size) ||
                    size != item.Size || size < 0 || size > Int32.MaxValue)
                    throw new InvalidDataException("旧版文档部件缺失或长度不符：" + item.Name);
                position = end + 1;
                if (batch.LongLength - position <= size || batch[position + (int)size] != (byte)'\n')
                    throw new InvalidDataException("旧版文档部件内容不完整：" + item.Name);
                var bytes = new byte[(int)size];
                Buffer.BlockCopy(batch, position, bytes, 0, bytes.Length);
                if (Sha(bytes) != item.Sha) throw new InvalidDataException("旧版文档部件校验失败：" + item.Name);
                parts.Add(new DocxPart { Name = item.Name, Bytes = bytes });
                position += bytes.Length + 1;
            }
            if (position != batch.Length || DocxFormat.Hex(DocxFormat.Fingerprint(parts)) != meta.Fingerprint)
                throw new InvalidDataException("旧版文档整体部件校验失败。");
            return DocxFormat.Write(parts);
        }

        public PreparedSnapshot PrepareSnapshot(byte[] bytes)
        {
            EnsureRepository();
            var chunks = DocumentSnapshot.Split(bytes);
            var objects = DocumentSnapshot.BlobObjects(chunks, Text(Run(null, null, "rev-parse", "--show-object-format")));
            var sizes = chunks.Select(chunk => chunk.Count).ToArray();
            ImportChunks(objects, sizes, MissingChunkIndexes(objects),
                (stream, index) => { var chunk = chunks[index]; stream.Write(chunk.Array, chunk.Offset, chunk.Count); });
            return StoreManifest(objects, sizes, bytes.Length, Sha(bytes));
        }

        // First streaming pass over the document file: computes the whole-file checksum and the
        // content-defined chunk plan without retaining any chunk content in memory.
        internal FileSnapshot PlanFile(string path)
        {
            EnsureRepository();
            var format = Text(Run(null, null, "rev-parse", "--show-object-format"));
            using (var stream = OpenShared(path))
            {
                var plan = DocumentSnapshot.Plan(stream, format);
                var info = new FileInfo(path);
                return new FileSnapshot { Plan = plan, WriteTimeUtc = info.LastWriteTimeUtc };
            }
        }

        // Second pass: seeks to and reads only the chunks git does not have yet, streaming them
        // into fast-import. The get-mark verification below proves every imported blob id matches
        // the plan, which also detects the file changing mid-import.
        internal PreparedSnapshot ImportPlan(string path, FileSnapshot file)
        {
            var plan = file.Plan;
            var missing = MissingChunkIndexes(plan.Objects);
            if (missing.Length > 0)
            {
                using (var stream = OpenShared(path))
                {
                    var info = new FileInfo(path);
                    if (info.Length != plan.Length || info.LastWriteTimeUtc != file.WriteTimeUtc)
                        throw new InvalidOperationException("文档在提交准备期间被修改，请重新提交。");
                    var buffer = new byte[64 * 1024];
                    try
                    {
                        ImportChunks(plan.Objects, plan.Sizes, missing, (target, index) =>
                        {
                            stream.Position = plan.Offsets[index];
                            var remaining = plan.Sizes[index];
                            while (remaining > 0)
                            {
                                var step = Math.Min(buffer.Length, remaining);
                                var read = 0;
                                while (read < step)
                                {
                                    var count = stream.Read(buffer, read, step - read);
                                    if (count <= 0)
                                        throw new InvalidOperationException("文档在提交准备期间被修改，请重新提交。");
                                    read += count;
                                }
                                target.Write(buffer, 0, step);
                                remaining -= step;
                            }
                        });
                    }
                    catch (InvalidDataException error)
                    {
                        var verify = new FileInfo(path);
                        if (error.Message.Contains("数据块写入校验失败") &&
                            (verify.Length != plan.Length || verify.LastWriteTimeUtc != file.WriteTimeUtc))
                            throw new InvalidOperationException("文档在提交准备期间被修改，请重新提交。");
                        throw;
                    }
                }
            }
            return StoreManifest(plan.Objects, plan.Sizes, plan.Length, plan.Checksum);
        }

        public PreparedSnapshot PrepareSnapshotFile(string path)
        {
            return ImportPlan(path, PlanFile(path));
        }

        private int[] MissingChunkIndexes(string[] objects)
        {
            var unique = objects.Select((id, index) => new { Id = id, Index = index }).GroupBy(item => item.Id)
                .Select(group => group.First()).ToArray();
            var check = Text(Run(Utf8.GetBytes(String.Join("\n", unique.Select(item => item.Id)) + "\n"),
                null, "cat-file", "--batch-check")).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (check.Length != unique.Length) throw new InvalidDataException("无法确认已保存的数据块。");
            return unique.Where((item, index) => check[index] == item.Id + " missing").Select(item => item.Index).ToArray();
        }

        private void ImportChunks(string[] objects, int[] sizes, int[] missing, Action<Stream, int> writeChunk)
        {
            if (missing.Length == 0) return;
            var imported = Checked(InvokeWithInput(stream =>
            {
                for (var i = 0; i < missing.Length; i++)
                {
                    var mark = (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    var header = Utf8.GetBytes("blob\nmark :" + mark + "\ndata " +
                        sizes[missing[i]].ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
                    stream.Write(header, 0, header.Length);
                    writeChunk(stream, missing[i]);
                    var tail = Utf8.GetBytes("\nget-mark :" + mark + "\n");
                    stream.Write(tail, 0, tail.Length);
                }
                var done = Utf8.GetBytes("done\n");
                stream.Write(done, 0, done.Length);
            }, null, System.Threading.Timeout.Infinite, true, "-c", "pack.threads=1",
                "fast-import", "--quiet", "--done"));
            var written = Text(imported).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(Id);
            if (!written.SequenceEqual(missing.Select(index => objects[index]))) throw new InvalidDataException("数据块写入校验失败。");
        }

        private PreparedSnapshot StoreManifest(string[] objects, int[] sizes, long totalLength, string checksum)
        {
            var manifest = DocumentSnapshot.Manifest(totalLength, checksum, sizes, objects);
            var manifestId = Id(Text(Run(manifest, null, "hash-object", "-w", "--no-filters", "--stdin")));
            var entries = String.Join("\n", objects.Distinct().OrderBy(id => id, StringComparer.Ordinal)
                .Select(id => "100644 blob " + id + "\t" + id)) + "\n";
            var chunksTree = Id(Text(Run(Utf8.GetBytes(entries), null, "mktree")));
            var rootEntries = "040000 tree " + chunksTree + "\t" + DocumentSnapshot.ChunksFolder + "\n" +
                "100644 blob " + manifestId + "\t" + DocumentSnapshot.ManifestFile + "\n";
            var tree = Id(Text(Run(Utf8.GetBytes(rootEntries), null, "mktree")));
            return new PreparedSnapshot(tree, checksum);
        }

        private static FileStream OpenShared(string path)
        {
            return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        }

        private void EnsureRepository()
        {
            Directory.CreateDirectory(root);
            if (!File.Exists(Path.Combine(root, "HEAD")))
                Checked(Invoke(null, null, "init", "--bare", "--initial-branch=main", "--template="));
        }

        internal string SnapshotChecksum(string revision)
        {
            var entry = Invoke(null, null, "cat-file", "blob", Id(revision) + ":" + DocumentSnapshot.ManifestFile);
            return entry.Code == 0 ? DocumentSnapshot.Read(entry.Bytes).Checksum : Sha(Extract(revision));
        }

        public string Commit(byte[] bytes, string message, string author, string expectedHead,
            string branch = null, string sourceBranch = null, string sourceTip = null, PreparedSnapshot prepared = null)
        {
            if (bytes == null || bytes.Length < 4) throw new ArgumentException("无法读取文档数据。");
            if (String.IsNullOrWhiteSpace(message) || message.IndexOf('\0') >= 0) throw new ArgumentException("请填写有效的提交说明。");
            EnsureRepository();
            branch = branch ?? ActiveBranch();
            if (ActiveBranch() != branch || BranchTip(branch) != expectedHead ||
                (sourceBranch != null && (sourceBranch == branch || BranchTip(sourceBranch) != sourceTip)))
                throw new InvalidOperationException("分支或版本历史已变化。请刷新后再操作。");
            var checksum = Sha(bytes);
            if (sourceBranch == null && expectedHead != null && checksum == SnapshotChecksum(expectedHead)) return null;
            prepared = prepared ?? PrepareSnapshot(bytes);
            if (prepared.Checksum != checksum) throw new InvalidDataException("待提交的文档与分块快照不一致。");
            return FinishCommit(prepared, message, author, expectedHead, branch, sourceBranch, sourceTip);
        }

        // Streaming counterpart of Commit: the document is read from disk in bounded buffers
        // instead of being loaded into memory first.
        public string CommitFile(string path, string message, string author, string expectedHead,
            string branch = null, string sourceBranch = null, string sourceTip = null, FileSnapshot planned = null)
        {
            if (String.IsNullOrWhiteSpace(message) || message.IndexOf('\0') >= 0) throw new ArgumentException("请填写有效的提交说明。");
            EnsureRepository();
            branch = branch ?? ActiveBranch();
            if (ActiveBranch() != branch || BranchTip(branch) != expectedHead ||
                (sourceBranch != null && (sourceBranch == branch || BranchTip(sourceBranch) != sourceTip)))
                throw new InvalidOperationException("分支或版本历史已变化。请刷新后再操作。");
            var file = planned ?? PlanFile(path);
            if (sourceBranch == null && expectedHead != null && file.Plan.Checksum == SnapshotChecksum(expectedHead)) return null;
            var prepared = ImportPlan(path, file);
            return FinishCommit(prepared, message, author, expectedHead, branch, sourceBranch, sourceTip);
        }

        // Commits an already-prepared snapshot (e.g. imported from another path) without
        // reading any document bytes again. Callers must hold the repository lock.
        internal string CommitPrepared(PreparedSnapshot prepared, string message, string author, string expectedHead,
            string branch = null, string sourceBranch = null, string sourceTip = null)
        {
            if (prepared == null) throw new ArgumentException("无法读取文档数据。");
            if (String.IsNullOrWhiteSpace(message) || message.IndexOf('\0') >= 0) throw new ArgumentException("请填写有效的提交说明。");
            EnsureRepository();
            branch = branch ?? ActiveBranch();
            if (ActiveBranch() != branch || BranchTip(branch) != expectedHead ||
                (sourceBranch != null && (sourceBranch == branch || BranchTip(sourceBranch) != sourceTip)))
                throw new InvalidOperationException("分支或版本历史已变化。请刷新后再操作。");
            if (sourceBranch == null && expectedHead != null && prepared.Checksum == SnapshotChecksum(expectedHead)) return null;
            return FinishCommit(prepared, message, author, expectedHead, branch, sourceBranch, sourceTip);
        }

        private string FinishCommit(PreparedSnapshot prepared, string message, string author, string expectedHead,
            string branch, string sourceBranch, string sourceTip)
        {
            var reference = BranchRef(branch);
            var args = new List<string> { "commit-tree", Id(prepared.Tree) };
            if (expectedHead != null) args.AddRange(new[] { "-p", Id(expectedHead) });
            if (sourceBranch != null)
            {
                if (expectedHead == null || MergeBase(expectedHead, sourceTip) == sourceTip)
                    throw new InvalidOperationException("来源分支已包含在当前分支。");
                args.AddRange(new[] { "-p", Id(sourceTip) });
            }
            args.AddRange(new[] { "-F", "-" });
            var identity = Regex.Replace(author ?? "", "[<>\\x00-\\x1f]", "").Trim();
            if (identity.Length == 0) identity = "WordGit";
            var commit = Id(Text(Run(Utf8.GetBytes(message.Trim() + "\n"), identity, args.ToArray())));
            var operations = new List<string> {
                "update " + reference + " " + commit + " " + (expectedHead ?? new string('0', commit.Length))
            };
            if (sourceBranch != null) operations.Add("verify " + HistoryReference(sourceBranch) + " " + Id(sourceTip));
            RefTransaction(identity, operations.ToArray());
            if (commitCreated != null) commitCreated(root);
            return commit;
        }

        private static string BranchRef(string branch)
        {
            if (String.IsNullOrWhiteSpace(branch) || branch.Length > 150 || branch.Any(character =>
                Char.IsWhiteSpace(character) || Char.IsControl(character) || "~^:?*[\\".IndexOf(character) >= 0))
                throw new ArgumentException("请输入有效的 Git 分支名，不包含空格或特殊符号。");
            return Heads + branch;
        }

        private void RefTransaction(string identity, params string[] operations)
        {
            Run(Utf8.GetBytes("start\n" + String.Join("\n", operations) + "\nprepare\ncommit\n"), identity, "update-ref", "--stdin");
        }

        private static string Id(string candidate)
        {
            if (!Regex.IsMatch(candidate ?? "", @"\A(?:[a-f0-9]{40}|[a-f0-9]{64})\z")) throw new InvalidOperationException("无效的 Git 版本标识。");
            return candidate;
        }

        private static string Text(Result result) { return Utf8.GetString(Checked(result).Bytes).Trim(' ', '\r', '\n'); }
        private Result Run(byte[] input, string identity, params string[] args) { return Checked(Invoke(input, identity, args)); }

        private static Result Checked(Result result)
        {
            if (result.Code != 0) throw new InvalidOperationException("Git 操作失败：" + result.Error.Trim());
            return result;
        }

        private Result Invoke(byte[] input, string identity, params string[] args)
        {
            return Invoke(input, identity, 45000, false, args);
        }

        private Result Invoke(byte[] input, string identity, int timeout, bool lowPriority, params string[] args)
        {
            Action<Stream> writer = input == null ? null : (Action<Stream>)(stream => stream.Write(input, 0, input.Length));
            return InvokeWithInput(writer, identity, timeout, lowPriority, args);
        }

        private ProcessStartInfo CreateInfo(string identity, bool redirectInput, params string[] args)
        {
            var command = new List<string>();
            command.AddRange(new[] { "-c", "core.longpaths=true" });
            command.Add("--git-dir=.");
            command.AddRange(args);
            var info = new ProcessStartInfo(git, String.Join(" ", command.Select(Quote)))
            {
                WorkingDirectory = root, UseShellExecute = false, CreateNoWindow = true,
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

        private static void TryLowerPriority(Process process)
        {
            try { process.PriorityClass = ProcessPriorityClass.BelowNormal; }
            catch (System.ComponentModel.Win32Exception) { }
            catch (InvalidOperationException) { }
        }

        private Result InvokeWithInput(Action<Stream> input, string identity, int timeout, bool lowPriority, params string[] args)
        {
            var info = CreateInfo(identity, input != null, args);
            using (var process = Process.Start(info))
            using (var stdout = new MemoryStream())
            {
                if (lowPriority) TryLowerPriority(process);
                var read = process.StandardOutput.BaseStream.CopyToAsync(stdout);
                var errors = process.StandardError.ReadToEndAsync();
                var write = input == null ? Task.FromResult(true) : Task.Run(() =>
                {
                    using (var stream = process.StandardInput.BaseStream) input(stream);
                    return true;
                });
                if (!process.WaitForExit(timeout))
                {
                    process.Kill();
                    process.WaitForExit();
                    throw new TimeoutException("Git 操作超时。");
                }
                Task.WaitAll(read, errors);
                if (process.ExitCode == 0) write.GetAwaiter().GetResult();
                // Surface writer errors that are not a consequence of git failing (a broken pipe
                // is just the downstream effect of git exiting; our own errors must not be lost).
                else if (write.IsFaulted && !(write.Exception.InnerException is IOException)) write.GetAwaiter().GetResult();
                else if (write.IsFaulted) GC.KeepAlive(write.Exception);
                return new Result { Code = process.ExitCode, Bytes = stdout.ToArray(), Error = errors.Result };
            }
        }

        // Like InvokeWithInput, but stdout is streamed to a reader callback instead of being
        // buffered into memory. If the reader reports a data problem, git is stopped and the
        // reader's error is thrown (it is more informative than git's exit code).
        private void InvokeWithOutput(byte[] input, string identity, int timeout, bool lowPriority,
            Action<Stream> reader, params string[] args)
        {
            var info = CreateInfo(identity, input != null, args);
            using (var process = Process.Start(info))
            {
                if (lowPriority) TryLowerPriority(process);
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
                    process.Kill();
                    process.WaitForExit();
                    throw new TimeoutException("Git 操作超时。");
                }
                if (readerError != null)
                {
                    try { process.Kill(); }
                    catch (InvalidOperationException) { }
                    process.WaitForExit();
                    Task.WaitAll(errors);
                    throw readerError;
                }
                if (!process.WaitForExit(timeout))
                {
                    process.Kill();
                    process.WaitForExit();
                    throw new TimeoutException("Git 操作超时。");
                }
                Task.WaitAll(read, errors);
                if (readerError != null) throw readerError;
                if (process.ExitCode != 0) throw new InvalidOperationException("Git 操作失败：" + errors.Result.Trim());
                write.GetAwaiter().GetResult();
            }
        }

        private static string Quote(string arg)
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

        private sealed class Result { public int Code; public byte[] Bytes; public string Error; }
    }

    // A planned chunked snapshot of a document file, plus the file stamp used to prove the file
    // did not change between the planning pass and the import pass.
    internal sealed class FileSnapshot
    {
        public SnapshotPlan Plan;
        public DateTime WriteTimeUtc;
    }
}
