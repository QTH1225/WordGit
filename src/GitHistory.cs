using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace WordGit
{
    internal sealed class Revision
    {
        public string Id, Author, Subject, Message;
        public string[] Parents, References;
        public DateTimeOffset Date;
        public override string ToString() { return Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm") + "   " + Id.Substring(0, 8) + "   " + Subject + "   · " + Author; }
    }

    internal sealed class ReviewImport
    {
        public string Id, Reference;
        public bool Reused;
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
        private const long ManifestLimit = 64L * 1024 * 1024;
        private const long MetaLimit = 16L * 1024 * 1024;
        // Unreachable objects (failed imports, expired remote backups) are dropped after this age.
        private const string CruftExpiration = "2.weeks.ago";
        private static readonly TimeSpan FullRepackInterval = TimeSpan.FromDays(30);
        private static readonly Encoding Utf8 = new UTF8Encoding(false, true);
        // Commit messages written by other Git clients are not guaranteed to be UTF-8.
        private static readonly Encoding LenientUtf8 = new UTF8Encoding(false, false);
        private static readonly string[] ReservedNames = { "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
        private readonly string root;
        private readonly GitProcess git;
        private readonly Action<string> commitCreated;
        public string RepositoryPath { get { return root; } }

        public GitHistory(string documentPath, string storage = null, string executable = null, Action<string> commitCreated = null)
            : this(RepositoryFor(documentPath, storage), executable, commitCreated)
        {
        }

        private GitHistory(string repositoryPath, string executable, Action<string> commitCreated)
        {
            root = repositoryPath;
            git = new GitProcess(executable ?? GitProcess.FindGit(), repositoryPath);
            this.commitCreated = commitCreated;
        }

        internal static string RepositoryFor(string documentPath, string storage = null)
        {
            return Path.Combine(storage ?? AppPaths.Repositories,
                Sha(Utf8.GetBytes(Path.GetFullPath(documentPath).ToUpperInvariant())) + ".git");
        }

        internal static GitHistory OpenRepository(string repositoryPath)
        {
            return new GitHistory(repositoryPath, null, (Action<string>)null);
        }

        public static string Sha(byte[] bytes)
        {
            using (var algorithm = SHA256.Create()) return Hex(algorithm.ComputeHash(bytes));
        }

        internal static string Hex(byte[] bytes)
        {
            return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
        }

        private bool Exists { get { return File.Exists(Path.Combine(root, "HEAD")); } }

        public string ActiveBranch()
        {
            if (!Exists) return "main";
            var reference = Text(Run(null, null, "symbolic-ref", "HEAD"));
            if (!reference.StartsWith(Heads, StringComparison.Ordinal)) throw new InvalidOperationException("HEAD 未指向本地分支。");
            return reference.Substring(Heads.Length);
        }

        public string Head() { return BranchTip(ActiveBranch()); }

        public string BranchTip(string branch)
        {
            if (!Exists) return null;
            var result = Invoke(null, null, "rev-parse", "--verify", "-q", HistoryReference(branch));
            if (result.Code == 1) return null;
            return Id(Text(result));
        }

        public string[] Branches(string mergedInto = null)
        {
            if (!Exists) return new string[0];
            var args = new List<string> { "for-each-ref", "--sort=refname", "--format=%(refname:strip=2)" };
            if (mergedInto != null) args.Add("--merged=" + Id(mergedInto));
            args.Add(Heads);
            return Text(Run(null, null, args.ToArray())).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        // HEAD's target plus every local branch tip: changes exactly when a commit, switch or
        // fast-forward is published. Fetching only moves remote-tracking and backup refs.
        internal string ReferenceState()
        {
            if (!Exists) return "";
            return Text(Run(null, null, "symbolic-ref", "HEAD")) + "\n" +
                Text(Run(null, null, "for-each-ref", "--format=%(objectname) %(refname)", Heads));
        }

        // A revision's own tree serves as the prepared snapshot when the working copy is replaced
        // by exactly that revision, so publishing a restore needs no second pass over the file.
        internal PreparedSnapshot PreparedFromRevision(string revision)
        {
            revision = Id(revision);
            return new PreparedSnapshot(Id(Text(Run(null, null, "rev-parse", revision + "^{tree}"))), SnapshotChecksum(revision));
        }

        public void CreateBranch(string name, string revision)
        {
            var reference = BranchRef(name);
            Run(null, null, "check-ref-format", reference);
            if (Branches().Any(branch => String.Equals(branch, name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("同名分支已存在。");
            revision = Id(revision);
            // Prove the starting version is intact before a branch can point at it.
            ExtractTo(revision, new DiscardStream());
            Run(null, null, "update-ref", reference, revision, new string('0', revision.Length));
        }

        internal static string ReviewBranchConflict(string name, IEnumerable<string> branches)
        {
            var comparison = StringComparison.OrdinalIgnoreCase;
            var conflict = branches.FirstOrDefault(branch => String.Equals(name, branch, comparison) ||
                name.StartsWith(branch + "/", comparison) || branch.StartsWith(name + "/", comparison));
            if (conflict == null) return null;
            return "返稿分支「" + name + "」与已有分支「" + conflict +
                "」冲突。请换一个名称，例如 review-老师审阅。";
        }

        internal const string ReviewArchivePrefix = "refs/wordgit/reviews/";

        // Called with the repository lock. The bytes and the actual sent version together
        // identify a return, regardless of the proposed branch name or its current main tip.
        internal ReviewImport ImportOrReuseReview(string path, string basis, string newBranch, string reviewer,
            string targetBranch, string expectedTarget)
        {
            basis = Id(basis);
            if (ActiveBranch() != targetBranch || BranchTip(targetBranch) != expectedTarget)
                throw new InvalidOperationException("当前分支或提交已变化，请刷新后重试。");

            using (var input = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                string checksum;
                using (var sha = SHA256.Create()) checksum = Hex(sha.ComputeHash(input));
                var previous = Text(Run(null, null, "log", "--all", "--format=%H", "--fixed-strings",
                    "--grep=WordGit-Review-SHA256: " + checksum, "--"));
                var matches = previous.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(Id).Where(id => ImportedReviewBase(id) == basis && SnapshotChecksum(id) == checksum).ToArray();
                var merged = matches.FirstOrDefault(id => IsAncestor(id, expectedTarget));
                if (merged != null) return new ReviewImport { Id = merged, Reused = true };
                var existing = matches.FirstOrDefault();
                if (existing != null)
                {
                    var branch = Text(Run(null, null, "for-each-ref", "--points-at=" + existing, "--format=%(refname:strip=2)", Heads))
                        .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    var reference = branch ?? (ReviewArchivePrefix + existing);
                    // Keep a stable ref when the branch that originally held the return has advanced.
                    if (branch == null && BranchTip(reference) == null)
                        RefTransaction(null, "create " + reference + " " + existing,
                            "verify " + BranchRef(targetBranch) + " " + Id(expectedTarget));
                    return new ReviewImport { Id = existing, Reference = reference, Reused = true };
                }
                var imported = ImportReviewFile(path, basis, newBranch, reviewer, targetBranch, expectedTarget);
                return new ReviewImport { Id = imported, Reference = newBranch ?? (ReviewArchivePrefix + imported) };
            }
        }

        public string ImportReviewFile(string path, string basis, string newBranch, string reviewer,
            string targetBranch, string expectedTarget)
        {
            if (String.IsNullOrWhiteSpace(reviewer) || reviewer.IndexOf('\0') >= 0)
                throw new ArgumentException("请填写返稿来源。");
            EnsureRepository();
            var reference = newBranch == null ? null : BranchRef(newBranch);
            if (reference != null)
            {
                Run(null, null, "check-ref-format", reference);
                var conflict = ReviewBranchConflict(newBranch, Branches());
                if (conflict != null) throw new InvalidOperationException(conflict);
            }
            basis = Id(basis);
            if (expectedTarget == null || ActiveBranch() != targetBranch || BranchTip(targetBranch) != expectedTarget)
                throw new InvalidOperationException("当前分支或提交已变化，请刷新后重新导入。");
            ExtractTo(basis, new DiscardStream());
            var prepared = PrepareSnapshotFile(path);
            if (ActiveBranch() != targetBranch || BranchTip(targetBranch) != expectedTarget)
                throw new InvalidOperationException("导入期间当前分支已变化，返稿分支尚未创建，请重新导入。");
            var identity = Regex.Replace(reviewer, @"[<>\x00-\x1f]", "").Trim();
            if (identity.Length == 0) throw new ArgumentException("请填写有效的返稿来源。");
            var message = "导入审阅稿：" + Path.GetFileName(path) + "\n\nWordGit-Review-Base: " + basis +
                "\nWordGit-Review-SHA256: " + prepared.Checksum + "\n";
            var commit = Id(Text(Run(Utf8.GetBytes(message), identity, "commit-tree", prepared.Tree, "-p", basis, "-F", "-")));
            var publish = "create ";
            if (reference == null)
            {
                reference = ReviewArchivePrefix + commit;
                if (BranchTip(reference) == commit) publish = "verify ";
            }
            RefTransaction(identity, publish + reference + " " + commit,
                "verify " + BranchRef(targetBranch) + " " + Id(expectedTarget));
            if (commitCreated != null) commitCreated(root);
            return commit;
        }

        internal string ImportedReviewBase(string revision)
        {
            var text = LenientText(Run(null, null, "show", "-s", "--format=%P%n%B", Id(revision)));
            var lines = text.Split('\n');
            var marker = lines.Where(line => line.StartsWith("WordGit-Review-Base: ", StringComparison.Ordinal)).ToArray();
            if (marker.Length != 1) return null;
            var parent = marker[0].Substring("WordGit-Review-Base: ".Length).Trim();
            return Regex.IsMatch(parent, @"\A(?:[a-f0-9]{40}|[a-f0-9]{64})\z") && lines[0].Trim() == parent ? parent : null;
        }

        internal string[] PreviouslyMergedReturns(string tip, string basis)
        {
            var previous = Text(Run(null, null, "log", "--first-parent", "--merges", "--format=%P",
                Id(basis) + ".." + Id(tip), "--"));
            return previous.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .SelectMany(line => line.Split(' ').Skip(1)).Distinct(StringComparer.Ordinal)
                .Where(parent => ImportedReviewBase(parent) == basis).ToArray();
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
            if (!IsAncestor(targetTip, head)) throw new InvalidOperationException("该分支还有提交未合入当前分支。");

            // Git locks HEAD and its referent. If HEAD now points to the target, the two operations conflict.
            RefTransaction(null, "verify HEAD " + head,
                "option no-deref", "delete " + BranchRef(target) + " " + targetTip);
        }

        public IDisposable Lock()
        {
            Directory.CreateDirectory(root);
            try { return new FileStream(Path.Combine(root, "wordgit.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException error)
            {
                if (IsSharingViolation(error)) throw new InvalidOperationException("另一处 WordGit 正在处理这个文档，请稍后重试。");
                throw;
            }
        }

        private static bool IsSharingViolation(IOException error)
        {
            var code = error.HResult & 0xFFFF;
            return code == 32 || code == 33;
        }

        // Full maintenance: reachable objects go into one pack, recent unreachable objects into a
        // cruft pack, and unreachable objects older than the expiration are dropped.
        public RepackResult Repack()
        {
            using (var packing = TryRepackLock())
            {
                if (packing == null) throw new InvalidOperationException("此文档的历史正在压缩。");
                var tips = RepackTips();
                if (tips.Length == 0) throw new InvalidOperationException("尚无提交可压缩。");

                var before = ObjectBytes();
                Repack(tips, false, true);
                return new RepackResult(before, ObjectBytes());
            }
        }

        internal bool TryAutoRepack(int threshold, Func<bool> canStart)
        {
            if (!Exists) return true;
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
                    "--max-count=" + threshold.ToString(CultureInfo.InvariantCulture), "--stdin"));
                if (Int32.Parse(count, CultureInfo.InvariantCulture) < threshold) return true;
                if (!canStart()) return false;

                var full = !File.Exists(FullRepackPath) || DateTime.UtcNow - File.GetLastWriteTimeUtc(FullRepackPath) > FullRepackInterval;
                Repack(tips, true, full);
                return true;
            }
        }

        private string PackedTipsPath { get { return Path.Combine(root, "pack-tips"); } }
        private string FullRepackPath { get { return Path.Combine(root, "pack-full"); } }

        private FileStream TryRepackLock()
        {
            try { return new FileStream(Path.Combine(root, "wordgit-repack.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException error)
            {
                if (IsSharingViolation(error)) return null;
                throw;
            }
        }

        private string[] RepackTips()
        {
            if (!Exists) return new string[0];
            return Text(Run(null, null, "for-each-ref", "--format=%(objectname)", Heads))
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(Id).Distinct().ToArray();
        }

        private void Repack(string[] tips, bool background, bool full)
        {
            // Neither mode deletes objects newer than the cruft expiration, so objects written by a
            // concurrent commit survive even before its branch update is published. Geometric
            // repacks only merge small packs, keeping routine maintenance cheap on large histories.
            var args = full
                ? new List<string> { "repack", "--cruft", "-d", "--cruft-expiration=" + CruftExpiration, "-n", "-q" }
                : new List<string> { "repack", "--geometric=2", "-d", "-n", "-q" };
            if (background) args.AddRange(new[] { "--threads=1", "--window-memory=64m" });
            Checked(git.Invoke(null, null, Timeout.Infinite, background, args.ToArray()));
            // repack only expires packed objects; loose unreachable objects need prune, which also
            // keeps everything newer than the expiration.
            if (full) Checked(git.Invoke(null, null, Timeout.Infinite, background, "prune", "--expire=" + CruftExpiration));

            // Record the starting tips so a commit made during packing is counted next time.
            ReplaceSmallFile(PackedTipsPath, tips);
            if (full) ReplaceSmallFile(FullRepackPath, new[] { DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) });
        }

        private static void ReplaceSmallFile(string path, IEnumerable<string> lines)
        {
            var temporary = path + ".tmp";
            try
            {
                File.WriteAllLines(temporary, lines, Utf8);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private long ObjectBytes()
        {
            var sizes = Text(Run(null, null, "count-objects", "-v")).Split('\n')
                .Where(line => line.StartsWith("size: ", StringComparison.Ordinal) ||
                    line.StartsWith("size-pack: ", StringComparison.Ordinal) ||
                    line.StartsWith("size-garbage: ", StringComparison.Ordinal));
            return sizes.Sum(line => Int64.Parse(line.Substring(line.IndexOf(':') + 1), CultureInfo.InvariantCulture)) * 1024;
        }

        internal bool IsAncestor(string ancestor, string descendant)
        {
            var result = Invoke(null, null, "merge-base", "--is-ancestor", Id(ancestor), Id(descendant));
            if (result.Code == 1) return false;
            Checked(result);
            return true;
        }

        // The basis for a Word revision merge. Criss-cross histories have several equally good
        // bases; Git's own choice is used, and the author still reviews every resulting revision.
        public string MergeBase(string left, string right)
        {
            var result = Invoke(null, null, "merge-base", Id(left), Id(right));
            if (result.Code == 1) throw new InvalidOperationException("两个分支没有共同祖先，无法创建可靠的修订稿。");
            return Id(Text(result));
        }

        public List<Revision> History(string branch = null)
        {
            if (!Exists) return new List<Revision>();
            // --all deliberately includes refs/wordgit/remote-backup/*: commits a collaborator
            // force-pushed away stay visible (unlabelled) instead of silently disappearing.
            var raw = LenientText(Run(null, null, "log", branch == null ? "--all" : HistoryReference(branch), "--topo-order",
                "--decorate=full", "-z", "--format=%H%x00%P%x00%aI%x00%an%x00%D%x00%s%x00%B", "--"));
            if (raw.Length == 0) return new List<Revision>();
            var fields = (raw.EndsWith("\0", StringComparison.Ordinal) ? raw.Substring(0, raw.Length - 1) : raw).Split('\0');
            if (fields.Length % 7 != 0) throw new InvalidOperationException("无法读取此仓库的版本历史。");
            var revisions = new List<Revision>();
            for (var index = 0; index < fields.Length; index += 7)
                revisions.Add(new Revision {
                    Id = Id(fields[index]),
                    Parents = fields[index + 1].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(Id).ToArray(),
                    Date = DateTimeOffset.Parse(fields[index + 2], CultureInfo.InvariantCulture),
                    Author = fields[index + 3],
                    References = fields[index + 4].Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries)
                        .Where(reference => !reference.StartsWith("refs/wordgit/", StringComparison.Ordinal))
                        .Select(reference => reference.Replace("refs/heads/", "").Replace("refs/tags/", "").Replace("refs/remotes/", "")).ToArray(),
                    Subject = fields[index + 5],
                    Message = fields[index + 6].TrimEnd()
                });
            return revisions;
        }

        private enum SnapshotKind { Chunked, Expanded, Whole }

        private sealed class StoredSnapshot
        {
            public SnapshotKind Kind;
            public string Object;
            public long Size;
        }

        // Finds which storage format a revision uses, without reading any content.
        private StoredSnapshot Locate(string revision)
        {
            var specs = new[] { DocumentSnapshot.ManifestFile, DocxFormat.MetaName, FileName }.Select(name => revision + ":" + name).ToArray();
            var lines = Text(Run(Utf8.GetBytes(String.Join("\n", specs) + "\n"), null, "cat-file", "--batch-check"))
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length != specs.Length) throw new InvalidDataException("无法读取版本格式。");
            var kinds = new[] { SnapshotKind.Chunked, SnapshotKind.Expanded, SnapshotKind.Whole };
            for (var index = 0; index < lines.Length; index++)
            {
                var fields = lines[index].Split(' ');
                long size;
                if (fields.Length == 3 && fields[1] == "blob" && Regex.IsMatch(fields[0], @"\A(?:[a-f0-9]{40}|[a-f0-9]{64})\z") &&
                    Int64.TryParse(fields[2], NumberStyles.None, CultureInfo.InvariantCulture, out size))
                    return new StoredSnapshot { Kind = kinds[index], Object = fields[0], Size = size };
            }
            throw new InvalidDataException("版本 " + revision.Substring(0, 8) +
                " 不包含可识别的文档快照。请保留仓库并检查版本格式；此操作不会删除历史。");
        }

        private byte[] ReadSmallBlob(StoredSnapshot stored, long limit)
        {
            if (stored.Size > limit) throw new InvalidDataException("版本元数据过大，已拒绝读取。");
            var bytes = Checked(Invoke(null, null, "cat-file", "blob", stored.Object)).Bytes;
            if (bytes.LongLength != stored.Size) throw new InvalidDataException("版本元数据读取不完整。");
            return bytes;
        }

        private DocumentSnapshot ReadManifest(StoredSnapshot stored)
        {
            return DocumentSnapshot.Read(ReadSmallBlob(stored, ManifestLimit));
        }

        // Restores a revision into the output stream with bounded memory, for every storage
        // format. The output is only valid when this returns: callers must discard it on error.
        public void ExtractTo(string revision, Stream output)
        {
            revision = Id(revision);
            var stored = Locate(revision);
            switch (stored.Kind)
            {
                case SnapshotKind.Chunked:
                    var snapshot = ReadManifest(stored);
                    var request = Utf8.GetBytes(String.Join("\n", snapshot.Objects.Select(id =>
                        revision + ":" + DocumentSnapshot.ChunksFolder + "/" + id)) + "\n");
                    git.Stream(request, Timeout.Infinite, stream => snapshot.Restore(stream, output), "cat-file", "--batch");
                    break;
                case SnapshotKind.Expanded:
                    ExtractExpanded(revision, DocxFormat.ParseMeta(Utf8.GetString(ReadSmallBlob(stored, MetaLimit))), output);
                    break;
                default:
                    git.Stream(null, Timeout.Infinite, stream => CopyExactly(stream, output, stored.Size), "cat-file", "blob", stored.Object);
                    break;
            }
        }

        private static void CopyExactly(Stream input, Stream output, long size)
        {
            var buffer = new byte[64 * 1024];
            var remaining = size;
            while (remaining > 0)
            {
                var count = input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
                if (count <= 0) throw new InvalidDataException("历史文档读取不完整。");
                output.Write(buffer, 0, count);
                remaining -= count;
            }
            if (input.ReadByte() >= 0) throw new InvalidDataException("历史文档长度不符。");
        }

        // Pre-chunking format: each ZIP member was stored as its own blob plus a manifest. The
        // container is rebuilt member by member straight into the output.
        private void ExtractExpanded(string revision, DocxMeta meta, Stream output)
        {
            if (meta.Entries.Count == 0 || meta.Entries.Select(item => item.Name).Distinct(StringComparer.Ordinal).Count() != meta.Entries.Count)
                throw new InvalidDataException("旧版文档清单为空或包含重复部件。");
            var request = Utf8.GetBytes(String.Join("\n", meta.Entries.Select(item => revision + ":" + item.Name)) + "\n");
            git.Stream(request, Timeout.Infinite, batch =>
            {
                var buffer = new byte[64 * 1024];
                using (var archive = new ZipArchive(output, ZipArchiveMode.Create, true))
                    foreach (var item in meta.Entries)
                    {
                        var fields = ReadLegacyHeader(batch).Split(' ');
                        long size;
                        if (fields.Length != 3 || fields[1] != "blob" || !Regex.IsMatch(fields[0], @"\A(?:[a-f0-9]{40}|[a-f0-9]{64})\z") ||
                            !Int64.TryParse(fields[2], NumberStyles.None, CultureInfo.InvariantCulture, out size) || size != item.Size)
                            throw new InvalidDataException("旧版文档部件缺失或长度不符：" + item.Name);
                        var entry = archive.CreateEntry(item.Name, DocxFormat.LevelFor(item.Name));
                        entry.LastWriteTime = DocxFormat.ZipTime;
                        using (var sha = SHA256.Create())
                        {
                            using (var target = entry.Open())
                            {
                                var remaining = size;
                                while (remaining > 0)
                                {
                                    var count = batch.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
                                    if (count <= 0) throw new InvalidDataException("旧版文档部件内容不完整：" + item.Name);
                                    sha.TransformBlock(buffer, 0, count, null, 0);
                                    target.Write(buffer, 0, count);
                                    remaining -= count;
                                }
                            }
                            sha.TransformFinalBlock(new byte[0], 0, 0);
                            if (Hex(sha.Hash) != item.Sha) throw new InvalidDataException("旧版文档部件校验失败：" + item.Name);
                        }
                        if (batch.ReadByte() != '\n') throw new InvalidDataException("旧版文档部件内容不完整：" + item.Name);
                    }
                if (batch.ReadByte() >= 0 || DocxFormat.Hex(DocxFormat.Fingerprint(meta.Entries)) != meta.Fingerprint)
                    throw new InvalidDataException("旧版文档整体部件校验失败。");
            }, "cat-file", "--batch");
        }

        private static string ReadLegacyHeader(Stream batch)
        {
            var bytes = new List<byte>(128);
            for (;;)
            {
                var next = batch.ReadByte();
                if (next < 0 || bytes.Count > 8192) throw new InvalidDataException("旧版文档部件读取不完整。");
                if (next == '\n') return LenientUtf8.GetString(bytes.ToArray());
                bytes.Add((byte)next);
            }
        }

        // First streaming pass over the document file: computes the whole-file checksum and the
        // content-defined chunk plan without retaining any chunk content in memory.
        internal FileSnapshot PlanFile(string path)
        {
            EnsureRepository();
            DocxFormat.ValidateFile(path);
            var format = Text(Run(null, null, "rev-parse", "--show-object-format"));
            var before = new FileInfo(path);
            var stamp = before.LastWriteTimeUtc;
            var length = before.Length;
            SnapshotPlan plan;
            using (var stream = OpenShared(path)) plan = DocumentSnapshot.Plan(stream, format);
            var after = new FileInfo(path);
            if (after.Length != length || after.LastWriteTimeUtc != stamp || plan.Length != length)
                throw new InvalidOperationException("文档在提交准备期间被修改，请重新提交。");
            return new FileSnapshot { Plan = plan, WriteTimeUtc = stamp };
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
            var imported = Checked(git.Invoke(stream =>
            {
                for (var i = 0; i < missing.Length; i++)
                {
                    var mark = (i + 1).ToString(CultureInfo.InvariantCulture);
                    var header = Utf8.GetBytes("blob\nmark :" + mark + "\ndata " +
                        sizes[missing[i]].ToString(CultureInfo.InvariantCulture) + "\n");
                    stream.Write(header, 0, header.Length);
                    writeChunk(stream, missing[i]);
                    var tail = Utf8.GetBytes("\nget-mark :" + mark + "\n");
                    stream.Write(tail, 0, tail.Length);
                }
                var done = Utf8.GetBytes("done\n");
                stream.Write(done, 0, done.Length);
            }, null, Timeout.Infinite, true, "-c", "pack.threads=1", "fast-import", "--quiet", "--done"));
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
            if (!Exists) Checked(Invoke(null, null, "init", "--bare", "--initial-branch=main", "--template="));
        }

        internal string SnapshotChecksum(string revision)
        {
            revision = Id(revision);
            var stored = Locate(revision);
            if (stored.Kind == SnapshotKind.Chunked) return ReadManifest(stored).Checksum;
            using (var hashing = new HashingStream())
            {
                ExtractTo(revision, hashing);
                return hashing.Finish();
            }
        }

        // Streams the document from disk in bounded buffers instead of loading it into memory.
        public string CommitFile(string path, string message, string author, string expectedHead,
            string branch = null, string sourceBranch = null, string sourceTip = null, FileSnapshot planned = null)
        {
            branch = VerifyCommitState(message, expectedHead, branch, sourceBranch, sourceTip);
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
            branch = VerifyCommitState(message, expectedHead, branch, sourceBranch, sourceTip);
            if (sourceBranch == null && expectedHead != null && prepared.Checksum == SnapshotChecksum(expectedHead)) return null;
            return FinishCommit(prepared, message, author, expectedHead, branch, sourceBranch, sourceTip);
        }

        private string VerifyCommitState(string message, string expectedHead, string branch, string sourceBranch, string sourceTip)
        {
            if (String.IsNullOrWhiteSpace(message) || message.IndexOf('\0') >= 0) throw new ArgumentException("请填写有效的提交说明。");
            EnsureRepository();
            branch = branch ?? ActiveBranch();
            if (ActiveBranch() != branch || BranchTip(branch) != expectedHead ||
                (sourceBranch != null && (sourceBranch == branch || BranchTip(sourceBranch) != sourceTip)))
                throw new InvalidOperationException("分支或版本历史已变化。请刷新后再操作。");
            return branch;
        }

        private string FinishCommit(PreparedSnapshot prepared, string message, string author, string expectedHead,
            string branch, string sourceBranch, string sourceTip)
        {
            var reference = BranchRef(branch);
            var args = new List<string> { "commit-tree", Id(prepared.Tree) };
            if (expectedHead != null) args.AddRange(new[] { "-p", Id(expectedHead) });
            if (sourceBranch != null)
            {
                if (expectedHead == null || IsAncestor(sourceTip, expectedHead))
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
                Char.IsWhiteSpace(character) || Char.IsControl(character) || "~^:?*[\\\"<>|".IndexOf(character) >= 0))
                throw new ArgumentException("请输入有效的 Git 分支名，不包含空格或特殊符号。");
            // Loose refs are files: Windows reserved device names and trailing dots cannot be stored.
            foreach (var segment in branch.Split('/'))
            {
                var stem = segment.Split('.')[0];
                if (segment.EndsWith(".", StringComparison.Ordinal) ||
                    ReservedNames.Contains(stem.ToUpperInvariant()))
                    throw new ArgumentException("分支名不能使用 Windows 保留名称（如 CON、NUL、COM1），也不能以句点结尾。");
            }
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

        private static string Text(GitResult result) { return Utf8.GetString(Checked(result).Bytes).Trim(' ', '\r', '\n'); }
        private static string LenientText(GitResult result) { return LenientUtf8.GetString(Checked(result).Bytes).Trim(' ', '\r', '\n'); }
        private GitResult Run(byte[] input, string identity, params string[] args) { return Checked(Invoke(input, identity, args)); }
        private GitResult Invoke(byte[] input, string identity, params string[] args) { return git.Invoke(input, identity, args); }

        private static GitResult Checked(GitResult result)
        {
            if (result.Code != 0) throw new InvalidOperationException("Git 操作失败：" + result.Error.Trim());
            return result;
        }
    }

    // A planned chunked snapshot of a document file, plus the file stamp used to prove the file
    // did not change between the planning pass and the import pass.
    internal sealed class FileSnapshot
    {
        public SnapshotPlan Plan;
        public DateTime WriteTimeUtc;
    }

    // Write-only sink that computes the SHA-256 of everything written to it.
    internal sealed class HashingStream : Stream
    {
        private readonly SHA256 sha = SHA256.Create();
        private long length;

        internal string Finish()
        {
            sha.TransformFinalBlock(new byte[0], 0, 0);
            return GitHistory.Hex(sha.Hash);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            sha.TransformBlock(buffer, offset, count, null, 0);
            length += count;
        }

        public override bool CanRead { get { return false; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return true; } }
        public override long Length { get { return length; } }
        public override long Position { get { return length; } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        protected override void Dispose(bool disposing) { if (disposing) sha.Dispose(); base.Dispose(disposing); }
    }

    // Write-only sink used when only the verification performed during a restore matters.
    internal sealed class DiscardStream : Stream
    {
        private long length;
        public override void Write(byte[] buffer, int offset, int count) { length += count; }
        public override bool CanRead { get { return false; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return true; } }
        public override long Length { get { return length; } }
        public override long Position { get { return length; } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
    }
}
