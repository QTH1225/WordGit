using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace WordGit
{
    internal static class AppPaths
    {
        public static readonly string Root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordGit");
        public static readonly string Repositories = Path.Combine(Root, "repositories");
        public static readonly string Sessions = Path.Combine(Root, "sessions");
        public static readonly string LogFile = Path.Combine(Root, "logs", "wordgit.log");
        public static string NewSession()
        {
            return Directory.CreateDirectory(Path.Combine(Sessions, Guid.NewGuid().ToString("N"))).FullName;
        }

        internal static bool IsSessionFile(string path)
        {
            return path.StartsWith(Sessions + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }

        public static int CleanSessions(int days) { return CleanSessions(Sessions, days); }

        internal static int CleanSessions(string sessionsRoot, int days)
        {
            var removed = 0;
            var cutoff = DateTime.UtcNow.AddDays(-days);
            try
            {
                foreach (var folder in Directory.GetDirectories(sessionsRoot))
                {
                    try
                    {
                        if (Directory.GetLastWriteTimeUtc(folder) >= cutoff) continue;
                        if (Directory.GetFiles(folder, "*.wordgit", SearchOption.AllDirectories).Length > 0) continue;
                        Directory.Delete(folder, true);
                        removed++;
                    }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            return removed;
        }

        public static void LogError(string message) { AppendLog(LogFile, message); }

        // The profile folder carries the Windows user name; logs are meant to be shareable in bug reports.
        internal static string Redact(string message)
        {
            var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (String.IsNullOrEmpty(profile) || message == null) return message;
            var result = new System.Text.StringBuilder();
            var index = 0;
            for (;;)
            {
                var found = message.IndexOf(profile, index, StringComparison.OrdinalIgnoreCase);
                if (found < 0) break;
                result.Append(message, index, found - index).Append("%USERPROFILE%");
                index = found + profile.Length;
            }
            return result.Append(message, index, message.Length - index).ToString();
        }

        internal static void AppendLog(string file, string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(file));
                if (File.Exists(file) && new FileInfo(file).Length > 1024 * 1024)
                {
                    var rolled = file + ".1";
                    try { if (File.Exists(rolled)) File.Delete(rolled); } catch (IOException) { }
                    File.Move(file, rolled);
                }
                File.AppendAllText(file, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + Redact(message) + Environment.NewLine,
                    new System.Text.UTF8Encoding(false));
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    internal enum WorkingCopyState { NoCommit, Unsaved, Changed, Clean }

    internal sealed class DocumentComment
    {
        public int Index, Page, ScopeStart;
        public string Author, Text;
        public string Caption { get { return Author + " · 第 " + Page + " 页"; } }
    }

    internal sealed class DocumentReviewItem
    {
        public DocumentComment Comment;
        public int Index, Page, ScopeStart, ScopeEnd, StoryType, StoryIndex, RevisionType;
        public string Author, Text, RangeText;
        public bool IsComment { get { return Comment != null; } }
        public string Kind
        {
            get
            {
                if (IsComment) return "批注";
                switch (RevisionType)
                {
                    case 1: return "新增";
                    case 2: return "删除";
                    case 3: case 4: case 8: case 10: case 11: case 12: case 13: return "格式";
                    case 5: return "域变更";
                    case 6: case 7: case 20: case 21: return "冲突";
                    case 9: return "替换";
                    case 14: return "移出";
                    case 15: return "移入";
                    case 16: return "新增单元格";
                    case 17: return "删除单元格";
                    case 18: case 19: return "单元格变更";
                    default: return "修订";
                }
            }
        }
        public string Caption { get { return Kind + " · " + Author + " · 第 " + Page + " 页"; } }
        public string Key { get { return (IsComment ? "comment" : "revision") + ":" + StoryType + ":" + StoryIndex + ":" + Index; } }

        public static DocumentReviewItem ForComment(DocumentComment comment)
        {
            return new DocumentReviewItem { Comment = comment, Index = comment.Index, Page = comment.Page,
                ScopeStart = comment.ScopeStart, StoryType = 1, Author = comment.Author, Text = comment.Text };
        }
    }

    internal sealed class WorkingCopy
    {
        private readonly string path;
        private readonly Func<bool> isSaved;
        private readonly Func<string> readHead;
        private string headHash;
        private DateTime lastWrite;
        private long lastLength;
        private bool measured;
        private WorkingCopyState lastState;

        public WorkingCopy(string path, Func<bool> isSaved, Func<string> readHead)
        {
            this.path = path;
            this.isSaved = isSaved;
            this.readHead = readHead;
        }

        public WorkingCopyState State()
        {
            if (!isSaved())
            {
                measured = false;
                return WorkingCopyState.Unsaved;
            }
            if (readHead == null) return WorkingCopyState.NoCommit;

            var file = new FileInfo(path);
            var writeTime = file.LastWriteTimeUtc;
            var length = file.Length;
            if (measured && writeTime == lastWrite && length == lastLength) return lastState;

            if (headHash == null) headHash = readHead();
            lastState = WordDocuments.FileChecksum(path) == headHash ? WorkingCopyState.Clean : WorkingCopyState.Changed;
            lastWrite = writeTime;
            lastLength = length;
            measured = true;
            return lastState;
        }
    }

    // Result of the background half of "restore an old version": the current draft is committed
    // and the old version staged, so only the file swap remains for Word's thread.
    internal sealed class RestorePlan
    {
        public string Path, Revision, Author, Branch, Head, CurrentHash, OldHash, Staged;
        public PreparedSnapshot Prepared;
        public bool Unchanged;
    }

    internal sealed class WordDocuments
    {
        private readonly dynamic word;
        public WordDocuments(object application) { word = application; }
        public IWin32Window Owner
        {
            get { return word.Windows.Count == 0 ? null : DialogPlacement.RootOwner(new WordWindow(new IntPtr((int)word.ActiveWindow.Hwnd))); }
        }

        public object Current()
        {
            if (word.Documents.Count == 0) throw new InvalidOperationException("请先在 Word 打开一份 .docx。");
            return word.ActiveDocument;
        }

        internal static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                try { Marshal.ReleaseComObject(value); }
                catch (ArgumentException) { }
        }

        public static List<DocumentComment> Comments(object value)
        {
            dynamic document = value;
            var result = new List<DocumentComment>();
            dynamic all = document.Comments;
            try
            {
                for (var index = 1; index <= (int)all.Count; index++)
                {
                    dynamic comment = null, scope = null, range = null;
                    try
                    {
                        comment = all[index];
                        scope = comment.Scope;
                        range = comment.Range;
                        result.Add(new DocumentComment
                        {
                            Index = index, Author = (string)comment.Author, Text = ((string)range.Text).Trim(),
                            ScopeStart = (int)scope.Start, Page = (int)scope.Information[3]
                        });
                    }
                    finally { Release(range); Release(scope); Release(comment); }
                }
            }
            finally { Release(all); }
            return result;
        }

        internal static int DocumentRevisionCount(object value)
        {
            dynamic revisions = ((dynamic)value).Revisions;
            try { return (int)revisions.Count; }
            finally { Release(revisions); }
        }

        // Changes whenever comments or revisions are added, removed or resolved, or the text
        // length changes; cheap enough to poll, unlike ReviewItems, which repaginates.
        public static string ReviewSignature(object value)
        {
            dynamic document = value;
            dynamic comments = null, revisions = null, content = null;
            try
            {
                comments = document.Comments;
                revisions = document.Revisions;
                content = document.Content;
                return (int)comments.Count + ":" + (int)revisions.Count + ":" + (int)content.End;
            }
            finally { Release(content); Release(revisions); Release(comments); }
        }

        public void GoToComment(object value, DocumentComment selected)
        {
            dynamic document = value;
            if (selected.Index > (int)document.Comments.Count)
                throw new InvalidOperationException("批注已更新，请刷新批注列表。");
            dynamic comment = document.Comments[selected.Index];
            if (((string)comment.Range.Text).Trim() != selected.Text || (string)comment.Author != selected.Author)
                throw new InvalidOperationException("批注已更新，请刷新批注列表。");
            document.Activate();
            dynamic scope = comment.Scope;
            scope.Select();
            word.ActiveWindow.ScrollIntoView(scope, true);
        }

        public static List<DocumentReviewItem> ReviewItems(object value)
        {
            dynamic document = value;
            var result = new List<DocumentReviewItem>();
            foreach (var comment in Comments(value)) result.Add(DocumentReviewItem.ForComment(comment));
            var seen = new HashSet<string>(StringComparer.Ordinal);
            dynamic stories = document.StoryRanges;
            try
            {
                foreach (dynamic first in stories)
                {
                    dynamic story = first;
                    var storyType = (int)story.StoryType;
                    var storyIndex = 0;
                    while (story != null)
                    {
                        if (storyType != 4) // Comment text is already represented by the comment entry.
                            CollectRevisions(story, storyType, storyIndex, seen, result);
                        dynamic next = story.NextStoryRange;
                        Release(story);
                        story = next;
                        storyIndex++;
                    }
                }
            }
            finally { Release(stories); }
            return result;
        }

        private static void CollectRevisions(dynamic story, int storyType, int storyIndex, HashSet<string> seen, List<DocumentReviewItem> result)
        {
            dynamic changes = story.Revisions;
            try
            {
                var count = (int)changes.Count;
                for (var index = 1; index <= count; index++)
                {
                    dynamic change = null, scope = null, anchor = null;
                    try
                    {
                        change = changes[index];
                        scope = change.Range;
                        anchor = scope.Duplicate;
                        anchor.Collapse(1);
                        var text = (string)scope.Text ?? "";
                        var format = (string)change.FormatDescription ?? "";
                        var description = String.IsNullOrWhiteSpace(format) ? text : format + Environment.NewLine + text;
                        var entry = new DocumentReviewItem { Index = index, Page = (int)anchor.Information[3],
                            ScopeStart = (int)scope.Start, ScopeEnd = (int)scope.End, StoryType = storyType, StoryIndex = storyIndex,
                            RevisionType = (int)change.Type, Author = (string)change.Author, RangeText = text,
                            Text = String.IsNullOrWhiteSpace(description) ? "段落、空白或结构变更" : description.Trim() };
                        if (seen.Add(entry.Key)) result.Add(entry);
                    }
                    finally { Release(anchor); Release(scope); Release(change); }
                }
            }
            finally { Release(changes); }
        }

        public void GoToReviewItem(object value, DocumentReviewItem selected)
        {
            if (selected.IsComment) { GoToComment(value, selected.Comment); return; }
            dynamic change = FindRevision(value, selected);
            dynamic document = value;
            document.Activate();
            word.ActiveWindow.View.ShowRevisionsAndComments = true;
            dynamic scope = change.Range;
            scope.Select();
            word.ActiveWindow.ScrollIntoView(scope, true);
        }

        private static object FindRevision(object value, DocumentReviewItem selected)
        {
            if (selected == null || selected.IsComment)
                throw new InvalidOperationException("请先选择一条修订。");
            dynamic document = value;
            dynamic story = document.StoryRanges[selected.StoryType];
            for (var index = 0; index < selected.StoryIndex && story != null; index++) story = story.NextStoryRange;
            if (story == null || selected.Index < 1 || selected.Index > (int)story.Revisions.Count)
                throw new InvalidOperationException("修订已更新，请刷新批注与修订列表。");
            dynamic change = story.Revisions[selected.Index];
            dynamic scope = change.Range;
            if ((int)change.Type != selected.RevisionType || (string)change.Author != selected.Author ||
                ((string)scope.Text ?? "") != selected.RangeText || (int)scope.Start != selected.ScopeStart || (int)scope.End != selected.ScopeEnd)
                throw new InvalidOperationException("修订已更新，请刷新批注与修订列表。");
            return change;
        }

        public void ResolveReviewItem(object value, DocumentReviewItem selected, bool accept)
        {
            dynamic document = value;
            if ((bool)document.ReadOnly)
                throw new InvalidOperationException("当前文档为只读，无法处理修订。");
            dynamic change = FindRevision(value, selected);
            document.Activate();
            if (accept) change.Accept();
            else change.Reject();
        }

        public static string LocalPath(object value, bool allowSave = false)
        {
            dynamic doc = value;
            if (String.IsNullOrEmpty((string)doc.Path) && allowSave) doc.Save();
            var filename = (string)doc.FullName;
            if (String.IsNullOrEmpty((string)doc.Path)) throw new InvalidOperationException("请先将这份文档另存为 .docx。");
            if (filename.Contains("://") || !Path.IsPathRooted(filename))
                throw new InvalidOperationException("WordGit 需要可直接读取的本地 .docx。OneDrive/SharePoint 文件请使用同步到本机的文件夹中的副本，或另存一份到本地。");
            if (!String.Equals(Path.GetExtension(filename), ".docx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("WordGit 只支持 .docx。");
            return Path.GetFullPath(filename);
        }

        public static string SaveDocument(object value)
        {
            dynamic doc = value;
            var path = LocalPath(value, true);
            if (!(bool)doc.Saved)
            {
                if ((bool)doc.ReadOnly) throw new InvalidOperationException("文档是只读的，当前修改无法保存。请先另存为 .docx。");
                doc.Save();
            }
            if (!(bool)doc.Saved) throw new InvalidOperationException("Word 尚未保存成功，未创建提交。");
            return path;
        }

        public static string FileChecksum(string path)
        {
            using (var sha = SHA256.Create())
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                return GitHistory.Hex(sha.ComputeHash(stream));
        }

        // The target only appears once the restored version has been fully verified, so a failed
        // or interrupted restore never leaves a truncated document under the chosen name.
        public static void ExtractToFile(GitHistory repo, string revision, string path)
        {
            if (File.Exists(path)) throw new IOException("目标文件已存在：" + path);
            var temporary = path + "." + Guid.NewGuid().ToString("N").Substring(0, 8) + ".wordgit-tmp";
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    repo.ExtractTo(revision, stream);
                    stream.Flush(true);
                }
                File.Move(temporary, path);
            }
            finally { TryRemove(temporary); }
        }

        public object Open(string path, bool readOnly, bool visible = true)
        {
            return word.Documents.Open(FileName: path, ConfirmConversions: false, ReadOnly: readOnly,
                AddToRecentFiles: false, Visible: visible);
        }

        public object ReviewFiles(GitHistory repo, string originalRevision, string revisedRevision, bool merge, string author = null)
        {
            var folder = AppPaths.NewSession();
            var leftFile = Path.Combine(folder, "基准版本.docx");
            var rightFile = Path.Combine(folder, "修订版本.docx");
            ExtractToFile(repo, originalRevision, leftFile);
            ExtractToFile(repo, revisedRevision, rightFile);
            return ReviewStaged(folder, leftFile, rightFile, merge, author);
        }

        // Compares or merges two documents given as files; both are copied first, so Word never
        // holds the caller's originals open.
        public object ReviewPaths(string original, string revised, bool merge, string author = null)
        {
            var folder = AppPaths.NewSession();
            var leftFile = Path.Combine(folder, "基准版本.docx");
            var rightFile = Path.Combine(folder, "修订版本.docx");
            DocxFormat.ValidateFile(original);
            DocxFormat.ValidateFile(revised);
            CopyShared(original, leftFile);
            CopyShared(revised, rightFile);
            return ReviewStaged(folder, leftFile, rightFile, merge, author);
        }

        public object CompareCurrentExtracted(GitHistory repo, string originalRevision, object current)
        {
            var folder = AppPaths.NewSession();
            var leftFile = Path.Combine(folder, "基准版本.docx");
            ExtractToFile(repo, originalRevision, leftFile);
            dynamic document = current;
            string rightFile;
            if ((bool)document.Saved && !String.IsNullOrEmpty((string)document.Path))
            {
                // The saved file is the current draft; copying it avoids materialising the whole
                // document (images included) as one WordOpenXML string.
                rightFile = Path.Combine(folder, "当前稿.docx");
                CopyShared(LocalPath(current), rightFile);
            }
            else
            {
                rightFile = Path.Combine(folder, "当前稿.xml");
                File.WriteAllText(rightFile, (string)document.WordOpenXML, new System.Text.UTF8Encoding(false));
            }
            return ReviewStaged(folder, leftFile, rightFile, false, null);
        }

        private object ReviewReturnAgainstBase(GitHistory repo, string basisRevision, string externalPath)
        {
            var folder = AppPaths.NewSession();
            var leftFile = Path.Combine(folder, "基准版本.docx");
            var rightFile = Path.Combine(folder, "修订版本.docx");
            ExtractToFile(repo, basisRevision, leftFile);
            CopyShared(externalPath, rightFile);
            return ReviewStaged(folder, leftFile, rightFile, true, null);
        }

        private static void CopyShared(string source, string destination)
        {
            using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                input.CopyTo(output);
        }

        private object ReviewStaged(string folder, string leftFile, string rightFile, bool merge, string author)
        {
            dynamic left = null, right = null;
            var screenUpdating = (bool)word.ScreenUpdating;
            word.ScreenUpdating = false;
            try
            {
                left = Open(leftFile, false, false);
                right = Open(rightFile, false, false);
                // Word can reject AcceptAllRevisions on an invisible, inactive document.
                left.Windows[1].Visible = true;
                left.Activate();
                if (!merge) { left.TrackRevisions = false; left.AcceptAllRevisions(); }
                right.Windows[1].Visible = true;
                right.Activate();
                if (!merge) { right.TrackRevisions = false; right.AcceptAllRevisions(); }
                dynamic result;
                if (merge)
                    result = word.MergeDocuments(OriginalDocument: left, RevisedDocument: right, Destination: 2,
                        CompareFormatting: true, CompareTables: true, CompareHeaders: true, CompareFootnotes: true,
                        CompareTextboxes: true, CompareFields: true, CompareComments: true, FormatFrom: 0);
                else
                    result = word.CompareDocuments(OriginalDocument: left, RevisedDocument: right, Destination: 2,
                        CompareFormatting: true, CompareTables: true, CompareHeaders: true, CompareFootnotes: true,
                        CompareTextboxes: true, CompareFields: true, CompareComments: true, RevisedAuthor: author ?? "WordGit",
                        IgnoreAllComparisonWarnings: true);
                result.Activate();
                return result;
            }
            finally
            {
                CloseReviewInput(right);
                CloseReviewInput(left);
                word.ScreenUpdating = screenUpdating;
                TryRemove(leftFile);
                TryRemove(rightFile);
                try { Directory.Delete(folder); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        public object ReviewBranchesExtracted(GitHistory repo, string basisRevision, string targetRevision, string incomingRevision, string targetBranch = "当前分支", string sourceBranch = "来源分支", bool preserveIncomingRevisions = false)
        {
            if (repo.SnapshotChecksum(basisRevision) == repo.SnapshotChecksum(targetRevision))
                return ReviewFiles(repo, targetRevision, incomingRevision, true);
            var folder = AppPaths.NewSession();
            var targetFile = Path.Combine(folder, "分支修订-当前.docx");
            var sourceFile = Path.Combine(folder, "分支修订-来源.docx");
            SaveBranchRevisionsToFile(repo, basisRevision, targetRevision, targetBranch, targetFile);
            SaveBranchRevisionsToFile(repo, basisRevision, incomingRevision, sourceBranch, sourceFile,
                repo.ImportedReviewBase(incomingRevision) != null || preserveIncomingRevisions);
            return ReviewStaged(folder, targetFile, sourceFile, true, null);
        }

        internal object ReviewReturnedDraft(GitHistory repo, string basisRevision, string targetRevision, string externalPath, string branch)
        {
            DocxFormat.ValidateFile(externalPath);
            if (repo.SnapshotChecksum(basisRevision) == repo.SnapshotChecksum(targetRevision))
                return ReviewReturnAgainstBase(repo, basisRevision, externalPath);

            var folder = AppPaths.NewSession();
            var targetFile = Path.Combine(folder, "当前稿修订.docx");
            var returnedFile = Path.Combine(folder, "返稿修订.docx");
            SaveBranchRevisionsToFile(repo, basisRevision, targetRevision, branch, targetFile);
            SaveReviewToFile(ReviewReturnAgainstBase(repo, basisRevision, externalPath), returnedFile);
            return ReviewStaged(folder, targetFile, returnedFile, true, null);
        }

        private void SaveBranchRevisionsToFile(GitHistory repo, string basisRevision, string versionRevision, string author, string outputPath,
            bool preserveRevisions = false)
        {
            SaveReviewToFile(ReviewFiles(repo, basisRevision, versionRevision, preserveRevisions, author), outputPath);
        }

        private void SaveReviewToFile(object document, string outputPath)
        {
            try
            {
                ((dynamic)document).SaveAs2(FileName: outputPath, FileFormat: 16, AddToRecentFiles: false);
            }
            finally
            {
                CloseReviewInput(document);
            }
        }

        // Background half of a restore (no Word COM). The caller holds the repository lock.
        public static RestorePlan PrepareRestore(GitHistory repo, string path, string revision, string author)
        {
            var plan = new RestorePlan { Path = path, Revision = revision, Author = author,
                CurrentHash = FileChecksum(path), OldHash = repo.SnapshotChecksum(revision) };
            if (plan.CurrentHash == plan.OldHash) { plan.Unchanged = true; return plan; }
            plan.Branch = repo.ActiveBranch();
            var head = repo.Head();
            plan.Head = repo.CommitFile(path, "恢复前保存当前稿", author, head, plan.Branch) ?? head;
            plan.Prepared = repo.PreparedFromRevision(revision);
            plan.Staged = Path.Combine(AppPaths.NewSession(), "恢复版本.docx");
            ExtractToFile(repo, revision, plan.Staged);
            return plan;
        }

        // Word half of a restore. The caller holds the repository lock.
        public string FinishRestore(GitHistory repo, RestorePlan plan)
        {
            if (plan.Unchanged) return "当前稿已经与所选旧版一致。";
            string saved = null;
            ReplaceWorkingCopy(repo, plan.Path, plan.CurrentHash, plan.Staged, plan.OldHash,
                () => saved = repo.CommitPrepared(plan.Prepared, "恢复版本 " + plan.Revision.Substring(0, 8), plan.Author, plan.Head, plan.Branch));
            return saved == null ? "已恢复旧版。" : "已恢复并创建新提交 " + saved.Substring(0, 8) + "。";
        }

        public static void EnsureRestorable(object source)
        {
            var path = LocalPath(source);
            EnsureReplaceable(path);
            if ((bool)((dynamic)source).ReadOnly) throw new InvalidOperationException("当前文档只读，请使用「另存旧版」。");
        }

        public string Restore(object source, GitHistory repo, string revision, string author)
        {
            EnsureRestorable(source);
            var path = SaveDocument(source);
            return FinishRestore(repo, PrepareRestore(repo, path, revision, author));
        }

        public static void EnsureReplaceable(string path)
        {
            if (path.StartsWith(@"\\", StringComparison.Ordinal))
                throw new InvalidOperationException("请在本地文件副本上切换、恢复或合并分支。");
            foreach (var folder in new[] { Environment.GetEnvironmentVariable("OneDrive"),
                Environment.GetEnvironmentVariable("OneDriveCommercial"), Environment.GetEnvironmentVariable("OneDriveConsumer") })
                if (!String.IsNullOrEmpty(folder) && path.StartsWith(folder.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("为避免覆盖共编修改，请在 OneDrive 同步目录之外的本地副本上切换、恢复或合并分支。");
        }

        // Swaps the working copy for a verified replacement and then publishes the matching ref
        // change. A journal in the repository makes the pair recoverable if Word dies in between.
        public void ReplaceWorkingCopy(GitHistory repo, string path, string currentHash, string replacementFile, string replacementHash, Action publish)
        {
            EnsureReplaceable(path);
            var leftover = WorkingCopyJournal.Recover(repo);
            if (leftover != null) throw new InvalidOperationException(leftover + Environment.NewLine + "请确认当前稿后重试。");
            dynamic doc = null;
            foreach (var candidate in word.Documents)
                if (String.Equals((string)((dynamic)candidate).FullName, path, StringComparison.OrdinalIgnoreCase)) { doc = candidate; break; }
            if (doc != null && ((bool)doc.ReadOnly || !(bool)doc.Saved))
                throw new InvalidOperationException("当前稿有未保存的修改，或处于只读状态。请先处理后重试。");
            if (FileChecksum(path) != currentHash)
                throw new InvalidOperationException("磁盘上的当前稿已变化，未替换文件。");
            if (currentHash == replacementHash) { publish(); return; }
            var folderPath = Path.GetDirectoryName(path);
            CleanRestoreLeftovers(folderPath);
            var staged = Path.Combine(folderPath, ".wordgit-restore-" + Guid.NewGuid().ToString("N") + ".tmp");
            var backup = staged + ".bak";
            File.Copy(replacementFile, staged);
            if (FileChecksum(staged) != replacementHash)
            {
                TryRemove(staged);
                throw new InvalidDataException("替换内容在操作期间发生变化，未覆盖当前稿。");
            }
            WorkingCopyJournal.Write(repo, new WorkingCopyJournal.Entry { Document = path, Backup = backup,
                OriginalHash = currentHash, ReplacementHash = replacementHash, References = repo.ReferenceState() });
            var closed = false;
            var replaced = false;
            try
            {
                if (doc != null) { doc.Close(SaveChanges: 0); closed = true; }
                if (FileChecksum(path) != currentHash)
                    throw new InvalidOperationException("磁盘文件在操作期间发生变化，未覆盖。请重新打开后重试。");
                File.Replace(staged, path, backup);
                replaced = true;
                try { publish(); }
                catch
                {
                    if (FileChecksum(path) != replacementHash)
                        throw new IOException("当前稿在操作期间被其他程序修改，原稿备份保留在：" + backup);
                    File.Replace(backup, path, null);
                    replaced = false;
                    throw;
                }
                TryRemove(backup);
                replaced = false;
            }
            finally
            {
                // A replaced file whose rollback failed keeps its journal for recovery on restart.
                if (!replaced) WorkingCopyJournal.Clear(repo);
                TryRemove(staged);
                if (closed) Open(path, false);
            }
        }

        internal static void CleanRestoreLeftovers(string folder)
        {
            var now = DateTime.UtcNow;
            foreach (var rule in new[] { new { Pattern = ".wordgit-restore-*.tmp", Age = TimeSpan.FromDays(1) },
                // Backups are the only copy of a draft when a rollback failed; keep them for a month.
                new { Pattern = ".wordgit-restore-*.tmp.bak", Age = TimeSpan.FromDays(30) } })
            {
                try
                {
                    foreach (var file in Directory.GetFiles(folder, rule.Pattern))
                    {
                        if (!file.EndsWith(rule.Pattern.Substring(rule.Pattern.LastIndexOf('*') + 1), StringComparison.OrdinalIgnoreCase)) continue;
                        try { if (File.GetLastWriteTimeUtc(file) < now - rule.Age) File.Delete(file); }
                        catch (IOException) { }
                        catch (UnauthorizedAccessException) { }
                    }
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }

        private static void TryRemove(string file)
        {
            try { File.Delete(file); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }

        private static void CloseReviewInput(object document)
        {
            if (document != null)
                try { ((dynamic)document).Close(SaveChanges: 0); } catch (COMException) { }
        }

        private sealed class WordWindow : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public WordWindow(IntPtr handle) { Handle = handle; }
        }
    }
}
