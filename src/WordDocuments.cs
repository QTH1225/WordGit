using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
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
        public string Author, Text, ScopeText;
        public DateTime Date;
        public bool Resolved;
        public string Caption { get { return Author + " · 第 " + Page + " 页"; } }
    }

    internal sealed class DocumentReviewItem
    {
        public DocumentComment Comment;
        public int Index, Page, ScopeStart, ScopeEnd, StoryType, StoryIndex, RevisionType;
        public string Author, Text, RangeText, RevisionFormat;
        public DateTime Date;
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
        private readonly Func<string> readState;
        private DateTime modified, created;
        private long length = -1;
        private WorkingCopyState cached;

        public WorkingCopy(string path, Func<string> readState)
        {
            this.path = path;
            this.readState = readState;
        }

        public WorkingCopyState State()
        {
            var file = new FileInfo(path);
            var size = file.Length;
            var lastWrite = file.LastWriteTimeUtc;
            var creation = file.CreationTimeUtc;
            if (size == length && lastWrite == modified && creation == created)
                return cached;

            // Capture the stamp before the service read so a save during that read is checked again.
            var state = readState();
            cached = state == "clean" ? WorkingCopyState.Clean :
                state == "noCommit" ? WorkingCopyState.NoCommit : WorkingCopyState.Changed;
            length = size;
            modified = lastWrite;
            created = creation;
            return cached;
        }
    }

    internal sealed partial class WordDocuments
    {
        private readonly dynamic word;
        public bool IsWpsHost { get; private set; }

        public bool SupportsTrackedMerge { get; private set; }

        public void EnsureTrackedMergeSupported()
        {
            if (!SupportsTrackedMerge)
                throw new InvalidOperationException("当前 WPS 文字暂不支持生成可保留双方修改的修订合并稿。可将返稿保存为独立分支，并对比历史版本。");
        }

        public WordDocuments(object application)
        {
            word = application;
            var path = (string)word.Path;
            IsWpsHost = !String.IsNullOrEmpty(path) && File.Exists(Path.Combine(path, "wps.exe"));
            SupportsTrackedMerge = !IsWpsHost;
        }
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

        private static object RequireReviewObject(object value, string part)
        {
            if (value == null) throw new InvalidOperationException("Word 暂未返回" + part + "，请等文档加载完成后重试。");
            return value;
        }

        public static List<DocumentComment> Comments(object value, bool includeReviewStatus = false)
        {
            dynamic document = RequireReviewObject(value, "文档");
            var result = new List<DocumentComment>();
            dynamic all = RequireReviewObject((object)document.Comments, "批注列表");
            try
            {
                for (var index = 1; index <= (int)all.Count; index++)
                {
                    dynamic comment = null, scope = null, range = null;
                    try
                    {
                        comment = RequireReviewObject((object)all[index], "批注");
                        scope = RequireReviewObject((object)comment.Scope, "批注锚点");
                        range = RequireReviewObject((object)comment.Range, "批注内容");
                        result.Add(new DocumentComment
                        {
                            Index = index, Author = (string)comment.Author, Text = ((string)range.Text ?? "").Trim(),
                            ScopeStart = (int)scope.Start, Page = (int)scope.Information[3],
                            ScopeText = includeReviewStatus ? ((string)scope.Text ?? "").Trim() : null,
                            Date = includeReviewStatus ? (DateTime)comment.Date : default(DateTime),
                            Resolved = includeReviewStatus && CommentResolved(comment)
                        });
                    }
                    finally { Release(range); Release(scope); Release(comment); }
                }
            }
            finally { Release(all); }
            return result;
        }

        private static bool CommentResolved(dynamic comment)
        {
            try { return (bool)comment.Done; }
            catch (COMException) { return false; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return false; }
        }

        internal static int DocumentRevisionCount(object value)
        {
            dynamic document = RequireReviewObject(value, "文档");
            dynamic revisions = RequireReviewObject((object)document.Revisions, "修订列表");
            try { return (int)revisions.Count; }
            finally { Release(revisions); }
        }

        // Changes whenever comments or revisions are added, removed or resolved, or the text
        // length changes; cheap enough to poll, unlike ReviewItems, which repaginates.
        public static string ReviewSignature(object value)
        {
            dynamic document = RequireReviewObject(value, "文档");
            dynamic comments = null, revisions = null, content = null;
            try
            {
                comments = RequireReviewObject((object)document.Comments, "批注列表");
                revisions = RequireReviewObject((object)document.Revisions, "修订列表");
                content = RequireReviewObject((object)document.Content, "文档内容");
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
            ActivateIfNeeded(value);
            dynamic scope = comment.Scope;
            scope.Select();
            word.ActiveWindow.ScrollIntoView(scope, true);
        }

        public static List<DocumentReviewItem> ReviewItems(object value, bool includeResolvedComments = true, bool includeRevisionDates = false)
        {
            dynamic document = RequireReviewObject(value, "文档");
            var result = new List<DocumentReviewItem>();
            foreach (var comment in Comments(value, !includeResolvedComments))
                if (includeResolvedComments || !comment.Resolved) result.Add(DocumentReviewItem.ForComment(comment));
            var seen = new HashSet<string>(StringComparer.Ordinal);
            dynamic stories = RequireReviewObject((object)document.StoryRanges, "文档范围");
            try
            {
                foreach (object first in stories)
                {
                    dynamic story = RequireReviewObject(first, "文档范围项目");
                    try
                    {
                        var storyType = (int)story.StoryType;
                        var storyIndex = 0;
                        while (story != null)
                        {
                            if (storyType != 4) // Comment text is already represented by the comment entry.
                                CollectRevisions(story, storyType, storyIndex, seen, result, includeRevisionDates);
                            dynamic next = story.NextStoryRange;
                            Release(story);
                            story = next;
                            storyIndex++;
                        }
                    }
                    finally { Release(story); }
                }
            }
            finally { Release(stories); }
            return result;
        }

        private static void CollectRevisions(dynamic story, int storyType, int storyIndex, HashSet<string> seen,
            List<DocumentReviewItem> result, bool includeRevisionDates)
        {
            dynamic changes = RequireReviewObject((object)story.Revisions, "修订范围");
            try
            {
                var count = (int)changes.Count;
                for (var index = 1; index <= count; index++)
                {
                    dynamic change = null, scope = null, anchor = null;
                    try
                    {
                        change = RequireReviewObject((object)changes[index], "修订");
                        scope = RequireReviewObject((object)change.Range, "修订位置");
                        anchor = RequireReviewObject((object)scope.Duplicate, "修订锚点");
                        anchor.Collapse(1);
                        var text = (string)scope.Text ?? "";
                        var format = (string)change.FormatDescription ?? "";
                        var description = String.IsNullOrWhiteSpace(format) ? text : format + Environment.NewLine + text;
                        var entry = new DocumentReviewItem { Index = index, Page = (int)anchor.Information[3],
                            ScopeStart = (int)scope.Start, ScopeEnd = (int)scope.End, StoryType = storyType, StoryIndex = storyIndex,
                            RevisionType = (int)change.Type, Author = (string)change.Author, RangeText = text,
                            Date = includeRevisionDates ? (DateTime)change.Date : default(DateTime), RevisionFormat = format,
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
            ActivateIfNeeded(value);
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
            ActivateIfNeeded(value);
            if (accept) change.Accept();
            else change.Reject();
        }

        internal void ActivateIfNeeded(object document)
        {
            if (word.Windows.Count == 0 || !ReferenceEquals((object)word.ActiveDocument, document))
                ((dynamic)document).Activate();
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
                return Hex(sha.ComputeHash(stream));
        }

        private static string Hex(byte[] bytes)
        {
            var text = new StringBuilder(bytes.Length * 2);
            foreach (var value in bytes) text.Append(value.ToString("x2", CultureInfo.InvariantCulture));
            return text.ToString();
        }

        public object Open(string path, bool readOnly, bool visible = true)
        {
            return word.Documents.Open(FileName: path, ConfirmConversions: false, ReadOnly: readOnly,
                AddToRecentFiles: false, Visible: visible);
        }

        public object ReviewFiles(string original, string revised, bool merge, string author = null, bool collaboration = false)
        {
            if (merge) EnsureTrackedMergeSupported();
            var folder = AppPaths.NewSession();
            var leftFile = Path.Combine(folder, "基准版本.docx");
            var rightFile = Path.Combine(folder, "修订版本.docx");
            CopyShared(original, leftFile);
            CopyShared(revised, rightFile);
            return ReviewStaged(folder, leftFile, rightFile, merge, author, collaboration);
        }

        public object ReviewPaths(string original, string revised, bool merge, string author = null)
        {
            return ReviewFiles(original, revised, merge, author);
        }

        /// Compares a revision the kernel already extracted with the unsaved or saved current draft.
        internal object CompareSnapshot(string original, string currentSnapshot)
        {
            var folder = Path.GetDirectoryName(currentSnapshot);
            var leftFile = Path.Combine(folder, "基准版本.docx");
            try
            {
                CopyShared(original, leftFile);
                return ReviewStaged(folder, leftFile, currentSnapshot, false, null);
            }
            finally { TryRemove(currentSnapshot); }
        }

        internal static string CaptureComparisonSource(object current, string folder)
        {
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
            return rightFile;
        }

        internal void BuildVersionReviewFile(string original, string revised, string snapshot, bool merge, string output)
        {
            var review = snapshot == null ? ReviewFiles(original, revised, merge) : CompareSnapshot(original, snapshot);
            SaveReviewToFile(review, output);
        }

        private static void CopyShared(string source, string destination)
        {
            using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                input.CopyTo(output);
        }

        private object ReviewStaged(string folder, string leftFile, string rightFile, bool merge, string author, bool collaboration = false)
        {
            dynamic left = null, right = null;
            var screenUpdating = (bool)word.ScreenUpdating;
            word.ScreenUpdating = false;
            try
            {
                left = Open(leftFile, false, false);
                right = Open(rightFile, false, false);
                PrepareReviewInput((object)left, merge);
                PrepareReviewInput((object)right, merge);
                dynamic result;
                if (merge)
                    result = word.MergeDocuments(OriginalDocument: left, RevisedDocument: right, Destination: 2,
                        CompareFormatting: true, CompareTables: true, CompareHeaders: true, CompareFootnotes: true,
                        CompareTextboxes: true, CompareFields: true, CompareComments: true, FormatFrom: 0);
                else
                    result = word.CompareDocuments(OriginalDocument: left, RevisedDocument: right, Destination: 2,
                        CompareFormatting: true, CompareTables: true, CompareHeaders: true, CompareFootnotes: true,
                        CompareTextboxes: true, CompareFields: true, CompareComments: !collaboration, CompareMoves: !collaboration,
                        Granularity: collaboration ? 0 : 1, RevisedAuthor: author ?? "WordGit",
                        IgnoreAllComparisonWarnings: true);
                if (result == null) throw new InvalidOperationException("编辑器未生成独立的修订稿。");
                ActivateIfNeeded((object)result);
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

        private void PrepareReviewInput(object value, bool merge)
        {
            dynamic document = value;
            // Word needs active input windows on its isolated desktop. WPS works with hidden inputs.
            if (!IsWpsHost)
            {
                document.Windows[1].Visible = true;
                ActivateIfNeeded(value);
            }
            if (!merge)
            {
                document.TrackRevisions = false;
                document.AcceptAllRevisions();
            }
        }

        public object ReviewBranchFiles(MergePlan plan)
        {
            EnsureTrackedMergeSupported();
            if (plan.SubmittedBase != null)
                return ReviewRepeatedReturn(plan);
            if (plan.BasisIsTarget)
                return ReviewFiles(plan.Target, plan.Incoming, true);
            var folder = AppPaths.NewSession();
            var targetFile = Path.Combine(folder, "分支修订-当前.docx");
            var sourceFile = Path.Combine(folder, "分支修订-来源.docx");
            SaveBranchRevisionsToFile(plan.Basis, plan.Target, plan.TargetBranch ?? "当前分支", targetFile);
            SaveBranchRevisionsToFile(plan.Basis, plan.Incoming, plan.SourceBranch ?? "来源分支", sourceFile,
                plan.SubmittedBase != null || plan.PreserveIncoming);
            return ReviewStaged(folder, targetFile, sourceFile, true, null);
        }

        internal void BuildBranchMergeFile(MergePlan plan, string output)
        {
            SaveReviewToFile(ReviewBranchFiles(plan), output);
        }

        private object ReviewRepeatedReturn(MergePlan plan)
        {
            var basis = plan.Basis;
            var target = plan.Target;
            var incoming = plan.Incoming;
            var folder = AppPaths.NewSession();
            var currentFile = Path.Combine(folder, "当前稿.docx");
            var currentRevisions = Path.Combine(folder, "当前稿修订.docx");
            var incomingRevisions = Path.Combine(folder, "返稿修订.docx");
            var proposedFile = Path.Combine(folder, "返稿建议稿.docx");
            var rawReturn = Path.Combine(folder, "返稿副本.docx");
            var previousReturn = Path.Combine(folder, "上轮返稿副本.docx");
            var currentAuthor = "当前主稿 " + Guid.NewGuid().ToString("N").Substring(0, 8);
            object result = null;
            try
            {
                CopyShared(target, currentFile);
                List<DocumentComment> comments;
                bool preserveCurrentRevisions;
                var current = Open(currentFile, true, false);
                try
                {
                    comments = Comments(current, true);
                    preserveCurrentRevisions = ReviewItems(current).Any(item => !item.IsComment);
                }
                finally { CloseReviewInput(current); }

                if (plan.BasisIsTarget)
                {
                    if (preserveCurrentRevisions)
                        result = ReviewFiles(target, incoming, false, plan.IncomingAuthor);
                    else
                    {
                        CopyShared(incoming, incomingRevisions);
                        result = ReviewStaged(folder, currentFile, incomingRevisions, true, null);
                    }
                }
                else
                {
                    var previousReturns = plan.PreviousReturns;
                    var continued = FindContinuedReturn(previousReturns, incoming, rawReturn, previousReturn);
                    var continuedFile = continued == null ? basis : continued.File;
                    string currentBody = null;
                    if (continued != null)
                    {
                        current = Open(currentFile, true, false);
                        try { currentBody = (string)((dynamic)current).Content.Text; }
                        finally { CloseReviewInput(current); }
                    }
                    SaveBranchRevisionsToFile(continuedFile, target, currentAuthor, currentRevisions);
                    var earlier = new List<List<DocumentReviewItem>>();
                    if (continued == null)
                    {
                        foreach (var previous in previousReturns)
                        {
                            var review = ReviewFiles(basis, previous.File, true);
                            try { earlier.Add(ReviewItems(review, includeRevisionDates: true)); }
                            finally { CloseReviewInput(review); }
                        }
                    }
                    var proposal = continued == null ? ReviewFiles(basis, incoming, true) :
                        ReviewFiles(continuedFile, incoming, false, "返稿");
                    List<DocumentReviewItem> incomingItems;
                    try
                    {
                        ReturnReviewDecisions.RemoveRepeated(this, proposal, earlier);
                        incomingItems = ReviewItems(proposal);
                        ((dynamic)proposal).SaveAs2(FileName: incomingRevisions, FileFormat: 16, AddToRecentFiles: false);
                    }
                    finally { CloseReviewInput(proposal); }

                    // Build the suggested result with native Word combine, then compare it against
                    // the current manuscript when its committed content already contains revisions.
                    // Those revisions belong to the manuscript and must not be offered for rejection.
                    bool hasConflict;
                    var trackedCurrent = Open(currentRevisions, false, true);
                    try
                    {
                        hasConflict = ReturnReviewDecisions.PrepareConflictProposal(this, trackedCurrent, incomingItems, continued != null);
                        if (hasConflict) ((dynamic)trackedCurrent).Save();
                    }
                    finally { CloseReviewInput(trackedCurrent); }
                    result = ReviewStaged(folder, currentRevisions, incomingRevisions, true, null);
                    if (hasConflict || preserveCurrentRevisions)
                    {
                        ((dynamic)result).AcceptAllRevisions();
                        ((dynamic)result).SaveAs2(FileName: proposedFile, FileFormat: 16, AddToRecentFiles: false);
                        CloseReviewInput(result);
                        result = null;
                        var authors = incomingItems.Where(item => !item.IsComment).Select(item => item.Author).Distinct().ToArray();
                        result = ReviewStaged(folder, currentFile, proposedFile, false, authors.Length == 1 ? authors[0] : "返稿");
                    }
                    else ReturnReviewDecisions.KeepCurrentChanges(this, result, currentAuthor,
                        continued == null ? null : incomingItems, currentBody);
                }
                RestoreResolvedReviewComments(ref result, comments);
                var draft = result;
                result = null;
                return draft;
            }
            finally
            {
                if (result != null) CloseReviewInput(result);
                TryRemove(currentFile);
                TryRemove(currentRevisions);
                TryRemove(incomingRevisions);
                TryRemove(proposedFile);
                TryRemove(rawReturn);
                TryRemove(previousReturn);
                try { Directory.Delete(folder); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        private static ReturnFile FindContinuedReturn(List<ReturnFile> previous, string incoming, string rawReturn, string previousReturn)
        {
            if (previous.Count == 0) return null;
            CopyShared(incoming, rawReturn);
            var proposed = TrackedReviewSnapshot.Read(rawReturn);
            if (proposed.Count == 0) return null;
            foreach (var revision in previous)
            {
                CopyShared(revision.File, previousReturn);
                var archived = TrackedReviewSnapshot.Read(previousReturn);
                TryRemove(previousReturn);
                if (archived.Count > 0 && archived.All(change => proposed.ContainsKey(change.Key) && proposed[change.Key] >= change.Value))
                    return revision;
            }
            return null;
        }

        private void RestoreResolvedReviewComments(ref object result, List<DocumentComment> comments)
        {
            var unresolved = ReturnReviewDecisions.KeepResolvedComments(result, comments);
            if (unresolved.Count == 0) return;

            // Some Word combine results refuse Comment.Done even though the DOCX supports it.
            // Save before editing the generated copy and keep the native result as a fallback.
            string patched = null;
            try
            {
                var folder = AppPaths.NewSession();
                var original = Path.Combine(folder, "合并稿.docx");
                patched = Path.Combine(folder, "批注已恢复的合并稿.docx");
                ((dynamic)result).SaveAs2(FileName: original, FileFormat: 16, AddToRecentFiles: false);
                CopyShared(original, patched);
                var restored = ReviewCommentStatus.Restore(patched, unresolved);
                if (restored != unresolved.Count)
                    AppPaths.LogError("Resolved comment status could not be matched for " + (unresolved.Count - restored) + " thread(s).");
                if (restored == 0) return;
                var reopened = Open(patched, false);
                CloseReviewInput(result);
                result = reopened;
                patched = null;
                TryRemove(original);
            }
            catch (Exception error) { AppPaths.LogError("Restore resolved comment status: " + error); }
            finally { if (patched != null) TryRemove(patched); }
        }

        private void SaveBranchRevisionsToFile(string basisFile, string versionFile, string author, string outputPath,
            bool preserveRevisions = false)
        {
            SaveReviewToFile(ReviewFiles(basisFile, versionFile, preserveRevisions, author), outputPath);
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

        public string Restore(object source, string plan)
        {
            EnsureRestorable(source);
            return ApplyReplacement(SaveDocument(source), plan, true);
        }

        public static void EnsureRestorable(object source)
        {
            var path = LocalPath(source);
            EnsureReplaceable(path);
            if ((bool)((dynamic)source).ReadOnly) throw new InvalidOperationException("当前文档只读，请使用「另存旧版」。");
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

        // Closes the working copy, lets the kernel swap in the file it prepared and publish the
        // matching ref change, then reopens the document. The kernel journal makes the pair
        // recoverable if Word dies in between.
        public string ApplyReplacement(string path, string plan, bool keepActiveWindow = false)
        {
            EnsureReplaceable(path);
            dynamic doc = null;
            foreach (var candidate in word.Documents)
                if (String.Equals((string)((dynamic)candidate).FullName, path, StringComparison.OrdinalIgnoreCase)) { doc = candidate; break; }
            if (doc != null && ((bool)doc.ReadOnly || !(bool)doc.Saved))
                throw new InvalidOperationException("当前稿有未保存的修改，或处于只读状态。请先处理后重试。");
            var closed = false;
            var restorePainting = false;
            var screenUpdating = true;
            var reopenInBackground = false;
            string message;
            try
            {
                if (doc != null && !IsWpsHost)
                {
                    reopenInBackground = keepActiveWindow && word.Windows.Count > 0 &&
                        !ReferenceEquals((object)word.ActiveDocument, (object)doc);
                    if (!reopenInBackground)
                    {
                        screenUpdating = (bool)word.ScreenUpdating;
                        word.ScreenUpdating = false;
                        restorePainting = true;
                    }
                }
                if (doc != null) { doc.Close(SaveChanges: 0); closed = true; }
                message = ServiceClient.Connect().ApplyPlan(plan);
            }
            catch
            {
                // A pending plan keeps the document lock; release it so the next attempt can start.
                ServiceClient.Connect().CancelPlan(plan);
                throw;
            }
            finally
            {
                try
                {
                    if (closed)
                    {
                        var reopened = Open(path, false, !reopenInBackground);
                        if (reopenInBackground) ((dynamic)reopened).Windows[1].Visible = true;
                    }
                }
                finally { if (restorePainting) word.ScreenUpdating = screenUpdating; }
            }
            return message;
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
