using System;
using System.IO;
using System.IO.Compression;
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
        public static string NewSession()
        {
            return Directory.CreateDirectory(Path.Combine(Root, "sessions", Guid.NewGuid().ToString("N"))).FullName;
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

            using (var sha = SHA256.Create())
            {
                if (headHash == null) headHash = readHead();
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    lastState = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant() == headHash
                        ? WorkingCopyState.Clean : WorkingCopyState.Changed;
            }
            lastWrite = writeTime;
            lastLength = length;
            measured = true;
            return lastState;
        }
    }

    internal sealed class WordDocuments
    {
        private readonly dynamic word;
        public WordDocuments(object application) { word = application; }
        public IWin32Window Owner
        {
            get { return word.Windows.Count == 0 ? null : new WordWindow(new IntPtr((int)word.ActiveWindow.Hwnd)); }
        }

        public object Current()
        {
            if (word.Documents.Count == 0) throw new InvalidOperationException("请先在 Word 打开一份 .docx。");
            return word.ActiveDocument;
        }

        public static System.Collections.Generic.List<DocumentComment> Comments(object value)
        {
            dynamic document = value;
            var result = new System.Collections.Generic.List<DocumentComment>();
            for (var index = 1; index <= (int)document.Comments.Count; index++)
            {
                dynamic comment = document.Comments[index];
                dynamic scope = comment.Scope;
                result.Add(new DocumentComment
                {
                    Index = index, Author = (string)comment.Author, Text = ((string)comment.Range.Text).Trim(),
                    ScopeStart = (int)scope.Start, Page = (int)scope.Information[3]
                });
            }
            return result;
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

        public static System.Collections.Generic.List<DocumentReviewItem> ReviewItems(object value)
        {
            dynamic document = value;
            var result = new System.Collections.Generic.List<DocumentReviewItem>();
            foreach (var comment in Comments(value)) result.Add(DocumentReviewItem.ForComment(comment));
            var seen = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
            foreach (dynamic first in document.StoryRanges)
            {
                dynamic story = first;
                var storyType = (int)story.StoryType;
                if (storyType == 4) continue; // Comment text is already represented by the comment entry.
                var storyIndex = 0;
                while (story != null)
                {
                    dynamic changes = story.Revisions;
                    var count = (int)changes.Count;
                    for (var index = 1; index <= count; index++)
                    {
                        dynamic change = changes[index];
                        dynamic scope = change.Range;
                        dynamic anchor = scope.Duplicate;
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
                    story = story.NextStoryRange;
                    storyIndex++;
                }
            }
            return result;
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
                throw new InvalidOperationException("当前只支持可直接读取的本地 .docx，请先在 Word 中将云端文件另存一份到本地。");
            if (!String.Equals(Path.GetExtension(filename), ".docx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("WordGit 第一版只支持 .docx。");
            return Path.GetFullPath(filename);
        }

        public static byte[] Snapshot(object value)
        {
            dynamic doc = value;
            var path = LocalPath(value, true);
            if (!(bool)doc.Saved)
            {
                if ((bool)doc.ReadOnly) throw new InvalidOperationException("文档是只读的，当前修改无法保存。请先另存为 .docx。");
                doc.Save();
            }
            if (!(bool)doc.Saved) throw new InvalidOperationException("Word 尚未保存成功，未创建提交。");
            return ReadDocx(path);
        }

        public static byte[] ReadDocx(string path)
        {
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var copy = new MemoryStream())
            {
                input.CopyTo(copy);
                using (var zip = new ZipArchive(copy, ZipArchiveMode.Read, true))
                    if (zip.GetEntry("[Content_Types].xml") == null || zip.GetEntry("word/document.xml") == null)
                        throw new InvalidDataException("文件不是可读取的 Word .docx。");
                return copy.ToArray();
            }
        }

        public static void WriteNew(string path, byte[] bytes)
        {
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                output.Write(bytes, 0, bytes.Length);
        }

        public object Open(string path, bool readOnly, bool visible = true)
        {
            return word.Documents.Open(FileName: path, ConfirmConversions: false, ReadOnly: readOnly,
                AddToRecentFiles: false, Visible: visible);
        }

        public object Review(byte[] original, byte[] revised, bool merge, string author = null)
        {
            return Review(original, revised, null, merge, author);
        }

        public object CompareCurrent(byte[] original, object current)
        {
            return Review(original, null, current, false, null);
        }

        private object Review(byte[] original, byte[] revised, object current, bool merge, string author)
        {
            var folder = AppPaths.NewSession();
            var leftFile = Path.Combine(folder, "基准版本.docx");
            var rightFile = Path.Combine(folder, current == null ? "修订版本.docx" : "当前稿.xml");
            dynamic left = null, right = null;
            try
            {
                WriteNew(leftFile, original);
                if (current == null) WriteNew(rightFile, revised);
                else File.WriteAllText(rightFile, (string)((dynamic)current).WordOpenXML, new System.Text.UTF8Encoding(false));
                left = Open(leftFile, false, false);
                right = Open(rightFile, false, false);
                if (!merge)
                {
                    left.TrackRevisions = false;
                    right.TrackRevisions = false;
                    left.AcceptAllRevisions();
                    right.AcceptAllRevisions();
                }
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
                TryRemove(leftFile);
                TryRemove(rightFile);
                try { Directory.Delete(folder); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        public object ReviewBranches(byte[] basis, byte[] target, byte[] incoming, string targetBranch = "当前分支", string sourceBranch = "来源分支")
        {
            if (basis.SequenceEqual(target)) return Review(target, incoming, true);
            var targetChanges = SaveBranchRevisions(basis, target, targetBranch);
            var sourceChanges = SaveBranchRevisions(basis, incoming, sourceBranch);
            return Review(targetChanges, sourceChanges, true);
        }

        private byte[] SaveBranchRevisions(byte[] basis, byte[] version, string author)
        {
            var folder = AppPaths.NewSession();
            var path = Path.Combine(folder, "分支修订.docx");
            object document = null;
            try
            {
                document = Review(basis, version, false, author);
                ((dynamic)document).SaveAs2(FileName: path, FileFormat: 16, AddToRecentFiles: false);
                return ReadDocx(path);
            }
            finally
            {
                CloseReviewInput(document);
                TryRemove(path);
                try { Directory.Delete(folder); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        public string Restore(object source, GitHistory repo, string revision, string author)
        {
            var path = LocalPath(source);
            EnsureReplaceable(path);
            if ((bool)((dynamic)source).ReadOnly) throw new InvalidOperationException("当前文档只读，请使用「另存旧版」。");
            var oldBytes = repo.Extract(revision);
            var current = Snapshot(source);
            if (current.SequenceEqual(oldBytes)) return "当前稿已经与所选旧版一致。";
            var branch = repo.ActiveBranch();
            var head = repo.Head();
            head = repo.Commit(current, "恢复前保存当前稿", author, head, branch) ?? head;
            string saved = null;
            ReplaceWorkingCopy(path, current, oldBytes,
                () => saved = repo.Commit(oldBytes, "恢复版本 " + revision.Substring(0, 8), author, head, branch));
            return saved == null ? "已恢复旧版。" : "已恢复并创建新提交 " + saved.Substring(0, 8) + "。";
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

        public void ReplaceWorkingCopy(string path, byte[] current, byte[] replacement, Action publish)
        {
            EnsureReplaceable(path);
            dynamic doc = null;
            foreach (var candidate in word.Documents)
                if (String.Equals((string)((dynamic)candidate).FullName, path, StringComparison.OrdinalIgnoreCase)) { doc = candidate; break; }
            if (doc != null && ((bool)doc.ReadOnly || !(bool)doc.Saved))
                throw new InvalidOperationException("当前稿有未保存的修改，或处于只读状态。请先处理后重试。");
            if (!ReadDocx(path).SequenceEqual(current))
                throw new InvalidOperationException("磁盘上的当前稿已变化，未替换文件。");
            if (current.SequenceEqual(replacement)) { publish(); return; }
            var folderPath = Path.GetDirectoryName(path);
            var staged = Path.Combine(folderPath, ".wordgit-restore-" + Guid.NewGuid().ToString("N") + ".tmp");
            var backup = staged + ".bak";
            WriteNew(staged, replacement);
            var closed = false;
            try
            {
                if (doc != null) { doc.Close(SaveChanges: 0); closed = true; }
                if (!ReadDocx(path).SequenceEqual(current))
                    throw new InvalidOperationException("磁盘文件在操作期间发生变化，未覆盖。请重新打开后重试。");
                File.Replace(staged, path, backup);
                try { publish(); }
                catch
                {
                    if (!ReadDocx(path).SequenceEqual(replacement))
                        throw new IOException("当前稿在操作期间被其他程序修改，原稿备份保留在：" + backup);
                    File.Replace(backup, path, null);
                    throw;
                }
                TryRemove(backup);
            }
            finally
            {
                TryRemove(staged);
                if (closed) Open(path, false);
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
