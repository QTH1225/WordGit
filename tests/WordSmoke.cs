using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace WordGit
{
    internal static class WordSmoke
    {
        [STAThread]
        private static void Main()
        {
            Console.OutputEncoding = Encoding.UTF8;
            var root = AppPaths.NewSession();
            dynamic application = null;
            bool ownsSession = false;
            try
            {
                var existing = Process.GetProcessesByName("WINWORD").Select(process => (uint)process.Id).ToArray();
                application = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application", true));
                if (existing.Length != 0)
                {
                    var added = Process.GetProcessesByName("WINWORD").Where(process => !existing.Contains((uint)process.Id)).ToList();
                    if (added.Count != 1 || (int)application.Documents.Count != 0)
                        throw new InvalidOperationException("无法确认 Word 测试使用独立进程。现有文档未动。");
                }
                ownsSession = true;
                application.Visible = false;
                application.DisplayAlerts = 0;
                var word = new WordDocuments(application);
                VerifyRevisionActions(application, word, root);
                var sourcePath = Path.Combine(root, "原始 试验报告.docx");
                dynamic source = application.Documents.Add();
                source.Content.Text = "试验报告\r本节记录配合比。\r";
                source.Sections[1].Headers[1].Range.Text = "保留的分节页眉";
                source.PageSetup.TopMargin = 55f;
                dynamic anchor = source.Range(source.Content.End - 1, source.Content.End - 1);
                dynamic table = source.Tables.Add(anchor, 2, 2);
                table.Cell(1, 1).Range.Text = "龄期";
                table.Cell(1, 2).Range.Text = "试件";
                source.SaveAs2(FileName: sourcePath, FileFormat: 16, AddToRecentFiles: false);
                var original = WordDocuments.Snapshot((object)source);
                var repo = new GitHistory(sourcePath, root);
                var first = repo.Commit(original, "初稿", "测试", null);
                source.TrackRevisions = true;
                source.Range(0, 0).InsertBefore("当前还有未保存修改\r");
                source.Comments.Add(source.Range(0, 2), "定位这条批注");
                var comments = WordDocuments.Comments((object)source);
                Require(comments.Count == 1 && comments[0].Page > 0 && comments[0].Text == "定位这条批注", "Comment listing");
                word.GoToComment((object)source, comments[0]);
                Require((int)application.Selection.Start == comments[0].ScopeStart, "Comment navigation");
                source.TrackFormatting = true;
                dynamic deletion = source.Content.Duplicate;
                Require((bool)deletion.Find.Execute(FindText: "本节记录配合比"), "Deletion fixture text");
                deletion.Delete();
                var formatStart = "当前还有未保存修改\r".Length;
                source.Range(formatStart, formatStart + 2).Font.Italic = -1;
                source.Sections[1].Headers[1].Range.InsertBefore("页眉修订 ");
                var reviewItems = WordDocuments.ReviewItems((object)source);
                Require(reviewItems.Count(item => item.IsComment) == 1 &&
                    reviewItems.Any(item => item.RevisionType == 1 && item.StoryType == 1) &&
                    reviewItems.Any(item => item.RevisionType == 2) &&
                    reviewItems.Any(item => item.Kind == "格式") &&
                    reviewItems.Any(item => !item.IsComment && item.StoryType != 1), "Review listing missed a tracked change");
                foreach (var item in reviewItems.Where(item => !item.IsComment))
                {
                    word.GoToReviewItem((object)source, item);
                    Require((int)application.Selection.Start == item.ScopeStart &&
                        (int)application.Selection.Range.StoryType == item.StoryType, "Revision navigation selected the wrong range");
                }
                word.GoToReviewItem((object)source, reviewItems.Single(item => item.IsComment));
                var existingRevisions = (int)source.Revisions.Count;
                dynamic liveComparison = word.CompareCurrent(original, (object)source);
                Require((int)liveComparison.Revisions.Count > 0 && (int)liveComparison.Tables.Count == 1 &&
                    ((string)liveComparison.Content.Text).Contains("当前还有未保存修改"), "Live Word comparison missed the unsaved edits");
                Require(((string)liveComparison.Sections[1].Headers[1].Range.Text).Contains("保留的分节页眉") &&
                    Math.Abs((float)liveComparison.PageSetup.TopMargin - 55f) < 0.1f, "Live comparison lost header or margins");
                liveComparison.Close(SaveChanges: 0);
                source.Activate();
                liveComparison = word.CompareCurrent(original, (object)source);
                Require((int)liveComparison.Revisions.Count > 0, "Repeated comparison with revisions failed");
                liveComparison.Close(SaveChanges: 0);
                Require(WordDocuments.ReviewItems((object)source).Count == reviewItems.Count, "Comparison changed the review list");
                Require((int)source.Revisions.Count == existingRevisions && (int)source.Comments.Count == 1, "Comparison changed source revisions or comments");
                Require(!(bool)source.Saved && WordDocuments.ReadDocx(sourcePath).SequenceEqual(original),
                    "Comparison saved or modified the source document");
                source.AcceptAllRevisions();
                Require(WordDocuments.ReviewItems((object)source).All(item => item.IsComment),
                    "Accepted revisions remained in the review list");
                var staleRejected = false;
                try { word.GoToReviewItem((object)source, reviewItems.First(item => !item.IsComment)); }
                catch (InvalidOperationException) { staleRejected = true; }
                Require(staleRejected, "A stale revision navigated to another change");
                source.Close(SaveChanges: 0);
                source = word.Open(sourcePath, false);
                var returnPath = Path.Combine(root, "导师返稿.docx");
                WordDocuments.WriteNew(returnPath, original);
                dynamic returned = word.Open(returnPath, false, false);
                returned.TrackRevisions = true;
                returned.Range(0, 0).InsertBefore("审阅补充\r");
                var revised = WordDocuments.Snapshot((object)returned);
                returned.Revisions.AcceptAll();
                var cleanRevision = WordDocuments.Snapshot((object)returned);
                returned.Close(SaveChanges: 0);
                dynamic comparison = word.Review(repo.Extract(first), cleanRevision, false);
                Require(comparison.Revisions.Count > 0 && comparison.Tables.Count == 1, "Compare results");
                comparison.Close(SaveChanges: 0);
                dynamic merge = word.Review(repo.Extract(first), revised, true);
                Require(merge.Revisions.Count > 0 && merge.Tables.Count == 1, "Merge results");
                Require(((string)merge.Sections[1].Headers[1].Range.Text).Contains("保留的分节页眉"), "Header lost");
                merge.Revisions.AcceptAll();
                Require(((string)merge.Content.Text).Contains("审阅补充"), "Accept changes");
                var mergePath = Path.Combine(root, "合并结果.docx");
                merge.SaveAs2(FileName: mergePath, FileFormat: 16, AddToRecentFiles: false);
                var mergeBytes = WordDocuments.Snapshot((object)merge);
                var merged = repo.Commit(mergeBytes, "提交合并稿", "测试", first);
                Require(repo.Extract(merged).SequenceEqual(mergeBytes) && WordDocuments.ReadDocx(sourcePath).SequenceEqual(original), "Merge must not overwrite original");
                source.Range(0, 0).InsertBefore("待恢复文本\r");
                word.Restore(source, repo, first, "测试");
                Require(WordDocuments.ReadDocx(sourcePath).SequenceEqual(original) && repo.History().Count == 4, "Restore must preserve source bytes and history");
                var targetTip = repo.Head();
                repo.CreateBranch("导师意见", first);
                word.ReplaceWorkingCopy(sourcePath, original, original,
                    () => repo.Switch("main", targetTip, "导师意见", first));
                dynamic active = word.Current();
                active.Range(0, 0).InsertBefore("分支补充\r");
                var branchBytes = WordDocuments.Snapshot((object)active);
                var side = repo.Commit(branchBytes, "导师分支", "测试", first, "导师意见");
                word.ReplaceWorkingCopy(sourcePath, branchBytes, original,
                    () => repo.Switch("导师意见", side, "main", targetTip));
                Require(WordDocuments.ReadDocx(sourcePath).SequenceEqual(original), "Branch checkout changed DOCX bytes");
                active = word.Current();
                var last = (int)active.Content.End - 1;
                active.Range(last, last).InsertAfter("\r主分支备注");
                var targetBytes = WordDocuments.Snapshot((object)active);
                var target = repo.Commit(targetBytes, "主分支修改", "测试", targetTip, "main");
                var common = repo.MergeBase(target, side);
                Console.WriteLine("Running native review for diverged branches…");
                dynamic resolved = word.ReviewBranches(repo.Extract(common), targetBytes, branchBytes);
                Require(resolved.Revisions.Count > 0 && resolved.Tables.Count == 1, "Branch merge missing revisions or table");
                resolved.Revisions.AcceptAll();
                var body = (string)resolved.Content.Text;
                Require(body.Contains("主分支备注") && body.Contains("分支补充"), "Both branch edits must survive acceptance");
                Require(((string)resolved.Sections[1].Headers[1].Range.Text).Contains("保留的分节页眉"), "Branch merge lost headers");
                Require(Math.Abs((float)resolved.PageSetup.TopMargin - 55f) < 0.1f, "Branch merge changed the original margin");
                var verified = Path.Combine(root, "已审阅合并稿.docx");
                resolved.SaveAs2(FileName: verified, FileFormat: 16, AddToRecentFiles: false);
                var verifiedBytes = WordDocuments.Snapshot((object)resolved);
                string mergeCommit = null;
                word.ReplaceWorkingCopy(sourcePath, targetBytes, verifiedBytes,
                    () => mergeCommit = repo.Commit(verifiedBytes, "合并两个分支", "测试", target, "main", "导师意见", side));
                Require(repo.Head() == mergeCommit && repo.History().First().Parents.SequenceEqual(new[] { target, side }) &&
                    WordDocuments.ReadDocx(sourcePath).SequenceEqual(verifiedBytes), "Confirmed branch merge was not applied");
                Console.WriteLine("PASS: native Word compare, review merge, exact restoration and branch checkout, and divergent branch merge retaining both changes, table and header.");
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error.GetType().Name + ": " + error.Message);
                Console.Error.WriteLine(error.StackTrace);
                Environment.ExitCode = 1;
            }
            finally
            {
                if (ownsSession) application.Quit(SaveChanges: 0);
                try { Directory.Delete(root, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        private static void VerifyRevisionActions(object value, WordDocuments word, string root)
        {
            dynamic application = value;
            dynamic document = application.Documents.Add();
            var path = Path.Combine(root, "逐条审阅测试.docx");
            document.Content.Text = "保留正文\r删除后接受\r删除后拒绝\r格式测试\r";
            document.Sections[1].Headers[1].Range.Text = "页眉原文";
            document.SaveAs2(FileName: path, FileFormat: 16, AddToRecentFiles: false);
            var original = WordDocuments.ReadDocx(path);
            document.TrackRevisions = true;
            document.TrackFormatting = true;
            document.Range(0, 0).InsertBefore("新增后接受\r");
            var end = (int)document.Content.End - 1;
            document.Range(end, end).InsertBefore("新增后拒绝\r");
            foreach (var text in new[] { "删除后接受", "删除后拒绝" })
            {
                dynamic range = document.Content.Duplicate;
                Require((bool)range.Find.Execute(FindText: text), "Revision fixture deletion");
                range.Delete();
            }
            dynamic format = document.Content.Duplicate;
            Require((bool)format.Find.Execute(FindText: "格式测试"), "Revision fixture formatting");
            format.Font.Italic = -1;
            document.Sections[1].Headers[1].Range.InsertBefore("页眉新增");
            document.Comments.Add(document.Range(0, 2), "保留批注");
            var first = WordDocuments.ReviewItems((object)document).Single(item => item.RevisionType == 1 && item.RangeText.Contains("新增后接受"));
            word.ResolveReviewItem((object)document, first, true);
            Require(((string)document.Content.Text).Contains("新增后接受") &&
                !WordDocuments.ReviewItems((object)document).Any(item => item.RevisionType == 1 && item.RangeText.Contains("新增后接受")), "Accept insertion");
            Require((bool)document.Undo() && WordDocuments.ReviewItems((object)document).Any(item => item.RevisionType == 1 && item.RangeText.Contains("新增后接受")),
                "Native revision resolution cannot be undone");
            first = WordDocuments.ReviewItems((object)document).Single(item => item.RevisionType == 1 && item.RangeText.Contains("新增后接受"));
            word.ResolveReviewItem((object)document, first, true);
            var beforeStale = WordDocuments.ReviewItems((object)document).Count;
            var staleRejected = false;
            try { word.ResolveReviewItem((object)document, first, false); }
            catch (InvalidOperationException) { staleRejected = true; }
            Require(staleRejected && WordDocuments.ReviewItems((object)document).Count == beforeStale, "A stale resolution changed another revision");
            var itemToResolve = WordDocuments.ReviewItems((object)document).Single(item => item.RevisionType == 1 && item.RangeText.Contains("新增后拒绝"));
            word.ResolveReviewItem((object)document, itemToResolve, false);
            Require(!((string)document.Content.Text).Contains("新增后拒绝"), "Reject insertion");
            itemToResolve = WordDocuments.ReviewItems((object)document).Single(item => item.RevisionType == 2 && item.RangeText.Contains("删除后接受"));
            word.ResolveReviewItem((object)document, itemToResolve, true);
            Require(!((string)document.Content.Text).Contains("删除后接受"), "Accept deletion");
            itemToResolve = WordDocuments.ReviewItems((object)document).Single(item => item.RevisionType == 2 && item.RangeText.Contains("删除后拒绝"));
            word.ResolveReviewItem((object)document, itemToResolve, false);
            Require(((string)document.Content.Text).Contains("删除后拒绝"), "Reject deletion");
            itemToResolve = WordDocuments.ReviewItems((object)document).Single(item => item.Kind == "格式");
            word.ResolveReviewItem((object)document, itemToResolve, false);
            format = document.Content.Duplicate;
            Require((bool)format.Find.Execute(FindText: "格式测试") && (int)format.Font.Italic == 0, "Reject formatting");
            var headerChanges = WordDocuments.ReviewItems((object)document).Where(item => !item.IsComment && item.StoryType != 1).ToList();
            Require(headerChanges.Count > 0, "Header revision fixture");
            Require(headerChanges.Select(item => item.Key).Distinct().Count() == headerChanges.Count &&
                headerChanges.All(item => item.StoryType != 4), "Duplicate or internal comment revisions");
            while (headerChanges.Count > 0)
            {
                itemToResolve = headerChanges[0];
                word.ResolveReviewItem((object)document, itemToResolve, false);
                var remainingHeaders = WordDocuments.ReviewItems((object)document).Where(item => !item.IsComment && item.StoryType != 1).ToList();
                Require(remainingHeaders.Count < headerChanges.Count, "Header revision resolution made no progress");
                headerChanges = remainingHeaders;
            }
            Require(!((string)document.Sections[1].Headers[1].Range.Text).Contains("页眉新增"), "Reject header revision");
            Require(WordDocuments.ReviewItems((object)document).All(item => item.IsComment) && (int)document.Comments.Count == 1 &&
                !(bool)document.Saved && WordDocuments.ReadDocx(path).SequenceEqual(original), "Review actions saved the file or changed comments");
            document.Close(SaveChanges: 0);
            document = word.Open(path, true, false);
            var readOnlyRejected = false;
            try { word.ResolveReviewItem((object)document, itemToResolve, true); }
            catch (InvalidOperationException) { readOnlyRejected = true; }
            Require(readOnlyRejected, "Read-only revision resolution");
            document.Close(SaveChanges: 0);
            Console.WriteLine("PASS: native accept/reject for insertion, deletion, formatting and header revisions, Undo, stale-item and read-only protection.");
        }

        private static void Require(bool condition, string detail)
        {
            if (!condition) throw new InvalidOperationException(detail);
        }
    }
}
