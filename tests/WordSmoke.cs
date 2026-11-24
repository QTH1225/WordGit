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
        private static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            var root = AppPaths.NewSession();
            dynamic application = null;
            bool ownsSession = false;
            var wps = args.Contains("--wps");
            var wordProcesses = wps ? Process.GetProcessesByName("WINWORD").Select(process => process.Id).ToArray() : new int[0];
            try
            {
                var processName = wps ? "wps" : "WINWORD";
                var executable = wps ? "wps.exe" : "WINWORD.EXE";
                var progId = wps ? "KWPS.Application" : "Word.Application";
                var existing = Process.GetProcessesByName(processName).Select(process => (uint)process.Id).ToArray();
                application = Activator.CreateInstance(Type.GetTypeFromProgID(progId, true));
                Require(File.Exists(Path.Combine((string)application.Path, executable)), "Automation opened a different editor from " + progId);
                if (existing.Length != 0)
                {
                    var added = Process.GetProcessesByName(processName).Where(process => !existing.Contains((uint)process.Id)).ToList();
                    if ((wps ? added.Count == 0 : added.Count != 1) || (int)application.Documents.Count != 0)
                        throw new InvalidOperationException("无法确认 Word 测试使用独立进程。现有文档未动。");
                }
                ownsSession = true;
                application.Visible = false;
                application.DisplayAlerts = 0;
                Console.WriteLine("Testing " + progId + " " + application.Version);
                if (args.Contains("--installed-addin")) VerifyInstalledAddIn(application);
                var word = new WordDocuments(application);
                if (wps)
                {
                    VerifyRevisionActions(application, word, root);
                    VerifyWpsCompatibility(application, word, root);
                    Require(Process.GetProcessesByName("WINWORD").All(process => wordProcesses.Contains(process.Id)),
                        "WPS started a Microsoft Word process");
                    Console.WriteLine("PASS: the WPS session did not start Word.");
                    return;
                }
                if (args.Contains("--dialogs-only")) { VerifyDialogOwners(application, word); return; }
                if (args.Contains("--pane-only")) { VerifyPaneLoading(application, root); return; }
                if (args.Contains("--remote-only")) { VerifyRemoteCollaboration(application, word, root); return; }
                if (args.Contains("--multi-review-only")) { VerifyRepeatedReviews(application, word, root); return; }
                if (args.Contains("--review-only"))
                {
                    VerifyImportedReview(application, word, root);
                    VerifyRepeatedReviews(application, word, root);
                    return;
                }
                if (!args.Contains("--local-only"))
                {
                    VerifyRevisionActions(application, word, root);
                    VerifyImportedReview(application, word, root);
                    VerifyRepeatedReviews(application, word, root);
                    VerifyRemoteCollaboration(application, word, root);
                }
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
                var original = TestFiles.Snapshot((object)source);
                var repo = new GitHistory(sourcePath, root);
                var first = repo.Commit(original, "初稿", "测试", null);
                var previewPath = Path.Combine(root, "初稿预览.docx");
                WordDocuments.ExtractToFile(repo, first, previewPath);
                dynamic preview = word.Open(previewPath, true);
                Require((bool)preview.ReadOnly && (int)preview.Tables.Count == 1 &&
                    TestFiles.ReadDocx(previewPath).SequenceEqual(original), "Word did not open the exact original in read-only mode");
                preview.Close(SaveChanges: 0);
                source.Activate();
                Console.WriteLine("PASS: Word commits and previews the byte-identical original read-only.");
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
                dynamic liveComparison = word.CompareCurrentExtracted(repo, first, (object)source);
                Require((int)liveComparison.Revisions.Count > 0 && (int)liveComparison.Tables.Count == 1 &&
                    ((string)liveComparison.Content.Text).Contains("当前还有未保存修改"), "Live Word comparison missed the unsaved edits");
                Require(((string)liveComparison.Sections[1].Headers[1].Range.Text).Contains("保留的分节页眉") &&
                    Math.Abs((float)liveComparison.PageSetup.TopMargin - 55f) < 0.1f, "Live comparison lost header or margins");
                liveComparison.Close(SaveChanges: 0);
                source.Activate();
                liveComparison = word.CompareCurrentExtracted(repo, first, (object)source);
                Require((int)liveComparison.Revisions.Count > 0, "Repeated comparison with revisions failed");
                liveComparison.Close(SaveChanges: 0);
                Require(WordDocuments.ReviewItems((object)source).Count == reviewItems.Count, "Comparison changed the review list");
                Require((int)source.Revisions.Count == existingRevisions && (int)source.Comments.Count == 1, "Comparison changed source revisions or comments");
                Require(!(bool)source.Saved && TestFiles.ReadDocx(sourcePath).SequenceEqual(original),
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
                // A saved draft is compared from its file instead of a WordOpenXML string.
                dynamic savedComparison = word.CompareCurrentExtracted(repo, first, (object)source);
                Require((int)savedComparison.Revisions.Count == 0 && (int)savedComparison.Tables.Count == 1 && (bool)source.Saved,
                    "Comparing the saved draft with its own commit reported changes");
                savedComparison.Close(SaveChanges: 0);
                source.Activate();
                var returnPath = Path.Combine(root, "导师返稿.docx");
                TestFiles.WriteNew(returnPath, original);
                dynamic returned = word.Open(returnPath, false, false);
                returned.TrackRevisions = true;
                returned.Range(0, 0).InsertBefore("审阅补充\r");
                var revised = TestFiles.Snapshot((object)returned);
                returned.Revisions.AcceptAll();
                var cleanRevision = TestFiles.Snapshot((object)returned);
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
                var mergeBytes = TestFiles.Snapshot((object)merge);
                var merged = repo.Commit(mergeBytes, "提交合并稿", "测试", first);
                Require(repo.Extract(merged).SequenceEqual(mergeBytes) && TestFiles.ReadDocx(sourcePath).SequenceEqual(original), "Merge must not overwrite original");
                source.Range(0, 0).InsertBefore("待恢复文本\r");
                word.Restore(source, repo, first, "测试");
                Require(TestFiles.ReadDocx(sourcePath).SequenceEqual(original) && repo.History().Count == 4, "Restore must preserve source bytes and history");
                var targetTip = repo.Head();
                repo.CreateBranch("导师意见", first);
                word.ReplaceWorkingCopy(repo, sourcePath, original, original,
                    () => repo.Switch("main", targetTip, "导师意见", first));
                dynamic active = word.Current();
                active.Range(0, 0).InsertBefore("分支补充\r");
                var branchBytes = TestFiles.Snapshot((object)active);
                var side = repo.Commit(branchBytes, "导师分支", "测试", first, "导师意见");
                word.ReplaceWorkingCopy(repo, sourcePath, branchBytes, original,
                    () => repo.Switch("导师意见", side, "main", targetTip));
                Require(TestFiles.ReadDocx(sourcePath).SequenceEqual(original) && !File.Exists(WorkingCopyJournal.PathFor(repo)),
                    "Branch checkout changed DOCX bytes or left its journal");
                active = word.Current();
                var last = (int)active.Content.End - 1;
                active.Range(last, last).InsertAfter("\r主分支备注");
                var targetBytes = TestFiles.Snapshot((object)active);
                var target = repo.Commit(targetBytes, "主分支修改", "测试", targetTip, "main");
                var common = repo.MergeBase(target, side);
                Console.WriteLine("Running native review for diverged branches…");
                dynamic resolved = word.ReviewBranchesExtracted(repo, common, target, side);
                Require(resolved.Revisions.Count > 0 && resolved.Tables.Count == 1, "Branch merge missing revisions or table");
                resolved.Revisions.AcceptAll();
                var body = (string)resolved.Content.Text;
                Require(body.Contains("主分支备注") && body.Contains("分支补充"), "Both branch edits must survive acceptance");
                Require(((string)resolved.Sections[1].Headers[1].Range.Text).Contains("保留的分节页眉"), "Branch merge lost headers");
                Require(Math.Abs((float)resolved.PageSetup.TopMargin - 55f) < 0.1f, "Branch merge changed the original margin");
                var verified = Path.Combine(root, "已审阅合并稿.docx");
                resolved.SaveAs2(FileName: verified, FileFormat: 16, AddToRecentFiles: false);
                var verifiedBytes = TestFiles.Snapshot((object)resolved);
                string mergeCommit = null;
                word.ReplaceWorkingCopy(repo, sourcePath, targetBytes, verifiedBytes,
                    () => mergeCommit = repo.Commit(verifiedBytes, "合并两个分支", "测试", target, "main", "导师意见", side));
                Require(repo.Head() == mergeCommit && repo.History().First().Parents.SequenceEqual(new[] { target, side }) &&
                    TestFiles.ReadDocx(sourcePath).SequenceEqual(verifiedBytes), "Confirmed branch merge was not applied");
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
                try {
                    var parent = Path.GetFullPath(Path.Combine(AppPaths.Root, "sessions")) + Path.DirectorySeparatorChar;
                    if (!Path.GetFullPath(root).StartsWith(parent, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Unexpected test session path");
                    foreach (var file in Directory.GetFiles(root, "*", SearchOption.AllDirectories))
                        File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
                    Directory.Delete(root, true);
                } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        private static void VerifyInstalledAddIn(dynamic application)
        {
            dynamic document = application.Documents.Add();
            try
            {
                var addIn = application.COMAddIns.Item("WordGit.Connect");
                Require((bool)addIn.Connect, "WordGit is registered but not connected in this editor");
                Console.WriteLine("PASS: the installed WordGit COM add-in loads automatically.");
            }
            finally { document.Close(SaveChanges: 0); }
        }

        private static void VerifyWpsCompatibility(dynamic application, WordDocuments word, string root)
        {
            var path = Path.Combine(root, "WPS 测试稿.docx");
            dynamic draft = application.Documents.Add();
            draft.Content.Text = "送审版\r正文段落\r";
            draft.Sections[1].Headers[1].Range.Text = "测试页眉";
            draft.PageSetup.TopMargin = 55f;
            dynamic table = draft.Tables.Add(draft.Range(draft.Content.End - 1, draft.Content.End - 1), 2, 2);
            table.Cell(1, 1).Range.Text = "试件";
            draft.SaveAs2(FileName: path, FileFormat: 16, AddToRecentFiles: false);
            var repo = new GitHistory(path, root);
            var first = repo.CommitFile(WordDocuments.SaveDocument((object)draft), "送审版", "测试", null);
            var original = TestFiles.ReadDocx(path);
            draft.Range(0, 0).InsertBefore("我方修改\r");
            var second = repo.CommitFile(WordDocuments.SaveDocument((object)draft), "修改稿", "测试", first);
            Require(first != null && second != null && repo.History().Count == 2, "WPS failed to commit both versions");
            Console.WriteLine("PASS: WPS commits two native DOCX versions.");

            var previewPath = Path.Combine(root, "送审版预览.docx");
            WordDocuments.ExtractToFile(repo, first, previewPath);
            dynamic preview = word.Open(previewPath, true);
            Require((bool)preview.ReadOnly && (int)preview.Tables.Count == 1 &&
                ((string)preview.Sections[1].Headers[1].Range.Text).Contains("测试页眉") &&
                Math.Abs((float)preview.PageSetup.TopMargin - 55f) < 0.1f &&
                TestFiles.ReadDocx(previewPath).SequenceEqual(original), "WPS preview failed to preserve the old DOCX");
            preview.Close(SaveChanges: 0);
            Console.WriteLine("PASS: WPS opens the exact old version read-only and preserves table, header and margin.");

            var latest = TestFiles.ReadDocx(path);
            draft.Activate();
            draft.TrackRevisions = true;
            draft.Range(0, 0).InsertBefore("未保存的修改\r");
            var sourceRevisions = (int)draft.Revisions.Count;
            for (int attempt = 0; attempt < 2; attempt++)
            {
                dynamic compared = word.CompareCurrentExtracted(repo, second, (object)draft);
                Require((int)compared.Revisions.Count > 0 &&
                    ((string)compared.Content.Text).Contains("未保存的修改") &&
                    (int)compared.Tables.Count == 1, "WPS comparison missed tracked edits or table");
                compared.Close(SaveChanges: 0);
                draft.Activate();
            }
            Require(!(bool)draft.Saved && (int)draft.Revisions.Count == sourceRevisions &&
                TestFiles.ReadDocx(path).SequenceEqual(latest), "WPS comparison changed the working copy");
            Console.WriteLine("PASS: WPS compares unsaved tracked edits repeatedly without saving the draft.");

            draft.Close(SaveChanges: 0);
            draft = word.Open(path, false);
            word.Restore((object)draft, repo, first, "测试");
            Require(TestFiles.ReadDocx(path).SequenceEqual(original) && repo.History().Count == 3 &&
                repo.Extract(second).SequenceEqual(latest), "WPS restore changed historical bytes or refs");
            Console.WriteLine("PASS: WPS restores identical bytes in a new commit.");

            var returnPath = Path.Combine(root, "老师返稿.docx");
            TestFiles.WriteNew(returnPath, original);
            dynamic returned = word.Open(returnPath, false);
            returned.Comments.Add(returned.Range(0, 2), "这里需要补充数据");
            returned.TrackRevisions = true;
            returned.Range(0, 0).InsertBefore("老师补充\r");
            WordDocuments.SaveDocument((object)returned);
            returned.Close(SaveChanges: 0);
            var tip = repo.Head();
            var review = repo.ImportReviewFile(returnPath, first, "review-wps", "老师", "main", tip);
            Require(repo.BranchTip("review-wps") == review && repo.History("review-wps").First().Parents.Single() == first &&
                repo.Extract(review).SequenceEqual(TestFiles.ReadDocx(returnPath)), "WPS archive altered the returned draft or its parent");
            var archivePath = Path.Combine(root, "返稿预览.docx");
            WordDocuments.ExtractToFile(repo, review, archivePath);
            dynamic archive = word.Open(archivePath, true);
            Require((int)archive.Comments.Count == 1 && (int)archive.Revisions.Count > 0,
                "WPS did not reopen the archived comments and tracked revisions");
            archive.Close(SaveChanges: 0);

            var blocked = false;
            try { word.ReviewBranchesExtracted(repo, first, tip, review); }
            catch (InvalidOperationException error) { blocked = error.Message.Contains("WPS"); }
            Require(blocked && repo.Head() == tip && repo.BranchTip("review-wps") == review &&
                TestFiles.ReadDocx(path).SequenceEqual(original), "WPS merge must leave the original and archived history intact");
            Console.WriteLine("PASS: WPS archives review bytes, comments and revisions; automatic tracked merge is unavailable.");
        }

        private static void VerifyRepeatedReviews(dynamic application, WordDocuments word, string root)
        {
            const string accepted = "第一轮接受的新增句";
            const string rejected = "第一轮拒绝的新增句";
            const string oldComment = "第一轮已解决的批注";
            const string newComment = "第二轮新增批注";
            const string newText = "第二轮新增说明";
            const string acceptedFormat = "保留第一轮斜体";
            const string rejectedFormat = "拒绝第一轮加粗";
            const string currentConflict = "作者改写一句";
            const string returnedConflict = "老师另改一句";
            var basisPath = Path.Combine(root, "多轮测试送审版.docx");
            dynamic document = application.Documents.Add();
            document.RemovePersonalInformation = false;
            document.Content.Text = "多轮测试标题\r接受锚点\r间隔段落\r拒绝锚点\r批注定位位置\r" +
                acceptedFormat + "\r" + rejectedFormat + "\r第二轮锚点\r第三轮冲突锚点\r";
            document.SaveAs2(FileName: basisPath, FileFormat: 16, AddToRecentFiles: false);
            document.Close(SaveChanges: 0);
            var history = new GitHistory(basisPath, root);
            var basis = history.CommitFile(basisPath, "原始送审版", "测试", null);

            var firstPath = Path.Combine(root, "第一轮返稿.docx");
            WordDocuments.ExtractToFile(history, basis, firstPath);
            document = word.Open(firstPath, false);
            document.TrackRevisions = true;
            InsertAtReviewAnchor(document, "接受锚点", accepted);
            InsertAtReviewAnchor(document, "拒绝锚点", rejected);
            document.TrackFormatting = true;
            FindReviewAnchor(document, acceptedFormat).Font.Italic = -1;
            FindReviewAnchor(document, rejectedFormat).Font.Bold = -1;
            document.Comments.Add(FindReviewAnchor(document, "批注定位位置"), oldComment);
            document.Save();
            document.Close(SaveChanges: 0);
            var first = history.ImportReviewFile(firstPath, basis, null, "老师", "main", basis);
            var firstRef = GitHistory.ReviewArchivePrefix + first;
            dynamic firstMerge = word.ReviewBranchesExtracted(history, basis, basis, first);
            var reject = WordDocuments.ReviewItems((object)firstMerge).Single(item => !item.IsComment && item.RangeText.Contains(rejected));
            word.ResolveReviewItem((object)firstMerge, reject, false);
            var rejectFormat = WordDocuments.ReviewItems((object)firstMerge).Single(item =>
                !item.IsComment && item.Kind == "格式" && item.RangeText.Contains(rejectedFormat));
            word.ResolveReviewItem((object)firstMerge, rejectFormat, false);
            firstMerge.AcceptAllRevisions();
            firstMerge.Comments[1].Done = true;
            var firstResult = Path.Combine(root, "第一轮确认后.docx");
            firstMerge.SaveAs2(FileName: firstResult, FileFormat: 16, AddToRecentFiles: false);
            firstMerge.Close(SaveChanges: 0);
            var target = history.CommitFile(firstResult, "第一轮完成", "测试", basis, "main", firstRef, first);
            Require(history.History("main").First().Parents.SequenceEqual(new[] { basis, first }) &&
                history.Branches().SequenceEqual(new[] { "main" }), "Unnamed original review was not retained as a parent");

            var secondPath = Path.Combine(root, "第二轮沿用第一轮原始返稿.docx");
            File.Copy(firstPath, secondPath);
            document = word.Open(secondPath, false);
            document.TrackRevisions = true;
            Require(!WordDocuments.Comments((object)document, true).Single(item => item.Text == oldComment).Resolved,
                "The raw return must still have the old comment open");
            dynamic adjacent = FindReviewAnchor(document, "间隔段落");
            var end = (int)adjacent.End;
            document.Range(end, end).InsertAfter("\r" + newText);
            document.Comments.Add(FindReviewAnchor(document, "多轮测试标题"), newComment);
            document.Save();
            document.Close(SaveChanges: 0);
            var second = history.ImportReviewFile(secondPath, basis, "review-round-two", "老师", "main", target);
            dynamic secondMerge = word.ReviewBranchesExtracted(history, basis, target, second);
            var secondItems = WordDocuments.ReviewItems((object)secondMerge, false);
            Require(secondItems.Count == 2 && secondItems.Any(item => item.IsComment && item.Text == newComment) &&
                secondItems.Any(item => !item.IsComment && item.RangeText.Contains(newText)),
                "The second round must only show the new revision and unresolved comment");
            Require(WordDocuments.Comments((object)secondMerge, true).Single(item => item.Text == oldComment).Resolved,
                "The first resolved comment was reopened");
            secondMerge.AcceptAllRevisions();
            secondMerge.TrackRevisions = false;
            dynamic conflictAnchor = FindReviewAnchor(secondMerge, "第三轮冲突锚点");
            conflictAnchor.Text = currentConflict;
            var secondResult = Path.Combine(root, "第二轮确认后.docx");
            secondMerge.SaveAs2(FileName: secondResult, FileFormat: 16, AddToRecentFiles: false);
            var body = (string)secondMerge.Content.Text;
            Require(body.Contains(accepted) && body.Contains(newText) && !body.Contains(rejected),
                "An accepted or rejected decision from the first round was lost");
            Require((int)FindReviewAnchor(secondMerge, acceptedFormat).Font.Italic == -1 &&
                (int)FindReviewAnchor(secondMerge, rejectedFormat).Font.Bold == 0,
                "An accepted or rejected formatting decision from the first round was lost");
            secondMerge.Close(SaveChanges: 0);
            var nextTip = history.CommitFile(secondResult, "第二轮完成", "测试", target, "main", "review-round-two", second);

            var thirdPath = Path.Combine(root, "第三轮修改同一句.docx");
            WordDocuments.ExtractToFile(history, basis, thirdPath);
            document = word.Open(thirdPath, false);
            document.TrackRevisions = true;
            dynamic thirdAnchor = FindReviewAnchor(document, "第三轮冲突锚点");
            thirdAnchor.Text = returnedConflict;
            document.Save();
            document.Close(SaveChanges: 0);
            var third = history.ImportReviewFile(thirdPath, basis, null, "老师", "main", nextTip);
            dynamic thirdMerge = word.ReviewBranchesExtracted(history, basis, nextTip, third);
            var thirdItems = WordDocuments.ReviewItems((object)thirdMerge, false);
            Require(thirdItems.Any(item => !item.IsComment && item.RangeText.Contains(returnedConflict)) &&
                thirdItems.Any(item => !item.IsComment && item.RangeText.Contains(currentConflict)),
                "Conflicting edits at the same base range must remain reviewable");
            Require(thirdItems.All(item => item.IsComment || (!item.RangeText.Contains(accepted) && !item.RangeText.Contains(rejected) &&
                !item.RangeText.Contains(newText))), "A previous non-conflicting decision returned to the review list");
            var thirdDraftPath = Path.Combine(root, "第三轮待确认稿.docx");
            thirdMerge.SaveAs2(FileName: thirdDraftPath, FileFormat: 16, AddToRecentFiles: false);
            thirdMerge.RejectAllRevisions();
            body = (string)thirdMerge.Content.Text;
            Require(body.Contains(currentConflict) && !body.Contains(returnedConflict) && !body.Contains("第三轮冲突锚点"),
                "Rejecting the third-round conflict did not retain the actual current text");
            thirdMerge.Close(SaveChanges: 0);
            thirdMerge = word.Open(thirdDraftPath, false);
            thirdMerge.AcceptAllRevisions();
            body = (string)thirdMerge.Content.Text;
            Require(body.Contains(returnedConflict) && !body.Contains(currentConflict) && !body.Contains("第三轮冲突锚点") &&
                body.Contains(accepted) && body.Contains(newText) && !body.Contains(rejected),
                "Accepting the third-round conflict duplicated text or changed earlier review decisions");
            thirdMerge.Close(SaveChanges: 0);
            Require(history.Head() == nextTip &&
                history.Extract(first).SequenceEqual(TestFiles.ReadDocx(firstPath)) &&
                history.Extract(second).SequenceEqual(TestFiles.ReadDocx(secondPath)) &&
                history.Extract(third).SequenceEqual(TestFiles.ReadDocx(thirdPath)),
                "Multi-round review modified an original archive or advanced the main branch");
            Console.WriteLine("PASS: three review rounds using the original baseline, accepted and rejected edits, resolved comments, a conflict, and byte-identical archives.");
        }

        private static dynamic FindReviewAnchor(dynamic document, string anchor)
        {
            dynamic range = document.Content.Duplicate;
            Require((bool)range.Find.Execute(FindText: anchor), "Missing review fixture anchor " + anchor);
            return range;
        }

        private static void InsertAtReviewAnchor(dynamic document, string anchor, string text)
        {
            var position = (int)FindReviewAnchor(document, anchor).Start;
            document.Range(position, position).InsertBefore(text + "\r");
        }

        private static void VerifyRemoteCollaboration(dynamic application, WordDocuments word, string root)
        {
            var originalName = (string)application.UserName;
            try
            {
                var seedPath = Path.Combine(root, "远程种子稿.docx");
                dynamic seed = application.Documents.Add();
                seed.RemovePersonalInformation = false;
                seed.Content.Text = "远程协作基准\r试验条件尚待核对\r";
                seed.Sections[1].Headers[1].Range.Text = "远程协作页眉";
                seed.PageSetup.TopMargin = 55f;
                dynamic table = seed.Tables.Add(seed.Range(seed.Content.End - 1, seed.Content.End - 1), 2, 2);
                table.Cell(1, 1).Range.Text = "龄期";
                var imagePath = Path.Combine(root, "远程协作图片.png");
                using (var bitmap = new System.Drawing.Bitmap(20,20))
                { bitmap.SetPixel(5,5,System.Drawing.Color.Blue); bitmap.Save(imagePath,System.Drawing.Imaging.ImageFormat.Png); }
                seed.InlineShapes.AddPicture(FileName: imagePath, LinkToFile: false, SaveWithDocument: true, Range: seed.Range(0,0));
                seed.SaveAs2(FileName: seedPath, FileFormat: 16, AddToRecentFiles: false);
                seed.Close(SaveChanges: 0);
                var server = new GitHistory(seedPath, Path.Combine(root, "server"));
                var basis = server.CommitFile(seedPath, "共同版本", "作者", null);
                var aPath = Path.Combine(root, "远程作者甲.docx");
                var bPath = Path.Combine(root, "远程作者乙.docx");
                var a = GitHistory.CloneRemote(server.RepositoryPath, "main", aPath, Path.Combine(root, "remote-a"), true);
                var b = GitHistory.CloneRemote(server.RepositoryPath, "main", bPath, Path.Combine(root, "remote-b"), true);
                dynamic aDoc = word.Open(aPath, false, true);
                aDoc.TrackRevisions = false;
                aDoc.Range(0,0).InsertBefore("作者甲新增的试验结果\r");
                WordDocuments.SaveDocument(aDoc);
                var aHash = WordDocuments.FileChecksum(aPath);
                var local = a.CommitFile(aPath, "甲修改", "甲", basis);
                // Keep the first client profile stable; raw revision authors are captured from the peer file.
                dynamic bDoc = word.Open(bPath, false, false);
                bDoc.RemovePersonalInformation = false;
                bDoc.TrackRevisions = true;
                dynamic range = bDoc.Content.Duplicate;
                Require((bool)range.Find.Execute(FindText: "试验条件尚待核对"), "Remote edit anchor");
                range.Text = "作者乙补充了试验条件";
                dynamic comment = bDoc.Comments.Add(bDoc.Range(0,2), "合作者乙的统计建议");
                comment.Author = "合作者乙";
                WordDocuments.SaveDocument(bDoc);
                bDoc.Close(SaveChanges: 0);
                var rawRemote = TestFiles.ReadDocx(bPath);
                dynamic rawPreview = word.Open(bPath, true, false);
                var rawAuthors = WordDocuments.ReviewItems((object)rawPreview).Where(item => !item.IsComment).Select(item => item.Author).Distinct().ToArray();
                var rawCommentAuthor = WordDocuments.Comments((object)rawPreview).Single(item => item.Text.Contains("统计建议")).Author;
                Require(rawAuthors.Length > 0, "Saved remote fixture has no revision authors");
                rawPreview.Close(SaveChanges: 0);
                var remote = b.CommitFile(bPath, "乙修改", "乙", basis);
                b.PushRemote(true);
                var state = a.FetchRemote(true);
                Require(a.Head() == local && WordDocuments.FileChecksum(aPath) == aHash &&
                    state.Ahead == 1 && state.Behind == 1, "Remote fetch altered the open draft");
                Require((bool)aDoc.Saved && !(bool)aDoc.ReadOnly, "Remote fixture is dirty/read-only before comparison");
                Console.WriteLine("Running native Word review for divergent remote branches…");
                dynamic review = word.ReviewBranchesExtracted(a, basis, local, remote, "main", "origin/main", true);
                var items = WordDocuments.ReviewItems((object)review);
                Require((bool)aDoc.Saved, "Native comparison marked the original draft as unsaved");
                Require(items.Any(item => item.IsComment && item.Author == rawCommentAuthor) &&
                    items.Any(item => !item.IsComment && rawAuthors.Contains(item.Author)), "Remote review lost original contributor metadata");
                review.AcceptAllRevisions();
                Require(((string)review.Content.Text).Contains("作者甲新增的试验结果") &&
                    ((string)review.Content.Text).Contains("作者乙补充了试验条件") &&
                    (int)review.Tables.Count == 1 && (int)review.InlineShapes.Count == 1 &&
                    ((string)review.Sections[1].Headers[1].Range.Text).Contains("远程协作页眉") &&
                    Math.Abs((float)review.PageSetup.TopMargin - 55f) < 0.1f, "Remote merge lost edits or layout objects");
                var mergedFile = Path.Combine(root, "远程确认结果.docx");
                review.SaveAs2(FileName: mergedFile, FileFormat: 16, AddToRecentFiles: false);
                var mergedHash = WordDocuments.FileChecksum(mergedFile);
                string merged = null;
                using (a.Lock())
                    word.ReplaceWorkingCopy(a, aPath, aHash, mergedFile, mergedHash,
                        () => merged = a.CommitFile(aPath, "确认远程修订", "甲", local, "main", state.RemoteReference, remote));
                Require(a.History().Single(row => row.Id == merged).Parents.SequenceEqual(new[] { local, remote }) &&
                    a.Extract(remote).SequenceEqual(rawRemote), "Remote merge lost either parent or raw peer revision");
                a.PushRemote(true);
                var peerState = b.FetchRemote(true);
                var peerStaged = Path.Combine(root, "乙下载的合并版本.docx");
                WordDocuments.ExtractToFile(b, merged, peerStaged);
                bDoc = word.Open(bPath, false, false);
                var peerHash = WordDocuments.FileChecksum(bPath);
                using (b.Lock())
                    word.ReplaceWorkingCopy(b, bPath, peerHash, peerStaged, mergedHash,
                        () => b.FastForwardRemote("main", remote, peerState.RemoteReference, merged));
                Require(b.Head() == merged && WordDocuments.FileChecksum(bPath) == mergedHash && server.Head() == merged,
                    "Merged document did not round-trip through the remote");
                bDoc = word.Current();
                bDoc.TrackRevisions = false;
                bDoc.Range(0,0).InsertBefore("乙下一轮更新\r");
                WordDocuments.SaveDocument(bDoc);
                var next = b.CommitFile(bPath, "乙下一轮", "乙", merged);
                b.PushRemote(true);
                state = a.FetchRemote(true);
                var staged = Path.Combine(root, "甲下载的下一轮.docx");
                WordDocuments.ExtractToFile(a, next, staged);
                var nextHash = WordDocuments.FileChecksum(staged);
                aDoc = word.Open(aPath, false, false);
                aDoc.TrackRevisions = false;
                aDoc.Range(0,0).InsertBefore("甲未保存的修改\r");
                var refused = false;
                try {
                    using (a.Lock())
                        word.ReplaceWorkingCopy(a, aPath, mergedHash, staged, nextHash,
                            () => a.FastForwardRemote("main", merged, state.RemoteReference, next));
                } catch (InvalidOperationException) { refused = true; }
                Require(refused && a.Head() == merged && WordDocuments.FileChecksum(aPath) == mergedHash &&
                    ((string)aDoc.Content.Text).Contains("甲未保存的修改"), "Pull overwrote unsaved Word edits");
                aDoc.Close(SaveChanges: 0);
                aDoc = word.Open(aPath, false, false);
                using (a.Lock())
                    word.ReplaceWorkingCopy(a, aPath, mergedHash, staged, nextHash,
                        () => a.FastForwardRemote("main", merged, state.RemoteReference, next));
                Require(a.Head() == next && WordDocuments.FileChecksum(aPath) == nextHash &&
                    a.Extract(remote).SequenceEqual(rawRemote), "Clean remote pull lost history or document bytes");
                Console.WriteLine("PASS: real Word two-client clone/fetch/native revision merge/two-parent commit/push/pull, contributor metadata, layout objects and unsaved-edit protection.");
            }
            finally { application.UserName = originalName; }
        }

        private static void VerifyImportedReview(dynamic application, WordDocuments word, string root)
        {
            var userName = (string)application.UserName;
            try
            {
                var path = Path.Combine(root, "返稿导入-主稿.docx");
                dynamic source = application.Documents.Add();
                source.RemovePersonalInformation = false;
                source.Content.Text = "送审标题\r原始试验说明\r应删除的旧描述\r";
                source.Sections[1].Headers[1].Range.Text = "送审页眉";
                source.PageSetup.TopMargin = 55f;
                dynamic end = source.Range(source.Content.End - 1, source.Content.End - 1);
                dynamic table = source.Tables.Add(end, 2, 2);
                table.Cell(1, 1).Range.Text = "龄期";
                table.Cell(1, 2).Range.Text = "试件";
                var picture = Path.Combine(root, "返稿导入-图片.png");
                using (var bitmap = new System.Drawing.Bitmap(24, 24))
                { bitmap.SetPixel(10, 10, System.Drawing.Color.Red); bitmap.Save(picture, System.Drawing.Imaging.ImageFormat.Png); }
                source.InlineShapes.AddPicture(FileName: picture, LinkToFile: false, SaveWithDocument: true,
                    Range: source.Range(0, 0));
                source.SaveAs2(FileName: path, FileFormat: 16, AddToRecentFiles: false);
                var repository = new GitHistory(path, root);
                var basis = repository.CommitFile(path, "送审 A", "作者", null);
                source.TrackRevisions = false;
                source.Range(0, 0).InsertBefore("我在送审后新增的试验结果\r");
                WordDocuments.SaveDocument(source);
                var target = repository.CommitFile(path, "继续修改 B", "作者", basis);
                var targetHash = WordDocuments.FileChecksum(path);
                application.UserName = "导师甲";
                var commentsPath = Path.Combine(root, "只有批注的老师返稿.docx");
                WordDocuments.ExtractToFile(repository, basis, commentsPath);
                dynamic comments = word.Open(commentsPath, false, false);
                comments.Comments.Add(comments.Range(0, 2), "仅批注：请补充统计方法");
                WordDocuments.SaveDocument(comments);
                comments.Close(SaveChanges: 0);
                var commentsId = repository.ImportReviewFile(commentsPath, basis, "review/仅批注", "导师甲", "main", target);
                dynamic commentsMerge = word.ReviewBranchesExtracted(repository, basis, target, commentsId);
                commentsMerge.AcceptAllRevisions();
                var keptCurrent = ((string)commentsMerge.Content.Text).Contains("我在送审后新增的试验结果");
                var keptComment = WordDocuments.Comments((object)commentsMerge).Any(item => item.Text.Contains("仅批注"));
                Require(keptCurrent && keptComment, "Comment-only review lost current edits or teacher comments (current=" +
                    keptCurrent + ", comment=" + keptComment + ")");
                commentsMerge.Close(SaveChanges: 0);

                var teacherPath = Path.Combine(root, "有批注和修订的老师返稿.docx");
                WordDocuments.ExtractToFile(repository, basis, teacherPath);
                dynamic teacher = word.Open(teacherPath, false, false);
                teacher.RemovePersonalInformation = false;
                teacher.TrackRevisions = true;
                dynamic replacement = teacher.Content.Duplicate;
                Require((bool)replacement.Find.Execute(FindText: "原始试验说明"), "Missing teacher edit anchor");
                replacement.Text = "老师修改后的试验说明";
                dynamic deletion = teacher.Content.Duplicate;
                Require((bool)deletion.Find.Execute(FindText: "应删除的旧描述"), "Missing teacher deletion anchor");
                deletion.Delete();
                dynamic teacherComment = teacher.Comments.Add(teacher.Range(0, 2), "请核对这个试验结论");
                teacherComment.Author = "导师甲";
                WordDocuments.SaveDocument(teacher);
                Require((int)teacher.Revisions.Count > 0, "Teacher revision fixture has no tracked changes");
                teacher.Close(SaveChanges: 0);
                var raw = TestFiles.ReadDocx(teacherPath);
                dynamic sourcePreview = word.Open(teacherPath, true, false);
                var originalComment = WordDocuments.Comments((object)sourcePreview).Single(item => item.Text.Contains("核对"));
                var originalAuthors = WordDocuments.ReviewItems((object)sourcePreview).Where(item => !item.IsComment).Select(item => item.Author).Distinct().ToArray();
                var originalRevisionCount = (int)sourcePreview.Revisions.Count;
                Require(originalRevisionCount > 0 && originalAuthors.Length > 0, "Saved teacher file lost its fixture revisions");
                sourcePreview.Close(SaveChanges: 0);

                var branchesBeforeDirect = repository.Branches();
                string directArchive;
                using (repository.Lock())
                    directArchive = repository.ImportReviewFile(teacherPath, basis, null, "导师甲", "main", target);
                Console.WriteLine("Running native review without a separate teacher branch…");
                dynamic direct = word.ReviewBranchesExtracted(repository, basis, target, directArchive);
                try
                {
                    var directItems = WordDocuments.ReviewItems((object)direct);
                    Require(directItems.Any(item => !item.IsComment && originalAuthors.Contains(item.Author)) &&
                        directItems.Any(item => item.IsComment && item.Text.Contains("核对")), "Direct merge lost teacher review metadata");
                    direct.AcceptAllRevisions();
                    var directBody = (string)direct.Content.Text;
                    Require(directBody.Contains("我在送审后新增的试验结果") && directBody.Contains("老师修改后的试验说明"),
                        "Direct merge lost either the teacher edits or post-submission work");
                    Require(repository.Head() == target && repository.Branches().SequenceEqual(branchesBeforeDirect) &&
                        repository.Extract(directArchive).SequenceEqual(raw) &&
                        WordDocuments.FileChecksum(path) == targetHash && TestFiles.ReadDocx(teacherPath).SequenceEqual(raw),
                        "Generating a direct merge changed a working branch or failed to archive the original");
                }
                finally { direct.Close(SaveChanges: 0); }

                string imported;
                using (repository.Lock())
                    imported = repository.ImportReviewFile(teacherPath, basis, "review/导师一审", "导师甲", "main", target);
                Require(repository.Head() == target && WordDocuments.FileChecksum(path) == targetHash &&
                    repository.History().Single(row => row.Id == imported).Parents.SequenceEqual(new[] { basis }),
                    "Import changed the active branch or submission ancestry");
                Require(repository.Extract(imported).SequenceEqual(raw), "Imported DOCX is not byte-exact");
                var previewPath = Path.Combine(root, "原始返稿预览.docx");
                WordDocuments.ExtractToFile(repository, imported, previewPath);
                dynamic preview = word.Open(previewPath, true, false);
                Require((int)preview.Revisions.Count > 0 && WordDocuments.Comments((object)preview).Any(item =>
                    item.Author == originalComment.Author && item.Text.Contains("核对")), "Preview lost teacher review metadata");
                preview.Close(SaveChanges: 0);
                Console.WriteLine("Running native review for imported teacher branch…");
                dynamic resolved = word.ReviewBranchesExtracted(repository, basis, target, imported, "main", "review/导师一审");
                var items = WordDocuments.ReviewItems((object)resolved);
                Require(items.Any(item => !item.IsComment && originalAuthors.Contains(item.Author)) &&
                    items.Any(item => item.IsComment && item.Author == originalComment.Author && item.Text.Contains("核对")),
                    "Native review lost original comment or revision authors");
                var reject = items.FirstOrDefault(item => !item.IsComment && item.RevisionType == 2 &&
                    item.RangeText.Contains("应删除的旧描述"));
                Require(reject != null, "Native merge lost the reviewable deletion");
                word.ResolveReviewItem((object)resolved, reject, false);
                resolved.AcceptAllRevisions();
                var body = (string)resolved.Content.Text;
                Require(body.Contains("我在送审后新增的试验结果") && body.Contains("老师修改后的试验说明") &&
                    body.Contains("应删除的旧描述"), "Accept/reject did not retain both sides and the rejected deletion");
                Require((int)resolved.Tables.Count == 1 && (int)resolved.InlineShapes.Count == 1 &&
                    ((string)resolved.Sections[1].Headers[1].Range.Text).Contains("送审页眉") &&
                    Math.Abs((float)resolved.PageSetup.TopMargin - 55f) < 0.1f, "Imported review merge changed layout objects");
                var mergedPath = Path.Combine(root, "返稿导入-确认结果.docx");
                resolved.SaveAs2(FileName: mergedPath, FileFormat: 16, AddToRecentFiles: false);
                var mergedHash = WordDocuments.FileChecksum(mergedPath);
                string merged = null;
                using (repository.Lock())
                    word.ReplaceWorkingCopy(repository, path, targetHash, mergedPath, mergedHash,
                        () => merged = repository.CommitFile(path, "确认老师一审", "作者", target,
                            "main", "review/导师一审", imported));
                Require(repository.Head() == merged && repository.History().Single(row => row.Id == merged).Parents.SequenceEqual(new[] { target, imported }) &&
                    WordDocuments.FileChecksum(path) == mergedHash, "Imported review did not create a two-parent merge");
                repository.Repack();
                Require(repository.Extract(imported).SequenceEqual(raw) && TestFiles.ReadDocx(teacherPath).SequenceEqual(raw),
                    "Review source or its commit changed after merging/repacking");
                Console.WriteLine("PASS: imported raw teacher DOCX, comments-only review, original authors and revisions, native accept/reject, current edits, image/table/header/margin and traceable merge.");
            }
            finally { application.UserName = userName; }
        }

        private static void VerifyRevisionActions(object value, WordDocuments word, string root)
        {
            dynamic application = value;
            dynamic document = application.Documents.Add();
            var path = Path.Combine(root, "逐条审阅测试.docx");
            document.Content.Text = "保留正文\r删除后接受\r删除后拒绝\r格式测试\r";
            document.Sections[1].Headers[1].Range.Text = "页眉原文";
            document.SaveAs2(FileName: path, FileFormat: 16, AddToRecentFiles: false);
            var original = TestFiles.ReadDocx(path);
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
            // Anchor to unchanged text: WPS splits a revision at a comment boundary.
            dynamic commentAnchor = document.Content.Duplicate;
            Require((bool)commentAnchor.Find.Execute(FindText: "保留正文"), "Revision fixture comment anchor");
            document.Comments.Add(commentAnchor, "保留批注");
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
                !(bool)document.Saved && TestFiles.ReadDocx(path).SequenceEqual(original), "Review actions saved the file or changed comments");
            document.Close(SaveChanges: 0);
            document = word.Open(path, true, false);
            var readOnlyRejected = false;
            try { word.ResolveReviewItem((object)document, itemToResolve, true); }
            catch (InvalidOperationException) { readOnlyRejected = true; }
            Require(readOnlyRejected, "Read-only revision resolution");
            document.Close(SaveChanges: 0);
            Console.WriteLine("PASS: native accept/reject for insertion, deletion, formatting and header revisions, Undo, stale-item and read-only protection.");
        }

        private static void VerifyPaneLoading(dynamic application, string root)
        {
            // User documents live outside WordGit\'s protected preview/session directory.
            var path = Path.Combine(Path.GetTempPath(), "wg-pane-" + Guid.NewGuid().ToString("N") + ".docx");
            var storage = Path.Combine(root, "pane-repositories");
            dynamic source = application.Documents.Add();
            try
            {
                source.Content.Text = "异步侧栏测试。\r";
                source.SaveAs2(path, 16);
                var repo = new GitHistory(path, storage);
                repo.CommitFile(path, "初稿", "WordGit test", null);
                using (var host = new System.Windows.Forms.Form {
                    ClientSize = new System.Drawing.Size(420, 720), Opacity = 0, ShowInTaskbar = false })
                using (var pane = new HistoryPane { Dock = System.Windows.Forms.DockStyle.Fill })
                using (var heartbeat = new System.Windows.Forms.Timer { Interval = 15 })
                {
                    host.Controls.Add(pane);
                    host.Show();
                    var ticks = 0;
                    heartbeat.Tick += (sender, args) => ticks++;
                    heartbeat.Start();
                    var method = typeof(Connect).GetMethod("LoadPaneAsync",
                        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                    pane.BindAsync(() => (System.Threading.Tasks.Task<HistoryPaneData>)method.Invoke(null,
                        new object[] { (object)source, storage }), (documentPath, action, first, second) => { });
                    var clock = Stopwatch.StartNew();
                    var loaded = pane.ReloadAsync();
                    var dispatch = clock.ElapsedMilliseconds;
                    var deadline = DateTime.UtcNow.AddSeconds(20);
                    while ((!loaded.IsCompleted || !PaneLabels(pane).Contains("与最新提交一致")) && DateTime.UtcNow < deadline)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        System.Threading.Thread.Sleep(1);
                    }
                    Require(loaded.IsCompleted, "Word pane loading timed out");
                    loaded.GetAwaiter().GetResult();
                    Require(ticks > 0 && PaneLabels(pane).Contains("当前分支  main") &&
                        PaneLabels(pane).Contains("与最新提交一致"), "Word pane did not update after background loading: " + String.Join(" / ", PaneLabels(pane)));
                    source.TrackRevisions = true;
                    source.Content.InsertAfter("继续编辑。");
                    pane.UpdateWorkingState();
                    Require(PaneLabels(pane).Contains("有未保存修改"), "Word pane lost edits made during loading");
                    Console.WriteLine("PASS: real Word asynchronous pane; dispatch " + dispatch +
                        " ms, complete " + clock.ElapsedMilliseconds + " ms, UI heartbeat " + ticks + ".");
                }
            }
            finally { source.Close(false); File.Delete(path); }
        }

        private static void VerifyDialogOwners(dynamic application, WordDocuments word)
        {
            dynamic document = application.Documents.Add();
            try
            {
                var owner = word.Owner;
                Require(owner != null && owner.Handle != IntPtr.Zero, "Word dialog owner was not found");
                using (var about = new AboutDialog { Opacity = 0 })
                {
                    DialogPlacement.Show(about, owner);
                    System.Windows.Forms.Application.DoEvents();
                    VerifyNativeCenter(about, owner);
                }
                System.Windows.Forms.TextBox subject, details;
                using (var commit = Dialogs.CommitForm("对话框位置验证", "", out subject, out details))
                {
                    commit.Opacity = 0;
                    DialogPlacement.Show(commit, owner);
                    System.Windows.Forms.Application.DoEvents();
                    VerifyNativeCenter(commit, owner);
                    Require(details.Height > 180 && subject.Height < details.Height, "Native commit input layout");
                }
                Console.WriteLine("PASS: Word-owned commit and about windows centered on the native document frame.");
            }
            finally { document.Close(false); }
        }

        private static void VerifyNativeCenter(System.Windows.Forms.Form dialog, System.Windows.Forms.IWin32Window owner)
        {
            DialogRectangle rectangle;
            Require(GetDialogWindowRect(owner.Handle, out rectangle), "Cannot read native Word window bounds");
            var bounds = System.Drawing.Rectangle.FromLTRB(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom);
            var expected = DialogPlacement.CenteredBounds(dialog.Size, bounds,
                System.Windows.Forms.Screen.FromHandle(owner.Handle).WorkingArea);
            Require(Math.Abs(expected.Left - dialog.Left) <= 2 && Math.Abs(expected.Top - dialog.Top) <= 2,
                "Word dialog did not appear in the center of its document window");
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct DialogRectangle { internal int Left, Top, Right, Bottom; }
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowRect")]
        private static extern bool GetDialogWindowRect(IntPtr handle, out DialogRectangle rectangle);

        private static System.Collections.Generic.IEnumerable<string> PaneLabels(System.Windows.Forms.Control root)
        {
            foreach (System.Windows.Forms.Control child in root.Controls)
            {
                if (child is System.Windows.Forms.Label) yield return child.Text;
                foreach (var label in PaneLabels(child)) yield return label;
            }
        }

        private static void Require(bool condition, string detail)
        {
            if (!condition) throw new InvalidOperationException(detail);
        }
    }
}
