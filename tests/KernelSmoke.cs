using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace WordGit
{
    // Real-machine smoke test for the Rust core: the Windows add-in talks to the local service,
    // and Word still produces the tracked-changes merge. Covers commit, preview, restore and a
    // branch merge on a throwaway document.
    internal static class KernelSmoke
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            var folder = Path.Combine(Path.GetTempPath(), "wordgit-kernelsmoke-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            dynamic application = null;
            ServiceClient client = null;
            try
            {
                application = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application", true));
                application.Visible = false;
                application.DisplayAlerts = 0;
                Console.WriteLine("Testing Word " + application.Version);
                client = TestKernel.Start();
                VerifyDocumentFlows(application, new WordDocuments(application), client, folder);
                VerifyReturnedDrafts(application, new WordDocuments(application), client, folder);
                Console.WriteLine("PASS: 提交、预览、恢复与分支合并都通过 Rust 内核完成。");
            }
            catch (Exception error)
            {
                Console.WriteLine("FAILED: " + Describe(error));
                throw;
            }
            finally
            {
                if (application != null)
                {
                    try { application.Quit(0); } catch (Exception) { }
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(application); } catch (Exception) { }
                }
                if (client != null) TestKernel.Stop(client);
                try { Directory.Delete(folder, true); } catch (IOException) { }
            }
        }

        private static string Describe(Exception error)
        {
            try
            {
                var text = error.GetType().FullName + ": " + error.Message;
                if (error.InnerException != null)
                    text += " / " + error.InnerException.GetType().FullName + ": " + error.InnerException.Message;
                return text;
            }
            catch (Exception)
            {
                return "unprintable exception";
            }
        }

        private static void VerifyDocumentFlows(dynamic application, WordDocuments word, ServiceClient client, string folder)
        {
            var path = Path.Combine(folder, "内核测试稿.docx");
            dynamic document = application.Documents.Add();
            document.Content.Text = "内核迁移测试稿\r第一节 提交与预览。\r";
            document.SaveAs2(FileName: path, FileFormat: 16, AddToRecentFiles: false);
            var first = TestFiles.Snapshot((object)document);
            var head1 = Commit(client, path, "第一节");
            Require(head1 != null, "第一次提交应产生提交");
            Require(client.Status(path).History.Count == 1, "提交历史应有一条记录");
            Console.WriteLine("PASS: 首次提交入库。");

            TestKernel.VerifyRestart(client);
            Require(client.Status(path).Head == head1, "内核重启后应保留原来的提交");
            Console.WriteLine("PASS: 内核退出后自动启动，原有提交可读取。");

            var workingCopy = new WorkingCopy(path, () => client.WorkingState(path));
            Require(workingCopy.State() == WorkingCopyState.Clean, "保存前应与最新提交一致");
            document.Content.InsertAfter("第二节 恢复与合并。\r");
            document.Save();
            Require(workingCopy.State() == WorkingCopyState.Changed, "Word 保存后侧栏应识别未提交修改");
            Console.WriteLine("PASS: Word 实际保存后识别出未提交修改。");
            var second = TestFiles.Snapshot((object)document);
            Require(!first.SequenceEqual(second), "第二次保存应改变文件内容");
            var head2 = Commit(client, path, "第二节");
            Require(head2 != null && head2 != head1, "第二次提交应产生新的提交");
            var status = client.Status(path);
            Require(status.History.Count == 2 && status.Head == head2, "历史应有两条提交且 HEAD 指向最新提交");
            Require(status.WorkingState == "clean", "刚提交后工作稿应与最新提交一致，实际为 " + status.WorkingState);
            Console.WriteLine("PASS: 第二次提交后工作稿与历史一致。");

            var preview = client.Call("version/preview", ServiceClient.Args("document", path, "revision", head1));
            var previewFile = Json.String(preview, "file");
            Require(previewFile != null && File.Exists(previewFile), "预览应返回已导出的文件");
            Require(TestFiles.ReadDocx(previewFile).SequenceEqual(first), "预览的第一版应与提交时的字节一致");
            try { File.Delete(previewFile); } catch (IOException) { }
            Console.WriteLine("PASS: 预览旧版逐字节还原。");

            var restore = client.Call("version/restore", ServiceClient.Args(
                "document", path, "revision", head1, "author", "小测试"));
            var plan = Json.String(restore, "plan");
            Require(plan != null, "恢复旧版应返回替换计划");
            document.Close(0);
            var message = client.ApplyPlan(plan);
            Require(!String.IsNullOrEmpty(message), "替换计划应返回结果说明");
            Require(TestFiles.ReadDocx(path).SequenceEqual(first), "恢复后工作稿应与第一版逐字节一致");
            var afterRestore = client.Status(path);
            Require(afterRestore.History.Count == 3, "恢复前应把当前稿再提交一份，历史应为三条，实际 " + afterRestore.History.Count);
            Console.WriteLine("PASS: 恢复到旧版并保留恢复前的版本。");

            document = application.Documents.Open(path);
            document.Content.InsertAfter("第二节 恢复与合并。\r");
            document.Save();
            Require(Commit(client, path, "重新写回第二节") != null, "恢复后应能继续提交");
            var mainBranch = afterRestore.Branch;
            Require(client.Call("branch/create", ServiceClient.Args(
                "document", path, "name", "实验分支", "author", "小测试")) != null, "应能新建分支");
            var switched = client.Call("branch/switch", ServiceClient.Args(
                "document", path, "name", "实验分支", "author", "小测试"));
            document.Close(0);
            client.ApplyPlan(Json.String(switched, "plan"));
            document = application.Documents.Open(path);
            Require(client.Status(path).Branch == "实验分支", "切换后应位于新分支");

            document.Content.InsertAfter("第三节 只存在于实验分支。\r");
            document.Save();
            Commit(client, path, "第三节");
            var switchedBack = client.Call("branch/switch", ServiceClient.Args(
                "document", path, "name", mainBranch, "author", "小测试"));
            document.Close(0);
            client.ApplyPlan(Json.String(switchedBack, "plan"));
            document = application.Documents.Open(path);
            Require(client.Status(path).Branch == mainBranch, "应回到主分支");

            var context = client.Call("merge/context", ServiceClient.Args("document", path, "source", "实验分支"));
            var merge = MergePlan.From(context);
            Require(merge.Draft != null && merge.Basis != null && merge.Incoming != null, "合并上下文应准备三个版本文件");
            word.BuildBranchMergeFile(merge, merge.Draft);
            Require(File.Exists(merge.Draft), "应生成合并稿");
            Require(merge.Review != null, "合并稿应带有待完成的审阅信息");
            client.Call("merge/register", ServiceClient.Args("draft", merge.Draft, "review", ReviewPayload(merge.Review)));
            var draft = (dynamic)word.Open(merge.Draft, false);
            Require(WordDocuments.DocumentRevisionCount(draft) > 0, "合并稿应包含 Word 修订");
            Console.WriteLine("PASS: 分支合并生成带修订的合并稿。");

            var finish = client.Call("merge/finish", ServiceClient.Args(
                "draft", merge.Draft, "message", "合并实验分支", "author", "小测试"));
            var finishPlan = Json.String(finish, "plan");
            Require(finishPlan != null, "完成合并应返回替换计划");
            draft.Close(0);
            document.Close(0);
            client.ApplyPlan(finishPlan);
            document = application.Documents.Open(path);
            var text = (string)document.Content.Text;
            Require(text.Contains("第二节") && text.Contains("第三节"), "合并后当前稿应包含两个分支的内容，实际：" + text.Replace("\r", " "));
            Require(client.Status(path).Review == null, "完成合并后不应再留有审阅信息");
            Console.WriteLine("PASS: 合并结果已写回当前稿。");
            document.Close(0);

            // 合并两个历史版本：不新建分支，只把两份旧版交给 Word 生成修订稿再提交。
            var versions = client.Status(path).History;
            Require(versions.Count >= 2, "应有至少两个历史版本可供合并");
            var pair = MergePlan.From(client.Call("merge/versions", ServiceClient.Args(
                "document", path, "revision", versions[0].Id, "other", versions[versions.Count - 1].Id, "author", "小测试")));
            word.BuildVersionReviewFile(pair.Original, pair.Revised, null, true, pair.Draft);
            Require(File.Exists(pair.Draft), "应生成两个历史版本的合并稿");
            client.Call("merge/register", ServiceClient.Args("draft", pair.Draft, "review", ReviewPayload(pair.Review)));
            var merged = (dynamic)word.Open(pair.Draft, false);
            Require(WordDocuments.DocumentRevisionCount(merged) > 0, "历史版本合并稿应包含 Word 修订");
            merged.Revisions.AcceptAll();
            merged.Save();
            merged.Close(0);
            var mergedCommit = Json.String(client.Call("commit", ServiceClient.Args(
                "document", pair.Draft, "message", "合并两个历史版本", "author", "小测试")), "commit");
            Require(mergedCommit != null, "合并稿应能提交");
            Console.WriteLine("PASS: 合并两个历史版本并提交合并稿。");
        }

        private static void VerifyReturnedDrafts(dynamic application, WordDocuments word, ServiceClient client, string folder)
        {
            var path = Path.Combine(folder, "主稿返稿验收.docx");
            dynamic main = application.Documents.Add();
            main.Content.Text = "论文初稿\r材料配比与研究背景。\r统计结果与结论。";
            main.Sections[1].Headers[1].Range.Text = "保留的页眉";
            main.SaveAs2(FileName: path, FileFormat: 16, AddToRecentFiles: false);
            var basis = Commit(client, path, "送审版");
            var sent = TestFiles.ReadDocx(path);
            main.TrackRevisions = true;
            main.Range(0, 0).InsertBefore("未保存对比标记\r");
            int originalRevisions = main.Revisions.Count;
            var comparison = ComparePlan.From(client.Call("version/compare", ServiceClient.Args(
                "document", path, "revision", basis, "currentInMemory", true)));
            var snapshot = WordDocuments.CaptureComparisonSource((object)main, AppPaths.NewSession());
            dynamic compared = word.CompareSnapshot(comparison.Original, snapshot);
            Require(((string)compared.Content.Text).Contains("未保存对比标记") && (int)compared.Revisions.Count > 0,
                "Word 对比漏掉未保存的文字。");
            Require(!(bool)main.Saved && (int)main.Revisions.Count == originalRevisions && TestFiles.ReadDocx(path).SequenceEqual(sent),
                "Word 对比修改了原稿或其状态。");
            compared.Close(0);
            main.Close(0);
            main = word.Open(path, false);
            Console.WriteLine("PASS: Word 原生对比包含未保存修订，且未保存或修改原稿。");

            var previous = Path.Combine(folder, "原始送审稿.docx");
            File.WriteAllBytes(previous, sent);
            for (var round = 1; round <= 2; round++)
            {
                var file = Path.Combine(folder, "导师返稿-" + round + ".docx");
                File.Copy(previous, file);
                dynamic returned = word.Open(file, false);
                returned.TrackRevisions = true;
                int end = (int)returned.Content.End - 1;
                var incoming = "导师补充第" + round + "轮";
                returned.Range(end, end).InsertAfter("\r" + incoming);
                returned.Comments.Add(returned.Range(0, 4), "审阅意见 " + round);
                returned.Save();
                returned.Close(0);
                previous = file;
                main.Activate();
                var author = "主稿补充第" + round + "轮";
                main.Range(0, 0).InsertBefore(author + "\r");
                Require(!(bool)main.Saved && client.Status(path).WorkingState == "clean", "要覆盖仅在编辑器内存中的未提交修改。");
                WordDocuments.SaveDocument((object)main);
                var submitted = Commit(client, path, "合并前提交本地修改");
                var request = ServiceClient.Args("document", path, "file", file, "basis", basis,
                    "branch", round == 1 ? null : "review-2", "reviewer", "导师", "archiveOnly", true,
                    "expectedBranch", client.Status(path).Branch, "expectedHead", submitted);
                var imported = client.Call("review/import", request);
                var source = Json.String(imported, "source");
                var archived = Json.String(imported, "commit");
                Require(!Json.Bool(imported, "reused"), "新一轮返稿应产生独立归档。");
                Require(client.Status(path).History.Any(r => r.Id == archived), "原始返稿没有进入提交图。");
                var preview = client.Call("version/preview", ServiceClient.Args("document", path, "revision", archived));
                Require(TestFiles.ReadDocx(Json.String(preview, "file")).SequenceEqual(TestFiles.ReadDocx(file)),
                    "原始返稿的批注、修订或排版没有按字节保留。");
                var plan = MergePlan.From(client.Call("merge/context", ServiceClient.Args("document", path, "source", source, "basis", basis)));
                Require(plan.SubmittedBase == basis && plan.Review.SourceTip == archived &&
                    plan.PreviousReturns.Count == round - 1, "合并没有保留所选送审版本或前一轮返稿。");
                word.BuildBranchMergeFile(plan, plan.Draft);
                client.Call("merge/register", ServiceClient.Args("draft", plan.Draft, "review", ReviewPayload(plan.Review)));
                dynamic draft = word.Open(plan.Draft, false);
                Require((int)draft.Revisions.Count > 0 && (int)draft.Comments.Count > 0, "应生成带原生修订和批注的合并稿。");
                draft.AcceptAllRevisions();
                Require(((string)draft.Content.Text).Contains(incoming) && ((string)draft.Content.Text).Contains(author) &&
                    ((string)draft.Sections[1].Headers[1].Range.Text).Contains("保留的页眉"), "合并丢失内容或页眉。");
                draft.Save();
                var resolved = TestFiles.ReadDocx(plan.Draft);
                var replace = client.Call("merge/finish", ServiceClient.Args("draft", plan.Draft, "message", "确认导师返稿", "author", "验收"));
                draft.Close(0);
                main.Close(0);
                client.ApplyPlan(Json.String(replace, "plan"));
                main = word.Open(path, false);
                var status = client.Status(path);
                Require(status.History[0].Parents.SequenceEqual(new[] { submitted, archived }) && TestFiles.ReadDocx(path).SequenceEqual(resolved),
                    "合并结果必须以原始返稿为第二父提交，并原样写回。");
                if (round == 2) request["branch"] = "review-duplicate";
                request["expectedHead"] = status.Head;
                var repeated = client.Call("review/import", request);
                var noChanges = client.Call("merge/context", ServiceClient.Args("document", path, "source", Json.String(repeated, "source")));
                Require(Json.Bool(repeated, "reused") && Json.String(repeated, "commit") == archived && Json.Bool(noChanges, "upToDate") &&
                    !client.Status(path).Branches.Contains("review-duplicate"), "重复提交同一返稿应直接复用。");
                Console.WriteLine("PASS: Word 返稿第 " + round + " 轮，原始返稿、双父合并和重复导入均可追溯。");
            }
            main.Close(0);
        }

        private static string Commit(ServiceClient client, string path, string message)
        {
            var status = client.Status(path);
            var body = ServiceClient.Args("document", path, "message", message, "author", "小测试");
            if (status.Head != null) body["expectedHead"] = status.Head;
            return Json.String(client.Call("commit", body), "commit");
        }

        private static System.Collections.Generic.Dictionary<string, object> ReviewPayload(PendingReview review)
        {
            return ServiceClient.Args(
                "original", review.Original, "head", review.Head, "branch", review.Branch,
                "sourceBranch", review.SourceBranch, "sourceTip", review.SourceTip,
                "applyToOriginal", review.ApplyToOriginal, "recordOnlyMerge", review.RecordOnlyMerge);
        }

        private static void Require(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }
    }
}
