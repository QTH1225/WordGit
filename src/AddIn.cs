using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: ComVisible(false)]
[assembly: System.Reflection.AssemblyVersion("0.9.6.0")]

namespace WordGit
{
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int mode,
            [MarshalAs(UnmanagedType.IDispatch)] object instance, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int mode, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    [ComVisible(true), Guid("C899209F-FEEC-43D3-A795-8CBA4C979E96"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IWordGitActions
    {
        [DispId(1)] void OnCommit([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(2)] void OnHistory([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(3)] void OnPreview([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(4)] void OnCompare([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(5)] void OnExport([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(6)] void OnRestore([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(7)] void OnMergeReview([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(8)] void OnCreateBranch([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(9)] void OnSwitchBranch([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(10)] void OnMergeBranch([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(11)] void OnDeleteBranch([MarshalAs(UnmanagedType.IDispatch)] object control);

        [DispId(12)] void OnCompressHistory([MarshalAs(UnmanagedType.IDispatch)] object control);
    }

    [ComVisible(true), Guid("C5707F76-F46F-4A11-800B-18E7D98377D3"), ProgId("WordGit.Connect")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IWordGitActions))]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility, IWordGitActions, ICustomTaskPaneConsumer
    {
        private dynamic application;
        private WordDocuments word;
        private TaskPanes panes;
        private HistoryMaintenance maintenance;
        private bool committing;

        public void OnConnection(object app, int mode, object instance, ref Array custom)
        {
            application = app;
            word = new WordDocuments(app);
            maintenance = new HistoryMaintenance(AppPaths.Repositories);
        }

        public void CTPFactoryAvailable(ICTPFactory factory) { panes = new TaskPanes(factory); }

        public void OnDisconnection(int mode, ref Array custom)
        {
            if (maintenance != null) maintenance.Dispose();
            maintenance = null;
            if (panes != null) panes.Dispose();
            panes = null;
            application = null;
            word = null;
        }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
<ribbon><tabs><tab id=""WordGitTab"" label=""WordGit"">
  <group id=""WordGitVersions"" label=""版本记录"">
    <button id=""WordGitCommit"" label=""提交版本"" size=""large"" imageMso=""FileSave"" onAction=""OnCommit"" supertip=""保存当前文档并创建本地 Git 提交。"" />
    <button id=""WordGitHistory"" label=""提交图"" size=""large"" imageMso=""ViewOutlineView"" onAction=""OnHistory"" supertip=""在 Word 右侧查看分支和提交记录，预览或对比历史版本。"" />
    <button id=""WordGitPreview"" label=""预览版本"" size=""large"" imageMso=""FilePrintPreview"" onAction=""OnPreview"" supertip=""选一个历史版本，在 Word 中只读打开。"" />
    <button id=""WordGitCompressHistory"" label=""压缩历史"" size=""large"" imageMso=""PicturesCompress"" onAction=""OnCompressHistory"" supertip=""无损打包当前文档的 Git 历史，完成后显示对象占用变化。"" />
  </group>
  <group id=""WordGitBranches"" label=""分支"">
    <button id=""WordGitCreateBranch"" label=""新建分支"" size=""large"" imageMso=""FileNew"" onAction=""OnCreateBranch"" supertip=""选择起始版本并输入新分支名称。"" />
    <button id=""WordGitSwitchBranch"" label=""切换分支"" size=""large"" imageMso=""WindowSwitchWindowsMenuWord"" onAction=""OnSwitchBranch"" supertip=""选择要切换到的本地分支。"" />
    <button id=""WordGitMergeBranch"" label=""合并分支"" size=""large"" imageMso=""ReviewAcceptChange"" onAction=""OnMergeBranch"" supertip=""将所选分支合入当前分支，先在 Word 中审阅修订。"" />
    <button id=""WordGitDeleteBranch"" label=""删除分支"" size=""large"" imageMso=""ReviewRejectChange"" onAction=""OnDeleteBranch"" supertip=""删除已经合入当前分支的其他本地分支。"" />
  </group>
  <group id=""WordGitReview"" label=""比较与审阅"">
    <button id=""WordGitCompare"" label=""对比版本"" size=""large"" imageMso=""ReviewCompareMenu"" onAction=""OnCompare"" supertip=""选一版与当前稿对比，或按 Ctrl 选两版对比。"" />
    <button id=""WordGitMergeReview"" label=""合并审阅稿"" size=""large"" imageMso=""ReviewTrackChanges"" onAction=""OnMergeReview"" supertip=""选择基准版和返稿，生成 Word 修订稿。"" />
  </group>
  <group id=""WordGitRetrieve"" label=""取回版本"">
    <button id=""WordGitExport"" label=""另存旧版"" size=""large"" imageMso=""FileSaveAs"" onAction=""OnExport"" supertip=""将所选版本另存为新的 Word 文档。"" />
    <button id=""WordGitRestore"" label=""恢复旧版"" size=""large"" imageMso=""Undo"" onAction=""OnRestore"" supertip=""将历史版本恢复为当前稿，并记录一次新提交。"" />
  </group>
</tab></tabs></ribbon></customUI>";
        }

        public void OnCommit(object control) { Guard(CommitCurrent, true); }
        public void OnHistory(object control)
        {
            if (panes == null) { OpenHistory(null); return; }
            Guard(() =>
            {
                word.Current();
                ShowPane(application.ActiveWindow);
            });
        }
        public void OnPreview(object control) { OpenHistory(HistoryAction.Preview); }
        public void OnCompare(object control) { OpenHistory(HistoryAction.Compare); }
        public void OnExport(object control) { OpenHistory(HistoryAction.Export); }
        public void OnRestore(object control) { OpenHistory(HistoryAction.Restore); }
        public void OnMergeReview(object control) { OpenHistory(HistoryAction.Merge); }
        public void OnCreateBranch(object control) { OpenHistory(HistoryAction.CreateBranch); }
        public void OnSwitchBranch(object control) { OpenHistory(HistoryAction.SwitchBranch); }
        public void OnMergeBranch(object control) { OpenHistory(HistoryAction.MergeBranch); }
        public void OnDeleteBranch(object control) { OpenHistory(HistoryAction.DeleteBranch); }

        public void OnCompressHistory(object control)
        {
            Guard(() =>
            {
                var path = HistoryPath(WordDocuments.LocalPath(word.Current()));
                var repo = new GitHistory(path);
                if (repo.Head() == null) { Notice("尚无提交可压缩。"); return; }
                Dialogs.ShowHistoryCompression(word.Owner, Path.GetFileName(path), repo.Repack);
            });
        }

        private void OpenHistory(HistoryAction? command)
        {
            var changesHistory = !command.HasValue || command == HistoryAction.CreateBranch ||
                command == HistoryAction.SwitchBranch || command == HistoryAction.Restore || command == HistoryAction.DeleteBranch;
            Guard(() => { while (ManageHistory(command) && !command.HasValue) { } }, changesHistory);
        }

        private void ShowPane(object window)
        {
            var handle = (int)((dynamic)window).Hwnd;
            panes.Show(window, () =>
            {
                var data = LoadPane(DocumentForWindow(handle));
                data.NavigateReviewItem = item => Guard(() => word.GoToReviewItem(DocumentForWindow(handle), item));
                data.ResolveReviewItem = (item, accept) =>
                {
                    var resolved = false;
                    Guard(() =>
                    {
                        var document = DocumentForWindow(handle);
                        if (!String.Equals(WordDocuments.LocalPath(document), data.DocumentPath, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("此窗口的文档已变化，请先刷新列表。");
                        word.ResolveReviewItem(document, item, accept);
                        resolved = true;
                    });
                    return resolved;
                };
                return data;
            },
                (shownPath, action, revision, other) => Guard(() =>
                {
                    var document = DocumentForWindow(handle);
                    var path = WordDocuments.LocalPath(document);
                    if (!String.Equals(shownPath, path, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("此窗口的文档已变化，请先刷新提交图。");
                    if (action == HistoryAction.Commit)
                    {
                        ((dynamic)document).Activate();
                        CommitCurrent();
                        return;
                    }
                    var historyPath = HistoryPath(path);
                    var repo = new GitHistory(historyPath);
                    if (action == HistoryAction.CompareLatest)
                    {
                        var head = repo.Head();
                        if (head == null) { Notice("尚无提交版本可对比。"); return; }
                        CompareVersions(document, repo, head, null);
                    }
                    else if (action == HistoryAction.Compare) CompareVersions(document, repo, revision, other);
                    else PreviewVersion(historyPath, repo, revision);
                }, action == HistoryAction.Commit));
        }

        private object DocumentForWindow(int handle)
        {
            if (application.Windows.Count > 0)
            {
                object active = application.ActiveWindow;
                if ((int)((dynamic)active).Hwnd == handle) return ((dynamic)active).Document;
                foreach (object window in application.Windows)
                    if ((int)((dynamic)window).Hwnd == handle) return ((dynamic)window).Document;
            }
            throw new InvalidOperationException("此文档窗口已关闭。");
        }

        private static HistoryPaneData LoadPane(object document)
        {
            var path = WordDocuments.LocalPath(document);
            var historyPath = HistoryPath(path);
            var repo = new GitHistory(historyPath);
            var branch = repo.ActiveBranch();
            var head = repo.BranchTip(branch);
            var workingCopy = new WorkingCopy(path, () => (bool)((dynamic)document).Saved,
                head == null ? (Func<string>)null : () => repo.SnapshotChecksum(head));
            return new HistoryPaneData
            {
                DocumentPath = path, HistoryPath = historyPath, Branch = branch,
                Head = head, Branches = repo.Branches(), Load = repo.History, WorkingState = workingCopy.State,
                LoadReviewItems = () => WordDocuments.ReviewItems(document),
                CanResolveRevisions = !(bool)((dynamic)document).ReadOnly
            };
        }

        private static string HistoryPath(string documentPath)
        {
            var review = PendingReview.Load(documentPath);
            return review == null ? documentPath : review.Original;
        }

        private void PreviewVersion(string historyPath, GitHistory repo, string revision)
        {
            var preview = Path.Combine(AppPaths.NewSession(), Path.GetFileNameWithoutExtension(historyPath) + "-" + revision.Substring(0, 8) + ".docx");
            WordDocuments.ExtractToFile(repo, revision, preview);
            word.Open(preview, true);
        }

        private void CompareVersions(object document, GitHistory repo, string basis, string other)
        {
            if (other == null) word.CompareCurrentExtracted(repo, basis, document);
            else word.ReviewFiles(repo, basis, other, false);
        }

        private string Author { get { return (string)application.UserName; } }

        private void CommitCurrent()
        {
            if (committing) { Notice("正在提交，请等待当前提交完成。"); return; }
            var document = word.Current();
            var path = WordDocuments.LocalPath(document, true);
            var review = PendingReview.Load(path);
            var repo = new GitHistory(review == null ? path : review.Original, commitCreated: maintenance.Schedule);
            var branch = repo.ActiveBranch();
            var head = repo.Head();
            if (review != null && (head != review.Head || branch != review.Branch))
                throw new InvalidOperationException("审阅期间原文档的历史已更新。请先比较最新版本，再重新合并。");
            if (review != null && review.SourceBranch != null)
            {
                FinishBranchMerge(document, path, review, repo);
                return;
            }
            var target = review == null ? Path.GetFileName(path) : Path.GetFileName(review.Original);
            var message = Dialogs.Ask(word.Owner, "提交到「" + target + "」 · " + branch, review == null ? "" : "合并审阅稿");
            if (message == null) return;
            var savedPath = WordDocuments.SaveDocument(document);
            var author = Author;
            committing = true;
            Dialogs.ShowBackgroundOperation(word.Owner, "提交版本 · " + target, async () =>
            {
                var id = await System.Threading.Tasks.Task.Run(() =>
                {
                    using (repo.Lock()) return repo.CommitFile(savedPath, message, author, head, branch);
                });
                if (review != null && id != null) review.Save(path, id);
                return id == null ? "文件内容与最新提交一致。" : "已提交 " + id.Substring(0, 8) +
                    "。提交时保存的版本已入库，之后的编辑可继续提交。";
            }, () =>
            {
                committing = false;
                RefreshPane(path);
            });
        }

        private bool ManageHistory(HistoryAction? command)
        {
            var document = word.Current();
            var activePath = WordDocuments.LocalPath(document);
            var review = PendingReview.Load(activePath);
            var historyPath = review == null ? activePath : review.Original;
            var repo = new GitHistory(historyPath, commitCreated: maintenance.Schedule);
            var branch = repo.ActiveBranch();
            var activeTip = repo.BranchTip(branch);
            if (activeTip == null) { Notice("还没有版本，请先点击「提交版本」。"); return false; }
            var branches = repo.Branches();
            var filename = Path.GetFileName(historyPath);
            var deleting = command == HistoryAction.DeleteBranch;
            var chooseBranch = command == HistoryAction.SwitchBranch || command == HistoryAction.MergeBranch || deleting;
            if (review != null && (chooseBranch || command == HistoryAction.CreateBranch || command == HistoryAction.Restore))
                throw new InvalidOperationException("请在原文档中执行此操作。");

            HistoryAction action;
            Revision basis = null, other = null;
            string selectedBranch;
            if (chooseBranch)
            {
                var alternatives = (deleting ? repo.Branches(activeTip) : branches)
                    .Where(name => !String.Equals(name, branch, StringComparison.OrdinalIgnoreCase)).ToArray();
                if (alternatives.Length == 0)
                {
                    Notice(deleting ? "没有可以删除的分支。这里仅列出已合入当前分支的其他分支。"
                        : "当前只有「" + branch + "」分支。请先新建分支。");
                    return false;
                }
                action = command.Value;
                selectedBranch = Dialogs.ChooseBranch(word.Owner, filename, branch, alternatives, action);
                if (selectedBranch == null) return false;
            }
            else
            {
                using (var dialog = new HistoryDialog(filename, repo.History(), review == null, activeTip, branch, branches, repo.History, command))
                {
                    if (dialog.ShowDialog(word.Owner) != DialogResult.OK) return false;
                    action = dialog.Action;
                    basis = dialog.Basis;
                    other = dialog.Other;
                    selectedBranch = dialog.SelectedBranch;
                }
            }
            switch (action)
            {
                case HistoryAction.CreateBranch:
                    var name = Dialogs.Ask(word.Owner, "从 " + basis.Id.Substring(0, 8) + " 新建分支", "", true);
                    if (name == null) return true;
                    using (repo.Lock()) repo.CreateBranch(name, basis.Id);
                    Notice("已创建分支「" + name + "」，当前仍在「" + branch + "」。");
                    return true;
                case HistoryAction.SwitchBranch:
                    SwitchBranch(document, historyPath, repo, branch, activeTip, selectedBranch);
                    return true;
                case HistoryAction.MergeBranch:
                    StartBranchMerge(document, historyPath, repo, branch, activeTip, selectedBranch);
                    break;
                case HistoryAction.DeleteBranch:
                    DeleteBranch(repo, branch, activeTip, selectedBranch);
                    break;
                case HistoryAction.Preview:
                    PreviewVersion(historyPath, repo, basis.Id);
                    break;
                case HistoryAction.Compare:
                    CompareVersions(document, repo, basis.Id, other == null ? null : other.Id);
                    break;
                case HistoryAction.Export:
                    using (var save = new SaveFileDialog { Filter = "Word 文档 (*.docx)|*.docx", FileName = Path.GetFileNameWithoutExtension(historyPath) + "-" + basis.Id.Substring(0, 8) + ".docx", OverwritePrompt = true })
                    {
                        if (save.ShowDialog(word.Owner) != DialogResult.OK) return true;
                        if (File.Exists(save.FileName)) throw new IOException("请选择一个新文件名，避免覆盖已有文件。");
                        WordDocuments.ExtractToFile(repo, basis.Id, save.FileName);
                        word.Open(save.FileName, false);
                    }
                    break;
                case HistoryAction.Restore:
                    if (MessageBox.Show(word.Owner, "先存下当前稿，再将所选旧版恢复到原文件，并创建一次新提交。是否继续？", "WordGit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        using (repo.Lock()) Notice(word.Restore(document, repo, basis.Id, Author));
                    break;
                case HistoryAction.Merge:
                    string externalPath = null;
                    if (other == null)
                    {
                        using (var open = new OpenFileDialog { Filter = "Word 审阅稿 (*.docx)|*.docx", Title = "选择返回的 Word 审阅稿" })
                        {
                            if (open.ShowDialog(word.Owner) != DialogResult.OK) return true;
                            externalPath = open.FileName;
                        }
                    }
                    var head = repo.Head();
                    dynamic result = word.MergeExtracted(repo, basis.Id, other == null ? null : other.Id, externalPath);
                    var resultPath = Path.Combine(AppPaths.NewSession(), Path.GetFileNameWithoutExtension(historyPath) + "-合并稿.docx");
                    result.SaveAs2(FileName: resultPath, FileFormat: 16, AddToRecentFiles: false);
                    new PendingReview { Original = historyPath, Branch = repo.ActiveBranch() }.Save(resultPath, head);
                    Notice("已打开合并稿。请在 Word 的「审阅」中接受或拒绝修改，再点击「提交版本」。");
                    break;
            }
            return false;
        }

        private void DeleteBranch(GitHistory repo, string current, string head, string target)
        {
            var tip = repo.BranchTip(target);
            if (tip == null) throw new InvalidOperationException("找不到要删除的分支。");
            if (MessageBox.Show(word.Owner, "删除已合入「" + current + "」的分支「" + target + "」？\n最后一次提交  " +
                tip.Substring(0, 8) + "。", "WordGit · 删除分支", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;

            using (repo.Lock()) repo.DeleteMergedBranch(current, head, target, tip);
            Notice("已删除分支「" + target + "」。");
        }

        private void SwitchBranch(object document, string path, GitHistory repo, string from, string expected, string target)
        {
            WordDocuments.EnsureReplaceable(path);
            var targetTip = repo.BranchTip(target);
            if (targetTip == null) throw new InvalidOperationException("找不到目标分支。");
            if (!(bool)((dynamic)document).Saved || WordDocuments.FileChecksum(path) != repo.SnapshotChecksum(expected))
            {
                var answer = MessageBox.Show(word.Owner, "当前稿有未提交的修改。是否先提交到「" + from + "」再切换？",
                    "WordGit", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer != DialogResult.Yes) return;
                var message = Dialogs.Ask(word.Owner, "切换前提交到「" + from + "」", "切换前保存当前稿");
                if (message == null) return;
                WordDocuments.SaveDocument(document);
                using (repo.Lock()) expected = repo.CommitFile(path, message, Author, expected, from) ?? expected;
            }
            var staged = Path.Combine(AppPaths.NewSession(), "切换目标.docx");
            WordDocuments.ExtractToFile(repo, targetTip, staged);
            using (repo.Lock())
                word.ReplaceWorkingCopy(path, repo.SnapshotChecksum(expected), staged, repo.SnapshotChecksum(targetTip),
                    () => repo.Switch(from, expected, target, targetTip));
            Notice("已切换到「" + target + "」。");
        }

        private void StartBranchMerge(object document, string path, GitHistory repo, string branch, string tip, string source)
        {
            WordDocuments.EnsureReplaceable(path);
            if (!(bool)((dynamic)document).Saved || WordDocuments.FileChecksum(path) != repo.SnapshotChecksum(tip))
                throw new InvalidOperationException("请先提交当前分支的修改，再合并分支。");
            var sourceTip = repo.BranchTip(source);
            if (sourceTip == null) throw new InvalidOperationException("找不到来源分支。");
            var basis = repo.MergeBase(tip, sourceTip);
            if (basis == sourceTip) { Notice("来源分支的提交已包含在当前分支。"); return; }
            dynamic result = word.ReviewBranchesExtracted(repo, basis, tip, sourceTip, branch, source);
            var resultPath = Path.Combine(AppPaths.NewSession(), Path.GetFileNameWithoutExtension(path) + "-分支合并稿.docx");
            result.SaveAs2(FileName: resultPath, FileFormat: 16, AddToRecentFiles: false);
            new PendingReview { Original = path, Branch = branch, SourceBranch = source, SourceTip = sourceTip }.Save(resultPath, tip);
            Notice("已打开「" + source + " → " + branch + "」的修订稿。请在 Word 审阅中接受或拒绝修改，再点击「提交版本」完成合并。此时分支尚未改变。");
        }

        private void FinishBranchMerge(object document, string path, PendingReview review, GitHistory repo)
        {
            if ((int)((dynamic)document).Revisions.Count > 0)
                throw new InvalidOperationException("这份分支合并稿仍有待处理修订。请先在 Word「审阅」中接受或拒绝修改。");
            if (repo.BranchTip(review.SourceBranch) != review.SourceTip)
                throw new InvalidOperationException("来源分支已更新，请重新生成合并修订稿。");
            var message = Dialogs.Ask(word.Owner, "确认合并到「" + review.Branch + "」", "合并分支 " + review.SourceBranch);
            if (message == null) return;
            WordDocuments.SaveDocument(document);
            var author = Author;
            committing = true;
            Dialogs.ShowBackgroundOperation(word.Owner, "合并提交 · " + Path.GetFileName(review.Original), async () =>
            {
                var prepared = await System.Threading.Tasks.Task.Run(() => repo.ImportPlan(path, repo.PlanFile(path)));
                var headHash = await System.Threading.Tasks.Task.Run(() => repo.SnapshotChecksum(review.Head));
                string commit = null;
                using (repo.Lock())
                    word.ReplaceWorkingCopy(review.Original, headHash, path, prepared.Checksum,
                        () => commit = repo.CommitPrepared(prepared, message, author, review.Head, review.Branch,
                            review.SourceBranch, review.SourceTip));
                review.SourceBranch = null;
                review.SourceTip = null;
                review.Save(path, commit);
                return "已合并到「" + review.Branch + "」，提交 " + commit.Substring(0, 8) + "。";
            }, () =>
            {
                committing = false;
                RefreshPane(review.Original);
            });
        }

        private void Guard(Action action, bool refreshPane = false)
        {
            if (word == null) return;
            var visiblePath = refreshPane ? VisiblePanePath() : null;
            try { action(); }
            catch (Exception error)
            {
                try { MessageBox.Show(word.Owner, error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                catch { MessageBox.Show(error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
            finally
            {
                if (refreshPane) RefreshPane(visiblePath);
            }
        }

        private string VisiblePanePath()
        {
            if (panes == null || application.Windows.Count == 0) return null;
            try { return panes.VisiblePath(application.ActiveWindow); }
            catch (COMException) { return null; }
        }

        private void RefreshPane(string previousPath)
        {
            if (panes == null || application.Windows.Count == 0) return;
            try
            {
                object window = application.ActiveWindow;
                if (panes.Refresh(window) || previousPath == null) return;
                var currentPath = (string)((dynamic)window).Document.FullName;
                if (String.Equals(previousPath, currentPath, StringComparison.OrdinalIgnoreCase)) ShowPane(window);
            }
            catch (COMException) { }
            catch (InvalidComObjectException) { }
        }

        private void Notice(string message)
        {
            MessageBox.Show(word.Owner, message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private sealed class PendingReview
        {
            public string Original, Head, Branch, SourceBranch, SourceTip;
            public static PendingReview Load(string path)
            {
                if (!path.StartsWith(Path.Combine(AppPaths.Root, "sessions") + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    || !File.Exists(path + ".wordgit")) return null;
                var saved = File.ReadAllLines(path + ".wordgit");
                if ((saved.Length != 2 && saved.Length != 5) || !Path.IsPathRooted(saved[0])) throw new InvalidDataException("合并稿的版本关联已损坏。");
                return new PendingReview { Original = saved[0], Head = saved[1], Branch = saved.Length == 2 ? "main" : saved[2],
                    SourceBranch = saved.Length == 5 && saved[3].Length > 0 ? saved[3] : null,
                    SourceTip = saved.Length == 5 && saved[4].Length > 0 ? saved[4] : null };
            }

            public void Save(string path, string head) { File.WriteAllLines(path + ".wordgit", new[] { Original, head, Branch, SourceBranch ?? "", SourceTip ?? "" }); }
        }
    }
}
