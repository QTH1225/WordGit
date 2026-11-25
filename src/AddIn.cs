using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: ComVisible(false)]
[assembly: System.Reflection.AssemblyVersion("0.11.2.0")]
[assembly: System.Reflection.AssemblyFileVersion("0.11.0.0")]
[assembly: System.Reflection.AssemblyInformationalVersion("0.11.0")]
[assembly: System.Reflection.AssemblyProduct("WordGit")]
[assembly: System.Reflection.AssemblyTitle("WordGit Word 版本管理插件")]

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
        [DispId(13)] void OnAbout([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(14)] void OnCheckUpdates([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(15)] void OnImportReview([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(16)] void OnLinkRemote([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(17)] void OnCloneRemote([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(18)] void OnFetchRemote([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(19)] void OnPullRemote([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(20)] void OnPushRemote([MarshalAs(UnmanagedType.IDispatch)] object control);
    }

    [ComVisible(true), Guid("C5707F76-F46F-4A11-800B-18E7D98377D3"), ProgId("WordGit.Connect")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IWordGitActions))]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility, IWordGitActions, ICustomTaskPaneConsumer
    {
        private dynamic application;
        private WordDocuments word;
        private TaskPanes panes;
        private readonly OperationTracker operations = new OperationTracker();
        private readonly List<string> startupNotices = new List<string>();

        // Every repository operation lives in the Rust kernel; the add-in only calls its local
        // service and drives Word's own compare and merge.
        private static ServiceClient Kernel { get { return ServiceClient.Connect(); } }

        public void OnConnection(object app, int mode, object instance, ref Array custom)
        {
            UiThread.EnsureContext();
            application = app;
            word = new WordDocuments(app);
            AppPaths.CleanSessions(1);
            StartKernel();
        }

        // Started in the background so Word's start-up is not delayed. The kernel finishes any
        // replacement a crash interrupted and hands back the messages it produced.
        private void StartKernel()
        {
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    var notices = Json.Strings(Kernel.Call("notices", ServiceClient.Args()), "notices");
                    if (notices.Length == 0) return;
                    lock (startupNotices) startupNotices.AddRange(notices);
                }
                catch (Exception error) { AppPaths.LogError("start kernel: " + error); }
            });
        }

        public void CTPFactoryAvailable(ICTPFactory factory) { panes = new TaskPanes(factory); }

        public void OnDisconnection(int mode, ref Array custom)
        {
            if (panes != null) panes.Dispose();
            panes = null;
            if (application != null) WordDocuments.Release((object)application);
            application = null;
            word = null;
        }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { ShowStartupNotices(); }

        // Word cannot be stopped from quitting here, but it waits while this call pumps messages,
        // so a commit or push that the user left running in the background is not cut off.
        public void OnBeginShutdown(ref Array custom)
        {
            UiThread.EnsureContext();
            try { operations.WaitForCompletion(); }
            catch (Exception error) { AppPaths.LogError("shutdown wait: " + error); }
        }

        private void ShowStartupNotices()
        {
            string[] pending;
            lock (startupNotices)
            {
                if (startupNotices.Count == 0) return;
                pending = startupNotices.ToArray();
                startupNotices.Clear();
            }
            MessageBox.Show(String.Join(Environment.NewLine + Environment.NewLine, pending), "WordGit",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public string GetCustomUI(string ribbonId)
        {
            return RibbonMarkup(word == null || !word.IsWpsHost);
        }

        internal static string RibbonMarkup(bool nativeIcons)
        {
            // Keep 2-3 character captions together. For longer captions, allow a line break only after the first pair.
            // WORD JOINER protects each line; ZERO WIDTH SPACE marks the intended break.
            var markup = @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
<ribbon><tabs><tab id=""WordGitTab"" label=""WordGit"">
  <group id=""WordGitVersions"" label=""本地提交"">
    <button id=""WordGitCommit"" label=""提&#x2060;交&#x200B;版&#x2060;本"" size=""large"" imageMso=""FileSave"" onAction=""OnCommit"" supertip=""保存当前文档并创建本地 Git 提交。"" />
    <button id=""WordGitHistory"" label=""提&#x2060;交&#x2060;图"" size=""large"" imageMso=""ViewOutlineView"" onAction=""OnHistory"" supertip=""在 Word 右侧查看分支和提交记录，预览或对比历史版本。"" />
    <button id=""WordGitPreview"" label=""预&#x2060;览&#x200B;版&#x2060;本"" size=""large"" imageMso=""FilePrintPreview"" onAction=""OnPreview"" supertip=""选一个历史版本，在 Word 中只读打开。"" />
  </group>
  <group id=""WordGitBranches"" label=""分支"">
    <button id=""WordGitCreateBranch"" label=""新&#x2060;建&#x200B;分&#x2060;支"" size=""large"" imageMso=""FileNew"" onAction=""OnCreateBranch"" supertip=""选择起始版本并输入新分支名称。"" />
    <button id=""WordGitSwitchBranch"" label=""切&#x2060;换&#x200B;分&#x2060;支"" size=""large"" imageMso=""WindowSwitchWindowsMenuWord"" onAction=""OnSwitchBranch"" supertip=""选择要切换到的本地分支。"" />
    <button id=""WordGitMergeBranch"" label=""合&#x2060;并&#x2060;分&#x2060;支"" size=""large"" imageMso=""ReviewAcceptChange"" onAction=""OnMergeBranch"" supertip=""将所选分支合入当前分支，先在 Word 中审阅修订。"" />
    <button id=""WordGitDeleteBranch"" label=""删&#x2060;除&#x2060;分&#x2060;支"" size=""large"" imageMso=""ReviewRejectChange"" onAction=""OnDeleteBranch"" supertip=""删除已经合入当前分支的其他本地分支。"" />
  </group>
  <group id=""WordGitReview"" label=""比较与审阅"">
    <button id=""WordGitImportReview"" label=""返&#x2060;稿&#x2060;合&#x2060;并"" size=""large"" imageMso=""FileOpen"" onAction=""OnImportReview"" supertip=""选择送审版本和返稿文件，生成 Word 修订合并稿，可选择将原始返稿保存为独立分支。"" />
    <button id=""WordGitCompare"" label=""对&#x2060;比&#x2060;版&#x2060;本"" size=""large"" imageMso=""ReviewCompareMenu"" onAction=""OnCompare"" supertip=""选一版与当前稿对比，或按 Ctrl 选两版对比。"" />
  </group>
  <group id=""WordGitRemote"" label=""远程协作"">
    <button id=""WordGitLinkRemote"" label=""关&#x2060;联&#x2060;仓&#x2060;库"" size=""large"" imageMso=""HyperlinkInsert"" onAction=""OnLinkRemote"" supertip=""为当前稿件关联一个 GitHub 仓库并获取远程历史。"" />
    <button id=""WordGitCloneRemote"" label=""克&#x2060;隆&#x2060;稿&#x2060;件"" size=""large"" imageMso=""FileOpen"" onAction=""OnCloneRemote"" supertip=""从已有 WordGit 仓库下载历史，还原为新的本地稿件。"" />
    <button id=""WordGitFetchRemote"" label=""获&#x2060;取"" size=""large"" imageMso=""Refresh"" onAction=""OnFetchRemote"" supertip=""下载远程提交和分支，不改动当前文档。"" />
    <button id=""WordGitPullRemote"" label=""拉&#x2060;取"" size=""large"" imageMso=""FileOpen"" onAction=""OnPullRemote"" supertip=""获取后更新当前稿；双方有修改时生成 Word 修订合并稿。"" />
    <button id=""WordGitPushRemote"" label=""推&#x2060;送"" size=""large"" imageMso=""FileSave"" onAction=""OnPushRemote"" supertip=""推送当前分支已经提交的版本。服务器拒绝非快进更新时请先拉取。"" />
  </group>
  <group id=""WordGitRetrieve"" label=""取回版本"">
    <button id=""WordGitExport"" label=""另&#x2060;存&#x200B;旧&#x2060;版"" size=""large"" imageMso=""FileSaveAs"" onAction=""OnExport"" supertip=""将所选版本另存为新的 Word 文档。"" />
    <button id=""WordGitRestore"" label=""恢&#x2060;复&#x2060;旧&#x2060;版"" size=""large"" imageMso=""Undo"" onAction=""OnRestore"" supertip=""将历史版本恢复为当前稿，并记录一次新提交。"" />
  </group>
  <group id=""WordGitInformation"" label=""关于与更新"">
    <button id=""WordGitAbout"" label=""WordGit {WordGitVersion}"" size=""large"" imageMso=""FileProperties"" onAction=""OnAbout"" supertip=""查看版本、作者、项目主页和反馈链接。"" />
    <button id=""WordGitCheckUpdates"" label=""检&#x2060;查&#x2060;更&#x2060;新"" size=""large"" imageMso=""Refresh"" onAction=""OnCheckUpdates"" supertip=""查询 GitHub 正式发布版本，并打开发布页下载安装。"" />
  </group>
</tab></tabs></ribbon></customUI>"
                .Replace("{WordGitVersion}", AppInfo.DisplayVersion);
            return nativeIcons ? markup : RibbonIcons.ForWps(markup);
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
        public void OnMergeReview(object control) { OnImportReview(control); }
        public void OnCreateBranch(object control) { OpenHistory(HistoryAction.CreateBranch); }
        public void OnSwitchBranch(object control) { OpenHistory(HistoryAction.SwitchBranch); }
        public void OnMergeBranch(object control) { OpenHistory(HistoryAction.MergeBranch); }
        public void OnDeleteBranch(object control) { OpenHistory(HistoryAction.DeleteBranch); }

        public void OnAbout(object control) { Guard(() => AboutDialog.ShowWindow(word.Owner, false)); }
        public void OnCheckUpdates(object control) { Guard(() => AboutDialog.ShowWindow(word.Owner, true)); }

        public void OnImportReview(object control) { OpenHistory(HistoryAction.ImportReview); }

        public void OnLinkRemote(object control) { Guard(LinkRemote); }
        public void OnCloneRemote(object control) { Guard(CloneRemote); }
        public void OnFetchRemote(object control) { Guard(() => SyncRemote(SyncOperation.Fetch)); }
        public void OnPullRemote(object control) { Guard(() => SyncRemote(SyncOperation.Pull)); }
        public void OnPushRemote(object control) { Guard(() => SyncRemote(SyncOperation.Push)); }

        private void OpenHistory(HistoryAction? command)
        {
            var changesHistory = !command.HasValue || command == HistoryAction.CreateBranch ||
                command == HistoryAction.SwitchBranch || command == HistoryAction.Restore || command == HistoryAction.DeleteBranch;
            Guard(() => { while (ManageHistory(command) && !command.HasValue) { } }, changesHistory);
        }

        private void ShowPane(object window)
        {
            var handle = (int)((dynamic)window).Hwnd;
            panes.Show(window, WordDocuments.LocalPath(DocumentForWindow(handle)), async () =>
            {
                var data = await LoadPaneAsync(DocumentForWindow(handle));
                if (!String.Equals(WordDocuments.LocalPath(DocumentForWindow(handle)), data.DocumentPath, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("此窗口的文档已变化，请刷新提交图。");
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
                    if (action == HistoryAction.ImportReview)
                    {
                        word.ActivateIfNeeded(document);
                        ManageHistory(HistoryAction.ImportReview);
                        return;
                    }
                    if (action == HistoryAction.Commit)
                    {
                        word.ActivateIfNeeded(document);
                        CommitCurrent();
                        return;
                    }
                    if (action == HistoryAction.LinkRemote || action == HistoryAction.FetchRemote || action == HistoryAction.PullRemote || action == HistoryAction.PushRemote)
                    {
                        word.ActivateIfNeeded(document);
                        if (action == HistoryAction.LinkRemote) LinkRemote();
                        else SyncRemote(action == HistoryAction.FetchRemote ? SyncOperation.Fetch : action == HistoryAction.PullRemote ? SyncOperation.Pull : SyncOperation.Push);
                        return;
                    }
                    if (action == HistoryAction.CompareLatest)
                    {
                        var head = Kernel.Status(path).Head;
                        if (head == null) { Notice("尚无提交版本可对比。"); return; }
                        CompareVersions(document, path, head, null);
                    }
                    else if (action == HistoryAction.Compare) CompareVersions(document, path, revision, other);
                    else using (Busy()) PreviewVersion(document, path, revision);
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

        private static async Task<HistoryPaneData> LoadPaneAsync(object document, string storage = null)
        {
            // Capture Word COM values on its UI thread; only the service round trip runs in the worker.
            var path = WordDocuments.LocalPath(document);
            var readOnly = (bool)((dynamic)document).ReadOnly;
            var client = ServiceClient.Connect();
            var result = await Task.Run(() =>
            {
                var status = client.Status(path);
                var review = status.Review;
                var workingCopy = new WorkingCopy(path, () => true, () => status.WorkingState);
                return new HistoryPaneData
                {
                    DocumentPath = path, HistoryPath = HistoryPath(path, review), Branch = status.Branch,
                    Head = status.Head, Branches = status.Branches, InitialHistory = status.History,
                    MergeTargetBranch = review != null && review.IsPendingMerge ? review.Branch : null,
                    Load = name => Revision.List(client.Call("history", ServiceClient.Args("document", path, "branch", name)), "history"),
                    WorkingState = workingCopy.State, Remote = status.Remote
                };
            });
            result.IsSaved = () => (bool)((dynamic)document).Saved;
            var pendingMerge = result.MergeTargetBranch != null;
            result.LoadReviewItems = () => WordDocuments.ReviewItems(document, !pendingMerge);
            result.ReviewSignature = () => WordDocuments.ReviewSignature(document);
            result.CanResolveRevisions = !readOnly;
            if (pendingMerge)
                result.DocumentRevisionCount = () => WordDocuments.DocumentRevisionCount(document);
            return result;
        }

        // The document whose history an action uses: a pending merge copy maps back to its
        // original. Other files under WordGit's sessions folder are previews and have no history.
        private static string HistoryPath(string documentPath, PendingReview review)
        {
            if (review != null) return review.Original;
            if (AppPaths.IsSessionFile(documentPath))
                throw new InvalidOperationException("这是 WordGit 生成的临时预览或比较稿，没有自己的版本历史。如需保留，请先「另存为」到其他位置。");
            return documentPath;
        }

        private void PreviewVersion(object document, string path, string revision)
        {
            var file = Json.String(Kernel.Call("version/preview", ServiceClient.Args(
                "document", path, "revision", revision, "author", Author)), "file");
            if (file == null) throw new InvalidOperationException("没能准备这个版本，请稍后重试。");
            word.Open(file, true);
        }

        private void CompareVersions(object document, string path, string basis, string other)
        {
            Dictionary<string, object> request;
            if (other == null) request = ServiceClient.Args("document", path, "revision", basis, "author", Author);
            else request = ServiceClient.Args("document", path, "revision", basis, "other", other, "author", Author);
            var plan = ComparePlan.From(Kernel.Call("version/compare", request));

            if (word.IsWpsHost)
            {
                using (Busy())
                {
                    if (other == null)
                    {
                        var snapshot = WordDocuments.CaptureComparisonSource(document, AppPaths.NewSession());
                        try { word.CompareSnapshot(plan.Original, snapshot); }
                        finally { DiscardReviewOutput(snapshot); }
                    }
                    else word.ReviewPaths(plan.Original, plan.Revised, false);
                }
                return;
            }

            EnsureIdle(path);
            var resultPath = Path.Combine(AppPaths.NewSession(), Path.GetFileNameWithoutExtension(path) + "-版本对比.docx");
            var current = other == null ? WordDocuments.CaptureComparisonSource(document, AppPaths.NewSession()) : null;
            try
            {
                RunInBackground(path, "对比版本", async token =>
                {
                    try
                    {
                        await OpenBackgroundReview(resultPath, worker => worker.BuildVersionReviewFile(
                            plan.Original, plan.Revised, current, false, resultPath), null, null, false);
                        return "已打开版本对比结果。";
                    }
                    finally { if (current != null) DiscardReviewOutput(current); }
                }, false, initialStatus: "正在生成 Word 版本对比…");
            }
            catch
            {
                if (current != null) DiscardReviewOutput(current);
                throw;
            }
        }

        private string Author { get { return (string)application.UserName; } }

        private static string Key(string historyPath) { return Path.GetFullPath(historyPath); }

        // Runs an operation in the modeless progress dialog and marks the document busy until
        // it ends; the busy mark is what OnBeginShutdown waits for.
        private void RunInBackground(string historyPath, string title, Func<CancellationToken, Task<string>> operation,
            bool cancellable, Action finished = null, string initialStatus = null)
        {
            var key = Key(historyPath);
            operations.Begin(key, title);
            Action done = () =>
            {
                operations.End(key);
                if (finished != null) finished();
            };
            try
            {
                if (cancellable) Dialogs.ShowBackgroundOperation(word.Owner, title, operation, done, initialStatus);
                else Dialogs.ShowBackgroundOperation(word.Owner, title, () => operation(CancellationToken.None), done, initialStatus);
            }
            catch
            {
                operations.End(key);
                throw;
            }
        }

        private void EnsureIdle(string historyPath)
        {
            if (operations.IsBusy(Key(historyPath)))
                throw new InvalidOperationException("这份文档正在后台" + operations.Describe() + "，请等待当前操作完成。");
        }

        // Long kernel operations are cancellable with an id the add-in can signal from the
        // progress dialog; the token the dialog cancels is forwarded to the service.
        private static string NewOperation() { return Guid.NewGuid().ToString("N"); }

        private static void OnCancel(CancellationToken token, string operation)
        {
            if (token.CanBeCanceled)
                token.Register(() =>
                {
                    try { ServiceClient.Connect().CancelOperation(operation); }
                    catch (Exception error) { AppPaths.LogError("cancel " + operation + ": " + error); }
                });
        }

        private void CommitCurrent()
        {
            var document = word.Current();
            var path = WordDocuments.LocalPath(document, true);
            var client = Kernel;
            var status = client.Status(path);
            var review = status.Review;
            var historyPath = HistoryPath(path, review);
            EnsureIdle(historyPath);
            var branch = status.Branch;
            var head = status.Head;
            if (review != null && (head != review.Head || branch != review.Branch))
                throw new InvalidOperationException("审阅期间原文档的历史已更新。请先比较最新版本，再重新合并。");
            var finishingMerge = review != null && review.IsPendingMerge;
            if (finishingMerge) CheckReviewComplete(document);
            if (review != null && (review.SourceBranch != null || review.ApplyToOriginal))
            {
                FinishWorkingCopyMerge(document, path, review);
                return;
            }
            var target = Path.GetFileName(historyPath);
            var message = Dialogs.Ask(word.Owner, (finishingMerge ? "完成合并到「" : "提交到「") + target + "」 · " + branch,
                finishingMerge ? "合并审阅稿" : "");
            if (message == null) return;
            WordDocuments.SaveDocument(document);
            var author = Author;
            RunInBackground(historyPath, "提交版本 · " + target, async token =>
            {
                var result = await Task.Run(() => client.Call("commit", ServiceClient.Args(
                    "document", path, "message", message, "author", author, "expectedHead", head)));
                var id = Json.String(result, "commit");
                return id == null ? "文件内容与最新提交一致。" : "已提交 " + id.Substring(0, 8) +
                    "。提交时保存的版本已入库，之后的编辑可继续提交。";
            }, false, () => {
                if (finishingMerge) RefreshMergePanes(historyPath, path);
                else RefreshPane(path);
            });
        }

        private bool ManageHistory(HistoryAction? command)
        {
            var document = word.Current();
            var activePath = WordDocuments.LocalPath(document);
            var client = Kernel;
            var status = client.Status(activePath);
            var review = status.Review;
            var historyPath = HistoryPath(activePath, review);
            var branch = status.Branch;
            var activeTip = status.Head;
            if (activeTip == null) { Notice("还没有版本，请先点击「提交版本」。"); return false; }
            var branches = status.Branches ?? new string[0];
            var filename = Path.GetFileName(historyPath);
            if (command == HistoryAction.MergeBranch || command == HistoryAction.Merge) word.EnsureTrackedMergeSupported();
            var deleting = command == HistoryAction.DeleteBranch;
            var chooseBranch = command == HistoryAction.SwitchBranch || command == HistoryAction.MergeBranch || deleting;
            if (review != null && (chooseBranch || command == HistoryAction.CreateBranch || command == HistoryAction.Restore || command == HistoryAction.ImportReview))
                throw new InvalidOperationException("请在原文档中执行此操作。");

            if (command == HistoryAction.ImportReview)
            {
                EnsureIdle(historyPath);
                MergeReturnedDraft(document, activePath, branch, activeTip, branches);
                return false;
            }

            HistoryAction action;
            Revision basis = null, other = null;
            string selectedBranch;
            if (chooseBranch)
            {
                var alternatives = (deleting ? DeletableBranches(client, activePath, branch) : branches)
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
                using (var dialog = new HistoryDialog(filename, status.History, review == null, activeTip, branch, branches,
                    name => Revision.List(client.Call("history", ServiceClient.Args("document", activePath, "branch", name)), "history"), command))
                {
                    if (dialog.ShowDialog(word.Owner) != DialogResult.OK) return false;
                    action = dialog.Action;
                    basis = dialog.Basis;
                    other = dialog.Other;
                    selectedBranch = dialog.SelectedBranch;
                }
            }
            UiThread.EnsureContext();
            var readOnlyAction = action == HistoryAction.Preview || action == HistoryAction.Compare || action == HistoryAction.Export;
            if (!readOnlyAction) EnsureIdle(historyPath);
            switch (action)
            {
                case HistoryAction.ImportReview:
                    MergeReturnedDraft(document, activePath, branch, activeTip, branches);
                    break;
                case HistoryAction.CreateBranch:
                    var name = Dialogs.Ask(word.Owner, "从 " + basis.Id.Substring(0, 8) + " 新建分支", "", true);
                    if (name == null) return true;
                    using (Busy())
                        Notice(Json.String(client.Call("branch/create", ServiceClient.Args(
                            "document", activePath, "name", name, "revision", basis.Id, "author", Author)), "message"));
                    return true;
                case HistoryAction.SwitchBranch:
                    SwitchBranch(document, activePath, branch, activeTip, selectedBranch);
                    break;
                case HistoryAction.MergeBranch:
                    CheckMergeTarget(activePath);
                    RunInBackground(historyPath, "合并分支", token => GenerateBranchMerge(document, activePath,
                        selectedBranch, null), false, initialStatus: "正在生成 Word 修订合并稿…");
                    break;
                case HistoryAction.DeleteBranch:
                    DeleteBranch(client, activePath, branch, selectedBranch);
                    break;
                case HistoryAction.Preview:
                    using (Busy()) PreviewVersion(document, activePath, basis.Id);
                    break;
                case HistoryAction.Compare:
                    CompareVersions(document, activePath, basis.Id, other == null ? null : other.Id);
                    break;
                case HistoryAction.Export:
                    using (Busy())
                    {
                        var exported = client.Call("version/export", ServiceClient.Args(
                            "document", activePath, "revision", basis.Id, "author", Author));
                        var file = Json.String(exported, "file");
                        if (file != null) word.Open(file, false);
                        Notice(Json.String(exported, "message") ?? "已另存为新的 Word 文档。");
                    }
                    break;
                case HistoryAction.Restore:
                    RestoreVersion(document, activePath, basis.Id);
                    break;
                case HistoryAction.Merge:
                    word.EnsureTrackedMergeSupported();
                    var versions = MergePlan.From(client.Call("merge/versions", ServiceClient.Args(
                        "document", activePath, "revision", basis.Id, "other", other.Id, "author", Author)));
                    RunInBackground(historyPath, "合并两个历史版本", async token =>
                    {
                        await OpenBackgroundReview(versions.Draft, worker => worker.BuildVersionReviewFile(
                            versions.Original, versions.Revised, null, true, versions.Draft), versions.Review);
                        return versions.Notice ?? "已打开合并稿。处理修订后点击右侧的「完成合并」。";
                    }, false, initialStatus: "正在生成 Word 修订合并稿…");
                    break;
            }
            return false;
        }

        private static string[] DeletableBranches(ServiceClient client, string path, string current)
        {
            var result = client.Call("branch/deletable", ServiceClient.Args("document", path));
            return Json.Strings(result, "branches");
        }

        private void DeleteBranch(ServiceClient client, string path, string current, string target)
        {
            var status = client.Status(path);
            foreach (var item in status.History)
            {
                if (!item.References.Any(reference => ReferenceText.Label(reference) == target)) continue;
                if (MessageBox.Show(word.Owner, "删除已合入「" + current + "」的分支「" + target + "」？\n最后一次提交  " +
                    item.Id.Substring(0, 8) + "。", "WordGit · 删除分支", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                break;
            }
            using (Busy())
                Notice(Json.String(client.Call("branch/delete", ServiceClient.Args(
                    "document", path, "name", target, "author", Author)), "message"));
        }

        private void RestoreVersion(object document, string path, string revision)
        {
            if (MessageBox.Show(word.Owner, "先存下当前稿，再将所选旧版恢复到原文件，并创建一次新提交。是否继续？", "WordGit",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            WordDocuments.EnsureRestorable(document);
            WordDocuments.SaveDocument(document);
            var author = Author;
            var client = Kernel;
            RunInBackground(path, "恢复旧版 " + revision.Substring(0, 8), async token =>
            {
                var prepared = await Task.Run(() => client.Call("version/restore", ServiceClient.Args(
                    "document", path, "revision", revision, "author", author)));
                if (Json.Bool(prepared, "unchanged")) return Json.String(prepared, "message") ?? "当前稿已经与所选旧版一致。";
                var plan = Json.String(prepared, "plan");
                if (plan == null) throw new InvalidOperationException("没能准备这个版本，请稍后重试。");
                return word.ApplyReplacement(Json.String(prepared, "document") ?? path, plan, true);
            }, false, () => RefreshPane(path), ReplacingStatus);
        }

        private const string ReplacingStatus = "正在后台准备版本，完成后当前稿会重新打开。请暂时不要编辑当前稿。";

        private void SwitchBranch(object document, string path, string from, string expected, string target)
        {
            WordDocuments.EnsureReplaceable(path);
            string message = null;
            if (!(bool)((dynamic)document).Saved || Kernel.Status(path).WorkingState == "changed")
            {
                var answer = MessageBox.Show(word.Owner, "当前稿有未提交的修改。是否先提交到「" + from + "」再切换？",
                    "WordGit", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer != DialogResult.Yes) return;
                message = Dialogs.Ask(word.Owner, "切换前提交到「" + from + "」", "切换前保存当前稿");
                if (message == null) return;
                WordDocuments.SaveDocument(document);
            }
            var author = Author;
            var client = Kernel;
            RunInBackground(path, "切换到「" + target + "」", async token =>
            {
                Dictionary<string, object> request;
                if (message == null) request = ServiceClient.Args("document", path, "name", target, "author", author);
                else request = ServiceClient.Args("document", path, "name", target, "message", message, "author", author);
                var prepared = await Task.Run(() => client.Call("branch/switch", request));
                var plan = Json.String(prepared, "plan");
                if (plan == null) throw new InvalidOperationException("没能准备目标分支，请稍后重试。");
                return word.ApplyReplacement(Json.String(prepared, "document") ?? path, plan, true);
            }, false, () => RefreshPane(path), ReplacingStatus);
        }

        // A merge draft only makes sense on top of committed work; the kernel enforces the same
        // rule, this check fails before the user waits for anything.
        private static void CheckMergeTarget(string path)
        {
            WordDocuments.EnsureReplaceable(path);
            if (ServiceClient.Connect().Status(path).WorkingState == "changed")
                throw new InvalidOperationException("请先提交当前分支的修改，再生成合并稿。");
        }

        private async Task<string> GenerateBranchMerge(object document, string path, string source, string submitted)
        {
            CheckMergeTarget(path);
            word.EnsureTrackedMergeSupported();
            var client = Kernel;
            Dictionary<string, object> request;
            if (submitted == null) request = ServiceClient.Args("document", path, "source", source);
            else request = ServiceClient.Args("document", path, "source", source, "basis", submitted);
            var plan = MergePlan.From(await Task.Run(() => client.Call("merge/context", request)));
            if (plan.UpToDate) return plan.Message ?? "来源分支的提交已包含在当前分支。";
            if (plan.Draft == null) throw new InvalidOperationException("没能准备合并稿，请稍后重试。");
            await OpenBackgroundReview(plan.Draft, worker => worker.BuildBranchMergeFile(plan, plan.Draft), plan.Review);
            return plan.Notice ?? "已打开合并稿。处理修订后点击右侧的「完成合并」。";
        }

        // The draft is produced by a second, hidden Word instance so the document the user is
        // reading keeps its focus and scroll position.
        private async Task OpenBackgroundReview(string resultPath, Action<WordDocuments> build,
            PendingReview review = null, Action validate = null, bool showPane = true)
        {
            var wordExe = word == null ? null : Path.Combine((string)application.Path, "WINWORD.EXE");
            var opened = false;
            try
            {
                await QuietWord.Run(wordExe, build);
                if (validate != null) validate();
                if (review != null) await Task.Run(() => ServiceClient.Connect().Call("merge/register", ServiceClient.Args(
                    "draft", resultPath, "review", ReviewPayload(review))));
                var draft = word.Open(resultPath, false);
                opened = true;
                if (review != null && showPane) ShowMergePane(draft);
            }
            finally
            {
                if (!opened) DiscardReviewOutput(resultPath);
            }
        }

        private static Dictionary<string, object> ReviewPayload(PendingReview review)
        {
            return new Dictionary<string, object>
            {
                { "original", review.Original }, { "head", review.Head }, { "branch", review.Branch },
                { "sourceBranch", review.SourceBranch }, { "sourceTip", review.SourceTip },
                { "applyToOriginal", review.ApplyToOriginal }, { "recordOnlyMerge", review.RecordOnlyMerge },
            };
        }

        private static void DiscardReviewOutput(string path)
        {
            foreach (var file in new[] { path, path + ".wordgit" })
                try { File.Delete(file); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            try { Directory.Delete(Path.GetDirectoryName(path)); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private void ShowMergePane(object result)
        {
            if (panes == null) return;
            try { ShowPane((object)((dynamic)result).ActiveWindow); }
            catch (Exception error) { AppPaths.LogError("open merge review pane: " + error); }
        }

        private void MergeReturnedDraft(object document, string path, string branch, string tip, string[] branches)
        {
            ImportReviewOptions options;
            bool chooseHistoricalMerge;
            var history = Kernel.Status(path).History;
            using (var picker = new ImportReviewDialog(path, history, branch, branches, word.SupportsTrackedMerge))
            {
                if (picker.ShowDialog(word.Owner) != DialogResult.OK) return;
                chooseHistoricalMerge = picker.ChooseHistoricalMerge;
                options = picker.Options;
            }
            if (chooseHistoricalMerge)
            {
                ManageHistory(HistoryAction.Merge);
                return;
            }

            var basis = options.Basis;
            string message = null;
            string author = null;
            if (options.CreateMergeDraft)
            {
                WordDocuments.EnsureReplaceable(path);
                if (Kernel.Status(path).WorkingState == "changed")
                {
                    message = Dialogs.Ask(word.Owner, "提交当前稿并继续返稿合并", "");
                    if (message == null) return;
                    WordDocuments.SaveDocument(document);
                    author = Author;
                }
            }
            UiThread.EnsureContext();
            var client = Kernel;
            RunInBackground(path, options.CreateMergeDraft ? "生成返稿合并稿" : "保存返稿", async token =>
            {
                if (message != null)
                    await Task.Run(() => client.Call("commit", ServiceClient.Args(
                        "document", path, "message", message, "author", author, "expectedHead", tip)));
                Dictionary<string, object> request;
                if (options.Branch == null) request = ServiceClient.Args("document", path, "file", options.FilePath,
                    "basis", basis.Id, "reviewer", options.Reviewer ?? "返稿");
                else request = ServiceClient.Args("document", path, "file", options.FilePath, "basis", basis.Id,
                    "branch", options.Branch, "reviewer", options.Reviewer ?? "返稿");
                var imported = await Task.Run(() => client.Call("review/import", request));
                var saved = Json.String(imported, "message") ?? "已保存返稿。";
                if (!options.CreateMergeDraft) return saved + " 可在提交图中预览或对比返稿。";
                try
                {
                    if (options.Branch != null)
                        return saved + " " + await GenerateBranchMerge(document, path, options.Branch, null);
                    var plan = MergePlan.From(imported);
                    if (plan.Draft == null) throw new InvalidOperationException("没能准备返稿合并稿，请稍后重试。");
                    await OpenBackgroundReview(plan.Draft, worker => worker.BuildBranchMergeFile(plan, plan.Draft), plan.Review);
                    return saved + " " + (plan.Notice ?? "");
                }
                catch (Exception error)
                {
                    AppPaths.LogError("import review merge: " + error);
                    return saved + Environment.NewLine + "生成合并稿未完成：" + error.Message +
                        Environment.NewLine + "再次选择同一返稿和送审版本时会继续使用这份记录。";
                }
            }, false, () => RefreshPane(path), options.CreateMergeDraft ? "正在准备返稿并生成 Word 修订合并稿…" : "正在保存返稿…");
        }

        private void LinkRemote()
        {
            var document = word.Current();
            var path = WordDocuments.LocalPath(document);
            var status = Kernel.Status(path);
            if (status.Review != null) throw new InvalidOperationException("请在原稿中关联远程仓库。");
            var historyPath = HistoryPath(path, null);
            EnsureIdle(historyPath);
            var state = status.Remote;
            RemoteOptions options;
            using (var picker = new RemoteDialog(false, state == null ? null : state.Url, state == null ? null : state.RemoteBranch))
            {
                if (picker.ShowDialog(word.Owner) != DialogResult.OK) return;
                options = picker.Options;
            }
            UiThread.EnsureContext();
            var operation = NewOperation();
            RunInBackground(historyPath, "关联远程仓库", async token => {
                OnCancel(token, operation);
                var result = await Task.Run(() => Kernel.Call("remote/link", ServiceClient.Args(
                    "document", path, "url", options.Url, "branch", options.Branch, "operation", operation)));
                return Json.String(result, "message") ?? "仓库已关联。";
            }, true, () => RefreshPane(historyPath));
        }

        private void CloneRemote()
        {
            RemoteOptions options;
            using (var picker = new RemoteDialog(true))
            {
                if (picker.ShowDialog(word.Owner) != DialogResult.OK) return;
                options = picker.Options;
            }
            UiThread.EnsureContext();
            WordDocuments.EnsureReplaceable(options.DocumentPath);
            var operation = NewOperation();
            RunInBackground(options.DocumentPath, "克隆远程稿件", async token => {
                OnCancel(token, operation);
                var result = await Task.Run(() => Kernel.Call("remote/clone", ServiceClient.Args(
                    "url", options.Url, "branch", options.Branch, "document", options.DocumentPath, "operation", operation)));
                var document = Json.String(result, "document") ?? options.DocumentPath;
                word.Open(document, false);
                try { if (panes != null) ShowPane(application.ActiveWindow); }
                catch (Exception error) { return (Json.String(result, "message") ?? "远程历史已下载。") + Environment.NewLine + "侧栏未能打开：" + error.Message + Environment.NewLine + "可点击「提交图」重试。"; }
                return Json.String(result, "message") ?? "远程历史已下载。";
            }, true);
        }

        private enum SyncOperation { Fetch, Pull, Push }

        private void SyncRemote(SyncOperation operation)
        {
            var document = word.Current();
            var path = WordDocuments.LocalPath(document);
            var status = Kernel.Status(path);
            if (status.Review != null) throw new InvalidOperationException("请在原稿中获取、拉取或推送；先完成或关闭当前合并审阅稿。");
            var historyPath = HistoryPath(path, null);
            EnsureIdle(historyPath);
            if (status.Remote == null || status.Remote.Url == null) throw new InvalidOperationException("请先点击「关联仓库」。");
            var branch = status.Branch;
            var head = status.Head;
            if (operation != SyncOperation.Fetch && head == null) throw new InvalidOperationException("请先提交当前稿，或使用「克隆稿件」。");
            var title = operation == SyncOperation.Push ? "推送当前分支" : operation == SyncOperation.Pull ? "拉取远程版本" : "获取远程历史";
            var id = NewOperation();
            var client = Kernel;
            RunInBackground(historyPath, title, async token => {
                OnCancel(token, id);
                var kind = operation == SyncOperation.Push ? "remote/push" : operation == SyncOperation.Pull ? "remote/pull" : "remote/fetch";
                var result = await Task.Run(() => client.Call(kind, ServiceClient.Args("document", path, "operation", id)));
                RefreshPane(path);
                if (operation == SyncOperation.Push || operation == SyncOperation.Fetch)
                    return Json.String(result, "message") ?? "已获取远程历史。";
                var action = Json.String(result, "action") ?? "none";
                if (action == "none") return Json.String(result, "message") ?? "当前没有需要拉取的新版本。";
                if (action == "commitFirst") return Json.String(result, "message") ?? "请先提交当前稿，再点击「拉取」。";
                if (action == "merge")
                    return await GenerateBranchMerge(document, path, Json.String(result, "source"), null);
                var plan = Json.String(result, "plan");
                if (plan == null) return Json.String(result, "message") ?? "远程历史已获取。";
                return word.ApplyReplacement(Json.String(result, "document") ?? path, plan, true);
            }, true);
        }

        private static void CheckReviewComplete(object document)
        {
            if (WordDocuments.DocumentRevisionCount(document) > 0 || WordDocuments.ReviewItems(document).Any(item => !item.IsComment))
                throw new InvalidOperationException("合并稿仍有待处理的修订，请先逐条接受或拒绝。");
        }

        private void FinishWorkingCopyMerge(object document, string path, PendingReview review)
        {
            var fromBranch = review.SourceBranch != null && !review.SourceBranch.StartsWith(ReferenceText.ReviewArchivePrefix, StringComparison.Ordinal);
            var summary = fromBranch ? "合并分支 " + ReferenceText.Label(review.SourceBranch) : "合并返稿";
            var message = Dialogs.Ask(word.Owner, "确认合并到「" + review.Branch + "」", summary);
            if (message == null) return;
            WordDocuments.SaveDocument(document);
            var author = Author;
            var original = review.Original;
            var client = Kernel;
            RunInBackground(original, "合并提交 · " + Path.GetFileName(original), async token =>
            {
                var prepared = await Task.Run(() => client.Call("merge/finish", ServiceClient.Args(
                    "draft", path, "message", message, "author", author)));
                var plan = Json.String(prepared, "plan");
                if (plan == null) throw new InvalidOperationException("没能准备合并结果，请稍后重试。");
                return word.ApplyReplacement(Json.String(prepared, "document") ?? original, plan, true);
            }, false, () => RefreshMergePanes(original, path));
        }

        // Every Ribbon and pane callback enters here: errors become a message box plus a log entry,
        // never an unhandled exception in Word's message loop.
        private void Guard(Action action, bool refreshPane = false)
        {
            if (word == null) return;
            UiThread.EnsureContext();
            ShowStartupNotices();
            string visiblePath = null;
            try
            {
                if (refreshPane) visiblePath = VisiblePanePath();
                action();
            }
            catch (Exception error)
            {
                AppPaths.LogError(error.ToString());
                try { MessageBox.Show(word.Owner, error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                catch (Exception) { MessageBox.Show(error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
            finally
            {
                if (refreshPane) RefreshPane(visiblePath);
            }
        }

        private static IDisposable Busy() { return new WaitCursor(); }

        private sealed class WaitCursor : IDisposable
        {
            private readonly Cursor previous = Cursor.Current;
            public WaitCursor() { Cursor.Current = Cursors.WaitCursor; }
            public void Dispose() { Cursor.Current = previous; }
        }

        private string VisiblePanePath()
        {
            if (panes == null || application.Windows.Count == 0) return null;
            try { return panes.VisiblePath(application.ActiveWindow); }
            catch (COMException) { return null; }
        }

        private void RefreshMergePanes(string originalPath, string draftPath)
        {
            if (panes != null && !String.Equals(VisiblePanePath(), draftPath, StringComparison.OrdinalIgnoreCase))
                panes.RefreshDocument(draftPath);
            RefreshPane(originalPath);
        }

        private void RefreshPane(string previousPath)
        {
            if (panes == null || application == null) return;
            try
            {
                if (application.Windows.Count == 0) return;
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
            if (message == null) return;
            MessageBox.Show(word.Owner, message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
