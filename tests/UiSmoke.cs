using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WordGit
{
    internal static class UiSmoke
    {
        [STAThread]
        public static void Main(string[] args)
        {
            try
            {
                Application.EnableVisualStyles();
                // The pane talks to the core; keep it away from real history.
                if (String.IsNullOrEmpty(Environment.GetEnvironmentVariable("WORDGIT_HOME")))
                    Environment.SetEnvironmentVariable("WORDGIT_HOME", Path.Combine(Path.GetTempPath(), "wordgit-ui-" + Guid.NewGuid().ToString("N")));
                var day = DateTimeOffset.Now;
                var one = new string('a', 40);
                var two = new string('b', 40);
                var three = new string('c', 40);
                var four = new string('d', 40);
                var items = new List<Revision> {
                    RevisionFor(one, new[] { two, three }, day, "合并导师审阅稿", "HEAD -> main"),
                    RevisionFor(two, new[] { four }, day.AddMinutes(-5), "调整页眉并补充试验条件"),
                    RevisionFor(three, new[] { four }, day.AddMinutes(-8), "导师审阅版", "review"),
                    RevisionFor(four, new string[0], day.AddHours(-1), "初稿", "v0.1")
                };
                VerifyCommitDialog(args[0]);
                VerifyDialogPlacement();
                VerifyDataDirectoryDialog(args[0]);
                using (var history = new HistoryDialog("论文独创性声明.docx", items, true, one, "main",
                    new[] { "main", "review" }, branch => branch == "review" ? new List<Revision> { items[2], items[3] } : items))
                {
                    Render(history, Path.Combine(args[0], "history-graph.png"));
                    history.Size = history.MinimumSize;
                    Render(history, Path.Combine(args[0], "history-graph-min.png"));
                    var grid = Descendants(history).OfType<DataGridView>().Single();
                    grid.Rows[2].Selected = true;
                    var buttons = Descendants(history).OfType<Button>().ToArray();
                    if (buttons.Single(button => button.Text == "打开预览").Enabled || !buttons.Single(button => button.Text == "对比版本").Enabled)
                        throw new InvalidOperationException("Invalid actions for multi-selection");
                    var filter = Descendants(history).OfType<ComboBox>().Single();
                    filter.SelectedItem = "review";
                    if (grid.RowCount != 2 || !buttons.Single(button => button.Text == "切换分支").Enabled ||
                        !buttons.Single(button => button.Text == "合并分支…").Enabled)
                        throw new InvalidOperationException("Invalid filtered branch state");
                    Render(history, Path.Combine(args[0], "history-branch-min.png"));
                }
                VerifyRibbon();
                VerifyRemoteUi(args[0], items);
                VerifyImportReview(args[0], items);
                VerifyAboutDialog(args[0]);
                VerifyUpdateDialog(args[0]);
                VerifyBackgroundOperationCancel();
                VerifyHistoryPane(args[0], items, one, three);
                VerifyRevisionActions(args[0]);
                VerifyAsyncHistoryPane(args[0], items);
                VerifyPaneReuse(items);
                VerifyMergeDraftPane(args[0], one, three);
                using (var compare = new HistoryDialog("论文独创性声明.docx", items, true, one, command: HistoryAction.Compare))
                {
                    compare.Size = compare.MinimumSize;
                    Render(compare, Path.Combine(args[0], "compare-picker.png"));
                    var grid = Descendants(compare).OfType<DataGridView>().Single();
                    var button = (Button)compare.AcceptButton;
                    if (!grid.MultiSelect || button.Text != "对比版本" || Descendants(compare).OfType<Button>().Count() != 2)
                        throw new InvalidOperationException("Compare entry should show only its own action and Cancel");
                    compare.Show();
                    grid.Rows[2].Selected = true;
                    button.PerformClick();
                    if (compare.DialogResult != DialogResult.OK || compare.Action != HistoryAction.Compare ||
                        compare.Basis.Id != three || compare.Other.Id != one)
                        throw new InvalidOperationException("Compare entry did not return the two chosen revisions");
                }
                using (var create = new HistoryDialog("论文独创性声明.docx", items, true, three, command: HistoryAction.CreateBranch))
                {
                    create.Show();
                    var grid = Descendants(create).OfType<DataGridView>().Single();
                    if (grid.MultiSelect || grid.SelectedRows.Count != 1 || grid.SelectedRows[0].Index != 2)
                        throw new InvalidOperationException("Create entry should preselect the current tip");
                    ((Button)create.AcceptButton).PerformClick();
                    if (create.DialogResult != DialogResult.OK || create.Action != HistoryAction.CreateBranch || create.Basis.Id != three)
                        throw new InvalidOperationException("Create entry did not return the selected revision");
                }
                ComboBox branchChoice;
                using (var picker = Dialogs.BranchForm("论文独创性声明.docx", "main", new[] { "review", "draft" }, HistoryAction.MergeBranch, out branchChoice))
                {
                    picker.Size = picker.MinimumSize;
                    Render(picker, Path.Combine(args[0], "branch-picker.png"));
                    if ((string)branchChoice.SelectedItem != "review" || ((Button)picker.AcceptButton).Text != "生成修订稿")
                        throw new InvalidOperationException("Invalid branch merge picker");
                }
                Console.WriteLine("PASS: Ribbon XML and callbacks, focused version and branch pickers, history pane, graph selection, default and minimum-size layouts.");
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error.GetType().Name + ": " + error.Message);
                Environment.ExitCode = 1;
            }
        }

        private static void VerifyCommitDialog(string output)
        {
            TextBox subject, body;
            using (var form = Dialogs.CommitForm("提交到「论文独创性声明.docx」 · main", "", out subject, out body))
            {
                var submit = Descendants(form).OfType<Button>().Single(button => button.Text == "提交");
                if (submit.Enabled || subject.Multiline || !body.Multiline || !body.AcceptsReturn)
                    throw new InvalidOperationException("Commit title/body inputs are invalid");
                body.Text = "补充试验条件。";
                if (submit.Enabled) throw new InvalidOperationException("A commit without a title was accepted");
                subject.Text = "校正封面";
                body.Text = "补充 Word 版本管理记录。\r\n\r\n第二段保留换行。";
                if (!submit.Enabled || Dialogs.ComposeMessage(subject.Text, body.Text) !=
                    "校正封面\n\n补充 Word 版本管理记录。\n\n第二段保留换行。" ||
                    Dialogs.ComposeMessage("  标题  ", "  ") != "标题")
                    throw new InvalidOperationException("Commit title/body message formatting changed");
                Render(form, Path.Combine(output, "commit-dialog.png"));
                VerifyNoOverlap(form);
                if (body.Height < 180) throw new InvalidOperationException("Commit details area is too small");
                form.Size = form.MinimumSize;
                Render(form, Path.Combine(output, "commit-dialog-min.png"));
                VerifyNoOverlap(form);
            }
            using (var form = Dialogs.CommitForm("确认合并", "旧标题\n\n第一段\n第二段", out subject, out body))
            {
                if (subject.Text != "旧标题" || body.Text.Replace("\r\n", "\n") != "第一段\n第二段")
                    throw new InvalidOperationException("Existing commit text was not split correctly");
                form.Font = new Font(form.Font.FontFamily, 13.5f);
                Render(form, Path.Combine(output, "commit-dialog-large-font.png"));
                VerifyNoOverlap(form);
            }
            using (var branch = Dialogs.CommitForm("从旧版新建分支", "review/老师", out subject, out body, true))
            {
                if (body != null || subject.Multiline || branch.AcceptButton == null)
                    throw new InvalidOperationException("The branch-name dialog acquired a commit body field");
                Render(branch, Path.Combine(output, "create-branch-input.png"));
            }
        }

        private static void VerifyDataDirectoryDialog(string output)
        {
            const string path = @"F:\试验文档\WordGit 提交历史";
            using (var dialog = new DataDirectoryDialog(path, true))
            {
                Render(dialog, Path.Combine(output, "data-directory-dialog.png"));
                VerifyNoOverlap(dialog);
                var directory = Descendants(dialog).OfType<TextBox>().Single();
                var browse = Descendants(dialog).OfType<Button>().Single(button => button.AccessibleName == "选择文件夹");
                var move = Descendants(dialog).OfType<Button>().Single(button => button.Text == "迁移历史");
                if (directory.Text != path || directory.ReadOnly || directory.BackColor != SystemColors.Window ||
                    directory.ClientSize.Height < directory.Font.Height || !browse.Enabled || browse.Left <= directory.Right ||
                    move.Enabled || dialog.SelectedDirectory != null ||
                    ((Button)dialog.CancelButton).DialogResult != DialogResult.Cancel)
                    throw new InvalidOperationException("The data directory should be editable with a folder picker on the right");
                directory.Text = path + "\\";
                if (move.Enabled) throw new InvalidOperationException("A trailing separator should not make the directory a new destination");
                const string target = @"F:\试验文档\新历史";
                directory.Text = target;
                if (!move.Enabled || dialog.SelectedDirectory != null)
                    throw new InvalidOperationException("Editing the location should wait for an explicit migration");
                dialog.Show();
                move.PerformClick();
                if (dialog.DialogResult != DialogResult.OK || dialog.SelectedDirectory != target)
                    throw new InvalidOperationException("The migration should use the edited directory");
            }
            using (var dialog = new DataDirectoryDialog(path, false))
            {
                var directory = Descendants(dialog).OfType<TextBox>().Single();
                if (directory.Text != path || !directory.ReadOnly ||
                    Descendants(dialog).OfType<Button>().Single(button => button.AccessibleName == "选择文件夹").Enabled)
                    throw new InvalidOperationException("A busy WordGit must still display the current directory");
            }
        }

        private static void VerifyNoOverlap(Control root)
        {
            var form = root as Form;
            var show = form != null && !form.Visible;
            if (show) { form.Show(); Application.DoEvents(); }
            try
            {
                foreach (Control parent in Descendants(root).Concat(new[] { root }))
                {
                    var controls = parent.Controls.Cast<Control>().Where(control => control.Visible &&
                        (control is Label || control is TextBoxBase || control is Button)).ToArray();
                    for (var i = 0; i < controls.Length; i++)
                        for (var j = i + 1; j < controls.Length; j++)
                            if (controls[i].Bounds.IntersectsWith(controls[j].Bounds))
                                throw new InvalidOperationException("Overlapping controls: " + controls[i].Text + " / " + controls[j].Text);
                }
            }
            finally { if (show) form.Hide(); }
        }

        private static void VerifyDialogPlacement()
        {
            var area = Screen.PrimaryScreen.WorkingArea;
            using (var owner = new Form { StartPosition = FormStartPosition.Manual,
                Bounds = new Rectangle(area.Left + 80, area.Top + 60, Math.Min(900, area.Width - 120), Math.Min(680, area.Height - 100)),
                Opacity = 0, ShowInTaskbar = false })
            {
                owner.Show();
                var gate = new System.Threading.Tasks.TaskCompletionSource<string>();
                Dialogs.ShowBackgroundOperation(owner, "居中测试", () => gate.Task);
                var dialog = Application.OpenForms.Cast<Form>().Single(form => form.Text == "WordGit · 居中测试");
                Application.DoEvents();
                var expected = DialogPlacement.CenteredBounds(dialog.Size, owner.Bounds, Screen.FromControl(owner).WorkingArea);
                if (Math.Abs(dialog.Left - expected.Left) > 2 || Math.Abs(dialog.Top - expected.Top) > 2)
                    throw new InvalidOperationException("The modeless status dialog was not centered on its document owner");
                gate.SetResult("完成。");
                PumpUntil(() => Descendants(dialog).OfType<Button>().Single(button => button.Text == "关闭").Enabled);
                dialog.Close();
            }
            var clipped = DialogPlacement.CenteredBounds(new Size(460, 220), new Rectangle(-3000, -900, 500, 300),
                new Rectangle(-1920, 0, 1920, 1080));
            if (clipped.Left < -1920 || clipped.Top < 0 || clipped.Right > 0 || clipped.Bottom > 1080)
                throw new InvalidOperationException("The centered dialog escaped its monitor");
        }

        private static void VerifyHistoryPane(string output, List<Revision> items, string mainTip, string reviewTip)
        {
            var current = "main";
            var state = WorkingCopyState.Clean;
            var selected = new List<string>();
            using (var host = new Form { ClientSize = new Size(420, 520), ForeColor = Color.White })
            using (var pane = new HistoryPane { Dock = DockStyle.Fill })
            {
                host.Controls.Add(pane);
                pane.Bind(() => new HistoryPaneData
                {
                    DocumentPath = @"C:\test\论文.docx", HistoryPath = @"C:\test\论文.docx", Branch = current,
                    Head = current == "main" ? mainTip : reviewTip, Branches = new[] { "main", "review" },
                    Load = filter => filter == "review" ? new List<Revision> { items[2], items[3] } : items,
                    WorkingState = () => state,
                    LoadReviewItems = () => new List<DocumentReviewItem> {
                        DocumentReviewItem.ForComment(new DocumentComment { Index = 1, Page = 2, Author = "导师", Text = "请补充试验条件。", ScopeStart = 12 }),
                        DocumentReviewItem.ForComment(new DocumentComment { Index = 2, Page = 5, Author = "作者", Text = "这一段需要核对数据。", ScopeStart = 80 }),
                        new DocumentReviewItem { Index = 1, Page = 3, Author = "作者", Text = "新增试验方法。", ScopeStart = 40, RevisionType = 1 },
                        new DocumentReviewItem { Index = 2, Page = 4, Author = "导师", Text = "删除重复描述。", ScopeStart = 60, RevisionType = 2 },
                        new DocumentReviewItem { Index = 3, Page = 5, Author = "导师", Text = "字体：加粗。", ScopeStart = 90, RevisionType = 3 }
                    },
                    NavigateReviewItem = item => selected.Add(item.IsComment ? "comment:" + item.Page : "change:" + item.RevisionType + ":" + item.Page)
                }, (path, action, first, second) =>
                {
                    if (path != @"C:\test\论文.docx") throw new InvalidOperationException("History pane lost the source path");
                    selected.Add(action + ":" + first + ":" + second);
                });
                pane.Reload();
                var commentList = Descendants(pane).OfType<ListBox>().Single();
                if (commentList.Items.Count != 5) throw new InvalidOperationException("Review list is incomplete");
                var reviewFilter = Descendants(pane).OfType<ComboBox>().Single(box => box.AccessibleName == "筛选批注与修订");
                reviewFilter.SelectedItem = "修订";
                if (commentList.Items.Count != 3) throw new InvalidOperationException("Revision filter is incomplete");
                commentList.SelectedIndex = 0;
                if (selected.Last() != "change:1:3") throw new InvalidOperationException("Revision navigation sent the wrong change");
                reviewFilter.SelectedItem = "批注";
                if (commentList.Items.Count != 2) throw new InvalidOperationException("Comment filter is incomplete");
                commentList.SelectedIndex = 1;
                if (selected.Last() != "comment:5") throw new InvalidOperationException("Comment navigation sent the wrong page");
                reviewFilter.SelectedItem = "全部";
                selected.Clear();
                Render(host, Path.Combine(output, "history-pane.png"));
                if (Descendants(pane).OfType<Button>().Any(button => button.ForeColor != SystemColors.ControlText))
                    throw new InvalidOperationException("The pane inherited dark theme text on its light surface");
                var grid = Descendants(pane).OfType<DataGridView>().Single();
                var filterBox = Descendants(pane).OfType<ComboBox>().Single(box => box.AccessibleName == "筛选分支");
                var preview = Descendants(pane).OfType<Button>().Single(button => button.Text == "预览");
                var compare = Descendants(pane).OfType<Button>().Single(button => button.Text == "对比当前稿");
                host.Opacity = 0;
                host.Show();
                preview.PerformClick();
                grid.Rows[2].Selected = true;
                compare.PerformClick();
                if (preview.Enabled || compare.Text != "对比所选版本" ||
                    !selected.SequenceEqual(new[] { "Preview:" + mainTip + ":", "Compare:" + reviewTip + ":" + mainTip }))
                    throw new InvalidOperationException("History pane sent the wrong revisions");
                filterBox.SelectedItem = "review";
                if (grid.RowCount != 2 || Descendants(pane).OfType<Label>().All(label => label.Text != "当前分支  main"))
                    throw new InvalidOperationException("Branch filtering changed the current branch");
                var compareLatest = Descendants(pane).OfType<Button>().Single(button => button.Text == "对比最新提交");
                compareLatest.PerformClick();
                if (selected.Last() != "CompareLatest::")
                    throw new InvalidOperationException("Latest comparison must ignore the selected row and filter");
                var commit = Descendants(pane).OfType<Button>().Single(button => button.Text == "提交版本");
                commit.PerformClick();
                if (!commit.Enabled || selected.Last() != "Commit::")
                    throw new InvalidOperationException("Pane commit must target its current document and branch");
                var import = Descendants(pane).OfType<Button>().Single(button => button.Text == "返稿合并");
                import.PerformClick();
                if (!import.Enabled || selected.Last() != "ImportReview::")
                    throw new InvalidOperationException("Pane review import did not target its document");
                state = WorkingCopyState.Unsaved;
                pane.UpdateWorkingState();
                if (Descendants(pane).OfType<Label>().All(label => label.Text != "有未保存修改"))
                    throw new InvalidOperationException("Unsaved status was not displayed");
                state = WorkingCopyState.Changed;
                pane.UpdateWorkingState();
                if (Descendants(pane).OfType<Label>().All(label => label.Text != "有未提交修改"))
                    throw new InvalidOperationException("Uncommitted status was not displayed");
                current = "review";
                pane.Reload();
                if ((string)filterBox.SelectedItem != "review" || Descendants(pane).OfType<Label>().All(label => label.Text != "当前分支  review"))
                    throw new InvalidOperationException("History pane did not refresh the current branch");
                host.ClientSize = new Size(300, 440);
                Render(host, Path.Combine(output, "history-pane-narrow.png"));
            }
        }

        private static HistoryPaneData AsyncPaneData(List<Revision> items, string branch)
        {
            return new HistoryPaneData {
                DocumentPath = @"C:\test\异步稿件.docx", HistoryPath = @"C:\test\异步稿件.docx",
                Branch = branch, Head = items[0].Id, Branches = new[] { "main", "review" },
                Load = filter => items, WorkingState = () => WorkingCopyState.Clean,
                LoadReviewItems = () => new List<DocumentReviewItem>()
            };
        }

        private static void PumpUntil(Func<bool> complete, [System.Runtime.CompilerServices.CallerLineNumber] int line = 0)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (!complete() && DateTime.UtcNow < deadline)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(1);
            }
            if (!complete()) throw new InvalidOperationException("Asynchronous pane check timed out at line " + line);
        }

        private static void VerifyAsyncHistoryPane(string output, List<Revision> items)
        {
            var uiThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
            var metadata = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
            using (var graphEntered = new System.Threading.ManualResetEventSlim())
            using (var graphRelease = new System.Threading.ManualResetEventSlim())
            using (var workingEntered = new System.Threading.ManualResetEventSlim())
            using (var workingRelease = new System.Threading.ManualResetEventSlim())
            using (var host = new Form { ClientSize = new Size(420, 720), Opacity = 0, ShowInTaskbar = false })
            using (var pane = new HistoryPane { Dock = DockStyle.Fill })
            using (var heartbeat = new Timer { Interval = 15 })
            {
                host.Controls.Add(pane);
                host.Show();
                var ticks = 0;
                heartbeat.Tick += (sender, args) => ticks++;
                heartbeat.Start();
                pane.BindAsync(() => metadata.Task, (path, action, first, second) => { });
                var clock = System.Diagnostics.Stopwatch.StartNew();
                var pending = pane.ReloadAsync();
                var dispatchMs = clock.ElapsedMilliseconds;
                PumpUntil(() => ticks > 0);
                var frame = Descendants(pane).Single(control => control.AccessibleName == "WordGit 正在加载");
                if (!frame.Visible || frame.Bounds != pane.ClientRectangle)
                    throw new InvalidOperationException("The initial loading surface did not cover the pane");
                Render(host, Path.Combine(output, "history-loading.png"));
                host.Show();
                if (pending.IsCompleted || Descendants(pane).OfType<Button>().Single(button => button.Text == "提交版本").Enabled)
                    throw new InvalidOperationException("The pane did not remain responsive while metadata was pending");

                var saved = true;
                var snapshot = AsyncPaneData(items, "main");
                snapshot.Load = filter => {
                    if (System.Threading.Thread.CurrentThread.ManagedThreadId == uiThread)
                        throw new InvalidOperationException("History I/O ran on the UI thread");
                    graphEntered.Set();
                    if (!graphRelease.Wait(4000)) throw new InvalidOperationException("History gate timed out");
                    return items;
                };
                snapshot.IsSaved = () => {
                    if (System.Threading.Thread.CurrentThread.ManagedThreadId != uiThread)
                        throw new InvalidOperationException("A Word saved-state query left its UI thread");
                    return saved;
                };
                snapshot.WorkingState = () => {
                    if (System.Threading.Thread.CurrentThread.ManagedThreadId == uiThread)
                        throw new InvalidOperationException("Document hashing ran on the UI thread");
                    workingEntered.Set();
                    if (!workingRelease.Wait(4000)) throw new InvalidOperationException("Working-state gate timed out");
                    return WorkingCopyState.Clean;
                };
                metadata.SetResult(snapshot);
                PumpUntil(() => graphEntered.IsSet);
                var before = ticks;
                PumpUntil(() => ticks > before);
                if (pending.IsCompleted || !frame.Visible) throw new InvalidOperationException("The loading frame was removed before history was ready");
                graphRelease.Set();
                PumpUntil(() => pending.IsCompleted && workingEntered.IsSet);
                pending.GetAwaiter().GetResult();
                saved = false;
                workingRelease.Set();
                PumpUntil(() => Descendants(pane).OfType<Label>().Any(label => label.Text == "有未保存修改"));
                if (frame.Visible || Descendants(pane).OfType<DataGridView>().Single().RowCount != items.Count ||
                    !Descendants(pane).OfType<Button>().Single(button => button.Text == "提交版本").Enabled)
                    throw new InvalidOperationException("The asynchronous pane did not apply its result");

                host.ClientSize = new Size(300, 520);
                pane.PrepareForShow();
                Render(host, Path.Combine(output, "history-loading-narrow.png"));
                host.Show();
                var reopened = pane.ReloadAsync();
                PumpUntil(() => reopened.IsCompleted);
                reopened.GetAwaiter().GetResult();
                if (frame.Visible) throw new InvalidOperationException("Reopening left the loading frame visible");

                var stale = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
                var requests = 0;
                pane.BindAsync(() => ++requests == 1 ? stale.Task :
                    System.Threading.Tasks.Task.FromResult(AsyncPaneData(items, "review")),
                    (path, action, first, second) => { });
                var older = pane.ReloadAsync();
                var newer = pane.ReloadAsync();
                PumpUntil(() => newer.IsCompleted);
                newer.GetAwaiter().GetResult();
                stale.SetResult(AsyncPaneData(items, "main"));
                PumpUntil(() => older.IsCompleted);
                older.GetAwaiter().GetResult();
                if (!Descendants(pane).OfType<Label>().Any(label => label.Text == "当前分支  review"))
                    throw new InvalidOperationException("An older request replaced the newest branch");

                var failed = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
                pane.BindAsync(() => failed.Task, (path, action, first, second) => { });
                pane.PrepareForShow();
                var failedLoad = pane.ReloadAsync();
                failed.SetException(new IOException("读取失败测试"));
                PumpUntil(() => failedLoad.IsCompleted);
                failedLoad.GetAwaiter().GetResult();
                if (frame.Visible || !Descendants(pane).OfType<Label>().Any(label => label.Text == "读取失败测试"))
                    throw new InvalidOperationException("A failed request left the loading cover stuck");

                var closing = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
                pane.BindAsync(() => closing.Task, (path, action, first, second) => { });
                var unfinished = pane.ReloadAsync();
                pane.Dispose();
                closing.SetResult(AsyncPaneData(items, "main"));
                PumpUntil(() => unfinished.IsCompleted);
                unfinished.GetAwaiter().GetResult();
                Console.WriteLine("PASS: pane dispatch " + dispatchMs + " ms; UI heartbeat, background history/hash, UI-only saved state, stale results and close.");
            }
        }

        private static void VerifyPaneReuse(List<Revision> items)
        {
            using (var host = new Form { ClientSize = new Size(420, 720), Opacity = 0, ShowInTaskbar = false })
            {
                host.Show();
                var factory = new PaneTestFactory(host);
                using (var panes = new TaskPanes(factory))
                {
                    var window = new PaneTestWindow { Hwnd = host.Handle.ToInt32() };
                    var path = @"C:\test\异步稿件.docx";
                    var otherPath = @"C:\test\另一稿件.docx";
                    var result = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
                    var reads = 0;
                    Func<System.Threading.Tasks.Task<HistoryPaneData>> load = () => { reads++; return result.Task; };
                    Action<string, HistoryAction, string, string> open = (p, action, one, two) => { };
                    panes.Show(window, path, load, open);
                    panes.Show(window, path, load, open);
                    PumpUntil(() => reads == 1);
                    var pane = factory.Pane.Control;
                    var frame = Descendants(pane).OfType<PaneLoadingFrame>().Single();
                    panes.Show(window, path, load, open);
                    Application.DoEvents();
                    if (reads != 1 || !frame.Visible || factory.Creates != 1)
                        throw new InvalidOperationException("Repeated clicks restarted a pending pane load");
                    var data = AsyncPaneData(items, "main");
                    data.InitialHistory = items;
                    result.SetResult(data);
                    var submit = Descendants(pane).OfType<Button>().Single(button => button.Text == "提交版本");
                    PumpUntil(() => !frame.Visible && submit.Enabled);
                    var filter = Descendants(pane).OfType<ComboBox>().Single(box => box.Items.Contains("review"));
                    filter.SelectedItem = "review";
                    var preview = Descendants(pane).OfType<Button>().Single(button => button.Text == "预览");
                    PumpUntil(() => preview.Enabled);
                    var grid = Descendants(pane).OfType<DataGridView>().Single();
                    grid.CurrentCell = grid.Rows[2].Cells[1];
                    grid.ClearSelection();
                    grid.Rows[2].Selected = true;
                    var clock = System.Diagnostics.Stopwatch.StartNew();
                    for (var i = 0; i < 10; i++) panes.Show(window, path, load, open);
                    Application.DoEvents();
                    if (reads != 1 || frame.Visible || (string)filter.SelectedItem != "review" ||
                        grid.SelectedRows.Count != 1 || grid.SelectedRows[0].Index != 2)
                        throw new InvalidOperationException("Showing an open pane reloaded or reset its selection");
                    Console.WriteLine("PASS: ten repeated pane clicks " + clock.ElapsedMilliseconds + " ms, zero extra loads.");
                    factory.Pane.Visible = false;
                    panes.Show(window, path, load, open);
                    if (reads != 1 || frame.Visible || !factory.Pane.Visible)
                        throw new InvalidOperationException("Reopening a cached pane restarted loading");
                    if (!panes.Refresh(window)) throw new InvalidOperationException("Explicit refresh was ignored");
                    PumpUntil(() => reads == 2 && submit.Enabled);
                    result = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
                    panes.Show(window, otherPath, load, open);
                    PumpUntil(() => reads == 3);
                    var other = AsyncPaneData(items, "review");
                    other.DocumentPath = other.HistoryPath = otherPath;
                    other.InitialHistory = items;
                    result.SetResult(other);
                    PumpUntil(() => !frame.Visible && submit.Enabled);
                    if (pane.DocumentPath != otherPath) throw new InvalidOperationException("A changed document reused the previous content");
                    result = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
                    panes.Show(window, path, load, open);
                    panes.Show(window, otherPath, load, open);
                    PumpUntil(() => reads == 4);
                    result.SetException(new IOException("加载失败"));
                    PumpUntil(() => !frame.Visible && !pane.CanReuse(otherPath));
                    result = new System.Threading.Tasks.TaskCompletionSource<HistoryPaneData>();
                    result.SetResult(other);
                    panes.Show(window, otherPath, load, open);
                    PumpUntil(() => reads == 5 && !frame.Visible && submit.Enabled);
                    factory.Pane.Visible = false;
                    if (panes.Refresh(window)) throw new InvalidOperationException("A hidden pane reported a visible refresh");
                    Application.DoEvents();
                    if (reads != 5) throw new InvalidOperationException("A hidden pane loaded eagerly");
                    panes.Show(window, otherPath, load, open);
                    PumpUntil(() => reads == 6 && !frame.Visible && submit.Enabled);
                    Console.WriteLine("PASS: queued loads, cached reopening, explicit refresh, document changes, retry and hidden-pane invalidation.");
                }
            }
        }

        private static void VerifyMergeDraftPane(string output, string head, string sourceTip)
        {
            using (var host = new Form { ClientSize = new Size(420, 720), Opacity = 0, ShowInTaskbar = false })
            {
                host.Show();
                var original = Path.Combine(output, "原稿.docx");
                var resultPath = Path.Combine(output, "分支合并稿.docx");
                var draftWindow = new PaneTestWindow { Hwnd = host.Handle.ToInt32() };
                var draft = new PaneTestDraft { ActiveWindow = draftWindow };
                draftWindow.Document = draft;
                var originalWindow = new PaneTestWindow { Hwnd = 0 };
                dynamic application = new System.Dynamic.ExpandoObject();
                application.ActiveWindow = originalWindow;
                application.Windows = new List<PaneTestWindow> { originalWindow, draftWindow };

                const System.Reflection.BindingFlags fields = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                var connect = new Connect();
                typeof(Connect).GetField("application", fields).SetValue(connect, (object)application);
                var factory = new PaneTestFactory(host);
                connect.CTPFactoryAvailable(factory);
                Array custom = null;
                try
                {
                    // The core owns the sidecar now; the add-in only reads it through status.
                    File.WriteAllLines(resultPath + ".wordgit", new[] { original, head, "main", "review", sourceTip });
                    draft.FullName = resultPath;
                    typeof(Connect).GetMethod("ShowMergePane", fields).Invoke(connect, new object[] { draft });
                    if (factory.Creates != 1 || !factory.Pane.Visible || !factory.Pane.Control.CanReuse(resultPath) ||
                        !File.ReadAllLines(resultPath + ".wordgit").SequenceEqual(new[] { original, head, "main", "review", sourceTip }))
                        throw new InvalidOperationException("The saved merge draft did not get its own visible review pane");
                    Console.WriteLine("PASS: the saved merge draft opens its own pane while preserving both history parents.");
                }
                finally
                {
                    connect.OnDisconnection(0, ref custom);
                    File.Delete(resultPath + ".wordgit");
                }
            }
        }

        private static void VerifyRevisionActions(string output)
        {
            var pending = new List<DocumentReviewItem> {
                DocumentReviewItem.ForComment(new DocumentComment { Index = 1, Page = 1, Author = "导师", Text = "核对数据", ScopeStart = 2 }),
                new DocumentReviewItem { Index = 1, StoryType = 1, Page = 1, RevisionType = 1, Author = "作者", Text = "新增方法", RangeText = "新增方法", ScopeStart = 10 },
                new DocumentReviewItem { Index = 2, StoryType = 1, Page = 2, RevisionType = 2, Author = "导师", Text = "重复段落", RangeText = "重复段落", ScopeStart = 30 },
                new DocumentReviewItem { Index = 3, StoryType = 1, Page = 2, RevisionType = 3, Author = "导师", Text = "字体加粗", RangeText = "试验结果", ScopeStart = 50 }
            };
            var navigated = new List<string>();
            var decisions = new List<string>();
            var fail = false;
            var editable = true;
            var merging = true;
            var failReviewRead = false;
            var failReviewSignature = false;
            var signature = "initial";
            var nativeCount = 0;
            var commits = 0;
            using (var host = new Form { ClientSize = new Size(420, 640), Opacity = 0 })
            using (var pane = new HistoryPane { Dock = DockStyle.Fill })
            {
                host.Controls.Add(pane);
                pane.Bind(() => new HistoryPaneData {
                    DocumentPath = @"C:\test\审阅.docx", HistoryPath = @"C:\test\审阅.docx", Branch = "main",
                    Branches = new string[0], Load = branch => new List<Revision>(), WorkingState = () => WorkingCopyState.Unsaved,
                    LoadReviewItems = () => {
                        if (failReviewRead) throw new NullReferenceException("模拟 Word 读取修订时返回空对象");
                        return pending.ToList();
                    },
                    ReviewSignature = () => {
                        if (failReviewSignature) throw new NullReferenceException("模拟 Word 切换窗口时批注不可用");
                        return signature;
                    },
                    NavigateReviewItem = item => navigated.Add(item.Text),
                    CanResolveRevisions = editable, MergeTargetBranch = merging ? "main" : null,
                    DocumentRevisionCount = () => nativeCount,
                    ResolveReviewItem = (item, accept) => {
                        if (fail) return false;
                        decisions.Add((accept ? "accept:" : "reject:") + item.Text);
                        pending.Remove(item);
                        if (item.RevisionType == 14 || item.RevisionType == 15)
                            pending.RemoveAll(entry => !entry.IsComment && entry.RangeText == item.RangeText &&
                                entry.RevisionType == (item.RevisionType == 14 ? 15 : 14));
                        var number = 1;
                        foreach (var change in pending.Where(entry => !entry.IsComment)) change.Index = number++;
                        return true;
                    }
                }, (path, action, first, second) => {
                    if (action == HistoryAction.Commit) commits++;
                });
                host.Show();
                pane.Reload();
                var finish = Descendants(pane).OfType<Button>().Single(button => button.Text == "完成合并");
                var labels = Descendants(pane).OfType<Label>().ToList();
                var reviewFilter = Descendants(pane).OfType<ComboBox>().Single(box => box.AccessibleName == "筛选批注与修订");
                reviewFilter.SelectedItem = "批注";
                if (finish.Enabled || labels.All(label => label.Text != "剩余修订 3 条 · 批注 1") ||
                    labels.All(label => label.Text != "合入分支  main"))
                    throw new InvalidOperationException("Filtering comments hid the pending revisions or enabled merge");
                reviewFilter.SelectedItem = "全部";
                var list = Descendants(pane).OfType<ListBox>().Single();
                var acceptButton = Descendants(pane).OfType<Button>().Single(button => button.Text == "接受");
                var rejectButton = Descendants(pane).OfType<Button>().Single(button => button.Text == "拒绝");
                list.SelectedIndex = 0;
                if (acceptButton.Enabled || rejectButton.Enabled) throw new InvalidOperationException("A comment can be resolved as a revision");
                list.SelectedIndex = 1;
                if (!acceptButton.Enabled || !rejectButton.Enabled) throw new InvalidOperationException("Revision actions are disabled");
                Render(host, Path.Combine(output, "revision-actions.png"));
                host.Show();
                acceptButton.PerformClick();
                if (decisions.Single() != "accept:新增方法" || navigated.Last() != "重复段落" || list.SelectedIndex != 1 ||
                    labels.All(label => label.Text != "剩余修订 2 条 · 批注 1") || finish.Enabled)
                    throw new InvalidOperationException("Accept did not update the remaining count or advance to the next revision");
                fail = true;
                var navigationCount = navigated.Count;
                rejectButton.PerformClick();
                if (pending.Count != 3 || navigated.Count != navigationCount || list.SelectedIndex != 1)
                    throw new InvalidOperationException("A failed resolution advanced the selection");
                fail = false;
                rejectButton.PerformClick();
                if (decisions.Last() != "reject:重复段落" || navigated.Last() != "字体加粗")
                    throw new InvalidOperationException("Reject did not advance to the next revision");
                acceptButton.PerformClick();
                if (pending.Count != 1 || list.SelectedIndex != -1 || acceptButton.Enabled || rejectButton.Enabled ||
                    Descendants(pane).OfType<Label>().All(label => label.Text != "修订已全部处理。"))
                    throw new InvalidOperationException("Finished revisions left an enabled action or lost a comment");
                if (!finish.Enabled || labels.All(label => label.Text != "剩余修订 0 条 · 批注 1"))
                    throw new InvalidOperationException("Remaining comments should not block a finished merge");
                var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                var poll = typeof(HistoryPane).GetMethod("PollReviewItems", flags);
                Action refreshReview = () => {
                    typeof(HistoryPane).GetField("commentPoll", flags).SetValue(pane, 3);
                    poll.Invoke(pane, null);
                };
                var commitGraph = Descendants(pane).OfType<DataGridView>().Single();
                signature = "changed";
                failReviewRead = true;
                refreshReview();
                if (finish.Enabled || list.Enabled || list.Items.Count != 1 || !commitGraph.Enabled ||
                    labels.All(label => label.Text != "批注与修订 · 稍后重试"))
                    throw new InvalidOperationException("A failed timed revision read escaped, cleared the pane or enabled an unsafe merge");
                failReviewRead = false;
                failReviewSignature = true;
                refreshReview();
                if (finish.Enabled || !commitGraph.Enabled)
                    throw new InvalidOperationException("A missing review signature blocked the commit graph or enabled an unsafe merge");
                failReviewSignature = false;
                refreshReview();
                if (!finish.Enabled || !list.Enabled || list.Items.Count != 1 ||
                    labels.All(label => label.Text != "剩余修订 0 条 · 批注 1"))
                    throw new InvalidOperationException("The next successful poll did not restore the review list or its actions");
                Console.WriteLine("PASS: the commit pane survives null objects during timed review reads and recovers on the next poll.");
                nativeCount = 1;
                pane.ReloadComments();
                if (finish.Enabled) throw new InvalidOperationException("A remaining Word revision was ignored");
                nativeCount = 0;
                pane.ReloadComments();
                finish.PerformClick();
                if (commits != 1) throw new InvalidOperationException("The finish button did not use the existing commit flow");

                pending.Add(new DocumentReviewItem { Index = 1, StoryType = 1, Page = 1, RevisionType = 1, Author = "作者", Text = "前一条", RangeText = "前一条", ScopeStart = 10 });
                pending.Add(new DocumentReviewItem { Index = 2, StoryType = 1, Page = 2, RevisionType = 1, Author = "作者", Text = "最后一条", RangeText = "最后一条", ScopeStart = 30 });
                pane.Reload();
                list.SelectedIndex = 2;
                rejectButton.PerformClick();
                if (navigated.Last() != "前一条") throw new InvalidOperationException("Last revision did not wrap to an unprocessed revision");
                pending.RemoveAll(item => !item.IsComment);
                pending.Add(new DocumentReviewItem { Index = 1, StoryType = 1, Page = 1, RevisionType = 14, Author = "作者", Text = "移动来源", RangeText = "同一移动段落", ScopeStart = 10 });
                pending.Add(new DocumentReviewItem { Index = 2, StoryType = 1, Page = 1, RevisionType = 1, Author = "作者", Text = "移动之前", RangeText = "移动之前", ScopeStart = 20 });
                pending.Add(new DocumentReviewItem { Index = 3, StoryType = 1, Page = 1, RevisionType = 15, Author = "作者", Text = "移动目标", RangeText = "同一移动段落", ScopeStart = 30 });
                pending.Add(new DocumentReviewItem { Index = 4, StoryType = 1, Page = 1, RevisionType = 2, Author = "导师", Text = "移动之后", RangeText = "移动之后", ScopeStart = 40 });
                pane.Reload();
                list.SelectedIndex = 3;
                acceptButton.PerformClick();
                if (navigated.Last() != "移动之后" || pending.Count != 3)
                    throw new InvalidOperationException("A linked move caused next-revision navigation to skip or jump backwards");
                editable = false;
                pane.Reload();
                if (acceptButton.Enabled || rejectButton.Enabled) throw new InvalidOperationException("Read-only review actions are enabled");
                host.ClientSize = new Size(300, 440);
                Render(host, Path.Combine(output, "revision-actions-narrow.png"));
                var graph = Descendants(pane).OfType<DataGridView>().Single();
                if (graph.Height < pane.Font.Height + 8) throw new InvalidOperationException("Review actions hid the commit graph");
                merging = false;
                pane.Reload();
                if (finish.Text != "提交版本" || !finish.Enabled)
                    throw new InvalidOperationException("The normal commit button did not return after merge completion");
            }
        }

        private static void VerifyRibbon()
        {
            VerifyRibbonMarkup(Connect.RibbonMarkup(true), true);
            VerifyRibbonMarkup(Connect.RibbonMarkup(false), false);
            if (new Connect().GetCustomUI("Microsoft.Word.Document") != Connect.RibbonMarkup(true))
                throw new InvalidOperationException("Ribbon without a host must default to native icons");
        }

        private static void VerifyRibbonMarkup(string markup, bool nativeIcons)
        {
            var ribbon = new System.Xml.XmlDocument();
            ribbon.LoadXml(markup);
            var tab = ribbon.GetElementsByTagName("tab").Cast<System.Xml.XmlElement>().Single();
            if (tab.GetAttribute("label") != "WordGit")
                throw new InvalidOperationException("Unexpected WordGit tab label");
            var groups = tab.ChildNodes.OfType<System.Xml.XmlElement>().Where(node => node.LocalName == "group").ToArray();
            if (!groups.Select(group => group.GetAttribute("label")).SequenceEqual(new[] { "本地提交", "分支", "比较与审阅", "远程协作", "取回版本", "关于与更新" }))
                throw new InvalidOperationException("Unexpected Ribbon group order");
            if (groups.Any(group => group.ChildNodes.OfType<System.Xml.XmlElement>().Any(node => node.GetAttribute("onAction") == "OnCompressHistory")))
                throw new InvalidOperationException("Storage optimization is exposed as a primary Ribbon button");
            var buttons = ribbon.GetElementsByTagName("button").Cast<System.Xml.XmlElement>().ToArray();
            var callbacks = buttons.Select(button => button.GetAttribute("onAction")).ToArray();
            if (buttons.Length != 19 || callbacks.Distinct().Count() != 19 || !callbacks.Contains("OnDeleteBranch") ||
                callbacks.Contains("OnMergeReview") ||
                !buttons.Any(button => button.GetAttribute("id") == "WordGitImportReview" &&
                    button.GetAttribute("label").Replace("\u2060", "").Replace("\u200B", "") == "返稿合并") ||
                !buttons.Any(button => button.GetAttribute("label").Replace('\u00A0', ' ') == "WordGit " + AppInfo.DisplayVersion && button.GetAttribute("onAction") == "OnAbout") ||
                !callbacks.Contains("OnCheckUpdates") || !callbacks.Contains("OnImportReview") ||
                !buttons.Any(button => button.GetAttribute("onAction") == "OnDataDirectory" &&
                    button.GetAttribute("label") == "数\u2060据\u200B目\u2060录") ||
                new[] { "OnLinkRemote", "OnCloneRemote", "OnFetchRemote", "OnPullRemote", "OnPushRemote" }.Any(action => !callbacks.Contains(action)) ||
                callbacks.Contains("OnCompressHistory") || ribbon.GetElementsByTagName("menu").Cast<System.Xml.XmlElement>().Any(menu => menu.GetAttribute("id") == "WordGitMaintenance") ||
                buttons.Any(button => String.IsNullOrEmpty(button.GetAttribute("imageMso")) == String.IsNullOrEmpty(button.GetAttribute("getImage"))) ||
                callbacks.Any(callback => typeof(IWordGitActions).GetMethod(callback) == null || typeof(Connect).GetMethod(callback) == null))
                throw new InvalidOperationException("Missing Ribbon button, icon, or COM callback");
            foreach (var pair in new[]
                {
                    new { Id = "WordGitHistory", Word = "ViewOutlineView", Wps = "FileProperties" },
                    new { Id = "WordGitPreview", Word = "FilePrintPreview", Wps = "FileOpen" },
                    new { Id = "WordGitCreateBranch", Word = "FileNew", Wps = "FileSaveAs" },
                    new { Id = "WordGitSwitchBranch", Word = "WindowSwitchWindowsMenuWord", Wps = "Refresh" },
                    new { Id = "WordGitMergeBranch", Word = "ReviewAcceptChange", Wps = "HyperlinkInsert" },
                    new { Id = "WordGitDeleteBranch", Word = "ReviewRejectChange", Wps = "Delete" },
                    new { Id = "WordGitCompare", Word = "ReviewCompareMenu", Wps = "Copy" }
                })
            {
                var button = buttons.Single(item => item.GetAttribute("id") == pair.Id);
                var expected = nativeIcons ? pair.Word : pair.Wps;
                if (button.GetAttribute("imageMso") != expected)
                    throw new InvalidOperationException("Wrong native icon for " + pair.Id);
            }
        }

        private static void VerifyRemoteUi(string output, List<Revision> items)
        {
            foreach (var clone in new[] { false, true })
            using (var dialog = new RemoteDialog(clone, clone ? null : "https://github.com/owner/paper.git", "main"))
            {
                Render(dialog, Path.Combine(output, clone ? "remote-clone.png" : "remote-link.png"));
                dialog.Url.Text = "https://github.com/owner/paper";
                dialog.DocumentPath.Text = Path.Combine(output, "远程克隆的新稿件.docx");
                dialog.Size = dialog.MinimumSize;
                Render(dialog, Path.Combine(output, clone ? "remote-clone-min.png" : "remote-link-min.png"));
                var note = Descendants(dialog).OfType<TextBox>().Single(box => box.Multiline);
                if (note.Bottom > ((Button)dialog.AcceptButton).Parent.Top)
                    throw new InvalidOperationException("Remote instructions overlap buttons");
                dialog.Show();
                ((Button)dialog.AcceptButton).PerformClick();
                if (dialog.Options == null || dialog.Options.Url != "https://github.com/owner/paper" || dialog.Options.Branch != "main")
                    throw new InvalidOperationException("Remote dialog lost connection options");
            }
            var actions = new List<HistoryAction>();
            string lastFilter = null;
            var data = new HistoryPaneData {
                DocumentPath = @"C:\test\当前稿.docx", HistoryPath = @"C:\test\当前稿.docx", Branch = "main", Head = items[0].Id,
                Branches = new[] { "main", ReferenceText.RemotePrefix + "main" },
                Load = filter => { lastFilter = filter; return filter == null ? items : items.Take(2).ToList(); },
                WorkingState = () => WorkingCopyState.Clean,
                LoadReviewItems = () => new List<DocumentReviewItem>(),
                Remote = new RemoteState { Url = "https://github.com/owner/paper.git", RemoteBranch = "main", RemoteTip = items[2].Id, LocalTip = items[0].Id, Ahead = 2, Behind = 1 }
            };
            using (var host = new Form { ClientSize = new Size(420, 720) })
            using (var pane = new HistoryPane { Dock = DockStyle.Fill })
            {
                host.Controls.Add(pane);
                pane.Bind(() => data, (path, action, first, second) => actions.Add(action));
                pane.Reload();
                Render(host, Path.Combine(output, "remote-history-pane.png"));
                if (!Descendants(pane).OfType<Label>().Any(label => label.Text.Contains("领先 2 / 落后 1")))
                    throw new InvalidOperationException("Remote synchronization state is missing");
                host.Show();
                foreach (var label in new[] { "获取", "拉取", "推送" })
                    Descendants(pane).OfType<Button>().Single(button => button.Text == label).PerformClick();
                if (!actions.SequenceEqual(new[] { HistoryAction.FetchRemote, HistoryAction.PullRemote, HistoryAction.PushRemote }))
                    throw new InvalidOperationException("Remote buttons sent the wrong actions");
                var filter = Descendants(pane).OfType<ComboBox>().Single(box => box.AccessibleName == "筛选分支");
                filter.SelectedItem = ReferenceText.RemotePrefix + "main";
                if (lastFilter != ReferenceText.RemotePrefix + "main" || Descendants(pane).OfType<DataGridView>().Single().RowCount != 2)
                    throw new InvalidOperationException("Remote branch view did not load its own history");
                host.ClientSize = new Size(300, 520);
                Render(host, Path.Combine(output, "remote-history-narrow.png"));
                data.DocumentPath = @"C:\test\合并稿.docx";
                pane.Reload();
                if (Descendants(pane).OfType<Button>().Any(button =>
                    new[] { "关联", "获取", "拉取", "推送" }.Contains(button.Text) && button.Enabled))
                    throw new InvalidOperationException("Unfinished review can synchronize as the source document");
            }
        }

        private static void VerifyImportReview(string output, List<Revision> items)
        {
            using (var picker = new HistoryDialog("论文.docx", items, true, items[1].Id, command: HistoryAction.Merge))
            {
                picker.Show();
                var grid = Descendants(picker).OfType<DataGridView>().Single();
                var confirm = (Button)picker.AcceptButton;
                if (confirm.Enabled) throw new InvalidOperationException("Historical merge enabled before two versions were selected");
                grid.ClearSelection();
                grid.Rows[3].Selected = true;
                grid.Rows[1].Selected = true;
                confirm.PerformClick();
                if (picker.Action != HistoryAction.Merge || picker.Basis.Id != items[3].Id || picker.Other.Id != items[1].Id)
                    throw new InvalidOperationException("Historical merge lost the selected versions");
            }
            var file = Path.Combine(output, "导入界面测试返稿.docx");
            File.WriteAllBytes(file, new byte[] { 1, 2, 3 });
            try
            {
                foreach (var mode in new[] { "直接合并", "先保存再合并", "仅保存为分支" })
                using (var dialog = new ImportReviewDialog(Path.Combine(output, "当前稿.docx"), items, "main", new[] { "main", "review" }))
                {
                    if (((Button)dialog.AcceptButton).Enabled || dialog.BasisPicker.SelectedIndex != 0)
                        throw new InvalidOperationException("The submission basis must not default to a commit");
                    Render(dialog, Path.Combine(output, "returned-draft-empty.png"));
                    dialog.FilePath.Text = file;
                    var saveDraft = Descendants(dialog).OfType<Button>().Single(button => button.Text == "仅保存为分支");
                    if (((Button)dialog.AcceptButton).Enabled || saveDraft.Enabled)
                        throw new InvalidOperationException("Choosing only the return file enabled the workflow");
                    var target = Descendants(dialog).OfType<TextBox>().Single(box => box.AccessibleName == "合入分支");
                    if (!target.ReadOnly || target.Text != "main") throw new InvalidOperationException("The checked-out target branch is not visible");
                    dialog.BasisPicker.SelectedIndex = 4;
                    if (!dialog.BranchName.Text.StartsWith("review-", StringComparison.Ordinal))
                        throw new InvalidOperationException("Default review branch conflicts with an existing review branch");
                    dialog.BranchName.Text = "review-老师一审";
                    dialog.Reviewer.Text = "导师甲";
                    dialog.Size = dialog.MinimumSize;
                    dialog.Font = new Font(dialog.Font.FontFamily, 16f);
                    var keep = mode != "直接合并";
                    dialog.SaveOriginal.Checked = keep;
                    if (!keep) dialog.BranchName.Text = "";
                    Render(dialog, Path.Combine(output, "返稿合并-" + mode + ".png"));
                    dialog.Show();
                    var historical = Descendants(dialog).OfType<LinkLabel>().Single(link => link.Text == "合并两个历史版本…");
                    var scroll = Descendants(dialog).OfType<Panel>().Single(panel => panel.AutoScroll);
                    var linkBounds = dialog.RectangleToClient(historical.Parent.RectangleToScreen(historical.Bounds));
                    var scrollBounds = dialog.RectangleToClient(scroll.Parent.RectangleToScreen(scroll.Bounds));
                    if (!historical.Visible || !dialog.ClientRectangle.Contains(linkBounds) || linkBounds.Top < scrollBounds.Bottom)
                        throw new InvalidOperationException("Historical merge link is clipped by the scrolling form");
                    var merge = (Button)dialog.AcceptButton;
                    var saveOnly = Descendants(dialog).OfType<Button>().Single(button => button.Text == "仅保存为分支");
                    if (dialog.BranchName.Visible != keep || !dialog.Reviewer.Visible || saveOnly.Visible != keep || !merge.Enabled)
                        throw new InvalidOperationException("Optional review branch settings did not follow the checkbox");
                    foreach (var button in Descendants(dialog).OfType<Button>().Where(button => button.Visible))
                        if (!dialog.ClientRectangle.Contains(dialog.RectangleToClient(button.Parent.RectangleToScreen(button.Bounds))))
                            throw new InvalidOperationException("Review merge action is outside the visible dialog");
                    var createMerge = mode != "仅保存为分支";
                    (createMerge ? merge : saveOnly).PerformClick();
                    if (dialog.DialogResult != DialogResult.OK || dialog.Options.FilePath != file || dialog.Options.Basis.Id != items[3].Id ||
                        dialog.Options.Branch != (keep ? "review-老师一审" : null) ||
                        dialog.Options.Reviewer != "导师甲" || dialog.Options.CreateMergeDraft != createMerge)
                        throw new InvalidOperationException("Review merge lost selected options");
                }
            }
            finally { File.Delete(file); }
        }

        private static void VerifyAboutDialog(string output)
        {
            var pending = new System.Threading.Tasks.TaskCompletionSource<UpdateResult>();
            var calls = 0;
            using (var dialog = new AboutDialog(token => { calls++; return pending.Task; }))
            {
                var links = Descendants(dialog).OfType<LinkLabel>().ToList();
                var urls = new[] { AppInfo.AuthorUrl, AppInfo.LicenseUrl, AppInfo.ProjectUrl, AppInfo.ReleasesUrl, AppInfo.IssuesUrl };
                if (!links.Select(link => (string)link.Tag).SequenceEqual(urls) ||
                    links.Any(link => link.Text.Contains("://")) || links.Single(link => (string)link.Tag == AppInfo.ProjectUrl).Text != "GitHub 仓库")
                    throw new InvalidOperationException("About dialog lost a named link or shows a raw URL");
                var status = Descendants(dialog).OfType<Label>().Single(label => label.Name == "updateStatus");
                VerifyAboutContent(dialog);
                if (status.Visible || Descendants(dialog).OfType<Button>().Count() != 2 ||
                    !Descendants(dialog).OfType<Label>().Any(label => label.Text == "版本 " + AppInfo.DisplayVersion))
                    throw new InvalidOperationException("About overview has redundant content or lacks the current version");
                Render(dialog, Path.Combine(output, "about-dialog.png"));
                dialog.Font = new Font(dialog.Font.FontFamily, 18f);
                VerifyAboutContent(dialog);
                Render(dialog, Path.Combine(output, "about-dialog-largest-font.png"));
                dialog.Font = SystemFonts.MessageBoxFont;
                dialog.Show();
                var running = dialog.CheckForUpdatesAsync();
                var duplicate = dialog.CheckForUpdatesAsync();
                var check = Descendants(dialog).OfType<Button>().Single(button => button.Text == "检查更新");
                if (calls != 1 || running.IsCompleted || check.Enabled || !status.Visible)
                    throw new InvalidOperationException("Update check blocked the UI or did not show progress");
                pending.SetResult(new UpdateResult(UpdateState.Available,
                    "发现新版本 0.9.10。请打开发布页下载安装；安装前保存并关闭 Word，现有历史会保留。"));
                var deadline = DateTime.UtcNow.AddSeconds(3);
                while (!running.IsCompleted && DateTime.UtcNow < deadline) Application.DoEvents();
                if (!running.IsCompleted || !check.Enabled || !status.Visible || !status.Text.Contains("0.9.10"))
                    throw new InvalidOperationException("Update result was not displayed");
                VerifyAboutContent(dialog);
                Render(dialog, Path.Combine(output, "about-update-result.png"));
                dialog.Font = new Font(dialog.Font.FontFamily, 18f);
                VerifyAboutContent(dialog);
                Render(dialog, Path.Combine(output, "about-update-large-font.png"));
            }
            var cancellation = new System.Threading.Tasks.TaskCompletionSource<UpdateResult>();
            var closing = new AboutDialog(token => {
                token.Register(() => cancellation.TrySetCanceled());
                return cancellation.Task;
            });
            closing.Show();
            var checking = closing.CheckForUpdatesAsync();
            closing.Close();
            Application.DoEvents();
            closing.Dispose();
            if (!cancellation.Task.IsCanceled) throw new InvalidOperationException("Closed update dialog kept its request alive");
        }

        private static void VerifyUpdateDialog(string output)
        {
            var result = new UpdateResult(UpdateState.Available, "当前版本 0.12.0  →  新版本 0.12.1-beta.2",
                "0.12.1-beta.2", String.Join("\n", Enumerable.Repeat("更新日志：修复和改进。", 40)),
                AppInfo.ReleasesUrl + "/download/v0.12.1-beta.2/WordGit-0.12.1-beta.2-Setup.exe");
            string ignored = null;
            Func<string, bool> ignore = version => { ignored = version; return true; };
            using (var dialog = new UpdateDialog(result, ignore))
            {
                dialog.Show();
                Application.DoEvents();
                var versions = Descendants(dialog).OfType<Label>().Single(label => label.Name == "releaseVersions");
                var notes = Descendants(dialog).OfType<TextBox>().Single(box => box.Name == "releaseNotes");
                var download = Descendants(dialog).OfType<Button>().Single(button => button.Text == "下载更新");
                var later = Descendants(dialog).OfType<Button>().Single(button => button.Text == "稍后提醒");
                var skip = Descendants(dialog).OfType<Button>().Single(button => button.Text == "忽略此版本");
                if (versions.Text != result.Message || !notes.ReadOnly || !notes.Text.Contains("更新日志") ||
                    !download.Visible || !download.Enabled)
                    throw new InvalidOperationException("Update popup lacks version, notes or download");
                if (skip.Left != 0 || later.Left <= skip.Right ||
                    download.Right != download.Parent.ClientSize.Width || download.Left <= later.Right)
                    throw new InvalidOperationException("Update actions are not aligned to the left and right edges");
                if (download.UseVisualStyleBackColor || download.BackColor != SystemColors.Highlight ||
                    download.ForeColor != SystemColors.HighlightText || !later.Focused)
                    throw new InvalidOperationException("Download must have its primary color independently of keyboard focus");
                VerifyNoOverlap(dialog);
                Render(dialog, Path.Combine(output, "update-dialog.png"));
                later.PerformClick();
                if (ignored != null || dialog.Visible)
                    throw new InvalidOperationException("Later should close without saving a decision");
            }
            using (var dialog = new UpdateDialog(result, ignore))
            {
                dialog.Size = dialog.MinimumSize;
                dialog.Show();
                Application.DoEvents();
                VerifyNoOverlap(dialog);
                var skip = Descendants(dialog).OfType<Button>().Single(button => button.Text == "忽略此版本");
                skip.PerformClick();
                if (ignored != result.Version || dialog.Visible)
                    throw new InvalidOperationException("Ignore must save the offered version and close the window");
            }
        }
        private static void VerifyAboutContent(AboutDialog dialog)
        {
            if (!dialog.Visible) dialog.Show();
            Application.DoEvents();
            VerifyNoOverlap(dialog);
            if (Descendants(dialog).OfType<ScrollableControl>().Any(control => control.AutoScroll) ||
                Descendants(dialog).OfType<TextBoxBase>().Any())
                throw new InvalidOperationException("About information requires scrolling");
            foreach (var control in Descendants(dialog).Where(control => control.Visible && (control is Label || control is Button)))
                if (!dialog.ClientRectangle.Contains(dialog.RectangleToClient(control.Parent.RectangleToScreen(control.Bounds))))
                    throw new InvalidOperationException("About information is outside the window: " + control.Text);
        }

        private static void VerifyBackgroundOperationCancel()
        {
            var gate = new System.Threading.Tasks.TaskCompletionSource<string>();
            var finished = 0;
            Dialogs.ShowBackgroundOperation(null, "取消测试", token =>
            {
                token.Register(() => gate.TrySetCanceled());
                return gate.Task;
            }, () => finished++);
            var dialog = Application.OpenForms.OfType<Form>().Last(form => form.Text == "WordGit · 取消测试");
            var cancel = Descendants(dialog).OfType<Button>().Single(button => button.Text == "取消");
            var close = Descendants(dialog).OfType<Button>().Single(button => button.Text == "关闭");
            cancel.PerformClick();
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (!close.Enabled && DateTime.UtcNow < deadline) Application.DoEvents();
            var status = Descendants(dialog).OfType<TextBox>().Single();
            if (!close.Enabled || cancel.Enabled || !status.Text.Contains("已取消") || finished != 1)
                throw new InvalidOperationException("Background operation did not finish after cancel");
            dialog.Close();
            Application.DoEvents();

            Dialogs.ShowBackgroundOperation(null, "普通测试", () => System.Threading.Tasks.Task.FromResult("完成。"));
            var plain = Application.OpenForms.OfType<Form>().Last(form => form.Text == "WordGit · 普通测试");
            if (Descendants(plain).OfType<Button>().Any(button => button.Text == "取消"))
                throw new InvalidOperationException("Non-cancellable operation shows a cancel button");
            deadline = DateTime.UtcNow.AddSeconds(5);
            var plainClose = Descendants(plain).OfType<Button>().Single(button => button.Text == "关闭");
            while (!plainClose.Enabled && DateTime.UtcNow < deadline) Application.DoEvents();
            if (!plainClose.Enabled || !Descendants(plain).OfType<TextBox>().Single().Text.Contains("完成。"))
                throw new InvalidOperationException("Background operation did not finish");
            plain.Close();
            Application.DoEvents();
        }

        private static Revision RevisionFor(string id, string[] parents, DateTimeOffset date, string title, params string[] refs)
        {
            return new Revision { Id = id, Parents = parents, References = refs, Author = "测试作者", Date = date,
                Subject = title, Message = title + "\r\n\r\n保留原稿排版并记录关键修改。" };
        }

        private static IEnumerable<Control> Descendants(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                yield return child;
                foreach (var nested in Descendants(child)) yield return nested;
            }
        }

        private static void Render(Form form, string path)
        {
            form.ShowInTaskbar = false;
            form.Opacity = 0;
            if (!form.Visible) form.Show();
            Application.DoEvents();
            form.PerformLayout();
            foreach (var control in Descendants(form)) { var handle = control.Handle; }
            form.PerformLayout();
            foreach (var control in Descendants(form).Where(control => control is Button || control is Label))
                if (!control.Parent.ClientRectangle.Contains(control.Bounds) ||
                    control.Height < TextRenderer.MeasureText(control.Text, control.Font).Height)
                    throw new InvalidOperationException("Clipped control: " + control.Text);
            using (var image = new Bitmap(form.Width, form.Height))
            {
                form.DrawToBitmap(image, new Rectangle(Point.Empty, image.Size));
                // DrawToBitmap visits child windows in reverse z-order; reapply the visible cover.
                foreach (var frame in Descendants(form).OfType<PaneLoadingFrame>().Where(control => control.Visible))
                {
                    var origin = form.PointToClient(frame.PointToScreen(Point.Empty));
                    origin.Offset((form.Width - form.ClientSize.Width) / 2,
                        form.Height - form.ClientSize.Height - (form.Width - form.ClientSize.Width) / 2);
                    frame.DrawToBitmap(image, new Rectangle(origin, frame.Size));
                }
                image.Save(path);
            }
            form.Hide();
        }
    }

    internal sealed class PaneTestWindow
    {
        public int Hwnd { get; set; }
        public object Document { get; set; }
    }

    internal sealed class PaneTestDraft
    {
        public PaneTestWindow ActiveWindow { get; set; }
        public string FullName { get; set; }
        public string Path { get { return System.IO.Path.GetDirectoryName(FullName); } }
    }

    internal sealed class PaneTestFactory : ICTPFactory
    {
        private readonly Form host;
        internal PaneTestNative Pane;
        internal int Creates;
        internal PaneTestFactory(Form host) { this.host = host; }
        public object CreateCTP(string controlId, string title, object window)
        {
            Creates++;
            Pane = new PaneTestNative { Control = new HistoryPane { Dock = DockStyle.Fill } };
            host.Controls.Add(Pane.Control);
            return Pane;
        }
    }

    internal sealed class PaneTestNative
    {
        internal HistoryPane Control;
        public object ContentControl { get { return Control; } }
        public int DockPosition { get; set; }
        public int Width { get; set; }
        public bool Visible { get { return Control.Visible; } set { Control.Visible = value; } }
        public void Delete() { Control.Dispose(); }
    }
}
