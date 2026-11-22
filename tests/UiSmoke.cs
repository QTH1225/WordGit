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
                TextBox field;
                using (var form = Dialogs.CommitForm("提交到「论文独创性声明.docx」", "", out field))
                {
                    Render(form, Path.Combine(args[0], "commit-dialog.png"));
                    var submit = Descendants(form).OfType<Button>().Single(button => button.Text == "提交");
                    if (submit.Enabled) throw new InvalidOperationException("Empty message is allowed");
                    field.Text = "校正封面\r\n补充 Word 版本管理记录。";
                    if (!submit.Enabled) throw new InvalidOperationException("Valid message is rejected");
                    form.Size = form.MinimumSize;
                    Render(form, Path.Combine(args[0], "commit-dialog-min.png"));
                }
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
                VerifyHistoryPane(args[0], items, one, three);
                VerifyRevisionActions(args[0]);
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
                var import = Descendants(pane).OfType<Button>().Single(button => button.Text == "导入返稿");
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
            using (var host = new Form { ClientSize = new Size(420, 640), Opacity = 0 })
            using (var pane = new HistoryPane { Dock = DockStyle.Fill })
            {
                host.Controls.Add(pane);
                pane.Bind(() => new HistoryPaneData {
                    DocumentPath = @"C:\test\审阅.docx", HistoryPath = @"C:\test\审阅.docx", Branch = "main",
                    Branches = new string[0], Load = branch => new List<Revision>(), WorkingState = () => WorkingCopyState.Unsaved,
                    LoadReviewItems = () => pending.ToList(), NavigateReviewItem = item => navigated.Add(item.Text),
                    CanResolveRevisions = editable,
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
                }, (path, action, first, second) => { });
                host.Show();
                pane.Reload();
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
                if (decisions.Single() != "accept:新增方法" || navigated.Last() != "重复段落" || list.SelectedIndex != 1)
                    throw new InvalidOperationException("Accept did not advance to the next revision");
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
            }
        }

        private static void VerifyRibbon()
        {
            var ribbon = new System.Xml.XmlDocument();
            ribbon.LoadXml(new Connect().GetCustomUI("Microsoft.Word.Document"));
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
                !buttons.Any(button => button.GetAttribute("label").Replace('\u00A0', ' ') == "WordGit " + AppInfo.DisplayVersion && button.GetAttribute("onAction") == "OnAbout") ||
                !callbacks.Contains("OnCheckUpdates") || !callbacks.Contains("OnImportReview") ||
                new[] { "OnLinkRemote", "OnCloneRemote", "OnFetchRemote", "OnPullRemote", "OnPushRemote" }.Any(action => !callbacks.Contains(action)) ||
                callbacks.Contains("OnCompressHistory") || ribbon.GetElementsByTagName("menu").Cast<System.Xml.XmlElement>().Any(menu => menu.GetAttribute("id") == "WordGitMaintenance") ||
                buttons.Any(button => String.IsNullOrEmpty(button.GetAttribute("imageMso"))) ||
                callbacks.Any(callback => typeof(IWordGitActions).GetMethod(callback) == null || typeof(Connect).GetMethod(callback) == null))
                throw new InvalidOperationException("Missing Ribbon button, icon, or COM callback");
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
                if (dialog.Options == null || dialog.Options.Url != "https://github.com/owner/paper.git" || dialog.Options.Branch != "main")
                    throw new InvalidOperationException("Remote dialog lost connection options");
            }
            var actions = new List<HistoryAction>();
            string lastFilter = null;
            var data = new HistoryPaneData {
                DocumentPath = @"C:\test\当前稿.docx", HistoryPath = @"C:\test\当前稿.docx", Branch = "main", Head = items[0].Id,
                Branches = new[] { "main", GitHistory.RemotePrefix + "main" },
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
                filter.SelectedItem = GitHistory.RemotePrefix + "main";
                if (lastFilter != GitHistory.RemotePrefix + "main" || Descendants(pane).OfType<DataGridView>().Single().RowCount != 2)
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
            using (var picker = new HistoryDialog("论文.docx", items, true, items[1].Id, command: HistoryAction.ImportReview))
            {
                Render(picker, Path.Combine(output, "review-basis-picker.png"));
                var grid = Descendants(picker).OfType<DataGridView>().Single();
                if (grid.MultiSelect || grid.SelectedRows[0].Index != 1 || Descendants(picker).OfType<Button>().Count() != 2)
                    throw new InvalidOperationException("Review import must select exactly one submission basis");
                picker.Show();
                grid.CurrentCell = grid.Rows[3].Cells[1];
                ((Button)picker.AcceptButton).PerformClick();
                if (picker.Action != HistoryAction.ImportReview || picker.Basis.Id != items[3].Id)
                    throw new InvalidOperationException("Review import lost the selected older basis");
            }
            var file = Path.Combine(output, "导入界面测试返稿.docx");
            File.WriteAllBytes(file, new byte[] { 1, 2, 3 });
            try
            {
                using (var dialog = new ImportReviewDialog(Path.Combine(output, "当前稿.docx"), items[3], "main"))
                {
                    if (((Button)dialog.AcceptButton).Enabled) throw new InvalidOperationException("Review file is optional");
                    Render(dialog, Path.Combine(output, "review-import.png"));
                    dialog.FilePath.Text = file;
                    dialog.BranchName.Text = "review/老师一审";
                    dialog.Reviewer.Text = "导师甲";
                    dialog.Size = dialog.MinimumSize;
                    Render(dialog, Path.Combine(output, "review-import-min.png"));
                    var note = Descendants(dialog).OfType<TextBox>().Single(box => box.Multiline);
                    if (note.Bottom > ((Button)dialog.AcceptButton).Parent.Top)
                        throw new InvalidOperationException("Review import instructions overlap the action row");
                    dialog.Show();
                    ((Button)dialog.AcceptButton).PerformClick();
                    if (dialog.DialogResult != DialogResult.OK || dialog.Options.FilePath != file ||
                        dialog.Options.Branch != "review/老师一审" || dialog.Options.Reviewer != "导师甲" || !dialog.Options.MergeAfterImport)
                        throw new InvalidOperationException("Review import lost selected options");
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
                Render(dialog, Path.Combine(output, "about-dialog.png"));
                dialog.Size = dialog.MinimumSize;
                Render(dialog, Path.Combine(output, "about-dialog-min.png"));
                if (!Descendants(dialog).OfType<Label>().Any(label => label.Text.Contains(AppInfo.DisplayVersion)) ||
                    !Descendants(dialog).OfType<LinkLabel>().Any(link => link.Text == AppInfo.ProjectUrl))
                    throw new InvalidOperationException("About dialog lost product information");
                dialog.Show();
                var running = dialog.CheckForUpdatesAsync();
                var duplicate = dialog.CheckForUpdatesAsync();
                var check = Descendants(dialog).OfType<Button>().Single(button => button.Text == "检查更新");
                if (calls != 1 || running.IsCompleted || check.Enabled)
                    throw new InvalidOperationException("Update check blocked the UI or allowed duplicate requests");
                pending.SetResult(new UpdateResult(UpdateState.Available, "发现新版本 0.9.10。请打开发布页下载安装。"));
                var deadline = DateTime.UtcNow.AddSeconds(3);
                while (!running.IsCompleted && DateTime.UtcNow < deadline) Application.DoEvents();
                if (!running.IsCompleted || !check.Enabled ||
                    !Descendants(dialog).OfType<TextBox>().Single().Text.Contains("0.9.10"))
                    throw new InvalidOperationException("Update result was not displayed");
                Render(dialog, Path.Combine(output, "about-update-result.png"));
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
                image.Save(path);
            }
            form.Hide();
        }
    }
}
