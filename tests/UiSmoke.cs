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
            var buttons = ribbon.GetElementsByTagName("button").Cast<System.Xml.XmlElement>().ToArray();
            var callbacks = buttons.Select(button => button.GetAttribute("onAction")).ToArray();
            if (buttons.Length != 12 || callbacks.Distinct().Count() != 12 || !callbacks.Contains("OnDeleteBranch") ||
                !buttons.Any(button => button.GetAttribute("label") == "压缩历史" && button.GetAttribute("onAction") == "OnCompressHistory") ||
                buttons.Any(button => String.IsNullOrEmpty(button.GetAttribute("imageMso"))) ||
                callbacks.Any(callback => typeof(IWordGitActions).GetMethod(callback) == null || typeof(Connect).GetMethod(callback) == null))
                throw new InvalidOperationException("Missing Ribbon button, icon, or COM callback");
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
