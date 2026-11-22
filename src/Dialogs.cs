using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordGit
{
    internal enum HistoryAction { Compare, Export, Restore, Merge, Preview, CreateBranch, SwitchBranch, MergeBranch, DeleteBranch, CompareLatest, Commit, ImportReview, LinkRemote, FetchRemote, PullRemote, PushRemote }

    internal static class DialogLayout
    {
        public static TableLayoutPanel Stack(int rows)
        {
            return new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = rows, Padding = new Padding(14) };
        }

        public static Label Label(string text)
        {
            return new Label { Text = text, AutoSize = true, Margin = new Padding(0, 4, 0, 7) };
        }

        public static Button Button(string label)
        {
            return new Button { Text = label, AutoSize = true, Padding = new Padding(8, 4, 8, 4), Margin = new Padding(0, 0, 8, 4) };
        }

        public static FlowLayoutPanel Buttons(bool rightAligned)
        {
            return new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill,
                WrapContents = false, Margin = new Padding(0, 10, 0, 0),
                FlowDirection = rightAligned ? FlowDirection.RightToLeft : FlowDirection.LeftToRight };
        }
    }

    internal sealed class HistoryDialog : Form
    {
        private readonly DataGridView history = new DataGridView();
        private readonly TextBox details = new TextBox { ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, TabStop = false };
        private readonly Dictionary<HistoryAction, Button> buttons = new Dictionary<HistoryAction, Button>();
        private List<Revision> revisions;
        private List<GraphRow> rows;
        private readonly List<Revision> allRevisions;
        private readonly string head;
        private readonly string branch;
        private readonly Func<string, List<Revision>> loadBranch;
        private readonly ComboBox branchFilter = new ComboBox { Name = "branchFilter", DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
        public HistoryAction Action { get; private set; }
        public Revision Basis { get; private set; }
        public Revision Other { get; private set; }
        public string SelectedBranch { get { return branchFilter.SelectedIndex > 0 ? (string)branchFilter.SelectedItem : null; } }

        public HistoryDialog(string name, List<Revision> revisions, bool allowRestore, string head, string branch = "main",
            string[] branches = null, Func<string, List<Revision>> loadBranch = null, HistoryAction? command = null)
        {
            this.revisions = revisions;
            allRevisions = revisions;
            rows = CommitGraph.Layout(revisions, head);
            this.head = head;
            this.branch = branch;
            this.loadBranch = loadBranch;
            Text = "WordGit · 提交图";
            StartPosition = FormStartPosition.CenterParent;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(940, 575);
            MinimumSize = new Size(790, 470);
            MinimizeBox = false;
            var layout = DialogLayout.Stack(7);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 72));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(layout);

            var displayName = name.Length <= 45 ? name : name.Substring(0, 42) + "…";
            var header = DialogLayout.Label(displayName + "   ·   当前分支：" + ShortBranch(branch));
            header.AccessibleDescription = name + "，当前分支 " + branch;
            header.Font = new Font(Font, FontStyle.Bold);
            layout.Controls.Add(header, 0, 0);
            var branchBar = DialogLayout.Buttons(false);
            branchBar.Margin = new Padding(0, 2, 0, 0);
            branchBar.Controls.Add(DialogLayout.Label("分支视图"));
            branchBar.Controls.Add(branchFilter);
            branchFilter.Items.Add("全部分支");
            foreach (var item in branches ?? new[] { branch }) branchFilter.Items.Add(item);
            branchFilter.SelectedIndex = 0;
            branchFilter.SelectedIndexChanged += (sender, args) =>
            {
                try { FilterGraph(); }
                catch (Exception error)
                {
                    branchFilter.SelectedIndex = 0;
                    MessageBox.Show(this, error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            if (allowRestore && !command.HasValue)
            {
                Add(branchBar, "从所选版本新建", HistoryAction.CreateBranch);
                Add(branchBar, "切换分支", HistoryAction.SwitchBranch);
                Add(branchBar, "合并分支…", HistoryAction.MergeBranch);
            }
            layout.Controls.Add(branchBar, 0, 1);
            layout.Controls.Add(DialogLayout.Label(SelectionHint(command)), 0, 2);
            history.MultiSelect = !command.HasValue || command == HistoryAction.Compare || command == HistoryAction.Merge;
            ConfigureGrid();
            layout.Controls.Add(history, 0, 3);
            layout.Controls.Add(DialogLayout.Label("提交详情"), 0, 4);
            layout.Controls.Add(details, 0, 5);
            var actions = DialogLayout.Buttons(false);
            layout.Controls.Add(actions, 0, 6);
            AddVersionAction(actions, "打开预览", HistoryAction.Preview, command);
            AddVersionAction(actions, "对比版本", HistoryAction.Compare, command);
            AddVersionAction(actions, "另存旧版", HistoryAction.Export, command);
            if (allowRestore) AddVersionAction(actions, "恢复为当前稿", HistoryAction.Restore, command);
            AddVersionAction(actions, "合并审阅稿", HistoryAction.Merge, command);
            if (command == HistoryAction.CreateBranch) AddVersionAction(actions, "新建分支", HistoryAction.CreateBranch, command);
            if (command == HistoryAction.ImportReview) AddVersionAction(actions, "选择送审版本", HistoryAction.ImportReview, command);
            var close = DialogLayout.Button(command.HasValue ? "取消" : "关闭");
            close.DialogResult = DialogResult.Cancel;
            actions.Controls.Add(close);
            CancelButton = close;
            history.SelectionChanged += (sender, args) => ShowSelection();
            history.CellDoubleClick += (sender, args) => { if (args.RowIndex >= 0 && history.SelectedRows.Count == 1) Complete(command == HistoryAction.ImportReview ? HistoryAction.ImportReview : HistoryAction.Preview); };
            if (revisions.Count > 0)
            {
                var preferred = command == HistoryAction.CreateBranch || command == HistoryAction.ImportReview ? revisions.FindIndex(revision => revision.Id == head) : 0;
                var row = history.Rows[Math.Max(0, preferred)];
                history.CurrentCell = row.Cells[1];
                row.Selected = true;
            }
            ShowSelection();
        }

        private void AddVersionAction(FlowLayoutPanel panel, string label, HistoryAction action, HistoryAction? command)
        {
            if (command.HasValue && command != action) return;
            Add(panel, label, action);
            if (command.HasValue)
            {
                Text = "WordGit · " + label;
                AcceptButton = buttons[action];
            }
        }

        private static string SelectionHint(HistoryAction? command)
        {
            switch (command)
            {
                case HistoryAction.Preview: return "选择要只读打开的版本。";
                case HistoryAction.Compare: return "选一版与当前稿对比，或按 Ctrl 选两个历史版本对比。";
                case HistoryAction.Export: return "选择要另存为新文档的版本。";
                case HistoryAction.Restore: return "选择要恢复为当前稿的版本。";
                case HistoryAction.Merge: return "选一版作为基准后选择返稿，或按 Ctrl 选两个历史版本。";
                case HistoryAction.CreateBranch: return "选择新分支的起始版本。";
                case HistoryAction.ImportReview: return "选择当时发给老师的送审版本，作为返稿分支的父提交。";
                default: return "单击查看详情，Ctrl 选两版对比；双击只读打开历史版本。";
            }
        }

        private static string ShortBranch(string value) { return value.Length > 24 ? value.Substring(0, 21) + "…" : value; }

        private void FilterGraph()
        {
            var selected = SelectedBranch;
            var filtered = selected == null || loadBranch == null ? allRevisions : loadBranch(selected);
            history.ClearSelection();
            history.RowCount = 0;
            revisions = filtered;
            rows = CommitGraph.Layout(revisions, head);
            history.RowCount = revisions.Count;
            if (revisions.Count > 0) history.Rows[0].Selected = true;
            history.Invalidate();
            ShowSelection();
        }

        private void ConfigureGrid()
        {
            history.Dock = DockStyle.Fill;
            history.BackgroundColor = SystemColors.Window;
            history.BorderStyle = BorderStyle.FixedSingle;
            history.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            history.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            history.ReadOnly = true;
            history.AllowUserToAddRows = false;
            history.AllowUserToDeleteRows = false;
            history.AllowUserToResizeRows = false;
            history.RowHeadersVisible = false;
            history.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            history.Margin = new Padding(0);
            history.RowTemplate.Height = Math.Max(28, Font.Height + 14);
            Column("提交图", Math.Max(64, 19 * rows.Select(row => row.Columns).DefaultIfEmpty(1).Max() + 16));
            Column("提交说明", 220).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Column("哈希", 90);
            Column("作者", 105);
            Column("时间", 161);
            history.VirtualMode = true;
            history.CellValueNeeded += (sender, args) =>
            {
                var revision = revisions[args.RowIndex];
                switch (args.ColumnIndex)
                {
                    case 1: args.Value = Description(revision); break;
                    case 2: args.Value = revision.Id.Substring(0, 8); break;
                    case 3: args.Value = revision.Author; break;
                    case 4: args.Value = revision.Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm"); break;
                }
            };
            history.RowCount = revisions.Count;
            history.CellPainting += PaintGraph;
        }

        private DataGridViewColumn Column(string name, int width)
        {
            var column = new DataGridViewTextBoxColumn { HeaderText = name, Width = width, MinimumWidth = Math.Min(width, 100),
                SortMode = DataGridViewColumnSortMode.NotSortable };
            history.Columns.Add(column);
            return column;
        }

        private static string Description(Revision revision)
        {
            return revision.References.Length == 0 ? revision.Subject : "[" + String.Join(" · ", revision.References) + "]  " + revision.Subject;
        }

        private void PaintGraph(object sender, DataGridViewCellPaintingEventArgs args)
        {
            if (args.RowIndex >= 0 && args.ColumnIndex == 0) CommitGraph.Paint(args, rows[args.RowIndex], 18);
        }

        private void Add(FlowLayoutPanel panel, string label, HistoryAction action)
        {
            var button = DialogLayout.Button(label);
            button.Click += (sender, args) => Complete(action);
            buttons.Add(action, button);
            panel.Controls.Add(button);
        }

        private void ShowSelection()
        {
            var selection = history.SelectedRows.Cast<DataGridViewRow>().OrderBy(row => row.Index).Select(row => revisions[row.Index]).ToArray();
            foreach (var action in buttons)
                action.Value.Enabled = action.Key == HistoryAction.SwitchBranch || action.Key == HistoryAction.MergeBranch
                    ? SelectedBranch != null && SelectedBranch != branch
                    : selection.Length == 1 ||
                    (selection.Length == 2 && (action.Key == HistoryAction.Compare || action.Key == HistoryAction.Merge));
            if (selection.Length == 0) { details.Text = "请选择一个提交查看详情。"; return; }
            var revision = selection[selection.Length - 1];
            if (selection.Length == 2)
            {
                details.Lines = new[] { "基准   " + revision.Id.Substring(0, 8) + "   " + Description(revision),
                    "目标   " + selection[0].Id.Substring(0, 8) + "   " + Description(selection[0]) };
                return;
            }
            details.Lines = new[] { revision.Message.Replace("\r\n", "\n").Replace("\n", Environment.NewLine), revision.Author + "   ·   " + revision.Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss zzz"),
                "提交    " + revision.Id, "父提交    " + (revision.Parents.Length == 0 ? "无" : String.Join(" · ", revision.Parents.Select(parent => parent.Substring(0, 8)))) };
        }

        private void Complete(HistoryAction action)
        {
            Button button;
            if (!buttons.TryGetValue(action, out button) || !button.Enabled) return;
            var selected = history.SelectedRows.Cast<DataGridViewRow>().Select(row => row.Index).OrderBy(index => index).ToArray();
            Basis = selected.Length > 0 ? revisions[selected[selected.Length - 1]] : null;
            Other = selected.Length == 2 ? revisions[selected[0]] : null;
            Action = action;
            DialogResult = DialogResult.OK;
        }
    }

    internal static class Dialogs
    {
        internal static Form BranchForm(string name, string current, string[] branches, HistoryAction action, out ComboBox field)
        {
            var merge = action == HistoryAction.MergeBranch;
            var delete = action == HistoryAction.DeleteBranch;
            var caption = delete ? "删除分支" : (merge ? "合并分支" : "切换分支");
            var dialog = new Form { Text = "WordGit · " + caption, StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable, MaximizeBox = false, MinimizeBox = false,
                ClientSize = new Size(510, 245), MinimumSize = new Size(440, 275),
                Font = SystemFonts.MessageBoxFont, AutoScaleMode = AutoScaleMode.Font };
            var layout = DialogLayout.Stack(6);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            dialog.Controls.Add(layout);
            var heading = DialogLayout.Label(name + "   ·   当前分支：" + current);
            heading.AutoEllipsis = true;
            heading.AutoSize = false;
            heading.Dock = DockStyle.Fill;
            heading.Height = dialog.Font.Height + 10;
            layout.Controls.Add(heading, 0, 0);
            var label = delete ? "已合入当前分支的分支" : (merge ? "来源分支" : "目标分支");
            layout.Controls.Add(DialogLayout.Label(label), 0, 1);
            field = new ComboBox { AccessibleName = label, Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
            field.Items.AddRange(branches);
            field.SelectedIndex = 0;
            layout.Controls.Add(field, 0, 2);
            var hint = delete ? "删除分支后，已合入的提交仍保留。"
                : (merge ? "将所选分支合入当前分支。\n请先在 Word 中审阅修订稿。"
                    : "切换前会检查当前稿。\n如有未提交的修改，可先提交。");
            layout.Controls.Add(DialogLayout.Label(hint), 0, 3);
            var actions = DialogLayout.Buttons(true);
            var cancel = DialogLayout.Button("取消");
            cancel.DialogResult = DialogResult.Cancel;
            var accept = DialogLayout.Button(delete ? "下一步" : (merge ? "生成修订稿" : "切换"));
            accept.DialogResult = DialogResult.OK;
            actions.Controls.Add(cancel);
            actions.Controls.Add(accept);
            layout.Controls.Add(actions, 0, 5);
            dialog.CancelButton = cancel;
            dialog.AcceptButton = accept;
            dialog.ActiveControl = field;
            return dialog;
        }

        public static string ChooseBranch(IWin32Window owner, string name, string current, string[] branches, HistoryAction action)
        {
            ComboBox field;
            using (var dialog = BranchForm(name, current, branches, action, out field))
                return dialog.ShowDialog(owner) == DialogResult.OK ? (string)field.SelectedItem : null;
        }

        internal static Form CommitForm(string title, string initial, out TextBox field, bool branchName = false)
        {
            var dialog = new Form { Text = branchName ? "WordGit · 新建分支" : "WordGit · 提交版本", StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable, MaximizeBox = false, MinimizeBox = false,
                ClientSize = new Size(570, 255), MinimumSize = new Size(440, 260), Font = SystemFonts.MessageBoxFont, AutoScaleMode = AutoScaleMode.Font };
            var layout = DialogLayout.Stack(5);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            dialog.Controls.Add(layout);
            var heading = DialogLayout.Label(title);
            heading.AutoEllipsis = true;
            heading.AutoSize = false;
            heading.Dock = DockStyle.Fill;
            heading.Height = dialog.Font.Height + 8;
            layout.Controls.Add(heading, 0, 0);
            var caption = branchName ? "分支名称" : "提交说明";
            layout.Controls.Add(DialogLayout.Label(caption), 0, 1);
            var input = new TextBox { Text = initial, Dock = branchName ? DockStyle.Top : DockStyle.Fill, Multiline = !branchName,
                ScrollBars = branchName ? ScrollBars.None : ScrollBars.Vertical,
                MaxLength = branchName ? 150 : 1000, AcceptsReturn = !branchName, AccessibleName = caption };
            layout.Controls.Add(input, 0, 2);
            layout.Controls.Add(DialogLayout.Label(branchName ? "例如 review/导师意见。新建后可用「切换分支」打开。"
                : "首行用作提交标题。Ctrl + Enter 提交，Esc 取消。"), 0, 3);
            var actions = DialogLayout.Buttons(true);
            var cancel = DialogLayout.Button("取消");
            cancel.DialogResult = DialogResult.Cancel;
            var commit = DialogLayout.Button(branchName ? "创建" : "提交");
            commit.DialogResult = DialogResult.OK;
            commit.Enabled = !String.IsNullOrWhiteSpace(initial);
            actions.Controls.Add(cancel);
            actions.Controls.Add(commit);
            layout.Controls.Add(actions, 0, 4);
            input.TextChanged += (sender, args) => commit.Enabled = !String.IsNullOrWhiteSpace(input.Text);
            input.KeyDown += (sender, args) => {
                if (args.Control && args.KeyCode == Keys.Enter && commit.Enabled)
                { args.SuppressKeyPress = true; dialog.DialogResult = DialogResult.OK; }
            };
            dialog.CancelButton = cancel;
            if (branchName) dialog.AcceptButton = commit;
            dialog.ActiveControl = input;
            field = input;
            return dialog;
        }

        public static string Ask(IWin32Window owner, string title, string initial, bool branchName = false)
        {
            TextBox field;
            using (var dialog = CommitForm(title, initial, out field, branchName))
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK) return null;
                var message = field.Text.Trim().Replace("\r\n", "\n");
                var newline = message.IndexOf('\n');
                return newline < 0 ? message : message.Substring(0, newline).TrimEnd() + "\n\n" + message.Substring(newline + 1).TrimStart();
            }
        }

        public static void ShowBackgroundOperation(IWin32Window owner, string title, Func<Task<string>> operation, Action finished = null,
            string initialStatus = null)
        {
            ShowBackgroundOperation(owner, title, cancellation => operation(), false, finished, initialStatus);
        }

        public static void ShowBackgroundOperation(IWin32Window owner, string title, Func<System.Threading.CancellationToken, Task<string>> operation,
            Action finished = null, string initialStatus = null)
        {
            ShowBackgroundOperation(owner, title, operation, true, finished, initialStatus);
        }

        private static void ShowBackgroundOperation(IWin32Window owner, string title,
            Func<System.Threading.CancellationToken, Task<string>> operation, bool cancellable, Action finished, string initialStatus)
        {
            var dialog = new Form
            {
                Text = "WordGit · " + title,
                StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
                ControlBox = false, ShowInTaskbar = false, ClientSize = new Size(460, 195),
                Font = SystemFonts.MessageBoxFont, AutoScaleMode = AutoScaleMode.Font
            };
            var cancellation = new System.Threading.CancellationTokenSource();
            var layout = DialogLayout.Stack(3);
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            dialog.Controls.Add(layout);
            var status = new TextBox
            {
                Text = initialStatus ?? "正在后台处理，可继续编辑正文。", Multiline = true, ReadOnly = true, BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill, BackColor = SystemColors.Control, ForeColor = SystemColors.ControlText,
                ScrollBars = ScrollBars.Vertical, TabStop = false
            };
            layout.Controls.Add(status, 0, 0);
            var progress = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee, Height = 14 };
            layout.Controls.Add(progress, 0, 1);
            var actions = DialogLayout.Buttons(true);
            var close = DialogLayout.Button("关闭");
            close.Enabled = false;
            close.Click += (sender, args) => dialog.Close();
            var cancel = DialogLayout.Button("取消");
            cancel.Click += (sender, args) =>
            {
                cancel.Enabled = false;
                status.Text = "正在取消…";
                cancellation.Cancel();
            };
            actions.Controls.Add(close);
            if (cancellable) actions.Controls.Add(cancel);
            layout.Controls.Add(actions, 0, 2);
            dialog.Shown += async (sender, args) =>
            {
                UiThread.EnsureContext();
                try
                {
                    var result = await operation(cancellation.Token);
                    if (dialog.IsDisposed) return;
                    status.Text = result;
                }
                catch (OperationCanceledException)
                {
                    if (!dialog.IsDisposed) status.Text = "已取消。未完成的部分不会影响文档与历史。";
                }
                catch (Exception error)
                {
                    AppPaths.LogError(title + ": " + error);
                    if (!dialog.IsDisposed) status.Text = "操作未完成。" + Environment.NewLine + error.Message;
                }
                finally
                {
                    if (!dialog.IsDisposed)
                    {
                        progress.Visible = false;
                        cancel.Enabled = false;
                        close.Enabled = true;
                    }
                    if (finished != null)
                        try { finished(); }
                        catch (Exception error) { AppPaths.LogError(title + " (finish): " + error); }
                }
            };
            dialog.FormClosed += (sender, args) => { cancellation.Cancel(); cancellation.Dispose(); };
            dialog.Show(owner);
        }
    }
}
