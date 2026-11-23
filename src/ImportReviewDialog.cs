using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class ImportReviewOptions
    {
        internal string FilePath, Branch, Reviewer;
        internal bool CreateMergeDraft;
    }

    internal sealed class ImportReviewDialog : Form
    {
        internal readonly TextBox FilePath = new TextBox { Dock = DockStyle.Fill, AccessibleName = "老师返稿文件" };
        internal readonly TextBox BranchName = new TextBox { Dock = DockStyle.Fill, AccessibleName = "返稿分支名" };
        internal readonly TextBox Reviewer = new TextBox { Dock = DockStyle.Fill, Text = "老师返稿", AccessibleName = "返稿来源" };
        internal readonly CheckBox SaveOriginal = new CheckBox { AutoSize = true, Checked = true,
            Text = "将原始返稿保存为独立分支", Margin = new Padding(0, 10, 0, 8) };
        internal ImportReviewOptions Options { get; private set; }
        private readonly string originalPath;
        private readonly Button merge = DialogLayout.Button("生成合并稿");
        private readonly Button saveOnly = DialogLayout.Button("仅保存为分支");

        internal ImportReviewDialog(string originalPath, Revision basis, string currentBranch, string[] branches)
        {
            this.originalPath = Path.GetFullPath(originalPath);
            Text = "WordGit · 返稿合并";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(600, 405);
            MinimumSize = new Size(550, 435);
            MinimizeBox = false;
            MaximizeBox = false;
            var outer = DialogLayout.Stack(2);
            outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(outer);

            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Margin = new Padding(0) };
            outer.Controls.Add(scroll, 0, 0);
            var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 6 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var row = 0; row < 6; row++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            scroll.Controls.Add(layout);
            var subject = basis.Subject.Length > 60 ? basis.Subject.Substring(0, 57) + "…" : basis.Subject;
            var branch = currentBranch.Length > 40 ? currentBranch.Substring(0, 37) + "…" : currentBranch;
            var caption = DialogLayout.Label("送审基准  " + basis.Id.Substring(0, 8) + " · " + subject +
                Environment.NewLine + "当前分支  " + branch);
            caption.AccessibleDescription = basis.Message + Environment.NewLine + currentBranch;
            layout.Controls.Add(caption, 0, 0);
            var fileRow = Field("返稿文件", FilePath);
            fileRow.ColumnCount = 3;
            fileRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var browse = DialogLayout.Button("选择文件…");
            browse.Margin = new Padding(8, 0, 0, 0);
            browse.Click += (sender, args) => {
                using (var picker = new OpenFileDialog { Filter = "Word 审阅稿 (*.docx)|*.docx", Title = "选择老师返回的审阅批注稿" })
                    if (picker.ShowDialog(this) == DialogResult.OK) FilePath.Text = picker.FileName;
            };
            fileRow.Controls.Add(browse, 2, 0);
            layout.Controls.Add(fileRow, 0, 1);
            var defaultBranch = "review-" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
            BranchName.Text = defaultBranch;
            for (var suffix = 2; GitHistory.ReviewBranchConflict(BranchName.Text, branches) != null; suffix++)
                BranchName.Text = defaultBranch + "-" + suffix;
            layout.Controls.Add(SaveOriginal, 0, 2);
            var branchRow = Field("返稿分支", BranchName);
            var reviewerRow = Field("返稿来源", Reviewer);
            layout.Controls.Add(branchRow, 0, 3);
            layout.Controls.Add(reviewerRow, 0, 4);
            var note = DialogLayout.Label("生成合并稿前，请先提交当前稿的修改。");
            layout.Controls.Add(note, 0, 5);
            Action wrap = () => {
                var width = Math.Max(120, scroll.ClientSize.Width - SystemInformation.VerticalScrollBarWidth);
                caption.MaximumSize = new Size(width, 0);
                note.MaximumSize = new Size(width, 0);
                SaveOriginal.MaximumSize = new Size(width, 0);
            };
            scroll.SizeChanged += (sender, args) => wrap();
            wrap();
            var buttons = DialogLayout.Buttons(true);
            var cancel = DialogLayout.Button("取消");
            cancel.DialogResult = DialogResult.Cancel;
            buttons.WrapContents = true;
            Action<bool> complete = createMergeDraft => {
                try {
                    var path = Path.GetFullPath(FilePath.Text.Trim());
                    if (!File.Exists(path) || !String.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("请选择已经保存的 .docx 返稿。");
                    if (String.Equals(path, this.originalPath, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("请选择收到的返稿文件。");
                    var branchName = SaveOriginal.Checked ? BranchName.Text.Trim() : null;
                    var conflict = branchName == null ? null : GitHistory.ReviewBranchConflict(branchName, branches);
                    if (conflict != null)
                    {
                        MessageBox.Show(this, conflict, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        BranchName.Focus();
                        BranchName.SelectAll();
                        return;
                    }
                    Options = new ImportReviewOptions { FilePath = path, Branch = branchName,
                        Reviewer = branchName == null ? null : Reviewer.Text.Trim(), CreateMergeDraft = createMergeDraft };
                    DialogResult = DialogResult.OK;
                }
                catch (Exception error) { MessageBox.Show(this, error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            merge.Click += (sender, args) => complete(true);
            saveOnly.Click += (sender, args) => complete(false);
            EventHandler validate = (sender, args) => {
                var keepBranch = SaveOriginal.Checked;
                var hasFile = File.Exists(FilePath.Text.Trim());
                var branchReady = !String.IsNullOrWhiteSpace(BranchName.Text) && !String.IsNullOrWhiteSpace(Reviewer.Text);
                branchRow.Visible = keepBranch;
                reviewerRow.Visible = keepBranch;
                saveOnly.Visible = keepBranch;
                saveOnly.Enabled = hasFile && branchReady;
                merge.Enabled = hasFile && (!keepBranch || branchReady);
            };
            FilePath.TextChanged += validate;
            BranchName.TextChanged += validate;
            Reviewer.TextChanged += validate;
            SaveOriginal.CheckedChanged += validate;
            buttons.Controls.Add(merge);
            buttons.Controls.Add(saveOnly);
            buttons.Controls.Add(cancel);
            outer.Controls.Add(buttons, 0, 1);
            AcceptButton = merge;
            CancelButton = cancel;
            validate(this, EventArgs.Empty);
        }

        private TableLayoutPanel Field(string label, Control value)
        {
            var row = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0, 7, 0, 7) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row.Controls.Add(DialogLayout.Label(label), 0, 0);
            row.Controls.Add(value, 1, 0);
            return row;
        }
    }
}
