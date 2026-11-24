using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class ImportReviewOptions
    {
        internal string FilePath, Branch, Reviewer;
        internal Revision Basis;
        internal bool CreateMergeDraft;
    }

    internal sealed class ImportReviewDialog : Form
    {
        internal readonly TextBox FilePath = new TextBox { Dock = DockStyle.Fill, AccessibleName = "返稿文件" };
        internal readonly ComboBox BasisPicker = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList,
            AccessibleName = "送审版本" };
        internal readonly TextBox BranchName = new TextBox { Dock = DockStyle.Fill, AccessibleName = "返稿分支名" };
        internal readonly TextBox Reviewer = new TextBox { Dock = DockStyle.Fill, Text = "老师返稿", AccessibleName = "返稿来源" };
        internal readonly CheckBox SaveOriginal = new CheckBox { AutoSize = true, Checked = true,
            Text = "将原始返稿保存为独立分支", Margin = new Padding(0, 10, 0, 8) };
        internal ImportReviewOptions Options { get; private set; }
        internal bool ChooseHistoricalMerge { get; private set; }
        private readonly string originalPath;
        private readonly Button merge = DialogLayout.Button("生成合并稿");
        private readonly Button saveOnly = DialogLayout.Button("仅保存为分支");

        internal ImportReviewDialog(string originalPath, List<Revision> revisions, string currentBranch, string[] branches, bool supportsTrackedMerge = true)
        {
            this.originalPath = Path.GetFullPath(originalPath);
            Text = supportsTrackedMerge ? "WordGit · 返稿合并" : "WordGit · 保存返稿";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(660, 478);
            MinimumSize = new Size(550, 455);
            MinimizeBox = false;
            MaximizeBox = false;
            var outer = DialogLayout.Stack(3);
            outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(outer);

            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Margin = new Padding(0) };
            outer.Controls.Add(scroll, 0, 0);
            var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 7 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var row = 0; row < 7; row++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            scroll.Controls.Add(layout);
            var fileRow = Field("返稿文件", FilePath);
            fileRow.ColumnCount = 3;
            fileRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var browse = DialogLayout.Button("选择文件…");
            browse.Margin = new Padding(8, 0, 0, 0);
            browse.Click += (sender, args) => {
                using (var picker = new OpenFileDialog { Filter = "Word 审阅稿 (*.docx)|*.docx", Title = "选择收到的 Word 返稿" })
                    if (picker.ShowDialog(this) == DialogResult.OK) FilePath.Text = picker.FileName;
            };
            fileRow.Controls.Add(browse, 2, 0);
            layout.Controls.Add(fileRow, 0, 0);
            BasisPicker.Items.Add("请选择实际发出的送审版本");
            foreach (var revision in revisions)
                BasisPicker.Items.Add(revision.Id.Substring(0, 8) + " · " +
                    revision.Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm") + " · " + revision.Subject);
            BasisPicker.SelectedIndex = 0;
            layout.Controls.Add(Field("送审版本", BasisPicker), 0, 1);
            var target = new TextBox { Text = currentBranch, Dock = DockStyle.Fill, ReadOnly = true, AccessibleName = "合入分支" };
            if (supportsTrackedMerge) layout.Controls.Add(Field("合入分支", target), 0, 2);
            var defaultBranch = "review-" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
            BranchName.Text = defaultBranch;
            for (var suffix = 2; GitHistory.ReviewBranchConflict(BranchName.Text, branches) != null; suffix++)
                BranchName.Text = defaultBranch + "-" + suffix;
            if (supportsTrackedMerge) layout.Controls.Add(SaveOriginal, 0, 3);
            var branchRow = Field("返稿分支", BranchName);
            var reviewerRow = Field("返稿来源", Reviewer);
            layout.Controls.Add(branchRow, 0, 4);
            layout.Controls.Add(reviewerRow, 0, 5);
            var note = DialogLayout.Label(supportsTrackedMerge ? "当前稿需要提交时，会先请你提交，再使用这里选好的返稿和送审版本继续。"
                : "当前 WPS 文字支持将原始返稿保存为分支，可在提交图中预览或对比。自动生成修订合并稿暂不可用。");
            layout.Controls.Add(note, 0, 6);
            var historical = new LinkLabel { Text = "合并两个历史版本…", AutoSize = true, Margin = new Padding(0, 2, 0, 0),
                Enabled = revisions.Count >= 2 };
            historical.Click += (sender, args) => {
                ChooseHistoricalMerge = true;
                DialogResult = DialogResult.OK;
            };
            if (supportsTrackedMerge) outer.Controls.Add(historical, 0, 1);
            Action wrap = () => {
                var width = Math.Max(120, scroll.ClientSize.Width - SystemInformation.VerticalScrollBarWidth);
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
                    if (BasisPicker.SelectedIndex <= 0) throw new InvalidOperationException("请选择实际发出的送审版本。");
                    var path = Path.GetFullPath(FilePath.Text.Trim());
                    if (!File.Exists(path) || !String.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("请选择已经保存的 .docx 返稿。");
                    if (String.Equals(path, this.originalPath, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("请选择收到的返稿文件。");
                    var branchName = SaveOriginal.Checked ? BranchName.Text.Trim() : null;
                    Options = new ImportReviewOptions { FilePath = path, Branch = branchName, Basis = revisions[BasisPicker.SelectedIndex - 1],
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
                var ready = hasFile && BasisPicker.SelectedIndex > 0;
                saveOnly.Enabled = ready && branchReady;
                merge.Enabled = ready && (!keepBranch || branchReady);
            };
            BasisPicker.SelectedIndexChanged += validate;
            FilePath.TextChanged += validate;
            BranchName.TextChanged += validate;
            Reviewer.TextChanged += validate;
            SaveOriginal.CheckedChanged += validate;
            if (supportsTrackedMerge) buttons.Controls.Add(merge);
            buttons.Controls.Add(saveOnly);
            buttons.Controls.Add(cancel);
            outer.Controls.Add(buttons, 0, 2);
            AcceptButton = supportsTrackedMerge ? merge : saveOnly;
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
