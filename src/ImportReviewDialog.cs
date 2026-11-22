using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class ImportReviewOptions
    {
        internal string FilePath, Branch, Reviewer;
        internal bool MergeAfterImport;
    }

    internal sealed class ImportReviewDialog : Form
    {
        internal readonly TextBox FilePath = new TextBox { Dock = DockStyle.Fill, AccessibleName = "老师返稿文件" };
        internal readonly TextBox BranchName = new TextBox { Dock = DockStyle.Fill, AccessibleName = "返稿分支名" };
        internal readonly TextBox Reviewer = new TextBox { Dock = DockStyle.Fill, Text = "老师返稿", AccessibleName = "返稿来源" };
        internal readonly CheckBox MergeAfterImport = new CheckBox { AutoSize = true, Checked = true,
            Text = "导入后与当前分支生成 Word 修订合并稿", Margin = new Padding(0, 10, 0, 8) };
        internal ImportReviewOptions Options { get; private set; }
        private readonly string originalPath;
        private readonly Button import = DialogLayout.Button("导入返稿");

        internal ImportReviewDialog(string originalPath, Revision basis, string currentBranch)
        {
            this.originalPath = Path.GetFullPath(originalPath);
            Text = "WordGit · 导入审阅稿为分支";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(600, 405);
            MinimumSize = new Size(550, 435);
            MinimizeBox = false;
            MaximizeBox = false;
            var layout = DialogLayout.Stack(7);
            for (var i = 0; i < 7; i++) layout.RowStyles.Add(new RowStyle(i == 5 ? SizeType.Percent : SizeType.AutoSize, i == 5 ? 100 : 0));
            Controls.Add(layout);
            var subject = basis.Subject.Length > 60 ? basis.Subject.Substring(0, 57) + "…" : basis.Subject;
            var branch = currentBranch.Length > 40 ? currentBranch.Substring(0, 37) + "…" : currentBranch;
            var caption = DialogLayout.Label("送审基准  " + basis.Id.Substring(0, 8) + " · " + subject +
                Environment.NewLine + "当前分支  " + branch);
            caption.AccessibleDescription = basis.Message + Environment.NewLine + currentBranch;
            caption.MaximumSize = new Size(480, 0);
            layout.Controls.Add(caption, 0, 0);
            var fileRow = Field("老师返稿", FilePath);
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
            BranchName.Text = "review/" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
            layout.Controls.Add(Field("返稿分支", BranchName), 0, 2);
            layout.Controls.Add(Field("返稿来源", Reviewer), 0, 3);
            layout.Controls.Add(MergeAfterImport, 0, 4);
            var note = new TextBox { Multiline = true, ReadOnly = true, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Control, ForeColor = SystemColors.ControlText, ScrollBars = ScrollBars.Vertical,
                MinimumSize = new Size(0, 95), TabStop = false,
                Text = "返稿按磁盘中已保存的原始文件入库，保留批注与修订。返稿来源只用于 Git 提交作者。\r\n\r\n导入不切换当前分支，也不覆盖当前稿。生成合并稿前，请先提交当前稿的修改；审阅确认后再提交合并结果。" };
            layout.Controls.Add(note, 0, 5);
            var buttons = DialogLayout.Buttons(true);
            var cancel = DialogLayout.Button("取消");
            cancel.DialogResult = DialogResult.Cancel;
            import.Enabled = false;
            import.Click += (sender, args) => {
                try {
                    var path = Path.GetFullPath(FilePath.Text.Trim());
                    if (!File.Exists(path) || !String.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("请选择已经保存的 .docx 返稿。");
                    if (String.Equals(path, this.originalPath, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("请选择独立的老师返稿文件，不能选择当前原稿。");
                    Options = new ImportReviewOptions { FilePath = path, Branch = BranchName.Text.Trim(),
                        Reviewer = Reviewer.Text.Trim(), MergeAfterImport = MergeAfterImport.Checked };
                    DialogResult = DialogResult.OK;
                }
                catch (Exception error) { MessageBox.Show(this, error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            EventHandler validate = (sender, args) => import.Enabled = File.Exists(FilePath.Text.Trim()) &&
                !String.IsNullOrWhiteSpace(BranchName.Text) && !String.IsNullOrWhiteSpace(Reviewer.Text);
            FilePath.TextChanged += validate;
            BranchName.TextChanged += validate;
            Reviewer.TextChanged += validate;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(import);
            layout.Controls.Add(buttons, 0, 6);
            AcceptButton = import;
            CancelButton = cancel;
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
