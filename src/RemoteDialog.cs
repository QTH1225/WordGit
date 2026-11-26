using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class RemoteOptions
    {
        internal string Url, Branch, DocumentPath;
    }

    internal sealed class RemoteDialog : Form
    {
        internal readonly TextBox Url = new TextBox { Dock = DockStyle.Fill, AccessibleName = "GitHub 仓库地址" };
        internal readonly TextBox Branch = new TextBox { Dock = DockStyle.Fill, AccessibleName = "远程分支" };
        internal readonly TextBox DocumentPath = new TextBox { Dock = DockStyle.Fill, AccessibleName = "克隆文档路径" };
        internal RemoteOptions Options { get; private set; }

        internal RemoteDialog(bool clone, string url = null, string branch = "main")
        {
            Text = clone ? "WordGit · 克隆远程稿件" : "WordGit · 关联 GitHub 仓库";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(620, 365);
            MinimumSize = new Size(560, 395);
            MinimizeBox = false;
            MaximizeBox = false;
            var layout = DialogLayout.Stack(8);
            for (var row = 0; row < 8; row++) layout.RowStyles.Add(new RowStyle(row == 6 ? SizeType.Percent : SizeType.AutoSize, row == 6 ? 100 : 0));
            Controls.Add(layout);
            layout.Controls.Add(DialogLayout.Label("一份稿件关联一个 GitHub 仓库，使用 WordGit 分块历史。"), 0, 0);
            layout.Controls.Add(DialogLayout.Label("GitHub 仓库地址（HTTPS / SSH）"), 0, 1);
            Url.Text = url ?? "";
            Url.ReadOnly = false;
            layout.Controls.Add(Url, 0, 2);
            layout.Controls.Add(DialogLayout.Label("远程分支"), 0, 3);
            Branch.Text = branch;
            layout.Controls.Add(Branch, 0, 4);
            var fileRow = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2, Visible = clone };
            fileRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fileRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var choose = DialogLayout.Button("保存位置…");
            choose.Margin = new Padding(8, 0, 0, 0);
            choose.Click += (sender, args) => {
                using (var save = new SaveFileDialog { Filter = "Word 文档 (*.docx)|*.docx", FileName = "协作稿件.docx" })
                    if (save.ShowDialog(this) == DialogResult.OK) DocumentPath.Text = save.FileName;
            };
            fileRow.Controls.Add(DocumentPath, 0, 0);
            fileRow.Controls.Add(choose, 1, 0);
            layout.Controls.Add(fileRow, 0, 5);
            var note = new TextBox { Multiline = true, ReadOnly = true, BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Control, ForeColor = SystemColors.ControlText, Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical, MinimumSize = new Size(0, 85), TabStop = false,
                Text = clone ? "从对方已推送的 WordGit 仓库下载历史，并还原为新的本地 DOCX。不会替换同名文档或已有历史。\r\n\r\n首次 HTTPS 连接可能由 Git for Windows 打开浏览器登录；SSH 需先配置密钥和主机信任。"
                    : "首次上传请选择 GitHub 上的空仓库。关联后先获取远程历史，推送由你另外点击。更换地址时旧远程历史会归档保留。\r\n\r\n首次 HTTPS 连接可能由 Git for Windows 打开浏览器登录。获取不会覆盖当前稿，双方修改通过 Word 修订合并。" };
            layout.Controls.Add(note, 0, 6);
            var actions = DialogLayout.Buttons(true);
            var cancel = DialogLayout.Button("取消");
            cancel.DialogResult = DialogResult.Cancel;
            var accept = DialogLayout.Button(clone ? "克隆稿件" : "关联并获取");
            accept.Enabled = false;
            EventHandler validate = (sender, args) => accept.Enabled = !String.IsNullOrWhiteSpace(Url.Text) &&
                !String.IsNullOrWhiteSpace(Branch.Text) && (!clone || !String.IsNullOrWhiteSpace(DocumentPath.Text));
            Url.TextChanged += validate;
            Branch.TextChanged += validate;
            DocumentPath.TextChanged += validate;
            accept.Click += (sender, args) => {
                try {
                    var normalized = Url.Text.Trim();
                    var path = clone ? Path.GetFullPath(DocumentPath.Text.Trim()) : null;
                    if (clone && (File.Exists(path) || Directory.Exists(path) || Path.GetExtension(path).ToLowerInvariant() != ".docx"))
                        throw new IOException("请选择一个尚不存在的 .docx 文件名，已有文件不会被覆盖。");
                    Options = new RemoteOptions { Url = normalized, Branch = Branch.Text.Trim(), DocumentPath = path };
                    DialogResult = DialogResult.OK;
                } catch (Exception error) { MessageBox.Show(this, error.Message, "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            actions.Controls.Add(cancel);
            actions.Controls.Add(accept);
            layout.Controls.Add(actions, 0, 7);
            AcceptButton = accept;
            CancelButton = cancel;
            validate(this, EventArgs.Empty);
        }
    }
}
