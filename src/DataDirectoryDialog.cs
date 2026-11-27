// SPDX-License-Identifier: GPL-3.0-only
using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class DataDirectoryDialog : Form
    {
        private readonly Font folderFont = new Font("Segoe MDL2 Assets", 12);
        private readonly ToolTip hints = new ToolTip();
        internal string SelectedDirectory { get; private set; }

        internal DataDirectoryDialog(string current, bool canMove)
        {
            Text = "WordGit · 数据目录";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(580, 116);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;

            var layout = DialogLayout.Stack(3);
            layout.Dock = DockStyle.Top;
            layout.AutoSize = true;
            layout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var row = 0; row < layout.RowCount; row++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var directory = new TextBox
            {
                AccessibleName = "数据目录", Text = current, ReadOnly = !canMove, BackColor = SystemColors.Window,
                Anchor = AnchorStyles.Left | AnchorStyles.Right, Margin = new Padding(0)
            };
            var browse = new Button
            {
                Text = "\uED25", Font = folderFont, AccessibleName = "选择文件夹", Enabled = canMove,
                Size = new Size(33, directory.PreferredHeight + 6), Margin = new Padding(6, 0, 0, 0)
            };
            hints.SetToolTip(browse, "选择文件夹");
            browse.Click += (sender, args) =>
            {
                using (var picker = new FolderBrowserDialog
                {
                    Description = "选择 WordGit 数据目录",
                    SelectedPath = current
                })
                    if (picker.ShowDialog(this) == DialogResult.OK) directory.Text = picker.SelectedPath;
            };

            var folder = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, Margin = new Padding(0) };
            folder.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            folder.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            folder.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            folder.Controls.Add(directory, 0, 0);
            folder.Controls.Add(browse, 1, 0);
            layout.Controls.Add(folder, 0, 0);
            layout.Controls.Add(DialogLayout.Label(canMove ? "迁移时会保留已有提交历史。"
                : "当前操作完成后即可更改目录。"), 0, 1);

            var buttons = DialogLayout.Buttons(true);
            var close = DialogLayout.Button("关闭");
            close.DialogResult = DialogResult.Cancel;
            var move = DialogLayout.Button("迁移历史");
            move.Enabled = false;
            directory.TextChanged += (sender, args) =>
            {
                var target = directory.Text.Trim().Trim('"').TrimEnd('\\', '/');
                move.Enabled = canMove && target.Length != 0 &&
                    !String.Equals(target, current.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
            };
            move.Click += (sender, args) =>
            {
                SelectedDirectory = directory.Text.Trim().Trim('"');
                DialogResult = DialogResult.OK;
            };
            buttons.Controls.Add(move);
            buttons.Controls.Add(close);
            layout.Controls.Add(buttons, 0, 2);
            layout.SizeChanged += (sender, args) =>
            {
                if (ClientSize.Height != layout.Height) ClientSize = new Size(ClientSize.Width, layout.Height);
            };
            Controls.Add(layout);
            CancelButton = close;
            ActiveControl = close;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                hints.Dispose();
                folderFont.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
