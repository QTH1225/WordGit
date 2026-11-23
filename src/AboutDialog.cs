using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class AboutDialog : Form
    {
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private readonly Func<CancellationToken, Task<UpdateResult>> check;
        private readonly Button checkButton = DialogLayout.Button("检查更新");
        private readonly TextBox updateStatus = new TextBox
        {
            Name = "updateStatus", AccessibleName = "更新状态", Multiline = true, ReadOnly = true,
            Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle,
            BackColor = SystemColors.Window, ForeColor = SystemColors.WindowText, TabStop = false,
            Text = "点击“检查更新”查询 GitHub 上的正式发布版本。"
        };
        private bool checking;

        internal AboutDialog(Func<CancellationToken, Task<UpdateResult>> check = null)
        {
            this.check = check ?? (token => UpdateChecker.CheckAsync(AppInfo.Version, token));
            Text = "关于 WordGit";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(560, 480);
            MinimumSize = new Size(500, 440);
            ShowInTaskbar = false;
            MaximizeBox = false;
            var outer = DialogLayout.Stack(2);
            outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(outer);
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Margin = new Padding(0) };
            outer.Controls.Add(scroll, 0, 0);
            var layout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 7, Margin = new Padding(0) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var row = 0; row < 7; row++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            scroll.Controls.Add(layout);
            var title = DialogLayout.Label("WordGit");
            title.Font = new Font(Font.FontFamily, 19, FontStyle.Bold);
            layout.Controls.Add(title, 0, 0);
            layout.Controls.Add(DialogLayout.Label("当前版本  " + AppInfo.DisplayVersion), 0, 1);
            layout.Controls.Add(Link("作者  " + AppInfo.Author, AppInfo.AuthorUrl), 0, 2);
            var project = Link(AppInfo.ProjectUrl, AppInfo.ProjectUrl);
            layout.Controls.Add(project, 0, 3);
            var links = DialogLayout.Buttons(false);
            links.Margin = new Padding(0, 2, 0, 12);
            links.WrapContents = true;
            links.Controls.Add(Link("发布记录", AppInfo.ReleasesUrl));
            links.Controls.Add(Link("反馈问题", AppInfo.IssuesUrl));
            links.Controls.Add(Link("许可证  GPL-3.0-only", AppInfo.LicenseUrl));
            layout.Controls.Add(links, 0, 4);
            updateStatus.Dock = DockStyle.Top;
            updateStatus.Height = Math.Max(80, Font.Height * 4 + 12);
            updateStatus.Margin = new Padding(0, 0, 0, 12);
            layout.Controls.Add(updateStatus, 0, 5);
            var note = DialogLayout.Label("© 2026 QTH1225 · GPLv3" + Environment.NewLine +
                "可按许可证复制、修改与分发，不提供担保。");
            layout.Controls.Add(note, 0, 6);
            Action wrap = () => {
                var width = Math.Max(120, scroll.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 4);
                note.MaximumSize = new Size(width, 0);
                project.MaximumSize = new Size(width, 0);
                updateStatus.Height = Math.Max(80, Font.Height * 4 + 12);
            };
            scroll.SizeChanged += (sender, args) => wrap();
            FontChanged += (sender, args) => wrap();
            wrap();
            var actions = DialogLayout.Buttons(true);
            actions.WrapContents = true;
            var close = DialogLayout.Button("关闭");
            close.Click += (sender, args) => Close();
            CancelButton = close;
            var releases = DialogLayout.Button("打开发布页");
            releases.Click += (sender, args) => AppInfo.OpenLink(this, AppInfo.ReleasesUrl);
            checkButton.Click += async (sender, args) => await CheckForUpdatesAsync();
            actions.Controls.Add(close);
            actions.Controls.Add(releases);
            actions.Controls.Add(checkButton);
            outer.Controls.Add(actions, 0, 1);
        }

        private LinkLabel Link(string text, string url)
        {
            var link = new LinkLabel { Text = text, AutoSize = true, Margin = new Padding(0, 4, 14, 7),
                AccessibleName = text, MaximumSize = new Size(460, 0), Tag = url };
            link.LinkClicked += (sender, args) => AppInfo.OpenLink(this, url);
            return link;
        }

        internal async Task CheckForUpdatesAsync()
        {
            if (checking || IsDisposed || Disposing) return;
            checking = true;
            checkButton.Enabled = false;
            updateStatus.Text = "正在检查 GitHub 正式发布版本…";
            try
            {
                var result = await check(cancellation.Token);
                if (!IsDisposed && !Disposing) updateStatus.Text = result.Message;
            }
            catch (OperationCanceledException) { }
            catch (Exception) { if (!IsDisposed && !Disposing) updateStatus.Text = "检查更新失败，请稍后重试，或打开发布页。"; }
            finally
            {
                checking = false;
                if (!IsDisposed && !Disposing) checkButton.Enabled = true;
            }
        }

        internal static void ShowWindow(IWin32Window owner, bool checkUpdates)
        {
            var dialog = Application.OpenForms.OfType<AboutDialog>().FirstOrDefault(form => !form.IsDisposed);
            if (dialog == null) { dialog = new AboutDialog(); DialogPlacement.Show(dialog, owner); }
            else { if (dialog.WindowState == FormWindowState.Minimized) dialog.WindowState = FormWindowState.Normal; DialogPlacement.Center(dialog, owner); dialog.Activate(); }
            if (checkUpdates) dialog.BeginInvoke(new Action(async () => await dialog.CheckForUpdatesAsync()));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !IsDisposed) { cancellation.Cancel(); cancellation.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
