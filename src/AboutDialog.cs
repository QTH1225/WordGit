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
        private readonly Label updateStatus = new Label
        {
            Name = "updateStatus", AccessibleName = "更新状态", AutoSize = true, Dock = DockStyle.Fill,
            BackColor = SystemColors.Window, ForeColor = SystemColors.WindowText, Visible = false,
            Padding = new Padding(10, 7, 10, 7), Margin = new Padding(0, 8, 0, 0)
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
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ClientSize = new Size(470, 188);
            ShowInTaskbar = false;
            MinimizeBox = MaximizeBox = false;

            var layout = DialogLayout.Stack(5);
            layout.Dock = DockStyle.Top;
            layout.AutoSize = true;
            layout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            layout.Padding = new Padding(18, 14, 18, 12);
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var row = 0; row < 5; row++) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var header = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 0, 0, 8) };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var title = DialogLayout.Label("WordGit");
            title.Font = new Font(Font.FontFamily, 19, FontStyle.Bold);
            title.Margin = new Padding(0, 0, 14, 0);
            var version = DialogLayout.Label("版本 " + AppInfo.DisplayVersion);
            version.Anchor = AnchorStyles.Left;
            version.ForeColor = SystemColors.GrayText;
            version.Margin = new Padding(0);
            header.Controls.Add(title, 0, 0);
            header.Controls.Add(version, 1, 0);
            layout.Controls.Add(header, 0, 0);

            var metadata = DialogLayout.Buttons(false);
            metadata.Margin = new Padding(0);
            metadata.WrapContents = true;
            metadata.Controls.Add(Link("作者 " + AppInfo.Author, AppInfo.AuthorUrl));
            metadata.Controls.Add(Link("GPL-3.0-only", AppInfo.LicenseUrl));
            layout.Controls.Add(metadata, 0, 1);

            var links = DialogLayout.Buttons(false);
            links.Margin = new Padding(0, 2, 0, 0);
            links.WrapContents = true;
            links.Controls.Add(Link("GitHub 仓库", AppInfo.ProjectUrl));
            links.Controls.Add(Link("发布页", AppInfo.ReleasesUrl));
            links.Controls.Add(Link("反馈问题", AppInfo.IssuesUrl));
            layout.Controls.Add(links, 0, 2);
            layout.Controls.Add(updateStatus, 0, 3);

            var actions = DialogLayout.Buttons(true);
            actions.WrapContents = true;
            var close = DialogLayout.Button("关闭");
            close.Click += (sender, args) => Close();
            CancelButton = close;
            checkButton.Click += async (sender, args) => await CheckForUpdatesAsync();
            actions.Controls.Add(close);
            actions.Controls.Add(checkButton);
            layout.Controls.Add(actions, 0, 4);

            // Keep the window width stable and let its height follow the visible rows.
            Action wrap = () => updateStatus.MaximumSize = new Size(
                Math.Max(140, ClientSize.Width - layout.Padding.Horizontal), 0);
            ClientSizeChanged += (sender, args) => wrap();
            layout.SizeChanged += (sender, args) => {
                if (ClientSize.Height != layout.Height) ClientSize = new Size(ClientSize.Width, layout.Height);
            };
            wrap();
            Controls.Add(layout);
        }

        private LinkLabel Link(string text, string url)
        {
            var link = new LinkLabel { Text = text, AutoSize = true, Margin = new Padding(0, 3, 16, 5),
                AccessibleName = text, LinkBehavior = LinkBehavior.HoverUnderline, Tag = url };
            link.LinkClicked += (sender, args) => AppInfo.OpenLink(this, url);
            return link;
        }

        internal async Task CheckForUpdatesAsync()
        {
            if (checking || IsDisposed || Disposing) return;
            checking = true;
            checkButton.Enabled = false;
            updateStatus.Text = "正在检查更新…";
            updateStatus.Visible = true;
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
