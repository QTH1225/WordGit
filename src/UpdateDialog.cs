using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;

namespace WordGit
{
    internal static class UpdatePreferences
    {
        private const string Key = @"Software\WordGit\Updates";
        internal static bool IncludeBeta
        {
            get { return Equals(Read("IncludeBeta", 0), 1); }
            set { Write("IncludeBeta", value ? 1 : 0, RegistryValueKind.DWord); }
        }

        private const string IgnoredVersion = "IgnoredVersion";

        internal static bool IsIgnored(ReleaseVersion version)
        {
            return AtOrBelowIgnoredVersion(version, Read(IgnoredVersion, "") as string);
        }

        internal static bool AtOrBelowIgnoredVersion(ReleaseVersion version, string ignoredVersion)
        {
            var ignored = ReleaseVersion.Parse(ignoredVersion);
            return ignored != null && version.CompareTo(ignored) <= 0;
        }

        internal static bool Ignore(string release)
        {
            var version = ReleaseVersion.Parse(release);
            return version != null && (IsIgnored(version) || Write(IgnoredVersion, version.Text, RegistryValueKind.String));
        }

        private static object Read(string name, object fallback)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(Key))
                    return key == null ? fallback : key.GetValue(name, fallback);
            }
            catch (Exception error)
            {
                AppPaths.LogError("read update preferences: " + error);
                return fallback;
            }
        }

        private static bool Write(string name, object value, RegistryValueKind kind)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(Key))
                    key.SetValue(name, value, kind);
                return true;
            }
            catch (Exception error)
            {
                AppPaths.LogError("write update preferences: " + error);
                return false;
            }
        }
    }

    internal sealed class UpdateDialog : Form
    {
        internal readonly string Version;

        internal UpdateDialog(UpdateResult result, Func<string, bool> ignoreRelease = null)
        {
            var ignore = ignoreRelease ?? UpdatePreferences.Ignore;
            Version = result.Version;
            Text = "WordGit · 发现新版本";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            ShowInTaskbar = false;
            MinimizeBox = MaximizeBox = false;
            ClientSize = new Size(555, 365);
            MinimumSize = new Size(445, 300);

            var layout = DialogLayout.Stack(4);
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(16, 14, 16, 10);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var versions = DialogLayout.Label(result.Message);
            versions.Name = "releaseVersions";
            versions.Margin = new Padding(0, 0, 0, 12);
            layout.Controls.Add(versions, 0, 0);
            layout.Controls.Add(DialogLayout.Label("更新日志"), 0, 1);
            var notes = new TextBox { Name = "releaseNotes", AccessibleName = "更新日志",
                Multiline = true, ReadOnly = true, TabStop = false, ScrollBars = ScrollBars.Vertical,
                Text = result.Notes.Replace("\r\n", "\n").Replace("\n", Environment.NewLine),
                Dock = DockStyle.Fill, Margin = new Padding(0, 5, 0, 8), BackColor = SystemColors.Window };
            layout.Controls.Add(notes, 0, 2);

            var actions = new TableLayoutPanel
            {
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill,
                ColumnCount = 3, RowCount = 1, Margin = new Padding(0, 10, 0, 0)
            };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var close = DialogLayout.Button("稍后提醒");
            close.Anchor = AnchorStyles.Right;
            close.Click += (sender, args) => Close();
            var skip = DialogLayout.Button("忽略此版本");
            skip.Anchor = AnchorStyles.Left;
            skip.AccessibleDescription = "暂不提醒这个版本，等到有更高版本时再提醒。";
            skip.Click += (sender, args) =>
            {
                if (ignore(Version)) Close();
                else MessageBox.Show(this, "无法保存忽略设置，请稍后再试。", "WordGit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };
            var download = DialogLayout.Button(result.DownloadUrl.StartsWith(AppInfo.ReleasesUrl + "/download/", StringComparison.Ordinal)
                ? "下载更新" : "打开发布页");
            download.Anchor = AnchorStyles.Right;
            download.Margin = new Padding(0, 0, 0, 4);
            download.FlatStyle = FlatStyle.Flat;
            download.UseVisualStyleBackColor = false;
            download.BackColor = SystemColors.Highlight;
            download.ForeColor = SystemColors.HighlightText;
            download.FlatAppearance.BorderColor = download.BackColor;
            download.FlatAppearance.MouseOverBackColor = download.BackColor;
            download.Click += (sender, args) => AppInfo.OpenLink(this, result.DownloadUrl);
            actions.Controls.Add(skip, 0, 0);
            actions.Controls.Add(close, 1, 0);
            actions.Controls.Add(download, 2, 0);
            layout.Controls.Add(actions, 0, 3);
            CancelButton = close;
            Controls.Add(layout);
            ActiveControl = close;
        }

        internal static void ShowResult(IWin32Window owner, UpdateResult result, bool manual)
        {
            var version = ReleaseVersion.Parse(result.Version);
            if (result.State != UpdateState.Available || version == null) return;
            if (!manual && (UpdatePreferences.IsIgnored(version) || (version.IsBeta && !UpdatePreferences.IncludeBeta))) return;
            var existing = Application.OpenForms.OfType<UpdateDialog>().FirstOrDefault(dialog => !dialog.IsDisposed);
            if (existing != null)
            {
                if (!manual) return;
                if (existing.Version == result.Version) { existing.Activate(); return; }
                existing.Close();
            }
            DialogPlacement.Show(new UpdateDialog(result), owner);
        }
    }
}
