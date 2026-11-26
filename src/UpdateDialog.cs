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
            set { Write("IncludeBeta", value ? 1 : 0); }
        }

        internal static bool WasShown(string version) { return Equals(Read("Shown." + version, 0), 1); }
        internal static void MarkShown(string version) { Write("Shown." + version, 1); }

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

        private static void Write(string name, int value)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(Key))
                    key.SetValue(name, value, RegistryValueKind.DWord);
            }
            catch (Exception error) { AppPaths.LogError("write update preferences: " + error); }
        }
    }

    internal sealed class UpdateDialog : Form
    {
        internal readonly string Version;

        internal UpdateDialog(UpdateResult result)
        {
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
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                Text = result.Notes.Replace("\r\n", "\n").Replace("\n", Environment.NewLine),
                Dock = DockStyle.Fill, Margin = new Padding(0, 5, 0, 8), BackColor = SystemColors.Window };
            layout.Controls.Add(notes, 0, 2);

            var actions = DialogLayout.Buttons(true);
            var close = DialogLayout.Button("稍后");
            close.Click += (sender, args) => Close();
            var download = DialogLayout.Button("下载更新");
            download.Click += (sender, args) => AppInfo.OpenLink(this, result.DownloadUrl);
            actions.Controls.Add(close);
            actions.Controls.Add(download);
            layout.Controls.Add(actions, 0, 3);
            CancelButton = close;
            Controls.Add(layout);
        }

        internal static void ShowResult(IWin32Window owner, UpdateResult result, bool manual)
        {
            if (result.State != UpdateState.Available || String.IsNullOrEmpty(result.Version)) return;
            if (!manual && (UpdatePreferences.WasShown(result.Version) ||
                (ReleaseVersion.Parse(result.Version).IsBeta && !UpdatePreferences.IncludeBeta))) return;
            var existing = Application.OpenForms.OfType<UpdateDialog>().FirstOrDefault(dialog => !dialog.IsDisposed);
            if (existing != null)
            {
                if (!manual) return;
                if (existing.Version == result.Version) { existing.Activate(); return; }
                existing.Close();
            }
            DialogPlacement.Show(new UpdateDialog(result), owner);
            UpdatePreferences.MarkShown(result.Version);
        }
    }
}
