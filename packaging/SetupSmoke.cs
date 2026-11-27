using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordGitSetup
{
    internal static class SetupSmoke
    {
        private static IEnumerable<Control> All(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                yield return child;
                foreach (var descendant in All(child)) yield return descendant;
            }
        }

        private static void VerifyDisclaimer(Payload payload, string output)
        {
            var markdown = File.ReadAllLines(Path.Combine(payload.Root, InstallDisclaimer.DocumentPath));
            var expected = Regex.Replace(String.Join("\n", markdown.Skip(1)), @"(?m)^#### ", "");
            expected = Regex.Replace(expected, @"(?m)^- ", "• ").Replace("**", "").Replace("`", "").Replace("\\\n", "\n").Trim();
            foreach (var approve in new[] { false, true })
            using (var dialog = new InstallDisclaimer(payload.Root))
            {
                dialog.ShowInTaskbar = false;
                dialog.Opacity = 0;
                var text = All(dialog).OfType<RichTextBox>().Single();
                var agree = All(dialog).OfType<CheckBox>().Single();
                var proceed = (Button)dialog.AcceptButton;
                var cancel = (Button)dialog.CancelButton;
                Exception failure = null;
                dialog.Shown += (sender, args) => dialog.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        if (!text.ReadOnly || text.ScrollBars != RichTextBoxScrollBars.Vertical ||
                            text.Text.Replace("\r\n", "\n").Trim() != expected)
                            throw new Exception("Installer disclaimer does not display the full source text.");
                        if (agree.Checked || proceed.Enabled)
                            throw new Exception("The installer must start with consent unchecked.");
                        proceed.PerformClick();
                        if (dialog.DialogResult != DialogResult.None)
                            throw new Exception("The installer can proceed without consent.");

                        foreach (var size in new[] { dialog.Size, dialog.MinimumSize })
                        {
                            dialog.Size = size;
                            Application.DoEvents();
                            if (agree.Left != 0 || proceed.Left <= agree.Right || cancel.Left <= proceed.Right ||
                                cancel.Right != cancel.Parent.ClientSize.Width - cancel.Margin.Right)
                                throw new Exception("Disclaimer footer controls are clipped or misaligned.");
                            if (!approve && size != dialog.MinimumSize)
                            using (var bitmap = new Bitmap(dialog.Width, dialog.Height))
                            {
                                dialog.DrawToBitmap(bitmap, new Rectangle(Point.Empty, dialog.Size));
                                bitmap.Save(Path.Combine(output, "setup-disclaimer.png"));
                            }
                        }

                        agree.Checked = true;
                        if (!proceed.Enabled) throw new Exception("Consent did not enable Continue.");
                        agree.Checked = false;
                        if (proceed.Enabled) throw new Exception("Clearing consent did not disable Continue.");
                        if (approve) { agree.Checked = true; proceed.PerformClick(); }
                        else cancel.PerformClick();
                    }
                    catch (Exception error) { failure = error; dialog.Close(); }
                }));
                var result = dialog.ShowDialog();
                if (failure != null) throw failure;
                if (result != (approve ? DialogResult.OK : DialogResult.Cancel))
                    throw new Exception("The disclaimer returned the wrong installation decision.");
            }
        }

        private static void VerifySetupCancellation(Payload payload, SetupForm form)
        {
            // Fail safely if the real form ever calls the install script after cancellation.
            var invoked = Path.Combine(payload.Root, "installer-invoked");
            File.WriteAllText(Path.Combine(payload.Root, "install.ps1"),
                "[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'installer-invoked'), 'test')");
            var install = All(form).OfType<Button>().Single(button => button.Text == "安装 / 更新");
            All(form).Single(control => control.AccessibleName == "数据目录").Text =
                Path.Combine(payload.Root, "cancelled-data");
            var cancelled = false;
            form.Show();
            using (var timer = new Timer { Interval = 100 })
            {
                timer.Tick += (sender, args) =>
                {
                    timer.Stop();
                    var disclaimer = Application.OpenForms.OfType<InstallDisclaimer>().SingleOrDefault();
                    if (disclaimer == null) return;
                    cancelled = true;
                    ((Button)disclaimer.CancelButton).PerformClick();
                };
                timer.Start();
                install.PerformClick();
            }
            if (!cancelled || !form.IsDisposed || File.Exists(invoked))
                throw new Exception("Cancelling the disclaimer did not close Setup before installation.");
        }

        [STAThread]
        private static int Main(string[] args)
        {
            try
            {
                Application.EnableVisualStyles();
                using (var payload = Payload.Extract())
                using (var form = new SetupForm(payload, false, false))
                {
                    VerifyDisclaimer(payload, args[0]);
                    form.ShowInTaskbar = false;
                    form.Opacity = 0;
                    foreach (var size in new[] { new Size(626, 459), new Size(560, 400) })
                    {
                        form.Size = size;
                        form.Show();
                        Application.DoEvents();
                        form.PerformLayout();
                        var directory = All(form).Single(control => control.AccessibleName == "数据目录");
                        if (!directory.Visible || directory.Text.Length == 0 || !directory.Parent.ClientRectangle.Contains(directory.Bounds))
                            throw new Exception("Missing or clipped data directory field.");
                        foreach (var control in All(form).Where(control => control is Button || control is Label))
                            if (control.Visible && !control.Parent.ClientRectangle.Contains(control.Bounds))
                                throw new Exception("Clipped setup control: " + control.Text);
                        using (var bitmap = new Bitmap(form.Width, form.Height))
                        {
                            form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size));
                            bitmap.Save(Path.Combine(args[0], "setup-" + size.Width + ".png"));
                        }
                        form.Hide();
                    }
                    VerifySetupCancellation(payload, form);
                }
                Console.WriteLine("PASS: setup embedded payload, disclaimer consent, cancellation before installation, default/minimum layouts.");
                return 0;
            }
            catch (Exception error) { Console.Error.WriteLine(error.Message); return 1; }
        }
    }
}
