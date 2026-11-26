using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

        [STAThread]
        private static int Main(string[] args)
        {
            try
            {
                Application.EnableVisualStyles();
                using (var payload = Payload.Extract())
                using (var form = new SetupForm(payload, false, false))
                {
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
                }
                Console.WriteLine("PASS: setup embedded payload and default/minimum layouts.");
                return 0;
            }
            catch (Exception error) { Console.Error.WriteLine(error.Message); return 1; }
        }
    }
}
