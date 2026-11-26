using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;

namespace WordGit
{
    // Guards the task pane's first-show order: the loading surface must reach the screen
    // before the pane starts reading the core, otherwise Word shows an unpainted pane.
    internal static class PaneProbe
    {
        private static bool started, startedEarly, singleLine, singleLineChecked;

        [STAThread]
        private static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Application.EnableVisualStyles();
            var form = new Form { Width = 430, Height = 660, Text = "pane probe", ShowInTaskbar = false };
            var pane = new HistoryPane();
            var watch = Stopwatch.StartNew();
            form.Controls.Add(pane);
            pane.Dock = DockStyle.Fill;
            form.Shown += (sender, e) =>
            {
                pane.BindAsync(async () =>
                {
                    started = true;
                    startedEarly = !pane.SurfacePainted;
                    await System.Threading.Tasks.Task.Delay(150);
                    return new HistoryPaneData
                    {
                        DocumentPath = "侧栏测试稿.docx", HistoryPath = "侧栏测试稿.docx",
                        Branch = "main", Head = "0123456", Branches = new string[0],
                        InitialHistory = new List<Revision>(), IsSaved = () => true,
                        WorkingState = () => WorkingCopyState.Clean
                    };
                }, (path, action, revision, other) => { });
                pane.PrepareForShow("侧栏测试稿.docx");
                pane.Refresh();
                pane.Update();
                pane.QueueReload();
            };
            var timer = new Timer { Interval = 200 };
            timer.Tick += (sender, e) =>
            {
                if (!singleLineChecked && watch.ElapsedMilliseconds >= 1200)
                {
                    singleLineChecked = true;
                    singleLine = pane.StatusIsSingleLine;
                }
                if (watch.ElapsedMilliseconds < 2500) return;
                timer.Stop();
                form.Close();
            };
            timer.Start();
            try { Application.Run(form); }
            finally
            {
                var failed = new List<string>();
                if (!started) failed.Add("the pane never started loading");
                if (startedEarly) failed.Add("loading started before the loading surface painted");
                if (!pane.SurfacePainted) failed.Add("the loading surface never painted");
                if (!singleLineChecked) failed.Add("the status line was never measured");
                else if (!singleLine) failed.Add("the empty-history status line wrapped onto a second line");
                if (failed.Count == 0) Console.WriteLine("PASS: the loading surface paints before the pane reads anything, and the empty-history status stays on one line.");
                foreach (var reason in failed) Console.WriteLine("FAILED: " + reason);
                Environment.ExitCode = failed.Count == 0 ? 0 : 1;
            }
        }
    }
}