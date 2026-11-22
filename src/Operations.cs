using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace WordGit
{
    // Word's thread has no WinForms message loop of its own. WinForms uninstalls its
    // SynchronizationContext whenever the outermost modal loop (ShowDialog) ends, after which
    // await continuations would resume on the thread pool and touch Word COM objects and
    // controls from the wrong thread. Every entry point re-installs the context first.
    internal static class UiThread
    {
        [ThreadStatic] private static WindowsFormsSynchronizationContext context;

        internal static void EnsureContext()
        {
            // Word's UI thread is STA; never install a WinForms context on a thread-pool (MTA) thread.
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) return;
            if (SynchronizationContext.Current is WindowsFormsSynchronizationContext) return;
            if (context == null) context = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(context);
        }
    }

    // Tracks background operations per repository (accessed on Word's thread only), so one
    // document's commit does not block another, and Word can wait for them before shutting down.
    internal sealed class OperationTracker
    {
        private readonly Dictionary<string, string> running = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        internal int Count { get { return running.Count; } }

        internal bool IsBusy(string key) { return running.ContainsKey(key); }

        internal void Begin(string key, string title)
        {
            string current;
            if (running.TryGetValue(key, out current))
                throw new InvalidOperationException("这份文档正在后台" + current + "，请等待当前操作完成。");
            running.Add(key, title);
        }

        internal void End(string key) { running.Remove(key); }

        internal string Describe() { return String.Join("、", running.Values); }

        // Shown from OnBeginShutdown: pumps messages so the background continuations can finish.
        internal void WaitForCompletion()
        {
            if (Count == 0) return;
            using (var dialog = new Form
            {
                Text = "WordGit", StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = true, ClientSize = new Size(420, 130),
                Font = SystemFonts.MessageBoxFont, TopMost = true
            })
            using (var timer = new System.Windows.Forms.Timer { Interval = 200 })
            {
                var layout = DialogLayout.Stack(3);
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var status = DialogLayout.Label("正在完成后台操作：" + Describe() + "。完成后 Word 会继续退出。");
                layout.Controls.Add(status, 0, 0);
                layout.Controls.Add(new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee, Height = 14 }, 0, 1);
                var actions = DialogLayout.Buttons(true);
                var skip = DialogLayout.Button("不再等待");
                skip.Click += (sender, args) => dialog.Close();
                actions.Controls.Add(skip);
                layout.Controls.Add(actions, 0, 2);
                dialog.Controls.Add(layout);
                timer.Tick += (sender, args) => { if (Count == 0) dialog.Close(); };
                timer.Start();
                dialog.ShowDialog();
            }
        }
    }
}
