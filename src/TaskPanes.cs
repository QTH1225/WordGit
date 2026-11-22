using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace WordGit
{
    [ComImport, Guid("000C033D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICTPFactory
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.Interface)]
        object CreateCTP([MarshalAs(UnmanagedType.BStr)] string controlId, [MarshalAs(UnmanagedType.BStr)] string title,
            [MarshalAs(UnmanagedType.Struct)] object window);
    }

    [ComImport, Guid("000C033E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)] void CTPFactoryAvailable([MarshalAs(UnmanagedType.Interface)] ICTPFactory factory);
    }

    internal sealed class TaskPanes
    {
        private const int DockRight = 2;
        private readonly ICTPFactory factory;
        private readonly Dictionary<int, Pane> panes = new Dictionary<int, Pane>();

        private sealed class Pane
        {
            public dynamic Native;
            public HistoryPane Control;
        }

        public TaskPanes(ICTPFactory factory) { this.factory = factory; }

        public void Show(object window, Func<Task<HistoryPaneData>> load, Action<string, HistoryAction, string, string> openVersion)
        {
            Prune();
            var handle = (int)((dynamic)window).Hwnd;
            Pane pane;
            if (!panes.TryGetValue(handle, out pane))
            {
                dynamic native = factory.CreateCTP(HistoryPane.ProgId, "WordGit · 批注、修订与提交", window);
                try
                {
                    var control = (HistoryPane)native.ContentControl;
                    native.DockPosition = DockRight;
                    native.Width = 420;
                    pane = new Pane { Native = native, Control = control };
                    control.BindAsync(load, openVersion);
                    panes.Add(handle, pane);
                }
                catch
                {
                    Delete(native);
                    throw;
                }
            }
            pane.Control.PrepareForShow();
            if (!(bool)pane.Native.Visible) pane.Native.Visible = true;
            // Paint the single loading surface before posting the background-load callback.
            pane.Control.Refresh();
            // Return to Word's message loop so the docked pane can paint before loading.
            pane.Control.BeginInvoke(new Action(pane.Control.Reload));
        }

        public string VisiblePath(object window)
        {
            Pane pane;
            return panes.TryGetValue((int)((dynamic)window).Hwnd, out pane) && (bool)pane.Native.Visible ? pane.Control.DocumentPath : null;
        }

        public bool Refresh(object window)
        {
            Prune();
            Pane pane;
            if (!panes.TryGetValue((int)((dynamic)window).Hwnd, out pane) || !(bool)pane.Native.Visible) return false;
            pane.Control.Reload();
            return true;
        }

        public void Dispose()
        {
            foreach (var pane in panes.Values) Delete(pane.Native);
            panes.Clear();
        }

        private void Prune()
        {
            foreach (var handle in panes.Where(entry => entry.Value.Control.IsDisposed || !IsWindow(new IntPtr(entry.Key))).Select(entry => entry.Key).ToArray())
            {
                var pane = panes[handle];
                panes.Remove(handle);
                // Stop the pane's timers and drop its document references before the host goes.
                if (!pane.Control.IsDisposed) pane.Control.Detach();
                Delete(pane.Native);
                WordDocuments.Release((object)pane.Native);
            }
        }

        private static void Delete(object pane)
        {
            if (pane == null) return;
            try { ((dynamic)pane).Delete(); }
            catch (COMException) { }
            catch (InvalidComObjectException) { }
        }

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr handle);
    }
}
