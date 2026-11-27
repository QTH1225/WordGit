using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Automation;

namespace WordGit
{
    // Word's Ribbon XML exposes button actions but no callback for selecting a custom tab.
    internal sealed class RibbonUpdateTrigger : IDisposable
    {
        private const int MouseHook = 7, LeftButtonUp = 0x202, RootWindow = 2;
        private const int RibbonSearchHeight = 192;
        private readonly Action activated;
        private readonly HookCallback callback;
        private readonly SynchronizationContext wordThread;
        private readonly int processId;
        private IntPtr hook;
        private int checkingPoint;
        private volatile bool disposed;

        internal RibbonUpdateTrigger(Action activated)
        {
            this.activated = activated;
            wordThread = SynchronizationContext.Current;
            callback = OnMouse;
            using (var process = Process.GetCurrentProcess())
                processId = process.Id;
            hook = SetWindowsHookEx(MouseHook, callback, IntPtr.Zero, GetCurrentThreadId());
        }

        private IntPtr OnMouse(int code, IntPtr message, IntPtr data)
        {
            if (code >= 0 && message.ToInt64() == LeftButtonUp && data != IntPtr.Zero)
            {
                var clicked = (MouseInput)Marshal.PtrToStructure(data, typeof(MouseInput));
                var root = GetAncestor(clicked.Window, RootWindow);
                Rectangle bounds;
                if (root != IntPtr.Zero && GetWindowRect(root, out bounds) &&
                    clicked.Y >= bounds.Top && clicked.Y < bounds.Top + RibbonSearchHeight &&
                    Interlocked.CompareExchange(ref checkingPoint, 1, 0) == 0)
                {
                    // Inspect only clicks near the ribbon, away from Word's UI thread.
                    ThreadPool.QueueUserWorkItem(_ => CheckTabAt(new Point(clicked.X, clicked.Y)));
                }
            }
            return CallNextHookEx(hook, code, message, data);
        }

        private void CheckTabAt(Point point)
        {
            try
            {
                if (disposed) return;
                var element = AutomationElement.FromPoint(point);
                for (var depth = 0; depth < 3 && element != null && element.Current.ProcessId == processId; depth++)
                {
                    if (element.Current.ControlType == ControlType.TabItem && element.Current.Name == "WordGit")
                    {
                        wordThread.Post(_ => { if (!disposed) activated(); }, null);
                        return;
                    }
                    element = TreeWalker.ControlViewWalker.GetParent(element);
                }
            }
            catch (ElementNotAvailableException) { }
            catch (COMException) { }
            catch (InvalidOperationException) { }
            finally { Interlocked.Exchange(ref checkingPoint, 0); }
        }

        public void Dispose()
        {
            disposed = true;
            if (hook != IntPtr.Zero) UnhookWindowsHookEx(hook);
            hook = IntPtr.Zero;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MouseInput { internal int X, Y; internal IntPtr Window; }

        [StructLayout(LayoutKind.Sequential)]
        private struct Rectangle { internal int Left, Top, Right, Bottom; }

        private delegate IntPtr HookCallback(int code, IntPtr message, IntPtr data);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", EntryPoint = "SetWindowsHookExW")]
        private static extern IntPtr SetWindowsHookEx(int id, HookCallback callback, IntPtr module, uint thread);
        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, int flags);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr window, out Rectangle bounds);
    }
}
