using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WordGit
{
    internal static class QuietWord
    {
        private const uint DesktopAccess = 0x01FF;
        private const uint NativeObjectModel = unchecked((uint)-16);
        private const uint StillRunning = 258;
        private static readonly Guid Dispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private static readonly object WorkerGate = new object();

        internal static Task Run(string executable, Action<WordDocuments> build)
        {
            var finished = new TaskCompletionSource<bool>();
            var thread = new Thread(() =>
            {
                try
                {
                    lock (WorkerGate) Build(executable, build);
                    finished.SetResult(true);
                }
                catch (Exception error) { finished.SetException(error); }
            }) { IsBackground = true, Name = "WordGit 修订处理" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return finished.Task;
        }

        private static void Build(string executable, Action<WordDocuments> build)
        {
            if (!File.Exists(executable)) throw new InvalidOperationException("找不到用于生成修订稿的 Microsoft Word。");
            var name = "WordGit-" + Guid.NewGuid().ToString("N");
            var desktop = CreateDesktop(name, IntPtr.Zero, IntPtr.Zero, 0, DesktopAccess, IntPtr.Zero);
            if (desktop == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error(), "无法创建 Word 修订窗口的独立桌面。");
            var process = new ProcessInformation();
            dynamic application = null;
            try
            {
                // Word requires active document windows during comparison. All those windows
                // belong to this process on a desktop that never receives user input.
                var startup = new StartupInfo { Size = Marshal.SizeOf(typeof(StartupInfo)), Desktop = name };
                var command = new StringBuilder("\"" + executable + "\" /w /q /a /m");
                if (!CreateProcess(executable, command, IntPtr.Zero, IntPtr.Zero, false, 0, IntPtr.Zero,
                    Path.GetDirectoryName(executable), ref startup, out process))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "无法启动用于生成修订稿的 Word 进程。");
                application = WaitForWord(desktop, process);
                application.DisplayAlerts = 0;
                build(new WordDocuments((object)application));
            }
            finally
            {
                if (application != null)
                {
                    try { application.Quit(SaveChanges: 0); }
                    catch (COMException error) { AppPaths.LogError("close review worker: " + error.Message); }
                    finally { WordDocuments.Release(application); }
                }
                if (process.Process != IntPtr.Zero)
                {
                    if (WaitForSingleObject(process.Process, 3000) == StillRunning) TerminateProcess(process.Process, 1);
                    CloseHandle(process.Process);
                    CloseHandle(process.Thread);
                }
                CloseDesktop(desktop);
            }
        }

        private static object WaitForWord(IntPtr desktop, ProcessInformation process)
        {
            var timer = Stopwatch.StartNew();
            do
            {
                foreach (var window in DocumentWindows(desktop, process.ProcessId))
                {
                    var dispatch = Dispatch;
                    object model;
                    if (AccessibleObjectFromWindow(window, NativeObjectModel, ref dispatch, out model) != 0) continue;
                    try { return ((dynamic)model).Application; }
                    catch (COMException) { }
                    finally { WordDocuments.Release(model); }
                }
            } while (timer.Elapsed < TimeSpan.FromSeconds(25) && WaitForSingleObject(process.Process, 125) == StillRunning);
            throw new InvalidOperationException("后台 Word 未能打开独立修订会话。请确认 Word 可以正常打开文档后重试。");
        }

        private static List<IntPtr> DocumentWindows(IntPtr desktop, uint process)
        {
            var windows = new List<IntPtr>();
            EnumDesktopWindows(desktop, (top, ignored) =>
            {
                uint owner;
                GetWindowThreadProcessId(top, out owner);
                if (owner == process)
                    EnumChildWindows(top, (child, unused) =>
                    {
                        var name = new StringBuilder(32);
                        if (GetClassName(child, name, name.Capacity) != 0 && name.ToString() == "_WwG") windows.Add(child);
                        return true;
                    }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            return windows;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct StartupInfo
        {
            internal int Size;
            internal string Reserved, Desktop, Title;
            internal uint X, Y, XSize, YSize, XCountChars, YCountChars, FillAttribute, Flags;
            internal ushort ShowWindow, ReservedSize;
            internal IntPtr ReservedPointer, StandardInput, StandardOutput, StandardError;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ProcessInformation
        {
            internal IntPtr Process, Thread;
            internal uint ProcessId, ThreadId;
        }

        private delegate bool WindowVisitor(IntPtr window, IntPtr context);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateDesktop(string name, IntPtr device, IntPtr mode, uint flags, uint access, IntPtr security);
        [DllImport("user32.dll")]
        private static extern bool CloseDesktop(IntPtr desktop);
        [DllImport("user32.dll")]
        private static extern bool EnumDesktopWindows(IntPtr desktop, WindowVisitor callback, IntPtr context);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr window, WindowVisitor callback, IntPtr context);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr window, uint id, ref Guid type,
            [MarshalAs(UnmanagedType.IDispatch)] out object model);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CreateProcess(string application, StringBuilder command, IntPtr processAttributes,
            IntPtr threadAttributes, bool inherit, uint flags, IntPtr environment, string directory,
            ref StartupInfo startup, out ProcessInformation process);
        [DllImport("kernel32.dll")]
        private static extern uint WaitForSingleObject(IntPtr handle, uint timeout);
        [DllImport("kernel32.dll")]
        private static extern bool TerminateProcess(IntPtr process, uint exitCode);
        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr handle);
    }
}
