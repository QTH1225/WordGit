// SPDX-License-Identifier: GPL-3.0-only
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordGit
{
    internal static class DialogPlacement
    {
        private sealed class WindowOwner : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            internal WindowOwner(IntPtr handle) { Handle = handle; }
        }

        internal static IWin32Window RootOwner(IWin32Window owner)
        {
            if (owner == null || owner.Handle == IntPtr.Zero) return null;
            var root = GetAncestor(owner.Handle, 2);
            return root == IntPtr.Zero || root == owner.Handle ? owner : new WindowOwner(root);
        }

        internal static IWin32Window OwnerFor(Control control)
        {
            var form = control.FindForm();
            return RootOwner(form ?? (IWin32Window)control);
        }

        internal static Rectangle CenteredBounds(Size size, Rectangle owner, Rectangle workArea)
        {
            var left = owner.Left + (owner.Width - size.Width) / 2;
            var top = owner.Top + (owner.Height - size.Height) / 2;
            left = Math.Max(workArea.Left, Math.Min(left, workArea.Right - size.Width));
            top = Math.Max(workArea.Top, Math.Min(top, workArea.Bottom - size.Height));
            return new Rectangle(new Point(left, top), size);
        }

        internal static void Center(Form dialog, IWin32Window owner)
        {
            owner = RootOwner(owner);
            var screen = owner == null ? Screen.FromPoint(Cursor.Position) : Screen.FromHandle(owner.Handle);
            var area = screen.WorkingArea;
            NativeRectangle rectangle;
            var bounds = owner != null && GetWindowRect(owner.Handle, out rectangle)
                ? Rectangle.FromLTRB(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom) : area;
            dialog.Location = CenteredBounds(dialog.Size, bounds, area).Location;
        }

        internal static IWin32Window Prepare(Form dialog, IWin32Window owner)
        {
            owner = RootOwner(owner);
            dialog.StartPosition = FormStartPosition.Manual;
            Center(dialog, owner);
            // Recalculate after Word applies its monitor's font/DPI scaling, before painting.
            dialog.Shown += (sender, args) => Center(dialog, owner);
            return owner;
        }

        internal static void Show(Form dialog, IWin32Window owner)
        {
            owner = Prepare(dialog, owner);
            if (owner == null) dialog.Show();
            else dialog.Show(owner);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRectangle { internal int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr handle, out NativeRectangle rectangle);
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr handle, uint flags);
    }
}
