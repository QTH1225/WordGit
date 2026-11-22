// SPDX-License-Identifier: GPL-3.0-only
using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordGit
{
    // A single opaque surface covers child HWNDs until the complete pane is ready.
    internal sealed class PaneLoadingFrame : Control
    {
        internal Func<int> ReviewHeight;
        internal Func<int> FooterHeight;

        internal PaneLoadingFrame()
        {
            Dock = DockStyle.Fill;
            BackColor = SystemColors.Control;
            TabStop = false;
            AccessibleName = "WordGit 正在加载";
            AccessibleDescription = "正在读取批注、修订和提交记录。";
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Opaque, true);
        }

        protected override void OnPaintBackground(PaintEventArgs args)
        {
            args.Graphics.Clear(BackColor);
        }

        protected override void OnPaint(PaintEventArgs args)
        {
            var graphics = args.Graphics;
            graphics.Clear(BackColor);
            if (ClientSize.Width < 40 || ClientSize.Height < 120) return;
            var padding = Math.Max(10, Font.Height / 2);
            var width = Math.Max(1, ClientSize.Width - 2 * padding);
            var footer = FooterHeight == null ? 24 : FooterHeight();
            var split = ReviewHeight == null ? ClientSize.Height * 3 / 8 : ReviewHeight();
            split = Math.Max(110, Math.Min(split, ClientSize.Height - footer - 170));
            using (var border = new Pen(SystemColors.ControlDark))
            using (var bar = new SolidBrush(SystemColors.ControlLight))
            using (var surface = new SolidBrush(SystemColors.Window))
            using (var heading = new Font(Font, FontStyle.Bold))
            {
                TextRenderer.DrawText(graphics, "批注与修订", heading,
                    new Rectangle(padding, padding, width, Font.Height + 6), ForeColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                var reviews = new Rectangle(padding, padding + Font.Height + 14,
                    width, Math.Max(36, split - Font.Height - padding - 52));
                graphics.FillRectangle(surface, reviews);
                graphics.DrawRectangle(border, reviews);
                DrawMessage(graphics, reviews, "正在加载批注与修订…");
                var buttonTop = reviews.Bottom + 8;
                graphics.FillRectangle(bar, padding, buttonTop, Math.Min(70, width / 3), 24);
                graphics.FillRectangle(bar, padding + Math.Min(70, width / 3) + 8,
                    buttonTop, Math.Min(70, width / 3), 24);

                var lower = split + 12;
                TextRenderer.DrawText(graphics, "提交记录", heading,
                    new Rectangle(padding, lower, width, Font.Height + 6), ForeColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                var lineTop = lower + Font.Height + 16;
                graphics.FillRectangle(bar, padding, lineTop, Math.Max(40, width * 2 / 3), 10);
                graphics.FillRectangle(bar, padding, lineTop + 20, Math.Max(30, width / 2), 10);
                var historyTop = lineTop + 44;
                var history = new Rectangle(padding, historyTop, width,
                    Math.Max(30, ClientSize.Height - footer - padding - historyTop));
                graphics.FillRectangle(surface, history);
                graphics.DrawRectangle(border, history);
                DrawMessage(graphics, history, "正在加载分支和提交记录…");
            }
        }

        private void DrawMessage(Graphics graphics, Rectangle bounds, string text)
        {
            bounds.Inflate(-8, -4);
            TextRenderer.DrawText(graphics, text, Font, bounds, SystemColors.GrayText,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
        }
    }
}
