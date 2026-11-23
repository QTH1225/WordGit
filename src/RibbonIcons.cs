using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace WordGit
{
    // Word renders imageMso icons from its own gallery; WPS lacks several of them
    // (ViewOutlineView, FileNew, ReviewCompareMenu) and shows a placeholder instead.
    // These three buttons carry their own images so both editors show the same icon.
    internal static class RibbonIcons
    {
        private static readonly Dictionary<string, object> Pictures = new Dictionary<string, object>();
        private static readonly object Gate = new object();
        private static readonly Color Primary = Color.FromArgb(43, 87, 154);   // Word blue
        private static readonly Color Secondary = Color.FromArgb(118, 118, 118);
        private static readonly Color Accent = Color.FromArgb(16, 124, 16);    // new branch green

        internal static object PictureFor(string id)
        {
            lock (Gate)
            {
                object picture;
                if (!Pictures.TryGetValue(id, out picture))
                {
                    using (var bitmap = BitmapFor(id))
                        picture = IconPicture.FromImage(bitmap);
                    Pictures.Add(id, picture);
                }
                return picture;
            }
        }

        internal static Bitmap BitmapFor(string id)
        {
            var bitmap = new Bitmap(32, 32, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Transparent);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                switch (id)
                {
                    case "graph": DrawGraph(graphics); break;
                    case "branch": DrawBranch(graphics); break;
                    case "compare": DrawCompare(graphics); break;
                    default: throw new ArgumentException("Unknown ribbon icon: " + id);
                }
            }
            return bitmap;
        }

        // Commit graph: main line with a branch joining from the upper right.
        private static void DrawGraph(Graphics graphics)
        {
            using (var pen = Stroke(Secondary)) graphics.DrawLine(pen, 9f, 5f, 9f, 27f);
            using (var pen = Stroke(Primary))
            {
                graphics.DrawLine(pen, 9f, 25f, 23f, 13f);
                graphics.DrawLine(pen, 23f, 13f, 23f, 5f);
            }
            Node(graphics, 9f, 6f, Secondary);
            Node(graphics, 9f, 16f, Secondary);
            Node(graphics, 9f, 26f, Secondary);
            Node(graphics, 23f, 5f, Primary);
            Node(graphics, 23f, 14f, Primary);
        }

        // New branch: the trunk forks into a fresh green tip.
        private static void DrawBranch(Graphics graphics)
        {
            using (var pen = Stroke(Secondary)) graphics.DrawLine(pen, 10f, 26f, 10f, 6f);
            using (var pen = Stroke(Accent)) graphics.DrawLine(pen, 10f, 16f, 22f, 7f);
            Node(graphics, 10f, 26f, Secondary);
            Node(graphics, 10f, 7f, Secondary);
            Node(graphics, 22f, 7f, Accent);
        }

        // Compare versions: a pair of opposing arrows between the two sides.
        private static void DrawCompare(Graphics graphics)
        {
            using (var pen = Stroke(Primary))
            {
                graphics.DrawLine(pen, 7f, 11f, 25f, 11f);
                graphics.DrawLine(pen, 25f, 11f, 18.5f, 5f);
                graphics.DrawLine(pen, 25f, 11f, 18.5f, 17f);
            }
            using (var pen = Stroke(Secondary))
            {
                graphics.DrawLine(pen, 25f, 21f, 7f, 21f);
                graphics.DrawLine(pen, 7f, 21f, 13.5f, 15f);
                graphics.DrawLine(pen, 7f, 21f, 13.5f, 27f);
            }
        }

        private static Pen Stroke(Color color)
        {
            var pen = new Pen(color, 3f);
            pen.StartCap = LineCap.Round;
            pen.EndCap = LineCap.Round;
            return pen;
        }

        private static void Node(Graphics graphics, float x, float y, Color color)
        {
            using (var brush = new SolidBrush(color))
                graphics.FillEllipse(brush, x - 3.6f, y - 3.6f, 7.2f, 7.2f);
        }

        private sealed class IconPicture : AxHost
        {
            private IconPicture() : base(String.Empty) { }
            internal static object FromImage(Image image) { return GetIPictureDispFromPicture(image); }
        }
    }
}
