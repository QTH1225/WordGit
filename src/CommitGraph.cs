using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class GraphRow
    {
        public int Node, NodeColor, Columns;
        public bool IsHead;
        public readonly List<GraphEdge> Edges = new List<GraphEdge>();
    }

    internal sealed class GraphEdge
    {
        public readonly int From, To, Color, Start, End;
        public GraphEdge(int from, int to, int color, int start, int end)
        { From = from; To = to; Color = color; Start = start; End = end; }
    }

    internal static class CommitGraph
    {
        private static readonly Color[] LaneColors = { Color.RoyalBlue, Color.SeaGreen, Color.DarkOrange, Color.MediumVioletRed, Color.DarkCyan, Color.DarkViolet };

        private static Color ColorFor(int index, bool selected)
        {
            if (selected) return SystemColors.HighlightText;
            // High contrast themes replace the palette with the system text colour; the lane position
            // still tells branches apart.
            if (SystemInformation.HighContrast) return SystemColors.WindowText;
            return LaneColors[index % LaneColors.Length];
        }

        public static void Paint(DataGridViewCellPaintingEventArgs args, GraphRow row, float laneSpacing)
        {
            args.PaintBackground(args.ClipBounds, true);
            var area = args.CellBounds;
            var scale = args.Graphics.DpiX / 96f;
            var step = laneSpacing * scale;
            var origin = area.Left + 13f * scale;
            var mid = area.Top + area.Height / 2f;
            var selected = (args.State & DataGridViewElementStates.Selected) != 0;
            var saved = args.Graphics.Save();
            try
            {
                args.Graphics.SetClip(Rectangle.Intersect(area, args.ClipBounds), CombineMode.Intersect);
                args.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                foreach (var edge in row.Edges)
                {
                    var startX = origin + edge.From * step;
                    var endX = origin + edge.To * step;
                    var startY = area.Top + edge.Start * area.Height / 2f;
                    var endY = area.Top + edge.End * area.Height / 2f;
                    var bend = (endY - startY) / 2f;
                    using (var pen = new Pen(ColorFor(edge.Color, selected), 2f * scale))
                        args.Graphics.DrawBezier(pen, startX, startY, startX, startY + bend, endX, endY - bend, endX, endY);
                }
                var center = origin + row.Node * step;
                var radius = 4f * scale;
                var color = ColorFor(row.NodeColor, selected);
                if (row.IsHead)
                    using (var pen = new Pen(color, 1.4f * scale))
                        args.Graphics.DrawEllipse(pen, center - radius * 1.7f, mid - radius * 1.7f, radius * 3.4f, radius * 3.4f);
                using (var fill = new SolidBrush(color))
                    args.Graphics.FillEllipse(fill, center - radius, mid - radius, radius * 2, radius * 2);
            }
            finally { args.Graphics.Restore(saved); }
            args.Handled = true;
        }

        private sealed class Lane
        {
            public string Id;
            public int Color;
        }

        public static List<GraphRow> Layout(IList<Revision> revisions, string head)
        {
            var result = new List<GraphRow>();
            var remaining = new HashSet<string>(revisions.Select(revision => revision.Id));
            var lanes = new List<Lane>();
            var nextColor = 0;
            foreach (var revision in revisions)
            {
                remaining.Remove(revision.Id);
                var incoming = lanes.ToList();
                var nodeIndex = lanes.FindIndex(lane => lane.Id == revision.Id);
                var active = nodeIndex >= 0;
                if (!active)
                {
                    nodeIndex = lanes.Count;
                    lanes.Add(new Lane { Id = revision.Id, Color = nextColor++ });
                }
                var node = lanes[nodeIndex];
                var row = new GraphRow { Node = nodeIndex, NodeColor = node.Color, IsHead = revision.Id == head };
                var parents = revision.Parents.Where(remaining.Contains).Distinct().ToArray();
                lanes.RemoveAt(nodeIndex);
                for (var index = 0; index < parents.Length; index++)
                {
                    var parent = parents[index];
                    if (lanes.Any(lane => lane.Id == parent)) continue;
                    var insertAt = System.Math.Min(nodeIndex + index, lanes.Count);
                    lanes.Insert(insertAt, new Lane { Id = parent, Color = index == 0 ? node.Color : nextColor++ });
                }
                var positions = new Dictionary<Lane, int>();
                for (var index = 0; index < lanes.Count; index++) positions[lanes[index]] = index;
                for (var index = 0; index < incoming.Count; index++)
                {
                    var lane = incoming[index];
                    if (lane == node) row.Edges.Add(new GraphEdge(index, nodeIndex, lane.Color, 0, 1));
                    else row.Edges.Add(new GraphEdge(index, positions[lane], lane.Color, 0, 2));
                }
                foreach (var parent in parents)
                {
                    var index = positions[lanes.Find(lane => lane.Id == parent)];
                    row.Edges.Add(new GraphEdge(nodeIndex, index, lanes[index].Color, 1, 2));
                }
                row.Columns = System.Math.Max(nodeIndex + 1, System.Math.Max(incoming.Count, lanes.Count));
                result.Add(row);
            }
            return result;
        }
    }
}
