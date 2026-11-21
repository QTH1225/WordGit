using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class HistoryPaneData
    {
        public string DocumentPath, HistoryPath, Branch, Head;
        public string[] Branches;
        public Func<string, List<Revision>> Load;
        public Func<WorkingCopyState> WorkingState;
        public Func<List<DocumentReviewItem>> LoadReviewItems;
        public Action<DocumentReviewItem> NavigateReviewItem;
        public Func<DocumentReviewItem, bool, bool> ResolveReviewItem;
        public bool CanResolveRevisions;
    }

    [ComVisible(true), Guid("F55EC3C1-B725-4ED8-9CA6-ED51D07C3361"), ProgId(HistoryPane.ProgId)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class HistoryPane : UserControl
    {
        internal const string ProgId = "WordGit.HistoryPane";

        private readonly Label branch = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
        private readonly Label filename = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, ForeColor = SystemColors.GrayText };
        private readonly Label status = DialogLayout.Label("");
        private readonly ComboBox filter = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = "筛选分支" };
        private readonly Label working = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
        private readonly Button compareLatest = DialogLayout.Button("对比最新提交");
        private readonly Timer workingTimer = new Timer { Interval = 1200 };
        private readonly Button commit = DialogLayout.Button("提交版本");
        private readonly Button preview = DialogLayout.Button("预览");
        private readonly Button compare = DialogLayout.Button("对比当前稿");
        private readonly ToolTip hints = new ToolTip();
        private readonly Label commentsTitle = DialogLayout.Label("批注与修订");
        private readonly Button acceptRevision = DialogLayout.Button("接受");
        private readonly Button rejectRevision = DialogLayout.Button("拒绝");
        private readonly Label reviewStatus = DialogLayout.Label("");
        private bool resolvingRevision;
        private bool reviewItemsCurrent;
        private readonly ComboBox reviewFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 78, AccessibleName = "筛选批注与修订" };
        private readonly ListBox comments = new ListBox
        {
            Dock = DockStyle.Fill, DrawMode = DrawMode.OwnerDrawFixed, IntegralHeight = false,
            AccessibleName = "批注与修订列表"
        };
        private List<DocumentReviewItem> commentItems = new List<DocumentReviewItem>();
        private List<DocumentReviewItem> allReviewItems = new List<DocumentReviewItem>();
        private int commentPoll;
        private readonly DataGridView graph = new DataGridView
        {
            Dock = DockStyle.Fill, BackgroundColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal, ColumnHeadersVisible = false,
            RowHeadersVisible = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false, ReadOnly = true, VirtualMode = true, MultiSelect = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect, AccessibleName = "Git 提交记录",
            Margin = new Padding(0, 8, 0, 0)
        };

        private Func<HistoryPaneData> load;
        private Action<string, HistoryAction, string, string> openVersion;
        private HistoryPaneData data;
        private List<Revision> revisions = new List<Revision>();
        private List<GraphRow> rows = new List<GraphRow>();
        private bool updating;

        internal string DocumentPath { get { return data == null ? null : data.DocumentPath; } }

        public HistoryPane()
        {
            Font = SystemFonts.MessageBoxFont;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            AutoScaleMode = AutoScaleMode.Font;
            Size = new Size(420, 640);
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal,
                Size = Size, SplitterDistance = 240, SplitterWidth = 6, Panel1MinSize = 120, Panel2MinSize = 280 };
            Controls.Add(split);
            var commentLayout = DialogLayout.Stack(3);
            commentLayout.Padding = new Padding(10, 8, 10, 2);
            commentLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            commentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            commentLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var reviewHeader = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0, 0, 0, 4) };
            reviewHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            reviewHeader.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            reviewHeader.Controls.Add(commentsTitle, 0, 0);
            reviewFilter.Items.AddRange(new object[] { "全部", "批注", "修订" });
            reviewFilter.SelectedIndex = 0;
            reviewFilter.SelectedIndexChanged += (sender, args) => { if (!updating) FilterReviewItems(); };
            reviewHeader.Controls.Add(reviewFilter, 1, 0);
            commentLayout.Controls.Add(reviewHeader, 0, 0);
            comments.ItemHeight = 3 * Font.Height + 12;
            comments.DrawItem += PaintComment;
            comments.SelectedIndexChanged += (sender, args) =>
            {
                UpdateReviewActions();
                if (!updating && !resolvingRevision && comments.SelectedIndex >= 0 && data != null && data.NavigateReviewItem != null)
                    data.NavigateReviewItem(commentItems[comments.SelectedIndex]);
            };
            comments.MouseMove += (sender, args) =>
            {
                var index = comments.IndexFromPoint(args.Location);
                var text = index < 0 || index >= commentItems.Count ? "" : commentItems[index].Caption + Environment.NewLine + commentItems[index].Text;
                if (hints.GetToolTip(comments) != text) hints.SetToolTip(comments, text);
            };
            commentLayout.Controls.Add(comments, 0, 1);
            var reviewActions = DialogLayout.Buttons(false);
            reviewActions.WrapContents = true;
            acceptRevision.Enabled = false;
            rejectRevision.Enabled = false;
            acceptRevision.Click += (sender, args) => ResolveSelectedRevision(true);
            rejectRevision.Click += (sender, args) => ResolveSelectedRevision(false);
            reviewActions.Controls.Add(acceptRevision);
            reviewActions.Controls.Add(rejectRevision);
            reviewActions.Controls.Add(reviewStatus);
            commentLayout.Controls.Add(reviewActions, 0, 2);
            hints.SetToolTip(acceptRevision, "接受选中的修订，并定位下一条。");
            hints.SetToolTip(rejectRevision, "拒绝选中的修订，并定位下一条。");
            split.Panel1.Controls.Add(commentLayout);
            var layout = DialogLayout.Stack(7);
            layout.Padding = new Padding(10);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            split.Panel2.Controls.Add(layout);

            var header = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0) };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            branch.Font = new Font(Font, FontStyle.Bold);
            branch.Height = Font.Height + 10;
            branch.Margin = new Padding(0, 0, 6, 0);
            var refresh = DialogLayout.Button("刷新");
            refresh.Margin = new Padding(0, 0, 0, 2);
            refresh.Click += (sender, args) => Reload();
            header.Controls.Add(branch, 0, 0);
            header.Controls.Add(refresh, 1, 0);
            layout.Controls.Add(header, 0, 0);
            filename.Height = Font.Height + 6;
            filename.Margin = new Padding(0, 2, 0, 6);
            layout.Controls.Add(filename, 0, 1);

            var workingRow = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0, 0, 0, 6) };
            workingRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            workingRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            working.Height = Font.Height + 10;
            working.Margin = new Padding(0, 0, 6, 0);
            compareLatest.Margin = new Padding(0);
            compareLatest.Enabled = false;
            compareLatest.Click += (sender, args) => openVersion(DocumentPath, HistoryAction.CompareLatest, null, null);
            workingRow.Controls.Add(working, 0, 0);
            workingRow.Controls.Add(compareLatest, 1, 0);
            layout.Controls.Add(workingRow, 0, 2);
            hints.SetToolTip(compareLatest, "用 Word 对比当前稿和当前分支的最新提交。");
            workingTimer.Tick += (sender, args) =>
            {
                if (!Visible || resolvingRevision) return;
                UpdateWorkingState();
                if (++commentPoll % 4 == 0) ReloadComments();
            };
            filter.Margin = new Padding(0);
            filter.SelectedIndexChanged += (sender, args) => { if (!updating && data != null) LoadGraph(null); };
            layout.Controls.Add(filter, 0, 3);

            graph.RowTemplate.Height = 3 * Font.Height + 21;
            graph.Columns.Add(new DataGridViewTextBoxColumn { Width = 42, MinimumWidth = 28, SortMode = DataGridViewColumnSortMode.NotSortable });
            graph.Columns.Add(new DataGridViewTextBoxColumn { AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable });
            graph.CellValueNeeded += (sender, args) =>
            {
                if (args.ColumnIndex != 1 || args.RowIndex >= revisions.Count) return;
                var revision = revisions[args.RowIndex];
                args.Value = Title(revision) + Environment.NewLine + Metadata(revision);
            };
            graph.CellToolTipTextNeeded += (sender, args) =>
            {
                if (args.RowIndex < 0 || args.RowIndex >= revisions.Count) return;
                var revision = revisions[args.RowIndex];
                args.ToolTipText = Title(revision) + Environment.NewLine + String.Join(" · ", revision.References) +
                    Environment.NewLine + Metadata(revision, "yyyy-MM-dd HH:mm") +
                    Environment.NewLine + revision.Message;
            };
            graph.CellPainting += PaintRow;
            graph.SelectionChanged += (sender, args) => { if (!updating) ShowSelection(); };
            graph.CellDoubleClick += (sender, args) => { if (args.RowIndex >= 0 && preview.Enabled) OpenSelected(false); };
            layout.Controls.Add(graph, 0, 5);

            status.MaximumSize = new Size(255, 0);
            layout.Controls.Add(status, 0, 6);
            var actions = DialogLayout.Buttons(false);
            actions.WrapContents = true;
            actions.Margin = new Padding(0, 8, 0, 0);
            preview.Click += (sender, args) => OpenSelected(false);
            compare.Click += (sender, args) => OpenSelected(true);
            commit.Enabled = false;
            commit.Click += (sender, args) => openVersion(DocumentPath, HistoryAction.Commit, null, null);
            actions.Controls.Add(commit);
            actions.Controls.Add(preview);
            actions.Controls.Add(compare);
            layout.Controls.Add(actions, 0, 4);
            hints.SetToolTip(commit, "保存当前文档，并提交到当前分支。");
            hints.SetToolTip(preview, "双击提交也可在 Word 中只读打开。");
            hints.SetToolTip(compare, "按 Ctrl 可选择两个历史版本。");
            ShowSelection();
        }

        internal void Bind(Func<HistoryPaneData> load, Action<string, HistoryAction, string, string> openVersion)
        {
            this.load = load;
            this.openVersion = openVersion;
            workingTimer.Start();
        }

        internal void Reload()
        {
            if (load == null) return;
            string previousFilter = filter.SelectedIndex > 0 ? (string)filter.SelectedItem : null;
            var previousSelection = Selected();
            var previousPath = DocumentPath;
            try
            {
                data = load();
                if (!String.Equals(previousPath, DocumentPath, StringComparison.OrdinalIgnoreCase))
                {
                    previousFilter = null;
                    previousSelection = new Revision[0];
                }
                updating = true;
                try
                {
                    branch.Text = "当前分支  " + data.Branch;
                    filename.Text = Path.GetFileName(data.HistoryPath);
                    hints.SetToolTip(branch, branch.Text);
                    hints.SetToolTip(filename, data.HistoryPath);
                    filter.Items.Clear();
                    filter.Items.Add("全部分支");
                    filter.Items.AddRange(data.Branches);
                    filter.SelectedIndex = previousFilter == null ? 0 : Math.Max(0, filter.Items.IndexOf(previousFilter));
                    filter.Enabled = data.Branches.Length > 0;
                }
                finally { updating = false; }
                LoadGraph(previousSelection.Select(item => item.Id).ToArray());
            }
            catch (Exception error)
            {
                data = null;
                branch.Text = "当前分支  —";
                filename.Text = "";
                filter.Enabled = false;
                ShowGraph(new List<Revision>(), new string[0]);
                status.Text = error.Message;
                hints.SetToolTip(status, error.Message);
            }
            commit.Enabled = data != null;
            compareLatest.Enabled = data != null && data.Head != null;
            UpdateWorkingState();
            ReloadComments();
        }

        internal bool ReloadComments()
        {
            try
            {
                var entries = data == null || data.LoadReviewItems == null ? new List<DocumentReviewItem>() :
                    data.LoadReviewItems().OrderBy(entry => entry.Page).ThenBy(entry => entry.ScopeStart)
                        .ThenBy(entry => entry.IsComment ? 0 : 1).ToList();
                var unchanged = entries.Count == allReviewItems.Count && entries.Zip(allReviewItems,
                    (left, right) => left.Key == right.Key && left.Text == right.Text && left.Author == right.Author &&
                        left.Page == right.Page && left.ScopeStart == right.ScopeStart && left.ScopeEnd == right.ScopeEnd &&
                        left.RevisionType == right.RevisionType).All(equal => equal);
                reviewItemsCurrent = true;
                commentsTitle.Text = "批注 " + entries.Count(entry => entry.IsComment) + " · 修订 " + entries.Count(entry => !entry.IsComment);
                if (!unchanged)
                {
                    allReviewItems = entries;
                    FilterReviewItems();
                }
                UpdateReviewActions();
                return true;
            }
            catch (COMException) { commentsTitle.Text = "批注与修订 · 请刷新"; }
            catch (InvalidComObjectException) { commentsTitle.Text = "批注与修订 · 请刷新"; }
            reviewItemsCurrent = false;
            UpdateReviewActions();
            return false;
        }

        private void FilterReviewItems()
        {
            var selected = comments.SelectedIndex < 0 || comments.SelectedIndex >= commentItems.Count ? null : commentItems[comments.SelectedIndex];
            var top = comments.TopIndex;
            var entries = allReviewItems.Where(entry => reviewFilter.SelectedIndex == 0 ||
                (reviewFilter.SelectedIndex == 1 ? entry.IsComment : !entry.IsComment)).ToList();
            updating = true;
            try
            {
                comments.BeginUpdate();
                comments.Items.Clear();
                commentItems = entries;
                foreach (var entry in entries) comments.Items.Add(entry.Caption + "  " + entry.Text);
                comments.SelectedIndex = selected == null ? -1 : entries.FindIndex(entry => entry.Key == selected.Key && entry.Text == selected.Text);
                if (entries.Count > 0) comments.TopIndex = Math.Min(top, entries.Count - 1);
            }
            finally { comments.EndUpdate(); updating = false; }
            UpdateReviewActions();
        }

        private void UpdateReviewActions()
        {
            var item = comments.SelectedIndex < 0 || comments.SelectedIndex >= commentItems.Count ? null : commentItems[comments.SelectedIndex];
            var enabled = reviewItemsCurrent && !resolvingRevision && item != null && !item.IsComment && data != null &&
                data.CanResolveRevisions && data.ResolveReviewItem != null;
            acceptRevision.Enabled = enabled;
            rejectRevision.Enabled = enabled;
        }

        private void ResolveSelectedRevision(bool accept)
        {
            var index = comments.SelectedIndex;
            if (!reviewItemsCurrent || resolvingRevision || index < 0 || index >= commentItems.Count || data == null ||
                !data.CanResolveRevisions || data.ResolveReviewItem == null || commentItems[index].IsComment) return;
            var selected = commentItems[index];
            var before = allReviewItems.Where(item => !item.IsComment).ToList();
            var ordinal = before.IndexOf(selected);
            DocumentReviewItem next = null;
            resolvingRevision = true;
            UpdateReviewActions();
            try
            {
                if (!data.ResolveReviewItem(selected, accept))
                {
                    reviewStatus.Text = "未处理，请刷新后重试。";
                    return;
                }
                if (!ReloadComments())
                {
                    reviewStatus.Text = "已处理，请刷新列表。";
                    return;
                }
                var remaining = allReviewItems.Where(item => !item.IsComment).ToList();
                var candidates = before.Skip(ordinal + 1).Concat(before.Take(ordinal));
                foreach (var candidate in candidates)
                {
                    next = remaining.Where(item => item.StoryType == candidate.StoryType && item.StoryIndex == candidate.StoryIndex &&
                        item.RevisionType == candidate.RevisionType && item.Author == candidate.Author && item.RangeText == candidate.RangeText &&
                        item.Text == candidate.Text).OrderBy(item => Math.Abs(remaining.IndexOf(item) - ordinal)).FirstOrDefault();
                    if (next != null) break;
                }
                if (next == null && remaining.Count > 0) next = remaining[ordinal < remaining.Count ? Math.Max(0, ordinal) : 0];
                updating = true;
                try { comments.SelectedIndex = next == null ? -1 : commentItems.IndexOf(next); }
                finally { updating = false; }
                reviewStatus.Text = next == null ? "修订已全部处理。" : accept ? "已接受，定位下一条。" : "已拒绝，定位下一条。";
                UpdateWorkingState();
            }
            finally
            {
                resolvingRevision = false;
                UpdateReviewActions();
            }
            if (next != null && data.NavigateReviewItem != null) data.NavigateReviewItem(next);
        }

        private void PaintComment(object sender, DrawItemEventArgs args)
        {
            if (args.Index < 0 || args.Index >= commentItems.Count) return;
            args.DrawBackground();
            var entry = commentItems[args.Index];
            var selected = (args.State & DrawItemState.Selected) != 0;
            var accent = entry.IsComment ? Color.SteelBlue :
                entry.RevisionType == 1 || entry.RevisionType == 15 || entry.RevisionType == 16 ? Color.SeaGreen :
                entry.RevisionType == 2 || entry.RevisionType == 14 || entry.RevisionType == 17 ? Color.Firebrick : Color.DarkViolet;
            using (var stripe = new SolidBrush(accent))
                args.Graphics.FillRectangle(stripe, args.Bounds.X, args.Bounds.Y + 3, 3, args.Bounds.Height - 6);
            var color = selected ? SystemColors.HighlightText : SystemColors.WindowText;
            var title = new Rectangle(args.Bounds.X + 8, args.Bounds.Y + 4, args.Bounds.Width - 16, Font.Height + 2);
            TextRenderer.DrawText(args.Graphics, entry.Caption, Font, title,
                selected ? color : accent, TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            var body = new Rectangle(title.X, title.Bottom + 2, title.Width, 2 * Font.Height + 4);
            TextRenderer.DrawText(args.Graphics, entry.Text.Replace("\r", " ").Replace("\n", " "), Font, body, color,
                TextFormatFlags.WordBreak | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            args.DrawFocusRectangle();
        }

        internal void UpdateWorkingState()
        {
            var text = "";
            try
            {
                if (data != null)
                    switch (data.WorkingState())
                    {
                        case WorkingCopyState.NoCommit: text = "尚无提交"; break;
                        case WorkingCopyState.Unsaved: text = "有未保存修改"; break;
                        case WorkingCopyState.Changed: text = "有未提交修改"; break;
                        case WorkingCopyState.Clean: text = "与最新提交一致"; break;
                    }
            }
            catch (COMException) { return; }
            catch (InvalidComObjectException) { return; }
            catch (IOException) { text = "暂无法读取修改状态"; }
            catch (UnauthorizedAccessException) { text = "暂无法读取修改状态"; }
            catch (InvalidOperationException) { text = "暂无法读取修改状态"; }
            if (working.Text == text) return;
            working.Text = text;
            hints.SetToolTip(working, text);
        }

        private void LoadGraph(string[] keep)
        {
            try
            {
                var selectedBranch = filter.SelectedIndex > 0 ? (string)filter.SelectedItem : null;
                ShowGraph(data.Load(selectedBranch), keep ?? new string[0]);
                hints.SetToolTip(status, null);
            }
            catch (Exception error)
            {
                ShowGraph(new List<Revision>(), new string[0]);
                status.Text = error.Message;
                hints.SetToolTip(status, error.Message);
            }
        }

        private void ShowGraph(List<Revision> entries, string[] keep)
        {
            updating = true;
            try
            {
                graph.RowCount = 0;
                revisions = entries;
                rows = CommitGraph.Layout(entries, data == null ? null : data.Head);
                using (var graphics = CreateGraphics())
                {
                    var lanes = rows.Select(row => row.Columns).DefaultIfEmpty(1).Max();
                    graph.Columns[0].Width = Math.Max(30, (int)Math.Ceiling((15 + 14 * lanes) * graphics.DpiX / 96));
                }
                graph.RowCount = entries.Count;
                if (entries.Count > 0)
                {
                    var index = Math.Max(0, entries.FindIndex(item => keep.Contains(item.Id)));
                    graph.CurrentCell = graph.Rows[index].Cells[1];
                    graph.ClearSelection();
                    graph.Rows[index].Selected = true;
                    foreach (var id in keep)
                    {
                        var i = entries.FindIndex(item => item.Id == id);
                        if (i >= 0) graph.Rows[i].Selected = true;
                    }
                }
            }
            finally { updating = false; }
            ShowSelection();
        }

        private void PaintRow(object sender, DataGridViewCellPaintingEventArgs args)
        {
            if (args.RowIndex < 0 || args.RowIndex >= rows.Count) return;
            if (args.ColumnIndex == 0)
            {
                CommitGraph.Paint(args, rows[args.RowIndex], 14);
                return;
            }
            args.PaintBackground(args.ClipBounds, true);
            var revision = revisions[args.RowIndex];
            var selected = (args.State & DataGridViewElementStates.Selected) != 0;
            var area = args.CellBounds;
            var height = Font.Height + 1;
            var top = area.Top + 5;
            var flags = TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix;
            var saved = args.Graphics.Save();
            try
            {
                args.Graphics.SetClip(args.ClipBounds, System.Drawing.Drawing2D.CombineMode.Intersect);
                TextRenderer.DrawText(args.Graphics, Title(revision), Font, new Rectangle(area.X + 3, top, area.Width - 8, height),
                    selected ? SystemColors.HighlightText : SystemColors.WindowText, flags | TextFormatFlags.PreserveGraphicsClipping);
                TextRenderer.DrawText(args.Graphics, Metadata(revision), Font, new Rectangle(area.X + 3, top + height, area.Width - 8, height),
                    selected ? SystemColors.HighlightText : SystemColors.GrayText, flags | TextFormatFlags.PreserveGraphicsClipping);
                PaintReferences(args.Graphics, revision.References,
                    new Rectangle(area.X + 3, top + 2 * height + 2, area.Width - 8, height + 2), selected);
            }
            finally { args.Graphics.Restore(saved); }
            args.Handled = true;
        }

        private void PaintReferences(Graphics graphics, string[] references, Rectangle area, bool selected)
        {
            var x = area.Left;
            foreach (var reference in references)
            {
                var name = reference.Replace("HEAD -> ", "").Replace("tag: ", "");
                var remaining = area.Right - x;
                if (remaining < 30) break;
                var width = Math.Min(remaining, TextRenderer.MeasureText(name, Font).Width + 34);
                var box = new Rectangle(x, area.Top, width - 3, area.Height - 1);
                var color = selected ? SystemColors.HighlightText : Color.SlateGray;
                using (var fill = new SolidBrush(selected ? SystemColors.Highlight : Color.AliceBlue))
                    graphics.FillRectangle(fill, box);
                using (var pen = new Pen(color))
                {
                    graphics.DrawRectangle(pen, box);
                    var unit = box.Height / 16f;
                    var left = x + 4 * unit;
                    var top = box.Top + 3 * unit;
                    graphics.DrawPolygon(pen, new[] { new PointF(left, top), new PointF(left + 6 * unit, top),
                        new PointF(left + 11 * unit, top + 5 * unit), new PointF(left + 6 * unit, top + 10 * unit),
                        new PointF(left, top + 10 * unit) });
                    graphics.DrawEllipse(pen, left + 2 * unit, top + 3 * unit, 2 * unit, 2 * unit);
                }
                TextRenderer.DrawText(graphics, name, Font, new Rectangle(x + 20, box.Top, Math.Max(1, box.Width - 22), box.Height),
                    color, TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter |
                        TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsClipping);
                x += width + 3;
            }
        }

        private static string Title(Revision revision)
        {
            return revision.Subject;
        }

        private static string Metadata(Revision revision, string dateFormat = "MM-dd HH:mm")
        {
            return revision.Id.Substring(0, 8) + "  ·  " + revision.Author + "  ·  " + revision.Date.ToLocalTime().ToString(dateFormat);
        }

        private Revision[] Selected()
        {
            return graph.SelectedRows.Cast<DataGridViewRow>().OrderBy(row => row.Index)
                .Where(row => row.Index >= 0 && row.Index < revisions.Count).Select(row => revisions[row.Index]).ToArray();
        }

        private void ShowSelection()
        {
            var count = Selected().Length;
            preview.Enabled = count == 1;
            compare.Enabled = count == 1 || count == 2;
            compare.Text = count == 2 ? "对比所选版本" : "对比当前稿";
            status.Text = revisions.Count == 0 ? "提交版本后，历史会显示在这里。"
                : (count > 2 ? "对比时请选择一条或两条提交。" : revisions.Count + " 条提交  ·  双击预览");
        }

        private void OpenSelected(bool diff)
        {
            var selected = Selected();
            if (selected.Length == 0 || selected.Length > (diff ? 2 : 1)) return;
            openVersion(DocumentPath, diff ? HistoryAction.Compare : HistoryAction.Preview,
                selected[selected.Length - 1].Id, selected.Length == 2 ? selected[0].Id : null);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                workingTimer.Dispose();
                hints.Dispose();
                branch.Font.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
