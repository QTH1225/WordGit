using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class HistoryPaneData
    {
        public string DocumentPath, HistoryPath, Branch, Head;
        public string[] Branches;
        public Func<string, List<Revision>> Load;
        public List<Revision> InitialHistory;
        public Func<WorkingCopyState> WorkingState;
        public Func<bool> IsSaved;
        public Func<List<DocumentReviewItem>> LoadReviewItems;
        public string MergeTargetBranch;
        public Func<int> DocumentRevisionCount;
        // Cheap change detector for the review list; null means "always reload".
        public Func<string> ReviewSignature;
        public Action<DocumentReviewItem> NavigateReviewItem;
        public Func<DocumentReviewItem, bool, bool> ResolveReviewItem;
        public bool CanResolveRevisions;
        public RemoteState Remote;
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
        private readonly Timer reloadFallback = new Timer { Interval = 700 };
        private bool reloadPending;
        private int reloadPendingGeneration;
        private readonly Button commit = DialogLayout.Button("提交版本");
        private readonly Button importReview = DialogLayout.Button("返稿合并");
        private readonly Label remoteStatus = new Label { AutoSize = false, Dock = DockStyle.Fill, AutoEllipsis = true, AccessibleName = "远程同步状态" };
        private readonly Button linkRemote = DialogLayout.Button("关联");
        private readonly Button fetchRemote = DialogLayout.Button("获取");
        private readonly Button pullRemote = DialogLayout.Button("拉取");
        private readonly Button pushRemote = DialogLayout.Button("推送");
        private readonly Button preview = DialogLayout.Button("预览");
        private readonly Button compare = DialogLayout.Button("对比当前稿");
        private readonly ToolTip hints = new ToolTip();
        private readonly Label commentsTitle = DialogLayout.Label("批注与修订");
        private readonly Button acceptRevision = DialogLayout.Button("接受");
        private readonly Button rejectRevision = DialogLayout.Button("拒绝");
        private readonly Label reviewStatus = DialogLayout.Label("");
        private bool resolvingRevision;
        private bool reviewItemsCurrent;
        private int remainingRevisions;
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
        private string reviewSignature;
        // Polling cadence on the 1.2 s timer: compare the signature every ~5 s, and rebuild the
        // list (which repaginates in Word) only when it changed or every ~30 s as a fallback.
        private const int SignatureTicks = 4, ForcedReviewTicks = 25;
        private string lastReviewError;
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
        private Func<Task<HistoryPaneData>> loadAsync;
        private int reloadGeneration, graphGeneration;
        private bool loading, graphLoading, workingQuery;
        private string graphError;
        private string requestedPath;
        private bool showQueued;
        private readonly Panel content = new Panel { Dock = DockStyle.Fill };
        private readonly PaneLoadingFrame loadingFrame = new PaneLoadingFrame();
        private Action<string, HistoryAction, string, string> openVersion;
        private HistoryPaneData data;
        private List<Revision> revisions = new List<Revision>();
        private List<GraphRow> rows = new List<GraphRow>();
        private bool updating;

        internal string DocumentPath { get { return data == null ? null : data.DocumentPath; } }

        public HistoryPane()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Font = SystemFonts.MessageBoxFont;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            AutoScaleMode = AutoScaleMode.Font;
            Size = new Size(420, 640);
            Controls.Add(content);
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal,
                Size = Size, SplitterDistance = 240, SplitterWidth = 6, Panel1MinSize = 120, Panel2MinSize = 280 };
            content.Controls.Add(split);
            var information = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Padding = new Padding(10, 4, 4, 4) };
            var about = new LinkLabel { Text = "WordGit " + AppInfo.DisplayVersion + " · " + AppInfo.Author,
                AutoSize = true, AccessibleName = "WordGit 版本与作者", Margin = new Padding(0, 0, 14, 2) };
            var updates = new LinkLabel { Text = "检查更新", AutoSize = true,
                AccessibleName = "检查 WordGit 更新", Margin = new Padding(0, 0, 0, 2) };
            about.LinkClicked += (sender, args) => AboutDialog.ShowWindow(DialogPlacement.OwnerFor(this), false);
            updates.LinkClicked += (sender, args) => AboutDialog.ShowWindow(DialogPlacement.OwnerFor(this), true);
            hints.SetToolTip(about, "查看版本、作者和项目链接。");
            hints.SetToolTip(updates, "查询 GitHub 正式发布版本。");
            information.Controls.Add(about);
            information.Controls.Add(updates);
            content.Controls.Add(information);
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
            var layout = DialogLayout.Stack(8);
            layout.Padding = new Padding(10);
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.Margin = new Padding(0);
            layout.MinimumSize = new Size(0, 450);
            split.Panel2.AutoScroll = true;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
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
            branch.Height = Font.Height + 16;
            branch.Margin = new Padding(0, 0, 6, 0);
            var refresh = DialogLayout.Button("刷新");
            StylePaneButton(refresh);
            refresh.Width = 64;
            refresh.Margin = new Padding(0);
            refresh.Click += (sender, args) => Reload();
            header.Controls.Add(branch, 0, 0);
            header.Controls.Add(refresh, 1, 0);
            layout.Controls.Add(header, 0, 0);
            filename.Height = Font.Height + 6;
            filename.Margin = new Padding(0, 4, 0, 8);
            filename.TextAlign = ContentAlignment.MiddleLeft;
            layout.Controls.Add(filename, 0, 1);

            var workingRow = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0, 0, 0, 6) };
            workingRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            workingRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            working.Height = Font.Height + 16;
            StylePaneButton(compareLatest);
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
                if (!Visible || resolvingRevision || loading) return;
                UpdateWorkingState();
                PollReviewItems();
            };
            var remoteFrame = new GroupBox { Text = "远程协作", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 8), Padding = new Padding(8, 6, 8, 8) };
            var remotePanel = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top, ColumnCount = 1, RowCount = 2, Margin = new Padding(0) };
            remotePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            remotePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            remotePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            remoteStatus.Height = Font.Height + 8;
            remoteStatus.TextAlign = ContentAlignment.MiddleLeft;
            remoteStatus.Margin = new Padding(0, 0, 0, 6);
            remotePanel.Controls.Add(remoteStatus, 0, 0);
            var syncControls = new[] { linkRemote, fetchRemote, pullRemote, pushRemote };
            var syncButtons = ButtonGrid(4, syncControls);
            var syncActions = new[] { HistoryAction.LinkRemote, HistoryAction.FetchRemote, HistoryAction.PullRemote, HistoryAction.PushRemote };
            for (var i = 0; i < syncControls.Length; i++)
            {
                var action = syncActions[i];
                syncControls[i].Enabled = false;
                syncControls[i].Click += (sender, args) => openVersion(DocumentPath, action, null, null);
            }
            remotePanel.Controls.Add(syncButtons, 0, 1);
            remoteFrame.Controls.Add(remotePanel);
            layout.Controls.Add(remoteFrame, 0, 3);
            hints.SetToolTip(fetchRemote, "下载远程历史，不覆盖当前稿。");
            hints.SetToolTip(pullRemote, "拉取新版本；双方有修改时生成 Word 修订合并稿。");
            hints.SetToolTip(pushRemote, "推送当前分支的已提交版本。");
            filter.FormattingEnabled = true;
            filter.Format += (sender, args) => {
                var reference = args.ListItem as string;
                if (reference != null && reference.StartsWith(ReferenceText.RemotePrefix, StringComparison.Ordinal))
                    args.Value = "远程 " + ReferenceText.Label(reference);
            };
            filter.Margin = new Padding(0);
            filter.SelectedIndexChanged += (sender, args) => { if (!updating && data != null) LoadGraph(null); };
            layout.Controls.Add(filter, 0, 4);

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
            layout.Controls.Add(graph, 0, 6);

            status.MaximumSize = new Size(StatusWidth(), 0);
            layout.Resize += (sender, args) => FitStatusWidth();
            layout.Controls.Add(status, 0, 7);
            var actions = ButtonGrid(2, commit, importReview, preview, compare);
            actions.Margin = new Padding(0, 8, 0, 0);
            preview.Click += (sender, args) => OpenSelected(false);
            compare.Click += (sender, args) => OpenSelected(true);
            commit.Enabled = false;
            commit.Click += (sender, args) => openVersion(DocumentPath, HistoryAction.Commit, null, null);
            importReview.Enabled = false;
            importReview.Click += (sender, args) => openVersion(DocumentPath, HistoryAction.ImportReview, null, null);
            layout.Controls.Add(actions, 0, 5);
            hints.SetToolTip(commit, "保存当前文档，并提交到当前分支。");
            hints.SetToolTip(importReview, "选择送审版本和返稿，生成修订合并稿，可选择保留原始返稿分支。");
            hints.SetToolTip(preview, "双击提交也可在 Word 中只读打开。");
            hints.SetToolTip(compare, "按 Ctrl 可选择两个历史版本。");
            loadingFrame.ReviewHeight = () => split.Panel1.Height;
            loadingFrame.FooterHeight = () => information.Height;
            Controls.Add(loadingFrame);
            loadingFrame.BringToFront();
            ShowSelection();
            reloadFallback.Tick += (sender, args) => StartQueuedReload(reloadPendingGeneration);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var parameters = base.CreateParams;
                // Composite the child HWNDs together instead of exposing their individual first paints.
                parameters.ExStyle |= 0x02000000; // WS_EX_COMPOSITED
                return parameters;
            }
        }

        internal void PrepareForShow(string path = null)
        {
            requestedPath = path;
            showQueued = path != null;
            ++reloadGeneration;
            ++graphGeneration;
            ShowLoadingFrame();
            CreateControl();
            PerformLayout();
            PrepareChildHandles(content);
            var ready = loadingFrame.Handle;
        }

        internal bool CanReuse(string path)
        {
            return !IsDisposed && !Disposing && String.Equals(requestedPath, path, StringComparison.OrdinalIgnoreCase) &&
                (showQueued || loading || (data != null && String.Equals(data.DocumentPath, path, StringComparison.OrdinalIgnoreCase)));
        }

        // The pane starts loading only once its loading surface has reached the screen, so Word
        // never sits on an unpainted pane while the core and the document are read.
        internal bool SurfacePainted { get { return loadingFrame.Painted; } }

        // Short status lines (the empty-history hint) must fit on one line, and the wrap limit must
        // follow the column width: a fixed cap wraps CJK text mid-phrase on scaled displays and
        // leaves an orphaned tail line.
        internal bool StatusIsSingleLine
        {
            get
            {
                var host = status.Parent;
                if (host == null) return true;
                var available = Math.Max(160, host.ClientSize.Width - status.Margin.Horizontal - 2);
                if (status.MaximumSize.Width != available) return false;
                var width = TextRenderer.MeasureText(status.Text, status.Font, new Size(int.MaxValue, int.MaxValue),
                    TextFormatFlags.NoPadding).Width;
                return width <= available;
            }
        }

        internal void QueueReload()
        {
            var generation = reloadGeneration;
            reloadPending = true;
            reloadPendingGeneration = generation;
            loadingFrame.FirstPaint = () => PostQueuedReload(generation);
            reloadFallback.Stop();
            if (loadingFrame.Painted) { PostQueuedReload(generation); return; }
            // Safety net: a pane that never paints must still load.
            reloadFallback.Interval = 700;
            reloadFallback.Start();
        }

        private void PostQueuedReload(int generation)
        {
            if (!IsHandleCreated || IsDisposed || Disposing) return;
            BeginInvoke(new Action(() => StartQueuedReload(generation)));
        }

        private void StartQueuedReload(int generation)
        {
            reloadFallback.Stop();
            loadingFrame.FirstPaint = null;
            if (!reloadPending || generation != reloadPendingGeneration || !CurrentLoad(generation)) return;
            reloadPending = false;
            showQueued = false;
            Reload();
        }

        private static void PrepareChildHandles(Control root)
        {
            var handle = root.Handle;
            root.PerformLayout();
            foreach (Control child in root.Controls) PrepareChildHandles(child);
        }

        private void ShowLoadingFrame()
        {
            content.Enabled = false;
            loadingFrame.Visible = true;
            loadingFrame.BringToFront();
            loadingFrame.Invalidate();
        }

        private void FinishPresentation()
        {
            SuspendLayout();
            try
            {
                content.Enabled = true;
                loadingFrame.Visible = false;
            }
            finally { ResumeLayout(true); }
            Invalidate(true);
            Update();
        }

        private void StylePaneButton(Button button)
        {
            var height = Font.Height + 16;
            button.AutoSize = false;
            button.MinimumSize = new Size(0, height);
            button.Size = new Size(1, height);
            button.Dock = DockStyle.Fill;
            button.Padding = new Padding(3, 0, 3, 0);
            button.TextAlign = ContentAlignment.MiddleCenter;
            button.AutoEllipsis = true;
            button.UseVisualStyleBackColor = true;
        }

        private TableLayoutPanel ButtonGrid(int columns, params Button[] buttons)
        {
            var rows = (buttons.Length + columns - 1) / columns;
            var grid = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, ColumnCount = columns, RowCount = rows, Margin = new Padding(0) };
            for (var column = 0; column < columns; column++)
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / columns));
            for (var row = 0; row < rows; row++)
                grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            for (var i = 0; i < buttons.Length; i++)
            {
                StylePaneButton(buttons[i]);
                buttons[i].Margin = new Padding(0, 0, i % columns == columns - 1 ? 0 : 6,
                    i / columns == rows - 1 ? 0 : 6);
                grid.Controls.Add(buttons[i], i % columns, i / columns);
            }
            return grid;
        }

        internal void Bind(Func<HistoryPaneData> load, Action<string, HistoryAction, string, string> openVersion)
        {
            this.load = load;
            loadAsync = null;
            ++reloadGeneration;
            ++graphGeneration;
            this.openVersion = openVersion;
            workingTimer.Start();
        }


        internal void BindAsync(Func<Task<HistoryPaneData>> load, Action<string, HistoryAction, string, string> openVersion)
        {
            loadAsync = load;
            this.load = null;
            ++reloadGeneration;
            ++graphGeneration;
            this.openVersion = openVersion;
            workingTimer.Start();
        }

        private bool CurrentLoad(int generation)
        {
            return !IsDisposed && !Disposing && generation == reloadGeneration;
        }

        internal void Reload()
        {
            var pending = ReloadAsync();
        }

        // Stops all background polling; the task pane hosting this control is going away.
        internal void Detach()
        {
            workingTimer.Stop();
            reloadFallback.Stop();
            reloadPending = false;
            loadingFrame.FirstPaint = null;
            ++reloadGeneration;
            ++graphGeneration;
            data = null;
            load = null;
            loadAsync = null;
        }

        private void PollReviewItems()
        {
            ++commentPoll;
            var request = data;
            if (commentPoll % SignatureTicks != 0 || request == null) return;
            var generation = reloadGeneration;
            try
            {
                if (commentPoll % ForcedReviewTicks != 0 && request.ReviewSignature != null)
                {
                    var signature = request.ReviewSignature();
                    if (signature == reviewSignature && reviewItemsCurrent) return;
                }
                if (CurrentLoad(generation) && data == request) ReloadComments();
            }
            catch (Exception error)
            {
                if (CurrentLoad(generation) && data == request) ReviewUnavailable(error);
            }
        }

        internal async Task ReloadAsync()
        {
            UiThread.EnsureContext();
            if (IsDisposed || Disposing || (load == null && loadAsync == null)) return;
            showQueued = false;
            var generation = ++reloadGeneration;
            ++graphGeneration;
            var previousFilter = filter.SelectedIndex > 0 ? (string)filter.SelectedItem : null;
            var previousSelection = Selected();
            var previousPath = DocumentPath;
            if (data == null || loadingFrame.Visible) ShowLoadingFrame();
            loading = true;
            graphLoading = true;
            graphError = null;
            comments.Enabled = reviewFilter.Enabled = false;
            graph.Enabled = false;
            filter.Enabled = false;
            commit.Enabled = importReview.Enabled = compareLatest.Enabled = false;
            linkRemote.Enabled = fetchRemote.Enabled = pullRemote.Enabled = pushRemote.Enabled = false;
            preview.Enabled = compare.Enabled = false;
            reviewItemsCurrent = false;
            UpdateReviewActions();
            status.Text = "正在读取提交记录…";
            try
            {
                var next = loadAsync == null ? load() : await loadAsync();
                if (!CurrentLoad(generation)) return;
                data = next;
                requestedPath = next.DocumentPath;
                if (!String.Equals(previousPath, DocumentPath, StringComparison.OrdinalIgnoreCase))
                {
                    previousFilter = null;
                    previousSelection = new Revision[0];
                }
                updating = true;
                try
                {
                    branch.Text = data.MergeTargetBranch == null ? "当前分支  " + data.Branch : "合入分支  " + data.MergeTargetBranch;
                    filename.Text = Path.GetFileName(data.HistoryPath);
                    hints.SetToolTip(branch, branch.Text);
                    hints.SetToolTip(filename, data.HistoryPath);
                    filter.Items.Clear();
                    filter.Items.Add("全部分支");
                    filter.Items.AddRange(data.Branches);
                    filter.SelectedIndex = previousFilter == null ? 0 : Math.Max(0, filter.Items.IndexOf(previousFilter));
                }
                finally { updating = false; }
                var keep = previousSelection.Select(item => item.Id).ToArray();
                if (previousFilter == null && data.InitialHistory != null) ShowGraph(data.InitialHistory, keep);
                else await LoadGraphAsync(keep);
                if (!CurrentLoad(generation)) return;
            }
            catch (Exception error)
            {
                if (!CurrentLoad(generation)) return;
                AppPaths.LogError("history pane load: " + error);
                data = null;
                branch.Text = "当前分支  —";
                filename.Text = "";
                ShowGraph(new List<Revision>(), new string[0]);
                graphError = error.Message;
                status.Text = error.Message;
                hints.SetToolTip(status, error.Message);
            }
            finally
            {
                if (CurrentLoad(generation))
                {
                    loading = false;
                    graphLoading = false;
                    graph.Enabled = true;
                    comments.Enabled = reviewFilter.Enabled = data != null;
                }
            }
            if (!CurrentLoad(generation)) return;
            filter.Enabled = data != null && data.Branches.Length > 0;
            var merging = data != null && data.MergeTargetBranch != null;
            commit.Text = merging ? "完成合并" : "提交版本";
            commit.Enabled = data != null && !merging;
            hints.SetToolTip(commit, merging ? "处理完剩余修订后，确认并提交合并结果。" : "保存当前文档，并提交到当前分支。");
            importReview.Enabled = data != null && data.Head != null && String.Equals(data.DocumentPath, data.HistoryPath, StringComparison.OrdinalIgnoreCase);
            compareLatest.Enabled = data != null && data.Head != null;
            var original = data != null && String.Equals(data.DocumentPath, data.HistoryPath, StringComparison.OrdinalIgnoreCase);
            remoteStatus.Text = data == null || data.Remote == null ? "尚未关联 GitHub 仓库" : data.Remote.Caption;
            hints.SetToolTip(remoteStatus, data == null || data.Remote == null ? "" : (data.Remote.Url ?? "") + Environment.NewLine + "同步数量以最近一次获取为准。");
            linkRemote.Enabled = original;
            var canFetchRemote = original && data.Remote != null && data.Remote.Linked;
            var canPushOrPull = canFetchRemote && data.Head != null;
            fetchRemote.Enabled = canFetchRemote;
            pullRemote.Enabled = pushRemote.Enabled = canPushOrPull;
            ShowSelection();
            // Reveal the graph as soon as it is ready: the review list needs a slow Word pass,
            // so it must not hold the whole pane on the loading surface.
            if (CurrentLoad(generation)) FinishPresentation();
            if (loadAsync != null) working.Text = "正在检查修改…";
            commentsTitle.Text = "正在读取批注与修订…";
            try
            {
                UpdateWorkingState();
                ReloadComments();
            }
            catch (Exception error)
            {
                if (!CurrentLoad(generation)) return;
                AppPaths.LogError("history pane lists: " + error);
                comments.Enabled = reviewFilter.Enabled = false;
                commentsTitle.Text = "批注与修订 · 稍后重试";
                status.Text = error.Message;
                hints.SetToolTip(status, error.Message);
            }
        }

        internal bool ReloadComments()
        {
            var request = data;
            var generation = reloadGeneration;
            try
            {
                // Read the signature first, so edits made while the list is built trigger a reload.
                var signature = request == null || request.ReviewSignature == null ? null : request.ReviewSignature();
                var entries = request == null || request.LoadReviewItems == null ? new List<DocumentReviewItem>() :
                    request.LoadReviewItems().OrderBy(entry => entry.Page).ThenBy(entry => entry.ScopeStart)
                        .ThenBy(entry => entry.IsComment ? 0 : 1).ToList();
                var merging = request != null && request.MergeTargetBranch != null;
                var wordCount = merging && request.DocumentRevisionCount != null ? request.DocumentRevisionCount() : 0;
                if (!CurrentLoad(generation) || data != request) return false;
                reviewSignature = signature;
                var unchanged = entries.Count == allReviewItems.Count && entries.Zip(allReviewItems,
                    (left, right) => left.Key == right.Key && left.Text == right.Text && left.Author == right.Author &&
                        left.Page == right.Page && left.ScopeStart == right.ScopeStart && left.ScopeEnd == right.ScopeEnd &&
                        left.RevisionType == right.RevisionType).All(equal => equal);
                var commentCount = entries.Count(entry => entry.IsComment);
                var listedRevisions = entries.Count - commentCount;
                remainingRevisions = Math.Max(listedRevisions, wordCount);
                reviewItemsCurrent = true;
                lastReviewError = null;
                comments.Enabled = reviewFilter.Enabled = request != null && !loading;
                commentsTitle.Text = merging ? "剩余修订 " + remainingRevisions + " 条 · 批注 " + commentCount :
                    "批注 " + commentCount + " · 修订 " + listedRevisions;
                if (!unchanged)
                {
                    allReviewItems = entries;
                    FilterReviewItems();
                }
                UpdateReviewActions();
                return true;
            }
            catch (Exception error)
            {
                if (CurrentLoad(generation) && data == request) ReviewUnavailable(error);
                return false;
            }
        }

        private void ReviewUnavailable(Exception error)
        {
            var message = error.GetType().FullName + ": " + error.Message;
            if (lastReviewError != message)
                AppPaths.LogError("review items: " + error);
            lastReviewError = message;
            reviewItemsCurrent = false;
            comments.Enabled = reviewFilter.Enabled = false;
            commentsTitle.Text = "批注与修订 · 稍后重试";
            UpdateReviewActions();
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
            if (data != null && data.MergeTargetBranch != null)
                commit.Enabled = !loading && reviewItemsCurrent && !resolvingRevision && remainingRevisions == 0 && data.CanResolveRevisions;
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
                reviewStatus.Text = next == null ? (remainingRevisions == 0 ? "修订已全部处理。" : "请在 Word 的「审阅」中处理剩余修订。")
                    : accept ? "已接受，定位下一条。" : "已拒绝，定位下一条。";
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

        internal async void UpdateWorkingState()
        {
            UiThread.EnsureContext();
            if (workingQuery || loading) return;
            var request = data;
            var generation = reloadGeneration;
            var text = "";
            workingQuery = true;
            try
            {
                if (request != null)
                {
                    var state = request.IsSaved == null ? request.WorkingState() :
                        (!request.IsSaved() ? WorkingCopyState.Unsaved : await Task.Run(request.WorkingState));
                    if (!CurrentLoad(generation) || data != request) return;
                    // Editing may continue while the saved file is being checked.
                    if (request.IsSaved != null && !request.IsSaved()) state = WorkingCopyState.Unsaved;
                    switch (state)
                    {
                        case WorkingCopyState.NoCommit: text = "尚无提交"; break;
                        case WorkingCopyState.Unsaved: text = "有未保存修改"; break;
                        case WorkingCopyState.Changed: text = "有未提交修改"; break;
                        case WorkingCopyState.Clean: text = "与最新提交一致"; break;
                    }
                }
            }
            catch (COMException) { return; }
            catch (InvalidComObjectException) { return; }
            catch (IOException) { text = "暂无法读取修改状态"; }
            catch (UnauthorizedAccessException) { text = "暂无法读取修改状态"; }
            catch (InvalidOperationException) { text = "暂无法读取修改状态"; }
            catch (Exception error)
            {
                // async void: anything that escapes would surface in Word's message loop.
                AppPaths.LogError("working state: " + error);
                text = "暂无法读取修改状态";
            }
            finally { workingQuery = false; }
            if (!CurrentLoad(generation) || data != request || working.Text == text) return;
            working.Text = text;
            hints.SetToolTip(working, text);
        }

        private void LoadGraph(string[] keep)
        {
            var pending = LoadGraphAsync(keep);
        }

        private async Task LoadGraphAsync(string[] keep)
        {
            UiThread.EnsureContext();
            var request = data;
            if (request == null) return;
            var generation = ++graphGeneration;
            graphError = null;
            var selectedBranch = filter.SelectedIndex > 0 ? (string)filter.SelectedItem : null;
            graphLoading = true;
            graph.Enabled = false;
            preview.Enabled = compare.Enabled = false;
            status.Text = "正在读取提交记录…";
            try
            {
                var entries = loadAsync == null ? request.Load(selectedBranch) :
                    await Task.Run(() => request.Load(selectedBranch));
                if (IsDisposed || Disposing || generation != graphGeneration || data != request) return;
                ShowGraph(entries, keep ?? new string[0]);
                hints.SetToolTip(status, null);
            }
            catch (Exception error)
            {
                if (IsDisposed || Disposing || generation != graphGeneration || data != request) return;
                AppPaths.LogError("history graph: " + error);
                ShowGraph(new List<Revision>(), new string[0]);
                graphError = error.Message;
                status.Text = error.Message;
                hints.SetToolTip(status, error.Message);
            }
            finally
            {
                if (!IsDisposed && !Disposing && generation == graphGeneration && data == request)
                {
                    graphLoading = false;
                    graph.Enabled = true;
                    ShowSelection();
                }
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
            preview.Enabled = !loading && !graphLoading && count == 1;
            compare.Enabled = !loading && !graphLoading && (count == 1 || count == 2);
            compare.Text = count == 2 ? "对比所选版本" : "对比当前稿";
            status.Text = graphError ?? ((loading || graphLoading) ? "正在读取提交记录…" :
                (revisions.Count == 0 ? "提交版本后，历史会显示在这里。"
                : (count > 2 ? "对比时请选择一条或两条提交。" : revisions.Count + " 条提交  ·  双击预览")));
        }

        // 状态行最多占满所在栏宽：短句（如空历史提示）一行放下，长错误消息才换行。
        private int StatusWidth()
        {
            var host = status.Parent;
            var width = host != null && host.ClientSize.Width > 0 ? host.ClientSize.Width : ClientSize.Width;
            return Math.Max(160, width - status.Margin.Horizontal - 2);
        }

        // 只有栏宽真正变化时才改上限，避免设置 MaximumSize 触发多余的布局。
        private void FitStatusWidth()
        {
            var width = StatusWidth();
            if (status.MaximumSize.Width != width) status.MaximumSize = new Size(width, 0);
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
                ++reloadGeneration;
                ++graphGeneration;
                workingTimer.Dispose();
                hints.Dispose();
                branch.Font.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
