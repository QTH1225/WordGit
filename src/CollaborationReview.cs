using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WordGit
{
    internal sealed class CollaborationReview
    {
        internal string File;
        internal List<DocumentReviewItem> Marks;
        internal List<CollaborationComment> Comments;
        internal Dictionary<string, object>[] Conflicts, Decisions;

        internal bool[] Choose(IWin32Window owner)
        {
            var selected = new bool[Conflicts.Length];
            for (var i = 0; i < selected.Length; i++)
            {
                using (var dialog = new Form { Text = "WordGit · 处理冲突 " + (i + 1) + " / " + selected.Length,
                    ShowInTaskbar = false, MinimizeBox = false, MaximizeBox = false, StartPosition = FormStartPosition.CenterParent,
                    AutoScaleMode = AutoScaleMode.Font, Font = SystemFonts.MessageBoxFont,
                    ClientSize = new Size(620, 345), MinimumSize = new Size(510, 305) })
                {
                    var layout = DialogLayout.Stack(3);
                    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    var original = Json.String(Conflicts[i], "original");
                    var heading = String.IsNullOrWhiteSpace(original) ? "双方修改了同一处。请选择这一处要保留的版本。" :
                        "双方修改了同一处 · 原文「" + Preview(original) + "」";
                    layout.Controls.Add(DialogLayout.Label(heading), 0, 0);
                    var choices = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
                    choices.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    choices.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    choices.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    choices.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                    var labels = new[] { "本地版本", "远端版本" };
                    var keys = new[] { "local", "remote" };
                    for (var j = 0; j < 2; j++)
                    {
                        choices.Controls.Add(DialogLayout.Label(labels[j]), j, 0);
                        choices.Controls.Add(new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Multiline = true,
                            ScrollBars = ScrollBars.Vertical, AccessibleName = labels[j],
                            Text = (Json.String(Conflicts[i], keys[j]) ?? "").Replace("\r", Environment.NewLine) }, j, 1);
                    }
                    layout.Controls.Add(choices, 0, 1);
                    var buttons = DialogLayout.Buttons(true);
                    var cancel = DialogLayout.Button("取消合并"); cancel.DialogResult = DialogResult.Cancel;
                    var local = DialogLayout.Button("保留本地"); local.DialogResult = DialogResult.No;
                    var remote = DialogLayout.Button("保留远端"); remote.DialogResult = DialogResult.Yes;
                    buttons.Controls.AddRange(new Control[] { cancel, remote, local });
                    layout.Controls.Add(buttons, 0, 2);
                    dialog.Controls.Add(layout);
                    dialog.CancelButton = cancel;
                    var choice = dialog.ShowDialog(owner);
                    if (choice != DialogResult.Yes && choice != DialogResult.No) return null;
                    selected[i] = choice == DialogResult.Yes;
                }
            }
            return selected;
        }

        private static string Preview(string value)
        {
            value = value.Trim().Replace("\r", " ").Replace("\n", " ");
            return value.Length > 60 ? value.Substring(0, 60) + "…" : value;
        }
    }

    internal sealed class CollaborationComment
    {
        internal string Author, Initial, Text, Anchor, Before, After, Story;
        internal DateTime Date;
        internal bool Resolved;
    }

    internal sealed partial class WordDocuments
    {
        internal CollaborationReview PrepareCollaboration(MergePlan plan, string stagingFile)
        {
            var local = ReadCollaborationDiff(plan.Basis, plan.Target, "本地", null);
            var remote = ReadCollaborationDiff(plan.Basis, plan.Incoming, plan.IncomingAuthor, null);
            var comparison = ReadCollaborationDiff(plan.Target, plan.Incoming, plan.IncomingAuthor, stagingFile);
            var answer = ServiceClient.Connect().Call("merge/collaboration", ServiceClient.Args(
                "local", CollaborationMarks(local), "remote", CollaborationMarks(remote),
                "comparison", CollaborationMarks(comparison)));
            var comments = ReadCollaborationComments(plan.Target);
            var incoming = ReadCollaborationComments(plan.Incoming);
            var matched = new HashSet<int>();
            foreach (var comment in incoming)
            {
                var duplicate = comments.FindIndex(existing => SameComment(existing, comment) && existing.Date == comment.Date
                    && existing.Before == comment.Before && existing.After == comment.After && !matched.Contains(comments.IndexOf(existing)));
                if (duplicate >= 0) matched.Add(duplicate);
                else { matched.Add(comments.Count); comments.Add(comment); }
            }
            return new CollaborationReview { File = stagingFile, Marks = comparison, Comments = comments,
                Conflicts = Json.Array(answer, "conflicts").Cast<Dictionary<string, object>>().ToArray(),
                Decisions = Json.Array(answer, "decisions").Cast<Dictionary<string, object>>().ToArray() };
        }

        private List<DocumentReviewItem> ReadCollaborationDiff(string basis, string revised, string author, string file)
        {
            var doc = ReviewFiles(basis, revised, false, author, true);
            try
            {
                var marks = ReviewItems(doc).Where(item => !item.IsComment).ToList();
                if (file != null) ((dynamic)doc).SaveAs2(FileName: file, FileFormat: 16, AddToRecentFiles: false);
                return marks;
            }
            finally { CloseReviewInput(doc); }
        }

        private static object[] CollaborationMarks(IEnumerable<DocumentReviewItem> marks)
        {
            return marks.Select(mark => (object)new Dictionary<string, object> {
                { "story", mark.StoryType + ":" + mark.StoryIndex }, { "kind", mark.RevisionType },
                { "start", mark.ScopeStart }, { "end", mark.ScopeEnd }, { "text", mark.RangeText }
            }).ToArray();
        }

        internal void CompleteCollaboration(CollaborationReview review, bool[] selected, string output)
        {
            if (selected.Length != review.Conflicts.Length || review.Marks.Count != review.Decisions.Length)
                throw new InvalidOperationException("合并选择不完整，请重新拉取。");
            dynamic doc = Open(review.File, false);
            try
            {
                var reloaded = ReviewItems((object)doc).Where(item => !item.IsComment).ToArray();
                if (reloaded.Length != review.Marks.Count || !reloaded.Zip(review.Marks, (a, b) =>
                    a.Key == b.Key && a.ScopeStart == b.ScopeStart && a.ScopeEnd == b.ScopeEnd &&
                    a.RevisionType == b.RevisionType && a.RangeText == b.RangeText).All(equal => equal))
                    throw new InvalidOperationException("编辑器重新打开后修订范围发生变化，请重新拉取。");
                foreach (var position in Enumerable.Range(0, reloaded.Length).OrderByDescending(i => reloaded[i].ScopeStart)
                    .ThenByDescending(i => reloaded[i].Index))
                {
                    var decision = review.Decisions[position];
                    var accept = decision.ContainsKey("conflict") ? selected[Convert.ToInt32(decision["conflict"])] : Json.Bool(decision, "incoming");
                    ResolveReviewItem((object)doc, reloaded[position], accept);
                }
                if (ReviewItems((object)doc).Any(item => !item.IsComment))
                    throw new InvalidOperationException("编辑器仍有未处理的修订，主稿保持原样。");
                RestoreCollaborationComments(doc, review.Comments);
                doc.SaveAs2(FileName: output, FileFormat: 16, AddToRecentFiles: false);
            }
            finally { CloseReviewInput(doc); }
        }

        private List<CollaborationComment> ReadCollaborationComments(string path)
        {
            dynamic doc = Open(path, true, false);
            try
            {
                PrepareReviewInput((object)doc, false);
                return CaptureCollaborationComments(doc);
            }
            finally { CloseReviewInput(doc); }
        }

        private static List<CollaborationComment> CaptureCollaborationComments(dynamic doc)
        {
            var notes = new List<CollaborationComment>();
            for (var i = 1; i <= (int)doc.Comments.Count; i++)
            {
                dynamic note = doc.Comments[i], scope = note.Scope, story = scope.Duplicate;
                story.Expand(6); // wdStory
                var end = (int)story.End;
                var start = (int)story.Start;
                var anchor = (string)scope.Text ?? "";
                var before = scope.Duplicate;
                before.SetRange(Math.Max(start, (int)scope.Start - 32), scope.Start);
                var after = scope.Duplicate;
                after.SetRange(scope.End, Math.Min(end, (int)scope.End + 32));
                notes.Add(new CollaborationComment { Author = note.Author, Initial = note.Initial, Date = note.Date,
                    Text = ((string)note.Range.Text ?? "").Trim(), Anchor = anchor, Before = before.Text ?? "",
                    After = after.Text ?? "", Resolved = CommentResolved(note), Story = ((int)scope.StoryType).ToString() });
            }
            return notes;
        }

        private static bool SameComment(CollaborationComment left, CollaborationComment right)
        {
            return left.Author == right.Author && left.Text == right.Text && left.Anchor == right.Anchor && left.Story == right.Story;
        }

        private static void RestoreCollaborationComments(dynamic doc, List<CollaborationComment> required)
        {
            doc.TrackRevisions = false;
            List<CollaborationComment> existing = CaptureCollaborationComments(doc);
            var retained = new HashSet<int>();
            var additions = new List<Tuple<CollaborationComment, object>>();
            foreach (var expected in required)
            {
                dynamic scope = FindCollaborationAnchor(doc, expected);
                var start = (int)scope.Start;
                var index = Enumerable.Range(0, existing.Count).Where(n =>
                    !retained.Contains(n) && SameComment(existing[n], expected) &&
                    (int)doc.Comments[n + 1].Scope.Start == start).DefaultIfEmpty(-1).First();
                if (index >= 0)
                {
                    retained.Add(index);
                    if (expected.Resolved && !existing[index].Resolved) doc.Comments[index + 1].Done = true;
                }
                else additions.Add(Tuple.Create(expected, (object)scope));
            }
            for (var i = existing.Count - 1; i >= 0; i--)
                if (!retained.Contains(i)) doc.Comments[i + 1].Delete();
            foreach (var addition in additions)
            {
                var expected = addition.Item1;
                dynamic note = doc.Comments.Add((dynamic)addition.Item2, expected.Text);
                note.Author = expected.Author;
                note.Initial = expected.Initial;
                if (expected.Resolved) note.Done = true;
            }
            List<CollaborationComment> actual = CaptureCollaborationComments(doc);
            if (actual.Count != required.Count ||
                required.Any(expected => !actual.Any(note => SameComment(expected, note) && (!expected.Resolved || note.Resolved))))
                throw new InvalidOperationException("编辑器未能完整保留批注，合并稿未生成。");
        }

        private static object FindCollaborationAnchor(dynamic doc, CollaborationComment note)
        {
            var candidates = new List<Tuple<int, object>>();
            dynamic story = doc.StoryRanges[Int32.Parse(note.Story)];
            while (story != null)
            {
                var body = (string)story.Text ?? "";
                for (var offset = body.IndexOf(note.Anchor, StringComparison.Ordinal); offset >= 0 && offset <= body.Length;
                    offset = offset == body.Length ? -1 : body.IndexOf(note.Anchor, offset + 1, StringComparison.Ordinal))
                {
                    var end = offset + note.Anchor.Length;
                    var score = (body.Substring(Math.Max(0, offset - note.Before.Length), Math.Min(offset, note.Before.Length)) == note.Before ? 1 : 0)
                        + (body.Substring(end, Math.Min(body.Length - end, note.After.Length)) == note.After ? 1 : 0);
                    if (note.Anchor.Length == 0 && score != 2) continue;
                    dynamic anchor = story.Duplicate;
                    anchor.SetRange((int)story.Start + offset, (int)story.Start + end);
                    if (((string)anchor.Text ?? "") == note.Anchor) candidates.Add(Tuple.Create(score, (object)anchor));
                }
                story = story.NextStoryRange;
            }
            if (candidates.Count > 0)
            {
                var best = candidates.Max(item => item.Item1);
                var matches = candidates.Where(item => item.Item1 == best).ToArray();
                if (matches.Length == 1) return matches[0].Item2;
            }
            throw new InvalidOperationException("无法准确定位「" + note.Author + "」的批注原文，合并稿未生成。");
        }
    }
}
