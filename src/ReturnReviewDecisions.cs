using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace WordGit
{
    internal static class ReturnReviewDecisions
    {
        private sealed class Span
        {
            internal DocumentReviewItem Item;
            internal int Start, End;

            internal bool SameStory(Span other)
            {
                return Item.StoryType == other.Item.StoryType && Item.StoryIndex == other.Item.StoryIndex;
            }

            internal bool Touches(Span other, bool separateAdjacentInsertions)
            {
                if (!SameStory(other)) return false;
                if (Start == End && other.Start == other.End) return Start == other.Start;
                if (!separateAdjacentInsertions && (Start == End || other.Start == other.End))
                    return Start <= other.End && other.Start <= End;
                return Start < other.End && other.Start < End;
            }

            internal bool SameProposal(Span other)
            {
                return SameStory(other) && Start == other.Start && End == other.End &&
                    Item.RevisionType == other.Item.RevisionType && Item.RangeText == other.Item.RangeText &&
                    Item.Author == other.Item.Author && Item.Date == other.Item.Date &&
                    Item.RevisionFormat == other.Item.RevisionFormat;
            }
        }

        private static List<Span> AtBaseline(IEnumerable<DocumentReviewItem> items)
        {
            var spans = new List<Span>();
            foreach (var story in items.Where(item => !item.IsComment).GroupBy(item => new { item.StoryType, item.StoryIndex }))
            {
                var insertions = story.Where(item => item.RevisionType == 1).OrderBy(item => item.ScopeStart).ToArray();
                foreach (var item in story)
                {
                    // In Word's markup coordinates, prior insertions take space and tracked deletions
                    // still occupy their original range. Project both ends onto the unedited story.
                    var start = BaselinePosition(item.ScopeStart, insertions);
                    var end = BaselinePosition(item.ScopeEnd, insertions);
                    spans.Add(new Span { Item = item, Start = start, End = end });
                }
            }
            return spans;
        }

        private static int BaselinePosition(int position, IEnumerable<DocumentReviewItem> insertions)
        {
            var total = 0;
            var previousEnd = 0;
            foreach (var insertion in insertions)
            {
                if (insertion.ScopeStart >= position) break;
                var start = System.Math.Max(previousEnd, insertion.ScopeStart);
                var end = System.Math.Min(position, insertion.ScopeEnd);
                if (end <= start) continue;
                total += end - start;
                previousEnd = end;
            }
            return position - total;
        }

        internal static void RemoveRepeated(WordDocuments word, object proposal, IEnumerable<List<DocumentReviewItem>> earlier)
        {
            var previous = earlier.SelectMany(AtBaseline).ToList();
            if (previous.Count == 0) return;
            var incoming = AtBaseline(WordDocuments.ReviewItems(proposal, includeRevisionDates: true));
            // Adjacent records can be part of one replacement; retain the group if any record changed.
            var repeated = new HashSet<Span>(incoming.Where(span =>
                (span.Item.RevisionType == 1 || span.Item.RevisionType == 2 ||
                    (span.Item.Kind == "格式" && !System.String.IsNullOrWhiteSpace(span.Item.RevisionFormat))) &&
                previous.Any(old => span.SameProposal(old))));
            Apply(word, proposal, KeepSeparate(incoming, repeated), false);
        }

        internal static bool PrepareConflictProposal(WordDocuments word, object current, IEnumerable<DocumentReviewItem> incoming,
            bool separateAdjacentInsertions)
        {
            var ours = AtBaseline(WordDocuments.ReviewItems(current));
            var proposal = AtBaseline(incoming);
            var separate = KeepSeparate(ours.Concat(proposal).ToList(), new HashSet<Span>(ours), separateAdjacentInsertions);
            var conflicts = ours.Except(separate).ToArray();
            if (conflicts.Length == 0) return false;
            Apply(word, current, conflicts, false);
            return true;
        }

        internal static void KeepCurrentChanges(WordDocuments word, object merged, string currentAuthor,
            List<DocumentReviewItem> continuedProposal = null)
        {
            var changes = AtBaseline(WordDocuments.ReviewItems(merged));
            var retained = new HashSet<Span>(changes.Where(span => span.Item.Author == currentAuthor));
            Apply(word, merged, KeepSeparate(changes, retained, continuedProposal != null), true);

            // Combining two comparisons can add a tracked final paragraph mark under the local
            // user name. Preserve explicit paragraph insertions from the actual return.
            if (continuedProposal == null || continuedProposal.Any(IsParagraphMark)) return;
            var end = (int)((dynamic)merged).Content.End;
            var extra = WordDocuments.ReviewItems(merged).Where(item => IsParagraphMark(item) &&
                item.StoryType == 1 && item.ScopeEnd >= end - 1 && item.Author != currentAuthor &&
                !continuedProposal.Any(proposal => !proposal.IsComment && proposal.Author == item.Author)).ToArray();
            if (extra.Length == 1) word.ResolveReviewItem(merged, extra[0], false);
        }

        private static bool IsParagraphMark(DocumentReviewItem item)
        {
            return !item.IsComment && item.RevisionType == 1 && item.RangeText == "\r";
        }

        private static IEnumerable<Span> KeepSeparate(List<Span> changes, HashSet<Span> candidates, bool separateAdjacentInsertions = false)
        {
            while (candidates.Count > 0)
            {
                var overlapping = candidates.Where(span => changes.Any(other =>
                    !candidates.Contains(other) && span.Touches(other, separateAdjacentInsertions))).ToArray();
                if (overlapping.Length == 0) break;
                candidates.ExceptWith(overlapping);
            }
            return candidates;
        }

        private static void Apply(WordDocuments word, object document, IEnumerable<Span> spans, bool accept)
        {
            foreach (var span in spans.OrderByDescending(value => value.Item.ScopeStart)
                .ThenByDescending(value => value.Item.Index).ToArray())
                word.ResolveReviewItem(document, span.Item, accept);
        }

        internal static List<DocumentComment> KeepResolvedComments(object merged, IEnumerable<DocumentComment> current)
        {
            var unresolved = new List<DocumentComment>();
            var resolved = current.Where(comment => comment.Resolved).ToArray();
            if (resolved.Length == 0) return unresolved;
            var comments = WordDocuments.Comments(merged, true);
            foreach (var candidate in comments.Where(comment => !comment.Resolved))
            {
                if (resolved.Count(old => SameComment(old, candidate)) == 1 &&
                    comments.Count(other => SameComment(other, candidate)) == 1 &&
                    !TryResolveComment(merged, candidate.Index))
                    unresolved.Add(candidate);
            }
            return unresolved;
        }

        private static bool TryResolveComment(object merged, int index)
        {
            dynamic comments = null, comment = null;
            try
            {
                comments = ((dynamic)merged).Comments;
                comment = comments[index];
                comment.Done = true;
                return (bool)comment.Done;
            }
            catch (COMException) { return false; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return false; }
            finally { WordDocuments.Release(comment); WordDocuments.Release(comments); }
        }

        private static bool SameComment(DocumentComment left, DocumentComment right)
        {
            return left.Author == right.Author && left.Text == right.Text && left.ScopeText == right.ScopeText &&
                left.Date == right.Date;
        }
    }
}

