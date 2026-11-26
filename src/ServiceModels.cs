using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;

namespace WordGit
{
    // JSON helpers for the local WordGit service protocol. The service speaks
    // camelCase for requests and most responses, but the pending review and the
    // remote state are serialized snake_case; every reader below names its keys
    // explicitly so the two conventions never get mixed up silently.
    internal static class Json
    {
        private static readonly JavaScriptSerializer Serializer = Create();

        private static JavaScriptSerializer Create()
        {
            var serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = Int32.MaxValue;
            serializer.RecursionLimit = 512;
            return serializer;
        }

        internal static string Send(IDictionary<string, object> body)
        {
            return Serializer.Serialize(body ?? new Dictionary<string, object>());
        }

        internal static Dictionary<string, object> Parse(string text)
        {
            var value = Serializer.DeserializeObject(text) as Dictionary<string, object>;
            if (value == null) throw new InvalidOperationException("WordGit 服务返回的数据无法识别。");
            return value;
        }

        internal static Dictionary<string, object> Object(Dictionary<string, object> value, string key)
        {
            object item;
            if (value == null || !value.TryGetValue(key, out item)) return null;
            return item as Dictionary<string, object>;
        }

        internal static object[] Array(Dictionary<string, object> value, string key)
        {
            object item;
            if (value == null || !value.TryGetValue(key, out item)) return new object[0];
            return item as object[] ?? new object[0];
        }

        internal static string[] Strings(Dictionary<string, object> value, string key)
        {
            var items = Array(value, key);
            var result = new string[items.Length];
            for (var index = 0; index < items.Length; index++) result[index] = items[index] as string;
            return result;
        }

        internal static string String(Dictionary<string, object> value, string key)
        {
            object item;
            if (value == null || !value.TryGetValue(key, out item)) return null;
            return item as string;
        }

        internal static bool Bool(Dictionary<string, object> value, string key)
        {
            object item;
            if (value == null || !value.TryGetValue(key, out item)) return false;
            return item is bool && (bool)item;
        }

        internal static int Int(Dictionary<string, object> value, string key)
        {
            object item;
            if (value == null || !value.TryGetValue(key, out item) || item == null) return 0;
            if (item is int) return (int)item;
            if (item is long) return (int)Math.Min(Int32.MaxValue, (long)item);
            var text = item as string;
            int parsed;
            return text != null && Int32.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : 0;
        }

        internal static DateTimeOffset Date(Dictionary<string, object> value, string key)
        {
            var text = String(value, key);
            DateTimeOffset parsed;
            if (text != null && DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal, out parsed))
                return parsed;
            return DateTimeOffset.MinValue;
        }
    }

    internal sealed class Revision
    {
        public string Id, Author, Subject, Message;
        public string[] Parents, References;
        public DateTimeOffset Date;
        public override string ToString() { return Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm") + "   " + Id.Substring(0, 8) + "   " + Subject + "   · " + Author; }

        internal static Revision From(Dictionary<string, object> value)
        {
            if (value == null) return null;
            return new Revision
            {
                Id = Json.String(value, "id"),
                Parents = Json.Strings(value, "parents"),
                Date = Json.Date(value, "date"),
                Author = Json.String(value, "author"),
                References = Json.Strings(value, "references"),
                Subject = Json.String(value, "subject"),
                Message = Json.String(value, "message"),
            };
        }

        internal static List<Revision> List(Dictionary<string, object> value, string key)
        {
            var items = Json.Array(value, key);
            var result = new List<Revision>(items.Length);
            for (var index = 0; index < items.Length; index++) result.Add(From(items[index] as Dictionary<string, object>));
            return result;
        }
    }

    internal sealed class ReviewImport
    {
        public string Id, Reference;
        public bool Reused;

        internal static ReviewImport From(Dictionary<string, object> value)
        {
            if (value == null) return null;
            return new ReviewImport { Id = Json.String(value, "commit"), Reference = Json.String(value, "reference"), Reused = Json.Bool(value, "reused") };
        }
    }

    internal sealed class RemoteState
    {
        internal string Url, LocalBranch, RemoteBranch, RemoteReference, LocalTip, RemoteTip;
        internal int Ahead, Behind;
        internal bool Unrelated;
        internal bool Linked { get { return !String.IsNullOrEmpty(Url); } }
        internal string Caption
        {
            get {
                if (!Linked) return "尚未关联远程仓库";
                if (RemoteTip == null) return "origin/" + RemoteBranch + " · 尚未获取到该分支";
                if (Unrelated) return "origin/" + RemoteBranch + " · 没有共同历史";
                return "origin/" + RemoteBranch + " · 领先 " + Ahead + " / 落后 " + Behind;
            }
        }

        internal static RemoteState From(Dictionary<string, object> value)
        {
            if (value == null) return null;
            return new RemoteState
            {
                Url = Json.String(value, "url"),
                LocalBranch = Json.String(value, "local_branch"),
                RemoteBranch = Json.String(value, "remote_branch"),
                RemoteReference = Json.String(value, "remote_reference"),
                LocalTip = Json.String(value, "local_tip"),
                RemoteTip = Json.String(value, "remote_tip"),
                Ahead = Json.Int(value, "ahead"),
                Behind = Json.Int(value, "behind"),
                Unrelated = Json.Bool(value, "unrelated"),
            };
        }
    }

    // The pending review recorded for a merge draft, as returned by the service.
    internal sealed class PendingReview
    {
        public string Original, Head, Branch, SourceBranch, SourceTip;
        public bool ApplyToOriginal, RecordOnlyMerge;
        public bool IsPendingMerge { get { return ApplyToOriginal || RecordOnlyMerge || SourceBranch != null; } }

        internal static PendingReview From(Dictionary<string, object> value)
        {
            if (value == null) return null;
            return new PendingReview
            {
                Original = Json.String(value, "original"),
                Head = Json.String(value, "head"),
                Branch = Json.String(value, "branch"),
                SourceBranch = Json.String(value, "source_branch") ?? Json.String(value, "sourceBranch"),
                SourceTip = Json.String(value, "source_tip") ?? Json.String(value, "sourceTip"),
                ApplyToOriginal = Json.Bool(value, "apply_to_original") || Json.Bool(value, "applyToOriginal"),
                RecordOnlyMerge = Json.Bool(value, "record_only_merge") || Json.Bool(value, "recordOnlyMerge"),
            };
        }
    }

    // A prepared replacement the add-in must apply after closing the document.
    internal sealed class ReplacementPlan
    {
        public string Id, Document, ChangesFile, Message;
        public bool Unchanged;

        internal static ReplacementPlan From(Dictionary<string, object> value)
        {
            if (value == null) throw new InvalidOperationException("WordGit 服务没有返回替换计划。");
            return new ReplacementPlan
            {
                Id = Json.String(value, "plan"),
                Document = Json.String(value, "document"),
                ChangesFile = Json.String(value, "changesFile"),
                Message = Json.String(value, "message"),
                Unchanged = Json.Bool(value, "unchanged"),
            };
        }
    }

    // One status response: everything the pane and the command handlers need.
    internal sealed class DocumentStatus
    {
        internal string Document, HistoryPath, Repository, Branch, Head, WorkingState;
        internal bool Session;
        internal string[] Branches;
        internal List<Revision> History;
        internal PendingReview Review;
        internal RemoteState Remote;

        internal static DocumentStatus From(Dictionary<string, object> value)
        {
            var history = Revision.List(value, "history");
            return new DocumentStatus
            {
                Document = Json.String(value, "document"),
                Session = Json.Bool(value, "session"),
                HistoryPath = Json.String(value, "historyPath"),
                Repository = Json.String(value, "repository"),
                Review = PendingReview.From(Json.Object(value, "review")),
                Branch = Json.String(value, "branch"),
                Head = Json.String(value, "head"),
                Branches = Json.Strings(value, "branches"),
                Remote = RemoteState.From(Json.Object(value, "remote")),
                History = history,
                WorkingState = Json.String(value, "workingState"),
            };
        }
    }

    // A merge draft prepared by merge/branch, merge/versions or review/import.
    // A return draft that already went into this branch, extracted for the add-in.
    internal sealed class ReturnFile
    {
        public string Id, File;

        public static ReturnFile From(Dictionary<string, object> value)
        {
            return value == null ? null : new ReturnFile { Id = Json.String(value, "id"), File = Json.String(value, "file") };
        }
    }

    internal sealed class MergePlan
    {
        public string Draft, Basis, Target, Incoming, Notice, IncomingAuthor;
        public string SubmittedBase, TargetBranch, SourceBranch, Message;
        // Set when two historical revisions are merged instead of a branch.
        public string Original, Revised;
        public bool UpToDate, BasisIsTarget, PreserveIncoming;
        public PendingReview Review;
        public List<ReturnFile> PreviousReturns = new List<ReturnFile>();

        public static MergePlan From(Dictionary<string, object> value)
        {
            if (value == null) throw new InvalidOperationException("WordGit 服务没有返回合并计划。");
            var plan = new MergePlan
            {
                Draft = Json.String(value, "draft"),
                Basis = Json.String(value, "basis"),
                Target = Json.String(value, "target"),
                Incoming = Json.String(value, "incoming"),
                Notice = Json.String(value, "notice"),
                Message = Json.String(value, "message"),
                IncomingAuthor = Json.String(value, "incomingAuthor"),
                SubmittedBase = Json.String(value, "submittedBase"),
                TargetBranch = Json.String(value, "targetBranch"),
                SourceBranch = Json.String(value, "sourceBranch"),
                Original = Json.String(value, "original"),
                Revised = Json.String(value, "revised"),
                UpToDate = Json.Bool(value, "upToDate"),
                BasisIsTarget = Json.Bool(value, "basisIsTarget"),
                PreserveIncoming = Json.Bool(value, "preserveIncoming"),
                Review = PendingReview.From(Json.Object(value, "review")),
            };
            foreach (var item in Json.Array(value, "previousReturns"))
            {
                var file = ReturnFile.From(item as Dictionary<string, object>);
                if (file != null && file.File != null) plan.PreviousReturns.Add(file);
            }
            return plan;
        }
    }

    // A pair of files prepared for Word's native comparison.
    internal sealed class ComparePlan
    {
        public string Session, Original, Revised;

        public static ComparePlan From(Dictionary<string, object> value)
        {
            if (value == null) throw new InvalidOperationException("WordGit 服务没有返回对比文件。");
            return new ComparePlan
            {
                Session = Json.String(value, "session"),
                Original = Json.String(value, "original"),
                Revised = Json.String(value, "revised"),
            };
        }
    }

    // References rendered by the pane and dialogs.
    internal static class ReferenceText
    {
        internal const string RemotePrefix = "refs/remotes/origin/";
        internal const string ReviewArchivePrefix = "refs/wordgit/reviews/";

        internal static string Label(string reference)
        {
            if (reference != null && reference.StartsWith(ReviewArchivePrefix, StringComparison.Ordinal)) return "返稿";
            return reference != null && reference.StartsWith(RemotePrefix, StringComparison.Ordinal)
                ? "origin/" + reference.Substring(RemotePrefix.Length) : reference;
        }

        // Branch names collide when they are equal or when one nests folders under the other.
        internal static string Conflict(string name, IEnumerable<string> branches)
        {
            var comparison = StringComparison.OrdinalIgnoreCase;
            string conflict = null;
            foreach (var branch in branches)
            {
                if (!String.Equals(name, branch, comparison) &&
                    !name.StartsWith(branch + "/", comparison) && !branch.StartsWith(name + "/", comparison)) continue;
                conflict = branch;
                break;
            }
            if (conflict == null) return null;
            return "返稿分支「" + name + "」与已有分支「" + conflict + "」冲突。请换一个名称，例如 review-老师审阅。";
        }
    }
}