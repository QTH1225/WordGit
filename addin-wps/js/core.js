// WordGit for WPS Writer — shared core used by the hidden main page, the task pane and dialogs.
//
// Talks to the local WordGit service (see crates/wordgit) and wraps the WPS document APIs that
// WordGit needs. Rules learned from probing WPS (docs/RUST_PLAN.md):
//  - In WPS's default single-window mode, opening/closing several documents in one synchronous
//    run crashes WPS. Every document operation here is async and yields to the event loop
//    between steps (`settle()`).
//  - JS can only pass positional arguments; CompareDocuments follows Word's order exactly,
//    including CompareMoves at index 13.
(function (global) {
    "use strict";

    var WG = global.WordGit = global.WordGit || {};
    var DEFAULT_PORT = 47821;

    WG.app = function () {
        return global.Application || (global.wps && global.wps.WpsApplication && global.wps.WpsApplication());
    };

    WG.baseUrl = function () {
        var href = decodeURI(location.href.split("?")[0].split("#")[0]);
        var folder = href.substring(0, href.lastIndexOf("/") + 1);
        // Pages under ui/ resolve against the add-in root.
        return folder.replace(/ui\/$/, "");
    };

    WG.errorText = function (error) {
        if (!error) return "未知错误";
        return String(error.message || error.description || error);
    };

    WG.settle = function (ms) {
        return new Promise(function (resolve) { setTimeout(resolve, ms || 0); });
    };

    // ---- local file access through WPS (ordinary web pages cannot do this) -------------------

    function fs() { return global.wps && global.wps.FileSystem; }

    WG.readText = function (path) {
        var files = fs();
        if (!files || !files.Exists(path)) return null;
        return String(files.ReadFile(path));
    };

    WG.writeComparisonSnapshot = function (doc, path) {
        var xml;
        try { xml = String(doc.WordOpenXML || ""); } catch (e) { xml = ""; }
        if (!xml || !xml.includes("http://schemas.microsoft.com/office/2006/xmlPackage")) {
            throw new Error("当前 WPS 无法生成未保存稿的对比副本。请改为对比两个已提交版本。");
        }
        var files = fs();
        files.WriteFile(path, xml);
        if (!files.Exists(path)) throw new Error("无法写入当前稿的临时对比副本。");
    };

    WG.removeComparisonSnapshot = function (path) {
        try { if (fs().Exists(path)) fs().Remove(path); } catch (e) { /* the service cleans old sessions */ }
    };

    function addinFolder() {
        var href = decodeURI(location.href.split("?")[0].split("#")[0]).replace(/^file:\/\//, "");
        if (/^\/[A-Za-z]:\//.test(href)) href = href.substring(1); // file:///C:/... on Windows
        var folder = href.substring(0, href.lastIndexOf("/"));
        return folder.replace(/\/ui$/, "");
    }

    // ---- service ------------------------------------------------------------------------------

    var service = null;

    // The add-in reads the service file the installer recorded, but an older registration can point
    // at a data folder that no longer runs the service (a stale token then fails every call). The
    // default folder is tried as well so a moved installation still works.
    function serviceFiles() {
        var folder = addinFolder();
        var files = [];
        var config = WG.readText(folder + "/wordgit-config.json");
        if (config) {
            try {
                var parsed = JSON.parse(config);
                if (parsed.serviceFile) files.push(parsed.serviceFile);
                if (parsed.home) files.push(String(parsed.home).replace(/[\\/]+$/, "") + "/service.json");
            } catch (e) { /* fall back to the default location below */ }
        }
        var profile = folder.replace(/[\\/]AppData[\\/]Roaming[\\/].*$/i, "");
        if (profile !== folder) files.push(profile + "/AppData/Local/WordGit/service.json");
        return files;
    }

    function loadServiceInfo() {
        var files = serviceFiles();
        for (var index = 0; index < files.length; index++) {
            var text = WG.readText(files[index]);
            if (!text) continue;
            try {
                var info = JSON.parse(text);
                if (info && info.token) return { base: "http://127.0.0.1:" + (info.port || DEFAULT_PORT), token: info.token };
            } catch (e) { /* try the next candidate */ }
        }
        return null;
    }

    WG.serviceMissing = function () {
        return new Error("没有连接到 WordGit 服务。请确认已运行安装程序（wordgit install），" +
            "或在终端执行 wordgit serve 后重试。");
    };

    WG.call = async function (endpoint, body) {
        for (var attempt = 0; attempt < 2; attempt++) {
            if (!service || attempt > 0) service = loadServiceInfo();
            if (!service) throw WG.serviceMissing();
            var response;
            try {
                response = await fetch(service.base + "/api/v1/" + endpoint, {
                    method: "POST",
                    headers: { "Content-Type": "application/json", "X-WordGit-Token": service.token },
                    body: JSON.stringify(body || {})
                });
            } catch (e) {
                if (attempt === 0) continue; // the service may have restarted with a new port/token
                throw WG.serviceMissing();
            }
            if (response.status === 401 && attempt === 0) continue;
            var data = {};
            try { data = await response.json(); } catch (e) { data = {}; }
            if (!response.ok) throw new Error(data.error || ("WordGit 服务返回错误 " + response.status));
            return data;
        }
        throw WG.serviceMissing();
    };

    WG.newOperationId = function () {
        return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
    };

    // ---- documents ----------------------------------------------------------------------------

    WG.activeDocument = function () {
        var app = WG.app();
        if (!app || app.Documents.Count === 0) throw new Error("请先打开一份 .docx 文档。");
        return app.ActiveDocument;
    };

    // Full path of a saved local .docx (saving first when asked to).
    WG.localPath = function (doc, allowSave) {
        if (!doc.Path && allowSave) doc.Save();
        var name = String(doc.FullName || "");
        if (!doc.Path) throw new Error("请先将这份文档保存为本地 .docx。");
        if (name.indexOf("://") >= 0) throw new Error("WordGit 需要可直接读取的本地 .docx。云文档请先另存一份到本地文件夹。");
        if (!/\.docx$/i.test(name)) throw new Error("WordGit 只支持 .docx。请先另存为 .docx 格式。");
        return name;
    };

    WG.saveDocument = function (doc) {
        var path = WG.localPath(doc, true);
        if (!doc.Saved) {
            if (doc.ReadOnly) throw new Error("文档是只读的，当前修改无法保存。请先另存为 .docx。");
            doc.Save();
        }
        if (!doc.Saved) throw new Error("文档尚未保存成功，未创建提交。");
        return path;
    };

    function samePath(a, b) {
        a = String(a || "").replace(/\\/g, "/");
        b = String(b || "").replace(/\\/g, "/");
        return /^[A-Za-z]:/.test(a) || /^\/Users\//.test(a) ? a.toLowerCase() === b.toLowerCase() : a === b;
    }
    WG.samePath = samePath;

    WG.findDocument = function (path) {
        var app = WG.app();
        for (var i = 1; i <= app.Documents.Count; i++) {
            var doc = app.Documents.Item(i);
            if (samePath(doc.FullName, path)) return doc;
        }
        return null;
    };

    WG.open = async function (path, readOnly) {
        var doc = WG.app().Documents.Open(path, false, !!readOnly);
        await WG.settle(50);
        return doc;
    };

    WG.close = async function (doc) {
        if (!doc) return;
        try { doc.Close(0); } catch (e) { /* already closed */ }
        await WG.settle(50);
    };

    WG.userName = function () {
        try {
            var name = String(WG.app().UserName || "");
            return name && name !== "未知" ? name : "";
        } catch (e) { return ""; }
    };

    // Documents produced by Compare/Merge: the return value when it is a document, else the
    // document that appeared during the call.
    function openNames() {
        var app = WG.app(), names = [];
        for (var i = 1; i <= app.Documents.Count; i++) names.push(app.Documents.Item(i));
        return names;
    }

    function produced(returned, before) {
        if (returned && typeof returned === "object" && returned.Revisions) return returned;
        var now = openNames();
        for (var i = 0; i < now.length; i++) if (before.indexOf(now[i]) < 0) return now[i];
        return WG.app().ActiveDocument;
    }

    // A new document marking the differences from `original` to `revised` as tracked changes by
    // `author`. Existing revisions in the inputs are accepted first unless `keepRevisions`.
    WG.compare = async function (originalPath, revisedPath, author, keepRevisions, options) {
        var app = WG.app();
        var left = null, right = null;
        try {
            left = await WG.open(originalPath, false);
            right = await WG.open(revisedPath, false);
            if (!keepRevisions) {
                left.TrackRevisions = false; left.AcceptAllRevisions();
                right.TrackRevisions = false; right.AcceptAllRevisions();
            }
            await WG.settle(50);
            var before = openNames();
            var returned = app.CompareDocuments(left, right, 2, options ? options.granularity : 1,
                true, true, true, true, true, true, true, true, !options || options.comments, true,
                author || "WordGit", true);
            await WG.settle(80);
            return produced(returned, before);
        } finally {
            await WG.close(right);
            await WG.close(left);
        }
    };


    // Both native comparisons end at the same returned document. Project ranges onto that
    // shared text so a rejected edit from the older return cannot mask a new change elsewhere.
    function reviewSpans(doc, revisedSide) {
        var spans = [], omitted = revisedSide ? [2, 14] : [1, 15];
        for (var i = 1; i <= doc.Revisions.Count; i++) {
            var revision = doc.Revisions.Item(i), range = revision.Range;
            spans.push({ index: i, type: revision.Type, story: range.StoryType, start: range.Start, end: range.End,
                text: String(range.Text || ""), format: String(revision.FormatDescription || "") });
        }
        var removed = spans.filter(function (span) { return omitted.indexOf(span.type) !== -1; });
        function position(offset, story) {
            var end = 0, hidden = 0;
            var ranges = removed.filter(function (span) { return span.story === story && span.start < offset; })
                .sort(function (left, right) { return left.start - right.start; });
            for (var range of ranges) {
                var first = Math.max(end, range.start), last = Math.min(offset, range.end);
                if (first >= last) continue;
                hidden += last - first;
                end = last;
            }
            return offset - hidden;
        }
        return spans.map(function (span) {
            return Object.assign({}, span, { start: position(span.start, span.story), end: position(span.end, span.story) });
        });
    }

    function intersectsReturn(left, right) {
        if (left.story !== right.story) return false;
        if (left.start === left.end) {
            return left.start === right.start || right.start < left.start && left.start < right.end;
        }
        return right.start === right.end
            ? left.start === right.start
            : left.start < right.end && right.start < left.end;
    }

    function revisionKeys(doc) {
        var counts = new Map();
        for (var i = 1; i <= doc.Revisions.Count; i++) {
            var revision = doc.Revisions.Item(i), range = revision.Range;
            var key = JSON.stringify([revision.Type, range.StoryType, revision.Author, String(revision.Date),
                String(range.Text || ""), String(revision.FormatDescription || "")]);
            counts.set(key, (counts.get(key) || 0) + 1);
        }
        return counts;
    }

    function readComments(doc) {
        var items = [];
        for (var i = 1; i <= doc.Comments.Count; i++) {
            var comment = doc.Comments.Item(i);
            items.push({ text: String(comment.Range.Text || ""), anchor: String(comment.Scope.Text || "").replace(/[\u0005\u0007]/g, "").trim(),
                author: String(comment.Author || ""), initials: String(comment.Initial || ""),
                date: String(comment.Date || ""), done: !!comment.Done });
        }
        return items;
    }

    function preserveComments(doc, current, incoming) {
        var existing = readComments(doc), used = new Set(), required = current.slice();
        function matching(notes, note, used) {
            var candidates = [];
            notes.forEach(function (item, index) {
                if ((!used || !used.has(index)) && item.author === note.author && item.text.trim() === note.text.trim()) {
                    candidates.push(index);
                }
            });
            var dated = candidates.filter(function (index) { return note.date && notes[index].date === note.date; });
            if (dated.length === 1) return dated[0];
            var anchored = candidates.filter(function (index) { return notes[index].anchor === note.anchor; });
            if (anchored.length === 1) return anchored[0];
            return candidates.length === 1 && (!note.anchor || !notes[candidates[0]].anchor) ? candidates[0] : -1;
        }
        incoming.forEach(function (note) {
            if (matching(current, note) < 0) required.push(note);
        });
        var tracking = doc.TrackRevisions;
        doc.TrackRevisions = false;
        try {
            required.forEach(function (note) {
                var index = matching(existing, note, used);
                if (index >= 0) {
                    used.add(index);
                    if (note.done && !existing[index].done) doc.Comments.Item(index + 1).Done = true;
                    return;
                }
                var range = doc.Content.Duplicate, find = range.Find;
                find.ClearFormatting();
                if (!note.anchor || !find.Execute(note.anchor, true, false, false, false, false, true, 0, false)) {
                    throw new Error("不能定位批注「" + note.text.trim() + "」的原文。原始返稿已保留。");
                }
                var match = range.Duplicate;
                var rest = doc.Range(match.End, doc.Content.End);
                if (rest.Find.Execute(note.anchor, true, false, false, false, false, true, 0, false)) {
                    throw new Error("批注「" + note.text.trim() + "」的原文出现多次。原始返稿已保留。");
                }
                var added = doc.Comments.Add(match, note.text.trimEnd());
                added.Author = note.author;
                if (note.initials) added.Initial = note.initials;
                if (note.done) added.Done = true;
            });
        } finally {
            doc.TrackRevisions = tracking;
        }
        var finalNotes = readComments(doc), retained = new Set();
        required.forEach(function (note) {
            var index = matching(finalNotes, note, retained);
            if (index < 0 || note.done && !finalNotes[index].done) {
                throw new Error("WPS 未能保留原有批注的状态。原始返稿已保留。");
            }
            retained.add(index);
        });
    }

    // Use only changes introduced by this return. The base of the output is always the
    // current manuscript, including its local changes and the decisions from earlier rounds.
    WG.buildReturnedReview = async function (plan) {
        var old = null, raw = null, current = null, proposal = null, diff = null, output = null;
        var snapshots = plan.previousReturns || [], base = plan.basis, known = [];
        var options = { comments: false, granularity: 0 };
        try {
            raw = await WG.open(plan.incoming, true);
            current = await WG.open(plan.target, true);
            var notes = readComments(raw), currentNotes = readComments(current), incoming = revisionKeys(raw);
            await WG.close(current); current = null;
            await WG.close(raw); raw = null;
            for (var entry of snapshots) {
                old = await WG.open(entry.file, true);
                var marks = revisionKeys(old);
                await WG.close(old); old = null;
                if (marks.size && Array.from(marks).every(function (pair) { return (incoming.get(pair[0]) || 0) >= pair[1]; })) {
                    base = entry.file;
                    break;
                }
            }
            if (base === plan.basis) {
                for (var previous of snapshots) {
                    diff = await WG.compare(plan.basis, previous.file, plan.incomingAuthor, false, options);
                    known = known.concat(reviewSpans(diff, false));
                    await WG.close(diff); diff = null;
                }
            }
            diff = await WG.compare(base, plan.incoming, plan.incomingAuthor, false, options);
            var fromBase = reviewSpans(diff, false), fromReturn = reviewSpans(diff, true);
            var allowed = fromReturn.filter(function (item, i) {
                var change = fromBase[i];
                return !known.some(function (before) {
                    return change.story === before.story && change.type === before.type &&
                        change.start === before.start && change.end === before.end &&
                        change.text === before.text && change.format === before.format;
                });
            });
            await WG.close(diff); diff = null;
            proposal = await WG.compare(plan.target, plan.incoming, plan.incomingAuthor, false, options);
            var proposed = reviewSpans(proposal, true);
            for (var i = proposed.length - 1; i >= 0; i--) {
                var item = proposed[i];
                if (!allowed.some(function (change) { return intersectsReturn(item, change); })) proposal.Revisions.Item(item.index).Reject();
            }
            preserveComments(proposal, currentNotes, notes);
            output = proposal; proposal = null;
            return output;
        } finally {
            for (var doc of [proposal, diff, current, raw, old]) if (doc) await WG.close(doc);
        }
    };


    WG.saveAs = function (doc, path) {
        doc.SaveAs2(path, 16);
    };

    // ---- review items (comments and revisions) -----------------------------------------------

    var KINDS = { 1: "新增", 2: "删除", 3: "格式", 4: "格式", 5: "域变更", 6: "冲突", 7: "冲突", 8: "格式", 9: "替换",
        10: "格式", 11: "格式", 12: "格式", 13: "格式", 14: "移出", 15: "移入", 16: "新增单元格", 17: "删除单元格",
        18: "单元格变更", 19: "单元格变更", 20: "冲突", 21: "冲突" };

    function pageOf(range) {
        try { return range.Information(3); } catch (e) { return 0; }
    }

    WG.reviewItems = function (doc, includeResolvedComments) {
        var items = [];
        var comments = doc.Comments;
        for (var c = 1; c <= comments.Count; c++) {
            var comment = comments.Item(c);
            if (includeResolvedComments === false) {
                try { if (comment.Done) continue; } catch (e) { /* older WPS hosts */ }
            }
            var text = "";
            try { text = String(comment.Range.Text).trim(); } catch (e) { text = ""; }
            items.push({ comment: true, index: c, kind: "批注", author: String(comment.Author || ""),
                text: text, page: pageOf(comment.Scope), start: comment.Scope.Start });
        }
        var revisions = doc.Revisions;
        for (var r = 1; r <= revisions.Count; r++) {
            var revision = revisions.Item(r);
            var range = revision.Range;
            var body = String(range.Text || "");
            var format = "";
            try { format = String(revision.FormatDescription || ""); } catch (e) { format = ""; }
            var description = format ? format + "\n" + body : body;
            items.push({ comment: false, index: r, type: revision.Type, kind: KINDS[revision.Type] || "修订",
                author: String(revision.Author || ""), text: description.trim() || "段落、空白或结构变更",
                rangeText: body, page: pageOf(range), start: range.Start, end: range.End });
        }
        items.sort(function (a, b) { return a.start - b.start; });
        return items;
    };

    WG.reviewSignature = function (doc) {
        return doc.Comments.Count + ":" + doc.Revisions.Count + ":" + doc.Content.End;
    };

    function findRevision(doc, item) {
        var revisions = doc.Revisions;
        if (item.index >= 1 && item.index <= revisions.Count) {
            var candidate = revisions.Item(item.index);
            if (candidate.Type === item.type && String(candidate.Author || "") === item.author &&
                candidate.Range.Start === item.start && candidate.Range.End === item.end && String(candidate.Range.Text || "") === item.rangeText)
                return candidate;
        }
        for (var i = 1; i <= revisions.Count; i++) {
            var revision = revisions.Item(i);
            if (revision.Type === item.type && String(revision.Author || "") === item.author &&
                revision.Range.Start === item.start && String(revision.Range.Text || "") === item.rangeText)
                return revision;
        }
        throw new Error("修订已更新，请刷新批注与修订列表。");
    }

    WG.goToItem = function (doc, item) {
        if (String(WG.activeDocument().FullName) !== String(doc.FullName)) doc.Activate();
        var range = item.comment ? doc.Comments.Item(item.index).Scope : findRevision(doc, item).Range;
        range.Select();
        try { WG.app().ActiveWindow.ScrollIntoView(range, true); } catch (e) { /* not supported everywhere */ }
    };

    WG.resolveItem = function (doc, item, accept) {
        if (doc.ReadOnly) throw new Error("当前文档为只读，无法处理修订。");
        var revision = findRevision(doc, item);
        if (accept) revision.Accept(); else revision.Reject();
    };

    // ---- dialogs ------------------------------------------------------------------------------

    // Opens ui/<page>.html as a WPS dialog; `data` is passed through PluginStorage and the
    // dialog's result comes back the same way.
    WG.dialog = function (page, title, width, height, data) {
        var app = WG.app();
        var key = "wordgit.dialog." + WG.newOperationId();
        app.PluginStorage.setItem(key, JSON.stringify({ data: data || {}, result: null }));
        var scale = global.devicePixelRatio || 1;
        app.ShowDialog(WG.baseUrl() + "ui/" + page + ".html?key=" + key, title, width * scale, height * scale, true);
        return new Promise(function (resolve) {
            var timer = setInterval(function () {
                var raw = app.PluginStorage.getItem(key);
                var state = null;
                try { state = JSON.parse(raw); } catch (e) { state = null; }
                if (state && state.closed) {
                    clearInterval(timer);
                    app.PluginStorage.removeItem(key);
                    resolve(state.result);
                }
            }, 200);
        });
    };

    // Inside a dialog page: input data and a way to finish.
    WG.dialogKey = function () {
        var match = /[?&]key=([^&]+)/.exec(location.search);
        return match ? decodeURIComponent(match[1]) : null;
    };

    WG.dialogData = function () {
        var key = WG.dialogKey();
        try { return JSON.parse(WG.app().PluginStorage.getItem(key)).data || {}; } catch (e) { return {}; }
    };

    WG.finishDialog = function (result) {
        var key = WG.dialogKey();
        if (key) WG.app().PluginStorage.setItem(key, JSON.stringify({ closed: true, result: result === undefined ? null : result }));
        setTimeout(function () { global.close(); }, 30);
    };

    WG.notify = function (message, title) {
        return WG.dialog("message", title || "WordGit", 440, 230, { message: message, kind: "info" });
    };

    WG.confirm = async function (message, title, okLabel) {
        var result = await WG.dialog("message", title || "WordGit", 440, 240, { message: message, kind: "confirm", ok: okLabel || "确定" });
        return result === true;
    };

    WG.ask = function (title, options) {
        return WG.dialog("ask", title, 500, options && options.multiline ? 360 : 250, options || {});
    };

    WG.shortId = function (id) { return id ? String(id).substring(0, 8) : ""; };

    // Native file chooser through the WPS object model (a web <input type=file> hides the path).
    // Returns the chosen path, or null when cancelled or unavailable.
    WG.pickFile = function (title, save, initial) {
        try {
            var dialog = WG.app().FileDialog(save ? 2 : 3); // msoFileDialogSaveAs / msoFileDialogFilePicker
            try { dialog.Title = title; } catch (e) { /* optional */ }
            try { dialog.AllowMultiSelect = false; } catch (e) { /* optional */ }
            if (initial) try { dialog.InitialFileName = initial; } catch (e) { /* optional */ }
            if (!save) {
                try { dialog.Filters.Clear(); dialog.Filters.Add("Word 文档", "*.docx"); } catch (e) { /* optional */ }
            }
            if (dialog.Show() === -1 && dialog.SelectedItems.Count > 0) return String(dialog.SelectedItems.Item(1));
        } catch (e) { /* fall back to typing the path */ }
        return null;
    };

    WG.homePath = function () {
        try { return String(WG.app().Env.GetHomePath()); } catch (e) { return ""; }
    };

    WG.formatDate = function (iso) {
        var date = new Date(iso);
        if (isNaN(date.getTime())) return iso || "";
        function pad(n) { return (n < 10 ? "0" : "") + n; }
        return date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate()) + " " + pad(date.getHours()) + ":" + pad(date.getMinutes());
    };

    WG.escape = function (text) {
        return String(text == null ? "" : text).replace(/[&<>"']/g, function (c) {
            return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c];
        });
    };

    // Broadcast to task panes that history or the document changed.
    WG.touch = function () {
        try { WG.app().PluginStorage.setItem("wordgit.changed", String(Date.now())); } catch (e) { /* ignore */ }
    };
})(window);
