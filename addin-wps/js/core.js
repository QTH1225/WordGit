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

    function addinFolder() {
        var href = decodeURI(location.href.split("?")[0].split("#")[0]).replace(/^file:\/\//, "");
        if (/^\/[A-Za-z]:\//.test(href)) href = href.substring(1); // file:///C:/... on Windows
        var folder = href.substring(0, href.lastIndexOf("/"));
        return folder.replace(/\/ui$/, "");
    }

    // ---- service ------------------------------------------------------------------------------

    var service = null;

    function loadServiceInfo() {
        var config = WG.readText(addinFolder() + "/wordgit-config.json");
        var serviceFile = null;
        if (config) {
            try { serviceFile = JSON.parse(config).serviceFile; } catch (e) { serviceFile = null; }
        }
        var text = serviceFile ? WG.readText(serviceFile) : null;
        if (!text) return null;
        try {
            var info = JSON.parse(text);
            return { base: "http://127.0.0.1:" + (info.port || DEFAULT_PORT), token: info.token };
        } catch (e) { return null; }
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
    WG.compare = async function (originalPath, revisedPath, author, keepRevisions) {
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
            var returned = app.CompareDocuments(left, right, 2, 1, true, true, true, true, true, true, true, true, true, true,
                author || "WordGit", true);
            await WG.settle(80);
            return produced(returned, before);
        } finally {
            await WG.close(right);
            await WG.close(left);
        }
    };

    // Applies every tracked change of `source` into `target` (same base text): WPS's
    // MergeDocuments does not combine two sides' revisions, so the incoming side's changes are
    // replayed as tracked changes by comparing the target with the incoming version.
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

    WG.reviewItems = function (doc) {
        var items = [];
        var comments = doc.Comments;
        for (var c = 1; c <= comments.Count; c++) {
            var comment = comments.Item(c);
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
            if (candidate.Type === item.type && candidate.Range.Start === item.start && candidate.Range.End === item.end)
                return candidate;
        }
        for (var i = 1; i <= revisions.Count; i++) {
            var revision = revisions.Item(i);
            if (revision.Type === item.type && revision.Range.Start === item.start && String(revision.Range.Text || "") === item.rangeText)
                return revision;
        }
        throw new Error("修订已更新，请刷新批注与修订列表。");
    }

    WG.goToItem = function (doc, item) {
        doc.Activate();
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
