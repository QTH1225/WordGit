// WordGit P0 probe for WPS Writer JS add-ins.
//
// Runs once after the add-in loads (and again from the ribbon) and answers the questions that
// decide whether WordGit can move to WPS: which origin the add-in runs under, whether it can talk
// to a loopback service, and which document APIs WordGit depends on actually work.
// Every step is reported to the probe service as it finishes, and the full result list is also
// saved as a text file through WPS itself, so results survive even when the network is blocked.

var SERVICE = "http://127.0.0.1:47820";
var RUN_ID = String(Date.now());
var results = [];
var seq = 0;
var probeRunning = false;

function app() {
    return window.Application || (window.wps && window.wps.WpsApplication && window.wps.WpsApplication());
}

function mode() {
    return location.protocol === "http:" || location.protocol === "https:" ? "online" : "offline";
}

function baseUrl() {
    var href = decodeURI(location.href.split("?")[0].split("#")[0]);
    return href.substring(0, href.lastIndexOf("/") + 1);
}

// Native WPS objects do not survive JSON.stringify; keep only plain data.
function plain(value, depth) {
    depth = depth || 0;
    if (value === null || value === undefined) return value === null ? null : "(undefined)";
    var type = typeof value;
    if (type === "string" || type === "number" || type === "boolean") return value;
    if (type === "function") return "(function)";
    if (depth > 3) return "(" + type + ")";
    if (Array.isArray(value)) return value.slice(0, 50).map(function (item) { return plain(item, depth + 1); });
    if (type === "object") {
        var out = {};
        var keys = Object.keys(value);
        if (keys.length === 0) return "(native object)";
        keys.slice(0, 80).forEach(function (key) {
            try { out[key] = plain(value[key], depth + 1); } catch (e) { out[key] = "(throws: " + errorText(e) + ")"; }
        });
        return out;
    }
    return "(" + type + ")";
}

function errorText(e) {
    if (!e) return String(e);
    return String(e.message || e.description || e);
}

function post(payload) {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open("POST", SERVICE + "/api/v1/report", true);
        xhr.setRequestHeader("Content-Type", "application/json");
        xhr.send(JSON.stringify(payload));
    } catch (e) {
        console.log("report failed", e);
    }
}

function report(name, data) {
    var entry = { run: RUN_ID, seq: ++seq, mode: mode(), step: name, data: plain(data) };
    results.push(entry);
    post(entry);
    return entry;
}

function step(name, fn) {
    if (skipped.indexOf(name) >= 0) {
        report(name, { ok: null, skipped: true });
        return undefined;
    }
    var started = Date.now();
    try {
        var value = fn();
        report(name, { ok: true, ms: Date.now() - started, value: value });
        return value;
    } catch (e) {
        report(name, { ok: false, ms: Date.now() - started, error: errorText(e), stack: e && e.stack ? String(e.stack) : "" });
        return undefined;
    }
}

async function asyncStep(name, fn) {
    var started = Date.now();
    try {
        var value = await fn();
        report(name, { ok: true, ms: Date.now() - started, value: value });
        return value;
    } catch (e) {
        report(name, { ok: false, ms: Date.now() - started, error: errorText(e) });
        return undefined;
    }
}

function listKeys(object) {
    var keys = [];
    try { for (var key in object) { keys.push(key); if (keys.length >= 300) break; } } catch (e) { keys.push("(enumeration throws: " + errorText(e) + ")"); }
    return keys;
}

function join(folder, name) {
    var separator = folder.indexOf("\\") >= 0 ? "\\" : "/";
    return folder.replace(/[\\/]+$/, "") + separator + name;
}

function saveAs(document, path) {
    var attempts = [
        ["SaveAs2(path, 16)", function () { document.SaveAs2(path, 16); }],
        ["SaveAs2(path, 12)", function () { document.SaveAs2(path, 12); }],
        ["SaveAs(path, 12)", function () { document.SaveAs(path, 12); }],
        ["SaveAs2(path)", function () { document.SaveAs2(path); }]
    ];
    var errors = [];
    for (var i = 0; i < attempts.length; i++) {
        try { attempts[i][1](); return { via: attempts[i][0], fullName: document.FullName, errors: errors }; }
        catch (e) { errors.push(attempts[i][0] + ": " + errorText(e)); }
    }
    throw new Error("all save variants failed: " + errors.join(" | "));
}

function closeQuietly(document) {
    if (!document) return;
    try { document.Close(0); } catch (e) { /* already closed */ }
}

function openDocument(path, readOnly) {
    return app().Documents.Open(path, false, !!readOnly);
}

function describeRevisionDocument(document) {
    var info = { name: document.Name, fullName: document.FullName, revisions: document.Revisions.Count, comments: document.Comments.Count, items: [] };
    var count = Math.min(info.revisions, 8);
    for (var i = 1; i <= count; i++) {
        var item = {};
        try {
            var revision = document.Revisions.Item(i);
            item.type = revision.Type;
            item.author = revision.Author;
            item.text = String(revision.Range.Text).substring(0, 80);
            try { item.page = revision.Range.Information(3); } catch (e) { item.page = "(throws: " + errorText(e) + ")"; }
        } catch (e) { item.error = errorText(e); }
        info.items.push(item);
    }
    return info;
}

function compareVariant(name, invoke, outputPath) {
    return step(name, function () {
        var a = null, b = null, result = null;
        try {
            a = openDocument(paths.a, false);
            b = openDocument(paths.b, false);
            result = invoke(a, b);
            if (!result) result = app().ActiveDocument;
            var info = describeRevisionDocument(result);
            if (outputPath) info.saved = saveAs(result, outputPath);
            return info;
        } finally {
            closeQuietly(result);
            closeQuietly(b);
            closeQuietly(a);
        }
    });
}

var paths = {};
var eventLog = [];
// <work>/probe-options.txt, one option per line, used to bisect host crashes:
//   skip <step>   skip that step
//   pause <ms>    yield to the host event loop for <ms> between document steps
var skipped = [];
var pauseMs = 0;

function pause() {
    return new Promise(function (resolve) { setTimeout(resolve, pauseMs); });
}

async function runProbe() {
    if (probeRunning) return;
    probeRunning = true;
    RUN_ID = String(Date.now());
    results = [];
    seq = 0;
    try {
        await runProbeSteps();
    } finally {
        probeRunning = false;
    }
}

async function runProbeSteps() {
    var App = app();
    report("start", { href: location.href, origin: location.origin, protocol: location.protocol });

    step("env", function () {
        return {
            userAgent: navigator.userAgent,
            platform: navigator.platform,
            hasWps: typeof window.wps,
            hasApplication: typeof window.Application,
            appName: App.Name,
            version: App.Version,
            build: App.Build,
            path: App.Path,
            userName: App.UserName,
            documents: App.Documents.Count,
            appKeys: listKeys(App),
            wpsKeys: listKeys(window.wps)
        };
    });

    var ping = await asyncStep("net.fetch.ping", async function () {
        var response = await fetch(SERVICE + "/api/v1/ping", { cache: "no-store" });
        return { status: response.status, body: await response.json() };
    });
    await asyncStep("net.xhr.ping", function () {
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open("GET", SERVICE + "/api/v1/ping", true);
            xhr.onload = function () { resolve({ status: xhr.status, body: xhr.responseText }); };
            xhr.onerror = function () { reject(new Error("xhr error, status " + xhr.status)); };
            xhr.send();
        });
    });
    await asyncStep("net.fetch.post-json", async function () {
        var response = await fetch(SERVICE + "/api/v1/report", {
            method: "POST", headers: { "Content-Type": "application/json", "X-WordGit-Token": "probe" },
            body: JSON.stringify({ run: RUN_ID, step: "net.fetch.post-json.body" })
        });
        return { status: response.status };
    });

    var work = ping && ping.body && ping.body.workDir ? ping.body.workDir : "/tmp";
    paths = {
        a: join(work, "probe-a.docx"),
        b: join(work, "probe-b.docx"),
        compared: join(work, "probe-compared.docx"),
        merged: join(work, "probe-merged.docx"),
        summary: join(work, "probe-summary-" + mode() + ".txt")
    };
    report("paths", paths);

    try {
        var optionsFile = join(work, "probe-options.txt");
        if (window.wps.FileSystem.Exists(optionsFile)) {
            String(window.wps.FileSystem.ReadFile(optionsFile)).split(/\r?\n/).forEach(function (line) {
                var parts = line.trim().split(/\s+/);
                if (parts[0] === "skip" && parts[1]) skipped.push(parts[1]);
                if (parts[0] === "pause" && parts[1]) pauseMs = parseInt(parts[1], 10) || 0;
            });
        }
    } catch (e) { skipped = []; pauseMs = 0; }
    report("options", { skip: skipped, pauseMs: pauseMs });

    step("events.register", function () {
        var registered = {};
        ["DocumentNew", "DocumentOpen", "DocumentBeforeSave", "DocumentBeforeClose", "DocumentAfterClose", "WindowActivate", "DocumentChange"].forEach(function (name) {
            try {
                App.ApiEvent.AddApiEventListener(name, function () { eventLog.push({ event: name, at: Date.now() }); });
                registered[name] = true;
            } catch (e) { registered[name] = errorText(e); }
        });
        return registered;
    });

    step("doc.create-a", function () {
        var d = App.Documents.Add();
        try {
            d.Content.Text = "WordGit 探测文档\r第一段：原稿内容。\r第二段：这一段会在返稿中被修改。\r第三段：这一段会被删除。";
            return saveAs(d, paths.a);
        } finally { closeQuietly(d); }
    });

    await pause();
    step("doc.create-b", function () {
        var d = App.Documents.Add();
        try {
            d.Content.Text = "WordGit 探测文档\r第一段：原稿内容。\r第二段：这一段已经在返稿中修改过。\r新增段落：返稿新加的内容。";
            var comment = null;
            try { comment = d.Comments.Add(d.Paragraphs.Item(2).Range, "老师的批注：请补充数据。"); } catch (e) { comment = "(Comments.Add throws: " + errorText(e) + ")"; }
            var saved = saveAs(d, paths.b);
            saved.commentAdded = typeof comment === "string" ? comment : true;
            saved.comments = d.Comments.Count;
            return saved;
        } finally { closeQuietly(d); }
    });

    await pause();
    step("doc.open-readonly", function () {
        var d = openDocument(paths.a, true);
        try { return { readOnly: d.ReadOnly, fullName: d.FullName, paragraphs: d.Paragraphs.Count }; }
        finally { closeQuietly(d); }
    });

    // WPS 12 follows Word's parameter order, including CompareMoves at index 13; the older WPS
    // reference that omits it is wrong. Arguments are positional because JS has no named ones.
    await pause();
    compareVariant("compare.CompareDocuments(objects)", function (a, b) {
        return App.CompareDocuments(a, b, 2, 1, true, true, true, true, true, true, true, true, true, true, "WordGit", true);
    }, paths.compared);

    await pause();
    compareVariant("compare.MergeDocuments(objects)", function (a, b) {
        return App.MergeDocuments(a, b, 2, 1, true, true, true, true, true, true, true, true, true, true, "原稿", "返稿", 0);
    }, paths.merged);

    await pause();
    step("compare.Document.Compare", function () {
        var a = null, result = null;
        try {
            a = openDocument(paths.a, false);
            var before = App.Documents.Count;
            var returned = a.Compare(paths.b, "WordGit", 2, true, true, false, false, false);
            result = returned || App.ActiveDocument;
            var info = describeRevisionDocument(result);
            info.returnedDocument = !!returned;
            info.documentsBefore = before;
            info.documentsAfter = App.Documents.Count;
            info.resultIsSource = result.FullName === a.FullName;
            return info;
        } finally {
            if (result && a && result.FullName !== a.FullName) closeQuietly(result);
            closeQuietly(a);
        }
    });

    await pause();
    step("revisions.operate", function () {
        var d = openDocument(paths.compared, false);
        try {
            var info = { before: d.Revisions.Count };
            if (info.before > 0) {
                var first = d.Revisions.Item(1);
                info.first = { type: first.Type, author: first.Author, text: String(first.Range.Text).substring(0, 80) };
                try { info.first.page = first.Range.Information(3); } catch (e) { info.first.page = "(throws: " + errorText(e) + ")"; }
                try { first.Range.Select(); info.selectWorks = true; } catch (e) { info.selectWorks = errorText(e); }
                d.Revisions.Item(1).Accept();
                info.afterAccept = d.Revisions.Count;
                if (d.Revisions.Count > 0) { d.Revisions.Item(1).Reject(); info.afterReject = d.Revisions.Count; }
            }
            if (d.Comments.Count > 0) {
                var comment = d.Comments.Item(1);
                info.comment = { author: comment.Author, text: String(comment.Range.Text), scope: String(comment.Scope.Text).substring(0, 80) };
            }
            d.TrackRevisions = true;
            info.trackRevisions = d.TrackRevisions;
            d.Range(0, 0).InsertBefore("修订模式插入。");
            info.afterTrackedInsert = d.Revisions.Count;
            d.AcceptAllRevisions();
            info.afterAcceptAll = d.Revisions.Count;
            return info;
        } finally { closeQuietly(d); }
    });

    await pause();
    step("doc.edit-save-reopen", function () {
        var d = openDocument(paths.a, false);
        var info = {};
        try {
            d.Content.InsertAfter("\r追加段落：编辑后保存。");
            info.savedFlagAfterEdit = d.Saved;
            d.Save();
            info.savedFlagAfterSave = d.Saved;
        } finally { closeQuietly(d); }
        var again = openDocument(paths.a, false);
        try {
            info.reopenedHasEdit = String(again.Content.Text).indexOf("追加段落") >= 0;
            info.fullName = again.FullName;
            info.documents = App.Documents.Count;
            return info;
        } finally { closeQuietly(again); }
    });

    step("ui.taskpane", function () {
        var pane = App.CreateTaskPane(baseUrl() + "ui/taskpane.html");
        pane.DockPosition = 2;
        pane.Visible = true;
        App.PluginStorage.setItem("taskpane_id", pane.ID);
        return { id: pane.ID, visible: pane.Visible, url: baseUrl() + "ui/taskpane.html" };
    });

    step("ui.dialog", function () {
        return App.ShowDialog(baseUrl() + "ui/dialog.html", "WordGit 探测对话框", 420, 300, false);
    });

    step("host.capabilities", function () {
        var info = {};
        [["wps.FileSystem", window.wps && window.wps.FileSystem], ["Application.FileSystem", App.FileSystem],
         ["Application.OAAssist", App.OAAssist], ["Application.Env", App.Env], ["Application.PluginStorage", App.PluginStorage]].forEach(function (pair) {
            info[pair[0]] = pair[1] ? listKeys(pair[1]) : "(missing)";
        });
        try { info.shellExecute = typeof App.OAAssist.ShellExecute; } catch (e) { info.shellExecute = errorText(e); }
        try { info.fileExists = App.FileSystem ? App.FileSystem.Exists(paths.a) : "(no FileSystem)"; } catch (e) { info.fileExists = errorText(e); }
        try {
            var probeFile = join(paths.a.replace(/[\\/][^\\/]*$/, ""), "probe-fs.txt");
            var written = window.wps.FileSystem.WriteFile(probeFile, "WordGit 文件系统探测");
            info.fsWriteRead = { written: written, readBack: window.wps.FileSystem.ReadFile(probeFile) };
        } catch (e) { info.fsWriteRead = errorText(e); }
        return info;
    });

    await new Promise(function (resolve) { setTimeout(resolve, 3000); });
    report("events.observed", eventLog.slice(0, 100));

    await pause();
    step("summary.save-text", function () {
        var d = App.Documents.Add();
        try {
            report("done", { steps: results.length });
            d.Content.Text = results.map(function (entry) { return JSON.stringify(entry); }).join("\r");
            d.SaveAs2(paths.summary, 2);
            return paths.summary;
        } finally { closeQuietly(d); }
    });
}

function OnAddinLoad(ribbonUI) {
    var App = app();
    if (App && typeof App.ribbonUI !== "object") App.ribbonUI = ribbonUI;
    report("addin.loaded", { href: location.href, origin: location.origin });
    setTimeout(runProbe, 5000);
    return true;
}

function OnAction(control) {
    switch (control.Id) {
        case "btnRunProbe":
            runProbe();
            break;
        case "btnShowPane": {
            var App = app();
            var id = App.PluginStorage.getItem("taskpane_id");
            var pane = id ? App.GetTaskPane(id) : null;
            if (!pane) {
                pane = App.CreateTaskPane(baseUrl() + "ui/taskpane.html");
                App.PluginStorage.setItem("taskpane_id", pane.ID);
            }
            pane.Visible = !pane.Visible;
            break;
        }
    }
    return true;
}
