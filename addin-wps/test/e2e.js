// End-to-end test of the WordGit WPS add-in against the real service, run inside WPS.
// Not part of the shipped add-in: docker/wps-linux injects it into the deployed copy.
// It drives the same functions the ribbon uses (minus the dialogs) and writes one JSON line
// per step to E2E_REPORT.
(function (global) {
    "use strict";
    var WG = global.WordGit;
    var HOME = "/home/ubuntu/e2e";
    var E2E_REPORT = "/out/e2e-report.jsonl";
    var MARKER = HOME + "/run-e2e";
    var lines = [];
    var failed = 0;

    function write() {
        try { global.wps.FileSystem.WriteFile(E2E_REPORT, lines.join("\n") + "\n"); } catch (e) { console.log(e); }
    }

    function record(step, ok, detail) {
        if (!ok) failed++;
        lines.push(JSON.stringify({ step: step, ok: ok, detail: detail === undefined ? null : detail, at: Date.now() }));
        write();
    }

    function check(condition, message) {
        if (!condition) throw new Error("断言失败：" + message);
    }

    async function step(name, body) {
        try {
            var detail = await body();
            record(name, true, detail);
        } catch (error) {
            record(name, false, WG.errorText(error) + (error && error.stack ? "\n" + error.stack : ""));
            throw error;
        }
        await WG.settle(60);
    }

    function text(doc) { return String(doc.Content.Text); }

    async function newDocument(path, content) {
        var doc = WG.app().Documents.Add();
        await WG.settle(50);
        doc.Content.Text = content;
        doc.SaveAs2(path, 16);
        await WG.settle(50);
        return doc;
    }

    async function setText(doc, content) {
        doc.Content.Text = content;
        doc.Save();
        await WG.settle(50);
    }

    // Accepts every tracked change in the active merge draft, as a reviewer would.
    async function acceptAll(doc) {
        doc.AcceptAllRevisions();
        doc.Save();
        await WG.settle(50);
    }

    async function run() {
        var app = WG.app();
        var doc = null, docPath = HOME + "/论文 初稿.docx";
        var v1 = "WordGit 端到端测试\r第一段：初稿内容。\r第二段：送审前的内容。";
        var v2 = "WordGit 端到端测试\r第一段：初稿内容，已补充。\r第二段：送审前的内容。";
        var ids = {};

        await step("service.info", async function () {
            var info = await WG.call("info", {});
            check(info.service === "wordgit", "服务标识");
            return { version: info.version, os: info.os };
        });

        await step("commit.first", async function () {
            doc = await newDocument(docPath, v1);
            var path = WG.localPath(doc, false);
            var status = await WG.call("status", { document: path });
            check(status.workingState === "noCommit", "初始状态应为无提交，实际 " + status.workingState);
            var result = await WG.call("commit", { document: path, message: "初稿", author: "端到端" });
            check(result.commit, "应创建提交");
            ids.v1 = result.commit;
            status = await WG.call("status", { document: path });
            check(status.workingState === "clean" && status.head === ids.v1, "提交后应一致");
            return result.message;
        });

        await step("commit.second", async function () {
            await setText(doc, v2);
            var status = await WG.call("status", { document: docPath });
            check(status.workingState === "changed", "修改后应为有未提交修改");
            var result = await WG.call("commit", { document: docPath, message: "补充第一段", author: "端到端", expectedHead: ids.v1 });
            ids.v2 = result.commit;
            var unchanged = await WG.call("commit", { document: docPath, message: "重复", author: "端到端" });
            check(unchanged.commit === null, "内容未变不应提交");
            return { v2: ids.v2 };
        });

        await step("version.preview", async function () {
            var result = await WG.call("version/preview", { document: docPath, revision: ids.v1 });
            var preview = await WG.open(result.file, true);
            var content = text(preview);
            check(preview.ReadOnly, "预览应只读");
            check(content.indexOf("已补充") < 0 && content.indexOf("初稿内容") >= 0, "预览内容应为第一版");
            await WG.close(preview);
            return result.file;
        });

        await step("version.compare", async function () {
            var files = await WG.call("version/compare", { document: docPath, revision: ids.v1 });
            var compared = await WG.compare(files.original, files.revised, "当前稿", false);
            var count = compared.Revisions.Count;
            var items = WG.reviewItems(compared);
            await WG.close(compared);
            check(count > 0, "对比稿应有修订");
            return { revisions: count, first: items[0] && items[0].kind + ":" + items[0].text };
        });

        await step("branch.create-and-switch", async function () {
            await WG.call("branch/create", { document: docPath, name: "review", revision: ids.v1 });
            var prepared = await WG.call("branch/switch", { document: docPath, name: "review", author: "端到端" });
            var message = await WG.applyReplacement(prepared);
            doc = WG.findDocument(docPath);
            check(doc, "切换后应重新打开文档");
            check(text(doc).indexOf("已补充") < 0, "切换后应为 review 分支内容");
            var status = await WG.call("status", { document: docPath });
            check(status.branch === "review" && status.workingState === "clean", "应位于 review 分支");
            return message;
        });

        await step("branch.commit-and-switch-back", async function () {
            await setText(doc, "WordGit 端到端测试\r第一段：初稿内容。\r第二段：送审前的内容，review 分支改写。");
            ids.review = (await WG.call("commit", { document: docPath, message: "review 分支修改", author: "审阅人" })).commit;
            var prepared = await WG.call("branch/switch", { document: docPath, name: "main" });
            await WG.applyReplacement(prepared);
            doc = WG.findDocument(docPath);
            check(text(doc).indexOf("已补充") >= 0 && text(doc).indexOf("review 分支改写") < 0, "应回到 main 内容");
            return { review: ids.review };
        });

        await step("merge.branch", async function () {
            var plan = await WG.call("merge/branch", { document: docPath, source: "review" });
            check(!plan.upToDate && plan.basisIsTarget === false, "main 与 review 双方都有修改");
            var notice = await WG.buildMergeDraft(plan, "main");
            var draft = WG.findDocument(plan.draft);
            check(draft, "合并稿应已打开");
            var revisions = draft.Revisions.Count;
            check(revisions > 0, "合并稿应有修订");
            var status = await WG.call("status", { document: plan.draft });
            check(status.review && status.review.source_branch === "review", "合并稿应关联原稿");
            // The reference document listing main's own changes is open too; close it unsaved.
            var docs = [];
            for (var i = 1; i <= app.Documents.Count; i++) docs.push(app.Documents.Item(i));
            for (var j = 0; j < docs.length; j++) {
                var name = String(docs[j].FullName);
                if (name.indexOf("/") < 0) await WG.close(docs[j]);
            }
            return { revisions: revisions, notice: notice.substring(0, 80), draft: plan.draft };
        });

        await step("merge.finish", async function () {
            var draft = null;
            for (var i = 1; i <= app.Documents.Count; i++) {
                if (/分支合并稿\.docx$/.test(app.Documents.Item(i).FullName)) draft = app.Documents.Item(i);
            }
            check(draft, "找不到合并稿");
            draft.Activate();
            // Keep main's own change: reject revisions that undo "已补充", accept the rest.
            var items = WG.reviewItems(draft);
            for (var k = items.length - 1; k >= 0; k--) {
                if (items[k].comment) continue;
                var undoesMain = items[k].type === 2 && items[k].rangeText.indexOf("已补充") >= 0 ||
                    items[k].type === 2 && items[k].rangeText.indexOf("，已补充") >= 0;
                WG.resolveItem(draft, items[k], !undoesMain);
            }
            draft.AcceptAllRevisions();
            draft.Save();
            await WG.settle(50);
            var draftPath = WG.localPath(draft, false);
            var prepared = await WG.call("merge/finish", { draft: draftPath, message: "合并 review 分支", author: "端到端" });
            var message = await WG.applyReplacement(prepared);
            doc = WG.findDocument(docPath);
            var log = await WG.call("history", { document: docPath, branch: "main" });
            check(log.history[0].parents.length === 2, "合并提交应有两个父提交");
            check(text(doc).indexOf("review 分支改写") >= 0, "原稿应包含 review 的修改");
            ids.merge = log.history[0].id;
            await WG.close(draft);
            return { message: message, text: text(doc).substring(0, 80) };
        });

        await step("branch.delete-merged", async function () {
            var deletable = await WG.call("branch/deletable", { document: docPath });
            check(deletable.branches.indexOf("review") >= 0, "review 已合入，应可删除");
            return (await WG.call("branch/delete", { document: docPath, name: "review" })).message;
        });

        await step("version.restore", async function () {
            var prepared = await WG.call("version/restore", { document: docPath, revision: ids.v1, author: "端到端" });
            var message = await WG.applyReplacement(prepared);
            doc = WG.findDocument(docPath);
            check(text(doc).indexOf("已补充") < 0 && text(doc).indexOf("review 分支改写") < 0, "应恢复到第一版");
            var status = await WG.call("status", { document: docPath });
            check(status.workingState === "clean" && status.head !== ids.merge, "恢复应创建新提交");
            ids.restored = status.head;
            return message;
        });

        await step("version.export", async function () {
            var result = await WG.call("version/export", { document: docPath, revision: ids.v2 });
            check(global.wps.FileSystem.Exists(result.file), "另存文件应存在");
            return result.file;
        });

        await step("review.import-branch", async function () {
            var returned = HOME + "/老师返稿.docx";
            var copy = await newDocument(returned, v1 + "\r老师补充：请说明实验条件。");
            await WG.close(copy);
            var result = await WG.call("review/import", { document: docPath, file: returned, basis: ids.v1, branch: "review-老师", reviewer: "老师" });
            check(result.commit, "返稿应保存为分支");
            var log = await WG.call("history", { document: docPath, branch: "review-老师" });
            check(log.history[0].author === "老师", "返稿提交作者应为来源");
            return result.message;
        });

        await step("review.apply-to-original", async function () {
            var returned = HOME + "/老师返稿.docx";
            var plan = await WG.call("review/import", { document: docPath, file: returned, basis: ids.v1, reviewer: "老师" });
            check(plan.basisIsTarget === true, "当前稿等于送审版，应为精确对比");
            await WG.buildMergeDraft(plan, "main");
            var draft = WG.findDocument(plan.draft);
            var items = WG.reviewItems(draft);
            check(items.some(function (i) { return !i.comment && i.author === "老师"; }), "修订作者应为老师");
            draft.Activate();
            await acceptAll(draft);
            var prepared = await WG.call("merge/finish", { draft: plan.draft, message: "合并老师返稿", author: "端到端" });
            await WG.applyReplacement(prepared);
            doc = WG.findDocument(docPath);
            check(text(doc).indexOf("老师补充") >= 0, "原稿应包含返稿内容");
            await WG.close(draft);
            return { revisions: items.length };
        });

        var shared = HOME + "/shared.git";
        var clonePath = HOME + "/协作者/论文.docx";
        await step("remote.link-push", async function () {
            var linked = await WG.call("remote/link", { document: docPath, url: shared, branch: "main", operation: "e2e-link" });
            check(linked.remote.linked, "应已关联");
            var pushed = await WG.call("remote/push", { document: docPath, operation: "e2e-push" });
            check(pushed.remote.ahead === 0 && pushed.remote.behind === 0, "推送后应同步");
            return pushed.remote.caption;
        });

        await step("remote.clone", async function () {
            var result = await WG.call("remote/clone", { url: shared, branch: "main", document: clonePath, operation: "e2e-clone" });
            var cloned = await WG.open(result.document, false);
            check(text(cloned).indexOf("老师补充") >= 0, "克隆稿内容应与原稿一致");
            var status = await WG.call("status", { document: clonePath });
            check(status.workingState === "clean" && status.remote.linked, "克隆稿应干净且已关联");
            await setText(cloned, text(cloned).replace(/\r$/, "") + "\r协作者新增段落。");
            await WG.call("commit", { document: clonePath, message: "协作者修改", author: "协作者" });
            var pushed = await WG.call("remote/push", { document: clonePath, operation: "e2e-push2" });
            await WG.close(cloned);
            return pushed.remote.caption;
        });

        await step("remote.pull-fast-forward", async function () {
            doc = WG.findDocument(docPath);
            doc.Activate();
            var result = await WG.call("remote/pull", { document: docPath, operation: "e2e-pull" });
            check(result.action === "replace", "应可快进拉取，实际 " + result.action);
            var message = await WG.applyReplacement(result);
            doc = WG.findDocument(docPath);
            check(text(doc).indexOf("协作者新增段落") >= 0, "拉取后原稿应包含协作者修改");
            var status = await WG.call("status", { document: docPath });
            check(status.remote.behind === 0 && status.workingState === "clean", "拉取后应同步且干净");
            return message;
        });

        await step("remote.pull-diverged", async function () {
            var cloned = await WG.open(clonePath, false);
            await setText(cloned, text(cloned).replace(/\r$/, "") + "\r协作者第二次修改。");
            await WG.call("commit", { document: clonePath, message: "协作者第二次修改", author: "协作者" });
            await WG.call("remote/push", { document: clonePath, operation: "e2e-push3" });
            await WG.close(cloned);
            doc = WG.findDocument(docPath);
            await setText(doc, text(doc).replace(/\r$/, "") + "\r作者本地修改。");
            await WG.call("commit", { document: docPath, message: "作者本地修改", author: "作者" });
            var result = await WG.call("remote/pull", { document: docPath, operation: "e2e-pull2" });
            check(result.action === "merge", "双方都有修改应生成合并稿，实际 " + result.action);
            var plan = await WG.call("merge/branch", { document: docPath, source: result.source });
            await WG.buildMergeDraft(plan, "main");
            var draft = WG.findDocument(plan.draft);
            check(draft && draft.Revisions.Count > 0, "拉取合并稿应有修订");
            return { revisions: draft.Revisions.Count };
        });

        await step("ui.pane", async function () {
            var pane = WG.showPane();
            // The pane renders asynchronously; wait until it has loaded the document's history.
            var rendered = false;
            for (var i = 0; i < 40 && !rendered; i++) {
                await WG.settle(250);
                rendered = !!WG.app().PluginStorage.getItem("wordgit.pane.rendered");
            }
            check(pane.Visible, "侧栏应可见");
            check(rendered, "侧栏应在 10 秒内完成渲染：" + (WG.app().PluginStorage.getItem("wordgit.pane.error") || "无脚本错误"));
            return { id: pane.ID, rendered: WG.app().PluginStorage.getItem("wordgit.pane.rendered") };
        });
    }

    async function main() {
        if (!global.wps || !global.wps.FileSystem || !global.wps.FileSystem.Exists(MARKER)) return;
        global.wps.FileSystem.Remove(MARKER);
        lines = [];
        record("start", true, { href: location.href });
        try { await run(); } catch (e) { /* recorded by step */ }
        record("done", failed === 0, { failed: failed });
    }

    // Wait for WPS to finish starting (and for start-up prompts to be confirmed).
    setTimeout(main, 6000);
})(window);
