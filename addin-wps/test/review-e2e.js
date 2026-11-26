// Loaded only by the WPS regression-test harness, with an isolated service and document folder.
(function (global) {
    "use strict";
    var WG = global.WordGit, home = global.WORDGIT_REVIEW_TEST_HOME;
    if (!home) return;

    var records = [], report = home + "/review-report.jsonl";
    var base = "论文初稿\r材料配比与研究背景。\r统计结果与结论。";
    var mainPath = home + "/主稿.docx";
    var app, doc, nextOptions, lastNotice, prompts = [], imports = [];

    function record(step, ok, detail) {
        records.push(JSON.stringify({ step: step, ok: ok, detail: detail || "" }));
        global.wps.FileSystem.WriteFile(report, records.join("\n") + "\n");
    }

    function check(value, message) {
        if (!value) throw new Error(message);
    }

    async function documentAt(path, content) {
        var item = app.Documents.Add();
        await WG.settle(60);
        item.Content.Text = content;
        WG.saveAs(item, path);
        await WG.settle(60);
        return item;
    }

    async function status() { return WG.call("status", { document: mainPath }); }

    async function test() {
        app = WG.app();
        var dialog = WG.dialog, ask = WG.ask, notify = WG.notify, call = WG.call, showPane = WG.showPane;
        WG.dialog = async function (kind) {
            check(kind === "import", "Unexpected dialog: " + kind);
            return nextOptions;
        };
        var cancelCommit = false;
        WG.ask = async function (title) {
            prompts.push(title);
            return cancelCommit ? null : { value: "验收提交 " + prompts.length, detail: "" };
        };
        WG.notify = async function (message) { lastNotice = String(message); };
        WG.call = async function (endpoint, args) {
            var response = await call(endpoint, args);
            if (endpoint === "review/import") imports.push(response);
            return response;
        };
        WG.showPane = function () { };
        try {
            doc = await documentAt(mainPath, base);
            var sent = (await WG.call("commit", { document: mainPath, message: "送审版本", author: "验收" })).commit;
            var compareEdit = "WPS 未保存对比标记";
            doc.TrackRevisions = true;
            doc.Range(0, 0).InsertBefore(compareEdit + "\r");
            var pending = doc.Revisions.Count;
            check(!doc.Saved && pending > 0, "测试稿必须有未保存修订");
            lastNotice = "";
            await WG.compareLatest();
            var compared = app.ActiveDocument;
            check(!lastNotice, lastNotice);
            check(String(compared.Content.Text).includes(compareEdit) && compared.Revisions.Count > 0, "对比稿漏掉未保存的改动");
            check(!doc.Saved && doc.Revisions.Count === pending && doc.TrackRevisions, "对比改动了原稿或其修订状态");
            check((await status()).workingState === "clean", "对比时自动保存了原稿");
            await WG.close(compared);
            await WG.close(doc);
            doc = await WG.open(mainPath, false);
            record("unsaved comparison", true);

            var cumulative = base;
            for (var round = 1; round <= 2; round++) {
                var returnPath = home + "/老师返稿-" + round + ".docx";
                var returned = await documentAt(returnPath, cumulative);
                returned.TrackRevisions = true;
                var end = returned.Content.End - 1, addition = "导师意见第" + round + "轮";
                returned.Range(end, end).InsertAfter("\r" + addition);
                returned.Save();
                await WG.close(returned);
                doc.Activate();
                var author = "LOCAL-" + round + " 作者补充实验说明。";
                doc.Range(0, 0).InsertBefore(author + "\r");
                check(!doc.Saved, "生成合并稿前应存在未保存改动");
                var before = await status();
                nextOptions = { file: returnPath, basis: sent, branch: round === 1 ? null : "review-round-2",
                    reviewer: "导师", createMergeDraft: true };
                if (round === 1) {
                    cancelCommit = true;
                    await WG.importReview();
                    check(!doc.Saved && imports.length === 0 && (await status()).head === before.head,
                        "取消提交后仍修改了稿件或导入了返稿");
                    cancelCommit = false;
                }
                lastNotice = "";
                await WG.importReview();
                var archived = imports[imports.length - 1], draft = app.ActiveDocument;
                check(archived && !archived.reused && !WG.samePath(draft.FullName, mainPath), lastNotice || "未打开合并稿");
                var working = await status();
                var review = await WG.call("status", { document: draft.FullName });
                check(working.head !== before.head && doc.Saved && prompts[prompts.length - 1].includes("提交当前稿并继续"),
                    "没有先提交带未保存修改的主稿");
                check(review.review && review.review.source_tip === archived.commit, "合并稿缺少原始返稿提交");
                var raw = await WG.call("version/preview", { document: mainPath, revision: archived.commit });
                var rawDoc = await WG.open(raw.file, true);
                check(rawDoc.Revisions.Count > 0, "原始返稿没有保留修订");
                await WG.close(rawDoc);
                draft.Activate();
                for (var index = draft.Revisions.Count; index >= 1; index--) {
                    var revision = draft.Revisions.Item(index);
                    if (revision.Type === 2 && String(revision.Range.Text).includes("LOCAL-")) revision.Reject();
                    else revision.Accept();
                }
                await WG.finishMerge();
                doc = WG.findDocument(mainPath);
                var final = await status();
                check(final.history[0].parents.join(" ") === working.head + " " + archived.commit, "合并提交缺少两个父节点");
                check(String(doc.Content.Text).includes(addition) && String(doc.Content.Text).includes(author),
                    "合并遗漏导师或作者的内容");

                doc.Activate();
                if (round === 2) nextOptions.branch = "review-repeat-" + round;
                await WG.importReview();
                var duplicate = imports[imports.length - 1], repeated = await status();
                check(duplicate.reused && duplicate.commit === archived.commit, "未复用同一送审版本的返稿");
                check(repeated.head === final.head && repeated.history.length === final.history.length &&
                    !repeated.branches.includes(nextOptions.branch), "重复合并新增了提交或分支");
                cumulative += "\r" + addition;
                record("round " + round, true);
            }
            record("done", true);
        } finally {
            WG.dialog = dialog;
            WG.ask = ask;
            WG.notify = notify;
            WG.call = call;
            WG.showPane = showPane;
        }
    }

    global.WordGitRunReviewTest = function () {
        setTimeout(function () {
            test().catch(function (error) { record("failed", false, WG.errorText(error) + "\n" + (error.stack || "")); });
        }, 0);
        return true;
    };
})(window);
