// WordGit for WPS Writer — user flows behind the ribbon and the task pane.
(function (global) {
    "use strict";
    var WG = global.WordGit;

    // ---- busy state and progress ----------------------------------------------------------------

    var busy = null;

    WG.run = async function (title, body) {
        if (busy) {
            await WG.notify("正在" + busy + "，请等待当前操作完成。");
            return;
        }
        busy = title;
        try {
            var message = await body();
            WG.touch();
            if (message) await WG.notify(message, "WordGit · " + title);
        } catch (error) {
            WG.touch();
            await WG.notify(WG.errorText(error), "WordGit · " + title);
        } finally {
            busy = null;
        }
    };

    // Long network operations show a cancellable progress dialog.
    WG.withProgress = async function (title, status, operation) {
        var app = WG.app();
        var id = WG.newOperationId();
        var key = "wordgit.progress." + id;
        app.PluginStorage.setItem(key, JSON.stringify({ title: title, status: status, operation: id, done: false }));
        var scale = global.devicePixelRatio || 1;
        app.ShowDialog(WG.baseUrl() + "ui/progress.html?key=" + key, "WordGit · " + title, 420 * scale, 190 * scale, false);
        try {
            return await operation(id);
        } finally {
            app.PluginStorage.setItem(key, JSON.stringify({ done: true }));
        }
    };

    // ---- context of the active document --------------------------------------------------------

    async function currentStatus(allowSave) {
        var doc = WG.activeDocument();
        var path = WG.localPath(doc, allowSave);
        var status = await WG.call("status", { document: path });
        if (status.session) {
            throw new Error("这是 WordGit 生成的临时预览或比较稿，没有自己的版本历史。如需保留，请先「另存为」到其他位置。");
        }
        return { doc: doc, path: path, status: status };
    }
    WG.currentStatus = currentStatus;

    function requireHead(status) {
        if (!status.head) throw new Error("还没有版本，请先点击「提交版本」。");
    }

    function requireOriginal(status, action) {
        if (status.review) throw new Error("请在原文档中" + action + "。");
    }

    function revisionChoices(status) {
        return (status.history || []).map(function (r) {
            return { id: r.id, label: WG.shortId(r.id) + " · " + WG.formatDate(r.date) + " · " + r.subject, author: r.author };
        });
    }

    // Opens the revision picker; resolves to { revision, other } or null.
    function pickRevision(status, title, hint, options) {
        return WG.dialog("pick", "WordGit · " + title, 760, 520, {
            hint: hint, history: status.history, graph: status.graph, head: status.head,
            branch: status.branch, multi: !!(options && options.multi), requireTwo: !!(options && options.requireTwo),
            action: (options && options.action) || "确定", preselectHead: !!(options && options.preselectHead)
        });
    }

    // ---- commit ---------------------------------------------------------------------------------

    WG.commit = function () {
        return WG.run("提交版本", async function () {
            var ctx = await currentStatus(true);
            var review = ctx.status.review;
            if (review && (review.source_branch || review.apply_to_original)) return WG.finishMerge();
            var target = ctx.status.historyPath.replace(/^.*[\\/]/, "");
            var finishing = review && review.record_only_merge;
            if (finishing && ctx.doc.Revisions.Count > 0) throw new Error("合并稿仍有待处理的修订，请先逐条接受或拒绝。");
            var answer = await WG.ask((finishing ? "完成合并到「" : "提交到「") + target + "」 · " + ctx.status.branch, {
                label: "提交标题", detailLabel: "详细内容（可选）", value: finishing ? "合并审阅稿" : "", multiline: true, ok: finishing ? "完成合并" : "提交"
            });
            if (!answer) return null;
            WG.saveDocument(ctx.doc);
            var message = answer.detail ? answer.value + "\n\n" + answer.detail : answer.value;
            var result = await WG.call("commit", { document: ctx.path, message: message, author: WG.userName(), expectedHead: ctx.status.head });
            return result.message;
        });
    };

    // ---- versions ---------------------------------------------------------------------------------

    WG.preview = function (revision) {
        return WG.run("预览版本", async function () {
            var ctx = await currentStatus(false);
            requireHead(ctx.status);
            if (!revision) {
                var picked = await pickRevision(ctx.status, "预览版本", "选择要只读打开的版本。", { action: "打开预览" });
                if (!picked) return null;
                revision = picked.revision;
            }
            var result = await WG.call("version/preview", { document: ctx.path, revision: revision });
            await WG.open(result.file, true);
            return null;
        });
    };

    WG.exportVersion = function () {
        return WG.run("另存旧版", async function () {
            var ctx = await currentStatus(false);
            requireHead(ctx.status);
            var picked = await pickRevision(ctx.status, "另存旧版", "选择要另存为新文档的版本。新文档保存在原稿所在文件夹。", { action: "另存" });
            if (!picked) return null;
            var result = await WG.call("version/export", { document: ctx.path, revision: picked.revision });
            await WG.open(result.file, false);
            return result.message;
        });
    };

    async function compareVersion(ctx, revision, other) {
        var useSnapshot = !other && !ctx.doc.Saved;
        var files = await WG.call("version/compare", { document: ctx.path, revision: revision,
            other: other || null, currentInMemory: useSnapshot });
        try {
            if (useSnapshot) WG.writeComparisonSnapshot(ctx.doc, files.revised);
            var result = await WG.compare(files.original, files.revised, other ? "所选版本" : "当前稿", false);
            result.Activate();
            return null;
        } finally {
            if (useSnapshot) WG.removeComparisonSnapshot(files.revised);
        }
    }

    WG.compareVersions = function (revision, other) {
        return WG.run("对比版本", async function () {
            var ctx = await currentStatus(false);
            requireHead(ctx.status);
            if (!revision) {
                var picked = await pickRevision(ctx.status, "对比版本", "选一版与当前稿对比，或按 Ctrl 选两个历史版本对比。", { multi: true, action: "对比" });
                if (!picked) return null;
                revision = picked.revision;
                other = picked.other;
            }
            return compareVersion(ctx, revision, other);
        });
    };

    WG.compareLatest = function () {
        return WG.run("对比最新提交", async function () {
            var ctx = await currentStatus(false);
            if (!ctx.status.head) return "尚无提交版本可对比。";
            return compareVersion(ctx, ctx.status.head, null);
        });
    };

    // ---- working-copy replacement (switch, restore, pull, finish merge) --------------------------

    // The service prepared a replacement; close the document, let the service swap it and
    // publish the branch change, then reopen. The document must be saved (no unsaved edits).
    async function applyReplacement(prepared) {
        if (prepared.unchanged) return prepared.message;
        var open = WG.findDocument(prepared.document);
        if (open && (!open.Saved || open.ReadOnly)) {
            await WG.call("plan/cancel", { plan: prepared.plan });
            throw new Error("当前稿有未保存的修改，或处于只读状态。请先处理后重试。");
        }
        if (open) await WG.close(open);
        var result;
        try {
            result = await WG.call("plan/apply", { plan: prepared.plan });
        } finally {
            if (open) await WG.open(prepared.document, false);
        }
        return result.message;
    }
    WG.applyReplacement = applyReplacement;

    WG.switchBranch = function () {
        return WG.run("切换分支", async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, "切换分支");
            requireHead(ctx.status);
            var locals = ctx.status.branches.filter(function (b) { return b.indexOf("refs/remotes/") !== 0 && b !== ctx.status.branch; });
            if (!locals.length) return "当前只有「" + ctx.status.branch + "」分支。请先新建分支。";
            var target = await WG.dialog("branch", "WordGit · 切换分支", 460, 380, {
                hint: "选择要切换到的本地分支。当前分支：" + ctx.status.branch, branches: locals, action: "切换"
            });
            if (!target) return null;
            var message = null;
            if (!ctx.doc.Saved || ctx.status.workingState === "changed") {
                if (!await WG.confirm("当前稿有未提交的修改。是否先提交到「" + ctx.status.branch + "」再切换？", "WordGit · 切换分支", "提交并切换")) return null;
                var answer = await WG.ask("切换前提交到「" + ctx.status.branch + "」", { label: "提交标题", value: "切换前保存当前稿", ok: "提交" });
                if (!answer) return null;
                message = answer.value;
                WG.saveDocument(ctx.doc);
            }
            var prepared = await WG.call("branch/switch", { document: ctx.path, name: target, message: message, author: WG.userName() });
            return applyReplacement(prepared);
        });
    };

    WG.restoreVersion = function (revision) {
        return WG.run("恢复旧版", async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, "恢复旧版");
            requireHead(ctx.status);
            if (ctx.doc.ReadOnly) throw new Error("当前文档只读，请使用「另存旧版」。");
            if (!revision) {
                var picked = await pickRevision(ctx.status, "恢复旧版", "选择要恢复为当前稿的版本。", { action: "恢复" });
                if (!picked) return null;
                revision = picked.revision;
            }
            if (!await WG.confirm("先存下当前稿，再将所选旧版恢复到原文件，并创建一次新提交。是否继续？", "WordGit · 恢复旧版", "恢复")) return null;
            WG.saveDocument(ctx.doc);
            var prepared = await WG.call("version/restore", { document: ctx.path, revision: revision, author: WG.userName() });
            return applyReplacement(prepared);
        });
    };

    // ---- branches -------------------------------------------------------------------------------

    WG.createBranch = function (revision) {
        return WG.run("新建分支", async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, "新建分支");
            requireHead(ctx.status);
            if (!revision) {
                var picked = await pickRevision(ctx.status, "新建分支", "选择新分支的起始版本。", { action: "下一步", preselectHead: true });
                if (!picked) return null;
                revision = picked.revision;
            }
            var answer = await WG.ask("从 " + WG.shortId(revision) + " 新建分支", { label: "分支名称", value: "", ok: "新建" });
            if (!answer) return null;
            var result = await WG.call("branch/create", { document: ctx.path, name: answer.value, revision: revision });
            return result.message;
        });
    };

    WG.deleteBranch = function () {
        return WG.run("删除分支", async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, "删除分支");
            requireHead(ctx.status);
            var deletable = await WG.call("branch/deletable", { document: ctx.path });
            if (!deletable.branches.length) return "没有可以删除的分支。这里仅列出已合入当前分支的其他分支。";
            var target = await WG.dialog("branch", "WordGit · 删除分支", 460, 380, {
                hint: "只能删除已经合入「" + deletable.current + "」的分支。", branches: deletable.branches, action: "删除"
            });
            if (!target) return null;
            if (!await WG.confirm("删除已合入「" + deletable.current + "」的分支「" + target + "」？", "WordGit · 删除分支", "删除")) return null;
            return (await WG.call("branch/delete", { document: ctx.path, name: target })).message;
        });
    };

    // Builds a merge draft in WPS from the files the service prepared, saves it in the session
    // folder, registers it and opens it for review.
    async function buildMergeDraft(plan, targetLabel) {
        var draft;
        if (plan.original) {
            draft = await WG.compare(plan.original, plan.revised, "修订版本", false);
        } else if (plan.submittedBase) {
            draft = await WG.buildReturnedReview(plan);
        } else {
            // Exact when the current branch did not change since the common version. Otherwise the
            // draft also shows the current branch's own changes as revisions that would undo them;
            // a reference document lists those so the author can reject them.
            draft = await WG.compare(plan.target, plan.incoming, plan.incomingAuthor, !!plan.preserveIncoming && plan.basisIsTarget);
        }
        WG.saveAs(draft, plan.draft);
        await WG.close(draft);
        await WG.settle(50);
        await WG.call("merge/register", { draft: plan.draft, review: {
            original: plan.review.original, head: plan.review.head, branch: plan.review.branch,
            sourceBranch: plan.review.sourceBranch || null, sourceTip: plan.review.sourceTip || null,
            applyToOriginal: !!plan.review.applyToOriginal, recordOnlyMerge: !!plan.review.recordOnlyMerge
        } });
        var notice = plan.notice;
        if (!plan.original && !plan.basisIsTarget && !plan.submittedBase) {
            await WG.compare(plan.basis, plan.target, targetLabel || "当前分支", false);
            notice += "\n\n双方都有修改：WPS 只能对比两份文稿，合并稿中也会出现「撤销当前分支修改」的修订。" +
                "已另外打开一份「当前分支的修改」对照稿；合并稿中与对照稿相反的修订请选择「拒绝」。";
        }
        var opened = await WG.open(plan.draft, false);
        opened.Activate();
        WG.showPane();
        return notice;
    }
    WG.buildMergeDraft = buildMergeDraft;

    WG.mergeBranch = function (source) {
        return WG.run("合并分支", async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, "合并分支");
            requireHead(ctx.status);
            if (!source) {
                var others = ctx.status.branches.filter(function (b) { return b !== ctx.status.branch; });
                if (!others.length) return "当前只有「" + ctx.status.branch + "」分支。请先新建分支。";
                source = await WG.dialog("branch", "WordGit · 合并分支", 460, 400, {
                    hint: "选择要合入「" + ctx.status.branch + "」的来源分支。", branches: others, action: "生成合并稿"
                });
                if (!source) return null;
            }
            if (!ctx.doc.Saved || ctx.status.workingState !== "clean") throw new Error("请先提交当前分支的修改，再生成合并稿。");
            var plan = await WG.call("merge/branch", { document: ctx.path, source: source });
            if (plan.upToDate) return plan.message;
            return buildMergeDraft(plan, ctx.status.branch);
        });
    };

    WG.importReview = function () {
        return WG.run("返稿合并", async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, "处理返稿");
            requireHead(ctx.status);
            var options = await WG.dialog("import", "WordGit · 返稿合并", 640, 560, {
                document: ctx.path, history: revisionChoices(ctx.status), branch: ctx.status.branch, branches: ctx.status.branches
            });
            if (!options) return null;
            if (options.historical) return WG.mergeTwoVersions(ctx);
            var expectedHead = ctx.status.head;
            if (options.createMergeDraft) {
                var current = await WG.call("status", { document: ctx.path });
                if (current.head !== expectedHead || current.branch !== ctx.status.branch) {
                    throw new Error("当前分支或提交已变化，请刷新后重新选择。");
                }
                if (!ctx.doc.Saved || current.workingState !== "clean") {
                    var answer = await WG.ask("提交当前稿并继续返稿合并", {
                        label: "提交标题", detailLabel: "详细内容（可选）", multiline: true, ok: "提交并继续"
                    });
                    if (!answer) return null;
                    if (!WG.samePath(WG.localPath(ctx.doc, false), ctx.path)) throw new Error("当前稿路径已变化，请重新选择返稿。");
                    WG.saveDocument(ctx.doc);
                    var message = answer.detail ? answer.value + "\n\n" + answer.detail : answer.value;
                    var committed = await WG.call("commit", { document: ctx.path, message: message,
                        author: WG.userName(), expectedHead: expectedHead, expectedBranch: ctx.status.branch });
                    expectedHead = committed.commit || expectedHead;
                }
            }
            var imported = await WG.call("review/import", { document: ctx.path, file: options.file, basis: options.basis,
                branch: options.branch, reviewer: options.reviewer, archiveOnly: true,
                expectedHead: expectedHead, expectedBranch: ctx.status.branch });
            if (!options.createMergeDraft) return imported.message + " 可随时从「合并分支」继续审阅。";
            try {
                if (!ctx.doc.Saved) throw new Error("当前稿有新的未保存修改，请提交后重试。");
                var branchPlan = await WG.call("merge/context", { document: ctx.path, source: imported.source, basis: options.basis });
                if (branchPlan.upToDate) return imported.message + " 这份返稿已经合入当前分支。";
                var notice = await buildMergeDraft(branchPlan, ctx.status.branch);
                return imported.message + "\n" + notice;
            } catch (error) {
                return imported.message + "\n生成合并稿未完成：" + WG.errorText(error) + "\n再次选择同一返稿和送审版本会复用这条记录。";
            }
        });
    };

    WG.mergeTwoVersions = async function (ctx) {
        var picked = await pickRevision(ctx.status, "合并两版", "按 Ctrl 选择两个历史版本，较早的一版作为基准。", { multi: true, requireTwo: true, action: "合并两版" });
        if (!picked) return null;
        var plan = await WG.call("merge/versions", { document: ctx.path, revision: picked.revision, other: picked.other });
        return buildMergeDraft(plan);
    };

    // Called from a merge draft once every revision is handled.
    WG.finishMerge = async function () {
        var doc = WG.activeDocument();
        var path = WG.localPath(doc, false);
        if (doc.Revisions.Count > 0) throw new Error("合并稿仍有待处理的修订，请先逐条接受或拒绝。");
        var status = await WG.call("status", { document: path });
        var review = status.review;
        if (!review) throw new Error("这不是待完成的合并稿。");
        var source = review.source_branch;
        var summary = source && !source.startsWith("refs/wordgit/reviews/") ? "合并分支 " + source.replace("refs/remotes/", "") : "合并返稿";
        var answer = await WG.ask("确认合并到「" + review.branch + "」", { label: "提交标题", value: summary, ok: "完成合并" });
        if (!answer) return null;
        WG.saveDocument(doc);
        var prepared = await WG.call("merge/finish", { draft: path, message: answer.value, author: WG.userName() });
        return applyReplacement(prepared);
    };

    WG.finishMergeAction = function () {
        return WG.run("完成合并", WG.finishMerge);
    };

    // ---- remote ---------------------------------------------------------------------------------

    WG.linkRemote = function () {
        return WG.run("关联仓库", async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, "关联远程仓库");
            var remote = ctx.status.remote || {};
            var options = await WG.dialog("remote", "WordGit · 关联远程仓库", 560, 380, {
                clone: false, url: remote.url || "", branch: remote.remote_branch || "main"
            });
            if (!options) return null;
            var result = await WG.withProgress("关联远程仓库", "正在关联并获取远程历史…", function (operation) {
                return WG.call("remote/link", { document: ctx.path, url: options.url, branch: options.branch, operation: operation });
            });
            return result.message;
        });
    };

    WG.cloneRemote = function () {
        return WG.run("克隆稿件", async function () {
            var options = await WG.dialog("remote", "WordGit · 克隆远程稿件", 560, 440, { clone: true, branch: "main" });
            if (!options) return null;
            var result = await WG.withProgress("克隆稿件", "正在下载远程历史…", function (operation) {
                return WG.call("remote/clone", { url: options.url, branch: options.branch, document: options.document, operation: operation });
            });
            await WG.open(result.document, false);
            return result.message;
        });
    };

    function sync(kind) {
        var titles = { fetch: "获取", push: "推送", pull: "拉取" };
        return WG.run(titles[kind], async function () {
            var ctx = await currentStatus(false);
            requireOriginal(ctx.status, titles[kind]);
            if (!ctx.status.remote || !ctx.status.remote.linked) throw new Error("请先点击「关联仓库」。");
            if (kind !== "fetch") requireHead(ctx.status);
            if (kind === "pull" && !ctx.doc.Saved) throw new Error("当前稿有未保存的修改，请先保存并提交，再拉取。");
            var result = await WG.withProgress(titles[kind], kind === "push" ? "正在推送当前分支…" : "正在获取远程历史…", function (operation) {
                return WG.call("remote/" + kind, { document: ctx.path, operation: operation });
            });
            if (kind !== "pull") return result.message;
            if (result.action === "merge") {
                var plan = await WG.call("merge/branch", { document: ctx.path, source: result.source });
                var notice = await buildMergeDraft(plan, ctx.status.branch);
                return "双方修改已生成修订合并稿。" + notice + "\n确认修订后点击「完成合并」，再回到原稿推送。";
            }
            if (result.action === "replace") return applyReplacement(result);
            return result.message;
        });
    }

    WG.fetchRemote = function () { return sync("fetch"); };
    WG.pushRemote = function () { return sync("push"); };
    WG.pullRemote = function () { return sync("pull"); };

    // ---- pane -------------------------------------------------------------------------------------

    WG.showPane = function () {
        var app = WG.app();
        var id = app.PluginStorage.getItem("wordgit.pane");
        var pane = null;
        try { pane = id ? app.GetTaskPane(id) : null; } catch (e) { pane = null; }
        if (!pane) {
            pane = app.CreateTaskPane(WG.baseUrl() + "ui/pane.html", "WordGit · 批注、修订与提交");
            app.PluginStorage.setItem("wordgit.pane", pane.ID);
            try { pane.DockPosition = 2; } catch (e) { /* right */ }
            try { pane.Width = 420; } catch (e) { /* keep the host default */ }
        }
        if (!pane.Visible) pane.Visible = true;
        return pane;
    };

    WG.about = function () {
        return WG.run("关于", async function () {
            var info = null, error = null;
            try { info = await WG.call("info", {}); } catch (e) { error = WG.errorText(e); }
            await WG.dialog("about", "WordGit", 460, 340, { info: info, error: error });
            return null;
        });
    };

    // Messages queued by the service at startup (e.g. an interrupted replacement repaired).
    WG.showNotices = async function () {
        try {
            var result = await WG.call("notices", {});
            if (result.notices && result.notices.length) await WG.notify(result.notices.join("\n\n"));
        } catch (e) { /* the service may not be running yet */ }
    };
})(window);
