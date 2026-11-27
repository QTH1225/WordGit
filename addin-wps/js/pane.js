// The WPS pane follows the Word pane's review and version controls.
window.onerror = function (message, source, line) {
    try { window.WordGit.app().PluginStorage.setItem("wordgit.pane.error", message + " @" + line); } catch (error) { /* host is closing */ }
};

(function () {
    "use strict";
    var WG = window.WordGit, app = WG.app();
    var $ = function (id) { return document.getElementById(id); };
    var currentKey = null, status = null, items = [], selectedItem = null, graphView = null, visibleHistory = [];
    var documentViews = new Map();
    var loading = false, reloadQueued = false, graphGeneration = 0, graphLoading = false;
    var resolving = false, reviewReady = false, remaining = 0, reviewSignature = "", reviewTicks = 0;
    var observedSave = null, lastChange = "", graphError = "";

    $("version").textContent = window.WORDGIT_VERSION || "";

    function currentDocument() {
        try {
            var doc = WG.activeDocument();
            var path = "", error = "";
            try { path = WG.localPath(doc, false); } catch (e) { error = WG.errorText(e); }
            return { doc: doc, key: path || String(doc.Name), path: path, name: String(doc.Name),
                saved: !!doc.Saved, readOnly: !!doc.ReadOnly, error: error };
        } catch (e) {
            return { doc: null, key: "", saved: true, readOnly: true };
        }
    }

    function pendingMerge() {
        var review = status && status.review;
        return !!(review && (review.source_branch || review.apply_to_original || review.record_only_merge));
    }

    function updateButtons(view, ids) {
        var ready = status && !status.session && view.key === currentKey && !loading;
        var merging = pendingMerge(), current = selectedItem;
        var remote = ready && status.remote;
        var original = ready && !status.review;
        var reviewAction = view.doc && view.key === currentKey && !loading && !view.readOnly &&
            reviewReady && !resolving && current && !current.comment;
        ids = ids || (graphView ? graphView.selected() : []);
        $("accept").disabled = $("reject").disabled = !reviewAction;
        $("filter").disabled = !reviewReady;
        $("commit").textContent = merging ? "完成合并" : "提交版本";
        $("commit").title = merging ? "处理完剩余修订后，确认并提交合并结果。" : "保存当前文档，并提交到当前分支。";
        $("commit").disabled = !ready || view.readOnly || (merging && (!reviewReady || resolving || remaining > 0));
        $("compareLatest").disabled = !ready || !status.head;
        $("importReview").disabled = !original || view.readOnly || !status.head;
        $("link").disabled = !original;
        $("fetch").disabled = !original || !remote || !remote.linked;
        $("pull").disabled = $("push").disabled = $("fetch").disabled || !status.head;
        $("branchFilter").disabled = !ready || !status.branches.length;
        $("preview").disabled = !ready || graphLoading || ids.length !== 1;
        $("compare").disabled = !ready || graphLoading || !ids.length;
        $("compare").textContent = ids.length === 2 ? "对比两版" : "对比当前稿";
        $("historyStatus").textContent = graphError || (loading || graphLoading ? "正在读取提交记录…" :
            ids.length === 2 ? "已选中两版，可以对比。" : ids.length === 1 ? "双击预览，" + (WG.modKey || "Ctrl") + " 可选两版对比。" : "");
    }

    function itemKey(item) {
        return [item.comment, item.type, item.author, item.start, item.end, item.text].join("|");
    }

    function visibleItems() {
        var filter = $("filter").value;
        return items.filter(function (item) { return filter === "all" || (filter === "comment") === item.comment; });
    }

    function renderItems(view) {
        var list = $("items"), shown = visibleItems();
        var position = list.scrollTop;
        var revisions = items.filter(function (entry) { return !entry.comment; }).length;
        var comments = items.length - revisions;
        $("reviewTitle").textContent = !reviewReady && view.doc ? "批注与修订 · 稍后重试" :
            pendingMerge() ? "剩余修订 " + remaining + " 条 · 批注 " + comments : "批注 " + comments + " · 修订 " + revisions;
        list.setAttribute("aria-busy", view.doc && !reviewReady ? "true" : "false");
        if (!shown.length) {
            list.innerHTML = '<div class="empty">' + (!view.doc ? "请先打开文档" :
                !reviewReady ? "正在读取批注与修订…" : "没有" + $("filter").selectedOptions[0].text.replace("全部", "批注或修订")) + "</div>";
        } else {
            list.innerHTML = shown.map(function (item) {
                var index = items.indexOf(item), selected = item === selectedItem;
                var detail = [item.author || "未知", item.page > 0 ? "第 " + item.page + " 页" : ""].filter(Boolean).join(" · ");
                var caption = item.comment ? detail : item.kind + " · " + detail;
                return '<div class="review-item' + (selected ? " selected" : "") + '" role="option" aria-selected="' + selected +
                    '" tabindex="0" data-index="' + index + '" title="' + WG.escape(caption + "\n" + item.text) + '">' +
                    '<div class="caption">' + WG.escape(caption) + '</div>' +
                    '<div class="body">' + WG.escape(item.text.replace(/[\r\n]+/g, " ")) + "</div></div>";
            }).join("");
            list.scrollTop = position;
        }
        updateButtons(view);
    }

    function loadItems(view) {
        try {
            var old = selectedItem ? itemKey(selectedItem) : null;
            items = view.doc ? WG.reviewItems(view.doc, !pendingMerge()) : [];
            selectedItem = old == null ? null : items.find(function (item) { return itemKey(item) === old; }) || null;
            var count = items.filter(function (item) { return !item.comment; }).length;
            remaining = pendingMerge() ? Math.max(count, Number(view.doc.Revisions.Count)) : count;
            reviewSignature = view.doc ? WG.reviewSignature(view.doc) : "";
            reviewReady = !!view.doc;
        } catch (error) { reviewReady = false; }
        renderItems(view);
    }

    function selectItem(event) {
        var node = event.target.closest(".review-item"), view = currentDocument();
        if (!node || !reviewReady || view.key !== currentKey) return;
        selectedItem = items[+node.dataset.index];
        renderItems(view);
        try { WG.goToItem(view.doc, selectedItem); } catch (error) { $("reviewStatus").textContent = WG.errorText(error); }
    }

    $("items").onclick = selectItem;
    $("items").onkeydown = function (event) {
        if (event.key === "Enter" || event.key === " ") { event.preventDefault(); selectItem(event); }
    };
    $("filter").onchange = function () {
        if (visibleItems().indexOf(selectedItem) < 0) selectedItem = null;
        renderItems(currentDocument());
    };

    function resolve(accept) {
        var view = currentDocument(), current = selectedItem;
        if (!reviewReady || resolving || !current || current.comment || view.readOnly || view.key !== currentKey) return;
        var before = items.filter(function (item) { return !item.comment; });
        var ordinal = before.indexOf(current);
        resolving = true;
        updateButtons(view);
        try {
            WG.resolveItem(view.doc, current, accept);
            loadItems(view);
            if (!reviewReady) {
                $("reviewStatus").textContent = "已处理，请刷新列表。";
                return;
            }
            var revisions = items.filter(function (item) { return !item.comment; });
            selectedItem = revisions[ordinal] || revisions[0] || null;
            if (selectedItem) {
                WG.goToItem(view.doc, selectedItem);
                $("reviewStatus").textContent = accept ? "已接受，定位下一条。" : "已拒绝，定位下一条。";
            } else {
                $("reviewStatus").textContent = remaining ? "请在 WPS 的「审阅」中处理剩余修订。" : "修订已全部处理。";
            }
        } catch (error) {
            $("reviewStatus").textContent = WG.errorText(error);
            loadItems(view);
        } finally {
            resolving = false;
            renderItems(view);
            var row = $("items").querySelector(".selected");
            if (row) row.scrollIntoView({ block: "nearest" });
        }
    }

    $("accept").onclick = function () { resolve(true); };
    $("reject").onclick = function () { resolve(false); };

    function renderGraph(history, graph, keep, empty) {
        visibleHistory = history;
        var scroll = keep.length ? $("graph").scrollTop : 0;
        graphView = WG.renderGraph($("graph"), history, graph, status && status.head, {
            pane: true, multi: true, selection: keep, selectFirst: true, empty: empty,
            onSelect: function (ids) { updateButtons(currentDocument(), ids); },
            onOpen: function (id) {
                if (!$("preview").disabled && graphView.selected()[0] === id) run(function () { return WG.preview(id); });
            }
        });
        $("graph").scrollTop = scroll;
    }

    function showStatus(view) {
        var review = status && status.review, remote = status && status.remote;
        var branch = pendingMerge() ? "合入分支  " + review.branch :
            status && !status.session ? "当前分支  " + status.branch : "当前分支 —";
        $("branch").textContent = $("branch").title = branch;
        var location = status && status.historyPath || view.path || view.name || "";
        $("file").textContent = location.replace(/^.*[\\/]/, "");
        $("file").title = location;
        var working = $("state");
        if (!view.doc) working.textContent = "未打开文档";
        else if (view.error) working.textContent = view.error;
        else if (status && status.session) working.textContent = "临时预览或比较稿";
        else if (loading) working.textContent = "正在检查修改…";
        else if (!status) working.textContent = "无法读取文档状态";
        else if (!view.saved) working.textContent = "有未保存修改";
        else if (pendingMerge()) working.textContent = "合并审阅稿";
        else if (status.workingState === "noCommit") working.textContent = "尚无提交";
        else if (status.workingState === "clean") working.textContent = "与最新提交一致";
        else working.textContent = "有未提交修改";
        working.title = working.textContent;
        $("remoteCaption").textContent = remote && remote.linked ? remote.caption : "尚未关联远程仓库";
        $("remoteCaption").title = remote && remote.url ? remote.url + "\n" + (remote.caption || "") : $("remoteCaption").textContent;
    }

    function showBranches(rememberedBranch) {
        var filter = $("branchFilter");
        var selected = rememberedBranch == null ? filter.value : rememberedBranch;
        var branches = status && status.branches || [];
        filter.innerHTML = '<option value="">全部分支</option>' + branches.map(function (branch) {
            return '<option value="' + WG.escape(branch) + '">' +
                WG.escape(branch.replace(/^refs\/remotes\//, "远程 ")) + "</option>";
        }).join("");
        if (branches.indexOf(selected) >= 0) filter.value = selected;
    }

    async function fetchGraph(view, keep) {
        var filter = $("branchFilter").value, generation = ++graphGeneration;
        graphError = "";
        graphLoading = !!filter;
        updateButtons(view, keep);
        try {
            var data = filter ? await WG.call("history", { document: view.path, branch: filter }) : status;
            if (generation !== graphGeneration || view.key !== currentDocument().key) return;
            renderGraph(data.history || [], data.graph || [], keep);
        } catch (error) {
            if (generation === graphGeneration) {
                graphError = WG.errorText(error);
                renderGraph([], [], [], "无法读取提交记录");
            }
        } finally {
            if (generation === graphGeneration) {
                graphLoading = false;
                updateButtons(currentDocument());
            }
        }
    }

    function switchDocument(view) {
        if (currentKey === view.key) return;
        if (status && !status.session && graphView) {
            documentViews.set(currentKey, { branch: $("branchFilter").value, ids: graphView.selected(), scroll: $("graph").scrollTop });
        }
        currentKey = view.key;
        ++graphGeneration;
        status = null;
        items = [];
        selectedItem = null;
        graphError = "";
        graphLoading = false;
        reviewReady = false;
        reviewTicks = 0;
        $("reviewStatus").textContent = "";
        showBranches();
        renderGraph([], [], [], view.path ? "正在读取提交记录…" : "保存文档后即可提交版本。");
    }

    async function load() {
        if (loading) { reloadQueued = true; return; }
        loading = true;
        var view = currentDocument();
        var remembered = view.key !== currentKey && documentViews.get(view.key);
        switchDocument(view);
        var keep = remembered ? remembered.ids : graphView.selected();
        showStatus(view);
        updateButtons(view);
        try {
            if (!view.path) { loadItems(view); return; }
            var data = await WG.call("status", { document: view.path });
            if (view.key !== currentDocument().key) { reloadQueued = true; return; }
            status = data;
            showBranches(remembered ? remembered.branch : null);
            if (status.session) renderGraph([], [], [], "这份文档没有自己的版本历史。");
            else await fetchGraph(view, keep);
            if (view.key !== currentDocument().key) { reloadQueued = true; return; }
            if (remembered) $("graph").scrollTop = remembered.scroll;
            loadItems(view);
            app.PluginStorage.setItem("wordgit.pane.rendered", JSON.stringify({ commits: (status.history || []).length,
                branch: status.branch, at: Date.now() }));
        } catch (error) {
            status = null;
            showBranches();
            graphError = WG.errorText(error);
            renderGraph([], [], [], "无法读取提交记录");
            loadItems(view);
        } finally {
            loading = false;
            var active = currentDocument();
            if (active.key === view.key) {
                observedSave = active.saved;
                showStatus(active);
                updateButtons(active);
            } else reloadQueued = true;
            if (reloadQueued) { reloadQueued = false; setTimeout(load, 0); }
        }
    }

    function run(action, refresh) {
        if (currentDocument().key !== currentKey) { load(); return; }
        var result = action();
        if (refresh) Promise.resolve(result).then(function () {
            lastChange = app.PluginStorage.getItem("wordgit.changed") || "";
            load();
        });
    }

    $("branchFilter").onchange = function () {
        var view = currentDocument();
        if (status && !status.session && view.key === currentKey) fetchGraph(view, graphView ? graphView.selected() : []);
    };
    $("refresh").onclick = load;
    $("commit").onclick = function () { run(WG.commit, true); };
    $("compareLatest").onclick = function () { run(WG.compareLatest); };
    $("importReview").onclick = function () { run(WG.importReview, true); };
    $("link").onclick = function () { run(WG.linkRemote, true); };
    $("fetch").onclick = function () { run(WG.fetchRemote, true); };
    $("pull").onclick = function () { run(WG.pullRemote, true); };
    $("push").onclick = function () { run(WG.pushRemote, true); };
    $("preview").onclick = function () {
        var ids = graphView.selected();
        if (ids.length === 1) run(function () { return WG.preview(ids[0]); });
    };
    $("compare").onclick = function () {
        var ids = graphView.selected();
        if (!ids.length) return;
        var order = visibleHistory.map(function (r) { return r.id; });
        ids.sort(function (a, b) { return order.indexOf(b) - order.indexOf(a); });
        run(function () { return WG.compareVersions(ids[0], ids[1] || null); });
    };
    $("about").onclick = function () { WG.about(); };

    var splitter = $("splitter"), split = $("paneContent");
    function resizeReview(height) {
        var total = split.clientHeight;
        var available = Math.max(106, total - 235);
        var size = Math.min(available, Math.max(106, height));
        split.style.setProperty("--review-height", size + "px");
        splitter.setAttribute("aria-valuenow", String(Math.round(size / total * 100)));
    }
    splitter.onpointerdown = function (event) { splitter.setPointerCapture(event.pointerId); };
    splitter.onpointermove = function (event) {
        if (splitter.hasPointerCapture(event.pointerId)) resizeReview(event.clientY - split.getBoundingClientRect().top);
    };
    splitter.onkeydown = function (event) {
        if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
        event.preventDefault();
        resizeReview($("items").closest(".pane-review").offsetHeight + (event.key === "ArrowDown" ? 16 : -16));
    };

    try {
        ["WindowActivate", "DocumentOpen", "DocumentAfterClose", "DocumentAfterSave"].forEach(function (name) {
            app.ApiEvent.AddApiEventListener(name, function () { setTimeout(checkDocument, 150); });
        });
    } catch (error) { /* polling also detects document changes */ }

    function checkDocument() {
        if (document.hidden) return;
        var view = currentDocument();
        if (view.key !== currentKey || (view.saved !== observedSave && view.saved)) { load(); return; }
        // Typing changes the saved indicator immediately; history is re-read only after a save.
        observedSave = view.saved;
        if (!loading) showStatus(view);
    }

    setInterval(function () {
        if (document.hidden || resolving) return;
        var changed = app.PluginStorage.getItem("wordgit.changed") || "";
        if (changed !== lastChange) { lastChange = changed; load(); return; }
        checkDocument();
        if (loading) return;
        var view = currentDocument();
        if (!view.doc || view.key !== currentKey) return;
        reviewTicks++;
        if (reviewTicks % 4 === 0 || !reviewReady) {
            var signature;
            try { signature = WG.reviewSignature(view.doc); } catch (error) { signature = null; }
            if (signature !== reviewSignature || !reviewReady || reviewTicks >= 24) {
                loadItems(view);
                reviewTicks = 0;
            }
        }
    }, 1200);
    document.addEventListener("visibilitychange", checkDocument);
    lastChange = app.PluginStorage.getItem("wordgit.changed") || "";
    load();
})();
