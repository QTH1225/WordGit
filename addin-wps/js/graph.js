// Commit graph rendering shared by the task pane and the version picker. Lane layout comes from
// the service (crates/wordgit-core/src/graph.rs), identical to the original Windows add-in.
(function (global) {
    "use strict";
    var WG = global.WordGit;
    var COLORS = ["#3370ff", "#2ea121", "#f0820f", "#c71585", "#008b8b", "#8a2be2"];
    var LANE = 14, LEFT = 10;
    var TAG_ICON = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M2 2h6l6 6-6 6-6-6z" fill="none" stroke="currentColor"/><circle cx="5" cy="5" r="1" fill="currentColor"/></svg>';

    function lanesSvg(row, columns, height) {
        var width = LEFT * 2 + Math.max(1, columns || row.columns) * LANE;
        var parts = ['<svg class="lanes" width="' + width + '" height="' + height + '" viewBox="0 0 ' + width + " " + height + '" aria-hidden="true">'];
        (row.edges || []).forEach(function (edge) {
            var x1 = LEFT + edge.from * LANE, x2 = LEFT + edge.to * LANE;
            var y1 = edge.start * height / 2, y2 = edge.end * height / 2, bend = (y2 - y1) / 2;
            parts.push('<path d="M' + x1 + " " + y1 + " C" + x1 + " " + (y1 + bend) + " " + x2 + " " + (y2 - bend) + " " + x2 + " " + y2 +
                '" stroke="' + COLORS[edge.color % COLORS.length] + '" stroke-width="2" fill="none"/>');
        });
        var cx = LEFT + row.node * LANE, color = COLORS[row.node_color % COLORS.length];
        if (row.is_head) parts.push('<circle class="ring" cx="' + cx + '" cy="' + height / 2 + '" r="7" stroke="' + color + '" stroke-width="1.5" fill="none"/>');
        parts.push('<circle class="dot" cx="' + cx + '" cy="' + height / 2 + '" r="4" fill="' + color + '"/>');
        parts.push("</svg>");
        return parts.join("");
    }

    function references(revision, pane) {
        return (revision.references || []).map(function (ref) {
            var label = ref.replace(/^HEAD -> /, "");
            if (label === "HEAD") return "";
            var kind = label.indexOf("origin/") === 0 ? "remote" : "branch";
            var content = pane ? TAG_ICON + "<span>" + WG.escape(label.replace(/^tag: /, "")) + "</span>" : WG.escape(label);
            return '<span class="tag ' + kind + '">' + content + "</span>";
        }).join("");
    }

    // Renders rows into the container. Ctrl or Cmd toggles a second selection.
    WG.renderGraph = function (container, history, graph, head, options) {
        options = options || {};
        var selected = (options.selection || []).filter(function (id) {
            return history.some(function (revision) { return revision.id === id; });
        }).slice(0, 2);
        container.onclick = container.ondblclick = container.onkeydown = null;
        if (!history.length) {
            container.innerHTML = '<div class="empty">' + WG.escape(options.empty || "提交版本后，历史会显示在这里。") + "</div>";
            return { selected: function () { return []; } };
        }
        var columns = options.pane ? Math.max.apply(null, graph.map(function (row) { return row.columns; }).concat(1)) : 0;
        var height = options.pane ? 61 : 46;
        container.innerHTML = history.map(function (revision, index) {
            var tags = references(revision, options.pane);
            var meta = WG.shortId(revision.id) + " · " + WG.escape(revision.author) + " · " + WG.formatDate(revision.date);
            var caption = options.pane ? '<div class="meta">' + meta + '</div><div class="refs">' + tags + '</div>' :
                '<div class="meta">' + tags + meta + '</div>';
            var row = graph[index] || { columns: 1, node: 0, node_color: 0, is_head: revision.id === head };
            return '<div class="commit' + (revision.id === head ? " current" : "") + '" data-index="' + index +
                '" role="option" aria-selected="false" tabindex="0" title="' + WG.escape(revision.message) + '">' +
                lanesSvg(row, columns, height) +
                '<div class="text"><div class="subject">' + WG.escape(revision.subject) + "</div>" + caption + "</div></div>";
        }).join("");
        function paint() {
            Array.prototype.forEach.call(container.children, function (node) {
                var active = selected.indexOf(history[+node.dataset.index].id) >= 0;
                node.classList.toggle("selected", active);
                node.setAttribute("aria-selected", String(active));
            });
            if (options.onSelect) options.onSelect(selected.slice());
        }
        function select(node, toggle) {
            var id = history[+node.dataset.index].id;
            if (toggle && options.multi) {
                var at = selected.indexOf(id);
                if (at >= 0) selected.splice(at, 1);
                else {
                    selected.push(id);
                    if (selected.length > 2) selected.shift();
                }
            } else selected = [id];
            paint();
        }
        container.onclick = function (event) {
            var node = event.target.closest(".commit");
            if (node) select(node, event.ctrlKey || event.metaKey);
        };
        container.ondblclick = function (event) {
            var node = event.target.closest(".commit");
            if (node && options.onOpen) options.onOpen(history[+node.dataset.index].id);
        };
        container.onkeydown = function (event) {
            var node = event.target.closest(".commit");
            if (!node) return;
            if (event.key === " " || event.key === "Enter") {
                event.preventDefault();
                select(node, event.ctrlKey || event.metaKey);
                if (event.key === "Enter" && options.onOpen && selected.length === 1) options.onOpen(selected[0]);
            }
        };
        if (!selected.length && options.preselect) selected = [options.preselect];
        if (!selected.length && options.selectFirst) selected = [history[0].id];
        if (selected.length) paint();
        return { selected: function () { return selected.slice(); } };
    };
})(window);
