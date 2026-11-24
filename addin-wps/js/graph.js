// Commit graph rendering shared by the task pane and the version picker. Lane layout comes from
// the service (crates/wordgit-core/src/graph.rs), identical to the original Windows add-in.
(function (global) {
    "use strict";
    var WG = global.WordGit;
    var COLORS = ["#3370ff", "#2ea121", "#f0820f", "#c71585", "#008b8b", "#8a2be2"];
    var LANE = 14, ROW = 46, LEFT = 10;

    function lanesSvg(row) {
        var width = LEFT * 2 + Math.max(1, row.columns) * LANE;
        var parts = ['<svg class="lanes" width="' + width + '" height="' + ROW + '" viewBox="0 0 ' + width + " " + ROW + '">'];
        row.edges.forEach(function (edge) {
            var x1 = LEFT + edge.from * LANE, x2 = LEFT + edge.to * LANE;
            var y1 = edge.start * ROW / 2, y2 = edge.end * ROW / 2, bend = (y2 - y1) / 2;
            parts.push('<path d="M' + x1 + " " + y1 + " C" + x1 + " " + (y1 + bend) + " " + x2 + " " + (y2 - bend) + " " + x2 + " " + y2 +
                '" stroke="' + COLORS[edge.color % COLORS.length] + '" stroke-width="2" fill="none"/>');
        });
        var cx = LEFT + row.node * LANE, color = COLORS[row.node_color % COLORS.length];
        if (row.is_head) parts.push('<circle cx="' + cx + '" cy="' + ROW / 2 + '" r="7" stroke="' + color + '" stroke-width="1.5" fill="none"/>');
        parts.push('<circle cx="' + cx + '" cy="' + ROW / 2 + '" r="4" fill="' + color + '"/>');
        parts.push("</svg>");
        return parts.join("");
    }

    function references(revision) {
        var tags = "";
        (revision.references || []).forEach(function (ref) {
            var label = ref.replace(/^HEAD -> /, "");
            if (label === "HEAD") return;
            tags += '<span class="tag ' + (label.indexOf("origin/") === 0 ? "remote" : "branch") + '">' + WG.escape(label) + "</span>";
        });
        return tags;
    }

    // Renders rows into `container`; `onSelect(ids)` receives the selected ids (Ctrl for two).
    WG.renderGraph = function (container, history, graph, head, options) {
        options = options || {};
        var selected = [];
        if (!history.length) {
            container.innerHTML = '<div class="empty">' + WG.escape(options.empty || "提交版本后，历史会显示在这里。") + "</div>";
            return { selected: function () { return []; } };
        }
        container.innerHTML = history.map(function (revision, index) {
            return '<div class="commit' + (revision.id === head ? " current" : "") + '" data-index="' + index + '" title="' + WG.escape(revision.message) + '">' +
                lanesSvg(graph[index]) +
                '<div class="text"><div class="subject">' + WG.escape(revision.subject) + "</div>" +
                '<div class="meta">' + references(revision) + WG.shortId(revision.id) + " · " + WG.escape(revision.author) + " · " + WG.formatDate(revision.date) + "</div></div></div>";
        }).join("");
        function paint() {
            Array.prototype.forEach.call(container.children, function (node) {
                node.classList.toggle("selected", selected.indexOf(history[+node.dataset.index].id) >= 0);
            });
            if (options.onSelect) options.onSelect(selected.slice());
        }
        container.onclick = function (event) {
            var node = event.target.closest(".commit");
            if (!node) return;
            var id = history[+node.dataset.index].id;
            if ((event.ctrlKey || event.metaKey) && options.multi) {
                var at = selected.indexOf(id);
                if (at >= 0) selected.splice(at, 1);
                else { selected.push(id); if (selected.length > 2) selected.shift(); }
            } else {
                selected = [id];
            }
            paint();
        };
        container.ondblclick = function (event) {
            var node = event.target.closest(".commit");
            if (node && options.onOpen) options.onOpen(history[+node.dataset.index].id);
        };
        if (options.preselect) {
            selected = [options.preselect];
            paint();
        }
        return { selected: function () { return selected.slice(); } };
    };
})(window);
