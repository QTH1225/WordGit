(function (global) {
    "use strict";
    var WG = global.WordGit;
    var api = "https://api.github.com/repos/QTH1225/WordGit/releases?per_page=25";
    var releasesUrl = "https://github.com/QTH1225/WordGit/releases";
    var dialogOpen = false;
    var started = false;
    var checkedManually = false;

    function version(text) {
        var match = typeof text === "string" &&
            /^v?(0|[1-9][0-9]{0,8})\.(0|[1-9][0-9]{0,8})\.(0|[1-9][0-9]{0,8})(?:\.(0|[1-9][0-9]{0,8})|-beta\.([1-9][0-9]{0,8}))?$/.exec(text);
        if (!match || match[0] !== text) return null;
        return { text: text.replace(/^v/, ""), number: [Number(match[1]), Number(match[2]), Number(match[3]), Number(match[4] || 0)],
            beta: Number(match[5] || 0) };
    }

    function compare(left, right) {
        for (var i = 0; i < left.number.length; i++) {
            if (left.number[i] !== right.number[i]) return left.number[i] - right.number[i];
        }
        if (left.beta === right.beta) return 0;
        if (left.beta === 0) return 1;
        if (right.beta === 0) return -1;
        return left.beta - right.beta;
    }

    function assetNames(value, platform) {
        if (/^Win/.test(platform)) return ["WordGit-" + value + "-windows.zip"];
        if (/^Linux (x86_64|amd64)$/.test(platform)) return ["WordGit-" + value + "-linux-x86_64.tar.gz"];
        if (/^Mac/i.test(platform)) {
            return [
                "WordGit-" + value + "-macos-universal.tar.gz",
                "WordGit-" + value + "-macos-arm64.tar.gz",
                "WordGit-" + value + "-macos-x86_64.tar.gz"
            ];
        }
        return [];
    }

    function selectRelease(releases, current, beta, platform) {
        var installed = version(current);
        if (!installed) return { message: "当前版本号无法识别。" };
        if (!Array.isArray(releases) || releases.some(function (entry) {
            return !entry || typeof entry.draft !== "boolean" || typeof entry.prerelease !== "boolean";
        })) return { message: "GitHub 返回的数据无法识别。" };

        var latest, selected;
        releases.forEach(function (entry) {
            var candidate = version(entry.tag_name);
            if (!candidate || entry.draft || (candidate.beta && !beta) || (entry.prerelease && !candidate.beta)) return;
            if (!latest || compare(candidate, latest) > 0) {
                latest = candidate;
                selected = entry;
            }
        });
        if (!latest) return { message: "当前渠道暂无可用版本。" };
        var order = compare(latest, installed);
        if (order < 0) return { message: "当前版本高于此渠道的最新版本（" + latest.text + "）。" };
        if (!order) return { message: "当前已是最新版（" + latest.text + "）。" };
        var names = assetNames(latest.text, platform);
        var matched = Array.isArray(selected.assets) ? names.find(function (name) {
            return selected.assets.some(function (asset) {
                return asset && asset.name === name && asset.state === "uploaded";
            });
        }) : null;
        var notes = typeof selected.body === "string" ? selected.body.trim() : "";
        return {
            version: latest.text,
            notes: notes.length > 24000 ? notes.slice(0, 24000) + "\n…" : notes || "此版本未填写更新日志。",
            url: matched ? releasesUrl + "/download/" + selected.tag_name + "/" + matched : releasesUrl + "/tag/" + selected.tag_name,
            direct: !!matched
        };
    }

    function preferencesFile() {
        var config = JSON.parse(WG.readText(WG.addinFolder() + "/wordgit-config.json") || "{}");
        if (typeof config.home !== "string" || !config.home) return null;
        return config.home.replace(/[\\/]+$/, "") + "/wps-updates.json";
    }

    function preferences() {
        try {
            var path = preferencesFile();
            var value = path && JSON.parse(WG.readText(path) || "{}") || {};
            return { beta: value.beta === true, ignoredVersion: typeof value.ignoredVersion === "string" ? value.ignoredVersion : "" };
        } catch (error) {
            return { beta: false, ignoredVersion: "" };
        }
    }

    function save(value) {
        var path = preferencesFile();
        if (!path) throw new Error("无法找到 WordGit 更新设置目录。");
        global.wps.FileSystem.WriteFile(path, JSON.stringify(value));
        var saved = preferences();
        if (saved.beta !== value.beta || saved.ignoredVersion !== value.ignoredVersion)
            throw new Error("无法保存 WordGit 更新设置。");
    }

    WG.updatePreferences = preferences;
    WG.setUpdateBeta = function (enabled) {
        var settings = preferences();
        settings.beta = enabled;
        save(settings);
    };

    WG.checkUpdates = async function (beta) {
        var controller = new AbortController();
        var timer = setTimeout(function () { controller.abort(); }, 12000);
        try {
            var response = await global.fetch(api, { credentials: "omit", redirect: "error", signal: controller.signal });
            if (response.status === 403 || response.status === 429)
                return { message: "GitHub 暂时限制了更新查询，请稍后重试。" };
            if (response.status === 404) return { message: "当前渠道暂无可用版本。" };
            if (!response.ok) return { message: "检查更新失败（HTTP " + response.status + "）。" };
            var data = await response.text();
            if (data.length > 1024 * 1024) return { message: "GitHub 返回的数据无法识别。" };
            return selectRelease(JSON.parse(data), global.WORDGIT_VERSION, beta, global.navigator.platform);
        } catch (error) {
            return { message: controller.signal.aborted ? "检查更新超时，请稍后重试。" : "无法读取 GitHub 更新，请稍后重试。" };
        } finally {
            clearTimeout(timer);
        }
    };

    WG.ignoreUpdate = function (release) {
        var candidate = version(release);
        if (!candidate) throw new Error("无法识别要忽略的版本。");
        var settings = preferences();
        var ignored = version(settings.ignoredVersion);
        if (!ignored || compare(candidate, ignored) > 0) {
            settings.ignoredVersion = candidate.text;
            save(settings);
        }
    };

    WG.showUpdates = async function (result) {
        if (!result) { started = true; checkedManually = true; }
        if (dialogOpen) return;
        dialogOpen = true;
        try {
            await WG.dialog("update", result ? "WordGit · 发现新版本" : "WordGit · 检查更新", 540, 410, { result: result }, false);
        } finally {
            dialogOpen = false;
        }
    };

    WG.scheduleUpdateCheck = function () {
        if (started) return;
        started = true;
        setTimeout(async function () {
            if (checkedManually) return;
            var result = await WG.checkUpdates(preferences().beta);
            var current = preferences();
            var candidate = version(result.version), ignored = version(current.ignoredVersion);
            if (!checkedManually && candidate && (!ignored || compare(candidate, ignored) > 0) &&
                (current.beta || !candidate.beta)) {
                try { await WG.showUpdates(result); } catch (error) { /* host is busy or closing */ }
            }
        }, 5000);
    };

    WG.isUpdateLink = function (url) {
        if (typeof url !== "string" || url.indexOf(releasesUrl + "/") !== 0) return false;
        var parts = url.slice(releasesUrl.length + 1).split("/");
        var parsed = version(parts[1]);
        if (!parsed) return false;
        return (parts[0] === "tag" && parts.length === 2) ||
            (parts[0] === "download" && parts.length === 3 && assetNames(parsed.text, global.navigator.platform).indexOf(parts[2]) >= 0);
    };
})(window);
