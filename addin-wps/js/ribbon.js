// WordGit for WPS Writer — ribbon callbacks (loaded by the hidden main page).
(function (global) {
    "use strict";
    var WG = global.WordGit;

    var ACTIONS = {
        WordGitCommit: function () { return WG.commit(); },
        WordGitHistory: function () { WG.showPane(); },
        WordGitPreview: function () { return WG.preview(); },
        WordGitCreateBranch: function () { return WG.createBranch(); },
        WordGitSwitchBranch: function () { return WG.switchBranch(); },
        WordGitMergeBranch: function () { return WG.mergeBranch(); },
        WordGitDeleteBranch: function () { return WG.deleteBranch(); },
        WordGitImportReview: function () { return WG.importReview(); },
        WordGitCompare: function () { return WG.compareVersions(); },
        WordGitLinkRemote: function () { return WG.linkRemote(); },
        WordGitCloneRemote: function () { return WG.cloneRemote(); },
        WordGitFetchRemote: function () { return WG.fetchRemote(); },
        WordGitPullRemote: function () { return WG.pullRemote(); },
        WordGitPushRemote: function () { return WG.pushRemote(); },
        WordGitExport: function () { return WG.exportVersion(); },
        WordGitRestore: function () { return WG.restoreVersion(); },
        WordGitAbout: function () { return WG.about(); }
    };

    var IMAGES = {
        WordGitCommit: "commit", WordGitHistory: "graph", WordGitPreview: "preview",
        WordGitCreateBranch: "branch-new", WordGitSwitchBranch: "branch-switch", WordGitMergeBranch: "merge",
        WordGitDeleteBranch: "branch-delete", WordGitImportReview: "review", WordGitCompare: "compare",
        WordGitLinkRemote: "link", WordGitCloneRemote: "clone", WordGitFetchRemote: "fetch",
        WordGitPullRemote: "pull", WordGitPushRemote: "push", WordGitExport: "export", WordGitRestore: "restore",
        WordGitAbout: "about"
    };

    global.OnAddinLoad = function (ribbonUI) {
        var app = WG.app();
        if (app && typeof app.ribbonUI !== "object") app.ribbonUI = ribbonUI;
        setTimeout(WG.showNotices, 3000);
        return true;
    };

    global.OnAction = function (control) {
        var action = ACTIONS[control.Id];
        if (action) {
            // Ribbon callbacks must return immediately; flows run asynchronously.
            setTimeout(function () {
                Promise.resolve(action()).catch(function (error) { WG.notify(WG.errorText(error)); });
            }, 0);
        }
        return true;
    };

    global.GetImage = function (control) {
        return "images/" + (IMAGES[control.Id] || "about") + ".svg";
    };

    global.GetLabel = function (control) {
        if (control.Id === "WordGitAbout") return "WordGit " + (global.WORDGIT_VERSION || "");
        return "";
    };
})(window);
