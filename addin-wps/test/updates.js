"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");

function release(tag, options) {
    return Object.assign({
        tag_name: tag, draft: false, prerelease: tag.includes("-beta."),
        html_url: "https://example.invalid/",
        body: "更新日志"
    }, options || {});
}

async function run() {
    const api = "https://api.github.com/repos/QTH1225/WordGit/releases?per_page=25";
    const releases = [
        release("v0.12.3", {draft: true}),
        release("v0.12.1-beta.2"),
        release("v0.12.0"),
        release("v0.12.1-beta.10", {prerelease: false, assets: [
            {name: "WordGit-0.12.1-beta.10-linux-x86_64.tar.gz", state: "uploaded",
                browser_download_url: "https://example.invalid/"}
        ]})
    ];
    const files = new Map([["/addin/wordgit-config.json", '{"home":"/test"}']]);
    const calls = [];
    const scheduled = [];
    const win = {
        WORDGIT_VERSION: "0.12.0", navigator: {platform: "Linux x86_64"},
        WordGit: {
            readText: name => files.get(name),
            addinFolder: () => "/addin",
            dialog: async (page, title, width, height, data, modal) => {
                assert.equal(page, "update");
                assert.equal(modal, false);
                calls.push(data.result);
                if (data.result) win.WordGit.rememberUpdate(data.result);
            }
        },
        wps: {FileSystem: {WriteFile: (name, value) => files.set(name, value)}},
        fetch: async (url, options) => {
            assert.equal(url, api);
            assert.equal(options.redirect, "error");
            assert.equal(options.credentials, "omit");
            return {status: 200, ok: true, text: async () => JSON.stringify(releases)};
        }
    };
    const inject = () => vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../js/updates.js"), "utf8"), {
        window: win, AbortController,
        setTimeout: (fn, delay) => { if (delay === 5000) scheduled.push(fn); },
        clearTimeout: () => {}
    });
    inject();
    const wg = win.WordGit;
    assert.equal(wg.updatePreferences().beta, false);
    assert.equal((await wg.checkUpdates(false)).version, undefined);
    let result = await wg.checkUpdates(true);
    assert.equal(result.version, "0.12.1-beta.10");
    assert.equal(result.notes, "更新日志");
    assert.equal(result.url, "https://github.com/QTH1225/WordGit/releases/download/v0.12.1-beta.10/WordGit-0.12.1-beta.10-linux-x86_64.tar.gz");
    assert.equal(wg.isUpdateLink(result.url), true);
    assert.equal(wg.isUpdateLink("https://example.invalid/"), false);
    win.navigator.platform = "Linux aarch64";
    result = await wg.checkUpdates(true);
    assert.equal(result.direct, false);
    assert.ok(result.url.endsWith("/tag/v0.12.1-beta.10"));
    win.WORDGIT_VERSION = "0.12.1-beta.11";
    assert.equal((await wg.checkUpdates(true)).version, undefined);
    win.WORDGIT_VERSION = "0.12.0";
    wg.setUpdateBeta(true);
    wg.scheduleUpdateCheck();
    wg.scheduleUpdateCheck();
    assert.equal(scheduled.length, 1);
    await scheduled.shift()();
    assert.equal(calls.length, 1);
    inject(); // Simulate restarting WPS with its persisted preferences.
    assert.equal(wg.updatePreferences().beta, true);
    wg.scheduleUpdateCheck();
    await scheduled.shift()();
    assert.equal(calls.length, 1);
    await wg.showUpdates();
    assert.equal(calls.length, 2);
    assert.equal(calls[1], undefined);
    console.log("PASS: WPS stable/Beta ordering, platform download, persistence, startup dedup and manual entry");
}
run().catch(error => { console.error(error); process.exitCode = 1; });
