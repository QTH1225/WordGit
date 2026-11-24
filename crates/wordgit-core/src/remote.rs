//! Remote collaboration: one `origin` per document history, hosted on GitHub or Gitee.
//!
//! Fetches never prune and archive force-pushed-away tips under `refs/wordgit/remote-backup/`,
//! so a collaborator rewriting history cannot make already-fetched versions disappear.

use crate::fsutil;
use crate::git::{Git, NETWORK_IDLE_TIMEOUT, NULL_DEVICE};
use crate::history::{self, branch_ref, lines, Repo, HEADS};
use crate::{docx, object_id, Error, Result};
use serde::Serialize;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::sync::atomic::AtomicBool;

pub const REMOTE_PREFIX: &str = "refs/remotes/origin/";
const BACKUP_ROOT: &str = "refs/wordgit/remote-backup/";
/// A run of timestamped backups protects against remote history rewrites; the cap prevents
/// unbounded growth.
const BACKUP_RETENTION: usize = 20;
const HOSTS: [&str; 2] = ["github.com", "gitee.com"];

#[derive(Clone, Debug, Default, Serialize)]
pub struct RemoteState {
    pub url: Option<String>,
    pub local_branch: String,
    pub remote_branch: String,
    pub remote_reference: String,
    pub local_tip: Option<String>,
    pub remote_tip: Option<String>,
    pub ahead: usize,
    pub behind: usize,
    pub unrelated: bool,
    pub linked: bool,
    pub caption: String,
}

impl RemoteState {
    pub fn finish(mut self) -> Self {
        self.linked = self.url.as_deref().is_some_and(|u| !u.is_empty());
        self.caption = if !self.linked {
            "尚未关联远程仓库".into()
        } else if self.remote_tip.is_none() {
            format!("origin/{} · 尚未获取到该分支", self.remote_branch)
        } else if self.unrelated {
            format!("origin/{} · 没有共同历史", self.remote_branch)
        } else {
            format!("origin/{} · 领先 {} / 落后 {}", self.remote_branch, self.ahead, self.behind)
        };
        self
    }
}

/// `refs/remotes/origin/x` → `origin/x`.
pub fn reference_label(reference: &str) -> String {
    match reference.strip_prefix(REMOTE_PREFIX) {
        Some(rest) => format!("origin/{rest}"),
        None => reference.to_string(),
    }
}

/// Accepts only plain GitHub/Gitee HTTPS or SSH repository addresses (no credentials, queries
/// or command tricks) and returns the canonical form. `allow_local` admits an existing local
/// directory (tests and offline collaboration).
pub fn normalize_remote_url(url: &str, allow_local: bool) -> Result<String> {
    let url = url.trim();
    if allow_local {
        let path = Path::new(url);
        if path.is_absolute() && path.is_dir() {
            return Ok(fsutil::full_path(path).to_string_lossy().into_owned());
        }
    }
    let invalid = || Error::invalid("请输入 GitHub 或 Gitee 仓库的 HTTPS 或 SSH 地址，不包含账号口令、查询参数或其他命令。");
    let lower = url.to_ascii_lowercase();
    let mut parsed = None;
    for host in HOSTS {
        for prefix in [format!("https://{host}/"), format!("git@{host}:"), format!("ssh://git@{host}/")] {
            if lower.starts_with(&prefix) {
                parsed = Some((host, prefix.starts_with("https://"), &url[prefix.len()..]));
            }
        }
    }
    let (host, https, rest) = parsed.ok_or_else(invalid)?;
    let rest = rest.strip_suffix('/').unwrap_or(rest);
    let rest = if rest.to_ascii_lowercase().ends_with(".git") { &rest[..rest.len() - 4] } else { rest };
    let mut parts = rest.split('/');
    let (Some(owner), Some(name), None) = (parts.next(), parts.next(), parts.next()) else { return Err(invalid()) };
    let valid = |s: &str| {
        !s.is_empty() && s != "." && s != ".." && s.chars().all(|c| c.is_ascii_alphanumeric() || "_.-".contains(c))
    };
    if !valid(owner) || !valid(name) {
        return Err(invalid());
    }
    Ok(if https { format!("https://{host}/{owner}/{name}.git") } else { format!("git@{host}:{owner}/{name}.git") })
}

impl Repo {
    fn config_value(&self, key: &str) -> Result<Option<String>> {
        if !self.exists() {
            return Ok(None);
        }
        let result = self.invoke(None, None, &["config", "--local", "--get", key])?;
        if result.code == 1 {
            return Ok(None);
        }
        Ok(Some(result.text()?))
    }

    pub fn remote_url(&self) -> Result<Option<String>> {
        self.config_value("remote.origin.url")
    }

    pub fn configure_remote(&self, url: &str, remote_branch: &str, allow_local: bool) -> Result<()> {
        let url = normalize_remote_url(url, allow_local)?;
        let reference = branch_ref(remote_branch)?;
        self.ensure_repository()?;
        self.run(None, None, &["check-ref-format", &reference])?;
        let old = self.remote_url()?;
        if let Some(old) = &old {
            let same = normalize_remote_url(old, allow_local).map(|o| o.eq_ignore_ascii_case(&url)).unwrap_or(false);
            if !same {
                self.archive_remote_tips()?;
                self.clean_remote_backups()?;
                let previous = self.remote_branches()?;
                if !previous.is_empty() {
                    let mut operations = Vec::new();
                    for reference in &previous {
                        if let Some(tip) = self.branch_tip(reference)? {
                            operations.push(format!("delete {reference} {tip}"));
                        }
                    }
                    self.ref_transaction(None, &operations)?;
                }
                self.run(None, None, &["config", "--local", "--add", "wordgit.previousRemote", old])?;
                self.run(None, None, &["remote", "set-url", "origin", &url])?;
            }
        } else {
            self.run(None, None, &["remote", "add", "origin", &url])?;
        }
        self.run(None, None, &["config", "--local", "remote.origin.fetch", "+refs/heads/*:refs/remotes/origin/*"])?;
        let branch = self.active_branch()?;
        self.run(None, None, &["config", "--local", &format!("branch.{branch}.remote"), "origin"])?;
        self.run(None, None, &["config", "--local", &format!("branch.{branch}.merge"), &reference])?;
        Ok(())
    }

    pub fn remote_branches(&self) -> Result<Vec<String>> {
        if self.remote_url()?.is_none() {
            return Ok(Vec::new());
        }
        let head = format!("{REMOTE_PREFIX}HEAD");
        Ok(lines(&self.text(&["for-each-ref", "--format=%(refname)", REMOTE_PREFIX])?)
            .into_iter()
            .filter(|reference| *reference != head)
            .collect())
    }

    pub fn sync_state(&self) -> Result<RemoteState> {
        let local = self.active_branch()?;
        let merge = self.config_value(&format!("branch.{local}.merge"))?;
        let branch = merge.as_deref().and_then(|m| m.strip_prefix(HEADS)).unwrap_or(&local).to_string();
        let mut state = RemoteState {
            url: self.remote_url()?,
            local_tip: self.branch_tip(&local)?,
            remote_reference: format!("{REMOTE_PREFIX}{branch}"),
            local_branch: local,
            remote_branch: branch,
            ..Default::default()
        };
        if state.url.is_some() {
            state.remote_tip = self.branch_tip(&state.remote_reference)?;
        }
        self.complete_sync_state(state)
    }

    pub(crate) fn complete_sync_state(&self, mut state: RemoteState) -> Result<RemoteState> {
        let (Some(local), Some(remote)) = (state.local_tip.clone(), state.remote_tip.clone()) else {
            return Ok(state.finish());
        };
        if state.url.is_none() || local == remote {
            return Ok(state.finish());
        }
        let counts = self.text(&["rev-list", "--left-right", "--count", &format!("{local}...{remote}")])?;
        let mut numbers = counts.split_whitespace().map(|n| n.parse::<usize>().unwrap_or(0));
        state.ahead = numbers.next().unwrap_or(0);
        state.behind = numbers.next().unwrap_or(0);
        if state.ahead > 0 && state.behind > 0 {
            let common = self.invoke(None, None, &["merge-base", &local, &remote])?;
            if common.code == 1 {
                state.unrelated = true;
            } else {
                object_id(&common.text()?)?;
            }
        }
        Ok(state.finish())
    }

    fn archive_remote_tips(&self) -> Result<()> {
        let mut tips = BTreeMap::new();
        for reference in self.remote_branches()? {
            if let Some(tip) = self.branch_tip(&reference)? {
                tips.insert(reference, tip);
            }
        }
        self.archive_tips(&tips)
    }

    fn archive_tips(&self, tips: &BTreeMap<String, String>) -> Result<()> {
        if tips.is_empty() {
            return Ok(());
        }
        let prefix = format!(
            "{BACKUP_ROOT}{}-{}/",
            crate::paths::compact_utc(crate::paths::now_seconds()),
            &crate::random_hex(3)
        );
        let operations: Vec<String> = tips
            .iter()
            .map(|(reference, tip)| Ok(format!("create {prefix}{} {}", &reference[REMOTE_PREFIX.len()..], object_id(tip)?)))
            .collect::<Result<_>>()?;
        self.ref_transaction(None, &operations)
    }

    fn backup_refs(&self) -> Result<Vec<String>> {
        Ok(lines(&self.text(&["for-each-ref", "--format=%(refname)", BACKUP_ROOT])?))
    }

    pub fn remote_backup_namespaces(&self) -> Result<Vec<String>> {
        let mut namespaces: Vec<String> = Vec::new();
        for reference in self.backup_refs()? {
            let namespace = reference[BACKUP_ROOT.len()..].split('/').next().unwrap_or("").to_string();
            if !namespaces.contains(&namespace) {
                namespaces.push(namespace);
            }
        }
        Ok(namespaces)
    }

    /// Keeps only the newest timestamped backups; GUID namespaces left by older versions always
    /// count as oldest.
    fn clean_remote_backups(&self) -> Result<()> {
        let mut groups: BTreeMap<String, Vec<String>> = BTreeMap::new();
        for reference in self.backup_refs()? {
            let namespace = reference[BACKUP_ROOT.len()..].split('/').next().unwrap_or("").to_string();
            groups.entry(namespace).or_default().push(reference);
        }
        let sort_key = |name: &str| {
            let timestamped = name.len() >= 8 && name.as_bytes()[..8].iter().all(u8::is_ascii_digit);
            format!("{}{name}", if timestamped { "1" } else { "0" })
        };
        let mut ordered: Vec<(String, Vec<String>)> = groups.into_iter().collect();
        ordered.sort_by_key(|group| std::cmp::Reverse(sort_key(&group.0)));
        let mut operations = Vec::new();
        for (_, references) in ordered.into_iter().skip(BACKUP_RETENTION) {
            for reference in references {
                let result = self.invoke(None, None, &["rev-parse", "--verify", "-q", &reference])?;
                if result.code == 0 {
                    operations.push(format!("delete {reference} {}", result.text()?));
                }
            }
        }
        if !operations.is_empty() {
            self.ref_transaction(None, &operations)?;
        }
        Ok(())
    }

    pub fn fetch_remote(&self, allow_local: bool, cancel: &AtomicBool) -> Result<RemoteState> {
        let url = self.remote_url()?.ok_or_else(|| Error::invalid("请先关联远程仓库。"))?;
        normalize_remote_url(&url, allow_local)?;
        let mut before = BTreeMap::new();
        for reference in self.remote_branches()? {
            if let Some(tip) = self.branch_tip(&reference)? {
                before.insert(reference, tip);
            }
        }
        self.network_run(
            allow_local,
            cancel,
            &[
                "-c", "transfer.fsckObjects=true", "-c", "fetch.fsckObjects=true", "fetch", "--progress", "--no-prune",
                "--no-tags", "--no-recurse-submodules", "origin", "+refs/heads/*:refs/remotes/origin/*",
            ],
        )?;
        // Only a forced (non-fast-forward) update can hide commits already fetched; their old
        // tips are kept under refs/wordgit/remote-backup so they stay in the history.
        let mut rewritten = BTreeMap::new();
        for (reference, old) in before {
            let now = self.branch_tip(&reference)?;
            let kept = match &now {
                Some(now) => *now == old || self.is_ancestor(&old, now)?,
                None => false,
            };
            if !kept {
                rewritten.insert(reference, old);
            }
        }
        self.archive_tips(&rewritten)?;
        self.clean_remote_backups()?;
        self.sync_state()
    }

    pub fn push_remote(&self, allow_local: bool, cancel: &AtomicBool) -> Result<RemoteState> {
        let state = self.sync_state()?;
        if !state.linked || state.local_tip.is_none() {
            return Err(Error::invalid("请先关联仓库并提交当前稿。"));
        }
        normalize_remote_url(state.url.as_deref().unwrap_or(""), allow_local)?;
        // A normal push: the server refuses non-fast-forward updates.
        let refspec = format!("{}:{}", branch_ref(&state.local_branch)?, branch_ref(&state.remote_branch)?);
        self.network_run(allow_local, cancel, &["push", "--progress", "--porcelain", "origin", &refspec])?;
        self.sync_state()
    }

    pub fn fast_forward_remote(&self, branch: &str, old_tip: &str, reference: &str, remote_tip: &str) -> Result<()> {
        if !reference.starts_with(REMOTE_PREFIX)
            || self.active_branch()? != branch
            || self.branch_tip(branch)?.as_deref() != Some(old_tip)
            || self.branch_tip(reference)?.as_deref() != Some(remote_tip)
        {
            return Err(Error::invalid("本地或远程跟踪分支已变化，请重新拉取。"));
        }
        let ancestor = self.invoke(None, None, &["merge-base", "--is-ancestor", &object_id(old_tip)?, &object_id(remote_tip)?])?;
        if ancestor.code != 0 {
            return Err(Error::invalid("双方已有不同修改，需要通过修订合并。"));
        }
        self.ref_transaction(
            None,
            &[
                format!("update {} {} {}", branch_ref(branch)?, object_id(remote_tip)?, object_id(old_tip)?),
                format!("verify {} {}", history::history_reference(reference)?, object_id(remote_tip)?),
            ],
        )
    }

    fn network_run(&self, allow_local: bool, cancel: &AtomicBool, arguments: &[&str]) -> Result<()> {
        normalize_remote_url(&self.remote_url()?.unwrap_or_default(), allow_local)?;
        let hooks = format!("core.hooksPath={NULL_DEVICE}");
        let file_protocol = if allow_local { "protocol.file.allow=always" } else { "protocol.file.allow=never" };
        let mut args: Vec<&str> = vec![
            "-c", &hooks, "-c", "gc.auto=0", "-c", "maintenance.auto=false", "-c", "push.followTags=false", "-c",
            "push.recurseSubmodules=no", "-c", "remote.origin.mirror=false", "-c", "protocol.ext.allow=never", "-c",
            file_protocol,
        ];
        args.extend_from_slice(arguments);
        let mut command = self.git().command(None, &args);
        // Restore the user's Git credential/proxy configuration only for network operations.
        command.env_remove("GIT_CONFIG_NOSYSTEM").env_remove("GIT_CONFIG_GLOBAL").env("GIT_TERMINAL_PROMPT", "0");
        configure_ssh(self.git(), &mut command);
        self.git().run_network(command, cancel, NETWORK_IDLE_TIMEOUT)
    }
}

/// Keeps a user's own SSH setup (a specific key, plink, custom known_hosts). Otherwise ssh must
/// never wait for terminal input, and a first connection records the host key.
fn configure_ssh(git: &Git, command: &mut std::process::Command) {
    let mut custom = false;
    for name in ["GIT_SSH_COMMAND", "GIT_SSH", "GIT_SSH_VARIANT"] {
        if let Some(value) = std::env::var_os(name).filter(|v| !v.is_empty()) {
            command.env(name, value);
            custom |= name != "GIT_SSH_VARIANT";
        }
    }
    if custom {
        return;
    }
    if let Ok(configured) = git.invoke_with_user_config(&["config", "--get", "core.sshCommand"]) {
        if configured.code == 0 && !configured.stdout.is_empty() {
            return;
        }
    }
    command.env("GIT_SSH_COMMAND", "ssh -o BatchMode=yes -o StrictHostKeyChecking=accept-new");
}

/// Downloads a remote WordGit history and restores its document at `document` (which must not
/// exist). The history only takes its final place once the document has been verified.
pub fn clone_remote(
    url: &str,
    remote_branch: &str,
    document: &Path,
    storage: &Path,
    executable: Option<&Path>,
    allow_local: bool,
    cancel: &AtomicBool,
) -> Result<Repo> {
    let url = normalize_remote_url(url, allow_local)?;
    let document = fsutil::full_path(document);
    let is_docx = document.extension().and_then(|e| e.to_str()).is_some_and(|e| e.eq_ignore_ascii_case("docx"));
    if !is_docx || document.exists() {
        return Err(Error::invalid("请选择一个尚不存在的本地 .docx 文件名。"));
    }
    let destination = Repo::for_document(&document, storage, executable, None)?;
    if destination.path().exists() {
        return Err(Error::invalid("该路径已有历史仓库，请选择新的文件名；原历史不会被替换。"));
    }
    let parent = destination.path().parent().map(Path::to_path_buf).unwrap_or_else(|| storage.to_path_buf());
    std::fs::create_dir_all(&parent)?;
    let folder = document.parent().map(Path::to_path_buf).ok_or_else(|| Error::invalid("无效的文档路径。"))?;
    let temporary: PathBuf = parent.join(format!("clone-{}.git", crate::random_hex(6)));
    let staged = folder.join(format!("wordgit-{}.tmp", crate::random_hex(16)));
    let result = (|| -> Result<()> {
        let repo = Repo::open(temporary.clone(), Some(destination.git().executable()), None)?;
        repo.ensure_repository()?;
        let reference = branch_ref(remote_branch)?;
        repo.run(None, None, &["check-ref-format", &reference])?;
        repo.run(None, None, &["symbolic-ref", "HEAD", &reference])?;
        repo.configure_remote(&url, remote_branch, allow_local)?;
        let state = repo.fetch_remote(allow_local, cancel)?;
        let tip = state
            .remote_tip
            .ok_or_else(|| Error::invalid("远程没有所选分支，请确认分支名或先由另一位作者推送版本。"))?;
        {
            let mut output = std::fs::OpenOptions::new().write(true).create_new(true).open(&staged)?;
            repo.extract_to(&tip, &mut output)?;
            output.sync_all()?;
        }
        docx::validate_file(&staged)?;
        repo.ref_transaction(None, &[format!("create {reference} {tip}")])?;
        std::fs::rename(&temporary, destination.path())?;
        if document.exists() {
            return Err(Error::invalid("请选择一个尚不存在的本地 .docx 文件名。"));
        }
        std::fs::rename(&staged, &document)?;
        fsutil::sync_parent(&document);
        // Content written by collaborators is treated like a download (Windows: Protected View).
        mark_from_internet(&document);
        Ok(())
    })();
    fsutil::try_remove(&staged);
    if temporary.is_dir() {
        let name = temporary.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default();
        if name.starts_with("clone-") && name.ends_with(".git") {
            let _ = std::fs::remove_dir_all(&temporary);
        }
    }
    result.map(|_| destination)
}

/// Writes the Windows "mark of the web" so Office treats a file like a download.
pub fn mark_from_internet(path: &Path) {
    #[cfg(windows)]
    {
        let mut stream = path.as_os_str().to_os_string();
        stream.push(":Zone.Identifier");
        let _ = std::fs::write(PathBuf::from(stream), b"[ZoneTransfer]\r\nZoneId=3\r\n");
    }
    #[cfg(not(windows))]
    let _ = path;
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn remote_urls() {
        assert_eq!(normalize_remote_url("https://github.com/a/b", false).unwrap(), "https://github.com/a/b.git");
        assert_eq!(normalize_remote_url("git@github.com:a/b.git", false).unwrap(), "git@github.com:a/b.git");
        assert_eq!(normalize_remote_url("ssh://git@gitee.com/zuo/x.git/", false).unwrap(), "git@gitee.com:zuo/x.git");
        assert!(normalize_remote_url("https://user:pw@github.com/a/b", false).is_err());
        assert!(normalize_remote_url("https://github.com/a/b?x=1", false).is_err());
        assert!(normalize_remote_url("https://github.com/../b", false).is_err());
        assert!(normalize_remote_url("https://example.com/a/b", false).is_err());
        assert!(normalize_remote_url("--upload-pack=x", false).is_err());
    }

    #[test]
    fn captions() {
        let state = RemoteState { remote_branch: "main".into(), ..Default::default() }.finish();
        assert_eq!(state.caption, "尚未关联远程仓库");
        assert_eq!(reference_label("refs/remotes/origin/dev"), "origin/dev");
    }
}
