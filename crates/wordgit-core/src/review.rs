//! Pending merge drafts and the working-copy swap.
//!
//! A merge draft lives in the sessions folder next to `<draft>.wordgit`, which ties it back to the
//! original document and the branch tip it was generated from (format shared with the original
//! Windows add-in). Replacing the user's document always goes through [`replace_working_copy`],
//! which journals the intent first so a crash can be repaired (see [`crate::journal`]).

use crate::fsutil;
use crate::history::Repo;
use crate::journal::{self, Entry};
use crate::paths::AppPaths;
use crate::{Error, Result};
use serde::Serialize;
use std::path::{Path, PathBuf};
use std::time::{Duration, SystemTime};

const APPLY_MARKER: &str = "apply-to-original";
const RECORD_MARKER: &str = "record-merge";

#[derive(Clone, Debug, Default, Serialize)]
pub struct PendingReview {
    pub original: PathBuf,
    pub head: String,
    pub branch: String,
    pub source_branch: Option<String>,
    pub source_tip: Option<String>,
    /// A returned draft merged into the current document (no source branch).
    pub apply_to_original: bool,
    /// Two historical versions merged; finishing records a normal commit.
    pub record_only_merge: bool,
}

impl PendingReview {
    pub fn is_pending_merge(&self) -> bool {
        self.apply_to_original || self.record_only_merge || self.source_branch.is_some()
    }

    fn metadata_path(draft: &Path) -> PathBuf {
        fsutil::sibling(draft, ".wordgit")
    }

    pub fn load(paths: &AppPaths, draft: &Path) -> Result<Option<PendingReview>> {
        if !paths.is_session_file(draft) {
            return Ok(None);
        }
        let text = match std::fs::read_to_string(Self::metadata_path(draft)) {
            Ok(text) => text,
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
            Err(error) => return Err(error.into()),
        };
        let saved: Vec<&str> = text.lines().collect();
        let apply = saved.len() == 6 && saved[5] == APPLY_MARKER;
        let record = saved.len() == 6 && saved[5] == RECORD_MARKER;
        if (saved.len() != 2 && saved.len() != 5 && !apply && !record) || !Path::new(saved[0]).is_absolute() {
            return Err(Error::data("合并稿的版本关联已损坏。"));
        }
        let optional = |index: usize| saved.get(index).filter(|v| !v.is_empty()).map(|v| v.to_string());
        Ok(Some(PendingReview {
            original: PathBuf::from(saved[0]),
            head: saved[1].to_string(),
            branch: if saved.len() == 2 { "main".into() } else { saved[2].to_string() },
            source_branch: if saved.len() >= 5 { optional(3) } else { None },
            source_tip: if saved.len() >= 5 { optional(4) } else { None },
            apply_to_original: apply,
            record_only_merge: record,
        }))
    }

    pub fn save(&self, draft: &Path, head: &str) -> Result<()> {
        let mut lines = vec![
            self.original.to_string_lossy().into_owned(),
            head.to_string(),
            self.branch.clone(),
            self.source_branch.clone().unwrap_or_default(),
            self.source_tip.clone().unwrap_or_default(),
        ];
        if self.apply_to_original {
            lines.push(APPLY_MARKER.into());
        } else if self.record_only_merge {
            lines.push(RECORD_MARKER.into());
        }
        let newline = if cfg!(windows) { "\r\n" } else { "\n" };
        let text = lines.join(newline) + newline;
        fsutil::write_durable(&Self::metadata_path(draft), text.as_bytes())
    }
}

/// The document whose history an action uses: a pending merge draft maps back to its original.
/// Other files under the sessions folder are previews and have no history.
pub fn history_path(paths: &AppPaths, document: &Path) -> Result<PathBuf> {
    if let Some(review) = PendingReview::load(paths, document)? {
        return Ok(review.original);
    }
    if paths.is_session_file(document) {
        return Err(Error::invalid(
            "这是 WordGit 生成的临时预览或比较稿，没有自己的版本历史。如需保留，请先「另存为」到其他位置。",
        ));
    }
    Ok(document.to_path_buf())
}

/// Rejects locations where replacing the file could fight a sync client or a network share.
pub fn ensure_replaceable(path: &Path) -> Result<()> {
    let text = path.to_string_lossy();
    if text.starts_with("\\\\") || text.starts_with("//") {
        return Err(Error::invalid("请在本地文件副本上切换、恢复或合并分支。"));
    }
    for variable in ["OneDrive", "OneDriveCommercial", "OneDriveConsumer"] {
        if let Some(folder) = std::env::var_os(variable).filter(|v| !v.is_empty()) {
            if crate::paths::starts_with_folder(path, Path::new(&folder)) {
                return Err(Error::invalid("为避免覆盖共编修改，请在 OneDrive 同步目录之外的本地副本上切换、恢复或合并分支。"));
            }
        }
    }
    Ok(())
}

/// Swaps the working copy for a verified replacement and then publishes the matching ref change.
/// The editor must have closed the document before this is called (and reopens it afterwards).
pub fn replace_working_copy(
    repo: &Repo,
    path: &Path,
    current_hash: &str,
    replacement_file: &Path,
    replacement_hash: &str,
    publish: &mut dyn FnMut() -> Result<()>,
) -> Result<()> {
    ensure_replaceable(path)?;
    if let Some(leftover) = journal::recover(repo)? {
        return Err(Error::invalid(format!("{leftover}\n请确认当前稿后重试。")));
    }
    if fsutil::checksum(path)?.as_deref() != Some(current_hash) {
        return Err(Error::invalid("磁盘上的当前稿已变化，未替换文件。"));
    }
    if current_hash == replacement_hash {
        return publish();
    }
    let folder = path.parent().ok_or_else(|| Error::invalid("无效的文档路径。"))?;
    clean_restore_leftovers(folder);
    let staged = folder.join(format!(".wordgit-restore-{}.tmp", crate::random_hex(16)));
    let backup = fsutil::sibling(&staged, ".bak");
    std::fs::copy(replacement_file, &staged)?;
    if fsutil::checksum(&staged)?.as_deref() != Some(replacement_hash) {
        fsutil::try_remove(&staged);
        return Err(Error::data("替换内容在操作期间发生变化，未覆盖当前稿。"));
    }
    journal::write(
        repo,
        &Entry {
            document: path.to_path_buf(),
            backup: backup.clone(),
            original_hash: current_hash.into(),
            replacement_hash: replacement_hash.into(),
            references: repo.reference_state()?,
        },
    )?;
    let mut replaced = false;
    let result = (|| -> Result<()> {
        if fsutil::checksum(path)?.as_deref() != Some(current_hash) {
            return Err(Error::invalid("磁盘文件在操作期间发生变化，未覆盖。请重新打开后重试。"));
        }
        fsutil::replace_with_backup(&staged, path, &backup)?;
        replaced = true;
        if let Err(error) = publish() {
            if fsutil::checksum(path)?.as_deref() != Some(replacement_hash) {
                return Err(Error::invalid(format!("当前稿在操作期间被其他程序修改，原稿备份保留在：{}", backup.display())));
            }
            std::fs::rename(&backup, path)?;
            fsutil::sync_parent(path);
            replaced = false;
            return Err(error);
        }
        fsutil::try_remove(&backup);
        replaced = false;
        Ok(())
    })();
    // A replaced file whose rollback failed keeps its journal for recovery on restart.
    if !replaced {
        journal::clear(repo);
    }
    fsutil::try_remove(&staged);
    result
}

/// Removes stale staging files (a day old) and backups (kept a month: they are the only copy of
/// a draft when a rollback failed).
pub fn clean_restore_leftovers(folder: &Path) {
    let now = SystemTime::now();
    let Ok(entries) = std::fs::read_dir(folder) else { return };
    for entry in entries.flatten() {
        let name = entry.file_name().to_string_lossy().into_owned();
        if !name.starts_with(".wordgit-restore-") {
            continue;
        }
        let age = if name.ends_with(".tmp.bak") {
            Duration::from_secs(30 * 86_400)
        } else if name.ends_with(".tmp") {
            Duration::from_secs(86_400)
        } else {
            continue;
        };
        let old = entry
            .metadata()
            .and_then(|m| m.modified())
            .map(|t| now.duration_since(t).unwrap_or_default() > age)
            .unwrap_or(false);
        if old {
            fsutil::try_remove(&entry.path());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pending_review_round_trip() {
        let paths = AppPaths::new(std::env::temp_dir().join(format!("wg-review-{}", crate::random_hex(4))));
        let session = paths.new_session().unwrap();
        let draft = session.join("合并稿.docx");
        let original = std::env::temp_dir().join("原稿.docx");
        let review = PendingReview {
            original: original.clone(),
            branch: "main".into(),
            source_branch: Some("review-1".into()),
            source_tip: Some("a".repeat(40)),
            ..Default::default()
        };
        review.save(&draft, &"b".repeat(40)).unwrap();
        let loaded = PendingReview::load(&paths, &draft).unwrap().unwrap();
        assert_eq!(loaded.original, original);
        assert_eq!(loaded.head, "b".repeat(40));
        assert_eq!(loaded.source_branch.as_deref(), Some("review-1"));
        assert!(loaded.is_pending_merge());
        assert_eq!(history_path(&paths, &draft).unwrap(), original);
        assert!(history_path(&paths, &session.join("预览.docx")).is_err());
        let _ = std::fs::remove_dir_all(&paths.root);
    }
}
