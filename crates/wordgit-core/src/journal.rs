//! Makes "replace the working copy, then publish refs" recoverable.
//!
//! Before the document file is replaced, an intent record is written into the repository; it is
//! removed once both steps are done. If the editor or the service dies in between, `recover`
//! finds the record and either rolls the file back (refs were not published) or keeps it (they
//! were), so the document and the branch never silently disagree. The record format is shared
//! with the original Windows add-in.

use crate::fsutil;
use crate::history::Repo;
use crate::paths::AppPaths;
use crate::{Error, Result};
use base64::Engine;
use std::path::{Path, PathBuf};

pub const FILE_NAME: &str = "wordgit-replace.journal";
const FORMAT: &str = "WORDGIT-REPLACE 1";

#[derive(Clone, Debug)]
pub struct Entry {
    pub document: PathBuf,
    pub backup: PathBuf,
    pub original_hash: String,
    pub replacement_hash: String,
    pub references: String,
}

pub fn path_for(repo: &Repo) -> PathBuf {
    repo.path().join(FILE_NAME)
}

pub fn write(repo: &Repo, entry: &Entry) -> Result<()> {
    let text = format!(
        "{FORMAT}\n{}\n{}\n{}\n{}\n{}\n",
        entry.document.display(),
        entry.backup.display(),
        entry.original_hash,
        entry.replacement_hash,
        base64::engine::general_purpose::STANDARD.encode(entry.references.as_bytes())
    );
    fsutil::write_durable(&path_for(repo), text.as_bytes())
}

pub fn clear(repo: &Repo) {
    fsutil::try_remove(&path_for(repo));
}

pub fn read(repo: &Repo) -> Result<Option<Entry>> {
    let path = path_for(repo);
    let text = match std::fs::read(&path) {
        Ok(bytes) => String::from_utf8(bytes).map_err(|_| damaged(&path))?,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(error) => return Err(error.into()),
    };
    let lines: Vec<&str> = text.trim_end_matches('\n').split('\n').collect();
    if lines.len() != 6 || lines[0] != FORMAT || !Path::new(lines[1]).is_absolute() || !Path::new(lines[2]).is_absolute() {
        return Err(damaged(&path));
    }
    let references = base64::engine::general_purpose::STANDARD
        .decode(lines[5])
        .ok()
        .and_then(|bytes| String::from_utf8(bytes).ok())
        .ok_or_else(|| damaged(&path))?;
    Ok(Some(Entry {
        document: PathBuf::from(lines[1]),
        backup: PathBuf::from(lines[2]),
        original_hash: lines[3].into(),
        replacement_hash: lines[4].into(),
        references,
    }))
}

fn damaged(path: &Path) -> Error {
    Error::data(format!("替换记录已损坏：{}", path.display()))
}

fn file_name(path: &Path) -> String {
    path.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default()
}

/// Repairs an interrupted replacement. Returns a message for the user when something was
/// repaired or needs attention.
pub fn recover(repo: &Repo) -> Result<Option<String>> {
    let entry = match read(repo) {
        Ok(Some(entry)) => entry,
        Ok(None) => return Ok(None),
        Err(error) => {
            let path = path_for(repo);
            let stamp = crate::paths::compact_utc(crate::paths::now_seconds());
            let _ = std::fs::rename(&path, fsutil::sibling(&path, &format!(".{stamp}.corrupt")));
            return Ok(Some(error.to_string()));
        }
    };
    let current = fsutil::checksum(&entry.document)?;
    let published = repo.reference_state()? != entry.references;
    let mut message = None;
    if !published
        && current.as_deref() == Some(entry.replacement_hash.as_str())
        && fsutil::checksum(&entry.backup)?.as_deref() == Some(entry.original_hash.as_str())
    {
        // The file was swapped but the branch never moved: put the original back.
        if std::fs::rename(&entry.backup, &entry.document).is_err() {
            // Usually the document is held open; the record stays for the next attempt.
            return Ok(Some(format!(
                "上次切换或恢复在完成前中断，「{}」需要还原为操作前的内容。请先在编辑器中关闭这份文档，再重试；原稿备份在：{}",
                file_name(&entry.document),
                entry.backup.display()
            )));
        }
        fsutil::sync_parent(&entry.document);
        message = Some(format!("上次切换或恢复在完成前中断，已将「{}」还原为操作前的内容。", file_name(&entry.document)));
    } else if current.as_deref() == Some(if published { entry.replacement_hash.as_str() } else { entry.original_hash.as_str() }) {
        // Consistent already: either fully published, or interrupted before the swap.
        fsutil::try_remove(&entry.backup);
    } else if entry.backup.exists() {
        message = Some(format!(
            "上次替换「{}」时被中断，且文件之后又有变化。操作前的原稿备份保留在：{}",
            file_name(&entry.document),
            entry.backup.display()
        ));
    }
    clear(repo);
    Ok(message)
}

/// Called when the service starts, before documents are worked on, for every repository with a
/// record.
pub fn recover_all(paths: &AppPaths, executable: Option<&Path>) -> Vec<String> {
    let mut messages = Vec::new();
    let Ok(entries) = std::fs::read_dir(paths.repositories()) else { return messages };
    for entry in entries.flatten() {
        let folder = entry.path();
        if folder.extension().and_then(|e| e.to_str()) != Some("git") || !folder.join(FILE_NAME).is_file() {
            continue;
        }
        let outcome = Repo::open(folder.clone(), executable, None).and_then(|repo| recover(&repo));
        match outcome {
            Ok(Some(message)) => messages.push(message),
            Ok(None) => {}
            Err(error) => paths.log_error(&format!("working-copy recovery {}: {error}", folder.display())),
        }
    }
    messages
}
