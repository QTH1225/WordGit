//! File helpers shared by the history and working-copy code: checksums, exclusive repository
//! locks, durable small-file writes and atomic replacement with a backup.

use crate::{Error, Result};
use sha2::{Digest, Sha256};
use std::fs::{self, File, OpenOptions};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::time::SystemTime;

/// SHA-256 of a file's contents, or `None` when it does not exist.
pub fn checksum(path: &Path) -> Result<Option<String>> {
    let mut file = match open_shared(path) {
        Ok(file) => file,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(error) => return Err(error.into()),
    };
    let mut sha = Sha256::new();
    let mut buffer = vec![0u8; 64 * 1024];
    loop {
        let count = file.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        sha.update(&buffer[..count]);
    }
    Ok(Some(crate::hex(&sha.finalize())))
}

/// Checksum of a file that must exist.
pub fn file_checksum(path: &Path) -> Result<String> {
    checksum(path)?.ok_or_else(|| Error::invalid(format!("找不到文件：{}", path.display())))
}

/// Opens for reading while letting the editor keep writing, renaming or deleting the file.
pub fn open_shared(path: &Path) -> std::io::Result<File> {
    let mut options = OpenOptions::new();
    options.read(true);
    #[cfg(windows)]
    {
        use std::os::windows::fs::OpenOptionsExt;
        const FILE_SHARE_READ: u32 = 1;
        const FILE_SHARE_WRITE: u32 = 2;
        const FILE_SHARE_DELETE: u32 = 4;
        options.share_mode(FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE);
    }
    options.open(path)
}

/// Size and modification time, used to prove a file did not change between two passes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Stamp {
    pub length: u64,
    pub modified: Option<SystemTime>,
}

pub fn stamp(path: &Path) -> Result<Stamp> {
    let meta = fs::metadata(path)?;
    Ok(Stamp { length: meta.len(), modified: meta.modified().ok() })
}

/// An exclusive, non-blocking lock on `path` held until dropped.
pub struct FileLock {
    _file: File,
}

/// `Ok(None)` when another process already holds the lock.
pub fn try_lock(path: &Path) -> Result<Option<FileLock>> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    #[cfg(windows)]
    {
        use std::os::windows::fs::OpenOptionsExt;
        match OpenOptions::new().read(true).write(true).create(true).truncate(false).share_mode(0).open(path) {
            Ok(file) => Ok(Some(FileLock { _file: file })),
            Err(error) if matches!(error.raw_os_error(), Some(32) | Some(33)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }
    #[cfg(unix)]
    {
        use std::os::unix::io::AsRawFd;
        let file = OpenOptions::new().read(true).write(true).create(true).truncate(false).open(path)?;
        // SAFETY: flock on a descriptor we own; released when the file is closed.
        let locked = unsafe { libc::flock(file.as_raw_fd(), libc::LOCK_EX | libc::LOCK_NB) };
        if locked == 0 {
            Ok(Some(FileLock { _file: file }))
        } else {
            let error = std::io::Error::last_os_error();
            if error.raw_os_error() == Some(libc::EWOULDBLOCK) {
                Ok(None)
            } else {
                Err(error.into())
            }
        }
    }
}

/// Writes a small file through a temporary sibling and an atomic rename, flushing data first.
pub fn write_durable(path: &Path, bytes: &[u8]) -> Result<()> {
    let temporary = sibling(path, ".tmp");
    {
        let mut file = File::create(&temporary)?;
        file.write_all(bytes)?;
        file.sync_all()?;
    }
    let result = replace_file(&temporary, path);
    if result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    result?;
    sync_parent(path);
    Ok(())
}

/// Atomically moves `source` over `target` (which may or may not exist).
pub fn replace_file(source: &Path, target: &Path) -> std::io::Result<()> {
    // std::fs::rename replaces existing targets atomically on both Unix and Windows
    // (MoveFileExW with MOVEFILE_REPLACE_EXISTING).
    fs::rename(source, target)
}

/// Replaces `target` with `source` and keeps the previous `target` as `backup`, like .NET's
/// `File.Replace`. `target` is never missing: it is hard-linked (or copied) to the backup first.
pub fn replace_with_backup(source: &Path, target: &Path, backup: &Path) -> std::io::Result<()> {
    let _ = fs::remove_file(backup);
    if fs::hard_link(target, backup).is_err() {
        fs::copy(target, backup)?;
        File::open(backup)?.sync_all()?;
    }
    fs::rename(source, target)?;
    sync_parent(target);
    Ok(())
}

/// Flushes a directory entry change (rename, create) on Unix; a no-op on Windows.
pub fn sync_parent(path: &Path) {
    #[cfg(unix)]
    if let Some(parent) = path.parent() {
        if let Ok(directory) = File::open(parent) {
            let _ = directory.sync_all();
        }
    }
    #[cfg(not(unix))]
    let _ = path;
}

/// `path` with `suffix` appended to its file name.
pub fn sibling(path: &Path, suffix: &str) -> PathBuf {
    let mut name = path.file_name().map(|n| n.to_os_string()).unwrap_or_default();
    name.push(suffix);
    path.with_file_name(name)
}

/// Copies a file the editor may still have open.
pub fn copy_shared(source: &Path, destination: &Path) -> Result<()> {
    let mut input = open_shared(source)?;
    let mut output = OpenOptions::new().write(true).create_new(true).open(destination)?;
    std::io::copy(&mut input, &mut output)?;
    output.sync_all()?;
    Ok(())
}

pub fn try_remove(path: &Path) {
    let _ = fs::remove_file(path);
}

/// Case-insensitive path equality on Windows and macOS, exact elsewhere.
pub fn same_path(a: &Path, b: &Path) -> bool {
    if cfg!(any(windows, target_os = "macos")) {
        a.to_string_lossy().to_lowercase() == b.to_string_lossy().to_lowercase()
    } else {
        a == b
    }
}

/// Absolute, lexically normalised path (no filesystem access, symlinks untouched), like
/// .NET's `Path.GetFullPath`.
pub fn full_path(path: &Path) -> PathBuf {
    let absolute = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir().unwrap_or_default().join(path)
    };
    let mut result = PathBuf::new();
    for component in absolute.components() {
        match component {
            std::path::Component::CurDir => {}
            std::path::Component::ParentDir => {
                result.pop();
            }
            other => result.push(other.as_os_str()),
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lock_is_exclusive() {
        let path = std::env::temp_dir().join(format!("wg-lock-{}", crate::random_hex(4)));
        let first = try_lock(&path).unwrap();
        assert!(first.is_some());
        assert!(try_lock(&path).unwrap().is_none());
        drop(first);
        assert!(try_lock(&path).unwrap().is_some());
        let _ = fs::remove_file(&path);
    }

    #[test]
    fn replace_keeps_backup() {
        let folder = std::env::temp_dir().join(format!("wg-replace-{}", crate::random_hex(4)));
        fs::create_dir_all(&folder).unwrap();
        let target = folder.join("doc.docx");
        let source = folder.join("new.tmp");
        let backup = folder.join("doc.bak");
        fs::write(&target, b"old").unwrap();
        fs::write(&source, b"new").unwrap();
        replace_with_backup(&source, &target, &backup).unwrap();
        assert_eq!(fs::read(&target).unwrap(), b"new");
        assert_eq!(fs::read(&backup).unwrap(), b"old");
        assert!(!source.exists());
        let _ = fs::remove_dir_all(&folder);
    }

    #[test]
    fn full_path_normalises() {
        let base = if cfg!(windows) { PathBuf::from(r"C:\a\b") } else { PathBuf::from("/a/b") };
        assert_eq!(full_path(&base.join("..").join(".").join("c")), base.parent().unwrap().join("c"));
    }
}
