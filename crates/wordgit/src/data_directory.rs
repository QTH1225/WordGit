//! The service file and open review sessions stay at their original paths. The history
//! location is stored separately so both hosts can change it without losing their connection.

use serde::{Deserialize, Serialize};
use std::fs::{self, File, Metadata, OpenOptions, TryLockError};
use std::io::ErrorKind;
use std::path::{Path, PathBuf};
use wordgit_core::fsutil;
use wordgit_core::paths::{starts_with_folder, AppPaths};
use wordgit_core::{Error, Result};

#[derive(Deserialize, Serialize)]
struct Location {
    directory: PathBuf,
}

fn config(paths: &AppPaths) -> PathBuf {
    paths.root.join("data-directory.json")
}

pub fn load(paths: AppPaths) -> Result<AppPaths> {
    match fs::read(config(&paths)) {
        Ok(bytes) => {
            let location: Location = serde_json::from_slice(&bytes).map_err(|_| Error::data("WordGit 数据目录配置无法读取。"))?;
            if !location.directory.is_absolute() {
                return Err(Error::data("WordGit 数据目录配置必须是绝对路径。"));
            }
            Ok(paths.with_data_directory(location.directory))
        }
        Err(error) if error.kind() == ErrorKind::NotFound => Ok(paths),
        Err(error) => Err(error.into()),
    }
}

pub fn ensure_access(paths: &AppPaths) -> Result<()> {
    if paths.has_custom_data_directory() && !paths.repositories().is_dir() {
        return Err(Error::invalid(format!("数据目录不可用：{}。请重新连接该磁盘，再使用 WordGit。", paths.data_directory().display())));
    }
    Ok(())
}

/// Runs CLI history operations and migrations under one cross-process lock. HTTP calls
/// are additionally coordinated by the service, which also pauses background maintenance.
pub fn access_lock(paths: &AppPaths, exclusive: bool) -> Result<File> {
    fs::create_dir_all(&paths.root)?;
    let file = OpenOptions::new().read(true).write(true).create(true).truncate(false)
        .open(paths.root.join("data-directory.lock"))?;
    let result = if exclusive { file.try_lock() } else { file.try_lock_shared() };
    match result {
        Ok(()) => Ok(file),
        Err(TryLockError::WouldBlock) => Err(Error::invalid("WordGit 正在读写历史，请完成当前操作后再试。")),
        Err(TryLockError::Error(error)) => Err(error.into()),
    }
}

pub struct Migration {
    pub paths: AppPaths,
    pub message: String,
}

pub fn migrate(paths: &AppPaths, directory: &Path) -> Result<Migration> {
    let _access = access_lock(paths, true)?;
    ensure_access(paths)?;
    if !directory.is_absolute() {
        return Err(Error::invalid("请选择数据目录的绝对路径。"));
    }
    let directory = fsutil::full_path(directory);
    let current = paths.data_directory();
    let old = paths.repositories();
    let mut forbidden = vec![old.clone(), paths.sessions(), paths.root.join("app"), paths.root.join("logs")];
    validate_target(&directory, &forbidden)?;
    fs::create_dir_all(&directory)?;
    let resolved = fs::canonicalize(&directory)?;
    if fs::canonicalize(current).is_ok_and(|current| fsutil::same_path(&resolved, &current)) || fsutil::same_path(&directory, current) {
        return Ok(Migration { paths: paths.clone(), message: "当前已使用这个数据目录。".into() });
    }
    forbidden.extend(forbidden.clone().iter().filter_map(|path| fs::canonicalize(path).ok()));
    validate_target(&resolved, &forbidden)?;
    let destination = directory.join("repositories");
    match fs::symlink_metadata(&destination) {
        Ok(meta) if !is_link(&meta) && meta.is_dir() && fs::read_dir(&destination)?.next().is_none() => {
            fs::remove_dir(&destination)?;
        }
        Ok(_) => return Err(Error::invalid("目标目录已有 repositories。请选择其他目录，以免覆盖其中的历史。")),
        Err(error) if error.kind() == ErrorKind::NotFound => {}
        Err(error) => return Err(error.into()),
    }
    let stage = directory.join(format!(".wordgit-migration-{}", wordgit_core::random_hex(12)));
    let mut builder = fs::DirBuilder::new();
    #[cfg(unix)]
    {
        use std::os::unix::fs::DirBuilderExt;
        builder.mode(0o700);
    }
    #[cfg(not(unix))]
    {
        // Keep the Windows folder's inherited permissions.
        builder.recursive(false);
    }
    builder.create(&stage)?;
    let mut temporary = Staging(stage);
    match fs::symlink_metadata(&old) {
        Ok(meta) if !is_link(&meta) && meta.is_dir() => copy_verified(&old, &temporary.0)?,
        Ok(_) => return Err(Error::invalid("当前历史目录不是普通文件夹，无法自动迁移。")),
        Err(error) if error.kind() == ErrorKind::NotFound => {}
        Err(error) => return Err(error.into()),
    }
    // Keep the old store authoritative until the whole copy and the new location are durable.
    // If saving the location fails, the staging guard removes just our verified copy.
    let settings = serde_json::to_vec(&Location { directory: directory.clone() })
        .map_err(|_| Error::invalid("无法保存这个数据目录的路径。"))?;
    fs::rename(&temporary.0, &destination)?;
    temporary.0 = destination;
    fsutil::sync_parent(&temporary.0);
    fsutil::write_durable(&config(paths), &settings)?;
    temporary.0.clear();

    // The destination is now authoritative. An interrupted cleanup cannot affect its history.
    let mut message = "数据目录已更新，已有提交历史已迁移。".to_string();
    if old.is_dir() {
        let backup = fsutil::sibling(&old, &format!(".migrated-{}", wordgit_core::random_hex(8)));
        let remaining = match fs::rename(&old, &backup) {
            Ok(()) => fs::remove_dir_all(&backup).err().map(|_| backup),
            Err(_) => Some(old),
        };
        if let Some(remaining) = remaining {
            message.push_str(&format!("\n旧文件尚未完全清理：{}", remaining.display()));
        }
    }
    Ok(Migration { paths: paths.clone().with_data_directory(directory), message })
}

fn validate_target(directory: &Path, forbidden: &[PathBuf]) -> Result<()> {
    let target_repo = directory.join("repositories");
    if forbidden.iter().any(|path| {
        fsutil::same_path(directory, path) || starts_with_folder(directory, path) || starts_with_folder(path, &target_repo)
    }) {
        return Err(Error::invalid("请选择仓库、审阅稿、日志和程序目录之外的文件夹。"));
    }
    Ok(())
}

fn is_link(metadata: &Metadata) -> bool {
    #[cfg(windows)]
    {
        use std::os::windows::fs::MetadataExt;
        metadata.file_attributes() & 0x400 != 0 // FILE_ATTRIBUTE_REPARSE_POINT, including junctions
    }
    #[cfg(not(windows))]
    {
        metadata.file_type().is_symlink()
    }
}

fn copy_verified(from: &Path, to: &Path) -> Result<()> {
    for entry in fs::read_dir(from)? {
        let entry = entry?;
        let source = entry.path();
        let target = to.join(entry.file_name());
        let meta = fs::symlink_metadata(&source)?;
        if is_link(&meta) {
            return Err(Error::invalid(format!("历史中包含文件链接，无法自动迁移：{}", source.display())));
        }
        if meta.is_dir() {
            fs::create_dir(&target)?;
            copy_verified(&source, &target)?;
        } else if meta.is_file() {
            let stamp = fsutil::stamp(&source)?;
            fsutil::copy_shared(&source, &target)?;
            if fsutil::file_checksum(&source)? != fsutil::file_checksum(&target)? || fsutil::stamp(&source)? != stamp {
                return Err(Error::data(format!("校验历史文件失败：{}", source.display())));
            }
            #[cfg(unix)]
            fs::set_permissions(&target, meta.permissions())?;
        } else {
            return Err(Error::invalid(format!("历史中包含不支持的文件类型：{}", source.display())));
        }
    }
    fsutil::sync_parent(&to.join("directory-entry"));
    Ok(())
}

struct Staging(PathBuf);

impl Drop for Staging {
    fn drop(&mut self) {
        if !self.0.as_os_str().is_empty() {
            let _ = fs::remove_dir_all(&self.0);
        }
    }
}
