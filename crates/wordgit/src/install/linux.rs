use crate::service::{self, DEFAULT_PORT};
use serde::Deserialize;
use std::fs;
use std::io::ErrorKind;
use std::os::fd::{AsRawFd, FromRawFd, OwnedFd};
use std::os::unix::fs::MetadataExt;
use std::path::{Path, PathBuf};
use std::process::Child;
use std::time::{Duration, Instant};
use wordgit_core::fsutil::{self, FileLock};
use wordgit_core::paths::AppPaths;
use wordgit_core::{Error, Result};

const WAIT: Duration = Duration::from_secs(12);
const POLL: Duration = Duration::from_millis(50);

#[derive(Deserialize)]
struct Instance {
    pid: i32,
}

fn instance(paths: &AppPaths) -> Result<Instance> {
    let file = fs::read(paths.root.join("service.json"))?;
    serde_json::from_slice(&file).map_err(|_| Error::invalid("无法确认原有 WordGit 服务。请退出该服务后重试。"))
}

pub fn lock_installation(paths: &AppPaths) -> Result<FileLock> {
    fsutil::try_lock(&paths.root.join("install.lock"))?.ok_or_else(|| Error::invalid("另一个 WordGit 安装程序正在运行。"))
}

pub fn stop_service(paths: &AppPaths) -> Result<()> {
    let lock_path = paths.root.join("service.lock");
    let deadline = Instant::now() + WAIT;
    let mut signalled = false;
    loop {
        if let Some(_lock) = fsutil::try_lock(&lock_path)? {
            fsutil::try_remove(&paths.root.join("service.json"));
            return Ok(());
        }
        if Instant::now() >= deadline {
            return Err(Error::Timeout("WordGit 服务尚未退出。请关闭 WPS，完成正在进行的操作后重试。".into()));
        }
        if !signalled {
            match instance(paths) {
                Ok(current) => {
                    signal_service(current.pid, &lock_path)?;
                    signalled = true;
                }
                Err(Error::Io(error)) if error.kind() == ErrorKind::NotFound => {}
                Err(error) => return Err(error),
            }
        }
        std::thread::sleep(POLL);
    }
}

fn signal_service(pid: i32, lock: &Path) -> Result<()> {
    let unknown = || Error::invalid("无法确认旧进程是此目录的 WordGit 服务。请退出该服务后重试。");
    if pid <= 1 {
        return Err(unknown());
    }
    // Pin the process before validating its identity so PID reuse cannot redirect the signal.
    // Older Linux kernels lack pidfds; the same ownership, executable and lock checks still apply.
    let fd = unsafe { libc::syscall(libc::SYS_pidfd_open, pid, 0u32) };
    let process = if fd >= 0 {
        Some(unsafe { OwnedFd::from_raw_fd(fd as i32) })
    } else {
        let error = std::io::Error::last_os_error();
        match error.raw_os_error() {
            Some(libc::ESRCH) => return Ok(()),
            Some(libc::ENOSYS) => None,
            _ => return Err(error.into()),
        }
    };
    let proc = PathBuf::from(format!("/proc/{pid}"));
    let executable = fs::read_link(proc.join("exe")).map_err(|_| unknown())?;
    let command = fs::read(proc.join("cmdline")).map_err(|_| unknown())?;
    let expected = fs::metadata(lock)?;
    let has_lock_file = fs::read_dir(proc.join("fd")).map_err(|_| unknown())?.flatten().any(|fd| {
        fd.path().metadata().is_ok_and(|meta| meta.dev() == expected.dev() && meta.ino() == expected.ino())
    });
    if proc.metadata()?.uid() != unsafe { libc::geteuid() } ||
        !matches!(executable.file_name().and_then(|name| name.to_str()), Some("wordgit" | "wordgit (deleted)")) ||
        command.split(|byte| *byte == 0).nth(1) != Some(b"serve") ||
        !has_lock_file
    {
        return Err(unknown());
    }
    let sent = if let Some(process) = process {
        unsafe { libc::syscall(libc::SYS_pidfd_send_signal, process.as_raw_fd(), libc::SIGTERM, std::ptr::null::<libc::siginfo_t>(), 0u32) }
    } else {
        unsafe { libc::kill(pid, libc::SIGTERM) as libc::c_long }
    };
    if sent != 0 {
        let error = std::io::Error::last_os_error();
        if error.raw_os_error() != Some(libc::ESRCH) {
            return Err(error.into());
        }
    }
    Ok(())
}

pub fn ensure_default_port() -> Result<()> {
    let address = std::net::SocketAddr::from(([127, 0, 0, 1], DEFAULT_PORT));
    if std::net::TcpStream::connect_timeout(&address, Duration::from_millis(500)).is_ok() {
        return Err(Error::invalid(format!("端口 {DEFAULT_PORT} 正在被其他程序使用。请释放该端口后重试。")));
    }
    Ok(())
}

pub fn install_executable(source: &Path, paths: &AppPaths) -> Result<PathBuf> {
    let app = paths.root.join("app");
    fs::create_dir_all(&app)?;
    let installed = fs::canonicalize(app)?.join("wordgit");
    if fs::canonicalize(source)? == installed {
        return Ok(installed);
    }
    let temporary = fsutil::sibling(&installed, &format!(".new-{}", wordgit_core::random_hex(8)));
    let result = fs::copy(source, &temporary).and_then(|_| fs::rename(&temporary, &installed));
    if result.is_err() {
        fsutil::try_remove(&temporary);
    }
    result?;
    Ok(installed)
}

pub fn wait_for_service(child: &mut Child, paths: &AppPaths) -> Result<()> {
    let deadline = Instant::now() + WAIT;
    loop {
        if child.try_wait()?.is_some() {
            return Err(Error::invalid("内核启动失败。请检查 Git 和本机服务端口后重试。"));
        }
        if instance(paths).is_ok_and(|running| running.pid as u32 == child.id()) && service::service_running(DEFAULT_PORT) {
            return Ok(());
        }
        if Instant::now() >= deadline {
            let _ = child.kill();
            let _ = child.wait();
            return Err(Error::Timeout("等待 WordGit 服务启动超时。".into()));
        }
        std::thread::sleep(POLL);
    }
}

pub fn uninstall_executable(paths: &AppPaths) -> Result<()> {
    let installed = paths.root.join("app/wordgit");
    match fs::remove_file(installed) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error.into()),
    }
}
