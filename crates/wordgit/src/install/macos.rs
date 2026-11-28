use crate::service::{self, DEFAULT_PORT};
use serde::Deserialize;
use std::ffi::{CStr, OsStr, OsString};
use std::fs;
use std::io::ErrorKind;
use std::os::unix::ffi::{OsStrExt, OsStringExt};
use std::os::unix::fs::{MetadataExt, PermissionsExt};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};
use wordgit_core::fsutil::{self, FileLock};
use wordgit_core::paths::{home_dir, AppPaths};
use wordgit_core::{Error, Result};

pub const LAUNCH_AGENT_LABEL: &str = "com.qth1225.wordgit";
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

pub fn plist_path() -> PathBuf {
    home_dir().join("Library/LaunchAgents").join(format!("{LAUNCH_AGENT_LABEL}.plist"))
}

/// Only interact with launchd when operating on the real user's LaunchAgents folder.
/// Isolated tests override `HOME` and must not stop or replace the user's live agent.
pub fn should_use_launchctl() -> bool {
    if std::env::var_os("WORDGIT_NO_LAUNCHCTL").is_some_and(|v| !v.is_empty()) {
        return false;
    }
    let Some(real_home) = real_home_dir() else {
        return false;
    };
    fsutil::same_path(&home_dir(), &real_home)
}

fn real_home_dir() -> Option<PathBuf> {
    let uid = unsafe { libc::geteuid() };
    let passwd = unsafe { libc::getpwuid(uid) };
    if passwd.is_null() {
        return None;
    }
    let dir = unsafe { (*passwd).pw_dir };
    if dir.is_null() {
        return None;
    }
    let cstr = unsafe { CStr::from_ptr(dir) };
    Some(PathBuf::from(OsStr::from_bytes(cstr.to_bytes())))
}

fn gui_domain() -> String {
    format!("gui/{}", unsafe { libc::geteuid() })
}

pub fn unload_launch_agent(plist: &Path) {
    if !should_use_launchctl() {
        return;
    }
    let target = format!("{}/{LAUNCH_AGENT_LABEL}", gui_domain());
    let booted_out = Command::new("launchctl")
        .args(["bootout", &target])
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .is_ok_and(|status| status.success());
    if !booted_out && plist.is_file() {
        let _ = Command::new("launchctl")
            .args(["unload", "-w"])
            .arg(plist)
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .status();
    }
}

pub fn load_launch_agent(plist: &Path) -> bool {
    if !should_use_launchctl() {
        return false;
    }
    unload_launch_agent(plist);
    let domain = gui_domain();
    if Command::new("launchctl")
        .args(["bootstrap", &domain])
        .arg(plist)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .is_ok_and(|status| status.success())
    {
        return true;
    }
    Command::new("launchctl")
        .args(["load", "-w"])
        .arg(plist)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .is_ok_and(|status| status.success())
}

pub fn lock_installation(paths: &AppPaths) -> Result<FileLock> {
    fsutil::try_lock(&paths.root.join("install.lock"))?.ok_or_else(|| Error::invalid("另一个 WordGit 安装程序正在运行。"))
}

pub fn stop_service(paths: &AppPaths) -> Result<()> {
    // Unload LaunchAgent first so KeepAlive does not respawn the service while we stop it.
    unload_launch_agent(&plist_path());
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

fn process_uid(pid: i32) -> Option<u32> {
    let mut info: libc::proc_bsdinfo = unsafe { std::mem::zeroed() };
    let size = std::mem::size_of::<libc::proc_bsdinfo>() as libc::c_int;
    let got = unsafe {
        libc::proc_pidinfo(pid, libc::PROC_PIDTBSDINFO, 0, (&mut info as *mut libc::proc_bsdinfo).cast(), size)
    };
    if got == size { Some(info.pbi_uid) } else { None }
}

fn process_executable(pid: i32) -> Option<PathBuf> {
    let mut buffer = vec![0u8; libc::PROC_PIDPATHINFO_MAXSIZE as usize];
    let len = unsafe { libc::proc_pidpath(pid, buffer.as_mut_ptr().cast(), buffer.len() as u32) };
    if len <= 0 {
        return None;
    }
    buffer.truncate(len as usize);
    Some(PathBuf::from(OsString::from_vec(buffer)))
}

fn process_args(pid: i32) -> Option<Vec<Vec<u8>>> {
    let mut mib = [libc::CTL_KERN, libc::KERN_PROCARGS2, pid];
    let mut size: libc::size_t = 0;
    if unsafe { libc::sysctl(mib.as_mut_ptr(), 3, std::ptr::null_mut(), &mut size, std::ptr::null_mut(), 0) } != 0
        || size < 4
    {
        return None;
    }
    let mut buffer = vec![0u8; size];
    if unsafe { libc::sysctl(mib.as_mut_ptr(), 3, buffer.as_mut_ptr().cast(), &mut size, std::ptr::null_mut(), 0) } != 0
        || size < 4
    {
        return None;
    }
    buffer.truncate(size);
    let argc = i32::from_ne_bytes(buffer[..4].try_into().ok()?);
    if argc <= 0 {
        return None;
    }
    let mut index = 4;
    while index < buffer.len() && buffer[index] != 0 {
        index += 1;
    }
    while index < buffer.len() && buffer[index] == 0 {
        index += 1;
    }
    let mut args = Vec::with_capacity(argc as usize);
    for _ in 0..argc {
        if index >= buffer.len() {
            break;
        }
        let start = index;
        while index < buffer.len() && buffer[index] != 0 {
            index += 1;
        }
        args.push(buffer[start..index].to_vec());
        index += 1;
    }
    Some(args)
}

const PROC_PIDFDVNODEINFO: libc::c_int = 1;

#[repr(C)]
struct ProcFileInfo {
    fi_openflags: u32,
    fi_status: u32,
    fi_offset: libc::off_t,
    fi_type: i32,
    fi_guardflags: u32,
}

#[repr(C)]
struct VnodeFdInfo {
    pfi: ProcFileInfo,
    pvi: libc::vnode_info,
}

fn process_holds_file(pid: i32, lock: &Path) -> Result<bool> {
    let expected = fs::metadata(lock)?;
    let size = unsafe { libc::proc_pidinfo(pid, libc::PROC_PIDLISTFDS, 0, std::ptr::null_mut(), 0) };
    if size <= 0 {
        return Ok(false);
    }
    let fd_size = std::mem::size_of::<libc::proc_fdinfo>();
    let count = (size as usize) / fd_size + 16;
    let mut fds = vec![unsafe { std::mem::zeroed::<libc::proc_fdinfo>() }; count];
    let bytes = unsafe {
        libc::proc_pidinfo(
            pid,
            libc::PROC_PIDLISTFDS,
            0,
            fds.as_mut_ptr().cast(),
            (fds.len() * fd_size) as libc::c_int,
        )
    };
    if bytes <= 0 {
        return Ok(false);
    }
    let actual = (bytes as usize) / fd_size;
    let vnode_size = std::mem::size_of::<VnodeFdInfo>() as libc::c_int;
    for fd in &fds[..actual] {
        if fd.proc_fdtype != libc::PROX_FDTYPE_VNODE as u32 {
            continue;
        }
        let mut info: VnodeFdInfo = unsafe { std::mem::zeroed() };
        let got = unsafe {
            libc::proc_pidfdinfo(
                pid,
                fd.proc_fd,
                PROC_PIDFDVNODEINFO,
                (&mut info as *mut VnodeFdInfo).cast(),
                vnode_size,
            )
        };
        if got == vnode_size
            && u64::from(info.pvi.vi_stat.vst_dev) == expected.dev()
            && info.pvi.vi_stat.vst_ino == expected.ino()
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn signal_service(pid: i32, lock: &Path) -> Result<()> {
    let unknown = || Error::invalid("无法确认旧进程是此目录的 WordGit 服务。请退出该服务后重试。");
    if pid <= 1 {
        return Err(unknown());
    }
    // Check whether the process still exists before validating identity.
    if unsafe { libc::kill(pid, 0) } != 0 {
        let error = std::io::Error::last_os_error();
        if error.raw_os_error() == Some(libc::ESRCH) {
            return Ok(());
        }
        return Err(unknown());
    }
    let Some(uid) = process_uid(pid) else { return Err(unknown()) };
    let Some(executable) = process_executable(pid) else { return Err(unknown()) };
    let Some(args) = process_args(pid) else { return Err(unknown()) };
    let has_lock_file = process_holds_file(pid, lock)?;
    if uid != unsafe { libc::geteuid() }
        || executable.file_name().and_then(|name| name.to_str()) != Some("wordgit")
        || args.get(1).map(Vec::as_slice) != Some(b"serve")
        || !has_lock_file
    {
        return Err(unknown());
    }
    let sent = unsafe { libc::kill(pid, libc::SIGTERM) };
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
    // Copy to a new sibling file then rename so macOS AMFI sees a fresh vnode and code signature.
    let temporary = fsutil::sibling(&installed, &format!(".new-{}", wordgit_core::random_hex(8)));
    let result = fs::copy(source, &temporary)
        .and_then(|_| fs::set_permissions(&temporary, fs::Permissions::from_mode(0o755)))
        .and_then(|_| fs::rename(&temporary, &installed));
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

pub fn wait_for_launchd_service(paths: &AppPaths) -> Result<()> {
    let deadline = Instant::now() + WAIT;
    loop {
        if instance(paths).is_ok() && service::service_running(DEFAULT_PORT) {
            return Ok(());
        }
        if Instant::now() >= deadline {
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

pub fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
