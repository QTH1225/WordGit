//! Every Git invocation: argument handling, environment isolation, durability settings,
//! timeouts and process-tree cleanup. Repository semantics live in [`crate::history`].

use crate::{Error, Result};
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, ExitStatus, Stdio};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

pub const DEFAULT_TIMEOUT: Duration = Duration::from_secs(45);
pub const NETWORK_IDLE_TIMEOUT: Duration = Duration::from_secs(180);

/// `core.fsync=committed` flushes loose objects, packs and refs before Git reports success, so a
/// power loss right after a commit cannot leave a branch pointing at an empty object.
const LOCAL_OPTIONS: [&str; 6] =
    ["-c", "core.longpaths=true", "-c", "core.fsync=committed", "-c", "core.fsyncMethod=batch"];

#[cfg(windows)]
pub const NULL_DEVICE: &str = "NUL";
#[cfg(not(windows))]
pub const NULL_DEVICE: &str = "/dev/null";

pub struct GitResult {
    pub code: i32,
    pub stdout: Vec<u8>,
    pub stderr: String,
}

impl GitResult {
    /// Fails with Git's own message unless the command succeeded.
    pub fn checked(self) -> Result<GitResult> {
        if self.code != 0 {
            return Err(Error::invalid(format!("Git 操作失败：{}", self.stderr.trim())));
        }
        Ok(self)
    }

    /// Strict UTF-8 output, trimmed like the original add-in (spaces and line breaks).
    pub fn text(self) -> Result<String> {
        let result = self.checked()?;
        let text = String::from_utf8(result.stdout).map_err(|_| Error::data("Git 输出不是有效的 UTF-8。"))?;
        Ok(trim_output(&text).to_string())
    }

    /// Lenient UTF-8: commit messages written by other clients are not guaranteed to be UTF-8.
    pub fn lenient_text(self) -> Result<String> {
        let result = self.checked()?;
        Ok(trim_output(&String::from_utf8_lossy(&result.stdout)).to_string())
    }
}

pub fn trim_output(text: &str) -> &str {
    text.trim_matches(|c| c == ' ' || c == '\r' || c == '\n')
}

/// Writes the process's standard input. Errors that are not plain I/O errors are the caller's
/// own failures and are always reported, even when Git also failed.
pub type InputWriter<'a> = &'a mut (dyn FnMut(&mut dyn Write) -> Result<()> + Send);

#[derive(Clone, Debug)]
pub struct Git {
    executable: PathBuf,
    directory: PathBuf,
}

impl Git {
    pub fn new(executable: impl Into<PathBuf>, directory: impl Into<PathBuf>) -> Self {
        Git { executable: executable.into(), directory: directory.into() }
    }

    pub fn executable(&self) -> &Path {
        &self.executable
    }

    /// `WORDGIT_GIT` (absolute path) wins; then `PATH`; then the usual Git for Windows folders.
    pub fn find() -> Result<PathBuf> {
        if let Some(explicit) = std::env::var_os("WORDGIT_GIT").filter(|v| !v.is_empty()) {
            let path = PathBuf::from(explicit);
            if !path.is_absolute() || !path.is_file() {
                return Err(Error::invalid("WORDGIT_GIT 必须指向有效的 git 可执行文件绝对路径。"));
            }
            return Ok(path);
        }
        let name = if cfg!(windows) { "git.exe" } else { "git" };
        let mut folders: Vec<PathBuf> =
            std::env::var_os("PATH").map(|p| std::env::split_paths(&p).collect()).unwrap_or_default();
        #[cfg(windows)]
        for (variable, suffix) in [
            ("ProgramFiles", "Git\\cmd"),
            ("ProgramFiles(x86)", "Git\\cmd"),
            ("LOCALAPPDATA", "Programs\\Git\\cmd"),
        ] {
            if let Some(base) = std::env::var_os(variable) {
                folders.push(PathBuf::from(base).join(suffix));
            }
        }
        #[cfg(target_os = "macos")]
        folders.extend(["/opt/homebrew/bin", "/usr/local/bin", "/usr/bin"].map(PathBuf::from));
        #[cfg(all(not(windows), not(target_os = "macos")))]
        folders.extend(["/usr/bin", "/usr/local/bin", "/opt/homebrew/bin"].map(PathBuf::from));
        for folder in folders {
            let folder = PathBuf::from(folder.to_string_lossy().trim().trim_matches('"'));
            let candidate = folder.join(name);
            if folder.is_absolute() && candidate.is_file() {
                return Ok(candidate);
            }
        }
        Err(Error::invalid(if cfg!(windows) {
            "找不到 Git for Windows。请安装 Git，或用 WORDGIT_GIT 指定 git.exe。"
        } else if cfg!(target_os = "macos") {
            "找不到 git。请先安装 Xcode 命令行工具（xcode-select --install）或通过 Homebrew 安装 Git，或用 WORDGIT_GIT 指定 git 的路径。"
        } else {
            "找不到 git。请先安装 Git（例如 sudo apt install git），或用 WORDGIT_GIT 指定 git 的路径。"
        }))
    }

    /// A Git command against this repository with the user's configuration shut out.
    pub fn command<S: AsRef<std::ffi::OsStr>>(&self, identity: Option<&str>, args: &[S]) -> Command {
        let mut command = Command::new(&self.executable);
        command.args(LOCAL_OPTIONS).arg("--git-dir=.").args(args).current_dir(&self.directory);
        for (key, _) in std::env::vars_os() {
            if key.to_string_lossy().to_ascii_uppercase().starts_with("GIT_") {
                command.env_remove(key);
            }
        }
        command
            .env("GIT_CONFIG_NOSYSTEM", "1")
            .env("GIT_CONFIG_GLOBAL", NULL_DEVICE)
            .env("GIT_TERMINAL_PROMPT", "0");
        if let Some(identity) = identity {
            command
                .env("GIT_AUTHOR_NAME", identity)
                .env("GIT_COMMITTER_NAME", identity)
                .env("GIT_AUTHOR_EMAIL", "wordgit@local.invalid")
                .env("GIT_COMMITTER_EMAIL", "wordgit@local.invalid");
        }
        command
    }

    pub fn invoke<S: AsRef<std::ffi::OsStr>>(&self, input: Option<&[u8]>, identity: Option<&str>, args: &[S]) -> Result<GitResult> {
        match input {
            Some(bytes) => {
                let mut writer = |stream: &mut dyn Write| -> Result<()> { Ok(stream.write_all(bytes)?) };
                self.invoke_with(Some(&mut writer), identity, Some(DEFAULT_TIMEOUT), false, args)
            }
            None => self.invoke_with(None, identity, Some(DEFAULT_TIMEOUT), false, args),
        }
    }

    pub fn invoke_with<S: AsRef<std::ffi::OsStr>>(
        &self,
        input: Option<InputWriter<'_>>,
        identity: Option<&str>,
        timeout: Option<Duration>,
        low_priority: bool,
        args: &[S],
    ) -> Result<GitResult> {
        run(self.command(identity, args), input, timeout, low_priority)
    }

    /// Reads the user's own global and system configuration (e.g. their SSH command).
    pub fn invoke_with_user_config<S: AsRef<std::ffi::OsStr>>(&self, args: &[S]) -> Result<GitResult> {
        let mut command = self.command(None, args);
        command.env_remove("GIT_CONFIG_NOSYSTEM").env_remove("GIT_CONFIG_GLOBAL");
        run(command, None, Some(DEFAULT_TIMEOUT), false)
    }

    /// Like `invoke`, but standard output goes to `reader` instead of memory. If the reader fails,
    /// Git is stopped and the reader's error (more informative than Git's exit code) is returned.
    pub fn stream<S: AsRef<std::ffi::OsStr>>(
        &self,
        input: Option<&[u8]>,
        timeout: Option<Duration>,
        reader: &mut (dyn FnMut(&mut dyn Read) -> Result<()> + Send),
        args: &[S],
    ) -> Result<()> {
        let mut command = self.command(None, args);
        command.stdin(if input.is_some() { Stdio::piped() } else { Stdio::null() });
        command.stdout(Stdio::piped()).stderr(Stdio::piped());
        let mut tree = ProcessTree::spawn(command)?;
        let stdin = tree.child.stdin.take();
        let mut stdout = tree.child.stdout.take().expect("piped stdout");
        let stderr = tree.child.stderr.take().expect("piped stderr");
        let deadline = timeout.map(|t| Instant::now() + t);
        std::thread::scope(|scope| -> Result<()> {
            let errors = scope.spawn(move || read_lossy(stderr));
            let writer = input.map(|bytes| {
                scope.spawn(move || -> io::Result<()> {
                    let mut stdin = stdin.expect("piped stdin");
                    stdin.write_all(bytes)
                })
            });
            let reading = scope.spawn(move || reader(&mut stdout));
            // Wait for the reader or the process, whichever ends first.
            let mut status: Option<ExitStatus> = None;
            loop {
                if reading.is_finished() {
                    break;
                }
                if status.is_none() {
                    status = tree.child.try_wait()?;
                }
                if status.is_some() && reading.is_finished() {
                    break;
                }
                if deadline.is_some_and(|d| Instant::now() > d) {
                    tree.kill();
                    let _ = reading.join();
                    return Err(Error::Timeout("Git 操作超时。".into()));
                }
                std::thread::sleep(Duration::from_millis(5));
            }
            let read_result = reading.join().unwrap_or_else(|_| Err(Error::data("读取 Git 输出时出错。")));
            if let Err(error) = read_result {
                tree.kill();
                return Err(error);
            }
            let status = match status {
                Some(status) => status,
                None => wait_until(&mut tree, deadline)?,
            };
            let stderr = errors.join().unwrap_or_default();
            if !status.success() {
                return Err(Error::invalid(format!("Git 操作失败：{}", stderr.trim())));
            }
            if let Some(writer) = writer {
                writer.join().unwrap_or(Ok(()))?;
            }
            Ok(())
        })
    }

    /// Network operations have no total deadline, only a deadline for silence: a first clone over
    /// a slow link can legitimately take far longer than any fixed limit. Progress output counts as
    /// activity. Cancellation and timeouts stop the whole tree, including remote helpers and ssh.
    pub fn run_network(&self, mut command: Command, cancel: &AtomicBool, idle_timeout: Duration) -> Result<()> {
        if cancel.load(Ordering::SeqCst) {
            return Err(Error::Cancelled);
        }
        command.stdin(Stdio::null()).stdout(Stdio::piped()).stderr(Stdio::piped());
        let mut tree = ProcessTree::spawn(command)?;
        let stdout = tree.child.stdout.take().expect("piped stdout");
        let stderr = tree.child.stderr.take().expect("piped stderr");
        let started = Instant::now();
        let last_activity = Arc::new(AtomicU64::new(0));
        let detail = Arc::new(Mutex::new(NetworkDetail::default()));
        std::thread::scope(|scope| -> Result<()> {
            let activity = Arc::clone(&last_activity);
            let collected = Arc::clone(&detail);
            let errors = scope.spawn(move || {
                pump(stderr, |chunk| {
                    activity.store(started.elapsed().as_millis() as u64, Ordering::SeqCst);
                    collected.lock().unwrap_or_else(|p| p.into_inner()).push(chunk);
                })
            });
            let activity = Arc::clone(&last_activity);
            let output = scope.spawn(move || {
                pump(stdout, |_| activity.store(started.elapsed().as_millis() as u64, Ordering::SeqCst))
            });
            let status = loop {
                if let Some(status) = tree.child.try_wait()? {
                    break status;
                }
                if cancel.load(Ordering::SeqCst) {
                    tree.kill();
                    return Err(Error::Cancelled);
                }
                let idle = started.elapsed().as_millis() as u64 - last_activity.load(Ordering::SeqCst);
                if idle > idle_timeout.as_millis() as u64 {
                    tree.kill();
                    return Err(Error::Timeout(
                        "远程 Git 操作长时间没有进展，已停止。请检查网络，或完成 Git 登录后重试。".into(),
                    ));
                }
                std::thread::sleep(Duration::from_millis(250));
            };
            let _ = output.join();
            let _ = errors.join();
            if cancel.load(Ordering::SeqCst) {
                return Err(Error::Cancelled);
            }
            if !status.success() {
                let mut message = detail.lock().unwrap_or_else(|p| p.into_inner()).message();
                if message.chars().count() > 1600 {
                    message = message.chars().rev().take(1600).collect::<Vec<_>>().into_iter().rev().collect();
                }
                return Err(Error::invalid(format!(
                    "远程 Git 操作未完成。推送被拒绝时请先获取并拉取；认证失败时请完成 Git 登录或凭据配置。\n{message}"
                )));
            }
            Ok(())
        })
    }
}

fn run(mut command: Command, input: Option<InputWriter<'_>>, timeout: Option<Duration>, low_priority: bool) -> Result<GitResult> {
    command.stdin(if input.is_some() { Stdio::piped() } else { Stdio::null() });
    command.stdout(Stdio::piped()).stderr(Stdio::piped());
    let mut tree = ProcessTree::spawn(command)?;
    if low_priority {
        tree.lower_priority();
    }
    let stdin = tree.child.stdin.take();
    let mut stdout = tree.child.stdout.take().expect("piped stdout");
    let stderr = tree.child.stderr.take().expect("piped stderr");
    let deadline = timeout.map(|t| Instant::now() + t);
    std::thread::scope(|scope| -> Result<GitResult> {
        let output = scope.spawn(move || {
            let mut bytes = Vec::new();
            stdout.read_to_end(&mut bytes).map(|_| bytes)
        });
        let errors = scope.spawn(move || read_lossy(stderr));
        let writer = input.map(|write| {
            scope.spawn(move || -> Result<()> {
                let mut stdin = stdin.expect("piped stdin");
                let result = write(&mut stdin);
                drop(stdin);
                result
            })
        });
        let status = wait_until(&mut tree, deadline)?;
        let stdout = output.join().unwrap_or_else(|_| Ok(Vec::new()))?;
        let stderr = errors.join().unwrap_or_default();
        let code = status.code().unwrap_or(-1);
        if let Some(writer) = writer {
            match writer.join().unwrap_or(Ok(())) {
                Ok(()) => {}
                // A broken pipe is just the downstream effect of Git exiting early.
                Err(Error::Io(_)) if code != 0 => {}
                Err(error) => return Err(error),
            }
        }
        Ok(GitResult { code, stdout, stderr })
    })
}

fn wait_until(tree: &mut ProcessTree, deadline: Option<Instant>) -> Result<ExitStatus> {
    loop {
        if let Some(status) = tree.child.try_wait()? {
            return Ok(status);
        }
        if deadline.is_some_and(|d| Instant::now() > d) {
            tree.kill();
            return Err(Error::Timeout("Git 操作超时。".into()));
        }
        std::thread::sleep(Duration::from_millis(2));
    }
}

fn read_lossy(mut stream: impl Read) -> String {
    let mut bytes = Vec::new();
    let _ = stream.read_to_end(&mut bytes);
    String::from_utf8_lossy(&bytes).into_owned()
}

fn pump(mut stream: impl Read, mut chunk: impl FnMut(&str)) {
    let mut buffer = [0u8; 4096];
    while let Ok(count) = stream.read(&mut buffer) {
        if count == 0 {
            break;
        }
        chunk(&String::from_utf8_lossy(&buffer[..count]));
    }
}

/// Collects a network command's stderr, dropping progress lines so the error message keeps only
/// what explains a failure.
#[derive(Default)]
struct NetworkDetail {
    detail: String,
    pending: String,
}

impl NetworkDetail {
    fn push(&mut self, chunk: &str) {
        self.pending.push_str(chunk);
        let Some(cut) = self.pending.rfind(['\n', '\r']) else { return };
        let complete: String = self.pending[..cut].to_string();
        self.pending = self.pending[cut + 1..].to_string();
        for line in complete.split(['\r', '\n']) {
            let line = line.trim();
            if !line.is_empty() && !is_progress_line(line) {
                self.detail.push_str(line);
                self.detail.push('\n');
            }
        }
        if self.detail.len() > 8000 {
            let mut start = self.detail.len() - 4000;
            while !self.detail.is_char_boundary(start) {
                start += 1;
            }
            self.detail = self.detail[start..].to_string();
        }
    }

    fn message(&self) -> String {
        format!("{}{}", self.detail, self.pending).trim().to_string()
    }
}

pub fn is_progress_line(line: &str) -> bool {
    let line = line.strip_prefix("remote: ").unwrap_or(line);
    const WORDS: [&str; 11] = [
        "Enumerating", "Counting", "Compressing", "Receiving", "Resolving", "Writing", "Total",
        "Delta compression", "Unpacking", "Checking", "Updating",
    ];
    WORDS.iter().any(|word| {
        line.strip_prefix(word).is_some_and(|rest| rest.chars().next().is_none_or(|c| !c.is_alphanumeric() && c != '_'))
    })
}

/// A started Git process plus a container for every helper it spawns (a process group on Unix, a
/// job object on Windows). Only an explicit kill stops the tree; a normal exit leaves
/// intentionally detached helpers (such as a credential cache daemon) alone.
struct ProcessTree {
    child: Child,
    #[cfg(windows)]
    job: windows_sys::Win32::Foundation::HANDLE,
}

impl ProcessTree {
    fn spawn(mut command: Command) -> Result<ProcessTree> {
        #[cfg(unix)]
        {
            use std::os::unix::process::CommandExt;
            command.process_group(0);
        }
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            const CREATE_NO_WINDOW: u32 = 0x0800_0000;
            command.creation_flags(CREATE_NO_WINDOW);
        }
        let child = command.spawn().map_err(|error| {
            if error.kind() == io::ErrorKind::NotFound {
                Error::invalid("无法启动 git，请确认已安装 Git。")
            } else {
                Error::Io(error)
            }
        })?;
        #[cfg(windows)]
        {
            use std::os::windows::io::AsRawHandle;
            use windows_sys::Win32::Foundation::CloseHandle;
            use windows_sys::Win32::System::JobObjects::{AssignProcessToJobObject, CreateJobObjectW};
            // SAFETY: plain Win32 calls on a handle owned by `child` and a job handle we own.
            let mut job = unsafe { CreateJobObjectW(std::ptr::null(), std::ptr::null()) };
            if !job.is_null() && unsafe { AssignProcessToJobObject(job, child.as_raw_handle() as _) } == 0 {
                unsafe { CloseHandle(job) };
                job = std::ptr::null_mut();
            }
            Ok(ProcessTree { child, job })
        }
        #[cfg(not(windows))]
        Ok(ProcessTree { child })
    }

    fn lower_priority(&self) {
        #[cfg(unix)]
        // SAFETY: setpriority on our own child's pid.
        unsafe {
            libc::setpriority(libc::PRIO_PROCESS, self.child.id() as libc::id_t, 10);
        }
        #[cfg(windows)]
        {
            use std::os::windows::io::AsRawHandle;
            use windows_sys::Win32::System::Threading::{SetPriorityClass, BELOW_NORMAL_PRIORITY_CLASS};
            // SAFETY: the handle belongs to the running child.
            unsafe { SetPriorityClass(self.child.as_raw_handle() as _, BELOW_NORMAL_PRIORITY_CLASS) };
        }
    }

    fn kill(&mut self) {
        #[cfg(unix)]
        // SAFETY: the child leads its own process group (process_group(0)).
        unsafe {
            libc::kill(-(self.child.id() as i32), libc::SIGKILL);
        }
        #[cfg(windows)]
        if !self.job.is_null() {
            // SAFETY: job handle owned by this tree.
            unsafe { windows_sys::Win32::System::JobObjects::TerminateJobObject(self.job, 1) };
        }
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}

impl Drop for ProcessTree {
    fn drop(&mut self) {
        // A caller that bailed out early must not leave Git (or its helpers) running.
        if let Ok(None) = self.child.try_wait() {
            self.kill();
        }
        #[cfg(windows)]
        if !self.job.is_null() {
            // SAFETY: job handle owned by this tree, closed exactly once.
            unsafe { windows_sys::Win32::Foundation::CloseHandle(self.job) };
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn progress_lines_are_recognised() {
        assert!(is_progress_line("remote: Counting objects: 100% (3/3), done."));
        assert!(is_progress_line("Receiving objects:  50%"));
        assert!(!is_progress_line("fatal: Authentication failed"));
        assert!(!is_progress_line("Totally different"));
    }

    #[test]
    fn network_detail_keeps_errors_only() {
        let mut detail = NetworkDetail::default();
        detail.push("Counting objects: 1\rfatal: could not read Username\nerror: fai");
        detail.push("led to push\n");
        assert_eq!(detail.message(), "fatal: could not read Username\nerror: failed to push");
    }
}
