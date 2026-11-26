//! Data locations, preview/review sessions and the shareable error log.

use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::{Duration, SystemTime};

/// Root of WordGit's per-user data:
/// - `repositories/` one bare Git repository per document,
/// - `sessions/` previews, comparisons and pending merge drafts,
/// - `logs/wordgit.log` errors, with the home folder redacted.
#[derive(Clone, Debug)]
pub struct AppPaths {
    pub root: PathBuf,
}

impl AppPaths {
    pub fn new(root: impl Into<PathBuf>) -> Self {
        AppPaths { root: root.into() }
    }

    /// `WORDGIT_HOME` overrides the platform default (used by tests and portable setups).
    ///
    /// - Windows: `%LOCALAPPDATA%\WordGit`, shared with the original Word add-in.
    /// - Linux: `~/.local/share/WordGit`; deliberately not `$XDG_DATA_HOME`, because the WPS
    ///   add-in cannot read environment variables and must find the service file by itself.
    /// - macOS: inside the WPS sandbox container so WPS can open session files directly.
    pub fn from_environment() -> Self {
        if let Some(home) = std::env::var_os("WORDGIT_HOME").filter(|v| !v.is_empty()) {
            return AppPaths::new(home);
        }
        AppPaths::new(default_root())
    }

    pub fn repositories(&self) -> PathBuf {
        self.root.join("repositories")
    }

    pub fn sessions(&self) -> PathBuf {
        self.root.join("sessions")
    }

    pub fn log_file(&self) -> PathBuf {
        self.root.join("logs").join("wordgit.log")
    }

    /// Fresh folder for one preview/compare/merge session.
    pub fn new_session(&self) -> std::io::Result<PathBuf> {
        let folder = self.sessions().join(crate::random_hex(16));
        fs::create_dir_all(&folder)?;
        Ok(folder)
    }

    pub fn is_session_file(&self, path: &Path) -> bool {
        starts_with_folder(path, &self.sessions())
    }

    /// Removes session folders older than `days` that hold no pending merge draft (`*.wordgit`).
    pub fn clean_sessions(&self, days: u64) -> usize {
        let cutoff = SystemTime::now() - Duration::from_secs(days * 86_400);
        let mut removed = 0;
        let Ok(entries) = fs::read_dir(self.sessions()) else { return 0 };
        for entry in entries.flatten() {
            let folder = entry.path();
            if !folder.is_dir() {
                continue;
            }
            let fresh = fs::metadata(&folder).and_then(|m| m.modified()).map(|t| t >= cutoff).unwrap_or(true);
            if fresh || contains_extension(&folder, "wordgit") {
                continue;
            }
            if fs::remove_dir_all(&folder).is_ok() {
                removed += 1;
            }
        }
        removed
    }

    /// Appends one line to the error log, rolling it over at 1 MiB.
    pub fn log_error(&self, message: &str) {
        append_log(&self.log_file(), message);
    }
}

fn default_root() -> PathBuf {
    #[cfg(windows)]
    {
        let base = std::env::var_os("LOCALAPPDATA")
            .map(PathBuf::from)
            .unwrap_or_else(|| home_dir().join("AppData").join("Local"));
        base.join("WordGit")
    }
    #[cfg(target_os = "macos")]
    {
        home_dir().join("Library/Containers/com.kingsoft.wpsoffice.mac/Data/WordGit")
    }
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        home_dir().join(".local/share/WordGit")
    }
}

pub fn home_dir() -> PathBuf {
    #[cfg(windows)]
    let home = std::env::var_os("USERPROFILE");
    #[cfg(not(windows))]
    let home = std::env::var_os("HOME");
    home.map(PathBuf::from).unwrap_or_else(|| PathBuf::from("."))
}

fn contains_extension(folder: &Path, extension: &str) -> bool {
    let Ok(entries) = fs::read_dir(folder) else { return true };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            if contains_extension(&path, extension) {
                return true;
            }
        } else if path.extension().and_then(|e| e.to_str()).is_some_and(|e| e.eq_ignore_ascii_case(extension)) {
            return true;
        }
    }
    false
}

/// True when `path` lies inside `folder` (case-insensitive on Windows and macOS).
pub fn starts_with_folder(path: &Path, folder: &Path) -> bool {
    let path = path.to_string_lossy();
    let mut folder = folder.to_string_lossy().into_owned();
    if !folder.ends_with(std::path::MAIN_SEPARATOR) {
        folder.push(std::path::MAIN_SEPARATOR);
    }
    if cfg!(any(windows, target_os = "macos")) {
        path.to_lowercase().starts_with(&folder.to_lowercase())
    } else {
        path.starts_with(&folder)
    }
}

/// Logs are meant to be attached to bug reports, so the user's home folder is replaced.
pub fn redact(message: &str) -> String {
    let home = home_dir().to_string_lossy().into_owned();
    if home.len() < 2 {
        return message.to_string();
    }
    let placeholder = if cfg!(windows) { "%USERPROFILE%" } else { "~" };
    if cfg!(windows) {
        let lower = message.to_lowercase();
        let needle = home.to_lowercase();
        // Lowercasing can change byte lengths for some scripts; only redact when it does not.
        if lower.len() != message.len() {
            return message.replace(&home, placeholder);
        }
        let mut result = String::with_capacity(message.len());
        let mut index = 0;
        while let Some(found) = lower[index..].find(&needle) {
            result.push_str(&message[index..index + found]);
            result.push_str(placeholder);
            index += found + needle.len();
        }
        result.push_str(&message[index..]);
        result
    } else {
        message.replace(&home, placeholder)
    }
}

pub fn append_log(file: &Path, message: &str) {
    let _ = (|| -> std::io::Result<()> {
        if let Some(parent) = file.parent() {
            fs::create_dir_all(parent)?;
        }
        if fs::metadata(file).map(|m| m.len() > 1024 * 1024).unwrap_or(false) {
            let rolled = file.with_extension("log.1");
            let _ = fs::remove_file(&rolled);
            fs::rename(file, &rolled)?;
        }
        let mut handle = fs::OpenOptions::new().create(true).append(true).open(file)?;
        let stamp = SystemTime::now().duration_since(SystemTime::UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0);
        writeln!(handle, "[{}] {}", format_utc(stamp), redact(message))
    })();
}

/// `YYYY-MM-DD HH:MM:SS` (UTC) without a date library.
pub fn format_utc(seconds: u64) -> String {
    let (year, month, day) = civil_from_days((seconds / 86_400) as i64);
    let rest = seconds % 86_400;
    format!("{year:04}-{month:02}-{day:02} {:02}:{:02}:{:02}", rest / 3600, rest % 3600 / 60, rest % 60)
}

/// `YYYYMMDDHHMMSS` (UTC), used for backup ref namespaces.
pub fn compact_utc(seconds: u64) -> String {
    let (year, month, day) = civil_from_days((seconds / 86_400) as i64);
    let rest = seconds % 86_400;
    format!("{year:04}{month:02}{day:02}{:02}{:02}{:02}", rest / 3600, rest % 3600 / 60, rest % 60)
}

pub fn now_seconds() -> u64 {
    SystemTime::now().duration_since(SystemTime::UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0)
}

// Howard Hinnant's days-from-civil inverse.
fn civil_from_days(days: i64) -> (i64, u32, u32) {
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let month = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    let year = yoe + era * 400 + i64::from(month <= 2);
    (year, month, day)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn utc_formatting() {
        assert_eq!(format_utc(0), "1970-01-01 00:00:00");
        assert_eq!(compact_utc(1_700_000_000), "20231114221320");
    }

    #[test]
    fn session_detection() {
        let paths = AppPaths::new(std::env::temp_dir().join("wg-paths"));
        let inside = paths.sessions().join("abc").join("x.docx");
        assert!(paths.is_session_file(&inside));
        assert!(!paths.is_session_file(&paths.root.join("sessions-other").join("x.docx")));
    }
}
