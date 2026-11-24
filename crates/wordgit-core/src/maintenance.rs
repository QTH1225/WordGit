//! Background storage maintenance: after enough new commits, a repository is repacked when the
//! service is idle. A full (cruft) repack runs at most once a month per repository.

use crate::history::Repo;
use crate::paths::AppPaths;
use std::collections::BTreeSet;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

pub const COMMIT_THRESHOLD: usize = 10;

/// Schedules repositories for repacking and runs them once the user has been idle for a while.
/// "Idle" means no API request for `idle` (the editor add-in calls the service whenever the user
/// works with WordGit).
pub struct Maintenance {
    pending: Mutex<BTreeSet<PathBuf>>,
    last_activity: AtomicU64,
    started: Instant,
    idle: Duration,
    stopped: AtomicBool,
    paths: AppPaths,
    executable: Option<PathBuf>,
}

impl Maintenance {
    pub fn new(paths: AppPaths, executable: Option<PathBuf>, idle: Duration) -> Arc<Self> {
        Arc::new(Maintenance {
            pending: Mutex::new(BTreeSet::new()),
            last_activity: AtomicU64::new(0),
            started: Instant::now(),
            idle,
            stopped: AtomicBool::new(false),
            paths,
            executable,
        })
    }

    pub fn schedule(&self, repository: &Path) {
        self.pending.lock().unwrap_or_else(|p| p.into_inner()).insert(repository.to_path_buf());
    }

    pub fn touch(&self) {
        self.last_activity.store(self.started.elapsed().as_millis() as u64, Ordering::SeqCst);
    }

    pub fn is_idle(&self) -> bool {
        let now = self.started.elapsed().as_millis() as u64;
        now.saturating_sub(self.last_activity.load(Ordering::SeqCst)) >= self.idle.as_millis() as u64
    }

    pub fn stop(&self) {
        self.stopped.store(true, Ordering::SeqCst);
    }

    /// Runs the maintenance loop on a background thread.
    pub fn spawn(self: &Arc<Self>) -> std::thread::JoinHandle<()> {
        let this = Arc::clone(self);
        std::thread::spawn(move || {
            // Pick up repositories that reached the threshold shortly before the last shutdown.
            if let Ok(entries) = std::fs::read_dir(this.paths.repositories()) {
                for entry in entries.flatten() {
                    let path = entry.path();
                    let stem = path.file_stem().map(|s| s.len()).unwrap_or(0);
                    if path.extension().and_then(|e| e.to_str()) == Some("git") && stem == 64 {
                        this.schedule(&path);
                    }
                }
            }
            while !this.stopped.load(Ordering::SeqCst) {
                std::thread::sleep(Duration::from_secs(5));
                this.run_pending();
            }
        })
    }

    fn take_next(&self) -> Option<PathBuf> {
        let mut pending = self.pending.lock().unwrap_or_else(|p| p.into_inner());
        let first = pending.iter().next().cloned()?;
        pending.remove(&first);
        Some(first)
    }

    fn run_pending(&self) {
        while self.is_idle() && !self.stopped.load(Ordering::SeqCst) {
            let Some(path) = self.take_next() else { return };
            let outcome = Repo::open(path.clone(), self.executable.as_deref(), None).and_then(|repo| {
                repo.try_auto_repack(COMMIT_THRESHOLD, &|| self.is_idle() && !self.stopped.load(Ordering::SeqCst))
            });
            match outcome {
                Ok(true) => {}
                Ok(false) => {
                    self.schedule(&path);
                    return;
                }
                Err(error) => self.paths.log_error(&format!("background repack check: {error}")),
            }
        }
    }
}
