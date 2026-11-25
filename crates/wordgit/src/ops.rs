//! WordGit operations behind the local service and the command line. Every function here is
//! blocking and independent of HTTP; the editor add-in drives the editor itself (open, compare,
//! close, reopen) and asks these operations for everything that touches history or files.

use serde::Deserialize;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::atomic::AtomicBool;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};
use wordgit_core::fsutil::{self, FileLock};
use wordgit_core::graph;
use wordgit_core::history::{self, PreparedSnapshot, Repo};
use wordgit_core::maintenance::Maintenance;
use wordgit_core::paths::AppPaths;
use wordgit_core::remote::{self, REMOTE_PREFIX};
use wordgit_core::review::{self, PendingReview};
use wordgit_core::{docx, journal, Error, Result};

const PLAN_LIFETIME: Duration = Duration::from_secs(600);

pub struct Context {
    pub paths: AppPaths,
    pub git: Option<PathBuf>,
    pub maintenance: Arc<Maintenance>,
    plans: Mutex<HashMap<String, ReplacePlan>>,
    cancels: Mutex<HashMap<String, Arc<AtomicBool>>>,
    notices: Mutex<Vec<String>>,
    /// Allows plain local directories as remotes (tests only).
    pub allow_local_remotes: bool,
}

/// What a working-copy replacement publishes once the file is swapped.
enum Publish {
    Switch { from: String, tip: String, to: String, target_tip: String },
    Commit {
        prepared: PreparedSnapshot,
        message: String,
        author: String,
        head: Option<String>,
        branch: String,
        source_branch: Option<String>,
        source_tip: Option<String>,
        /// Merge draft whose metadata is updated after a finished merge.
        draft: Option<PathBuf>,
    },
    FastForward { branch: String, head: String, reference: String, remote_tip: String },
}

/// Prepared in the background while the document is still open; applied after the editor closed it.
struct ReplacePlan {
    repo: Repo,
    document: PathBuf,
    current_hash: String,
    staged: PathBuf,
    replacement_hash: String,
    publish: Publish,
    done_message: String,
    created: Instant,
    /// Held from preparation until the swap is applied or abandoned.
    _lock: FileLock,
}

impl Context {
    pub fn new(paths: AppPaths, git: Option<PathBuf>, maintenance: Arc<Maintenance>) -> Arc<Context> {
        Arc::new(Context {
            paths,
            git,
            maintenance,
            plans: Mutex::new(HashMap::new()),
            cancels: Mutex::new(HashMap::new()),
            notices: Mutex::new(Vec::new()),
            allow_local_remotes: std::env::var_os("WORDGIT_ALLOW_LOCAL_REMOTES").is_some(),
        })
    }

    pub fn add_notices(&self, messages: Vec<String>) {
        self.notices.lock().unwrap_or_else(|p| p.into_inner()).extend(messages);
    }

    fn repo_for_history(&self, history_path: &Path) -> Result<Repo> {
        let maintenance = Arc::clone(&self.maintenance);
        let hook: history::CommitHook = Arc::new(move |path: &Path| maintenance.schedule(path));
        Repo::for_document(history_path, &self.paths.repositories(), self.git.as_deref(), Some(hook))
    }

    fn cancel_flag(&self, operation: Option<&str>) -> Arc<AtomicBool> {
        let flag = Arc::new(AtomicBool::new(false));
        if let Some(id) = operation {
            self.cancels.lock().unwrap_or_else(|p| p.into_inner()).insert(id.to_string(), Arc::clone(&flag));
        }
        flag
    }

    fn release_cancel(&self, operation: Option<&str>) {
        if let Some(id) = operation {
            self.cancels.lock().unwrap_or_else(|p| p.into_inner()).remove(id);
        }
    }

    fn store_plan(&self, plan: ReplacePlan) -> String {
        let id = wordgit_core::random_hex(12);
        let mut plans = self.plans.lock().unwrap_or_else(|p| p.into_inner());
        plans.retain(|_, plan| plan.created.elapsed() < PLAN_LIFETIME);
        plans.insert(id.clone(), plan);
        id
    }
}

// ---- request types -------------------------------------------------------------------------

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentRequest {
    pub document: String,
    #[serde(default)]
    pub branch: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CommitRequest {
    pub document: String,
    pub message: String,
    #[serde(default)]
    pub author: String,
    /// The head the user saw; the commit is refused if history moved meanwhile.
    #[serde(default)]
    pub expected_head: Option<String>,
    #[serde(default)]
    pub expected_branch: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RevisionRequest {
    pub document: String,
    pub revision: String,
    #[serde(default)]
    pub other: Option<String>,
    #[serde(default)]
    pub current_in_memory: bool,
    #[serde(default)]
    pub author: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BranchRequest {
    pub document: String,
    pub name: String,
    #[serde(default)]
    pub revision: Option<String>,
    /// Commit the working copy first with this message (switching with uncommitted changes).
    #[serde(default)]
    pub message: Option<String>,
    #[serde(default)]
    pub author: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportRequest {
    pub document: String,
    pub file: String,
    pub basis: String,
    #[serde(default)]
    pub branch: Option<String>,
    #[serde(default)]
    pub reviewer: Option<String>,
    #[serde(default)]
    pub archive_only: bool,
    #[serde(default)]
    pub expected_head: Option<String>,
    #[serde(default)]
    pub expected_branch: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeBranchRequest {
    pub document: String,
    pub source: String,
}

/// Like `MergeBranchRequest`, but the caller can name the submitted version a return draft was
/// based on; the add-in knows it from the import, the kernel cannot always recover it.
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeContextRequest {
    pub document: String,
    pub source: String,
    pub basis: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RegisterRequest {
    pub draft: String,
    pub review: ReviewInput,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ReviewInput {
    pub original: String,
    pub head: String,
    pub branch: String,
    #[serde(default)]
    pub source_branch: Option<String>,
    #[serde(default)]
    pub source_tip: Option<String>,
    #[serde(default)]
    pub apply_to_original: bool,
    #[serde(default)]
    pub record_only_merge: bool,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FinishMergeRequest {
    pub draft: String,
    pub message: String,
    #[serde(default)]
    pub author: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PlanRequest {
    pub plan: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RemoteRequest {
    pub document: String,
    #[serde(default)]
    pub url: Option<String>,
    #[serde(default)]
    pub branch: Option<String>,
    #[serde(default)]
    pub operation: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CloneRequest {
    pub url: String,
    pub branch: String,
    pub document: String,
    #[serde(default)]
    pub operation: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CancelRequest {
    pub operation: String,
}

// ---- helpers -------------------------------------------------------------------------------

fn short(id: &str) -> &str {
    &id[..8.min(id.len())]
}

fn file_stem(path: &Path) -> String {
    path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_else(|| "文档".into())
}

/// A saved local `.docx` the service can read.
fn document_path(document: &str) -> Result<PathBuf> {
    if document.trim().is_empty() {
        return Err(Error::invalid("请先将这份文档保存为本地 .docx。"));
    }
    if document.contains("://") {
        return Err(Error::invalid("WordGit 需要可直接读取的本地 .docx。云文档请先另存一份到本地文件夹。"));
    }
    let path = fsutil::full_path(Path::new(document));
    let is_docx = path.extension().and_then(|e| e.to_str()).is_some_and(|e| e.eq_ignore_ascii_case("docx"));
    if !is_docx {
        return Err(Error::invalid("WordGit 只支持 .docx。请先另存为 .docx 格式。"));
    }
    if !path.is_file() {
        return Err(Error::invalid(format!("找不到文档：{}。请先保存。", path.display())));
    }
    Ok(path)
}

fn revisions_json(repo: &Repo, revisions: &[history::Revision], head: Option<&str>) -> Value {
    let _ = repo;
    json!({ "history": revisions, "graph": graph::layout(revisions, head) })
}

struct Opened {
    path: PathBuf,
    review: Option<PendingReview>,
    history_path: PathBuf,
    repo: Repo,
}

impl Context {
    fn open(&self, document: &str) -> Result<Opened> {
        let path = document_path(document)?;
        let review = PendingReview::load(&self.paths, &path)?;
        let history_path = review::history_path(&self.paths, &path)?;
        let repo = self.repo_for_history(&history_path)?;
        Ok(Opened { path, review, history_path, repo })
    }

    /// The original document of an action that must run on the original (not a merge draft).
    fn open_original(&self, document: &str, action: &str) -> Result<Opened> {
        let opened = self.open(document)?;
        if opened.review.is_some() {
            return Err(Error::invalid(format!("请在原文档中{action}。")));
        }
        Ok(opened)
    }
}

/// Fails unless the working copy equals the branch tip (merges start from committed content).
fn check_merge_target(repo: &Repo, path: &Path, tip: &str) -> Result<()> {
    review::ensure_replaceable(path)?;
    if fsutil::file_checksum(path)? != repo.snapshot_checksum(tip)? {
        return Err(Error::invalid("请先提交当前分支的修改，再生成合并稿。"));
    }
    Ok(())
}

// ---- operations ----------------------------------------------------------------------------

impl Context {
    pub fn ping(&self) -> Value {
        json!({
            "ok": true,
            "service": "wordgit",
            "version": env!("CARGO_PKG_VERSION"),
            "api": 1,
            "os": std::env::consts::OS,
            "home": self.paths.root,
            "sessions": self.paths.sessions(),
        })
    }

    pub fn take_notices(&self) -> Value {
        let notices = std::mem::take(&mut *self.notices.lock().unwrap_or_else(|p| p.into_inner()));
        json!({ "notices": notices })
    }

    /// Everything the history pane needs for one document.
    pub fn status(&self, request: &DocumentRequest) -> Result<Value> {
        let path = document_path(&request.document)?;
        let review = PendingReview::load(&self.paths, &path)?;
        if review.is_none() && self.paths.is_session_file(&path) {
            return Ok(json!({ "document": path, "session": true }));
        }
        let opened = self.open(&request.document)?;
        let view = opened.repo.read_view()?;
        let head = view.remote.local_tip.clone();
        let working_state = match &head {
            None => "noCommit",
            Some(tip) if review.is_none() => {
                if fsutil::file_checksum(&path)? == opened.repo.snapshot_checksum(tip)? {
                    "clean"
                } else {
                    "changed"
                }
            }
            Some(_) => "review",
        };
        let graph = graph::layout(&view.history, head.as_deref());
        Ok(json!({
            "document": path,
            "session": false,
            "historyPath": opened.history_path,
            "repository": opened.repo.path(),
            "review": review,
            "branch": view.remote.local_branch,
            "head": head,
            "branches": view.branches,
            "remote": view.remote,
            "history": view.history,
            "graph": graph,
            "workingState": working_state,
        }))
    }

    pub fn history(&self, request: &DocumentRequest) -> Result<Value> {
        let opened = self.open(&request.document)?;
        let revisions = opened.repo.history(request.branch.as_deref())?;
        let head = opened.repo.head()?;
        Ok(revisions_json(&opened.repo, &revisions, head.as_deref()))
    }

    pub fn commit(&self, request: &CommitRequest) -> Result<Value> {
        let opened = self.open(&request.document)?;
        let repo = &opened.repo;
        let branch = repo.active_branch()?;
        let head = repo.head()?;
        if request.expected_branch.as_ref().is_some_and(|expected| expected != &branch) {
            return Err(Error::invalid("分支已变化，请刷新后再操作。"));
        }
        if let Some(expected) = &request.expected_head {
            if head.as_deref() != Some(expected.as_str()) {
                return Err(Error::invalid("分支或版本历史已变化。请刷新后再操作。"));
            }
        }
        if let Some(review) = &opened.review {
            if head.as_deref() != Some(review.head.as_str()) || branch != review.branch {
                return Err(Error::invalid("审阅期间原文档的历史已更新。请先比较最新版本，再重新合并。"));
            }
            if review.source_branch.is_some() || review.apply_to_original {
                return Err(Error::invalid("这是待完成的合并稿，请点击「完成合并」。"));
            }
        }
        let id = {
            let _lock = repo.lock()?;
            repo.commit_file(&opened.path, &request.message, &request.author, head.as_deref(), Some(&branch), None, None)?
        };
        if let Some(mut review) = opened.review {
            if id.is_some() || review.is_pending_merge() {
                review.record_only_merge = false;
                review.save(&opened.path, id.as_deref().or(head.as_deref()).unwrap_or_default())?;
            }
        }
        let message = match &id {
            Some(id) => format!("已提交 {}。提交时保存的版本已入库，之后的编辑可继续提交。", short(id)),
            None => "文件内容与最新提交一致。".into(),
        };
        Ok(json!({ "commit": id, "message": message, "branch": branch }))
    }

    pub fn create_branch(&self, request: &BranchRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "新建分支")?;
        let revision = match &request.revision {
            Some(revision) => revision.clone(),
            None => opened.repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?,
        };
        let _lock = opened.repo.lock()?;
        opened.repo.create_branch(request.name.trim(), &revision)?;
        Ok(json!({ "message": format!("已创建分支「{}」，当前仍在「{}」。", request.name.trim(), opened.repo.active_branch()?) }))
    }

    pub fn deletable_branches(&self, request: &DocumentRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "删除分支")?;
        let branch = opened.repo.active_branch()?;
        let head = opened.repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?;
        let branches: Vec<String> =
            opened.repo.branches(Some(&head))?.into_iter().filter(|b| !b.eq_ignore_ascii_case(&branch)).collect();
        Ok(json!({ "branches": branches, "current": branch }))
    }

    pub fn delete_branch(&self, request: &BranchRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "删除分支")?;
        let repo = &opened.repo;
        let current = repo.active_branch()?;
        let head = repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?;
        let tip = repo.branch_tip(&request.name)?.ok_or_else(|| Error::invalid("找不到要删除的分支。"))?;
        let _lock = repo.lock()?;
        repo.delete_merged_branch(&current, &head, &request.name, &tip)?;
        Ok(json!({ "message": format!("已删除分支「{}」。", request.name) }))
    }

    /// Restores a revision to a fresh session file for read-only preview.
    pub fn preview(&self, request: &RevisionRequest) -> Result<Value> {
        let opened = self.open(&request.document)?;
        let folder = self.paths.new_session()?;
        let file = folder.join(format!("{}-{}.docx", file_stem(&opened.history_path), short(&request.revision)));
        history::extract_to_file(&opened.repo, &request.revision, &file)?;
        Ok(json!({ "file": file }))
    }

    /// Saves a revision as a new document next to the original, never overwriting a file.
    pub fn export(&self, request: &RevisionRequest) -> Result<Value> {
        let opened = self.open(&request.document)?;
        let folder = opened.history_path.parent().map(Path::to_path_buf).unwrap_or_default();
        let stem = format!("{}-{}", file_stem(&opened.history_path), short(&request.revision));
        let mut file = folder.join(format!("{stem}.docx"));
        let mut index = 2;
        while file.exists() {
            file = folder.join(format!("{stem}-{index}.docx"));
            index += 1;
        }
        history::extract_to_file(&opened.repo, &request.revision, &file)?;
        Ok(json!({ "file": file, "message": format!("已另存为：{}", file.display()) }))
    }

    /// Files for comparing a revision with the current draft, or two revisions.
    pub fn compare_files(&self, request: &RevisionRequest) -> Result<Value> {
        let opened = self.open(&request.document)?;
        let folder = self.paths.new_session()?;
        let left = folder.join("基准版本.docx");
        history::extract_to_file(&opened.repo, &request.revision, &left)?;
        let right = match &request.other {
            Some(other) => {
                let file = folder.join("修订版本.docx");
                history::extract_to_file(&opened.repo, other, &file)?;
                file
            }
            None if request.current_in_memory => folder.join("当前稿.xml"),
            None => {
                let file = folder.join("当前稿.docx");
                fsutil::copy_shared(&opened.path, &file)?;
                file
            }
        };
        Ok(json!({ "session": folder, "original": left, "revised": right }))
    }

    /// Prepares switching the working copy to another branch.
    pub fn switch_prepare(&self, request: &BranchRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "切换分支")?;
        let repo = opened.repo.clone();
        review::ensure_replaceable(&opened.path)?;
        let from = repo.active_branch()?;
        let expected = repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?;
        let target = request.name.clone();
        let target_tip = repo.branch_tip(&target)?.ok_or_else(|| Error::invalid("找不到目标分支。"))?;
        let lock = repo.lock()?;
        let tip = match &request.message {
            Some(message) => repo
                .commit_file(&opened.path, message, &request.author, Some(&expected), Some(&from), None, None)?
                .unwrap_or(expected),
            None => {
                if fsutil::file_checksum(&opened.path)? != repo.snapshot_checksum(&expected)? {
                    return Err(Error::invalid("当前稿有未提交的修改，请先提交再切换。"));
                }
                expected
            }
        };
        let staged = self.paths.new_session()?.join("切换目标.docx");
        history::extract_to_file(&repo, &target_tip, &staged)?;
        let current_hash = repo.snapshot_checksum(&tip)?;
        let replacement_hash = repo.snapshot_checksum(&target_tip)?;
        let changes_file = current_hash != replacement_hash;
        let plan = self.store_plan(ReplacePlan {
            repo,
            document: opened.path.clone(),
            current_hash,
            staged,
            replacement_hash,
            publish: Publish::Switch { from, tip, to: target.clone(), target_tip },
            done_message: format!("已切换到「{target}」。"),
            created: Instant::now(),
            _lock: lock,
        });
        Ok(json!({ "plan": plan, "document": opened.path, "changesFile": changes_file }))
    }

    /// Prepares restoring an old version: the current draft is committed first.
    pub fn restore_prepare(&self, request: &RevisionRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "恢复旧版")?;
        let repo = opened.repo.clone();
        review::ensure_replaceable(&opened.path)?;
        let lock = repo.lock()?;
        let current_hash = fsutil::file_checksum(&opened.path)?;
        let old_hash = repo.snapshot_checksum(&request.revision)?;
        if current_hash == old_hash {
            return Ok(json!({ "unchanged": true, "message": "当前稿已经与所选旧版一致。" }));
        }
        let branch = repo.active_branch()?;
        let head = repo.head()?;
        let head = repo
            .commit_file(&opened.path, "恢复前保存当前稿", &request.author, head.as_deref(), Some(&branch), None, None)?
            .or(head);
        let prepared = repo.prepared_from_revision(&request.revision)?;
        let staged = self.paths.new_session()?.join("恢复版本.docx");
        history::extract_to_file(&repo, &request.revision, &staged)?;
        let plan = self.store_plan(ReplacePlan {
            repo,
            document: opened.path.clone(),
            current_hash,
            staged,
            replacement_hash: old_hash,
            publish: Publish::Commit {
                prepared,
                message: format!("恢复版本 {}", short(&request.revision)),
                author: request.author.clone(),
                head,
                branch,
                source_branch: None,
                source_tip: None,
                draft: None,
            },
            done_message: "已恢复旧版。".into(),
            created: Instant::now(),
            _lock: lock,
        });
        Ok(json!({ "plan": plan, "document": opened.path, "changesFile": true }))
    }

    /// Prepares writing a finished merge draft back to the original and committing it.
    pub fn finish_merge_prepare(&self, request: &FinishMergeRequest) -> Result<Value> {
        let draft = document_path(&request.draft)?;
        let review = PendingReview::load(&self.paths, &draft)?.ok_or_else(|| Error::invalid("这不是待完成的合并稿。"))?;
        if !(review.source_branch.is_some() || review.apply_to_original) {
            return Err(Error::invalid("这份合并稿直接提交即可。"));
        }
        let repo = self.repo_for_history(&review.original)?;
        if repo.active_branch()? != review.branch || repo.head()?.as_deref() != Some(review.head.as_str()) {
            return Err(Error::invalid("审阅期间原文档的历史已更新。请先比较最新版本，再重新合并。"));
        }
        if let Some(source) = &review.source_branch {
            if repo.branch_tip(source)? != review.source_tip {
                return Err(Error::invalid("来源分支已更新，请重新生成合并修订稿。"));
            }
        }
        review::ensure_replaceable(&review.original)?;
        let lock = repo.lock()?;
        let prepared = repo.prepare_snapshot_file(&draft)?;
        let head_hash = repo.snapshot_checksum(&review.head)?;
        let staged = self.paths.new_session()?.join("合并结果.docx");
        fsutil::copy_shared(&draft, &staged)?;
        if fsutil::file_checksum(&staged)? != prepared.checksum {
            return Err(Error::invalid("合并稿在准备期间被修改，请保存后重试。"));
        }
        let summary = match &review.source_branch {
            Some(source) => format!("已合并「{}」到「{}」", remote::reference_label(source), review.branch),
            None => format!("已将返稿合并到「{}」", review.branch),
        };
        let changes_file = head_hash != prepared.checksum;
        let plan = self.store_plan(ReplacePlan {
            repo,
            document: review.original.clone(),
            current_hash: head_hash,
            staged,
            replacement_hash: prepared.checksum.clone(),
            publish: Publish::Commit {
                prepared,
                message: request.message.clone(),
                author: request.author.clone(),
                head: Some(review.head.clone()),
                branch: review.branch.clone(),
                source_branch: review.source_branch.clone(),
                source_tip: review.source_tip.clone(),
                draft: Some(draft),
            },
            done_message: summary,
            created: Instant::now(),
            _lock: lock,
        });
        Ok(json!({ "plan": plan, "document": review.original, "changesFile": changes_file }))
    }

    /// Swaps the working copy and publishes. The editor must have closed the document.
    pub fn apply_plan(&self, request: &PlanRequest) -> Result<Value> {
        let plan = self
            .plans
            .lock()
            .unwrap_or_else(|p| p.into_inner())
            .remove(&request.plan)
            .ok_or_else(|| Error::invalid("操作已过期，请重新开始。"))?;
        let mut commit: Option<String> = None;
        let repo = plan.repo.clone();
        let mut publish = || -> Result<()> {
            match &plan.publish {
                Publish::Switch { from, tip, to, target_tip } => repo.switch(from, tip, to, target_tip),
                Publish::Commit { prepared, message, author, head, branch, source_branch, source_tip, .. } => {
                    commit = repo.commit_prepared(
                        prepared,
                        message,
                        author,
                        head.as_deref(),
                        Some(branch),
                        source_branch.as_deref(),
                        source_tip.as_deref(),
                    )?;
                    Ok(())
                }
                Publish::FastForward { branch, head, reference, remote_tip } => {
                    repo.fast_forward_remote(branch, head, reference, remote_tip)
                }
            }
        };
        let result = review::replace_working_copy(
            &plan.repo,
            &plan.document,
            &plan.current_hash,
            &plan.staged,
            &plan.replacement_hash,
            &mut publish,
        );
        fsutil::try_remove(&plan.staged);
        result?;
        let mut message = plan.done_message.clone();
        if let Publish::Commit { draft: Some(draft), head, .. } = &plan.publish {
            if let Some(mut review) = PendingReview::load(&self.paths, draft)? {
                review.source_branch = None;
                review.source_tip = None;
                review.apply_to_original = false;
                review.save(draft, commit.as_deref().or(head.as_deref()).unwrap_or_default())?;
            }
            message = match &commit {
                Some(id) => format!("{message}，提交 {}。", short(id)),
                None => "合并结果与最新提交一致。".into(),
            };
        } else if let (Publish::Commit { .. }, Some(id)) = (&plan.publish, &commit) {
            message = format!("已恢复并创建新提交 {}。", short(id));
        }
        Ok(json!({ "message": message, "commit": commit, "document": plan.document }))
    }

    pub fn cancel_plan(&self, request: &PlanRequest) -> Value {
        let removed = self.plans.lock().unwrap_or_else(|p| p.into_inner()).remove(&request.plan);
        if let Some(plan) = removed {
            fsutil::try_remove(&plan.staged);
        }
        json!({ "ok": true })
    }

    /// Files and options for merging a branch (or the remote branch) into the current one. The
    /// add-in builds the draft from them with the editor's compare function.
    pub fn merge_branch_plan(&self, request: &MergeBranchRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "合并分支")?;
        let repo = &opened.repo;
        let branch = repo.active_branch()?;
        let tip = repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?;
        check_merge_target(repo, &opened.path, &tip)?;
        let source_tip = repo.branch_tip(&request.source)?.ok_or_else(|| Error::invalid("找不到来源分支。"))?;
        if repo.is_ancestor(&source_tip, &tip)? {
            return Ok(json!({ "upToDate": true, "message": "来源分支的提交已包含在当前分支。" }));
        }
        let imported = repo.imported_review_base(&source_tip)?;
        let basis = match &imported {
            Some(basis) => basis.clone(),
            None => repo.merge_base(&tip, &source_tip)?,
        };
        let label = remote::reference_label(&request.source);
        let incoming_author = if imported.is_some() { repo.revision_author(&source_tip)? } else { label.clone() };
        let preserve = imported.is_some() || request.source.starts_with(REMOTE_PREFIX);
        let folder = self.paths.new_session()?;
        let files = self.three_way_files(repo, &folder, &basis, &tip, &source_tip)?;
        let draft = folder.join(format!("{}-分支合并稿.docx", file_stem(&opened.path)));
        Ok(json!({
            "draft": draft,
            "basis": files.0, "target": files.1, "incoming": files.2,
            "basisIsTarget": files.3,
            "incomingAuthor": incoming_author,
            "preserveIncoming": preserve,
            "review": {
                "original": opened.path, "head": tip, "branch": branch,
                "sourceBranch": request.source, "sourceTip": source_tip,
            },
            "notice": format!("已打开「{label} → {branch}」的修订稿。请接受或拒绝修订，再点击「完成合并」。"),
        }))
    }

    /// Everything the Word add-in needs to build a branch-merge draft.
    ///
    /// Same three-way files as `merge_branch_plan`, plus the return-draft history of the incoming
    /// branch: which version the return was based on and which earlier rounds of that same review
    /// already went into this branch. The add-in merges a repeated return on top of the last round
    /// instead of re-merging the whole return, which would duplicate text.
    pub fn merge_context(&self, request: &MergeContextRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "合并分支")?;
        let repo = &opened.repo;
        let branch = repo.active_branch()?;
        let tip = repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?;
        check_merge_target(repo, &opened.path, &tip)?;
        let source_tip = repo.branch_tip(&request.source)?.ok_or_else(|| Error::invalid("找不到来源分支。"))?;
        if repo.is_ancestor(&source_tip, &tip)? {
            return Ok(json!({ "upToDate": true, "message": "来源分支的提交已包含在当前分支。" }));
        }
        let imported = repo.imported_review_base(&source_tip)?;
        let submitted = request.basis.clone().or_else(|| imported.clone());
        let basis = match &submitted {
            Some(base) => base.clone(),
            None => repo.merge_base(&tip, &source_tip)?,
        };
        let label = remote::reference_label(&request.source);
        let preserve = imported.is_some() || request.source.starts_with(REMOTE_PREFIX);
        let folder = self.paths.new_session()?;
        let files = self.three_way_files(repo, &folder, &basis, &tip, &source_tip)?;
        let mut previous = Vec::new();
        if submitted.is_some() {
            for (index, revision) in repo.previously_merged_returns(&tip, &basis)?.into_iter().enumerate() {
                let file = folder.join(format!("历史返稿-{}.docx", index + 1));
                history::extract_to_file(repo, &revision, &file)?;
                previous.push(json!({ "id": revision, "file": file }));
            }
        }
        let draft = folder.join(format!("{}-分支合并稿.docx", file_stem(&opened.path)));
        Ok(json!({
            "draft": draft,
            "basis": files.0, "target": files.1, "incoming": files.2,
            "basisIsTarget": files.3,
            "incomingAuthor": label,
            "preserveIncoming": preserve,
            "submittedBase": submitted,
            "previousReturns": previous,
            "targetBranch": branch, "sourceBranch": request.source,
            "review": {
                "original": opened.path, "head": tip, "branch": branch,
                "sourceBranch": request.source, "sourceTip": source_tip,
            },
            "notice": format!("已打开「{label} → {branch}」的修订稿。请接受或拒绝修订，再点击「完成合并」。"),
        }))
    }

    fn three_way_files(&self, repo: &Repo, folder: &Path, basis: &str, target: &str, incoming: &str) -> Result<(PathBuf, PathBuf, PathBuf, bool)> {
        let basis_file = folder.join("共同基准.docx");
        let target_file = folder.join("当前分支.docx");
        let incoming_file = folder.join("来源分支.docx");
        history::extract_to_file(repo, basis, &basis_file)?;
        history::extract_to_file(repo, target, &target_file)?;
        history::extract_to_file(repo, incoming, &incoming_file)?;
        let same = repo.snapshot_checksum(basis)? == repo.snapshot_checksum(target)?;
        Ok((basis_file, target_file, incoming_file, same))
    }

    /// Preserves the original return as a commit, optionally exposes it as a local branch, and
    /// prepares a merge for clients that request the combined operation.
    pub fn import_review(&self, request: &ImportRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "处理返稿")?;
        let repo = &opened.repo;
        let file = fsutil::full_path(Path::new(request.file.trim()));
        let is_docx = file.extension().and_then(|e| e.to_str()).is_some_and(|e| e.eq_ignore_ascii_case("docx"));
        if !file.is_file() || !is_docx {
            return Err(Error::invalid("请选择已经保存的 .docx 返稿。"));
        }
        if fsutil::same_path(&file, &opened.path) {
            return Err(Error::invalid("请选择收到的返稿文件。"));
        }
        docx::validate_file(&file)?;
        let prepare_merge = request.branch.is_none() && !request.archive_only;
        let imported = {
            let _lock = repo.lock()?;
            let branch = repo.active_branch()?;
            let tip = repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?;
            if request.expected_head.as_ref().is_some_and(|expected| expected != &tip)
                || request.expected_branch.as_ref().is_some_and(|expected| expected != &branch)
            {
                return Err(Error::invalid("当前分支或提交已变化，请刷新后重新导入。"));
            }
            if prepare_merge {
                check_merge_target(repo, &opened.path, &tip)?;
            }
            repo.import_review_file(&file, &request.basis, request.branch.as_deref().map(str::trim),
                request.reviewer.as_deref().unwrap_or("返稿"), &branch, &tip)?
        };
        let mut result = if prepare_merge {
            self.merge_branch_plan(&MergeBranchRequest { document: request.document.clone(), source: imported.source.clone() })?
        } else {
            json!({})
        };
        let message = if result["upToDate"] == true {
            "这份返稿已经合入当前分支。".to_string()
        } else {
            let label = if imported.reused { "已复用原始返稿" } else { "已归档原始返稿" };
            match &imported.branch {
                Some(name) => format!("{label} {}，分支「{name}」。", short(&imported.commit)),
                None => format!("{label} {}。", short(&imported.commit)),
            }
        };
        result["message"] = json!(message);
        result["commit"] = json!(imported.commit);
        result["source"] = json!(imported.source);
        result["branch"] = json!(imported.branch);
        result["reused"] = json!(imported.reused);
        Ok(result)
    }

    /// Files for merging two historical versions into a draft that is then committed normally.
    pub fn merge_versions_plan(&self, request: &RevisionRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "合并两个版本")?;
        let other = request.other.clone().ok_or_else(|| Error::invalid("请选择两个版本。"))?;
        let branch = opened.repo.active_branch()?;
        let tip = opened.repo.head()?.ok_or_else(|| Error::invalid("还没有版本，请先提交。"))?;
        let folder = self.paths.new_session()?;
        let original = folder.join("基准版本.docx");
        let revised = folder.join("修订版本.docx");
        history::extract_to_file(&opened.repo, &request.revision, &original)?;
        history::extract_to_file(&opened.repo, &other, &revised)?;
        Ok(json!({
            "draft": folder.join(format!("{}-合并稿.docx", file_stem(&opened.path))),
            "original": original, "revised": revised,
            "review": { "original": opened.path, "head": tip, "branch": branch, "recordOnlyMerge": true },
            "notice": "已打开合并稿。请接受或拒绝修订，再点击「完成合并」。",
        }))
    }

    /// Records the merge draft the add-in saved, tying it to the original history.
    pub fn register_review(&self, request: &RegisterRequest) -> Result<Value> {
        let draft = document_path(&request.draft)?;
        if !self.paths.is_session_file(&draft) {
            return Err(Error::invalid("合并稿必须保存在 WordGit 的会话目录中。"));
        }
        let input = &request.review;
        let review = PendingReview {
            original: fsutil::full_path(Path::new(&input.original)),
            head: wordgit_core::object_id(&input.head)?,
            branch: input.branch.clone(),
            source_branch: input.source_branch.clone(),
            source_tip: input.source_tip.clone(),
            apply_to_original: input.apply_to_original,
            record_only_merge: input.record_only_merge,
        };
        review.save(&draft, &review.head)?;
        Ok(json!({ "ok": true }))
    }

    // ---- remote ------------------------------------------------------------------------------

    fn with_cancel<T>(&self, operation: Option<&str>, run: impl FnOnce(&AtomicBool) -> Result<T>) -> Result<T> {
        let flag = self.cancel_flag(operation);
        let result = run(&flag);
        self.release_cancel(operation);
        result
    }

    pub fn cancel(&self, request: &CancelRequest) -> Value {
        if let Some(flag) = self.cancels.lock().unwrap_or_else(|p| p.into_inner()).get(&request.operation) {
            flag.store(true, std::sync::atomic::Ordering::SeqCst);
        }
        json!({ "ok": true })
    }

    pub fn remote_link(&self, request: &RemoteRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "关联远程仓库")?;
        let url = request.url.clone().ok_or_else(|| Error::invalid("请输入仓库地址。"))?;
        let branch = request.branch.clone().unwrap_or_else(|| "main".into());
        let state = self.with_cancel(request.operation.as_deref(), |cancel| {
            let _lock = opened.repo.lock()?;
            opened.repo.configure_remote(&url, &branch, self.allow_local_remotes)?;
            opened.repo.fetch_remote(self.allow_local_remotes, cancel)
        })?;
        Ok(json!({ "remote": state, "message": format!("仓库已关联。获取不会覆盖当前稿。\n{}", state.caption) }))
    }

    pub fn remote_fetch(&self, request: &RemoteRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "获取")?;
        let state = self.with_cancel(request.operation.as_deref(), |cancel| {
            let _lock = opened.repo.lock()?;
            opened.repo.fetch_remote(self.allow_local_remotes, cancel)
        })?;
        Ok(json!({ "remote": state, "message": format!("已获取远程历史，当前稿未变。\n{}", state.caption) }))
    }

    pub fn remote_push(&self, request: &RemoteRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "推送")?;
        let state = self.with_cancel(request.operation.as_deref(), |cancel| {
            let _lock = opened.repo.lock()?;
            opened.repo.push_remote(self.allow_local_remotes, cancel)
        })?;
        Ok(json!({ "remote": state, "message": format!("已推送当前分支的已提交版本。\n{}", state.caption) }))
    }

    /// Fetches, then reports what pulling needs: nothing, a commit first, a merge draft, or a
    /// working-copy replacement plan (fast-forward).
    pub fn remote_pull(&self, request: &RemoteRequest) -> Result<Value> {
        let opened = self.open_original(&request.document, "拉取")?;
        let repo = opened.repo.clone();
        if repo.remote_url()?.is_none() {
            return Err(Error::invalid("请先点击「关联仓库」。"));
        }
        let branch = repo.active_branch()?;
        let head = repo.head()?.ok_or_else(|| Error::invalid("请先提交当前稿，或使用「克隆稿件」。"))?;
        let state = self.with_cancel(request.operation.as_deref(), |cancel| {
            let _lock = repo.lock()?;
            repo.fetch_remote(self.allow_local_remotes, cancel)
        })?;
        if state.unrelated {
            return Err(Error::invalid("本地与远程没有共同历史，不能自动合并。请确认关联的是同一份稿件的 WordGit 仓库。"));
        }
        let Some(remote_tip) = state.remote_tip.clone().filter(|_| state.behind > 0) else {
            return Ok(json!({ "action": "none", "remote": state, "message": format!("当前没有需要拉取的新版本。\n{}", state.caption) }));
        };
        if repo.active_branch()? != branch || repo.head()?.as_deref() != Some(head.as_str()) {
            return Err(Error::invalid("获取期间当前分支已变化，请重新拉取。"));
        }
        let head_hash = repo.snapshot_checksum(&head)?;
        if fsutil::file_checksum(&opened.path)? != head_hash {
            return Ok(json!({
                "action": "commitFirst", "remote": state,
                "message": "远程历史已获取。当前稿仍有未提交修改，请先提交，再点击「拉取」；当前稿没有被覆盖。",
            }));
        }
        if state.ahead > 0 {
            return Ok(json!({ "action": "merge", "source": state.remote_reference, "remote": state }));
        }
        review::ensure_replaceable(&opened.path)?;
        let lock = repo.lock()?;
        let staged = self.paths.new_session()?.join("远程更新.docx");
        history::extract_to_file(&repo, &remote_tip, &staged)?;
        docx::validate_file(&staged)?;
        let remote_hash = repo.snapshot_checksum(&remote_tip)?;
        let changes_file = remote_hash != head_hash;
        let plan = self.store_plan(ReplacePlan {
            repo,
            document: opened.path.clone(),
            current_hash: head_hash,
            staged,
            replacement_hash: remote_hash,
            publish: Publish::FastForward { branch, head, reference: state.remote_reference.clone(), remote_tip: remote_tip.clone() },
            done_message: format!("已拉取远程版本 {}。原有提交继续保留。", short(&remote_tip)),
            created: Instant::now(),
            _lock: lock,
        });
        Ok(json!({ "action": "replace", "plan": plan, "document": opened.path, "changesFile": changes_file, "remote": state }))
    }

    pub fn remote_clone(&self, request: &CloneRequest) -> Result<Value> {
        let document = fsutil::full_path(Path::new(request.document.trim()));
        review::ensure_replaceable(&document)?;
        if let Some(parent) = document.parent() {
            std::fs::create_dir_all(parent)?;
        }
        self.with_cancel(request.operation.as_deref(), |cancel| {
            remote::clone_remote(
                &request.url,
                &request.branch,
                &document,
                &self.paths.repositories(),
                self.git.as_deref(),
                self.allow_local_remotes,
                cancel,
            )
        })?;
        Ok(json!({
            "document": document,
            "message": format!("远程历史已下载，稿件已还原到：\n{}", document.display()),
        }))
    }

    /// Startup work: recover interrupted replacements and clean old sessions.
    pub fn startup(&self) {
        let messages = journal::recover_all(&self.paths, self.git.as_deref());
        self.add_notices(messages);
        self.paths.clean_sessions(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn document_paths_are_validated() {
        assert!(document_path("").is_err());
        assert!(document_path("https://x/y.docx").is_err());
        assert!(document_path("/tmp/a.doc").is_err());
    }
}
