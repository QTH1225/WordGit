//! A document's version history: one bare Git repository per document, chunked snapshots,
//! branches, imported review drafts and storage maintenance.

use crate::docx;
use crate::fsutil::{self, FileLock, Stamp};
use crate::git::{Git, GitResult};
use crate::snapshot::{self, DocumentSnapshot, ObjectFormat, SnapshotPlan};
use crate::{object_id, Error, Result};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::sync::OnceLock;
use std::time::{Duration, SystemTime};

const FILE_NAME: &str = "document.docx";
pub const HEADS: &str = "refs/heads/";
pub const REVIEW_ARCHIVE_PREFIX: &str = "refs/wordgit/reviews/";
const MANIFEST_LIMIT: u64 = 64 * 1024 * 1024;
const META_LIMIT: u64 = 16 * 1024 * 1024;
/// Unreachable objects (failed imports, expired remote backups) are dropped after this age.
const CRUFT_EXPIRATION: &str = "2.weeks.ago";
const FULL_REPACK_INTERVAL: Duration = Duration::from_secs(30 * 86_400);
const RESERVED_NAMES: [&str; 22] = [
    "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1",
    "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
];

#[derive(Clone, Debug, Serialize)]
pub struct Revision {
    pub id: String,
    pub parents: Vec<String>,
    /// Strict ISO 8601 author date (`%aI`).
    pub date: String,
    pub author: String,
    pub references: Vec<String>,
    pub subject: String,
    pub message: String,
}

#[derive(Clone, Debug, Serialize)]
pub struct RepackResult {
    pub before_bytes: u64,
    pub after_bytes: u64,
}

/// A snapshot tree already written to the object store, plus the document's SHA-256.
#[derive(Clone, Debug, Serialize, serde::Deserialize)]
pub struct PreparedSnapshot {
    pub tree: String,
    pub checksum: String,
}

#[derive(Debug, Serialize)]
pub struct ImportedReview {
    pub commit: String,
    pub source: String,
    pub branch: Option<String>,
    pub reused: bool,
}

/// A planned chunked snapshot plus the file stamp proving the file did not change between the
/// planning pass and the import pass.
#[derive(Clone, Debug)]
pub struct FileSnapshot {
    pub plan: SnapshotPlan,
    pub stamp: Stamp,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SnapshotKind {
    Chunked,
    Expanded,
    Whole,
}

struct StoredSnapshot {
    kind: SnapshotKind,
    object: String,
    size: u64,
}

pub type CommitHook = std::sync::Arc<dyn Fn(&Path) + Send + Sync>;

#[derive(Clone)]
pub struct Repo {
    root: PathBuf,
    git: Git,
    commit_created: Option<CommitHook>,
}

/// Location of a document's history: `<storage>/<sha256(full path)>.git`. On Windows and macOS,
/// whose file systems ignore case by default, the path is upper-cased first (as the original
/// Windows add-in does), so existing Windows histories keep their location.
pub fn repository_for(document: &Path, storage: &Path) -> PathBuf {
    let full = fsutil::full_path(document).to_string_lossy().into_owned();
    let key = if cfg!(any(windows, target_os = "macos")) { upper_invariant(&full) } else { full };
    storage.join(format!("{}.git", sha256_hex(key.as_bytes())))
}

/// .NET `ToUpperInvariant`: one-to-one simple case mapping, characters without a single-character
/// upper case form are kept.
fn upper_invariant(text: &str) -> String {
    text.chars()
        .map(|c| {
            let mut upper = c.to_uppercase();
            match (upper.next(), upper.next()) {
                (Some(single), None) => single,
                _ => c,
            }
        })
        .collect()
}

pub fn sha256_hex(bytes: &[u8]) -> String {
    crate::hex(&Sha256::digest(bytes))
}

/// `update-ref --stdin` learned `symref-update` in Git 2.46.
fn git_version(git: &Git) -> (u32, u32) {
    static VERSION: OnceLock<(u32, u32)> = OnceLock::new();
    *VERSION.get_or_init(|| {
        let text = git.invoke(None, None, &["version"]).and_then(GitResult::text).unwrap_or_default();
        let numbers: Vec<u32> = text
            .split(|c: char| !c.is_ascii_digit())
            .filter(|part| !part.is_empty())
            .take(2)
            .filter_map(|part| part.parse().ok())
            .collect();
        (numbers.first().copied().unwrap_or(0), numbers.get(1).copied().unwrap_or(0))
    })
}

impl Repo {
    pub fn for_document(document: &Path, storage: &Path, executable: Option<&Path>, hook: Option<CommitHook>) -> Result<Repo> {
        Repo::open(repository_for(document, storage), executable, hook)
    }

    pub fn open(root: PathBuf, executable: Option<&Path>, hook: Option<CommitHook>) -> Result<Repo> {
        let executable = match executable {
            Some(path) => path.to_path_buf(),
            None => Git::find()?,
        };
        Ok(Repo { git: Git::new(executable, &root), root, commit_created: hook })
    }

    pub fn path(&self) -> &Path {
        &self.root
    }

    pub fn git(&self) -> &Git {
        &self.git
    }

    pub fn exists(&self) -> bool {
        self.root.join("HEAD").is_file()
    }

    // ---- low-level helpers -------------------------------------------------------------------

    pub(crate) fn invoke(&self, input: Option<&[u8]>, identity: Option<&str>, args: &[&str]) -> Result<GitResult> {
        if !self.root.is_dir() {
            std::fs::create_dir_all(&self.root)?;
        }
        self.git.invoke(input, identity, args)
    }

    pub(crate) fn run(&self, input: Option<&[u8]>, identity: Option<&str>, args: &[&str]) -> Result<GitResult> {
        self.invoke(input, identity, args)?.checked()
    }

    pub(crate) fn text(&self, args: &[&str]) -> Result<String> {
        self.invoke(None, None, args)?.text()
    }

    pub(crate) fn text_with(&self, input: &[u8], identity: Option<&str>, args: &[&str]) -> Result<String> {
        self.invoke(Some(input), identity, args)?.text()
    }

    pub(crate) fn ref_transaction(&self, identity: Option<&str>, operations: &[String]) -> Result<()> {
        let script = format!("start\n{}\nprepare\ncommit\n", operations.join("\n"));
        self.run(Some(script.as_bytes()), identity, &["update-ref", "--stdin"])?;
        Ok(())
    }

    pub(crate) fn ensure_repository(&self) -> Result<()> {
        std::fs::create_dir_all(&self.root)?;
        if !self.exists() {
            self.run(None, None, &["init", "--bare", "--initial-branch=main", "--template="])?;
        }
        Ok(())
    }

    fn object_format(&self) -> Result<ObjectFormat> {
        ObjectFormat::parse(&self.text(&["rev-parse", "--show-object-format"])?)
    }

    // ---- branches ----------------------------------------------------------------------------

    pub fn active_branch(&self) -> Result<String> {
        if !self.exists() {
            return Ok("main".into());
        }
        let reference = self.text(&["symbolic-ref", "HEAD"])?;
        reference
            .strip_prefix(HEADS)
            .map(str::to_string)
            .ok_or_else(|| Error::invalid("HEAD 未指向本地分支。"))
    }

    pub fn head(&self) -> Result<Option<String>> {
        let branch = self.active_branch()?;
        self.branch_tip(&branch)
    }

    /// Tip of a local branch name, or of a full `refs/remotes/origin/...` reference.
    pub fn branch_tip(&self, branch: &str) -> Result<Option<String>> {
        if !self.exists() {
            return Ok(None);
        }
        let reference = history_reference(branch)?;
        let result = self.invoke(None, None, &["rev-parse", "--verify", "-q", &reference])?;
        if result.code == 1 {
            return Ok(None);
        }
        Ok(Some(object_id(&result.text()?)?))
    }

    pub fn branches(&self, merged_into: Option<&str>) -> Result<Vec<String>> {
        if !self.exists() {
            return Ok(Vec::new());
        }
        let merged = match merged_into {
            Some(tip) => Some(format!("--merged={}", object_id(tip)?)),
            None => None,
        };
        let mut args = vec!["for-each-ref", "--sort=refname", "--format=%(refname:strip=2)"];
        if let Some(merged) = merged.as_deref() {
            args.push(merged);
        }
        args.push(HEADS);
        Ok(lines(&self.text(&args)?))
    }

    /// HEAD's target, local branches and archived returns. Fetching only moves remote-tracking
    /// and backup refs.
    pub fn reference_state(&self) -> Result<String> {
        if !self.exists() {
            return Ok(String::new());
        }
        Ok(format!(
            "{}\n{}",
            self.text(&["symbolic-ref", "HEAD"])?,
            self.text(&["for-each-ref", "--format=%(objectname) %(refname)", HEADS, REVIEW_ARCHIVE_PREFIX])?
        ))
    }

    /// A revision's own tree serves as the prepared snapshot when the working copy is replaced
    /// by exactly that revision, so publishing a restore needs no second pass over the file.
    pub fn prepared_from_revision(&self, revision: &str) -> Result<PreparedSnapshot> {
        let revision = object_id(revision)?;
        let tree = object_id(&self.text(&["rev-parse", &format!("{revision}^{{tree}}")])?)?;
        Ok(PreparedSnapshot { tree, checksum: self.snapshot_checksum(&revision)? })
    }

    pub fn create_branch(&self, name: &str, revision: &str) -> Result<()> {
        let reference = branch_ref(name)?;
        self.run(None, None, &["check-ref-format", &reference])?;
        if self.branches(None)?.iter().any(|b| b.eq_ignore_ascii_case(name) || b.to_lowercase() == name.to_lowercase()) {
            return Err(Error::invalid("同名分支已存在。"));
        }
        let revision = object_id(revision)?;
        // Prove the starting version is intact before a branch can point at it.
        self.extract_to(&revision, &mut std::io::sink())?;
        let zeros = "0".repeat(revision.len());
        self.run(None, None, &["update-ref", &reference, &revision, &zeros])?;
        Ok(())
    }

    /// Archives the return once per exact snapshot and submitted version. The caller holds the
    /// repository lock. Older imports are found through their commit trailers without rewriting
    /// their history.
    pub fn import_review_file(
        &self,
        path: &Path,
        basis: &str,
        new_branch: Option<&str>,
        reviewer: &str,
        target_branch: &str,
        expected_target: &str,
    ) -> Result<ImportedReview> {
        let identity = sanitize_identity(reviewer);
        if identity.is_empty() || reviewer.contains('\0') {
            return Err(Error::invalid("请填写返稿来源。"));
        }
        let basis = object_id(basis)?;
        let expected = object_id(expected_target)?;
        let unchanged = || -> Result<bool> {
            Ok(self.active_branch()? == target_branch && self.head()?.as_deref() == Some(expected.as_str()))
        };
        if !unchanged()? {
            return Err(Error::invalid("当前分支或提交已变化，请刷新后重新导入。"));
        }
        let file = self.plan_file(path)?;
        let checksum = &file.plan.checksum;
        let marker = format!("--grep=WordGit-Review-SHA256: {checksum}");
        let matches = self.text(&["log", "--all", "--format=%H", "--fixed-strings", &marker, "--"])?;
        let mut reusable = None;
        for candidate in matches.lines() {
            if self.imported_review_base(candidate)?.as_deref() != Some(&basis) || self.snapshot_checksum(candidate)? != *checksum {
                continue;
            }
            if self.is_ancestor(candidate, &expected)? {
                reusable = Some(candidate);
                break;
            }
            reusable.get_or_insert(candidate);
        }
        let reused = reusable.is_some();
        let commit = if let Some(candidate) = reusable {
            candidate.to_string()
        } else {
            self.extract_to(&basis, &mut std::io::sink())?;
            let prepared = self.import_plan(path, &file)?;
            let file_name = path.file_name().map(|n| n.to_string_lossy().replace(['\r', '\n'], " ")).unwrap_or_default();
            let message = format!("导入审阅稿：{file_name}\n\nWordGit-Review-Base: {basis}\nWordGit-Review-SHA256: {checksum}\n");
            object_id(&self.text_with(
                message.as_bytes(), Some(&identity), &["commit-tree", &prepared.tree, "-p", &basis, "-F", "-"],
            )?)?
        };
        let source = format!("{REVIEW_ARCHIVE_PREFIX}{commit}");
        let mut updates = Vec::new();
        if self.branch_tip(&source)?.is_none() {
            updates.push(format!("create {source} {commit}"));
        }
        let branch = if let Some(name) = new_branch {
            let points_at = format!("--points-at={commit}");
            let aliases = self.text(&["for-each-ref", &points_at, "--format=%(refname:strip=2)", HEADS])?;
            if let Some(existing) = aliases.lines().find(|alias| *alias != target_branch) {
                Some(existing.to_string())
            } else {
                let reference = branch_ref(name)?;
                self.run(None, None, &["check-ref-format", &reference])?;
                if let Some(conflict) = review_branch_conflict(name, &self.branches(None)?) {
                    return Err(Error::invalid(conflict));
                }
                updates.push(format!("create {reference} {commit}"));
                Some(name.to_string())
            }
        } else {
            None
        };
        if !unchanged()? || fsutil::stamp(path)? != file.stamp {
            return Err(Error::invalid("导入期间当前分支或返稿文件发生变化，请重试。"));
        }
        if !updates.is_empty() {
            updates.push(format!("verify {} {expected}", branch_ref(target_branch)?));
            self.ref_transaction(Some(&identity), &updates)?;
            self.notify_commit();
        }
        Ok(ImportedReview { commit, source, branch, reused })
    }

    /// The basis recorded by `import_review_file`, when `revision` is such an import.
    pub fn imported_review_base(&self, revision: &str) -> Result<Option<String>> {
        let text = self.invoke(None, None, &["show", "-s", "--format=%P%n%B", &object_id(revision)?])?.lenient_text()?;
        let lines: Vec<&str> = text.split('\n').collect();
        let markers: Vec<&str> = lines.iter().filter_map(|l| l.strip_prefix("WordGit-Review-Base: ")).collect();
        if markers.len() != 1 {
            return Ok(None);
        }
        let parent = markers[0].trim();
        Ok((crate::is_object_id(parent) && lines[0].trim() == parent).then(|| parent.to_string()))
    }

    pub fn revision_author(&self, revision: &str) -> Result<String> {
        self.text(&["show", "-s", "--format=%an", &object_id(revision)?])
    }

    /// Commits that merged a return draft of `basis` into `tip`, in `git log` order.
    ///
    /// A return draft records its basis in `WordGit-Review-Base`, so the merges that brought an
    /// earlier round of the same review into this branch can be found again. The add-in uses them
    /// to keep a repeated return from being merged on top of itself.
    pub fn previously_merged_returns(&self, tip: &str, basis: &str) -> Result<Vec<String>> {
        let range = format!("{}..{}", object_id(basis)?, object_id(tip)?);
        let text = self.invoke(None, None, &["log", "--first-parent", "--merges", "--format=%P", &range, "--"])?.lenient_text()?;
        let mut parents: Vec<String> = Vec::new();
        for line in text.split('\n') {
            // `%P` lists every parent; the first one is on the first-parent chain, the rest came
            // in with the merge and are the candidates.
            for parent in line.split(' ').skip(1).filter(|p| crate::is_object_id(p)) {
                if !parents.iter().any(|known| known == parent) {
                    parents.push(parent.to_string());
                }
            }
        }
        let mut returns = Vec::new();
        for parent in parents {
            if self.imported_review_base(&parent)?.as_deref() == Some(basis) {
                returns.push(parent);
            }
        }
        Ok(returns)
    }

    /// Moves HEAD from `from` (at `old_tip`) to `to` (at `new_tip`) in one checked transaction.
    pub fn switch(&self, from: &str, old_tip: &str, to: &str, new_tip: &str) -> Result<()> {
        let current_ref = branch_ref(from)?;
        let target_ref = branch_ref(to)?;
        if self.active_branch()? != from || self.branch_tip(from)?.as_deref() != Some(old_tip) {
            return Err(Error::invalid("当前分支已变化，请刷新后重试。"));
        }
        let new_tip = object_id(new_tip)?;
        if git_version(&self.git) >= (2, 46) {
            return self.ref_transaction(
                None,
                &[
                    format!("verify {target_ref} {new_tip}"),
                    "option no-deref".into(),
                    format!("symref-update HEAD {target_ref} ref {current_ref}"),
                ],
            );
        }
        // Older Git: the caller holds the repository lock, so verifying and then pointing HEAD
        // at the target cannot interleave with another WordGit operation.
        if self.branch_tip(to)?.as_deref() != Some(new_tip.as_str()) {
            return Err(Error::invalid("目标分支已变化，请刷新后重试。"));
        }
        self.run(None, None, &["symbolic-ref", "HEAD", &target_ref])?;
        Ok(())
    }

    pub fn delete_merged_branch(&self, current: &str, head: &str, target: &str, target_tip: &str) -> Result<()> {
        let active = self.active_branch()?;
        if current.eq_ignore_ascii_case(target) || active.eq_ignore_ascii_case(target) {
            return Err(Error::invalid("不能删除当前分支。"));
        }
        if active != current
            || self.branch_tip(current)?.as_deref() != Some(head)
            || self.branch_tip(target)?.as_deref() != Some(target_tip)
        {
            return Err(Error::invalid("分支或版本历史已变化，请刷新后重试。"));
        }
        if !self.is_ancestor(target_tip, head)? {
            return Err(Error::invalid("该分支还有提交未合入当前分支。"));
        }
        // Git locks HEAD and its referent. If HEAD now points to the target, the two operations conflict.
        self.ref_transaction(
            None,
            &[
                format!("verify HEAD {}", object_id(head)?),
                "option no-deref".into(),
                format!("delete {} {}", branch_ref(target)?, object_id(target_tip)?),
            ],
        )
    }

    /// Exclusive lock for one WordGit operation on this repository (shared with the Windows add-in).
    pub fn lock(&self) -> Result<FileLock> {
        std::fs::create_dir_all(&self.root)?;
        fsutil::try_lock(&self.root.join("wordgit.lock"))?
            .ok_or_else(|| Error::invalid("另一处 WordGit 正在处理这个文档，请稍后重试。"))
    }

    // ---- maintenance -------------------------------------------------------------------------

    /// Full maintenance: reachable objects into one pack, recent unreachable objects into a cruft
    /// pack, and unreachable objects older than the expiration dropped.
    pub fn repack(&self) -> Result<RepackResult> {
        let Some(_packing) = self.try_repack_lock()? else {
            return Err(Error::invalid("此文档的历史正在压缩。"));
        };
        let tips = self.repack_tips()?;
        if tips.is_empty() {
            return Err(Error::invalid("尚无提交可压缩。"));
        }
        let before = self.object_bytes()?;
        self.repack_with(&tips, false, true)?;
        Ok(RepackResult { before_bytes: before, after_bytes: self.object_bytes()? })
    }

    /// Repacks when at least `threshold` commits were added since the last pack. Returns false when
    /// it could not start now (busy, or `can_start` declined) and should be retried later.
    pub fn try_auto_repack(&self, threshold: usize, can_start: &dyn Fn() -> bool) -> Result<bool> {
        if !self.exists() {
            return Ok(true);
        }
        let Some(_packing) = self.try_repack_lock()? else { return Ok(false) };
        let tips = self.repack_tips()?;
        if tips.is_empty() {
            return Ok(true);
        }
        let mut input: Vec<String> = tips.clone();
        if let Ok(previous) = std::fs::read_to_string(self.packed_tips_path()) {
            for tip in previous.lines().map(str::trim).filter(|l| !l.is_empty()) {
                input.push(format!("^{}", object_id(tip)?));
            }
        }
        let max_count = format!("--max-count={threshold}");
        let count = self.text_with(
            format!("{}\n", input.join("\n")).as_bytes(),
            None,
            &["rev-list", "--count", "--ignore-missing", &max_count, "--stdin"],
        )?;
        if count.parse::<usize>().unwrap_or(0) < threshold {
            return Ok(true);
        }
        if !can_start() {
            return Ok(false);
        }
        let full = std::fs::metadata(self.full_repack_path())
            .and_then(|m| m.modified())
            .map(|t| SystemTime::now().duration_since(t).unwrap_or_default() > FULL_REPACK_INTERVAL)
            .unwrap_or(true);
        self.repack_with(&tips, true, full)?;
        Ok(true)
    }

    fn packed_tips_path(&self) -> PathBuf {
        self.root.join("pack-tips")
    }

    fn full_repack_path(&self) -> PathBuf {
        self.root.join("pack-full")
    }

    fn try_repack_lock(&self) -> Result<Option<FileLock>> {
        fsutil::try_lock(&self.root.join("wordgit-repack.lock"))
    }

    fn repack_tips(&self) -> Result<Vec<String>> {
        if !self.exists() {
            return Ok(Vec::new());
        }
        let mut tips = Vec::new();
        for line in lines(&self.text(&["for-each-ref", "--format=%(objectname)", HEADS])?) {
            let id = object_id(&line)?;
            if !tips.contains(&id) {
                tips.push(id);
            }
        }
        Ok(tips)
    }

    fn repack_with(&self, tips: &[String], background: bool, full: bool) -> Result<()> {
        // Neither mode deletes objects newer than the cruft expiration, so objects written by a
        // concurrent commit survive even before its branch update is published. Geometric repacks
        // only merge small packs, keeping routine maintenance cheap on large histories.
        let expiration = format!("--cruft-expiration={CRUFT_EXPIRATION}");
        let mut args: Vec<&str> = if full {
            vec!["repack", "--cruft", "-d", &expiration, "-n", "-q"]
        } else {
            vec!["repack", "--geometric=2", "-d", "-n", "-q"]
        };
        if background {
            args.extend(["--threads=1", "--window-memory=64m"]);
        }
        self.git.invoke_with(None, None, None, background, &args)?.checked()?;
        // repack only expires packed objects; loose unreachable objects need prune, which also
        // keeps everything newer than the expiration.
        if full {
            let prune = format!("--expire={CRUFT_EXPIRATION}");
            self.git.invoke_with(None, None, None, background, &["prune", prune.as_str()])?.checked()?;
        }
        // Record the starting tips so a commit made during packing is counted next time.
        fsutil::write_durable(&self.packed_tips_path(), format!("{}\n", tips.join("\n")).as_bytes())?;
        if full {
            let stamp = crate::paths::format_utc(crate::paths::now_seconds());
            fsutil::write_durable(&self.full_repack_path(), format!("{stamp}\n").as_bytes())?;
        }
        Ok(())
    }

    fn object_bytes(&self) -> Result<u64> {
        let text = self.text(&["count-objects", "-v"])?;
        let mut kib = 0u64;
        for line in text.lines() {
            for key in ["size: ", "size-pack: ", "size-garbage: "] {
                if let Some(value) = line.strip_prefix(key) {
                    kib += value.trim().parse::<u64>().unwrap_or(0);
                }
            }
        }
        Ok(kib * 1024)
    }

    // ---- history -----------------------------------------------------------------------------

    pub fn is_ancestor(&self, ancestor: &str, descendant: &str) -> Result<bool> {
        let result = self.invoke(
            None,
            None,
            &["merge-base", "--is-ancestor", &object_id(ancestor)?, &object_id(descendant)?],
        )?;
        if result.code == 1 {
            return Ok(false);
        }
        result.checked()?;
        Ok(true)
    }

    /// Basis for an editor revision merge. Criss-cross histories have several equally good bases;
    /// Git's own choice is used, and the author still reviews every resulting revision.
    pub fn merge_base(&self, left: &str, right: &str) -> Result<String> {
        let result = self.invoke(None, None, &["merge-base", &object_id(left)?, &object_id(right)?])?;
        if result.code == 1 {
            return Err(Error::invalid("两个分支没有共同祖先，无法创建可靠的修订稿。"));
        }
        object_id(&result.text()?)
    }

    /// Commits of one branch, or of every ref (`None`). `--all` deliberately includes
    /// `refs/wordgit/remote-backup/*`: commits a collaborator force-pushed away stay visible
    /// (unlabelled) instead of silently disappearing.
    pub fn history(&self, branch: Option<&str>) -> Result<Vec<Revision>> {
        if !self.exists() {
            return Ok(Vec::new());
        }
        let selector = match branch {
            Some(branch) => history_reference(branch)?,
            None => "--all".into(),
        };
        let raw = self
            .invoke(
                None,
                None,
                &[
                    "log",
                    &selector,
                    "--topo-order",
                    "--decorate=full",
                    "-z",
                    "--format=%H%x00%P%x00%aI%x00%an%x00%D%x00%s%x00%B",
                    "--",
                ],
            )?
            .lenient_text()?;
        if raw.is_empty() {
            return Ok(Vec::new());
        }
        let raw = raw.strip_suffix('\0').unwrap_or(&raw);
        let fields: Vec<&str> = raw.split('\0').collect();
        if !fields.len().is_multiple_of(7) {
            return Err(Error::data("无法读取此仓库的版本历史。"));
        }
        let mut revisions = Vec::with_capacity(fields.len() / 7);
        for record in fields.chunks(7) {
            revisions.push(Revision {
                // Records after the first start with the newline `log -z` puts between entries.
                id: object_id(record[0].trim_start_matches('\n'))?,
                parents: record[1].split(' ').filter(|p| !p.is_empty()).map(object_id).collect::<Result<_>>()?,
                date: record[2].to_string(),
                author: record[3].to_string(),
                references: record[4]
                    .split(", ")
                    .filter(|r| !r.is_empty() && !r.starts_with("refs/wordgit/"))
                    .map(|r| r.replace("refs/heads/", "").replace("refs/tags/", "").replace("refs/remotes/", ""))
                    .collect(),
                subject: record[5].to_string(),
                message: record[6].trim_end().to_string(),
            });
        }
        Ok(revisions)
    }

    // ---- snapshots ---------------------------------------------------------------------------

    /// Which storage format a revision uses, without reading any content.
    fn locate(&self, revision: &str) -> Result<StoredSnapshot> {
        let specs: Vec<String> =
            [snapshot::MANIFEST_FILE, docx::META_NAME, FILE_NAME].iter().map(|name| format!("{revision}:{name}")).collect();
        let output = self.text_with(format!("{}\n", specs.join("\n")).as_bytes(), None, &["cat-file", "--batch-check"])?;
        let found = lines(&output);
        if found.len() != specs.len() {
            return Err(Error::data("无法读取版本格式。"));
        }
        for (line, kind) in found.iter().zip([SnapshotKind::Chunked, SnapshotKind::Expanded, SnapshotKind::Whole]) {
            let fields: Vec<&str> = line.split(' ').collect();
            if fields.len() == 3 && fields[1] == "blob" && crate::is_object_id(fields[0]) {
                if let Ok(size) = fields[2].parse::<u64>() {
                    return Ok(StoredSnapshot { kind, object: fields[0].to_string(), size });
                }
            }
        }
        Err(Error::data(format!(
            "版本 {} 不包含可识别的文档快照。请保留仓库并检查版本格式；此操作不会删除历史。",
            &revision[..8.min(revision.len())]
        )))
    }

    fn read_small_blob(&self, stored: &StoredSnapshot, limit: u64) -> Result<Vec<u8>> {
        if stored.size > limit {
            return Err(Error::data("版本元数据过大，已拒绝读取。"));
        }
        let bytes = self.run(None, None, &["cat-file", "blob", &stored.object])?.stdout;
        if bytes.len() as u64 != stored.size {
            return Err(Error::data("版本元数据读取不完整。"));
        }
        Ok(bytes)
    }

    fn read_manifest(&self, stored: &StoredSnapshot) -> Result<DocumentSnapshot> {
        DocumentSnapshot::read(&self.read_small_blob(stored, MANIFEST_LIMIT)?)
    }

    /// Restores a revision into `output` with bounded memory, for every storage format. The output
    /// is only valid when this returns `Ok`: callers must discard it on error.
    pub fn extract_to(&self, revision: &str, output: &mut (dyn Write + Send)) -> Result<()> {
        let revision = object_id(revision)?;
        let stored = self.locate(&revision)?;
        match stored.kind {
            SnapshotKind::Chunked => {
                let manifest = self.read_manifest(&stored)?;
                let request: String = manifest
                    .objects
                    .iter()
                    .map(|id| format!("{revision}:{}/{id}\n", snapshot::CHUNKS_FOLDER))
                    .collect();
                self.git.stream(
                    Some(request.as_bytes()),
                    None,
                    &mut |stream: &mut dyn Read| manifest.restore(stream, &mut *output),
                    &["cat-file", "--batch"],
                )
            }
            SnapshotKind::Expanded => {
                let text = String::from_utf8(self.read_small_blob(&stored, META_LIMIT)?)
                    .map_err(|_| Error::data("无法读取此仓库的版本元数据。"))?;
                self.extract_expanded(&revision, &docx::parse_meta(&text)?, output)
            }
            SnapshotKind::Whole => {
                let size = stored.size;
                self.git.stream(
                    None,
                    None,
                    &mut |stream: &mut dyn Read| copy_exactly(stream, &mut *output, size),
                    &["cat-file", "blob", &stored.object],
                )
            }
        }
    }

    /// Pre-chunking format: each ZIP member stored as its own blob plus a manifest. The container
    /// is rebuilt member by member straight into the output.
    fn extract_expanded(&self, revision: &str, meta: &docx::DocxMeta, output: &mut (dyn Write + Send)) -> Result<()> {
        let mut names: Vec<&str> = meta.entries.iter().map(|e| e.name.as_str()).collect();
        names.sort_unstable();
        names.dedup();
        if meta.entries.is_empty() || names.len() != meta.entries.len() {
            return Err(Error::data("旧版文档清单为空或包含重复部件。"));
        }
        let request: String = meta.entries.iter().map(|e| format!("{revision}:{}\n", e.name)).collect();
        self.git.stream(
            Some(request.as_bytes()),
            None,
            &mut |batch: &mut dyn Read| -> Result<()> {
                let mut reader = snapshot::ByteReader::new(batch);
                let mut zip = docx::ZipWriter::new(&mut *output);
                let mut buffer = vec![0u8; 64 * 1024];
                for item in &meta.entries {
                    let header = reader.header(8192).ok_or_else(|| Error::data("旧版文档部件读取不完整。"))?;
                    let fields: Vec<&str> = header.split(' ').collect();
                    let size_ok = fields.len() == 3
                        && fields[1] == "blob"
                        && crate::is_object_id(fields[0])
                        && fields[2].parse::<u64>().ok() == Some(item.size);
                    if !size_ok {
                        return Err(Error::data(format!("旧版文档部件缺失或长度不符：{}", item.name)));
                    }
                    let mut entry = zip.start_entry(&item.name)?;
                    let mut sha = Sha256::new();
                    let mut remaining = item.size;
                    while remaining > 0 {
                        let step = remaining.min(buffer.len() as u64) as usize;
                        reader
                            .read_exact(&mut buffer[..step])
                            .map_err(|_| Error::data(format!("旧版文档部件内容不完整：{}", item.name)))?;
                        sha.update(&buffer[..step]);
                        zip.write_entry(&mut entry, &buffer[..step])?;
                        remaining -= step as u64;
                    }
                    zip.end_entry(entry)?;
                    if crate::hex(&sha.finalize()) != item.sha {
                        return Err(Error::data(format!("旧版文档部件校验失败：{}", item.name)));
                    }
                    if reader.byte() != Some(b'\n') {
                        return Err(Error::data(format!("旧版文档部件内容不完整：{}", item.name)));
                    }
                }
                zip.finish()?;
                if reader.byte().is_some() || docx::fingerprint(&meta.entries)? != meta.fingerprint {
                    return Err(Error::data("旧版文档整体部件校验失败。"));
                }
                Ok(())
            },
            &["cat-file", "--batch"],
        )
    }

    /// First streaming pass over the document: whole-file checksum and the chunk plan, without
    /// retaining chunk content in memory.
    pub fn plan_file(&self, path: &Path) -> Result<FileSnapshot> {
        self.ensure_repository()?;
        docx::validate_file(path)?;
        let format = self.object_format()?;
        let before = fsutil::stamp(path)?;
        let plan = snapshot::plan(&mut fsutil::open_shared(path)?, format)?;
        let after = fsutil::stamp(path)?;
        if after != before || plan.length != before.length {
            return Err(Error::invalid("文档在提交准备期间被修改，请重新提交。"));
        }
        Ok(FileSnapshot { plan, stamp: before })
    }

    /// Second pass: reads only the chunks Git does not have yet and streams them into
    /// fast-import. Verifying every imported blob id against the plan also detects the file
    /// changing mid-import.
    pub fn import_plan(&self, path: &Path, file: &FileSnapshot) -> Result<PreparedSnapshot> {
        let plan = &file.plan;
        let changed = || Error::invalid("文档在提交准备期间被修改，请重新提交。");
        let missing = self.missing_chunk_indexes(&plan.objects)?;
        if !missing.is_empty() {
            let mut stream = fsutil::open_shared(path)?;
            if fsutil::stamp(path)? != file.stamp {
                return Err(changed());
            }
            let result = self.import_chunks(&plan.objects, &plan.sizes, &missing, &mut |target: &mut dyn Write, index| {
                stream.seek(SeekFrom::Start(plan.offsets[index]))?;
                let mut buffer = vec![0u8; plan.sizes[index]];
                stream.read_exact(&mut buffer).map_err(|_| changed())?;
                target.write_all(&buffer)?;
                Ok(())
            });
            if let Err(Error::Data(message)) = &result {
                if message.contains("数据块写入校验失败") && fsutil::stamp(path).map(|s| s != file.stamp).unwrap_or(true) {
                    return Err(changed());
                }
            }
            result?;
        }
        self.store_manifest(&plan.objects, &plan.sizes, plan.length, &plan.checksum)
    }

    pub fn prepare_snapshot_file(&self, path: &Path) -> Result<PreparedSnapshot> {
        let file = self.plan_file(path)?;
        self.import_plan(path, &file)
    }

    fn missing_chunk_indexes(&self, objects: &[String]) -> Result<Vec<usize>> {
        let mut unique: Vec<(usize, &String)> = Vec::new();
        let mut seen = std::collections::HashSet::new();
        for (index, id) in objects.iter().enumerate() {
            if seen.insert(id) {
                unique.push((index, id));
            }
        }
        let request: String = unique.iter().map(|(_, id)| format!("{id}\n")).collect();
        let check = lines(&self.text_with(request.as_bytes(), None, &["cat-file", "--batch-check"])?);
        if check.len() != unique.len() {
            return Err(Error::data("无法确认已保存的数据块。"));
        }
        Ok(unique
            .iter()
            .zip(&check)
            .filter(|((_, id), line)| **line == format!("{id} missing"))
            .map(|((index, _), _)| *index)
            .collect())
    }

    #[allow(clippy::type_complexity)]
    fn import_chunks(
        &self,
        objects: &[String],
        sizes: &[usize],
        missing: &[usize],
        write_chunk: &mut (dyn FnMut(&mut dyn Write, usize) -> Result<()> + Send),
    ) -> Result<()> {
        if missing.is_empty() {
            return Ok(());
        }
        let mut writer = |stream: &mut dyn Write| -> Result<()> {
            for (i, &index) in missing.iter().enumerate() {
                let mark = i + 1;
                stream.write_all(format!("blob\nmark :{mark}\ndata {}\n", sizes[index]).as_bytes())?;
                write_chunk(stream, index)?;
                stream.write_all(format!("\nget-mark :{mark}\n").as_bytes())?;
            }
            stream.write_all(b"done\n")?;
            Ok(())
        };
        let imported = self
            .git
            .invoke_with(Some(&mut writer), None, None, true, &["-c", "pack.threads=1", "fast-import", "--quiet", "--done"])?
            .text()?;
        let written = lines(&imported);
        let expected: Vec<&String> = missing.iter().map(|&i| &objects[i]).collect();
        if written.len() != expected.len() || written.iter().zip(&expected).any(|(a, b)| a != *b) {
            return Err(Error::data("数据块写入校验失败。"));
        }
        Ok(())
    }

    fn store_manifest(&self, objects: &[String], sizes: &[usize], length: u64, checksum: &str) -> Result<PreparedSnapshot> {
        let manifest = snapshot::manifest(length, checksum, sizes, objects)?;
        let manifest_id = object_id(&self.text_with(&manifest, None, &["hash-object", "-w", "--no-filters", "--stdin"])?)?;
        let mut distinct: Vec<&String> = objects.iter().collect();
        distinct.sort();
        distinct.dedup();
        let entries: String = distinct.iter().map(|id| format!("100644 blob {id}\t{id}\n")).collect();
        let chunks_tree = object_id(&self.text_with(entries.as_bytes(), None, &["mktree"])?)?;
        let root_entries = format!(
            "040000 tree {chunks_tree}\t{}\n100644 blob {manifest_id}\t{}\n",
            snapshot::CHUNKS_FOLDER,
            snapshot::MANIFEST_FILE
        );
        let tree = object_id(&self.text_with(root_entries.as_bytes(), None, &["mktree"])?)?;
        Ok(PreparedSnapshot { tree, checksum: checksum.to_string() })
    }

    /// SHA-256 of the document stored in `revision`.
    pub fn snapshot_checksum(&self, revision: &str) -> Result<String> {
        let revision = object_id(revision)?;
        let stored = self.locate(&revision)?;
        if stored.kind == SnapshotKind::Chunked {
            return Ok(self.read_manifest(&stored)?.checksum);
        }
        let mut hashing = HashingWriter::default();
        self.extract_to(&revision, &mut hashing)?;
        Ok(hashing.finish())
    }

    // ---- commits -----------------------------------------------------------------------------

    /// Commits the file at `path`; returns `None` when it equals the expected head (no-op).
    #[allow(clippy::too_many_arguments)]
    pub fn commit_file(
        &self,
        path: &Path,
        message: &str,
        author: &str,
        expected_head: Option<&str>,
        branch: Option<&str>,
        source_branch: Option<&str>,
        source_tip: Option<&str>,
    ) -> Result<Option<String>> {
        let branch = self.verify_commit_state(message, expected_head, branch, source_branch, source_tip)?;
        let file = self.plan_file(path)?;
        if source_branch.is_none() {
            if let Some(head) = expected_head {
                if file.plan.checksum == self.snapshot_checksum(head)? {
                    return Ok(None);
                }
            }
        }
        let prepared = self.import_plan(path, &file)?;
        self.finish_commit(&prepared, message, author, expected_head, &branch, source_branch, source_tip).map(Some)
    }

    /// Commits an already-prepared snapshot without reading the document again. Callers must hold
    /// the repository lock.
    #[allow(clippy::too_many_arguments)]
    pub fn commit_prepared(
        &self,
        prepared: &PreparedSnapshot,
        message: &str,
        author: &str,
        expected_head: Option<&str>,
        branch: Option<&str>,
        source_branch: Option<&str>,
        source_tip: Option<&str>,
    ) -> Result<Option<String>> {
        let branch = self.verify_commit_state(message, expected_head, branch, source_branch, source_tip)?;
        if source_branch.is_none() {
            if let Some(head) = expected_head {
                if prepared.checksum == self.snapshot_checksum(head)? {
                    return Ok(None);
                }
            }
        }
        self.finish_commit(prepared, message, author, expected_head, &branch, source_branch, source_tip).map(Some)
    }

    fn verify_commit_state(
        &self,
        message: &str,
        expected_head: Option<&str>,
        branch: Option<&str>,
        source_branch: Option<&str>,
        source_tip: Option<&str>,
    ) -> Result<String> {
        if message.trim().is_empty() || message.contains('\0') {
            return Err(Error::invalid("请填写有效的提交说明。"));
        }
        self.ensure_repository()?;
        let branch = match branch {
            Some(branch) => branch.to_string(),
            None => self.active_branch()?,
        };
        let stale = self.active_branch()? != branch
            || self.branch_tip(&branch)?.as_deref() != expected_head
            || match source_branch {
                Some(source) => source == branch || self.branch_tip(source)?.as_deref() != source_tip,
                None => false,
            };
        if stale {
            return Err(Error::invalid("分支或版本历史已变化。请刷新后再操作。"));
        }
        Ok(branch)
    }

    #[allow(clippy::too_many_arguments)]
    fn finish_commit(
        &self,
        prepared: &PreparedSnapshot,
        message: &str,
        author: &str,
        expected_head: Option<&str>,
        branch: &str,
        source_branch: Option<&str>,
        source_tip: Option<&str>,
    ) -> Result<String> {
        let reference = branch_ref(branch)?;
        let tree = object_id(&prepared.tree)?;
        let mut args: Vec<String> = vec!["commit-tree".into(), tree];
        if let Some(head) = expected_head {
            args.extend(["-p".into(), object_id(head)?]);
        }
        if source_branch.is_some() {
            let source_tip = source_tip.ok_or_else(|| Error::invalid("来源分支已包含在当前分支。"))?;
            match expected_head {
                Some(head) if !self.is_ancestor(source_tip, head)? => {}
                _ => return Err(Error::invalid("来源分支已包含在当前分支。")),
            }
            args.extend(["-p".into(), object_id(source_tip)?]);
        }
        args.extend(["-F".into(), "-".into()]);
        let mut identity = sanitize_identity(author);
        if identity.is_empty() {
            identity = "WordGit".into();
        }
        let argv: Vec<&str> = args.iter().map(String::as_str).collect();
        let commit = object_id(&self.text_with(format!("{}\n", message.trim()).as_bytes(), Some(&identity), &argv)?)?;
        let old = expected_head.map(str::to_string).unwrap_or_else(|| "0".repeat(commit.len()));
        let mut operations = vec![format!("update {reference} {commit} {old}")];
        if let (Some(source), Some(tip)) = (source_branch, source_tip) {
            operations.push(format!("verify {} {}", history_reference(source)?, object_id(tip)?));
        }
        self.ref_transaction(Some(&identity), &operations)?;
        self.notify_commit();
        Ok(commit)
    }

    fn notify_commit(&self) {
        if let Some(hook) = &self.commit_created {
            hook(&self.root);
        }
    }
}

/// Removes characters Git rejects in identities.
pub fn sanitize_identity(name: &str) -> String {
    name.chars().filter(|&c| c != '<' && c != '>' && (c as u32) >= 0x20).collect::<String>().trim().to_string()
}

/// Validates a local branch name and returns its full reference. Branch names must also be valid
/// file names on Windows, because collaborators may use the Windows add-in.
pub fn branch_ref(branch: &str) -> Result<String> {
    let invalid = || Error::invalid("请输入有效的 Git 分支名，不包含空格或特殊符号。");
    if branch.trim().is_empty()
        || branch.encode_utf16().count() > 150
        || branch.chars().any(|c| c.is_whitespace() || c.is_control() || "~^:?*[\\\"<>|".contains(c))
    {
        return Err(invalid());
    }
    for segment in branch.split('/') {
        let stem = segment.split('.').next().unwrap_or("");
        if segment.ends_with('.') || RESERVED_NAMES.contains(&upper_invariant(stem).as_str()) {
            return Err(Error::invalid("分支名不能使用 Windows 保留名称（如 CON、NUL、COM1），也不能以句点结尾。"));
        }
    }
    Ok(format!("{HEADS}{branch}"))
}

/// A local branch name, remote-tracking branch, or a full archived return reference.
pub fn history_reference(branch: &str) -> Result<String> {
    if let Some(commit) = branch.strip_prefix(REVIEW_ARCHIVE_PREFIX) {
        return Ok(format!("{REVIEW_ARCHIVE_PREFIX}{}", object_id(commit)?));
    }
    match branch.strip_prefix(crate::remote::REMOTE_PREFIX) {
        Some(rest) => Ok(format!("{}{}", crate::remote::REMOTE_PREFIX, &branch_ref(rest)?[HEADS.len()..])),
        None => branch_ref(branch),
    }
}

/// Why a review branch name would clash with existing branches (also as a Git ref directory).
pub fn review_branch_conflict(name: &str, branches: &[String]) -> Option<String> {
    let lower = name.to_lowercase();
    branches
        .iter()
        .find(|branch| {
            let other = branch.to_lowercase();
            other == lower || lower.starts_with(&format!("{other}/")) || other.starts_with(&format!("{lower}/"))
        })
        .map(|conflict| format!("返稿分支「{name}」与已有分支「{conflict}」冲突。请换一个名称，例如 review-老师审阅。"))
}

fn copy_exactly(input: &mut dyn Read, output: &mut dyn Write, size: u64) -> Result<()> {
    let mut buffer = vec![0u8; 64 * 1024];
    let mut remaining = size;
    while remaining > 0 {
        let want = remaining.min(buffer.len() as u64) as usize;
        let count = input.read(&mut buffer[..want])?;
        if count == 0 {
            return Err(Error::data("历史文档读取不完整。"));
        }
        output.write_all(&buffer[..count])?;
        remaining -= count as u64;
    }
    let mut extra = [0u8; 1];
    if input.read(&mut extra)? > 0 {
        return Err(Error::data("历史文档长度不符。"));
    }
    Ok(())
}

pub(crate) fn lines(text: &str) -> Vec<String> {
    text.split(['\r', '\n']).filter(|l| !l.is_empty()).map(str::to_string).collect()
}

/// Write-only sink computing the SHA-256 of everything written.
#[derive(Default)]
pub struct HashingWriter {
    sha: Sha256,
}

impl HashingWriter {
    pub fn finish(self) -> String {
        crate::hex(&self.sha.finalize())
    }
}

impl Write for HashingWriter {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.sha.update(buffer);
        Ok(buffer.len())
    }
    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// Restores `revision` to `path`. The target only appears once the version has been fully
/// verified, so a failed restore never leaves a truncated document under the chosen name.
pub fn extract_to_file(repo: &Repo, revision: &str, path: &Path) -> Result<()> {
    if path.exists() {
        return Err(Error::invalid(format!("目标文件已存在：{}", path.display())));
    }
    let temporary = fsutil::sibling(path, &format!(".{}.wordgit-tmp", crate::random_hex(4)));
    let result = (|| -> Result<()> {
        let mut file = std::fs::OpenOptions::new().write(true).create_new(true).open(&temporary)?;
        repo.extract_to(revision, &mut file)?;
        file.sync_all()?;
        drop(file);
        if path.exists() {
            return Err(Error::invalid(format!("目标文件已存在：{}", path.display())));
        }
        std::fs::rename(&temporary, path)?;
        fsutil::sync_parent(path);
        Ok(())
    })();
    fsutil::try_remove(&temporary);
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn branch_names() {
        assert_eq!(branch_ref("review/老师").unwrap(), "refs/heads/review/老师");
        assert!(branch_ref("a b").is_err());
        assert!(branch_ref("con").is_err());
        assert!(branch_ref("x/com1.txt").is_err());
        assert!(branch_ref("end.").is_err());
        assert!(history_reference("refs/remotes/origin/main").unwrap() == "refs/remotes/origin/main");
    }

    #[test]
    fn review_conflicts() {
        let branches = vec!["main".to_string(), "review".to_string()];
        assert!(review_branch_conflict("Main", &branches).is_some());
        assert!(review_branch_conflict("review/x", &branches).is_some());
        assert!(review_branch_conflict("review-2", &branches).is_none());
    }

    #[test]
    fn identities() {
        assert_eq!(sanitize_identity(" <张三>\n "), "张三");
    }

    #[test]
    fn invariant_upper_case() {
        assert_eq!(upper_invariant("c:\\users\\straße\\ä.docx"), "C:\\USERS\\STRAßE\\Ä.DOCX");
    }
}
