//! End-to-end tests of the core against a real `git`: storage formats, branches, review imports,
//! two-client remote collaboration, working-copy replacement, crash recovery and maintenance.

use std::path::{Path, PathBuf};
use std::sync::atomic::AtomicBool;
use wordgit_core::docx::{self, ZipWriter};
use wordgit_core::history::{self, Repo};
use wordgit_core::journal;
use wordgit_core::paths::AppPaths;
use wordgit_core::remote;
use wordgit_core::review;

struct Sandbox {
    root: PathBuf,
}

impl Sandbox {
    fn new(name: &str) -> Sandbox {
        let root = std::env::temp_dir().join(format!("wordgit-it-{name}-{}", wordgit_core::random_hex(4)));
        std::fs::create_dir_all(&root).unwrap();
        Sandbox { root }
    }

    fn storage(&self) -> PathBuf {
        self.root.join("repositories")
    }

    fn repo(&self, document: &Path) -> Repo {
        Repo::for_document(document, &self.storage(), None, None).unwrap()
    }
}

impl Drop for Sandbox {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.root);
    }
}

fn picture(seed: u32, size: usize) -> Vec<u8> {
    let mut state = seed;
    (0..size)
        .map(|_| {
            state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
            (state >> 24) as u8
        })
        .collect()
}

fn fixture(text: &str, image: &[u8]) -> Vec<u8> {
    let mut zip = ZipWriter::new(Vec::new());
    for (name, body) in [
        ("[Content_Types].xml", b"<?xml version=\"1.0\"?><Types/>".to_vec()),
        ("_rels/.rels", b"<Relationships/>".to_vec()),
        ("word/document.xml", format!("<w:document><w:body><w:p><w:t>{text}</w:t></w:p></w:body></w:document>").into_bytes()),
        ("word/media/image1.png", image.to_vec()),
    ] {
        let mut entry = zip.start_entry(name).unwrap();
        zip.write_entry(&mut entry, &body).unwrap();
        zip.end_entry(entry).unwrap();
    }
    zip.finish().unwrap()
}

fn extract(repo: &Repo, revision: &str) -> Vec<u8> {
    let mut out = Vec::new();
    repo.extract_to(revision, &mut out).unwrap();
    out
}

fn commit_bytes(repo: &Repo, folder: &Path, bytes: &[u8], message: &str, head: Option<&str>, branch: Option<&str>) -> Option<String> {
    let path = folder.join(format!("commit-{}.docx", wordgit_core::random_hex(4)));
    std::fs::write(&path, bytes).unwrap();
    let id = repo.commit_file(&path, message, "测试作者", head, branch, None, None).unwrap();
    std::fs::remove_file(&path).unwrap();
    id
}

#[test]
fn commits_branches_and_history() {
    let sandbox = Sandbox::new("history");
    let document = sandbox.root.join("论文 初稿.docx");
    let image = picture(1, 300 * 1024);
    let v1 = fixture("第一版", &image);
    let v2 = fixture("第二版：修改了实验条件", &image);
    std::fs::write(&document, &v1).unwrap();
    let repo = sandbox.repo(&document);
    assert_eq!(repo.head().unwrap(), None);
    assert!(repo.history(None).unwrap().is_empty());

    let first = repo.commit_file(&document, "初稿", "张三", None, None, None, None).unwrap().unwrap();
    assert_eq!(extract(&repo, &first), v1);
    // Unchanged content is not committed again.
    assert_eq!(repo.commit_file(&document, "再次提交", "张三", Some(&first), None, None, None).unwrap(), None);
    assert_eq!(repo.snapshot_checksum(&first).unwrap(), history::sha256_hex(&v1));

    std::fs::write(&document, &v2).unwrap();
    let second = repo.commit_file(&document, "第二版\n\n详细说明", "李四", Some(&first), None, None, None).unwrap().unwrap();
    assert_eq!(extract(&repo, &second), v2);
    let log = repo.history(None).unwrap();
    assert_eq!(log.len(), 2);
    assert_eq!(log[0].id, second);
    assert_eq!(log[0].subject, "第二版");
    assert_eq!(log[0].message, "第二版\n\n详细说明");
    assert_eq!(log[0].author, "李四");
    assert_eq!(log[0].parents, vec![first.clone()]);
    assert!(log[0].references.iter().any(|r| r.contains("main")));

    // Stale expectations are refused.
    assert!(repo.commit_file(&document, "过期", "x", Some(&first), None, None, None).is_err());

    // Branches.
    repo.create_branch("review/老师", &first).unwrap();
    assert!(repo.create_branch("REVIEW/老师", &first).is_err());
    assert_eq!(repo.branches(None).unwrap(), vec!["main".to_string(), "review/老师".to_string()]);
    repo.switch("main", &second, "review/老师", &first).unwrap();
    assert_eq!(repo.active_branch().unwrap(), "review/老师");
    assert_eq!(repo.head().unwrap().as_deref(), Some(first.as_str()));
    repo.switch("review/老师", &first, "main", &second).unwrap();
    // review/老师 is an ancestor of main, so it can be deleted.
    repo.delete_merged_branch("main", &second, "review/老师", &first).unwrap();
    assert_eq!(repo.branches(None).unwrap(), vec!["main".to_string()]);

    // Unchanged image data is shared between versions: only the changed chunks were new.
    let view = repo.read_view().unwrap();
    assert_eq!(view.history.len(), 2);
    assert_eq!(view.remote.local_tip.as_deref(), Some(second.as_str()));
    assert!(!view.remote.linked);
}

#[test]
fn chunks_are_deduplicated_and_storage_repacks() {
    let sandbox = Sandbox::new("dedup");
    let document = sandbox.root.join("图片.docx");
    let image = picture(7, 384 * 1024);
    let repo = sandbox.repo(&document);
    let mut head = None;
    let mut versions = Vec::new();
    for index in 0..10 {
        let bytes = fixture(&format!("修改正文 {index}"), &image);
        head = commit_bytes(&repo, &sandbox.root, &bytes, &format!("第 {} 版", index + 1), head.as_deref(), None);
        versions.push((head.clone().unwrap(), bytes));
    }
    let result = repo.repack().unwrap();
    assert!(result.after_bytes < (image.len() * 3) as u64, "image data duplicated: {result:?}");
    for (id, bytes) in &versions {
        assert_eq!(&extract(&repo, id), bytes);
    }
    // Automatic maintenance: nothing new since the last pack.
    assert!(repo.try_auto_repack(10, &|| true).unwrap());
}

/// Builds a commit in the retired whole-file format with plain Git plumbing.
fn legacy_commit(repo: &Repo, tree_entries: &[(String, Vec<u8>)]) -> String {
    let git = repo.git();
    let hash = |bytes: &[u8]| {
        git.invoke(Some(bytes), None, &["hash-object", "-w", "--no-filters", "--stdin"]).unwrap().text().unwrap()
    };
    let mut listing = String::new();
    for (name, bytes) in tree_entries {
        listing.push_str(&format!("100644 blob {}\t{name}\n", hash(bytes)));
    }
    let tree = git.invoke(Some(listing.as_bytes()), None, &["mktree"]).unwrap().text().unwrap();
    let commit =
        git.invoke(Some(b"legacy\n"), Some("旧版"), &["commit-tree", &tree, "-F", "-"]).unwrap().text().unwrap();
    git.invoke(None, None, &["update-ref", "refs/heads/main", &commit]).unwrap().checked().unwrap();
    commit
}

#[test]
fn reads_retired_storage_formats() {
    let sandbox = Sandbox::new("legacy");
    let document = sandbox.root.join("旧格式.docx");
    let repo = sandbox.repo(&document);
    let image = picture(3, 50_000);
    let bytes = fixture("整包格式", &image);
    // Bootstrap the repository through a normal commit, then add legacy commits on top.
    commit_bytes(&repo, &sandbox.root, &bytes, "init", None, None).unwrap();

    // 1. Whole document.docx blob.
    let whole = legacy_commit(&repo, &[("document.docx".into(), bytes.clone())]);
    assert_eq!(extract(&repo, &whole), bytes);
    assert_eq!(repo.snapshot_checksum(&whole).unwrap(), history::sha256_hex(&bytes));

    // 2. Expanded members + wordgit.meta (a nested member name exercises subtrees).
    let members: Vec<(String, Vec<u8>)> = vec![
        ("[Content_Types].xml".into(), b"<Types/>".to_vec()),
        ("word/document.xml".into(), "<w:t>拆包格式</w:t>".as_bytes().to_vec()),
        ("word/media/image1.png".into(), image.clone()),
    ];
    let entries: Vec<docx::MetaEntry> = members
        .iter()
        .map(|(name, body)| docx::MetaEntry { name: name.clone(), sha: history::sha256_hex(body), size: body.len() as u64 })
        .collect();
    let mut meta = format!("wordgit 2\nfingerprint {}\n", docx::fingerprint(&entries).unwrap());
    for entry in &entries {
        use base64::Engine;
        meta.push_str(&format!(
            "{} {} {}\n",
            base64::engine::general_purpose::STANDARD.encode(&entry.name),
            entry.size,
            entry.sha
        ));
    }
    let word: Vec<(String, Vec<u8>)> = vec![("document.xml".into(), members[1].1.clone())];
    // Flat tree with "word/..." needs real subtrees: build word/ and word/media/ by hand.
    let git = repo.git();
    let hash = |bytes: &[u8]| git.invoke(Some(bytes), None, &["hash-object", "-w", "--stdin"]).unwrap().text().unwrap();
    let media = git
        .invoke(Some(format!("100644 blob {}\timage1.png\n", hash(&image)).as_bytes()), None, &["mktree"])
        .unwrap()
        .text()
        .unwrap();
    let word_tree = git
        .invoke(
            Some(format!("100644 blob {}\tdocument.xml\n040000 tree {media}\tmedia\n", hash(&word[0].1)).as_bytes()),
            None,
            &["mktree"],
        )
        .unwrap()
        .text()
        .unwrap();
    let root = format!(
        "100644 blob {}\t[Content_Types].xml\n100644 blob {}\twordgit.meta\n040000 tree {word_tree}\tword\n",
        hash(&members[0].1),
        hash(meta.as_bytes())
    );
    let tree = git.invoke(Some(root.as_bytes()), None, &["mktree"]).unwrap().text().unwrap();
    let expanded = git.invoke(Some(b"expanded\n"), Some("旧版"), &["commit-tree", &tree, "-F", "-"]).unwrap().text().unwrap();
    let rebuilt = extract(&repo, &expanded);
    let path = sandbox.root.join("rebuilt.docx");
    std::fs::write(&path, &rebuilt).unwrap();
    let names = docx::zip_entry_names(&mut std::fs::File::open(&path).unwrap()).unwrap();
    assert_eq!(names, members.iter().map(|m| m.0.clone()).collect::<Vec<_>>());
    for (name, body) in &members {
        assert_eq!(&read_member(&rebuilt, name), body, "member {name}");
    }
}

/// Minimal reader for members written by our ZipWriter (data descriptors, deflate or stored).
fn read_member(zip: &[u8], wanted: &str) -> Vec<u8> {
    use std::io::Read;
    let le16 = |at: usize| u16::from_le_bytes([zip[at], zip[at + 1]]) as usize;
    let le32 = |at: usize| u32::from_le_bytes(zip[at..at + 4].try_into().unwrap()) as usize;
    let eocd = (0..zip.len() - 21).rev().find(|&i| zip[i..i + 4] == [0x50, 0x4b, 5, 6]).unwrap();
    let mut cursor = le32(eocd + 16);
    for _ in 0..le16(eocd + 10) {
        let method = le16(cursor + 10);
        let compressed = le32(cursor + 20);
        let name_length = le16(cursor + 28);
        let name = std::str::from_utf8(&zip[cursor + 46..cursor + 46 + name_length]).unwrap();
        let offset = le32(cursor + 42);
        if name == wanted {
            let data = offset + 30 + le16(offset + 26) + le16(offset + 28);
            let raw = &zip[data..data + compressed];
            return if method == 8 {
                let mut out = Vec::new();
                flate2::read::DeflateDecoder::new(raw).read_to_end(&mut out).unwrap();
                out
            } else {
                raw.to_vec()
            };
        }
        cursor += 46 + name_length + le16(cursor + 30) + le16(cursor + 32);
    }
    panic!("member {wanted} not found");
}

#[test]
fn imports_returned_drafts_and_merges_branches() {
    let sandbox = Sandbox::new("review");
    let document = sandbox.root.join("送审稿.docx");
    let image = picture(9, 20_000);
    let repo = sandbox.repo(&document);
    let sent = commit_bytes(&repo, &sandbox.root, &fixture("送审版", &image), "送审", None, None).unwrap();
    let later = commit_bytes(&repo, &sandbox.root, &fixture("送审后继续修改", &image), "继续写", Some(&sent), None).unwrap();

    let returned = sandbox.root.join("老师返稿.docx");
    let returned_bytes = fixture("老师改过的送审版", &image);
    std::fs::write(&returned, &returned_bytes).unwrap();
    assert!(repo.import_review_file(&returned, &sent, Some("review-1"), "老师", "main", &sent).is_err(), "stale target");
    let original = repo.import_review_file(&returned, &sent, None, "老师<a>", "main", &later).unwrap();
    assert!(!original.reused && original.branch.is_none());
    assert!(repo.history(None).unwrap().iter().any(|revision| revision.id == original.commit));
    let named = repo.import_review_file(&returned, &sent, Some("review-1"), "老师", "main", &later).unwrap();
    assert!(named.reused && named.branch.as_deref() == Some("review-1"));
    assert_eq!(named.commit, original.commit);
    let repeated = repo.import_review_file(&returned, &sent, Some("review-2"), "老师", "main", &later).unwrap();
    assert!(repeated.reused);
    assert_eq!(repeated.branch.as_deref(), Some("review-1"));
    assert_eq!(repo.branches(None).unwrap(), vec!["main", "review-1"]);
    let imported = original.commit;
    assert_eq!(repo.imported_review_base(&imported).unwrap().as_deref(), Some(sent.as_str()));
    assert_eq!(repo.imported_review_base(&later).unwrap(), None);
    assert_eq!(extract(&repo, &imported), returned_bytes);
    assert_eq!(repo.history(Some("review-1")).unwrap()[0].author, "老师a");
    assert!(repo.import_review_file(&returned, &later, Some("Review-1"), "老师", "main", &later).is_err(), "case conflict");
    let another_basis = repo.import_review_file(&returned, &later, None, "老师", "main", &later).unwrap();
    assert!(!another_basis.reused && another_basis.commit != imported);
    assert_eq!(repo.merge_base(&later, &imported).unwrap(), sent);
    assert!(!repo.is_ancestor(&imported, &later).unwrap());

    // Finish the branch merge with a two-parent commit.
    let merged_bytes = fixture("合并后的稿件", &image);
    let merged_path = sandbox.root.join("merged.docx");
    std::fs::write(&merged_path, &merged_bytes).unwrap();
    let prepared = repo.prepare_snapshot_file(&merged_path).unwrap();
    let commit = repo
        .commit_prepared(&prepared, "合并老师返稿", "张三", Some(&later), Some("main"), Some(&original.source), Some(&imported))
        .unwrap()
        .unwrap();
    let log = repo.history(Some("main")).unwrap();
    assert_eq!(log[0].id, commit);
    assert_eq!(log[0].parents, vec![later.clone(), imported.clone()]);
    // Now review-1 is merged and can be deleted.
    repo.delete_merged_branch("main", &commit, "review-1", &imported).unwrap();
    let after_merge = repo.import_review_file(&returned, &sent, None, "老师", "main", &commit).unwrap();
    assert!(after_merge.reused);
    assert_eq!(after_merge.commit, imported);
    assert_eq!(extract(&repo, &imported), returned_bytes);
}

#[test]
fn two_clients_collaborate_through_a_remote() {
    let sandbox = Sandbox::new("remote");
    let cancel = AtomicBool::new(false);
    let bare = sandbox.root.join("shared.git");
    std::fs::create_dir_all(&bare).unwrap();
    let setup = Repo::open(bare.clone(), None, None).unwrap();
    setup.git().invoke(None, None, &["init", "--bare", "--initial-branch=main"]).unwrap().checked().unwrap();
    let image = picture(5, 40_000);

    // Client A commits and pushes.
    let a_doc = sandbox.root.join("A").join("稿件.docx");
    std::fs::create_dir_all(a_doc.parent().unwrap()).unwrap();
    std::fs::write(&a_doc, fixture("A 的初稿", &image)).unwrap();
    let a = sandbox.repo(&a_doc);
    let a1 = a.commit_file(&a_doc, "初稿", "A", None, None, None, None).unwrap().unwrap();
    let bare_url = bare.to_string_lossy().into_owned();
    a.configure_remote(&bare_url, "main", true).unwrap();
    let state = a.push_remote(true, &cancel).unwrap();
    assert_eq!(state.remote_tip.as_deref(), Some(a1.as_str()));
    assert_eq!((state.ahead, state.behind), (0, 0));

    // Client B clones into a new document path.
    let b_doc = sandbox.root.join("B").join("稿件.docx");
    std::fs::create_dir_all(b_doc.parent().unwrap()).unwrap();
    let b = remote::clone_remote(&bare_url, "main", &b_doc, &sandbox.storage(), None, true, &cancel).unwrap();
    assert_eq!(std::fs::read(&b_doc).unwrap(), std::fs::read(&a_doc).unwrap());
    assert_eq!(b.head().unwrap().as_deref(), Some(a1.as_str()));
    assert!(remote::clone_remote(&bare_url, "main", &b_doc, &sandbox.storage(), None, true, &cancel).is_err());

    // B commits and pushes; A fetches and fast-forwards.
    std::fs::write(&b_doc, fixture("B 修改后", &image)).unwrap();
    let b2 = b.commit_file(&b_doc, "B 修改", "B", Some(&a1), None, None, None).unwrap().unwrap();
    b.push_remote(true, &cancel).unwrap();
    let state = a.fetch_remote(true, &cancel).unwrap();
    assert_eq!((state.ahead, state.behind), (0, 1));
    a.fast_forward_remote("main", &a1, &state.remote_reference, &b2).unwrap();
    assert_eq!(a.head().unwrap().as_deref(), Some(b2.as_str()));

    // Divergence: both commit; pushing A is refused, fetch reports ahead/behind.
    std::fs::write(&a_doc, fixture("A 再改", &image)).unwrap();
    let a3 = a.commit_file(&a_doc, "A 再改", "A", Some(&b2), None, None, None).unwrap().unwrap();
    std::fs::write(&b_doc, fixture("B 再改", &image)).unwrap();
    let b3 = b.commit_file(&b_doc, "B 再改", "B", Some(&b2), None, None, None).unwrap().unwrap();
    b.push_remote(true, &cancel).unwrap();
    assert!(a.push_remote(true, &cancel).is_err());
    let state = a.fetch_remote(true, &cancel).unwrap();
    assert_eq!((state.ahead, state.behind), (1, 1));
    assert_eq!(a.merge_base(&a3, &b3).unwrap(), b2);
    assert!(a.fast_forward_remote("main", &a3, &state.remote_reference, &b3).is_err());

    // A force-push by B hides b3 on the remote; A keeps it under a backup ref.
    b.git()
        .invoke(None, None, &["push", "--force", "origin", &format!("{b2}:refs/heads/main")])
        .unwrap()
        .checked()
        .unwrap();
    a.fetch_remote(true, &cancel).unwrap();
    assert_eq!(a.remote_backup_namespaces().unwrap().len(), 1);
    assert!(a.history(None).unwrap().iter().any(|r| r.id == b3), "rewritten commit must stay visible");

    // Cancellation is honoured before starting.
    let cancelled = AtomicBool::new(true);
    assert!(matches!(a.fetch_remote(true, &cancelled), Err(wordgit_core::Error::Cancelled)));
}

#[test]
fn replaces_working_copy_and_recovers_from_crashes() {
    let sandbox = Sandbox::new("replace");
    let document = sandbox.root.join("当前稿.docx");
    let image = picture(11, 10_000);
    let old = fixture("旧版", &image);
    let new = fixture("新版", &image);
    std::fs::write(&document, &old).unwrap();
    let repo = sandbox.repo(&document);
    let first = repo.commit_file(&document, "旧版", "x", None, None, None, None).unwrap().unwrap();
    let staged = sandbox.root.join("staged.docx");
    std::fs::write(&staged, &new).unwrap();
    let old_hash = history::sha256_hex(&old);
    let new_hash = history::sha256_hex(&new);

    // A failing publish rolls the document back and clears the journal.
    let result = review::replace_working_copy(&repo, &document, &old_hash, &staged, &new_hash, &mut || {
        Err(wordgit_core::Error::invalid("publish failed"))
    });
    assert!(result.is_err());
    assert_eq!(std::fs::read(&document).unwrap(), old);
    assert!(!journal::path_for(&repo).exists());

    // Success: file replaced, branch published, backup removed.
    review::replace_working_copy(&repo, &document, &old_hash, &staged, &new_hash, &mut || {
        repo.commit_file(&staged, "新版", "x", Some(&first), None, None, None).map(|_| ())
    })
    .unwrap();
    assert_eq!(std::fs::read(&document).unwrap(), new);
    let leftovers: Vec<_> = std::fs::read_dir(&sandbox.root)
        .unwrap()
        .flatten()
        .filter(|e| e.file_name().to_string_lossy().starts_with(".wordgit-restore-"))
        .collect();
    assert!(leftovers.is_empty());

    // Simulated crash after the swap but before publishing: recovery restores the original.
    let backup = sandbox.root.join(".wordgit-restore-test.tmp.bak");
    std::fs::write(&backup, &new).unwrap();
    std::fs::write(&document, &old).unwrap();
    journal::write(
        &repo,
        &journal::Entry {
            document: document.clone(),
            backup: backup.clone(),
            original_hash: new_hash.clone(),
            replacement_hash: old_hash.clone(),
            references: repo.reference_state().unwrap(),
        },
    )
    .unwrap();
    let paths = AppPaths::new(sandbox.root.clone());
    let messages = journal::recover_all(&paths, None);
    assert_eq!(messages.len(), 1, "{messages:?}");
    assert_eq!(std::fs::read(&document).unwrap(), new);
    assert!(!journal::path_for(&repo).exists());

    // A corrupt record is set aside with a message instead of blocking forever.
    std::fs::write(journal::path_for(&repo), b"garbage").unwrap();
    assert!(journal::recover(&repo).unwrap().is_some());
    assert!(!journal::path_for(&repo).exists());
}

#[test]
fn repository_location_is_stable() {
    let storage = PathBuf::from("/data/repositories");
    let a = history::repository_for(Path::new("/home/u/论文.docx"), &storage);
    let b = history::repository_for(Path::new("/home/u/./x/../论文.docx"), &storage);
    assert_eq!(a, b);
    assert_eq!(a.file_name().unwrap().len(), 64 + 4);
}

#[test]
fn locks_are_exclusive_per_repository() {
    let sandbox = Sandbox::new("lock");
    let repo = sandbox.repo(&sandbox.root.join("a.docx"));
    let held = repo.lock().unwrap();
    assert!(repo.lock().is_err());
    drop(held);
    repo.lock().unwrap();
}
