//! Runs the real `wordgit serve` binary and exercises the HTTP API: security checks (host, origin,
//! token) and a commit/branch/switch round trip through the same endpoints the add-in uses.

use std::io::{Read, Write};
use std::net::TcpStream;
use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

struct Service {
    child: Child,
    home: PathBuf,
    port: u16,
    token: String,
}

impl Drop for Service {
    fn drop(&mut self) {
        let _ = self.child.kill();
        let _ = self.child.wait();
        let _ = std::fs::remove_dir_all(&self.home);
    }
}

fn free_port() -> u16 {
    std::net::TcpListener::bind("127.0.0.1:0").unwrap().local_addr().unwrap().port()
}

fn start() -> Service {
    let home = std::env::temp_dir().join(format!("wordgit-svc-{}", std::process::id() as u64 * 1000 + free_port() as u64));
    std::fs::create_dir_all(&home).unwrap();
    let port = free_port();
    let child = Command::new(env!("CARGO_BIN_EXE_wordgit"))
        .args(["serve", "--port", &port.to_string()])
        .env("WORDGIT_HOME", &home)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .unwrap();
    let service_file = home.join("service.json");
    let started = Instant::now();
    while !service_file.exists() {
        assert!(started.elapsed() < Duration::from_secs(20), "service did not start");
        std::thread::sleep(Duration::from_millis(50));
    }
    std::thread::sleep(Duration::from_millis(100));
    let info: serde_json::Value = serde_json::from_slice(&std::fs::read(&service_file).unwrap()).unwrap();
    Service { child, home, port, token: info["token"].as_str().unwrap().to_string() }
}

struct Reply {
    status: u16,
    headers: String,
    body: serde_json::Value,
}

fn request(port: u16, method: &str, path: &str, headers: &[(&str, &str)], body: &str) -> Reply {
    let mut stream = TcpStream::connect(("127.0.0.1", port)).unwrap();
    stream.set_read_timeout(Some(Duration::from_secs(60))).unwrap();
    let mut text = format!("{method} {path} HTTP/1.1\r\nConnection: close\r\nContent-Length: {}\r\n", body.len());
    if !headers.iter().any(|(k, _)| k.eq_ignore_ascii_case("host")) {
        text.push_str(&format!("Host: 127.0.0.1:{port}\r\n"));
    }
    for (key, value) in headers {
        text.push_str(&format!("{key}: {value}\r\n"));
    }
    text.push_str("\r\n");
    text.push_str(body);
    stream.write_all(text.as_bytes()).unwrap();
    let mut response = Vec::new();
    stream.read_to_end(&mut response).unwrap();
    let response = String::from_utf8_lossy(&response).into_owned();
    let (head, body) = response.split_once("\r\n\r\n").unwrap_or((&response, ""));
    let status = head.split(' ').nth(1).unwrap().parse().unwrap();
    let body = serde_json::from_str(body).unwrap_or(serde_json::Value::Null);
    Reply { status, headers: head.to_lowercase(), body }
}

fn api(service: &Service, endpoint: &str, body: serde_json::Value) -> Reply {
    request(
        service.port,
        "POST",
        &format!("/api/v1/{endpoint}"),
        &[("Content-Type", "application/json"), ("X-WordGit-Token", &service.token), ("Origin", "file://")],
        &body.to_string(),
    )
}

/// Minimal valid DOCX (stored members) built without extra crates.
fn docx(text: &str) -> Vec<u8> {
    let parts: [(&str, String); 2] =
        [("[Content_Types].xml", "<Types/>".into()), ("word/document.xml", format!("<w:document><w:t>{text}</w:t></w:document>"))];
    let mut out = Vec::new();
    let mut central = Vec::new();
    for (name, body) in &parts {
        let offset = out.len() as u32;
        let crc = crc32(body.as_bytes());
        let mut header = Vec::new();
        header.extend_from_slice(&0x0403_4b50u32.to_le_bytes());
        header.extend_from_slice(&[20, 0, 0, 8, 0, 0, 0, 0, 0x21, 0]);
        header.extend_from_slice(&crc.to_le_bytes());
        header.extend_from_slice(&(body.len() as u32).to_le_bytes());
        header.extend_from_slice(&(body.len() as u32).to_le_bytes());
        header.extend_from_slice(&(name.len() as u16).to_le_bytes());
        header.extend_from_slice(&0u16.to_le_bytes());
        header.extend_from_slice(name.as_bytes());
        out.extend_from_slice(&header);
        out.extend_from_slice(body.as_bytes());
        central.extend_from_slice(&0x0201_4b50u32.to_le_bytes());
        central.extend_from_slice(&[20, 0, 20, 0, 0, 8, 0, 0, 0, 0, 0x21, 0]);
        central.extend_from_slice(&crc.to_le_bytes());
        central.extend_from_slice(&(body.len() as u32).to_le_bytes());
        central.extend_from_slice(&(body.len() as u32).to_le_bytes());
        central.extend_from_slice(&(name.len() as u16).to_le_bytes());
        central.extend_from_slice(&[0u8; 12]);
        central.extend_from_slice(&offset.to_le_bytes());
        central.extend_from_slice(name.as_bytes());
    }
    let directory = out.len() as u32;
    out.extend_from_slice(&central);
    out.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
    out.extend_from_slice(&[0, 0, 0, 0, 2, 0, 2, 0]);
    out.extend_from_slice(&(central.len() as u32).to_le_bytes());
    out.extend_from_slice(&directory.to_le_bytes());
    out.extend_from_slice(&[0, 0]);
    out
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for &byte in bytes {
        crc ^= byte as u32;
        for _ in 0..8 {
            crc = if crc & 1 != 0 { (crc >> 1) ^ 0xEDB8_8320 } else { crc >> 1 };
        }
    }
    !crc
}

#[test]
fn security_checks() {
    let service = start();
    let port = service.port;
    assert_eq!(request(port, "GET", "/api/v1/ping", &[], "").status, 200);
    // DNS rebinding: wrong Host is refused even with a valid token.
    let rebinding = request(port, "POST", "/api/v1/info", &[("Host", "evil.example"), ("X-WordGit-Token", &service.token)], "{}");
    assert_eq!(rebinding.status, 421);
    // Foreign web pages are refused.
    let foreign = request(port, "POST", "/api/v1/info", &[("Origin", "https://evil.example"), ("X-WordGit-Token", &service.token)], "{}");
    assert_eq!(foreign.status, 403);
    // Missing or wrong token.
    assert_eq!(request(port, "POST", "/api/v1/info", &[("Origin", "file://")], "{}").status, 401);
    assert_eq!(request(port, "POST", "/api/v1/info", &[("X-WordGit-Token", "wrong")], "{}").status, 401);
    // The add-in's preflight from file:// pages succeeds with CORS headers.
    let preflight = request(port, "OPTIONS", "/api/v1/status", &[("Origin", "file://"), ("Access-Control-Request-Method", "POST")], "");
    assert_eq!(preflight.status, 204);
    assert!(preflight.headers.contains("access-control-allow-origin: file://"));
    let ok = api(&service, "info", serde_json::json!({}));
    assert_eq!(ok.status, 200);
    assert!(ok.headers.contains("access-control-allow-origin: file://"));
    // Service file is private on Unix.
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        let mode = std::fs::metadata(service.home.join("service.json")).unwrap().permissions().mode();
        assert_eq!(mode & 0o077, 0);
    }
}

#[test]
fn commit_branch_switch_round_trip() {
    let service = start();
    let folder = service.home.join("docs");
    std::fs::create_dir_all(&folder).unwrap();
    let document = folder.join("论文.docx");
    std::fs::write(&document, docx("第一版")).unwrap();
    let path = document.to_string_lossy().into_owned();

    let status = api(&service, "status", serde_json::json!({ "document": path }));
    assert_eq!(status.status, 200, "{:?}", status.body);
    assert_eq!(status.body["workingState"], "noCommit");

    let first = api(&service, "commit", serde_json::json!({ "document": path, "message": "初稿", "author": "测试" }));
    assert_eq!(first.status, 200, "{:?}", first.body);
    let first_id = first.body["commit"].as_str().unwrap().to_string();

    std::fs::write(&document, docx("第二版")).unwrap();
    let stale = api(&service, "commit", serde_json::json!({ "document": path, "message": "x", "expectedHead": "0".repeat(40) }));
    assert_eq!(stale.status, 400);
    assert!(stale.body["error"].as_str().unwrap().contains("已变化"));
    let second = api(&service, "commit", serde_json::json!({ "document": path, "message": "第二版", "author": "测试" }));
    let second_id = second.body["commit"].as_str().unwrap().to_string();

    let created = api(&service, "branch/create", serde_json::json!({ "document": path, "name": "review", "revision": first_id }));
    assert_eq!(created.status, 200, "{:?}", created.body);

    let prepared = api(&service, "branch/switch", serde_json::json!({ "document": path, "name": "review" }));
    assert_eq!(prepared.status, 200, "{:?}", prepared.body);
    assert_eq!(prepared.body["changesFile"], true);
    let applied = api(&service, "plan/apply", serde_json::json!({ "plan": prepared.body["plan"] }));
    assert_eq!(applied.status, 200, "{:?}", applied.body);
    assert_eq!(std::fs::read(&document).unwrap(), docx("第一版"));
    // A plan is single-use.
    assert_eq!(api(&service, "plan/apply", serde_json::json!({ "plan": prepared.body["plan"] })).status, 400);

    let status = api(&service, "status", serde_json::json!({ "document": path }));
    assert_eq!(status.body["branch"], "review");
    assert_eq!(status.body["workingState"], "clean");
    assert_eq!(status.body["graph"].as_array().unwrap().len(), 2);

    // Restore the second version onto review: commits the current draft first, then a new commit.
    let restore = api(&service, "version/restore", serde_json::json!({ "document": path, "revision": second_id, "author": "测试" }));
    assert_eq!(restore.status, 200, "{:?}", restore.body);
    let applied = api(&service, "plan/apply", serde_json::json!({ "plan": restore.body["plan"] }));
    assert_eq!(applied.status, 200, "{:?}", applied.body);
    assert_eq!(std::fs::read(&document).unwrap(), docx("第二版"));

    // Preview and compare produce session files.
    let preview = api(&service, "version/preview", serde_json::json!({ "document": path, "revision": first_id }));
    let file = PathBuf::from(preview.body["file"].as_str().unwrap());
    assert_eq!(std::fs::read(&file).unwrap(), docx("第一版"));
    let preview_status = api(&service, "status", serde_json::json!({ "document": file.to_string_lossy() }));
    assert_eq!(preview_status.body["session"], true);

    // Malformed requests get a readable error.
    let bad = api(&service, "commit", serde_json::json!({ "message": "缺少文档" }));
    assert_eq!(bad.status, 400);
    assert!(bad.body["error"].as_str().unwrap().contains("请求格式错误"));
}

#[test]
fn moves_history_without_invalidating_review_sessions() {
    use serde_json::json;
    use std::fs;

    let service = start();
    let document = service.home.join("原稿.docx");
    let old_content = docx("初稿");
    fs::write(&document, &old_content).unwrap();
    let first = api(&service, "commit", json!({ "document": document, "message": "初稿" }));
    assert_eq!(first.status, 200, "{:?}", first.body);
    let first_id = &first.body["commit"];
    fs::write(&document, docx("新稿")).unwrap();
    let latest = api(&service, "commit", json!({ "document": document, "message": "新稿" }));
    assert_eq!(latest.status, 200, "{:?}", latest.body);
    let source = service.home.join("repositories");
    let subdirectory = source.join("错误");
    assert_eq!(api(&service, "settings/data-directory", json!({ "directory": subdirectory })).status, 400);
    assert!(!subdirectory.exists());

    let preview = api(&service, "version/preview", json!({ "document": document, "revision": first_id }));
    let preview_path = preview.body["file"].as_str().unwrap();
    let target = service.home.join("新的历史目录");
    fs::create_dir_all(target.join("repositories")).unwrap();
    let sentinel = target.join("repositories/keep.txt");
    fs::write(&sentinel, b"existing history").unwrap();
    assert_eq!(api(&service, "settings/data-directory", json!({ "directory": target })).status, 400);
    assert_eq!(fs::read(&sentinel).unwrap(), b"existing history");
    assert!(source.is_dir());
    fs::remove_file(&sentinel).unwrap();

    let pending = api(&service, "version/restore", json!({ "document": document, "revision": first_id }));
    assert_eq!(pending.status, 200, "{:?}", pending.body);
    assert_eq!(api(&service, "settings/data-directory", json!({ "directory": target })).status, 400);
    assert_eq!(api(&service, "plan/cancel", json!({ "plan": pending.body["plan"] })).status, 200);

    let moved = api(&service, "settings/data-directory", json!({ "directory": target }));
    assert_eq!(moved.status, 200, "{:?}", moved.body);
    assert_eq!(moved.body["dataDirectory"], json!(target));
    assert!(!source.exists());
    let status = api(&service, "status", json!({ "document": document }));
    assert_eq!(status.body["head"], latest.body["commit"]);
    assert_eq!(api(&service, "status", json!({ "document": preview_path })).body["session"], true);
    assert_eq!(fs::read(preview_path).unwrap(), old_content);

    let cli = Command::new(env!("CARGO_BIN_EXE_wordgit"))
        .arg("data-directory")
        .env("WORDGIT_HOME", &service.home)
        .output()
        .unwrap();
    assert!(cli.status.success());
    assert_eq!(String::from_utf8_lossy(&cli.stdout).trim(), target.to_string_lossy());
    let offline = service.home.join("离线磁盘");
    fs::rename(&target, &offline).unwrap();
    assert_eq!(api(&service, "status", json!({ "document": document })).status, 400);
    assert!(!target.exists());
    fs::rename(offline, &target).unwrap();

    let restore = api(&service, "version/restore", json!({ "document": document, "revision": first_id }));
    let applied = api(&service, "plan/apply", json!({ "plan": restore.body["plan"] }));
    assert_eq!(applied.status, 200, "{:?}", applied.body);
    assert_eq!(fs::read(&document).unwrap(), old_content);
    let returned = api(&service, "settings/data-directory", json!({ "directory": service.home }));
    assert_eq!(returned.status, 200, "{:?}", returned.body);
    assert!(source.is_dir());
    assert_eq!(api(&service, "status", json!({ "document": document })).body["workingState"], "clean");
}
