//! `wordgit` — local service and command line for WordGit.
//!
//! ```text
//! wordgit serve [--port N]            run the service the WPS add-in talks to
//! wordgit install [--addin DIR]       deploy the WPS add-in, enable autostart, start the service
//! wordgit uninstall                   remove the add-in and autostart (history is kept)
//! wordgit call <endpoint> [JSON]      run one API operation in-process and print the JSON result
//! wordgit status|log <doc>            shortcuts for common calls
//! wordgit commit <doc> -m <message>
//! wordgit version
//! ```

mod collaboration;
mod data_directory;
mod install;
mod ops;
mod service;

use ops::Context;
use serde_json::{json, Value};
use std::path::PathBuf;
use std::process::ExitCode;
use std::sync::Arc;
use std::time::Duration;
use wordgit_core::maintenance::Maintenance;
use wordgit_core::paths::AppPaths;

const IDLE_FOR_MAINTENANCE: Duration = Duration::from_secs(20);

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let command = args.first().map(String::as_str).unwrap_or("help");
    let result = match command {
        "serve" => serve(&args[1..]),
        "install" => install(&args[1..]),
        "uninstall" => uninstall(),
        "data-directory" => configure_directory(&args[1..]),
        "call" => call(&args[1..]),
        "status" | "log" | "commit" => shortcut(command, &args[1..]),
        "version" | "--version" | "-V" => {
            println!("wordgit {}", env!("CARGO_PKG_VERSION"));
            Ok(())
        }
        _ => {
            println!("{}", include_str!("usage.txt"));
            Ok(())
        }
    };
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(message) => {
            eprintln!("{message}");
            ExitCode::FAILURE
        }
    }
}

fn option(args: &[String], name: &str) -> Option<String> {
    args.iter().position(|a| a == name).and_then(|i| args.get(i + 1)).cloned()
}

fn context(paths: AppPaths) -> Result<Arc<Context>, String> {
    let paths = data_directory::load(paths).map_err(|error| error.to_string())?;
    let git = wordgit_core::git::Git::find().map_err(|e| e.to_string())?;
    Ok(make_context(paths, Some(git)))
}

fn make_context(paths: AppPaths, git: Option<PathBuf>) -> Arc<Context> {
    let maintenance = Maintenance::new(paths.clone(), git.clone(), IDLE_FOR_MAINTENANCE);
    Context::new(paths, git, maintenance)
}

fn service_lock(paths: &AppPaths) -> Result<wordgit_core::fsutil::FileLock, String> {
    wordgit_core::fsutil::try_lock(&paths.root.join("service.lock")).map_err(|error| error.to_string())?
        .ok_or_else(|| "WordGit 服务正在使用此数据目录。".into())
}

fn configure_directory(args: &[String]) -> Result<(), String> {
    let base = AppPaths::from_environment();
    if let Some(target) = args.first() {
        let _service = service_lock(&base)?;
        if service::service_running(service::DEFAULT_PORT) {
            return Err("WordGit 服务正在运行。请使用功能区的「数据目录」更改位置。".into());
        }
        let paths = data_directory::load(base).map_err(|error| error.to_string())?;
        let moved = data_directory::migrate(&paths, std::path::Path::new(target)).map_err(|error| error.to_string())?;
        println!("{}\n{}", moved.paths.data_directory().display(), moved.message);
    } else {
        let paths = data_directory::load(base).map_err(|error| error.to_string())?;
        println!("{}", paths.data_directory().display());
    }
    Ok(())
}

fn serve(args: &[String]) -> Result<(), String> {
    let port: u16 = option(args, "--port").map(|p| p.parse().map_err(|_| "无效端口")).transpose()?.unwrap_or(service::DEFAULT_PORT);
    if service::service_running(port) {
        println!("WordGit 服务已在运行（127.0.0.1:{port}）。");
        return Ok(());
    }
    let base = AppPaths::from_environment();
    let _service = service_lock(&base)?;
    let ctx = context(base)?;
    ctx.startup();
    let token = wordgit_core::random_hex(24);
    let runtime = tokio::runtime::Builder::new_multi_thread().enable_all().build().map_err(|e| e.to_string())?;
    let result = runtime.block_on(async {
        #[cfg(unix)]
        let mut terminate = tokio::signal::unix::signal(tokio::signal::unix::SignalKind::terminate()).map_err(|error| error.to_string())?;
        let listener = tokio::net::TcpListener::bind(("127.0.0.1", port))
            .await
            .map_err(|e| format!("无法监听 127.0.0.1:{port}：{e}"))?;
        let service_file = ctx.paths.root.join("service.json");
        service::write_service_file(&service_file, port, &token).map_err(|e| e.to_string())?;
        eprintln!("WordGit 服务 {} 监听 http://127.0.0.1:{port}，数据目录 {}", env!("CARGO_PKG_VERSION"), ctx.paths.data_directory().display());
        let state = service::AppState::new(ctx, token, port);
        let app = service::router(state);
        let shutdown = async move {
            #[cfg(unix)]
            tokio::select! {
                _ = tokio::signal::ctrl_c() => {},
                _ = terminate.recv() => {},
            }
            #[cfg(not(unix))]
            let _ = tokio::signal::ctrl_c().await;
        };
        let served = axum::serve(listener, app).with_graceful_shutdown(shutdown).await.map_err(|e| e.to_string());
        wordgit_core::fsutil::try_remove(&service_file);
        served
    });
    result
}

fn install(args: &[String]) -> Result<(), String> {
    let paths = AppPaths::from_environment();
    let source = option(args, "--addin").map(PathBuf::from).or_else(install::bundled_addin).ok_or("找不到加载项文件，请用 --addin 指定。")?;
    let autostart = !args.iter().any(|a| a == "--no-autostart");
    #[cfg(target_os = "linux")]
    let _installation = {
        if !source.join("ribbon.xml").is_file() {
            return Err(format!("找不到加载项文件：{}", source.display()));
        }
        let lock = install::linux::lock_installation(&paths).map_err(|error| error.to_string())?;
        if autostart {
            install::linux::stop_service(&paths).and_then(|()| install::linux::ensure_default_port()).map_err(|error| error.to_string())?;
        }
        lock
    };
    if let Some(directory) = option(args, "--data-directory") {
        configure_directory(&[directory])?;
    }
    let folder = install::install_addin(&source, &paths, env!("CARGO_PKG_VERSION")).map_err(|e| e.to_string())?;
    println!("已安装 WPS 加载项：{}", folder.display());
    if autostart {
        let executable = std::env::current_exe().map_err(|e| e.to_string())?;
        #[cfg(target_os = "linux")]
        let executable = install::linux::install_executable(&executable, &paths).map_err(|error| error.to_string())?;
        println!("{}", install::enable_autostart(&executable).map_err(|e| e.to_string())?);
        if cfg!(target_os = "linux") || !service::service_running(service::DEFAULT_PORT) {
            let mut command = std::process::Command::new(&executable);
            command.arg("serve").stdin(std::process::Stdio::null()).stdout(std::process::Stdio::null()).stderr(std::process::Stdio::null());
            #[cfg(windows)]
            {
                use std::os::windows::process::CommandExt;
                command.creation_flags(0x08000000); // CREATE_NO_WINDOW
            }
            #[cfg(unix)]
            {
                use std::os::unix::process::CommandExt;
                command.process_group(0);
            }
            let _child = command.spawn().map_err(|e| format!("无法启动服务：{e}"))?;
            #[cfg(target_os = "linux")]
            {
                let mut child = _child;
                install::linux::wait_for_service(&mut child, &paths).map_err(|error| error.to_string())?;
            }
            println!("WordGit 服务已启动。");
        }
    }
    println!("请重新启动 WPS 文字。首次载入时 WPS 会询问是否信任 WordGit 加载项，请选择「确定」。");
    Ok(())
}

fn uninstall() -> Result<(), String> {
    #[cfg(target_os = "linux")]
    let (paths, _installation) = {
        let paths = AppPaths::from_environment();
        let lock = install::linux::lock_installation(&paths).map_err(|error| error.to_string())?;
        install::linux::stop_service(&paths).map_err(|error| error.to_string())?;
        (paths, lock)
    };
    install::uninstall_addin().map_err(|e| e.to_string())?;
    install::disable_autostart().map_err(|e| e.to_string())?;
    #[cfg(target_os = "linux")]
    install::linux::uninstall_executable(&paths).map_err(|error| error.to_string())?;
    println!("已移除 WPS 加载项与自启动。稿件历史（repositories）保留不动。");
    Ok(())
}

/// Runs one operation in-process (no HTTP), printing the JSON result. Used by tests and scripts.
pub fn dispatch(ctx: &Context, endpoint: &str, body: Value) -> Result<Value, String> {
    fn parse<T: serde::de::DeserializeOwned>(body: Value) -> Result<T, String> {
        serde_json::from_value(body).map_err(|e| format!("请求格式错误：{e}"))
    }
    let result = match endpoint.trim_start_matches('/').trim_start_matches("api/v1/") {
        "info" | "ping" => Ok(ctx.ping()),
        "notices" => Ok(ctx.take_notices()),
        "status" => ctx.status(&parse(body)?),
        "history" => ctx.history(&parse(body)?),
        "commit" => ctx.commit(&parse(body)?),
        "branch/create" => ctx.create_branch(&parse(body)?),
        "branch/deletable" => ctx.deletable_branches(&parse(body)?),
        "branch/delete" => ctx.delete_branch(&parse(body)?),
        "branch/switch" => ctx.switch_prepare(&parse(body)?),
        "version/preview" => ctx.preview(&parse(body)?),
        "version/export" => ctx.export(&parse(body)?),
        "version/compare" => ctx.compare_files(&parse(body)?),
        "version/restore" => ctx.restore_prepare(&parse(body)?),
        "merge/collaboration" => collaboration::plan(&parse(body)?),
        "merge/branch" => ctx.merge_branch_plan(&parse(body)?),
        "merge/versions" => ctx.merge_versions_plan(&parse(body)?),
        "merge/register" => ctx.register_review(&parse(body)?),
        "merge/finish" => ctx.finish_merge_prepare(&parse(body)?),
        "review/import" => ctx.import_review(&parse(body)?),
        "plan/apply" => ctx.apply_plan(&parse(body)?),
        "plan/cancel" => Ok(ctx.cancel_plan(&parse(body)?)),
        "remote/link" => ctx.remote_link(&parse(body)?),
        "remote/fetch" => ctx.remote_fetch(&parse(body)?),
        "remote/push" => ctx.remote_push(&parse(body)?),
        "remote/pull" => ctx.remote_pull(&parse(body)?),
        "remote/clone" => ctx.remote_clone(&parse(body)?),
        other => return Err(format!("未知操作：{other}")),
    };
    result.map_err(|e| e.to_string())
}

fn call(args: &[String]) -> Result<(), String> {
    let endpoint = args.first().ok_or("用法：wordgit call <endpoint> [JSON]")?;
    let body: Value = match args.get(1) {
        Some(text) => serde_json::from_str(text).map_err(|e| format!("JSON 无效：{e}"))?,
        None => json!({}),
    };
    let base = AppPaths::from_environment();
    let _access = data_directory::access_lock(&base, false).map_err(|error| error.to_string())?;
    let ctx = context(base)?;
    if endpoint.trim_start_matches('/').trim_start_matches("api/v1/") != "info" {
        data_directory::ensure_access(&ctx.paths).map_err(|error| error.to_string())?;
    }
    let value = dispatch(&ctx, endpoint, body)?;
    println!("{}", serde_json::to_string_pretty(&value).unwrap_or_default());
    Ok(())
}

fn shortcut(command: &str, args: &[String]) -> Result<(), String> {
    let document = args.first().ok_or("请指定 .docx 文档路径。")?;
    let base = AppPaths::from_environment();
    let _access = data_directory::access_lock(&base, false).map_err(|error| error.to_string())?;
    let ctx = context(base)?;
    data_directory::ensure_access(&ctx.paths).map_err(|error| error.to_string())?;
    let value = match command {
        "status" => dispatch(&ctx, "status", json!({ "document": document }))?,
        "log" => dispatch(&ctx, "history", json!({ "document": document }))?,
        _ => {
            let message = option(args, "-m").ok_or("请用 -m 填写提交说明。")?;
            let author = option(args, "--author").unwrap_or_else(|| std::env::var("USER").or_else(|_| std::env::var("USERNAME")).unwrap_or_default());
            dispatch(&ctx, "commit", json!({ "document": document, "message": message, "author": author }))?
        }
    };
    if command == "log" {
        for revision in value["history"].as_array().into_iter().flatten() {
            println!(
                "{}  {}  {}  · {}",
                &revision["id"].as_str().unwrap_or("")[..8.min(revision["id"].as_str().unwrap_or("").len())],
                revision["date"].as_str().unwrap_or(""),
                revision["subject"].as_str().unwrap_or(""),
                revision["author"].as_str().unwrap_or("")
            );
        }
    } else {
        println!("{}", serde_json::to_string_pretty(&value).unwrap_or_default());
    }
    Ok(())
}
