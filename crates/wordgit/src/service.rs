//! The local HTTP service the editor add-in talks to.
//!
//! Security: it listens on 127.0.0.1 only, answers only requests addressed to loopback (DNS
//! rebinding), accepts browser requests only from the add-in's own `file://` pages, and requires a
//! random token on every call except `ping`. The token lives in `service.json`, readable only by
//! the user; the add-in reads it through WPS's file API, which ordinary web pages do not have.

use crate::ops::{self, Context};
use axum::body::Body;
use axum::extract::{Request, State};
use axum::http::{header, HeaderMap, HeaderValue, Method, StatusCode};
use axum::middleware::{self, Next};
use axum::response::{IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use serde::de::DeserializeOwned;
use serde_json::{json, Value};
use std::path::Path;
use std::sync::Arc;
use wordgit_core::Error;

pub const DEFAULT_PORT: u16 = 47821;
pub const TOKEN_HEADER: &str = "x-wordgit-token";

#[derive(Clone)]
pub struct AppState {
    pub ctx: Arc<Context>,
    pub token: String,
    pub port: u16,
}

fn error_response(error: &Error) -> Response {
    let status = match error {
        Error::Invalid(_) => StatusCode::BAD_REQUEST,
        Error::Data(_) => StatusCode::UNPROCESSABLE_ENTITY,
        Error::Timeout(_) => StatusCode::GATEWAY_TIMEOUT,
        Error::Cancelled => StatusCode::CONFLICT,
        Error::Io(_) => StatusCode::INTERNAL_SERVER_ERROR,
    };
    (status, Json(json!({ "error": error.to_string(), "kind": error.kind() }))).into_response()
}

/// Runs a blocking operation off the async runtime.
async fn blocking<T, F>(state: AppState, body: Value, run: F) -> Response
where
    T: DeserializeOwned + Send + 'static,
    F: FnOnce(&Context, T) -> wordgit_core::Result<Value> + Send + 'static,
{
    let request: T = match serde_json::from_value(body) {
        Ok(request) => request,
        Err(error) => {
            return (StatusCode::BAD_REQUEST, Json(json!({ "error": format!("请求格式错误：{error}"), "kind": "request" })))
                .into_response()
        }
    };
    let ctx = Arc::clone(&state.ctx);
    ctx.maintenance.touch();
    let outcome = tokio::task::spawn_blocking(move || {
        let result = run(&ctx, request);
        ctx.maintenance.touch();
        if let Err(error) = &result {
            if matches!(error, Error::Io(_) | Error::Data(_)) {
                ctx.paths.log_error(&format!("{}: {error}", error.kind()));
            }
        }
        result
    })
    .await;
    match outcome {
        Ok(Ok(value)) => Json(value).into_response(),
        Ok(Err(error)) => error_response(&error),
        Err(panic) => {
            state.ctx.paths.log_error(&format!("operation panicked: {panic}"));
            error_response(&Error::invalid("WordGit 服务内部错误，详情见日志。"))
        }
    }
}

macro_rules! route {
    ($router:ident, $path:literal, $method:path) => {
        $router = $router.route(
            $path,
            post(|State(state): State<AppState>, Json(body): Json<Value>| async move {
                blocking(state, body, |ctx, request| $method(ctx, &request)).await
            }),
        );
    };
}

pub fn router(state: AppState) -> Router {
    let mut api = Router::new()
        .route("/api/v1/ping", get(ping))
        .route(
            "/api/v1/info",
            post(|State(state): State<AppState>| async move { Json(state.ctx.ping()) }),
        )
        .route(
            "/api/v1/notices",
            post(|State(state): State<AppState>| async move { Json(state.ctx.take_notices()) }),
        )
        .route(
            "/api/v1/cancel",
            post(|State(state): State<AppState>, Json(body): Json<ops::CancelRequest>| async move {
                Json(state.ctx.cancel(&body))
            }),
        )
        .route(
            "/api/v1/plan/cancel",
            post(|State(state): State<AppState>, Json(body): Json<ops::PlanRequest>| async move {
                Json(state.ctx.cancel_plan(&body))
            }),
        );
    route!(api, "/api/v1/status", Context::status);
    route!(api, "/api/v1/history", Context::history);
    route!(api, "/api/v1/commit", Context::commit);
    route!(api, "/api/v1/branch/create", Context::create_branch);
    route!(api, "/api/v1/branch/deletable", Context::deletable_branches);
    route!(api, "/api/v1/branch/delete", Context::delete_branch);
    route!(api, "/api/v1/branch/switch", Context::switch_prepare);
    route!(api, "/api/v1/version/preview", Context::preview);
    route!(api, "/api/v1/version/export", Context::export);
    route!(api, "/api/v1/version/compare", Context::compare_files);
    route!(api, "/api/v1/version/restore", Context::restore_prepare);
    route!(api, "/api/v1/merge/branch", Context::merge_branch_plan);
    route!(api, "/api/v1/merge/versions", Context::merge_versions_plan);
    route!(api, "/api/v1/merge/register", Context::register_review);
    route!(api, "/api/v1/merge/finish", Context::finish_merge_prepare);
    route!(api, "/api/v1/review/import", Context::import_review);
    route!(api, "/api/v1/plan/apply", Context::apply_plan);
    route!(api, "/api/v1/remote/link", Context::remote_link);
    route!(api, "/api/v1/remote/fetch", Context::remote_fetch);
    route!(api, "/api/v1/remote/push", Context::remote_push);
    route!(api, "/api/v1/remote/pull", Context::remote_pull);
    route!(api, "/api/v1/remote/clone", Context::remote_clone);
    api.layer(middleware::from_fn_with_state(state.clone(), guard)).with_state(state)
}

async fn ping() -> Json<Value> {
    Json(json!({ "ok": true, "service": "wordgit", "version": env!("CARGO_PKG_VERSION"), "api": 1 }))
}

/// The add-in's pages load from `file://`; Chromium reports their origin as `file://` or `null`.
fn origin_allowed(origin: &str) -> bool {
    origin == "file://" || origin == "null"
}

fn cors(headers: &mut HeaderMap, origin: Option<&HeaderValue>) {
    if let Some(origin) = origin {
        headers.insert(header::ACCESS_CONTROL_ALLOW_ORIGIN, origin.clone());
        headers.insert(header::VARY, HeaderValue::from_static("Origin"));
        headers.insert(header::ACCESS_CONTROL_ALLOW_METHODS, HeaderValue::from_static("GET, POST, OPTIONS"));
        headers.insert(header::ACCESS_CONTROL_ALLOW_HEADERS, HeaderValue::from_static("Content-Type, X-WordGit-Token"));
        headers.insert("access-control-allow-private-network", HeaderValue::from_static("true"));
        headers.insert(header::ACCESS_CONTROL_MAX_AGE, HeaderValue::from_static("600"));
    }
}

fn refuse(status: StatusCode, message: &str) -> Response {
    (status, Json(json!({ "error": message, "kind": "forbidden" }))).into_response()
}

async fn guard(State(state): State<AppState>, request: Request, next: Next) -> Response {
    let host = request.headers().get(header::HOST).and_then(|h| h.to_str().ok()).unwrap_or("");
    let expected = [format!("127.0.0.1:{}", state.port), format!("localhost:{}", state.port)];
    if !expected.iter().any(|e| e.eq_ignore_ascii_case(host)) {
        return refuse(StatusCode::MISDIRECTED_REQUEST, "unexpected host");
    }
    let origin = request.headers().get(header::ORIGIN).cloned();
    if let Some(origin) = &origin {
        if !origin_allowed(origin.to_str().unwrap_or("")) {
            return refuse(StatusCode::FORBIDDEN, "origin not allowed");
        }
    }
    if request.method() == Method::OPTIONS {
        let mut response = Response::new(Body::empty());
        *response.status_mut() = StatusCode::NO_CONTENT;
        cors(response.headers_mut(), origin.as_ref());
        return response;
    }
    let public = request.uri().path() == "/api/v1/ping";
    let token_ok = request
        .headers()
        .get(TOKEN_HEADER)
        .and_then(|t| t.to_str().ok())
        .is_some_and(|t| constant_time_eq(t.as_bytes(), state.token.as_bytes()));
    let mut response = if public || token_ok {
        next.run(request).await
    } else {
        refuse(StatusCode::UNAUTHORIZED, "WordGit 服务口令不正确，请重新启动编辑器。")
    };
    response.headers_mut().insert(header::CACHE_CONTROL, HeaderValue::from_static("no-store"));
    cors(response.headers_mut(), origin.as_ref());
    response
}

fn constant_time_eq(a: &[u8], b: &[u8]) -> bool {
    a.len() == b.len() && a.iter().zip(b).fold(0u8, |acc, (x, y)| acc | (x ^ y)) == 0
}

/// `service.json` tells the add-in where the service listens and the token it must send.
pub fn write_service_file(path: &Path, port: u16, token: &str) -> std::io::Result<()> {
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    let text = serde_json::to_string_pretty(&json!({
        "port": port,
        "token": token,
        "pid": std::process::id(),
        "version": env!("CARGO_PKG_VERSION"),
    }))
    .unwrap_or_default();
    let temporary = wordgit_core::fsutil::sibling(path, ".tmp");
    {
        let mut options = std::fs::OpenOptions::new();
        options.write(true).create(true).truncate(true);
        #[cfg(unix)]
        {
            use std::os::unix::fs::OpenOptionsExt;
            options.mode(0o600);
        }
        use std::io::Write;
        let mut file = options.open(&temporary)?;
        file.write_all(text.as_bytes())?;
        file.sync_all()?;
    }
    std::fs::rename(&temporary, path)
}

/// Whether a WordGit service already answers on `port`.
pub fn service_running(port: u16) -> bool {
    use std::io::{Read, Write};
    let Ok(mut stream) = std::net::TcpStream::connect_timeout(
        &std::net::SocketAddr::from(([127, 0, 0, 1], port)),
        std::time::Duration::from_millis(500),
    ) else {
        return false;
    };
    let _ = stream.set_read_timeout(Some(std::time::Duration::from_secs(2)));
    let request = format!("GET /api/v1/ping HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\nConnection: close\r\n\r\n");
    if stream.write_all(request.as_bytes()).is_err() {
        return false;
    }
    let mut response = String::new();
    let _ = stream.read_to_string(&mut response);
    response.contains("\"service\":\"wordgit\"")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn origins_and_tokens() {
        assert!(origin_allowed("file://"));
        assert!(origin_allowed("null"));
        assert!(!origin_allowed("http://evil.example"));
        assert!(constant_time_eq(b"abc", b"abc"));
        assert!(!constant_time_eq(b"abc", b"abd"));
        assert!(!constant_time_eq(b"abc", b"ab"));
    }
}
