//! WordGit P0 probe server.
//!
//! A deliberately dependency-free HTTP/1.1 server used only to answer feasibility questions
//! about WPS JS add-ins: which origin the add-in runs under, whether it may call a loopback
//! service, and which document APIs work. It is not the production service.
//!
//! Routes:
//!   GET  /api/v1/ping     service identity and the probe work directory
//!   POST /api/v1/report   append one probe result (JSON body) to the report file
//!   GET  /addin/<path>    static files of the probe add-in (online deployment mode)
//!   OPTIONS *             CORS preflight

use std::collections::HashMap;
use std::fs::{self, OpenOptions};
use std::io::{BufRead, BufReader, Read, Write};
use std::net::{TcpListener, TcpStream};
use std::path::{Component, Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::{Duration, SystemTime, UNIX_EPOCH};

const MAX_BODY: usize = 4 * 1024 * 1024;

struct Config {
    port: u16,
    addin_dir: PathBuf,
    report: PathBuf,
    work_dir: PathBuf,
}

struct Request {
    method: String,
    path: String,
    headers: HashMap<String, String>,
    body: Vec<u8>,
}

struct Response {
    status: u16,
    content_type: &'static str,
    body: Vec<u8>,
}

impl Response {
    fn json(status: u16, body: String) -> Self {
        Response { status, content_type: "application/json; charset=utf-8", body: body.into_bytes() }
    }
    fn text(status: u16, body: &str) -> Self {
        Response { status, content_type: "text/plain; charset=utf-8", body: body.as_bytes().to_vec() }
    }
}

fn main() {
    let config = Arc::new(parse_args());
    fs::create_dir_all(&config.work_dir).expect("cannot create work dir");
    if let Some(parent) = config.report.parent() {
        fs::create_dir_all(parent).expect("cannot create report dir");
    }
    let listener = TcpListener::bind(("127.0.0.1", config.port)).expect("cannot bind 127.0.0.1");
    eprintln!(
        "wordgit-probe listening on http://127.0.0.1:{} addin={} report={} work={}",
        config.port,
        config.addin_dir.display(),
        config.report.display(),
        config.work_dir.display()
    );
    let report_lock = Arc::new(Mutex::new(()));
    for stream in listener.incoming() {
        let Ok(stream) = stream else { continue };
        let config = Arc::clone(&config);
        let report_lock = Arc::clone(&report_lock);
        std::thread::spawn(move || {
            if let Err(error) = serve(stream, &config, &report_lock) {
                eprintln!("connection error: {error}");
            }
        });
    }
}

fn parse_args() -> Config {
    let mut config = Config {
        port: 47820,
        addin_dir: PathBuf::from("addin"),
        report: PathBuf::from("report.jsonl"),
        work_dir: std::env::temp_dir().join("wordgit-probe"),
    };
    let mut args = std::env::args().skip(1);
    while let Some(flag) = args.next() {
        let value = args.next().unwrap_or_else(|| panic!("missing value for {flag}"));
        match flag.as_str() {
            "--port" => config.port = value.parse().expect("invalid --port"),
            "--addin-dir" => config.addin_dir = PathBuf::from(value),
            "--report" => config.report = PathBuf::from(value),
            "--work-dir" => config.work_dir = PathBuf::from(value),
            other => panic!("unknown argument {other}"),
        }
    }
    config
}

fn serve(stream: TcpStream, config: &Config, report_lock: &Mutex<()>) -> std::io::Result<()> {
    stream.set_read_timeout(Some(Duration::from_secs(15)))?;
    let mut reader = BufReader::new(stream.try_clone()?);
    let mut writer = stream;
    // Keep-alive: serve requests until the client closes or a request is malformed.
    loop {
        let request = match read_request(&mut reader) {
            Ok(Some(request)) => request,
            Ok(None) => return Ok(()),
            Err(error) => {
                write_response(&mut writer, None, &Response::text(400, &error))?;
                return Ok(());
            }
        };
        let origin = request.headers.get("origin").cloned();
        log_request(&request);
        let response = route(&request, config, report_lock);
        write_response(&mut writer, origin.as_deref(), &response)?;
        let close = request
            .headers
            .get("connection")
            .is_some_and(|value| value.eq_ignore_ascii_case("close"));
        if close {
            return Ok(());
        }
    }
}

fn read_request(reader: &mut BufReader<TcpStream>) -> Result<Option<Request>, String> {
    let mut line = String::new();
    match reader.read_line(&mut line) {
        Ok(0) => return Ok(None),
        Ok(_) => {}
        Err(_) => return Ok(None),
    }
    let mut parts = line.split_whitespace();
    let method = parts.next().ok_or("empty request line")?.to_string();
    let path = parts.next().ok_or("missing request target")?.to_string();
    let mut headers = HashMap::new();
    loop {
        let mut header = String::new();
        reader.read_line(&mut header).map_err(|e| e.to_string())?;
        let header = header.trim_end();
        if header.is_empty() {
            break;
        }
        if let Some((name, value)) = header.split_once(':') {
            headers.insert(name.trim().to_ascii_lowercase(), value.trim().to_string());
        }
    }
    let length: usize = headers
        .get("content-length")
        .map(|value| value.parse().map_err(|_| "invalid content-length"))
        .transpose()?
        .unwrap_or(0);
    if length > MAX_BODY {
        return Err("body too large".into());
    }
    let mut body = vec![0; length];
    reader.read_exact(&mut body).map_err(|e| e.to_string())?;
    Ok(Some(Request { method, path, headers, body }))
}

fn log_request(request: &Request) {
    let header = |name: &str| request.headers.get(name).map(String::as_str).unwrap_or("-");
    eprintln!(
        "{} {} {} host={} origin={} referer={} ua={}",
        now_millis(),
        request.method,
        request.path,
        header("host"),
        header("origin"),
        header("referer"),
        header("user-agent")
    );
}

fn route(request: &Request, config: &Config, report_lock: &Mutex<()>) -> Response {
    // DNS-rebinding guard: a loopback service must only answer requests addressed to loopback.
    let host = request.headers.get("host").map(String::as_str).unwrap_or("");
    let allowed_hosts = [format!("127.0.0.1:{}", config.port), format!("localhost:{}", config.port)];
    if !allowed_hosts.iter().any(|allowed| allowed.eq_ignore_ascii_case(host)) {
        return Response::text(421, "unexpected host");
    }
    if request.method == "OPTIONS" {
        return Response { status: 204, content_type: "text/plain", body: Vec::new() };
    }
    let path = request.path.split('?').next().unwrap_or("");
    match (request.method.as_str(), path) {
        ("GET", "/api/v1/ping") => Response::json(
            200,
            format!(
                "{{\"ok\":true,\"service\":\"wordgit-probe\",\"version\":\"{}\",\"os\":\"{}\",\"workDir\":{}}}",
                env!("CARGO_PKG_VERSION"),
                std::env::consts::OS,
                json_string(&config.work_dir.to_string_lossy())
            ),
        ),
        ("POST", "/api/v1/report") => match append_report(request, config, report_lock) {
            Ok(()) => Response::json(200, "{\"ok\":true}".into()),
            Err(error) => Response::text(500, &error.to_string()),
        },
        ("GET", _) if path == "/addin" || path.starts_with("/addin/") => serve_static(config, path),
        _ => Response::text(404, "not found"),
    }
}

fn append_report(request: &Request, config: &Config, report_lock: &Mutex<()>) -> std::io::Result<()> {
    let body = String::from_utf8_lossy(&request.body);
    let body = body.trim();
    // The add-in posts JSON; anything else is stored as a string so the report stays parseable.
    let report = if body.starts_with('{') || body.starts_with('[') { body.to_string() } else { json_string(body) };
    let origin = request.headers.get("origin").map(String::as_str).unwrap_or("");
    let line = format!(
        "{{\"time\":{},\"origin\":{},\"report\":{}}}\n",
        now_millis(),
        json_string(origin),
        report.replace('\n', " ")
    );
    let _guard = report_lock.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    let mut file = OpenOptions::new().create(true).append(true).open(&config.report)?;
    file.write_all(line.as_bytes())?;
    file.flush()
}

fn serve_static(config: &Config, path: &str) -> Response {
    let relative = path.trim_start_matches("/addin").trim_start_matches('/');
    let relative = percent_decode(relative);
    let relative = if relative.is_empty() || relative.ends_with('/') { format!("{relative}index.html") } else { relative };
    let relative = Path::new(&relative);
    if relative.components().any(|part| !matches!(part, Component::Normal(_))) {
        return Response::text(400, "invalid path");
    }
    let full = config.addin_dir.join(relative);
    match fs::read(&full) {
        Ok(body) => Response { status: 200, content_type: content_type(&full), body },
        Err(_) => Response::text(404, "not found"),
    }
}

fn content_type(path: &Path) -> &'static str {
    match path.extension().and_then(|ext| ext.to_str()).unwrap_or("") {
        "html" => "text/html; charset=utf-8",
        "js" => "text/javascript; charset=utf-8",
        "css" => "text/css; charset=utf-8",
        "xml" => "application/xml; charset=utf-8",
        "json" => "application/json; charset=utf-8",
        "svg" => "image/svg+xml",
        "png" => "image/png",
        _ => "application/octet-stream",
    }
}

fn write_response(writer: &mut TcpStream, origin: Option<&str>, response: &Response) -> std::io::Result<()> {
    let reason = match response.status {
        200 => "OK",
        204 => "No Content",
        400 => "Bad Request",
        404 => "Not Found",
        421 => "Misdirected Request",
        _ => "Error",
    };
    let mut head = format!(
        "HTTP/1.1 {} {}\r\nContent-Type: {}\r\nContent-Length: {}\r\nCache-Control: no-store\r\n",
        response.status,
        reason,
        response.content_type,
        response.body.len()
    );
    // P0 reflects every origin so the report shows which origins WPS actually uses.
    // The production service must allow-list the add-in origin instead.
    if let Some(origin) = origin {
        head.push_str(&format!(
            "Access-Control-Allow-Origin: {origin}\r\nVary: Origin\r\n\
             Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n\
             Access-Control-Allow-Headers: Content-Type, X-WordGit-Token\r\n\
             Access-Control-Allow-Private-Network: true\r\n"
        ));
    }
    head.push_str("\r\n");
    writer.write_all(head.as_bytes())?;
    writer.write_all(&response.body)?;
    writer.flush()
}

fn percent_decode(input: &str) -> String {
    let bytes = input.as_bytes();
    let mut output = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' && index + 2 < bytes.len() {
            let hex = std::str::from_utf8(&bytes[index + 1..index + 3]).unwrap_or("");
            if let Ok(value) = u8::from_str_radix(hex, 16) {
                output.push(value);
                index += 3;
                continue;
            }
        }
        output.push(bytes[index]);
        index += 1;
    }
    String::from_utf8_lossy(&output).into_owned()
}

fn json_string(value: &str) -> String {
    let mut output = String::with_capacity(value.len() + 2);
    output.push('"');
    for ch in value.chars() {
        match ch {
            '"' => output.push_str("\\\""),
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            c if (c as u32) < 0x20 => output.push_str(&format!("\\u{:04x}", c as u32)),
            c => output.push(c),
        }
    }
    output.push('"');
    output
}

fn now_millis() -> u128 {
    SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_millis()).unwrap_or(0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn json_string_escapes_controls_and_quotes() {
        assert_eq!(json_string("a\"b\\c\n\u{1}中"), "\"a\\\"b\\\\c\\n\\u0001中\"");
    }

    #[test]
    fn percent_decode_handles_utf8_and_truncation() {
        assert_eq!(percent_decode("ui/%E4%B8%AD.html"), "ui/中.html");
        assert_eq!(percent_decode("a%2"), "a%2");
        assert_eq!(percent_decode("a%"), "a%");
    }
}
