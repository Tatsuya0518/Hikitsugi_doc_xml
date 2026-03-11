//! 引継ぎドキュメントシステム – Rust / Axum バックエンド
//!
//! 起動: `RUST_LOG=info cargo run`
//! デバッグ: `RUST_LOG=debug cargo run`
mod converter;
mod exporter;

use std::fs;
use std::net::SocketAddr;
use std::path::{Component, Path, PathBuf};
use std::sync::Arc;

use anyhow::Result;
use axum::{
    body::Body,
    extract::{Path as AxumPath, Query, State},
    http::{header, HeaderMap, StatusCode},
    response::{IntoResponse, Response},
    routing::{get, post},
    Json, Router,
};
use serde::{Deserialize, Serialize};
use tower_http::services::ServeDir;
use tracing::info;

// ── Application state ─────────────────────────────────────────────────────────

#[derive(Clone)]
struct AppState {
    data_dir: PathBuf,
    cache_dir: PathBuf,
    template_dir: PathBuf,
}

// ── Main ─────────────────────────────────────────────────────────────────────

#[tokio::main]
async fn main() -> Result<()> {
    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::from_default_env()
                .add_directive("hikitsugi_doc_xml=info".parse()?),
        )
        .init();

    let base = std::env::current_dir()?;
    let data_dir = base.join("data").join("watch");
    let cache_dir = base.join("cache");
    let template_dir = base.join("templates");

    fs::create_dir_all(&data_dir)?;
    fs::create_dir_all(&cache_dir)?;

    let state = Arc::new(AppState {
        data_dir,
        cache_dir,
        template_dir,
    });

    let app = Router::new()
        .route("/", get(index_handler))
        .route("/api/structure", get(api_structure))
        .route(
            "/api/document/:team/:code/:filename",
            get(api_get_document),
        )
        .route(
            "/api/document/:team/:code/:filename/save",
            post(api_save_document),
        )
        .route(
            "/api/document/:team/:code/:filename/export",
            get(api_export_document),
        )
        .route("/api/search", get(api_search))
        .nest_service("/static", ServeDir::new("static"))
        .with_state(state);

    let port: u16 = std::env::var("PORT")
        .ok()
        .and_then(|p| p.parse().ok())
        .unwrap_or(5000);
    let addr = SocketAddr::from(([0, 0, 0, 0], port));
    info!("Starting server on http://{}", addr);

    let listener = tokio::net::TcpListener::bind(addr).await?;
    axum::serve(listener, app).await?;
    Ok(())
}

// ── Route handlers ────────────────────────────────────────────────────────────

/// Serve templates/index.html
async fn index_handler(State(st): State<Arc<AppState>>) -> impl IntoResponse {
    let path = st.template_dir.join("index.html");
    match fs::read_to_string(&path) {
        Ok(html) => Response::builder()
            .status(StatusCode::OK)
            .header(header::CONTENT_TYPE, "text/html; charset=utf-8")
            .body(Body::from(html))
            .unwrap(),
        Err(_) => Response::builder()
            .status(StatusCode::NOT_FOUND)
            .body(Body::from("index.html not found"))
            .unwrap(),
    }
}

// ── /api/structure ────────────────────────────────────────────────────────────

#[derive(Serialize)]
struct TeamInfo {
    name: String,
    codes: Vec<CodeInfo>,
}

#[derive(Serialize)]
struct CodeInfo {
    name: String,
    documents: Vec<String>,
}

#[derive(Serialize)]
struct StructureResponse {
    teams: Vec<TeamInfo>,
}

async fn api_structure(State(st): State<Arc<AppState>>) -> impl IntoResponse {
    let teams = read_structure(&st.data_dir);
    Json(StructureResponse { teams })
}

fn read_structure(data_dir: &Path) -> Vec<TeamInfo> {
    let mut teams = Vec::new();
    let Ok(team_entries) = dir_entries_sorted(data_dir) else {
        return teams;
    };
    for team_path in team_entries {
        if !team_path.is_dir() { continue; }
        let team_name = file_name_str(&team_path);
        if team_name.starts_with('.') { continue; }
        let Ok(code_entries) = dir_entries_sorted(&team_path) else { continue; };
        let mut codes = Vec::new();
        for code_path in code_entries {
            if !code_path.is_dir() { continue; }
            let code_name = file_name_str(&code_path);
            if code_name.starts_with('.') { continue; }
            let Ok(doc_entries) = dir_entries_sorted(&code_path) else { continue; };
            let docs: Vec<String> = doc_entries
                .iter()
                .filter(|p| {
                    let n = file_name_str(p);
                    n.ends_with(".xlsx") && !n.starts_with("~$")
                })
                .map(|p| file_name_str(p))
                .collect();
            if !docs.is_empty() {
                codes.push(CodeInfo { name: code_name, documents: docs });
            }
        }
        if !codes.is_empty() {
            teams.push(TeamInfo { name: team_name, codes });
        }
    }
    teams
}

// ── /api/document/:team/:code/:filename ──────────────────────────────────────

async fn api_get_document(
    State(st): State<Arc<AppState>>,
    AxumPath((team, code, filename)): AxumPath<(String, String, String)>,
) -> impl IntoResponse {
    let Some(xlsx_path) = safe_path(&st.data_dir, &[&team, &code, &filename]) else {
        return error_json(StatusCode::BAD_REQUEST, "不正なパスです");
    };
    if !xlsx_path.exists() {
        return error_json(StatusCode::NOT_FOUND, "ファイルが見つかりません");
    }
    let xml_path = cache_path(&st.cache_dir, &st.data_dir, &xlsx_path);
    match ensure_converted(&xlsx_path, &xml_path) {
        Err(e) => error_json(StatusCode::INTERNAL_SERVER_ERROR, &e.to_string()),
        Ok(()) => match fs::read(&xml_path) {
            Err(e) => error_json(StatusCode::INTERNAL_SERVER_ERROR, &e.to_string()),
            Ok(bytes) => Response::builder()
                .status(StatusCode::OK)
                .header(header::CONTENT_TYPE, "application/xml; charset=utf-8")
                .body(Body::from(bytes))
                .unwrap(),
        },
    }
}

// ── /api/document/:team/:code/:filename/save ─────────────────────────────────

async fn api_save_document(
    State(st): State<Arc<AppState>>,
    AxumPath((team, code, filename)): AxumPath<(String, String, String)>,
    body: String,
) -> impl IntoResponse {
    let Some(xlsx_path) = safe_path(&st.data_dir, &[&team, &code, &filename]) else {
        return error_json(StatusCode::BAD_REQUEST, "不正なパスです");
    };
    // Validate XML
    let mut reader = quick_xml::Reader::from_str(&body);
    let mut buf = Vec::new();
    let mut valid = false;
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(_)) => { valid = true; }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(e) => {
                return error_json(
                    StatusCode::BAD_REQUEST,
                    &format!("XMLの形式が不正です: {}", e),
                );
            }
            _ => {}
        }
    }

    let xml_path = cache_path(&st.cache_dir, &st.data_dir, &xlsx_path);
    if let Some(parent) = xml_path.parent() {
        let _ = fs::create_dir_all(parent);
    }
    match fs::write(&xml_path, body.as_bytes()) {
        Ok(()) => Json(serde_json::json!({"success": true})).into_response(),
        Err(e) => error_json(StatusCode::INTERNAL_SERVER_ERROR, &e.to_string()),
    }
}

// ── /api/document/:team/:code/:filename/export ───────────────────────────────

async fn api_export_document(
    State(st): State<Arc<AppState>>,
    AxumPath((team, code, filename)): AxumPath<(String, String, String)>,
) -> impl IntoResponse {
    let Some(xlsx_path) = safe_path(&st.data_dir, &[&team, &code, &filename]) else {
        return error_json(StatusCode::BAD_REQUEST, "不正なパスです");
    };
    let xml_path = cache_path(&st.cache_dir, &st.data_dir, &xlsx_path);

    // Ensure we have a cached XML to export from
    if !xml_path.exists() {
        if !xlsx_path.exists() {
            return error_json(StatusCode::NOT_FOUND, "ファイルが見つかりません");
        }
        if let Err(e) = ensure_converted(&xlsx_path, &xml_path) {
            return error_json(StatusCode::INTERNAL_SERVER_ERROR, &e.to_string());
        }
    }

    match exporter::export_xml_to_excel_bytes(&xml_path) {
        Err(e) => error_json(StatusCode::INTERNAL_SERVER_ERROR, &e.to_string()),
        Ok(bytes) => {
            let cd = format!(
                "attachment; filename=\"{}\"",
                filename.replace('"', "")
            );
            Response::builder()
                .status(StatusCode::OK)
                .header(
                    header::CONTENT_TYPE,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                .header(header::CONTENT_DISPOSITION, cd)
                .body(Body::from(bytes))
                .unwrap()
        }
    }
}

// ── /api/search?q=… ──────────────────────────────────────────────────────────

#[derive(Deserialize)]
struct SearchQuery {
    q: Option<String>,
}

#[derive(Serialize)]
struct SearchResult {
    team: String,
    code: String,
    filename: String,
    sheet: String,
    cell: String,
    value: String,
}

#[derive(Serialize)]
struct SearchResponse {
    results: Vec<SearchResult>,
    query: String,
}

async fn api_search(
    State(st): State<Arc<AppState>>,
    Query(params): Query<SearchQuery>,
) -> impl IntoResponse {
    let query = params.q.unwrap_or_default();
    let query = query.trim().to_string();
    if query.is_empty() {
        return Json(SearchResponse { results: vec![], query });
    }

    let query_lower = query.to_lowercase();
    let mut results = Vec::new();

    if let Ok(team_entries) = dir_entries_sorted(&st.data_dir) {
        for team_path in team_entries {
            if !team_path.is_dir() { continue; }
            let team_name = file_name_str(&team_path);
            if team_name.starts_with('.') { continue; }
            let Ok(code_entries) = dir_entries_sorted(&team_path) else { continue; };
            for code_path in code_entries {
                if !code_path.is_dir() { continue; }
                let code_name = file_name_str(&code_path);
                if code_name.starts_with('.') { continue; }
                let Ok(doc_entries) = dir_entries_sorted(&code_path) else { continue; };
                for doc_path in doc_entries {
                    let doc_name = file_name_str(&doc_path);
                    if !doc_name.ends_with(".xlsx") || doc_name.starts_with("~$") {
                        continue;
                    }
                    let xml_path = cache_path(&st.cache_dir, &st.data_dir, &doc_path);
                    if let Ok(()) = ensure_converted(&doc_path, &xml_path) {
                        if let Ok(matches) = search_in_xml(&xml_path, &query_lower) {
                            for (sheet, cell, value) in matches {
                                results.push(SearchResult {
                                    team: team_name.clone(),
                                    code: code_name.clone(),
                                    filename: doc_name.clone(),
                                    sheet,
                                    cell,
                                    value,
                                });
                            }
                        }
                    }
                }
            }
        }
    }

    Json(SearchResponse { results, query })
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/// Validate and construct a file path from untrusted path components.
/// Returns `None` if any component contains `..` or path separators.
fn safe_path(base: &Path, parts: &[&str]) -> Option<PathBuf> {
    let mut path = base.to_path_buf();
    for part in parts {
        // Reject empty, "..", or components containing path separators
        if part.is_empty() || part.contains('/') || part.contains('\\') || *part == ".." {
            return None;
        }
        // Extra check via PathBuf parsing
        let p = Path::new(part);
        let mut comps = p.components();
        match comps.next() {
            Some(Component::Normal(_)) => {}
            _ => return None,
        }
        if comps.next().is_some() {
            return None;
        }
        path.push(part);
    }
    Some(path)
}

/// Compute cache XML path for a given xlsx file.
fn cache_path(cache_dir: &Path, data_dir: &Path, xlsx_path: &Path) -> PathBuf {
    let rel = xlsx_path.strip_prefix(data_dir).unwrap_or(xlsx_path);
    let xml_name = rel.with_extension("xml");
    cache_dir.join(xml_name)
}

/// Convert xlsx to XML if cache is missing or stale.
fn ensure_converted(xlsx_path: &Path, xml_path: &Path) -> Result<()> {
    if converter::is_cache_valid(xlsx_path, xml_path) {
        return Ok(());
    }
    converter::convert_excel_to_xml(xlsx_path, xml_path)
}

/// Return sorted directory entries.
fn dir_entries_sorted(dir: &Path) -> std::io::Result<Vec<PathBuf>> {
    let mut entries: Vec<PathBuf> = fs::read_dir(dir)?
        .filter_map(|e| e.ok().map(|e| e.path()))
        .collect();
    entries.sort();
    Ok(entries)
}

fn file_name_str(p: &Path) -> String {
    p.file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("")
        .to_string()
}

/// Full-text search within a cached XML file.
/// Returns (sheet_name, cell_ref, value) for each hit.
fn search_in_xml(xml_path: &Path, query_lower: &str) -> Result<Vec<(String, String, String)>> {
    let xml_bytes = fs::read(xml_path)?;
    let xml_str = std::str::from_utf8(&xml_bytes)?;

    let mut results = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml_str);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_sheet = String::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                match e.local_name().as_ref() {
                    b"sheet" => {
                        current_sheet = e
                            .try_get_attribute("name")
                            .ok()
                            .flatten()
                            .and_then(|a| a.unescape_value().ok())
                            .map(|c| c.into_owned())
                            .unwrap_or_default();
                    }
                    b"cell" => {
                        let value = e
                            .try_get_attribute("value")
                            .ok()
                            .flatten()
                            .and_then(|a| a.unescape_value().ok())
                            .map(|c| c.into_owned())
                            .unwrap_or_default();
                        if !value.is_empty() && value.to_lowercase().contains(query_lower) {
                            let row = e
                                .try_get_attribute("row")
                                .ok()
                                .flatten()
                                .and_then(|a| a.unescape_value().ok())
                                .map(|c| c.into_owned())
                                .unwrap_or_default();
                            let col: u32 = e
                                .try_get_attribute("col")
                                .ok()
                                .flatten()
                                .and_then(|a| a.unescape_value().ok())
                                .and_then(|c| c.parse().ok())
                                .unwrap_or(0);
                            let cell_ref = format!("{}{}", converter::col_num_to_str(col), row);
                            let trimmed = if value.len() > 300 {
                                value[..300].to_string()
                            } else {
                                value.clone()
                            };
                            results.push((current_sheet.clone(), cell_ref, trimmed));
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    Ok(results)
}

/// Build a JSON error response.
fn error_json(status: StatusCode, msg: &str) -> Response {
    Response::builder()
        .status(status)
        .header(header::CONTENT_TYPE, "application/json")
        .body(Body::from(
            serde_json::json!({"error": msg}).to_string(),
        ))
        .unwrap()
}
