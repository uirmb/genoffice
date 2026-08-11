use std::{
    collections::HashMap,
    fs,
    net::SocketAddr,
    path::{Path as FsPath, PathBuf},
    sync::Arc,
};

use axum::{
    body::{Body, Bytes},
    extract::{DefaultBodyLimit, Path, Query, State},
    http::{header, HeaderValue, StatusCode},
    response::Response,
    routing::{get, post},
    Json, Router,
};
use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};
use ironcalc::{base::Model, export::save_to_xlsx};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use sha2::{Digest, Sha256};
use tokio::sync::{Mutex, RwLock};
use uuid::Uuid;
use xlsx_sidecar::{
    archive::{archive_manifest, read_entries_to_dir, save_archive, scan_entries_for_text, EntryContent},
    CellRange, WorkbookSessions,
};

const XLSX_MIME: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WorkbookSession {
    session_id: String,
    source: SessionSource,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
enum SessionSource {
    Blank,
    Uploaded,
}

trait SessionStore: Send + Sync {
    fn kind(&self) -> &'static str;
}

#[derive(Default)]
struct MemorySessionStore {
    sessions: RwLock<HashMap<String, WorkbookSession>>,
}

impl SessionStore for MemorySessionStore {
    fn kind(&self) -> &'static str {
        "memory"
    }
}

struct EngineState {
    workbooks: WorkbookSessions,
    files: HashMap<String, PathBuf>,
    metadata: HashMap<String, Value>,
}

impl Default for EngineState {
    fn default() -> Self {
        Self {
            workbooks: WorkbookSessions::new(),
            files: HashMap::new(),
            metadata: HashMap::new(),
        }
    }
}

#[derive(Clone)]
struct AppState {
    sessions: Arc<MemorySessionStore>,
    engine: Arc<Mutex<EngineState>>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CreateSessionRequest {
    #[serde(default)]
    source: Option<String>,
}

#[derive(Debug, Deserialize)]
struct OpenWorkbookQuery {
    name: Option<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ReadRangeRequest {
    sheet_id: String,
    range: CellRange,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ReadFormulaCellsRequest {
    sheet_id: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveEntriesRequest {
    entries: Vec<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveScanRequest {
    entries: Vec<String>,
    needle: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveContentRequest {
    name: String,
    content_base64: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveSaveRequest {
    name: Option<String>,
    #[serde(default)]
    replacements: Vec<ArchiveContentRequest>,
    #[serde(default)]
    removals: Vec<String>,
    #[serde(default)]
    additions: Vec<ArchiveContentRequest>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct HealthResponse {
    ok: bool,
    service: &'static str,
    session_store: &'static str,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct CreateSessionResponse {
    session_id: String,
    source: SessionSource,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveReadEntry {
    name: String,
    content_base64: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveReadResponse {
    entries: Vec<ArchiveReadEntry>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveScanResponse {
    matches: Vec<String>,
}

type ApiError = (StatusCode, String);

async fn health(State(state): State<AppState>) -> Json<HealthResponse> {
    Json(HealthResponse {
        ok: true,
        service: "xlsx-engine-service",
        session_store: state.sessions.kind(),
    })
}

async fn create_session(
    State(state): State<AppState>,
    Json(request): Json<CreateSessionRequest>,
) -> (StatusCode, Json<CreateSessionResponse>) {
    let source = match request.source.as_deref() {
        Some("uploaded") => SessionSource::Uploaded,
        _ => SessionSource::Blank,
    };

    // This lightweight endpoint only reserves an opaque application session.
    // Real workbook editing sessions are created by /v1/workbooks or
    // /v1/workbooks/blank and are backed by xlsx-sidecar WorkbookSessions.
    let session_id = format!("xls_{}", Uuid::new_v4().simple());
    let session = WorkbookSession {
        session_id: session_id.clone(),
        source: source.clone(),
    };

    state
        .sessions
        .sessions
        .write()
        .await
        .insert(session_id.clone(), session);

    (
        StatusCode::CREATED,
        Json(CreateSessionResponse { session_id, source }),
    )
}

async fn create_blank_workbook(
    State(state): State<AppState>,
    Query(query): Query<OpenWorkbookQuery>,
) -> Result<(StatusCode, Json<Value>), ApiError> {
    let name = safe_workbook_name(query.name.as_deref().unwrap_or("Untitled.xlsx"));
    let directory = workbook_directory();
    fs::create_dir_all(&directory).map_err(internal_error)?;
    let path = directory.join(&name);

    let mut model = Model::new_empty("Untitled", "en", "UTC", "en")
        .map_err(internal_error)?;
    model.evaluate();
    save_to_xlsx(&model, path.to_string_lossy().as_ref()).map_err(internal_error)?;

    register_workbook_path(&state, path, name).await
}

async fn open_workbook(
    State(state): State<AppState>,
    Query(query): Query<OpenWorkbookQuery>,
    bytes: Bytes,
) -> Result<(StatusCode, Json<Value>), ApiError> {
    let name = safe_workbook_name(query.name.as_deref().unwrap_or("workbook.xlsx"));
    let directory = workbook_directory();
    fs::create_dir_all(&directory).map_err(internal_error)?;
    let path = directory.join(&name);
    fs::write(&path, &bytes).map_err(internal_error)?;
    register_workbook_path(&state, path, name).await
}

async fn register_workbook_path(
    state: &AppState,
    path: PathBuf,
    name: String,
) -> Result<(StatusCode, Json<Value>), ApiError> {
    let bytes = fs::read(&path).map_err(internal_error)?;
    let sha256 = format!("{:x}", Sha256::digest(&bytes));
    let mut engine = state.engine.lock().await;
    let metadata = match engine.workbooks.open(&path) {
        Ok(metadata) => metadata,
        Err(error) => {
            if let Some(parent) = path.parent() {
                let _ = fs::remove_dir_all(parent);
            }
            return Err((StatusCode::UNPROCESSABLE_ENTITY, error.to_string()));
        }
    };
    let session_id = metadata.session_id.clone();
    let value = web_metadata_value(metadata, &name, &sha256)?;
    engine.files.insert(session_id.clone(), path);
    engine.metadata.insert(session_id, value.clone());
    Ok((StatusCode::CREATED, Json(value)))
}

async fn get_session_metadata(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
) -> Result<Json<Value>, ApiError> {
    let engine = state.engine.lock().await;
    let metadata = engine
        .metadata
        .get(&session_id)
        .cloned()
        .ok_or_else(|| (StatusCode::NOT_FOUND, "Unknown workbook session.".to_string()))?;
    Ok(Json(metadata))
}

async fn read_range(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
    Json(request): Json<ReadRangeRequest>,
) -> Result<Json<Value>, ApiError> {
    let mut engine = state.engine.lock().await;
    let result = engine
        .workbooks
        .read_range(&session_id, &request.sheet_id, &request.range)
        .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
    Ok(Json(serde_json::to_value(result).map_err(internal_error)?))
}

async fn read_formula_cells(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
    Json(request): Json<ReadFormulaCellsRequest>,
) -> Result<Json<Value>, ApiError> {
    let mut engine = state.engine.lock().await;
    let result = engine
        .workbooks
        .read_formula_cells(&session_id, &request.sheet_id)
        .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
    Ok(Json(serde_json::to_value(result).map_err(internal_error)?))
}

async fn archive_manifest_for_session(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
) -> Result<Json<Value>, ApiError> {
    let path = session_path(&state, &session_id).await?;
    let entries = archive_manifest(&path)
        .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
    Ok(Json(serde_json::json!({ "entries": entries })))
}

async fn archive_read_for_session(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
    Json(request): Json<ArchiveEntriesRequest>,
) -> Result<Json<ArchiveReadResponse>, ApiError> {
    let path = session_path(&state, &session_id).await?;
    let directory = scratch_directory("read");
    fs::create_dir_all(&directory).map_err(internal_error)?;

    let result = (|| -> Result<Vec<ArchiveReadEntry>, ApiError> {
        let extracted = read_entries_to_dir(&path, &request.entries, &directory)
            .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
        extracted
            .into_iter()
            .map(|entry| {
                let content = fs::read(entry.path).map_err(internal_error)?;
                Ok(ArchiveReadEntry {
                    name: entry.name,
                    content_base64: BASE64.encode(content),
                })
            })
            .collect()
    })();

    let _ = fs::remove_dir_all(&directory);
    Ok(Json(ArchiveReadResponse { entries: result? }))
}

async fn archive_scan_for_session(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
    Json(request): Json<ArchiveScanRequest>,
) -> Result<Json<ArchiveScanResponse>, ApiError> {
    let path = session_path(&state, &session_id).await?;
    let matches = scan_entries_for_text(&path, &request.entries, &request.needle)
        .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
    Ok(Json(ArchiveScanResponse { matches }))
}

async fn archive_save_for_session(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
    Json(request): Json<ArchiveSaveRequest>,
) -> Result<Response, ApiError> {
    let source_path = session_path(&state, &session_id).await?;
    let name = safe_workbook_name(request.name.as_deref().unwrap_or("workbook.xlsx"));
    let directory = workbook_directory();
    let content_directory = directory.join("patch");
    fs::create_dir_all(&content_directory).map_err(internal_error)?;
    let target_path = directory.join(&name);

    let save_result = (|| -> Result<(), ApiError> {
        let replacements = write_archive_content(&content_directory, "replace", &request.replacements)?;
        let additions = write_archive_content(&content_directory, "add", &request.additions)?;
        save_archive(
            &source_path,
            &target_path,
            &replacements,
            &request.removals,
            &additions,
        )
        .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
        Ok(())
    })();
    let _ = fs::remove_dir_all(&content_directory);
    save_result?;

    let bytes = fs::read(&target_path).map_err(internal_error)?;
    let sha256 = format!("{:x}", Sha256::digest(&bytes));

    let mut engine = state.engine.lock().await;
    let metadata = engine
        .workbooks
        .open(&target_path)
        .map_err(|error| (StatusCode::UNPROCESSABLE_ENTITY, error.to_string()))?;
    let saved_session_id = metadata.session_id.clone();
    let value = web_metadata_value(metadata, &name, &sha256)?;
    engine.files.insert(saved_session_id.clone(), target_path);
    engine.metadata.insert(saved_session_id.clone(), value);
    drop(engine);

    let mut response = Response::new(Body::from(bytes));
    *response.status_mut() = StatusCode::OK;
    response.headers_mut().insert(
        header::CONTENT_TYPE,
        HeaderValue::from_static(XLSX_MIME),
    );
    response.headers_mut().insert(
        "x-xlsx-session",
        HeaderValue::from_str(&saved_session_id).map_err(internal_error)?,
    );
    Ok(response)
}

async fn delete_session(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
) -> StatusCode {
    if state
        .sessions
        .sessions
        .write()
        .await
        .remove(&session_id)
        .is_some()
    {
        return StatusCode::NO_CONTENT;
    }

    let mut engine = state.engine.lock().await;
    let path = engine.files.remove(&session_id);
    engine.metadata.remove(&session_id);
    let closed = engine.workbooks.close(&session_id).is_ok();
    drop(engine);

    if let Some(path) = path {
        if let Some(parent) = path.parent() {
            let _ = fs::remove_dir_all(parent);
        }
    }

    if closed {
        StatusCode::NO_CONTENT
    } else {
        StatusCode::NOT_FOUND
    }
}

async fn session_path(state: &AppState, session_id: &str) -> Result<PathBuf, ApiError> {
    let engine = state.engine.lock().await;
    engine
        .files
        .get(session_id)
        .cloned()
        .ok_or_else(|| (StatusCode::NOT_FOUND, "Unknown workbook session.".to_string()))
}

fn write_archive_content(
    directory: &FsPath,
    prefix: &str,
    items: &[ArchiveContentRequest],
) -> Result<Vec<EntryContent>, ApiError> {
    items
        .iter()
        .enumerate()
        .map(|(index, item)| {
            let content = BASE64
                .decode(&item.content_base64)
                .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
            let path = directory.join(format!("{prefix}-{index}.bin"));
            fs::write(&path, content).map_err(internal_error)?;
            Ok(EntryContent {
                name: item.name.clone(),
                content_path: path,
            })
        })
        .collect()
}

fn workbook_directory() -> PathBuf {
    std::env::temp_dir()
        .join("genoffice-xlsx-engine")
        .join(Uuid::new_v4().to_string())
}

fn scratch_directory(kind: &str) -> PathBuf {
    std::env::temp_dir()
        .join("genoffice-xlsx-engine-scratch")
        .join(format!("{kind}-{}", Uuid::new_v4()))
}

fn web_metadata_value<T: Serialize>(metadata: T, name: &str, sha256: &str) -> Result<Value, ApiError> {
    let mut value = serde_json::to_value(metadata).map_err(internal_error)?;
    let object = value
        .as_object_mut()
        .ok_or_else(|| internal_error("Workbook metadata was not an object."))?;
    object.remove("path");
    object.insert("name".into(), Value::String(name.to_string()));
    object.insert("sha256".into(), Value::String(sha256.to_string()));
    object.insert("readOnly".into(), Value::Bool(false));
    object.insert("needsSaveAs".into(), Value::Bool(false));
    Ok(value)
}

fn safe_workbook_name(input: &str) -> String {
    let mut name = input
        .chars()
        .map(|character| match character {
            '/' | '\\' | '\0' => '_',
            other => other,
        })
        .collect::<String>();
    if name.trim().is_empty() {
        name = "workbook.xlsx".to_string();
    }
    if !name.to_ascii_lowercase().ends_with(".xlsx") {
        name.push_str(".xlsx");
    }
    name
}

fn internal_error(error: impl ToString) -> ApiError {
    (StatusCode::INTERNAL_SERVER_ERROR, error.to_string())
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let listen_addr = std::env::var("XLSX_ENGINE_LISTEN")
        .unwrap_or_else(|_| "127.0.0.1:7301".to_string());
    let address: SocketAddr = listen_addr.parse()?;

    let state = AppState {
        sessions: Arc::new(MemorySessionStore::default()),
        engine: Arc::new(Mutex::new(EngineState::default())),
    };

    let app = Router::new()
        .route("/health", get(health))
        .route("/v1/sessions", post(create_session))
        .route("/v1/workbooks", post(open_workbook))
        .route("/v1/workbooks/blank", post(create_blank_workbook))
        .route("/v1/sessions/{session_id}", get(get_session_metadata).delete(delete_session))
        .route("/v1/sessions/{session_id}/ranges", post(read_range))
        .route(
            "/v1/sessions/{session_id}/formulas",
            post(read_formula_cells),
        )
        .route(
            "/v1/sessions/{session_id}/archive/manifest",
            get(archive_manifest_for_session),
        )
        .route(
            "/v1/sessions/{session_id}/archive/read",
            post(archive_read_for_session),
        )
        .route(
            "/v1/sessions/{session_id}/archive/scan",
            post(archive_scan_for_session),
        )
        .route(
            "/v1/sessions/{session_id}/archive/save",
            post(archive_save_for_session),
        )
        .layer(DefaultBodyLimit::disable())
        .with_state(state);

    let listener = tokio::net::TcpListener::bind(address).await?;
    println!("xlsx-engine-service listening on http://{address}");

    axum::serve(listener, app)
        .with_graceful_shutdown(shutdown_signal())
        .await?;

    Ok(())
}

async fn shutdown_signal() {
    let _ = tokio::signal::ctrl_c().await;
}
