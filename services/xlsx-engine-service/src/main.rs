use std::{
    collections::HashMap,
    fs, io,
    net::SocketAddr,
    path::{Path as FsPath, PathBuf},
    sync::Arc,
    time::{Duration, Instant},
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
    recalc::{recalc_cells, RecalcCache, RecalcEdit, RecalcRead},
    CellRange, WorkbookSessions,
};

const XLSX_MIME: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DEFAULT_MAX_WORKBOOK_MB: usize = 100;
// Archive mutation JSON carries base64 content, so it needs headroom above the
// raw workbook upload limit while still remaining bounded in production.
const DEFAULT_MAX_REQUEST_MB: usize = 384;
const DEFAULT_SESSION_TTL_SECS: u64 = 60 * 60;
const DEFAULT_CLEANUP_INTERVAL_SECS: u64 = 60;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WorkbookSession {
    session_id: String,
    source: SessionSource,
    #[serde(skip)]
    last_access: Instant,
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
    recalc: RecalcCache,
    files: HashMap<String, PathBuf>,
    metadata: HashMap<String, Value>,
    last_access: HashMap<String, Instant>,
}

impl Default for EngineState {
    fn default() -> Self {
        Self {
            workbooks: WorkbookSessions::new(),
            recalc: RecalcCache::new(),
            files: HashMap::new(),
            metadata: HashMap::new(),
            last_access: HashMap::new(),
        }
    }
}

#[derive(Clone)]
struct AppState {
    sessions: Arc<MemorySessionStore>,
    engine: Arc<Mutex<EngineState>>,
    max_workbook_bytes: usize,
    session_ttl: Duration,
    cleanup_interval: Duration,
    workbook_root: PathBuf,
    scratch_root: PathBuf,
}

struct EngineWorkspace {
    root: PathBuf,
    workbook_root: PathBuf,
    scratch_root: PathBuf,
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
struct RecalcRequest {
    edits: Vec<RecalcEdit>,
    reads: Vec<RecalcRead>,
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
        last_access: Instant::now(),
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
    let directory = workbook_directory(&state);
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
    if bytes.len() > state.max_workbook_bytes {
        return Err((
            StatusCode::PAYLOAD_TOO_LARGE,
            format!(
                "Workbook exceeds the configured {}MB upload limit.",
                state.max_workbook_bytes / (1024 * 1024)
            ),
        ));
    }

    let name = safe_workbook_name(query.name.as_deref().unwrap_or("workbook.xlsx"));
    let directory = workbook_directory(&state);
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
    engine.metadata.insert(session_id.clone(), value.clone());
    engine.last_access.insert(session_id, Instant::now());
    Ok((StatusCode::CREATED, Json(value)))
}

fn touch_workbook_session(engine: &mut EngineState, session_id: &str) {
    if engine.files.contains_key(session_id) {
        engine
            .last_access
            .insert(session_id.to_string(), Instant::now());
    }
}

async fn get_session_metadata(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
) -> Result<Json<Value>, ApiError> {
    let mut engine = state.engine.lock().await;
    let metadata = engine
        .metadata
        .get(&session_id)
        .cloned()
        .ok_or_else(|| (StatusCode::NOT_FOUND, "Unknown workbook session.".to_string()))?;
    touch_workbook_session(&mut engine, &session_id);
    Ok(Json(metadata))
}

async fn read_range(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
    Json(request): Json<ReadRangeRequest>,
) -> Result<Json<Value>, ApiError> {
    let mut engine = state.engine.lock().await;
    touch_workbook_session(&mut engine, &session_id);
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
    touch_workbook_session(&mut engine, &session_id);
    let result = engine
        .workbooks
        .read_formula_cells(&session_id, &request.sheet_id)
        .map_err(|error| (StatusCode::BAD_REQUEST, error.to_string()))?;
    Ok(Json(serde_json::to_value(result).map_err(internal_error)?))
}

async fn recalc_workbook(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
    Json(request): Json<RecalcRequest>,
) -> Result<Json<Value>, ApiError> {
    let mut engine = state.engine.lock().await;
    let path = engine
        .files
        .get(&session_id)
        .cloned()
        .ok_or_else(|| (StatusCode::NOT_FOUND, "Unknown workbook session.".to_string()))?;
    touch_workbook_session(&mut engine, &session_id);
    let result = recalc_cells(&mut engine.recalc, &path, &request.edits, &request.reads)
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
    let directory = scratch_directory(&state, "read");
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
    let directory = workbook_directory(&state);
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
    engine
        .last_access
        .insert(saved_session_id.clone(), Instant::now());
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
    engine.last_access.remove(&session_id);
    if let Some(path) = path.as_deref() {
        engine.recalc.purge(path);
    }
    let closed = engine.workbooks.close(&session_id).is_ok();
    drop(engine);

    remove_workbook_directory(path);

    if closed {
        StatusCode::NO_CONTENT
    } else {
        StatusCode::NOT_FOUND
    }
}

fn remove_workbook_directory(path: Option<PathBuf>) {
    if let Some(path) = path {
        if let Some(parent) = path.parent() {
            let _ = fs::remove_dir_all(parent);
        }
    }
}

async fn cleanup_expired_sessions(state: &AppState) {
    let now = Instant::now();
    {
        let mut sessions = state.sessions.sessions.write().await;
        sessions.retain(|_, session| now.duration_since(session.last_access) < state.session_ttl);
    }

    let expired_paths = {
        let mut engine = state.engine.lock().await;
        let expired_ids = engine
            .last_access
            .iter()
            .filter_map(|(session_id, last_access)| {
                (now.duration_since(*last_access) >= state.session_ttl).then(|| session_id.clone())
            })
            .collect::<Vec<_>>();
        let mut paths = Vec::with_capacity(expired_ids.len());

        for session_id in expired_ids {
            engine.last_access.remove(&session_id);
            let path = engine.files.remove(&session_id);
            engine.metadata.remove(&session_id);
            if let Some(path) = path.as_deref() {
                engine.recalc.purge(path);
            }
            let _ = engine.workbooks.close(&session_id);
            if let Some(path) = path {
                paths.push(path);
            }
        }
        paths
    };

    for path in expired_paths {
        remove_workbook_directory(Some(path));
    }
}

async fn session_cleanup_loop(state: AppState) {
    let mut ticker = tokio::time::interval(state.cleanup_interval);
    ticker.set_missed_tick_behavior(tokio::time::MissedTickBehavior::Skip);
    loop {
        ticker.tick().await;
        cleanup_expired_sessions(&state).await;
    }
}

async fn session_path(state: &AppState, session_id: &str) -> Result<PathBuf, ApiError> {
    let mut engine = state.engine.lock().await;
    let path = engine
        .files
        .get(session_id)
        .cloned()
        .ok_or_else(|| (StatusCode::NOT_FOUND, "Unknown workbook session.".to_string()))?;
    engine
        .last_access
        .insert(session_id.to_string(), Instant::now());
    Ok(path)
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

fn workspace_key(address: SocketAddr) -> String {
    address
        .to_string()
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() || character == '-' || character == '_' {
                character
            } else {
                '_'
            }
        })
        .collect()
}

fn prepare_workspace(address: SocketAddr) -> Result<EngineWorkspace, io::Error> {
    let base = match std::env::var("XLSX_ENGINE_WORK_ROOT") {
        Ok(path) if !path.trim().is_empty() => PathBuf::from(path),
        Ok(_) => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "XLSX_ENGINE_WORK_ROOT must not be empty.",
            ))
        }
        Err(std::env::VarError::NotPresent) => {
            std::env::temp_dir().join("genoffice-xlsx-engine-v2")
        }
        Err(error) => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("Unable to read XLSX_ENGINE_WORK_ROOT: {error}"),
            ))
        }
    };
    let root = base.join(workspace_key(address));

    // main() binds the listener before reaching this function. Therefore no
    // other healthy Engine can own this exact endpoint while we remove leftovers
    // from a previous crashed process. Other ports/addresses use different roots.
    if root.exists() {
        fs::remove_dir_all(&root)?;
    }
    let workbook_root = root.join("workbooks");
    let scratch_root = root.join("scratch");
    fs::create_dir_all(&workbook_root)?;
    fs::create_dir_all(&scratch_root)?;

    Ok(EngineWorkspace {
        root,
        workbook_root,
        scratch_root,
    })
}

fn workbook_directory(state: &AppState) -> PathBuf {
    state
        .workbook_root
        .join(Uuid::new_v4().to_string())
}

fn scratch_directory(state: &AppState, kind: &str) -> PathBuf {
    state
        .scratch_root
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

fn configured_limit_bytes(name: &str, default_mb: usize) -> Result<usize, io::Error> {
    let megabytes = match std::env::var(name) {
        Ok(raw) => raw.parse::<usize>().map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("{name} must be a positive integer number of MiB."),
            )
        })?,
        Err(std::env::VarError::NotPresent) => default_mb,
        Err(error) => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("Unable to read {name}: {error}"),
            ))
        }
    };

    if megabytes == 0 {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("{name} must be greater than zero."),
        ));
    }

    megabytes.checked_mul(1024 * 1024).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("{name} is too large for this platform."),
        )
    })
}

fn configured_seconds(name: &str, default_seconds: u64) -> Result<Duration, io::Error> {
    let seconds = match std::env::var(name) {
        Ok(raw) => raw.parse::<u64>().map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("{name} must be a positive integer number of seconds."),
            )
        })?,
        Err(std::env::VarError::NotPresent) => default_seconds,
        Err(error) => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("Unable to read {name}: {error}"),
            ))
        }
    };

    if seconds == 0 {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("{name} must be greater than zero."),
        ));
    }
    Ok(Duration::from_secs(seconds))
}

fn internal_error(error: impl ToString) -> ApiError {
    (StatusCode::INTERNAL_SERVER_ERROR, error.to_string())
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let listen_addr = std::env::var("XLSX_ENGINE_LISTEN")
        .unwrap_or_else(|_| "127.0.0.1:7301".to_string());
    let address: SocketAddr = listen_addr.parse()?;
    let max_workbook_bytes =
        configured_limit_bytes("XLSX_ENGINE_MAX_WORKBOOK_MB", DEFAULT_MAX_WORKBOOK_MB)?;
    let max_request_bytes =
        configured_limit_bytes("XLSX_ENGINE_MAX_REQUEST_MB", DEFAULT_MAX_REQUEST_MB)?;
    if max_request_bytes < max_workbook_bytes {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "XLSX_ENGINE_MAX_REQUEST_MB must be greater than or equal to XLSX_ENGINE_MAX_WORKBOOK_MB.",
        )
        .into());
    }
    let session_ttl = configured_seconds("XLSX_ENGINE_SESSION_TTL_SECS", DEFAULT_SESSION_TTL_SECS)?;
    let cleanup_interval = configured_seconds(
        "XLSX_ENGINE_CLEANUP_INTERVAL_SECS",
        DEFAULT_CLEANUP_INTERVAL_SECS,
    )?;
    if cleanup_interval > session_ttl {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "XLSX_ENGINE_CLEANUP_INTERVAL_SECS must be less than or equal to XLSX_ENGINE_SESSION_TTL_SECS.",
        )
        .into());
    }

    // Bind first. Only after the endpoint is exclusively ours is it safe to
    // remove this endpoint's workspace left by an ungraceful previous process.
    let listener = tokio::net::TcpListener::bind(address).await?;
    let workspace = prepare_workspace(address)?;
    let workspace_root = workspace.root.clone();

    let state = AppState {
        sessions: Arc::new(MemorySessionStore::default()),
        engine: Arc::new(Mutex::new(EngineState::default())),
        max_workbook_bytes,
        session_ttl,
        cleanup_interval,
        workbook_root: workspace.workbook_root,
        scratch_root: workspace.scratch_root,
    };
    tokio::spawn(session_cleanup_loop(state.clone()));

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
            "/v1/sessions/{session_id}/recalc",
            post(recalc_workbook),
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
        .layer(DefaultBodyLimit::max(max_request_bytes))
        .with_state(state);

    println!(
        "xlsx-engine-service listening on http://{address} (workbook max {}MB, request max {}MB, session ttl {}s)",
        max_workbook_bytes / (1024 * 1024),
        max_request_bytes / (1024 * 1024),
        session_ttl.as_secs()
    );

    let result = axum::serve(listener, app)
        .with_graceful_shutdown(shutdown_signal())
        .await;

    // Graceful shutdown leaves no workspace. If the process is killed, the
    // next process that successfully binds this endpoint removes it at startup.
    let _ = fs::remove_dir_all(&workspace_root);
    result?;
    Ok(())
}

async fn shutdown_signal() {
    let _ = tokio::signal::ctrl_c().await;
}
