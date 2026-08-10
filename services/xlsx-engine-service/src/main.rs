use std::{
    collections::HashMap,
    fs,
    net::SocketAddr,
    path::PathBuf,
    sync::Arc,
};

use axum::{
    body::Bytes,
    extract::{DefaultBodyLimit, Path, Query, State},
    http::StatusCode,
    routing::{delete, get, post},
    Json, Router,
};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use sha2::{Digest, Sha256};
use tokio::sync::{Mutex, RwLock};
use uuid::Uuid;
use xlsx_sidecar::{CellRange, WorkbookSessions};

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
}

impl Default for EngineState {
    fn default() -> Self {
        Self {
            workbooks: WorkbookSessions::new(),
            files: HashMap::new(),
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

    // Keep identifiers opaque to browser clients. A future sharded store can
    // change placement without changing the Sheets Web or UC Excel contract.
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

async fn open_workbook(
    State(state): State<AppState>,
    Query(query): Query<OpenWorkbookQuery>,
    bytes: Bytes,
) -> Result<(StatusCode, Json<Value>), ApiError> {
    let name = safe_workbook_name(query.name.as_deref().unwrap_or("workbook.xlsx"));
    let upload_id = Uuid::new_v4();
    let directory = std::env::temp_dir()
        .join("genoffice-xlsx-engine")
        .join(upload_id.to_string());
    fs::create_dir_all(&directory).map_err(internal_error)?;
    let path = directory.join(&name);
    fs::write(&path, &bytes).map_err(internal_error)?;

    let sha256 = format!("{:x}", Sha256::digest(&bytes));
    let mut engine = state.engine.lock().await;
    let metadata = match engine.workbooks.open(&path) {
        Ok(metadata) => metadata,
        Err(error) => {
            let _ = fs::remove_dir_all(&directory);
            return Err((StatusCode::UNPROCESSABLE_ENTITY, error.to_string()));
        }
    };
    let session_id = metadata.session_id.clone();
    engine.files.insert(session_id.clone(), path);

    let mut value = serde_json::to_value(metadata).map_err(internal_error)?;
    let object = value
        .as_object_mut()
        .ok_or_else(|| internal_error("Workbook metadata was not an object."))?;
    // Electron uses the local absolute path for CELL("filename"). In Web mode
    // this would reveal the server's temporary filesystem, so the browser only
    // receives Host-facing identity and workbook metadata.
    object.remove("path");
    object.insert("name".into(), Value::String(name));
    object.insert("sha256".into(), Value::String(sha256));
    object.insert("readOnly".into(), Value::Bool(false));
    object.insert("needsSaveAs".into(), Value::Bool(false));

    Ok((StatusCode::CREATED, Json(value)))
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
        .route("/v1/sessions/{session_id}/ranges", post(read_range))
        .route("/v1/sessions/{session_id}", delete(delete_session))
        // Do not impose an application-specific XLSX body limit here. Deployment
        // policy belongs to the reverse proxy/operator, not the engine contract.
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
