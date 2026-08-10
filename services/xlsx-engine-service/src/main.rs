use std::{collections::HashMap, net::SocketAddr, sync::Arc};

use axum::{
    extract::{Path, State},
    http::StatusCode,
    routing::{delete, get, post},
    Json, Router,
};
use serde::{Deserialize, Serialize};
use tokio::sync::RwLock;
use uuid::Uuid;

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

#[derive(Clone)]
struct AppState {
    sessions: Arc<MemorySessionStore>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CreateSessionRequest {
    #[serde(default)]
    source: Option<String>,
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

    // Keep the identifier opaque to browser clients. A future sharded store can
    // change its internal placement strategy without changing the HTTP contract.
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

async fn delete_session(
    State(state): State<AppState>,
    Path(session_id): Path<String>,
) -> StatusCode {
    if state.sessions.sessions.write().await.remove(&session_id).is_some() {
        StatusCode::NO_CONTENT
    } else {
        StatusCode::NOT_FOUND
    }
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let listen_addr = std::env::var("XLSX_ENGINE_LISTEN")
        .unwrap_or_else(|_| "127.0.0.1:7301".to_string());
    let address: SocketAddr = listen_addr.parse()?;

    let state = AppState {
        sessions: Arc::new(MemorySessionStore::default()),
    };

    let app = Router::new()
        .route("/health", get(health))
        .route("/v1/sessions", post(create_session))
        .route("/v1/sessions/{session_id}", delete(delete_session))
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
