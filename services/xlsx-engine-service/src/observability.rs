use std::{
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

use axum::{
    body::Body,
    extract::{Request, State},
    http::{header, HeaderValue},
    middleware::Next,
    response::Response,
};
use serde_json::json;
use uuid::Uuid;

use crate::AppState;

const REQUEST_ID_HEADER: &str = "x-request-id";

#[derive(Default)]
pub(crate) struct ServiceMetrics {
    requests_total: AtomicU64,
    server_errors_total: AtomicU64,
    heavy_admission_rejects_total: AtomicU64,
}

impl ServiceMetrics {
    pub(crate) fn record_heavy_admission_reject(&self) {
        self.heavy_admission_rejects_total
            .fetch_add(1, Ordering::Relaxed);
    }

    fn requests_total(&self) -> u64 {
        self.requests_total.load(Ordering::Relaxed)
    }

    fn server_errors_total(&self) -> u64 {
        self.server_errors_total.load(Ordering::Relaxed)
    }

    fn heavy_admission_rejects_total(&self) -> u64 {
        self.heavy_admission_rejects_total.load(Ordering::Relaxed)
    }
}

fn incoming_request_id(request: &Request) -> Option<String> {
    let value = request.headers().get(REQUEST_ID_HEADER)?.to_str().ok()?;
    if value.is_empty() || value.len() > 128 {
        return None;
    }
    if !value
        .bytes()
        .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.' | b':'))
    {
        return None;
    }
    Some(value.to_string())
}

fn generated_request_id() -> String {
    format!("req_{}", Uuid::new_v4().simple())
}

pub(crate) async fn request_observability(
    State(state): State<AppState>,
    request: Request,
    next: Next,
) -> Response {
    let request_id = incoming_request_id(&request).unwrap_or_else(generated_request_id);
    let method = request.method().as_str().to_string();
    let path = request.uri().path().to_string();
    let started = Instant::now();

    state.metrics.requests_total.fetch_add(1, Ordering::Relaxed);
    let mut response = next.run(request).await;
    let status = response.status();
    if status.is_server_error() {
        state
            .metrics
            .server_errors_total
            .fetch_add(1, Ordering::Relaxed);
    }

    if let Ok(value) = HeaderValue::from_str(&request_id) {
        response.headers_mut().insert(REQUEST_ID_HEADER, value);
    }

    println!(
        "{}",
        json!({
            "event": "http_request",
            "requestId": request_id,
            "method": method,
            "path": path,
            "status": status.as_u16(),
            "durationMs": started.elapsed().as_millis(),
        })
    );

    response
}

pub(crate) async fn metrics(State(state): State<AppState>) -> Response {
    let workbook_sessions = {
        let engine = state.engine.lock().await;
        engine.metadata.len()
    };
    let lightweight_sessions = state.sessions.sessions.read().await.len();

    let body = format!(
        concat!(
            "# HELP genoffice_xlsx_requests_total HTTP requests observed by the XLSX engine.\n",
            "# TYPE genoffice_xlsx_requests_total counter\n",
            "genoffice_xlsx_requests_total {}\n",
            "# HELP genoffice_xlsx_server_errors_total HTTP 5xx responses returned by the XLSX engine.\n",
            "# TYPE genoffice_xlsx_server_errors_total counter\n",
            "genoffice_xlsx_server_errors_total {}\n",
            "# HELP genoffice_xlsx_heavy_admission_rejects_total Heavy requests rejected before work started because the queue timed out.\n",
            "# TYPE genoffice_xlsx_heavy_admission_rejects_total counter\n",
            "genoffice_xlsx_heavy_admission_rejects_total {}\n",
            "# HELP genoffice_xlsx_heavy_slots Configured heavy-work admission slots.\n",
            "# TYPE genoffice_xlsx_heavy_slots gauge\n",
            "genoffice_xlsx_heavy_slots {}\n",
            "# HELP genoffice_xlsx_heavy_slots_available Heavy-work slots currently available.\n",
            "# TYPE genoffice_xlsx_heavy_slots_available gauge\n",
            "genoffice_xlsx_heavy_slots_available {}\n",
            "# HELP genoffice_xlsx_workbook_sessions Active workbook sessions.\n",
            "# TYPE genoffice_xlsx_workbook_sessions gauge\n",
            "genoffice_xlsx_workbook_sessions {}\n",
            "# HELP genoffice_xlsx_lightweight_sessions Active lightweight reserved sessions.\n",
            "# TYPE genoffice_xlsx_lightweight_sessions gauge\n",
            "genoffice_xlsx_lightweight_sessions {}\n"
        ),
        state.metrics.requests_total(),
        state.metrics.server_errors_total(),
        state.metrics.heavy_admission_rejects_total(),
        state.max_heavy_requests,
        state.heavy_slots.available_permits(),
        workbook_sessions,
        lightweight_sessions,
    );

    let mut response = Response::new(Body::from(body));
    response.headers_mut().insert(
        header::CONTENT_TYPE,
        HeaderValue::from_static("text/plain; version=0.0.4; charset=utf-8"),
    );
    response
}
