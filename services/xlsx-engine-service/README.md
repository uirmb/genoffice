# XLSX Engine Service

The Web Sheets runtime keeps the existing React + Univer editor in the browser and moves XLSX parsing/preservation/export behind this Rust HTTP service.

## Development

```bash
npm run dev:xlsx-engine
```

The default listener is `127.0.0.1:7301`. Override it with `XLSX_ENGINE_LISTEN`.

The Sheets Vite development and preview servers proxy same-origin browser calls from `/xlsx-engine/*` to this service.

## Current API

- `GET /health`
- `POST /v1/sessions`
- `POST /v1/workbooks`
- `POST /v1/workbooks/blank`
- `GET /v1/sessions/:sessionId`
- `DELETE /v1/sessions/:sessionId`
- `POST /v1/sessions/:sessionId/ranges`
- `POST /v1/sessions/:sessionId/formulas`
- `POST /v1/sessions/:sessionId/recalc`
- `GET /v1/sessions/:sessionId/archive/manifest`
- `POST /v1/sessions/:sessionId/archive/read`
- `POST /v1/sessions/:sessionId/archive/scan`
- `POST /v1/sessions/:sessionId/archive/save`

Workbook APIs are session-addressed. Browser clients treat `sessionId` as an opaque string and may send it in `X-Xlsx-Session` as a future routing key.

The service directly reuses the existing `xlsx-sidecar` workbook sessions and IronCalc recalculation support. Browser preservation-save planning stays in the shared Sheets gateway, while this service owns workbook/session access and archive assembly.

## Request limits

The service keeps HTTP bodies bounded in production while leaving enough room for preservation saves, whose archive mutations carry base64-encoded package parts.

- `XLSX_ENGINE_MAX_WORKBOOK_MB` — maximum raw `.xlsx` upload size, default `100` MiB.
- `XLSX_ENGINE_MAX_REQUEST_MB` — maximum HTTP request body size, default `384` MiB.

`XLSX_ENGINE_MAX_REQUEST_MB` must be greater than or equal to `XLSX_ENGINE_MAX_WORKBOOK_MB`. Invalid or zero values fail service startup rather than silently disabling the protection. Raw workbook uploads that exceed their configured limit return HTTP `413 Payload Too Large`.

The request limit intentionally exceeds the raw workbook limit because base64 content expands binary payloads by roughly one third and a preservation save can contain multiple replaced or added package parts.

## Session expiry

Workbook sessions are intentionally in-memory for the first single-node milestone, but abandoned browser sessions are no longer allowed to live forever.

- `XLSX_ENGINE_SESSION_TTL_SECS` — idle workbook/session lifetime, default `3600` seconds.
- `XLSX_ENGINE_CLEANUP_INTERVAL_SECS` — expired-session sweep interval, default `60` seconds.

The cleanup interval must be less than or equal to the session TTL. Invalid or zero values fail service startup.

Metadata reads, range/formula reads, recalculation and archive operations refresh a workbook session's last-access timestamp. A preservation save registers the newly produced workbook session with a fresh TTL. When a session expires, the service closes the native workbook session, removes metadata and recalculation cache entries, and deletes the session's temporary workbook directory. Explicit `DELETE /v1/sessions/:sessionId` performs the same cleanup immediately.

## Production deployment

The first production topology is intentionally single-node:

```text
Nginx
├─ /                -> Sheets Web static files
└─ /xlsx-engine/*   -> 127.0.0.1:7301
```

No Redis, database, object storage, or message queue is required for the first milestone.

The service boundary is intentionally prepared for later horizontal expansion without changing Sheets Web, UC Excel Host, or the `office:*` iframe protocol. Session placement and routing can be introduced behind the same API when multiple Rust instances are needed.

The Rust service must remain independent of UC Web OS authentication and storage concepts: it does not receive tenant IDs, JWTs, CSRF tokens, FsNode IDs, or plugin permissions. UC owns files and authorization; the engine owns spreadsheet processing.

The remaining single-node hardening work is crash/startup orphan-directory cleanup, concurrency/time-budget controls, and production metrics/structured logging. Those concerns stay inside the engine service and do not change the browser or UC Host contracts.
