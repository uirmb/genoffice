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
- `GET /metrics`
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

## Heavy request admission

Workbook parsing, range/formula reads, recalculation and archive operations are admitted through a bounded semaphore before the expensive work starts.

- `XLSX_ENGINE_MAX_HEAVY_REQUESTS` — maximum admitted heavy requests, default `4`.
- `XLSX_ENGINE_HEAVY_QUEUE_TIMEOUT_SECS` — maximum time a request may wait for a heavy-work slot, default `15` seconds.

Both values must be positive integers. A request that cannot obtain a slot within the configured queue timeout returns HTTP `503 Service Unavailable` **before** its XLSX operation starts.

The admission timeout is intentionally not an in-flight execution timeout. Most sidecar/archive operations are synchronous today; aborting only the HTTP future would not reliably stop the underlying file operation and could leave a half-finished save. Once a request receives a slot, it is allowed to finish atomically. A future hard execution deadline should be added only together with cooperative cancellation or an isolated worker-process boundary.

`GET /health` reports `maxHeavyRequests`, `availableHeavySlots` and `heavyQueueTimeoutSecs` so a deployment can observe whether the configured pool is saturated.

## Observability

Every HTTP response carries `X-Request-Id`.

- If a caller supplies a safe `X-Request-Id` containing only letters, digits, `-`, `_`, `.`, or `:` and no more than 128 characters, the service preserves it.
- Otherwise the service generates an opaque `req_<uuid>` value.

Each completed request writes one JSON line to stdout with only operational fields:

```json
{"event":"http_request","requestId":"req_...","method":"POST","path":"/v1/workbooks","status":201,"durationMs":42}
```

The log intentionally records the URL **path only**. It does not log query strings, workbook names, request bodies, file bytes, UC users, tenants, FsNode IDs, JWTs, or plugin permissions.

Service startup also emits a JSON `service_started` event containing only listener and Engine limit configuration.

`GET /metrics` exposes a small Prometheus-text-compatible operational surface:

- `genoffice_xlsx_requests_total`
- `genoffice_xlsx_server_errors_total`
- `genoffice_xlsx_heavy_admission_rejects_total`
- `genoffice_xlsx_heavy_slots`
- `genoffice_xlsx_heavy_slots_available`
- `genoffice_xlsx_workbook_sessions`
- `genoffice_xlsx_lightweight_sessions`

The endpoint intentionally has no user-, tenant-, workbook-name-, or file-level labels, keeping metric cardinality bounded and avoiding storage metadata leakage. Production ingress may restrict `/metrics` to the internal monitoring network while keeping `/health` available to the load balancer.

## Session expiry

Workbook sessions are intentionally in-memory for the first single-node milestone, but abandoned browser sessions are no longer allowed to live forever.

- `XLSX_ENGINE_SESSION_TTL_SECS` — idle workbook/session lifetime, default `3600` seconds.
- `XLSX_ENGINE_CLEANUP_INTERVAL_SECS` — expired-session sweep interval, default `60` seconds.

The cleanup interval must be less than or equal to the session TTL. Invalid or zero values fail service startup.

Metadata reads, range/formula reads, recalculation and archive operations refresh a workbook session's last-access timestamp. A preservation save registers the newly produced workbook session with a fresh TTL. When a session expires, the service closes the native workbook session, removes metadata and recalculation cache entries, and deletes the session's temporary workbook directory. Explicit `DELETE /v1/sessions/:sessionId` performs the same cleanup immediately.

## Workspace isolation and crash recovery

Workbook and scratch files live under an endpoint-specific workspace. The default base directory is the operating system temp directory plus `genoffice-xlsx-engine-v2`; it can be overridden with `XLSX_ENGINE_WORK_ROOT`.

For example, `127.0.0.1:7301` uses a root similar to:

```text
<work-root>/127_0_0_1_7301/
├─ workbooks/
└─ scratch/
```

The service binds its TCP listener before touching this directory. Only after the endpoint is exclusively owned does startup remove leftovers from a previous crashed process and create a clean workspace. Different listen addresses/ports use different roots and are not deleted by each other. A graceful shutdown removes the current endpoint workspace immediately; an ungraceful process exit leaves it for the next successful owner of that endpoint to clean.

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

The remaining single-node hardening work is primarily deployment tuning and, if required by real workloads, a cooperative/worker-process execution deadline. Those concerns stay inside the engine service and do not change the browser or UC Host contracts.
