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
