# XLSX Engine Service

The Web Sheets runtime keeps the existing React + Univer editor in the browser and moves XLSX parsing/preservation/export behind this Rust HTTP service.

## Development

```bash
npm run dev:xlsx-engine
```

The default listener is `127.0.0.1:7301`. Override it with `XLSX_ENGINE_LISTEN`.

The Sheets Vite development server proxies same-origin browser calls from `/xlsx-engine/*` to this service.

## Initial API

- `GET /health`
- `POST /v1/sessions`
- `DELETE /v1/sessions/:sessionId`

All future workbook APIs are session-addressed. Browser clients must treat `sessionId` as an opaque string and may send it in `X-Xlsx-Session` as a future routing key.

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
