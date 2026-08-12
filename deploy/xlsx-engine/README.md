# Sheets Web + XLSX Engine single-node deployment

This directory is the production-oriented baseline for the first GenOffice Sheets Web milestone.

The intended topology is:

```text
Browser / UC Web OS
        |
      HTTPS
        |
      Nginx
      |   \
      |    +-- /xlsx-engine/* -> 127.0.0.1:7301
      |
      +------- Sheets Web static files
                    |
                    +-- office:* iframe protocol when embedded by UC
```

The Rust process is deliberately private and UC-agnostic. It must not receive UC JWTs, tenant IDs, FsNode IDs, plugin permissions, or storage credentials.

## 1. Build

From a clean checkout:

```bash
npm ci
npm run build:web:sheets
npm run build:xlsx-engine
```

Outputs used by this runbook:

```text
apps/sheets/dist-web/
services/xlsx-engine-service/target/release/xlsx-engine-service
```

The UC reference Host can be built separately with:

```bash
npm run build:uc-webos-xlsx-host
```

It belongs inside the UC plugin/static-host deployment, not inside the Rust service.

## 2. Install files

Create a dedicated unprivileged account and directories using the operating system's normal account-management tools, then install approximately as follows:

```text
/opt/genoffice/sheets-web/              <- contents of apps/sheets/dist-web/
/opt/genoffice/bin/xlsx-engine-service  <- Rust release binary
/etc/genoffice/xlsx-engine.env          <- copy of xlsx-engine.env.example
/var/lib/genoffice-xlsx-engine/         <- owned by genoffice:genoffice
/etc/systemd/system/genoffice-xlsx-engine.service
/etc/nginx/conf.d/genoffice-sheets.conf
```

Recommended ownership:

```text
/opt/genoffice                       root:root, read-only to the service
/etc/genoffice                       root:root
/etc/genoffice/xlsx-engine.env       root:genoffice, mode 0640
/var/lib/genoffice-xlsx-engine       genoffice:genoffice
```

The systemd unit uses `ProtectSystem=strict` and permits writes only under `/var/lib/genoffice-xlsx-engine`.

## 3. Configure the Engine

Copy `xlsx-engine.env.example` to `/etc/genoffice/xlsx-engine.env` and tune it for the server.

Important defaults:

```text
raw workbook max        100 MiB
total request max       384 MiB
idle session TTL        3600 s
cleanup interval        60 s
heavy request slots     4
heavy queue timeout     15 s
```

Start conservatively. The current native workbook-session layer still serializes some work internally, so increasing the admission-slot count far above CPU capacity does not automatically improve throughput.

## 4. Configure Nginx

Copy `nginx-sheets.conf.example`, then change at least:

- `server_name`;
- TLS configuration;
- Sheets Web static root if different;
- `frame-ancestors` to the exact UC Web OS origin(s);
- `/metrics` allow-list if Prometheus is not local.

The example forwards Nginx `$request_id` as `X-Request-Id`. The Engine echoes the ID back and writes it into its JSON request log, which makes a browser/proxy request traceable without logging workbook data.

The public `/xlsx-engine/` prefix is stripped by Nginx. Sheets Web therefore continues calling same-origin URLs such as:

```text
/xlsx-engine/v1/workbooks
/xlsx-engine/v1/sessions/<opaque-session-id>/ranges
```

while Rust itself continues exposing `/v1/*`.

## 5. Start and verify

After installing the unit/configuration:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now genoffice-xlsx-engine
sudo systemctl reload nginx
```

Local Engine smoke checks:

```bash
curl -fsS http://127.0.0.1:7301/health
curl -fsS http://127.0.0.1:7301/metrics
```

Expected health fields include:

```text
ok
service
sessionStore
maxHeavyRequests
availableHeavySlots
heavyQueueTimeoutSecs
```

Check request correlation:

```bash
curl -i -H 'X-Request-Id: deployment-smoke-1' \
  http://127.0.0.1:7301/health
```

The response should contain:

```text
X-Request-Id: deployment-smoke-1
```

and stdout/journald should contain one JSON `http_request` event with the same ID.

## 6. Production checks before traffic

Verify all of these before enabling UC users:

1. Nginx can serve the Sheets Web `index.html` and hashed assets.
2. `/xlsx-engine/health` works through the same public origin as Sheets Web.
3. `/xlsx-engine/metrics` is not publicly reachable unless intentionally allowed.
4. An ordinary XLSX opens, edits, saves and reopens.
5. An unchanged workbook can Save As to a new file.
6. In UC, a stale normal Save produces `VERSION_CONFLICT` and does not call the final file-write API.
7. `Ctrl/Cmd+S`, `Ctrl/Cmd+Shift+S` and `Ctrl/Cmd+O` are intercepted by Sheets Web rather than by the browser.
8. Stopping the service removes its endpoint workspace; a forced crash is cleaned on the next successful bind of that same endpoint.
9. Request bodies beyond the configured raw workbook limit return HTTP 413.
10. Saturated heavy-work admission returns HTTP 503 before XLSX work starts.

## 7. Logs and metrics

Use journald or the platform log collector for the service stdout. Request logs contain operational metadata only:

```json
{"event":"http_request","requestId":"...","method":"POST","path":"/v1/workbooks","status":201,"durationMs":42}
```

Do not add workbook names, query strings, UC identifiers, or request bodies to these log lines.

The Prometheus-text metrics are intentionally low-cardinality and contain no user/file labels. At minimum alert on sustained:

- growth in `genoffice_xlsx_server_errors_total`;
- growth in `genoffice_xlsx_heavy_admission_rejects_total`;
- `genoffice_xlsx_heavy_slots_available` staying at zero;
- unexpectedly high `genoffice_xlsx_workbook_sessions` relative to expected active editors.

## 8. Upgrade / rollback rule

Treat the Sheets Web static bundle and Rust Engine as one release even though they deploy independently. Upgrade both from the same tested commit. Keep the previous static bundle and Engine binary available so rollback can restore the pair together.

The UC Host/iframe protocol and Rust `/v1/*` boundary are intentionally stable so future multi-node routing can be added behind these interfaces without changing the editor.
