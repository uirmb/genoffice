# Web Office Foundation

This branch family introduces the Web Office foundation for running GenOffice editors in browsers and embedding them into a host system through iframe-based integration.

Current architecture:

- Keep the existing React editor UIs and preserve the Electron applications.
- Add Web entry points behind the shared `OfficeHostApi` / versioned `office:*` iframe protocol.
- Docs and Slides Web reuse their browser-compatible document engines.
- Sheets Web reuses the existing React + Univer renderer while XLSX parsing, workbook sessions, formula evaluation and archive assembly stay behind the independent Rust XLSX engine service.
- Browser Sheets calls only same-origin `/xlsx-engine/*`; Vite/Nginx owns the reverse proxy.
- UC Web OS owns authentication, FsNode/storage permissions, file open/save and platform pickers. The Rust engine never receives UC JWTs, tenants, users, FsNode IDs or plugin permissions.
- Standalone hosts may implement the same Office Host API with ordinary browser file pickers.

## Sheets Web status

The Excel Web foundation now has permanent TypeScript, Rust, compatibility and Chromium gates covering:

- blank workbook creation;
- real `.xlsx` open and lazy range reads;
- iframe `office:new` / `office:init` lifecycle;
- formula discovery and IronCalc recalculation;
- preservation save with formula cached-value persistence while retaining `<f>` formulas;
- filters and row visibility;
- hyperlinks;
- conditional formatting and data validation;
- sheet protection;
- page setup;
- legacy notes/comments and VML relationship creation;
- sheet add, duplicate, rename, remove, hide/unhide and reorder;
- structural row/column insert, remove and move;
- row/column size, hidden and outline state;
- merge/unmerge and reference movement;
- defined names;
- linked chart edits;
- workbook image/media reads plus image/visual insertion and anchor edits;
- native tables;
- x14 sparklines;
- Pivot definition reads;
- Pivot cache `refreshOnLoad`;
- native Pivot creation;
- existing Pivot output-layout refresh/expansion, including fail-closed checks when newly occupied cells already contain ordinary worksheet data;
- host save/download and reopen verification.

All major renderer XLSX save journals now have a Web preservation path and browser coverage. Unsupported document constructs discovered by the shared planner continue to fail closed rather than producing a partially corrupted package.

## Host file integration

Sheets Web supports both buffer and token-backed Host files.

Permanent Chromium coverage includes:

- Host-selected workbook open;
- `office:pick-file -> token -> office:read-file` for image insertion;
- normal save through the Host;
- clean-workbook Save As without a synthetic edit;
- Web `Ctrl/Cmd+S`, `Ctrl/Cmd+Shift+S` and `Ctrl/Cmd+O` routing through the same renderer menu-action path instead of browser page commands;
- UC-style optimistic version checks where a stale normal Save returns `VERSION_CONFLICT` before the final Host write; Save As remains independent of the source version.

The reference UC bridge lives in `examples/uc-webos-xlsx-host`. Opening the file that launched the plugin, normal Save and Save As use the established UC file RPCs. The only remaining platform API gap is a confirmed UC interactive picker for choosing an arbitrary second system file from inside Office; the generic `office:pick-file` boundary is already isolated so the editor does not need to change when that UC RPC is finalized.

## XLSX Engine production baseline

The Rust service now includes:

- configurable raw-workbook and total-request size limits;
- idle session TTL and cleanup;
- endpoint-isolated workspaces with crash/startup and graceful-shutdown cleanup;
- bounded heavy-work admission with a pre-execution queue timeout;
- `X-Request-Id` correlation;
- JSON request logs containing operational fields only;
- low-cardinality Prometheus-text `/metrics`;
- no UC platform coupling.

A single-node deployment baseline is under `deploy/xlsx-engine/` with environment, systemd, Nginx and rollout/smoke-check examples.

## Compatibility gates

In addition to feature-specific Chromium tests, `compat:web-engine` runs the repository's five generated XLSX compatibility fixtures through the real Web production save architecture:

```text
shared planner
  -> Rust archive manifest/read/scan
  -> Rust archive save
  -> new workbook session
  -> decompressed OOXML entry SHA256 comparison
```

Only entries explicitly touched/added/removed by the mutation plan may differ. This is a generated regression corpus, not a claim that the files were produced by Excel, WPS or LibreOffice; real-application files should be added as a separate corpus when available.

## Remaining milestone work

The Excel Web foundation itself is substantially complete. Remaining work is primarily release/platform work:

- integrate the UC Host bridge into the actual UC Web OS plugin source;
- replace the temporary browser fallback for arbitrary second-file selection once the final UC system-picker RPC is confirmed;
- audit and remediate production dependency vulnerabilities;
- measure/tune Web bundle loading rather than adding speculative chunking;
- add a real Excel/WPS/LibreOffice compatibility corpus;
- tune Engine admission/session limits with production telemetry;
- add a hard in-flight Engine execution deadline only if XLSX operations gain cooperative cancellation or move behind an isolated worker-process boundary.

Development branches:

- `agent/web-office-foundation`
- `agent/ppt-web-foundation`
- `agent/xlsx-web-foundation`
