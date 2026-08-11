# UC Web OS XLSX Host

Reference host for embedding GenOffice Sheets Web inside the existing UC Web OS plugin iframe.

## Topology

```text
UC Web OS
└─ UC Excel plugin iframe (this example)
   └─ GenOffice Sheets iframe
      └─ /xlsx-engine/* -> Rust XLSX Engine
```

The UC plugin remains the platform boundary. GenOffice Sheets never receives UC tenant IDs, JWTs, FsNode permissions or storage APIs directly.

## UC RPC used

The host talks to its parent with the existing UC plugin RPC envelope:

```text
uc-plugin-rpc-request
uc-plugin-rpc-response
```

The implementation uses the currently established file APIs:

- `uc.ready`
- `uc.host.getLaunchParams`
- `uc.fs.requestSelectedFileAccess`
- `uc.fs.readSelectedFile`
- `uc.fs.pickSaveDestination`
- `uc.fs.saveResultFile`

Normal save requests `writeMode: 'selected'`. Save As first chooses a destination, then requests `writeMode: 'result'`. A successful Save As must return a new `nodeId`/`id`; otherwise the host rejects the result so later Ctrl+S cannot accidentally write back to the original file.

### Optimistic version protection

The file version returned when the workbook is opened is carried into GenOffice as the editor's `baseVersion`. Before a normal Save, the UC Host requests fresh selected-file access and compares the latest `version`/`fileVersion` with that base version.

- same version → `uc.fs.saveResultFile` may run;
- different version → the Host returns `VERSION_CONFLICT` **before** `saveResultFile` is called;
- missing version on either side → backward-compatible save behavior is retained;
- Save As is exempt from the original file's version comparison because it creates a result file instead of overwriting the selected file.

After a successful save, the version returned by UC becomes the next `baseVersion`. This prevents a later Ctrl+S from silently overwriting changes made by another editor/session.

## GenOffice protocol

The nested editor uses only the shared `office:*` iframe protocol:

- `office:init`
- `office:pick-file`
- `office:read-file`
- `office:save-document`
- dirty/title/state events

The loaded UC `blob` is converted to an `ArrayBuffer` and transferred to Sheets through `office:init`. Save bytes travel in the opposite direction and are wrapped in an XLSX Blob for `uc.fs.saveResultFile`.

## Embedded URL

The plugin accepts:

- `sheetsUrl` — Sheets Web URL, default `http://127.0.0.1:5275`
- `ucHostOrigin` — UC Web OS parent origin; when omitted, the plugin tries `document.referrer`
- `pluginId` — UC plugin ID, default `thirdparty.plugin.excel-online`
- `locale` — fallback locale, default `zh-CN`

Example:

```text
http://127.0.0.1:8083/?sheetsUrl=http://127.0.0.1:5275&ucHostOrigin=http://127.0.0.1:5173&pluginId=thirdparty.plugin.excel-online
```

## Current picker boundary

The confirmed UC plugin contract used by the existing Office host does not yet provide a confirmed interactive open-file picker RPC for choosing an arbitrary second file from inside Office. Therefore `office:pick-file` is deliberately isolated in `openLocalAssetPicker()` and currently uses a browser file input, matching the existing UC Office-plugin fallback.

This does **not** affect opening the workbook that launched the plugin, normal save, or Save As — those already use UC file APIs. Once the final UC interactive picker API is fixed, replace only `openLocalAssetPicker()`; the Sheets editor and `office:*` protocol do not change.

## Development

```bash
npm run dev:uc-webos-xlsx-host
```

Build:

```bash
npm run build:uc-webos-xlsx-host
```
