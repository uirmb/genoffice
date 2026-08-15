# UC Web OS XLSX Host

Reference bridge for embedding GenOffice Sheets Web inside the UC Web OS plugin iframe.

## Topology

```text
UC Web OS / Host
└─ UC Office Bridge (this example)
   └─ GenOffice Sheets iframe
      └─ /xlsx-engine/* -> Rust XLSX Engine
```

UC remains the platform boundary. GenOffice Sheets never receives UC JWTs, file-access tokens, storage permissions or write locks directly. The XLSX Engine also remains independent of UC tenant/user/FsNode concerns.

## Stable UC RPC boundary

The Bridge uses the ordinary UC plugin RPC envelope:

```text
uc-plugin-rpc-request
uc-plugin-rpc-response
```

The stable Office file lifecycle maps to these UC capabilities:

- `uc.ready`
- `uc.host.getLaunchParams` (locale/mode only; file access does not depend on launch-path state)
- `uc.fs.readCurrentFile`
- `uc.fs.pickFile`
- `uc.fs.bindCurrentFile`
- `uc.fs.releasePickedFile`
- `uc.fs.pickAssets`
- `uc.fs.saveCurrentFile`
- `uc.fs.saveFileAs`
- `uc.fs.createFileVersion`
- `uc.download.saveFile`

The Bridge does **not** call `/plugins/file-access` or `/plugins/save`, does not retain UC access tokens, and does not implement a system picker or Save As destination state itself.

### Open document transaction

```text
office:pick-document
  -> uc.fs.pickFile
  -> selectionId + buffer
  -> editor parses/loads the workbook
  -> office:document-opened
  -> uc.fs.bindCurrentFile
```

If the editor cannot load the selected workbook it sends `office:document-open-failed`; the Bridge calls `uc.fs.releasePickedFile`. Until `bindCurrentFile` succeeds, the previously bound UC document remains current.

The old `office:pick-file` / `office:read-file` path remains a v1 compatibility alias only. Because that protocol has no editor-load acknowledgement, legacy document selection binds immediately; new Office code must not depend on that behavior.

### Insert assets

`office:pick-assets` maps directly to `uc.fs.pickAssets`.

Asset selection is read-only and buffer-based. It does not acquire a write token, bind a document, change the window title, or change the target of a later Ctrl/Cmd+S.

The reference Bridge no longer creates a local browser `<input type="file">` for embedded UC asset selection.

### Save

Normal save is one UC operation:

```text
office:save-document(mode=save)
  -> uc.fs.saveCurrentFile({ blob, filename, baseVersion })
```

Save As (including the first persistence of a new workbook) is also one UC operation:

```text
office:save-document(mode=saveAs | newDocument=true)
  -> uc.fs.saveFileAs({ blob, suggestedName, fileTypes, ...pickerText })
```

`saveFileAs` must create and return a new `nodeId` even if the chosen name equals the source name. File identity is never inferred from the name.

The Bridge no longer performs a separate `requestSelectedFileAccess -> saveResultFile` sequence and no longer performs optimistic-version checking itself. UC `saveCurrentFile` owns the short-lived write access, lock and `VERSION_CONFLICT` decision and returns the structured platform error code unchanged.

### History version

`office:save-history-version` maps to `uc.fs.createFileVersion`. A successful response must contain the actual latest file descriptor/version; GenOffice uses that version as the next save baseline.

### Local download

`office:download-document` maps to `uc.download.saveFile`.

Download is not UC persistence: it does not create/modify an FsNode and must not change current document identity, version, dirty state or window title. The legacy `office:export-document` message is translated to the same download behavior for protocol-v1 compatibility.

### Window close

Window ownership belongs to UC Host. The Bridge forwards stable `office:close-approved` / `office:close-cancelled` control messages to its parent rather than manipulating the parent DOM. The UC window layer should pair those messages with its existing `office:request-close` transaction/state machine.

## Initial document

After `office:ready`, the Bridge reads the already-bound file with `uc.fs.readCurrentFile` and sends `office:init` with `bytes` and `transport: 'buffer'`. If there is no bound file it sends `office:new`.

The editor must not issue a second read for `office:init` content.

## Embedded URL

The plugin accepts:

- `sheetsUrl` — Sheets Web URL, default `http://127.0.0.1:5275`
- `ucHostOrigin` — UC Web OS parent origin; when omitted, the Bridge tries `document.referrer`
- `pluginId` — UC plugin ID, default `thirdparty.plugin.excel-online`
- `locale` — fallback locale, default `zh-CN`

Example:

```text
http://127.0.0.1:8083/?sheetsUrl=http://127.0.0.1:5275&ucHostOrigin=http://127.0.0.1:5173&pluginId=thirdparty.plugin.excel-online
```

## Development

```bash
npm run dev:uc-webos-xlsx-host
```

Build:

```bash
npm run build:uc-webos-xlsx-host
```
