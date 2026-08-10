# GenOffice PPT Web iframe Host Demo

This is a browser-only Host used to validate the Slides Web `office:*` integration contract before wiring the editor into UC Web OS.

## Run

From the repository root:

```bash
npm ci
npm run dev:web:slides-host
```

Then open:

```text
http://127.0.0.1:8081
```

Services:

- Slides Web: `http://127.0.0.1:5274`
- PPT Host Demo: `http://127.0.0.1:8081`

You can override the editor URL:

```text
http://127.0.0.1:8081/?slidesUrl=http://127.0.0.1:5274
```

## What it validates

- App Center style blank presentation through `office:new`
- real `.pptx` open through `office:init`
- Host-owned file/image picker through `office:pick-file`
- Save and first-save/new-document identity
- Save As creating a new Host file identity
- Save as history version
- Export current presentation as a local `.pptx` download without changing identity
- Host-owned Exit request
- runtime `view` / `edit` switching
- runtime locale switching
- dirty state reporting
- exact iframe source/origin validation

The demo stores files and history in memory. It is **not** a production storage layer. UC Web OS should replace the demo implementations while keeping the same `office:*` protocol.
