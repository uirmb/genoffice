# PDF Web Host Demo

A minimal iframe Host used to validate the GenOffice PDF Web viewer against the stable `office:*` protocol.

Run from the repository root:

```bash
npm run dev:web:pdf-host
```

The Host starts on `http://localhost:8083` and expects PDF Web on `http://localhost:5276` unless `?pdfUrl=` overrides it.

The demo intentionally exposes no save or editing capability. Initial files use `office:init`; editor-driven Open uses the transactional `office:pick-document` / `office:document-opened` lifecycle.
