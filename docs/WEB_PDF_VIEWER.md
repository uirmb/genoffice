# GenOffice PDF Web — Viewer-only phase

PDF Web phase 1 is intentionally a read-only viewer. It is a separate browser entry point and does not replace or reduce the existing Electron PDF editor.

## Included

- UC/Office Host `office:init` with `kind: 'pdf'`
- transactional `office:pick-document` → parse → `office:document-opened` binding
- standalone browser PDF selection for development
- PDF.js canvas rendering and selectable text layer
- lazy page and thumbnail rendering
- page navigation and zoom
- fit width / fit page
- thumbnails and outline navigation
- PDF link annotations
- full-text search and match-page navigation
- clean close with permanently `dirty: false`

## Explicitly excluded

- PDF mutation and Save / Save As
- annotations and comments
- AcroForm editing
- signatures and stamps
- drawing
- image or text editing
- crop / cutout
- AI editing

These exclusions are enforced by `src/web/viewer-policy.ts` and its unit test. Web code must not import the Electron edit layers to add viewer functionality.

## Development

```bash
npm run dev:web:pdf
```

Standalone PDF Web runs on port `5276` by default.

For the iframe Host demo:

```bash
npm run dev:web:pdf-host
```

The demo Host runs on port `8083` and uses the same stable `office:*` file lifecycle used by the other Web Office applications.
