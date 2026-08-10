# Web Office Foundation

This branch introduces the Web Office foundation for running GenOffice editors in browsers and embedding them into a host system through iframe-based integration.

Initial scope:

- Keep the existing React editor UIs.
- Preserve the current Electron applications while adding Web entry points.
- Add a host API abstraction for file open/save, file picking, locale, dirty state, and lifecycle events.
- Add a versioned iframe message protocol and bridge.
- Prioritize Docs and Slides Web support first.
- Keep Sheets Web UI separate from the XLSX processing service design.
- Avoid changes to `docx-engine`, `pptx-engine`, and `pptx-render` unless browser compatibility requires them.

Development branch: `agent/web-office-foundation`.
