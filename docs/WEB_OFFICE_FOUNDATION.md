# Web Office Foundation

This branch family introduces the Web Office foundation for running GenOffice editors in browsers and embedding them into a host system through iframe-based integration.

Current scope:

- Keep the existing React editor UIs.
- Preserve the current Electron applications while adding Web entry points.
- Add a host API abstraction for file open/save, file picking, locale, dirty state, and lifecycle events.
- Add a versioned iframe message protocol and bridge.
- Docs and Slides Web reuse their existing browser-compatible document engines.
- Sheets Web reuses the existing React + Univer renderer while XLSX parsing, session access, formula evaluation, and package preservation stay behind the Rust XLSX engine service.
- Sheets Web uses same-origin `/xlsx-engine/*` calls so deployment and later engine routing remain host-controlled.
- Keep UC/Web OS authentication, storage, tenant, and permission concepts outside the XLSX engine service.

## Sheets Web status

The Excel Web foundation now has a production-build Chromium round-trip covering:

- blank workbook creation;
- real `.xlsx` open and lazy range reads;
- iframe `office:new` / `office:init` lifecycle;
- cell-edit preservation save;
- formula discovery;
- IronCalc recalculation;
- recalculated formula cached-value persistence while preserving `<f>` formulas;
- workbook image reads from the session archive with a 20MB preview limit;
- host save state and download of the saved workbook;
- reopen/read and downloaded-package verification after save.

The remaining Sheets Web work is primarily higher-level save journal coverage (sheet/structure, filters, hyperlinks, conditional formatting, data validation, notes, protection, page setup, charts/visuals, tables, pivots, sparklines, defined names), plus the remaining platform-facing capabilities such as pivot definition reads, local-image insertion, export, and final UC Web OS host integration.

Development branches:

- `agent/web-office-foundation`
- `agent/ppt-web-foundation`
- `agent/xlsx-web-foundation`
