# GenOffice Excel Web Host Demo

This example hosts GenOffice Sheets Web in an iframe and exercises the versioned `office:*` protocol.

It demonstrates the host responsibilities expected from UC Web OS or another embedding platform:

- create a blank workbook with `office:new`;
- open a real `.xlsx` with `office:init`;
- answer editor file-pick/read requests;
- receive dirty/title state;
- persist editor save bytes through `office:save-document`;
- switch edit/view mode and locale;
- download the currently saved workbook for round-trip verification.

## Development

```bash
npm run dev:xlsx-engine
npm run dev:web:sheets
npm run dev:web-xlsx-host
```

Default URLs:

- Sheets Web: `http://127.0.0.1:5275`
- XLSX Host Demo: `http://127.0.0.1:8082`
- XLSX Engine Service: `http://127.0.0.1:7301`

The host can point at another Sheets Web origin with the `sheetsUrl` query parameter. Sheets Web itself only calls same-origin `/xlsx-engine/*`; the serving layer proxies that path to the Rust engine.

The permanent Sheets Web GitHub Actions workflow builds the production Web surfaces, starts both previews plus the Rust engine, and runs a Chromium iframe round-trip against this host.
