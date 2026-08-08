# GenOffice Docs Web iframe Host Demo

This example simulates the future Web OS host around the standalone Docs Web app.

## Run

From the repository root:

```bash
npm ci
npm run dev:web:docs-host
```

Open <http://127.0.0.1:8080>.

The host expects Docs Web at <http://127.0.0.1:5273> by default. Override it with:

```text
http://127.0.0.1:8080/?docsUrl=http://127.0.0.1:5273
```

## Manual acceptance flow

1. Click **打开 DOCX** and choose a `.docx` file.
2. Confirm the document renders inside the iframe.
3. Edit content and confirm the host status changes to `dirty`.
4. Click **保存**.
5. Wait for the host status to become `saved`.
6. Click **下载当前 DOCX** and open the downloaded file in Microsoft Word.
7. Use **切换为只读** / **切换为编辑** to verify host-controlled editor mode.
8. Use an editor command that inserts an image. The editor sends `office:pick-file`; the host opens its own file-selection dialog. In the real Web OS this dialog will be replaced by the system file manager instead of a local browser picker.

## Host responsibilities demonstrated

- validate iframe `event.source` and `event.origin`
- send `office:init` only after `office:ready`
- provide document bytes and version metadata
- receive `office:save-document` and acknowledge with `office:save-document-result`
- receive dirty/title state
- trigger parent-requested save
- control `view` / `edit` mode
- service `office:pick-file` and `office:read-file`

The editor remains unaware of Web OS storage APIs. The production Web OS adapter should replace only the host-side file picker, read, save, permission, and version implementations.
