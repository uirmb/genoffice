# GenOffice Docs Web iframe Host Demo

This example simulates the future Web OS host around the Docs Web app.

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

## Word Web product policy

The demo sends production-like `office:init.capabilities`:

- legacy Genspark AI ribbon actions, context actions, and the AI side panel are hidden
- editor-owned AutoSave is hidden; autosave policy belongs to the host
- Open / Save / Save As remain available as document commands
- Save writes back to the current host file
- Save As is explicitly marked as `mode: 'saveAs'`; this demo prompts for a new name and creates a new host file id
- editing-area corner marks are enabled as a visual-only page overlay
- file/image selection is host-owned

Electron/desktop behavior is not changed by this Web-only policy.

## Manual acceptance flow

1. Click **打开 DOCX** and choose a `.docx` file.
2. Confirm the document renders inside the iframe and the Genspark AI UI / AutoSave toggle are absent.
3. Confirm the four corner marks identify the body editing area inside the page margins.
4. Edit content and confirm the host status changes to `dirty`.
5. Click **保存** and wait for the host status to become `saved`.
6. Edit again, use **文件 → 另存为**, choose a new name, and confirm the Host Demo switches to the new document identity.
7. Click **下载当前 DOCX** and open the downloaded file in Microsoft Word.
8. Use **切换为只读** / **切换为编辑** to verify host-controlled editor mode.
9. Use an editor command that inserts an image. The editor sends `office:pick-file`; the host opens its own file-selection dialog. In the real Web OS this dialog will be replaced by the system file manager instead of a local browser picker.

## Host responsibilities demonstrated

- validate iframe `event.source` and `event.origin`
- send `office:init` only after `office:ready`
- provide document bytes, version metadata, editor mode, and host capabilities
- receive `office:save-document` and acknowledge with `office:save-document-result`
- distinguish current-file Save from host-owned Save As
- receive dirty/title state
- trigger parent-requested save
- control `view` / `edit` mode
- service `office:pick-file` and `office:read-file`

The editor remains unaware of Web OS storage APIs. The production Web OS adapter should replace only the host-side file picker, read, save/save-as, permission, and version implementations.
