# GenOffice Web unified deployment

This deployment keeps the five browser applications as static files and containerizes only the stateful XLSX Engine.

Production URLs are expected to be:

```text
https://web.office.com/docx/
https://web.office.com/xlsx/
https://web.office.com/pptx/
https://web.office.com/md/
https://web.office.com/pdf/
https://web.office.com/xlsx-engine/*
```

## 1. Build the unified Web bundle

From the repository root:

```bash
npm ci
npm run build:web:all
```

The cross-platform build script injects a production Vite base path for each application and assembles one deployable directory:

```text
dist/web/
├── docx/
├── xlsx/
├── pptx/
├── md/
└── pdf/
```

Individual `npm run dev:web:*` commands still use `/` as their base path, so local development and the existing Host/E2E flows keep their current URLs.

Install the static bundle, for example:

```bash
sudo mkdir -p /opt/genoffice/web
sudo rsync -a --delete dist/web/ /opt/genoffice/web/
```

The Web bundle is static. It does not require Node.js or an application process on the production host after the build finishes.

## 2. Build the XLSX Engine image

Build from the repository root so the Docker build context contains both the service and its local `xlsx-sidecar` dependency:

```bash
docker build \
  -f deploy/xlsx-engine/Dockerfile \
  -t genoffice-xlsx-engine:0.1.0 \
  .
```

The image uses a Rust builder stage and a small Debian runtime stage. Node.js, Vite and the Web applications are not included in the Engine image.

## 3. Run one or more Engine containers

A three-instance single-host example:

```bash
docker run -d \
  --name genoffice-xlsx-1 \
  --restart unless-stopped \
  -p 127.0.0.1:7301:7301 \
  genoffice-xlsx-engine:0.1.0

docker run -d \
  --name genoffice-xlsx-2 \
  --restart unless-stopped \
  -p 127.0.0.1:7302:7301 \
  genoffice-xlsx-engine:0.1.0

docker run -d \
  --name genoffice-xlsx-3 \
  --restart unless-stopped \
  -p 127.0.0.1:7303:7301 \
  genoffice-xlsx-engine:0.1.0
```

Tune Engine limits with environment variables when required, for example:

```bash
-e XLSX_ENGINE_MAX_HEAVY_REQUESTS=4 \
-e XLSX_ENGINE_SESSION_TTL_SECS=3600
```

The work directory lives inside the container by default. Workbook sessions are temporary processing state, not UC file storage, so the first deployment milestone does not require a persistent volume.

## 4. Configure NGINX

Use `nginx-genoffice.conf.example` as the baseline.

Cookie-based upstream affinity requires NGINX Open Source 1.29.6 or newer. NGINX 1.30 stable or newer is recommended for production.

The important split is:

```text
/docx/  /xlsx/  /pptx/  /md/  /pdf/  -> /opt/genoffice/web static files
/xlsx-engine/*                         -> xlsx-engine Docker upstream
```

The current XLSX Engine keeps `WorkbookSessions`, recalculation state, metadata and workbook paths in the local process. A workbook session must therefore keep returning to the Engine instance that created it.

Phase 1 achieves that with NGINX sticky-cookie affinity. The existing `X-Xlsx-Session` request header remains unchanged and is reserved for a later session-aware router.

Important operational consequence: if an Engine container that owns active sessions is restarted, removed or becomes unhealthy, those in-memory sessions are lost. NGINX may send later requests to another Engine, but that Engine cannot recover the old session. Drain instances before planned removal and treat active workbook sessions as disposable processing state.

## 5. Validate the deployment

Check the five static applications:

```bash
curl -I https://web.office.com/docx/
curl -I https://web.office.com/xlsx/
curl -I https://web.office.com/pptx/
curl -I https://web.office.com/md/
curl -I https://web.office.com/pdf/
```

Check the Engine path through NGINX and capture the sticky cookie:

```bash
curl -i -c /tmp/genoffice-cookies.txt \
  https://web.office.com/xlsx-engine/health

curl -i -b /tmp/genoffice-cookies.txt \
  https://web.office.com/xlsx-engine/health
```

Then complete the product smoke test from UC Web OS:

1. Open a DOCX, edit, save and reopen it.
2. Open an XLSX, edit cells, recalculate if applicable, save and reopen it.
3. Open a PPTX, edit, save and reopen it.
4. Open a Markdown file, edit, save and close it through the UC Host flow.
5. Open a PDF and verify viewer navigation/search.
6. Open more than one XLSX and confirm repeated Engine requests remain healthy while multiple Engine containers are running.

## 6. Release and rollback rule

Build the static Web bundle and the XLSX Engine image from the same tested Git commit and version them as one GenOffice Web release.

Keep the previous `dist/web` bundle and Engine image tag available. A rollback should restore both from the same previous release so the browser and Engine API remain a tested pair.
