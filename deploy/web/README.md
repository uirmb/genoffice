# GenOffice Web container

This directory packages all five browser applications into one Docker image with NGINX:

```text
/docx/
/xlsx/
/pptx/
/md/
/pdf/
/xlsx-engine/*
```

The Web image is independent from the XLSX Engine deployment. It receives the Engine endpoints at container startup through `XLSX_ENGINE_SERVERS`; no Engine hostname is embedded in the Vite bundle.

For the complete production topology, separate-server deployment, Engine deployment, validation and rollback procedures, see [`../README.md`](../README.md).

## Build

From the repository root:

```bash
docker build \
  -f deploy/web/Dockerfile \
  -t genoffice-web:0.1.0 \
  .
```

The builder runs `npm run build:web:all` and copies `dist/web/{docx,xlsx,pptx,md,pdf}` into the NGINX runtime image.

## Compose

```bash
cd deploy/web
cp .env.example .env
```

Set one Engine:

```dotenv
XLSX_ENGINE_SERVERS=10.10.1.21:7301
```

or multiple independent Engine servers:

```dotenv
XLSX_ENGINE_SERVERS=10.10.1.21:7301,10.10.1.22:7301,10.10.1.23:7301
```

Start:

```bash
docker compose up -d --build
```

Check:

```bash
docker compose ps
docker compose logs -f web
curl http://127.0.0.1:8080/healthz
```

Inspect the generated Engine upstream configuration:

```bash
docker compose exec web cat /etc/nginx/conf.d/10-xlsx-upstreams.conf
```

## Sticky session

The Engine currently owns workbook sessions in process memory. NGINX therefore applies cookie affinity on `/xlsx-engine/` requests. The startup script generates the upstream block from `XLSX_ENGINE_SERVERS` and uses the fixed cookie name `genoffice_xlsx_route`.

`XLSX_ENGINE_STICKY_COOKIE_SECURE=false` is convenient for HTTP testing. Set it to `true` when the browser-facing deployment uses HTTPS.

The cookie is intentionally a browser session cookie instead of using a fixed expiry because the Engine's session TTL is based on idle time.
