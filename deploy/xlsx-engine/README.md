# GenOffice XLSX Engine container

This directory deploys the Rust XLSX processing service as an independent Docker unit. It may run on the same server as `genoffice-web` or on one or more separate Engine servers.

For the complete topology, separate-server examples, Web deployment, validation and rollback procedures, see [`../README.md`](../README.md).

## Build

From the repository root:

```bash
docker build \
  -f deploy/xlsx-engine/Dockerfile \
  -t genoffice-xlsx-engine:0.1.0 \
  .
```

## Compose

```bash
cd deploy/xlsx-engine
cp .env.example .env
docker compose up -d --build
```

Check status and logs:

```bash
docker compose ps
docker compose logs -f xlsx-engine
curl http://127.0.0.1:7301/health
```

The image publishes port `7301` by default and has an image-level health check against `/health`.

## Runtime configuration

The Compose file exposes the Engine controls through `.env`:

```dotenv
XLSX_ENGINE_SESSION_TTL_SECS=3600
XLSX_ENGINE_CLEANUP_INTERVAL_SECS=60
XLSX_ENGINE_MAX_HEAVY_REQUESTS=4
XLSX_ENGINE_HEAVY_QUEUE_TIMEOUT_SECS=15
XLSX_ENGINE_MAX_WORKBOOK_MB=100
XLSX_ENGINE_MAX_REQUEST_MB=384
```

`XLSX_ENGINE_BIND_ADDRESS` controls which host interface publishes port 7301. In a separate-server deployment, firewall that port so it is reachable from the GenOffice Web server or Web ingress nodes, not directly from browsers.

## Session behavior

Workbook sessions are held in process memory. Restarting or removing the Engine invalidates sessions owned by that process. When multiple Engine instances are configured, the Web NGINX layer keeps a browser on the Engine that created its workbook sessions.

The work directory contains disposable processing state. UC remains the owner of user files and permissions, so this deployment does not require a persistent workbook volume in the current session model.

For planned maintenance, drain active XLSX work before terminating an Engine that may still own sessions.

## Local development without Docker

Development mode:

```bash
npm run dev:xlsx-engine
```

Release build:

```bash
npm run build:xlsx-engine
```

Equivalent Cargo command:

```bash
cargo build \
  --release \
  --manifest-path services/xlsx-engine-service/Cargo.toml
```

The default local listener is `127.0.0.1:7301`. Set `XLSX_ENGINE_LISTEN=0.0.0.0:7301` when another container or host must reach a locally running Engine process.
