# GenOffice Web deployment

GenOffice Web is deployed as two independent Docker units so the browser applications and the XLSX processing service can be placed, upgraded and scaled independently.

```text
UC Web OS / Browser
        |
        v
+---------------------------+
| genoffice-web             |
| NGINX + five Web apps     |
|                           |
| /docx/                    |
| /xlsx/                    |
| /pptx/                    |
| /md/                      |
| /pdf/                     |
| /xlsx-engine/* -----------+-------------------+
|     sticky cookie         |                   |
+---------------------------+                   |
                                                v
                               +-------------------------------+
                               | genoffice-xlsx-engine         |
                               | Rust XLSX processing service  |
                               | :7301                         |
                               +-------------------------------+
```

The two deployment units intentionally have separate Compose files:

```text
deploy/
├── web/
│   ├── Dockerfile
│   ├── compose.yaml
│   ├── .env.example
│   ├── nginx-genoffice.conf
│   └── docker-entrypoint.d/
│       └── 15-genoffice-xlsx-upstreams.sh
└── xlsx-engine/
    ├── Dockerfile
    ├── compose.yaml
    └── .env.example
```

They may run on the same Docker host, but production does not depend on a shared Compose project or Docker network. The Web deployment only needs TCP access to every configured XLSX Engine endpoint.

## 1. Build the images

Run Docker builds from the repository root because both images use the repository as their build context.

### Web image

```bash
docker build \
  -f deploy/web/Dockerfile \
  -t genoffice-web:0.1.0 \
  .
```

The Web image is a multi-stage image. The builder runs `npm run build:web:all`; the runtime contains NGINX and only the generated browser assets:

```text
/usr/share/nginx/html/
├── docx/
├── xlsx/
├── pptx/
├── md/
└── pdf/
```

The runtime does not require Node.js.

### XLSX Engine image

```bash
docker build \
  -f deploy/xlsx-engine/Dockerfile \
  -t genoffice-xlsx-engine:0.1.0 \
  .
```

The image compiles `services/xlsx-engine-service` in a Rust builder stage and copies the release binary into a small Debian runtime image.

## 2. Deploy XLSX Engine

On every Engine server:

```bash
cd deploy/xlsx-engine
cp .env.example .env
```

Review `.env` before starting. A normal single instance uses:

```dotenv
XLSX_ENGINE_IMAGE=genoffice-xlsx-engine:0.1.0
XLSX_ENGINE_BIND_ADDRESS=0.0.0.0
XLSX_ENGINE_PORT=7301

XLSX_ENGINE_SESSION_TTL_SECS=3600
XLSX_ENGINE_CLEANUP_INTERVAL_SECS=60
XLSX_ENGINE_MAX_HEAVY_REQUESTS=4
XLSX_ENGINE_HEAVY_QUEUE_TIMEOUT_SECS=15
XLSX_ENGINE_MAX_WORKBOOK_MB=100
XLSX_ENGINE_MAX_REQUEST_MB=384
```

Build and start from source:

```bash
docker compose up -d --build
```

If the image already exists in a registry or on the host:

```bash
docker compose pull
docker compose up -d --no-build
```

Check status and logs:

```bash
docker compose ps
docker compose logs -f xlsx-engine
curl http://127.0.0.1:7301/health
```

The container has its own Docker health check against `/health`.

### Multiple Engine servers

For horizontal capacity, deploy the same Engine Compose unit independently on multiple servers. Example:

```text
10.10.1.21:7301
10.10.1.22:7301
10.10.1.23:7301
```

Each instance keeps workbook sessions in its own process memory. Do not put a second generic load balancer between GenOffice Web and these Engine instances unless that layer preserves the same session affinity contract.

## 3. Deploy Web

On the Web server:

```bash
cd deploy/web
cp .env.example .env
```

Configure the Engine endpoints that are reachable from the Web container.

Single Engine:

```dotenv
XLSX_ENGINE_SERVERS=10.10.1.21:7301
```

Multiple Engines:

```dotenv
XLSX_ENGINE_SERVERS=10.10.1.21:7301,10.10.1.22:7301,10.10.1.23:7301
```

`XLSX_ENGINE_SERVERS` accepts only comma-separated `host:port` endpoints. Do not include `http://`, `https://` or URL paths.

For local HTTP testing:

```dotenv
XLSX_ENGINE_STICKY_COOKIE_SECURE=false
```

When browsers reach GenOffice through HTTPS, use:

```dotenv
XLSX_ENGINE_STICKY_COOKIE_SECURE=true
```

Start the Web deployment:

```bash
docker compose up -d --build
```

Or use a prebuilt image:

```bash
docker compose pull
docker compose up -d --no-build
```

Check status and logs:

```bash
docker compose ps
docker compose logs -f web
curl http://127.0.0.1:8080/healthz
```

Expected browser paths are:

```text
http://WEB_HOST:8080/docx/
http://WEB_HOST:8080/xlsx/
http://WEB_HOST:8080/pptx/
http://WEB_HOST:8080/md/
http://WEB_HOST:8080/pdf/
```

Sheets Web continues to call the same-origin path `/xlsx-engine/*`; no Engine hostname is embedded in the Vite build.

## 4. Runtime Engine routing

At container startup, `15-genoffice-xlsx-upstreams.sh` validates `XLSX_ENGINE_SERVERS` and writes `/etc/nginx/conf.d/10-xlsx-upstreams.conf`.

For three endpoints the generated configuration is equivalent to:

```nginx
upstream genoffice_xlsx_engine {
    least_conn;
    sticky cookie genoffice_xlsx_route path=/xlsx-engine/ httponly samesite=lax;

    server 10.10.1.21:7301 max_fails=2 fail_timeout=10s;
    server 10.10.1.22:7301 max_fails=2 fail_timeout=10s;
    server 10.10.1.23:7301 max_fails=2 fail_timeout=10s;
}
```

The cookie is a browser session cookie by default. It is intentionally not given a fixed one-hour expiry: the Engine session TTL is an idle timeout and active workbook sessions can live longer than one hour.

You can inspect the generated runtime configuration with:

```bash
docker compose exec web cat /etc/nginx/conf.d/10-xlsx-upstreams.conf
docker compose exec web nginx -T
```

## 5. Separate-server production example

Recommended network shape:

```text
Web server
10.10.1.10
  genoffice-web :8080
        |
        +----> 10.10.1.21:7301
        +----> 10.10.1.22:7301
        +----> 10.10.1.23:7301

Engine server 1               Engine server 2               Engine server 3
10.10.1.21                    10.10.1.22                    10.10.1.23
genoffice-xlsx-engine         genoffice-xlsx-engine         genoffice-xlsx-engine
:7301                         :7301                         :7301
```

Recommended firewall policy:

- expose the Web service only to the UC ingress / user network that needs it;
- allow TCP 7301 on Engine servers only from the Web server or Web ingress nodes;
- do not expose XLSX Engine directly to browsers or the public Internet.

The Engine does not implement UC authentication or file permissions. UC owns authorization and file storage; XLSX Engine is an internal document-processing service.

## 6. Same-server deployment with separate Compose projects

The two Compose files remain independent even on one host.

Start Engine first:

```bash
cd deploy/xlsx-engine
docker compose up -d --build
```

Then configure `deploy/web/.env` with an address by which the Web container can reach the host-published Engine port. Prefer the host's LAN/private address rather than `127.0.0.1`, because `127.0.0.1` inside `genoffice-web` refers to the Web container itself.

Then start Web:

```bash
cd ../web
docker compose up -d --build
```

An optional shared external Docker network can be introduced for a same-host installation, but production correctness does not depend on one because the two units are designed for cross-host deployment.

## 7. Manual Docker run

### Engine

```bash
docker run -d \
  --name genoffice-xlsx-engine \
  --restart unless-stopped \
  -p 7301:7301 \
  genoffice-xlsx-engine:0.1.0
```

### Web

```bash
docker run -d \
  --name genoffice-web \
  --restart unless-stopped \
  -p 8080:80 \
  -e XLSX_ENGINE_SERVERS=10.10.1.21:7301 \
  -e XLSX_ENGINE_STICKY_COOKIE_SECURE=false \
  genoffice-web:0.1.0
```

## 8. Local XLSX Engine development without Docker

Requirements:

```bash
rustc --version
cargo --version
```

Development mode:

```bash
npm run dev:xlsx-engine
```

or directly:

```bash
cargo run --manifest-path services/xlsx-engine-service/Cargo.toml
```

Release build:

```bash
npm run build:xlsx-engine
```

or:

```bash
cargo build \
  --release \
  --manifest-path services/xlsx-engine-service/Cargo.toml
```

Windows output:

```text
services/xlsx-engine-service/target/release/xlsx-engine-service.exe
```

Linux/macOS output:

```text
services/xlsx-engine-service/target/release/xlsx-engine-service
```

Default local listener is `127.0.0.1:7301`. To accept connections from another container or host, set `XLSX_ENGINE_LISTEN=0.0.0.0:7301` before starting the process.

## 9. Validation

Validate Web routes:

```bash
curl -I http://WEB_HOST:8080/docx/
curl -I http://WEB_HOST:8080/xlsx/
curl -I http://WEB_HOST:8080/pptx/
curl -I http://WEB_HOST:8080/md/
curl -I http://WEB_HOST:8080/pdf/
```

Validate Engine through the Web proxy and save the sticky cookie:

```bash
curl -i \
  -c genoffice-cookies.txt \
  http://WEB_HOST:8080/xlsx-engine/health

curl -i \
  -b genoffice-cookies.txt \
  http://WEB_HOST:8080/xlsx-engine/health
```

Then complete the UC product smoke test:

1. DOCX: open, edit, save, close and reopen.
2. XLSX: open, edit cells, save, close and reopen.
3. XLSX: open two or more workbooks and continue editing long enough to exercise session affinity.
4. PPTX: open, edit, save, close and reopen.
5. Markdown: edit, save, and test the unsaved-close three-option flow.
6. PDF: open and verify navigation/search.

## 10. Operations and failure semantics

XLSX workbook sessions remain in process memory in this deployment phase.

- Restarting or removing an Engine loses the active sessions owned by that Engine.
- Sticky routing cannot reconstruct a lost session on another Engine.
- Drain an Engine before planned maintenance or scale-down.
- A Web restart is safe as long as its Engine endpoint list remains stable; browsers may receive a new routing cookie on the next Engine request.
- Keep all Web replicas configured with the same Engine endpoint set and ordering.

If an Engine endpoint is temporarily unreachable, DOCX, PPTX, Markdown and PDF remain available; XLSX Engine calls through `/xlsx-engine/*` return a gateway error until an Engine can serve the request.

## 11. TLS

The `genoffice-web` container listens on HTTP port 80. Production TLS may terminate at the host reverse proxy, cloud load balancer or ingress layer in front of this container.

When the browser-facing URL is HTTPS, set:

```dotenv
XLSX_ENGINE_STICKY_COOKIE_SECURE=true
```

Keep UC Web OS iframe / CSP origin configuration at the browser-facing ingress layer where the real production hostname is known.

## 12. Release and rollback

Web and XLSX Engine are independent deployment units and may have independent image versions, for example:

```text
genoffice-web:0.1.3
genoffice-xlsx-engine:0.1.1
```

Before deploying an incompatible Engine API change, update and test the Web/Engine compatibility contract first.

For rollback, keep the previously tested image tags available:

```bash
docker compose down
# change *_IMAGE in .env back to the previous tag
docker compose up -d --no-build
```

Do not expect active XLSX sessions to survive an Engine rollback or restart.
