FROM rust:1.89-bookworm AS builder

WORKDIR /src

COPY services/xlsx-engine-service/Cargo.toml services/xlsx-engine-service/Cargo.toml
COPY services/xlsx-engine-service/src services/xlsx-engine-service/src
COPY apps/sheets/native/xlsx-engine/Cargo.toml apps/sheets/native/xlsx-engine/Cargo.toml
COPY apps/sheets/native/xlsx-engine/src apps/sheets/native/xlsx-engine/src

RUN cargo build --release --manifest-path services/xlsx-engine-service/Cargo.toml

FROM debian:bookworm-slim AS runtime

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl \
    && rm -rf /var/lib/apt/lists/* \
    && useradd --system --uid 10001 --create-home --home-dir /var/lib/genoffice-xlsx-engine genoffice \
    && mkdir -p /var/lib/genoffice-xlsx-engine/work \
    && chown -R genoffice:genoffice /var/lib/genoffice-xlsx-engine

COPY --from=builder \
    /src/services/xlsx-engine-service/target/release/xlsx-engine-service \
    /usr/local/bin/xlsx-engine-service

USER genoffice

ENV XLSX_ENGINE_LISTEN=0.0.0.0:7301 \
    XLSX_ENGINE_WORK_ROOT=/var/lib/genoffice-xlsx-engine/work \
    XLSX_ENGINE_SESSION_TTL_SECS=3600 \
    XLSX_ENGINE_CLEANUP_INTERVAL_SECS=60 \
    XLSX_ENGINE_MAX_HEAVY_REQUESTS=4 \
    XLSX_ENGINE_HEAVY_QUEUE_TIMEOUT_SECS=15 \
    XLSX_ENGINE_MAX_WORKBOOK_MB=100 \
    XLSX_ENGINE_MAX_REQUEST_MB=384

EXPOSE 7301

HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
    CMD curl --fail --silent --show-error http://127.0.0.1:7301/health > /dev/null || exit 1

ENTRYPOINT ["/usr/local/bin/xlsx-engine-service"]
