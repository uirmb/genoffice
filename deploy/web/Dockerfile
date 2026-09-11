FROM node:24.15.0-bookworm-slim AS builder

WORKDIR /src

COPY . .

RUN npm ci \
    && npm run build:web:all

FROM nginx:1.30.4-alpine AS runtime

RUN rm -f /etc/nginx/conf.d/default.conf

COPY deploy/web/nginx-genoffice.conf /etc/nginx/conf.d/20-genoffice.conf
COPY deploy/web/docker-entrypoint.d/15-genoffice-xlsx-upstreams.sh /docker-entrypoint.d/15-genoffice-xlsx-upstreams.sh
COPY --from=builder /src/dist/web/ /usr/share/nginx/html/

RUN chmod +x /docker-entrypoint.d/15-genoffice-xlsx-upstreams.sh

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
    CMD wget -q -O /dev/null http://127.0.0.1/healthz || exit 1
