#!/bin/sh
set -eu

output=/etc/nginx/conf.d/10-xlsx-upstreams.conf
servers=${XLSX_ENGINE_SERVERS:-}
secure=${XLSX_ENGINE_STICKY_COOKIE_SECURE:-false}

if [ -z "$servers" ]; then
  echo "[genoffice-web] XLSX_ENGINE_SERVERS is required, for example: 10.0.0.21:7301,10.0.0.22:7301" >&2
  exit 1
fi

case "$secure" in
  1|true|TRUE|yes|YES)
    secure_attr=" secure"
    ;;
  0|false|FALSE|no|NO|"")
    secure_attr=""
    ;;
  *)
    echo "[genoffice-web] XLSX_ENGINE_STICKY_COOKIE_SECURE must be true or false" >&2
    exit 1
    ;;
esac

server_lines=$(printf '%s' "$servers" | tr ',' '\n')
server_count=0
tmp="${output}.tmp"

{
  echo "upstream genoffice_xlsx_engine {"
  echo "    least_conn;"
  printf '    sticky cookie genoffice_xlsx_route path=/xlsx-engine/ httponly samesite=lax%s;\n' "$secure_attr"

  while IFS= read -r raw; do
    server=$(printf '%s' "$raw" | tr -d '[:space:]')
    if [ -z "$server" ]; then
      echo "[genoffice-web] XLSX_ENGINE_SERVERS contains an empty endpoint" >&2
      exit 1
    fi

    if ! printf '%s\n' "$server" | grep -Eq '^([A-Za-z0-9._-]+|\[[0-9A-Fa-f:]+\]):[0-9]{1,5}$'; then
      echo "[genoffice-web] invalid XLSX engine endpoint: $server; expected host:port" >&2
      exit 1
    fi

    port=${server##*:}
    if [ "$port" -lt 1 ] || [ "$port" -gt 65535 ]; then
      echo "[genoffice-web] invalid XLSX engine port in endpoint: $server" >&2
      exit 1
    fi

    printf '    server %s max_fails=2 fail_timeout=10s;\n' "$server"
    server_count=$((server_count + 1))
  done <<EOF
$server_lines
EOF

  echo "}"
} > "$tmp"

if [ "$server_count" -lt 1 ]; then
  rm -f "$tmp"
  echo "[genoffice-web] XLSX_ENGINE_SERVERS must contain at least one endpoint" >&2
  exit 1
fi

mv "$tmp" "$output"
echo "[genoffice-web] configured $server_count XLSX engine upstream(s)"
