#!/usr/bin/env bash
set -euo pipefail

BASE_URL="${1:-http://127.0.0.1:7301}"
BASE_URL="${BASE_URL%/}"
REQUEST_ID="deploy-smoke-$(date +%s)"
TMP_DIR="$(mktemp -d)"
trap 'rm -rf "$TMP_DIR"' EXIT

fail() {
  printf 'FAIL: %s\n' "$1" >&2
  exit 1
}

printf 'Checking XLSX Engine at %s\n' "$BASE_URL"

curl -fsS \
  -D "$TMP_DIR/health.headers" \
  -H "X-Request-Id: $REQUEST_ID" \
  "$BASE_URL/health" \
  > "$TMP_DIR/health.json"

grep -qi "^x-request-id: ${REQUEST_ID}" "$TMP_DIR/health.headers" \
  || fail 'health response did not echo X-Request-Id'

node - "$TMP_DIR/health.json" <<'NODE'
const fs = require('fs')
const path = process.argv[2]
const value = JSON.parse(fs.readFileSync(path, 'utf8'))
if (value.ok !== true || value.service !== 'xlsx-engine-service') process.exit(1)
for (const field of ['maxHeavyRequests', 'availableHeavySlots', 'heavyQueueTimeoutSecs']) {
  if (!Number.isFinite(value[field])) process.exit(1)
}
NODE
printf '  health/request-id: ok\n'

curl -fsS "$BASE_URL/metrics" > "$TMP_DIR/metrics.txt"
for metric in \
  genoffice_xlsx_requests_total \
  genoffice_xlsx_server_errors_total \
  genoffice_xlsx_heavy_admission_rejects_total \
  genoffice_xlsx_heavy_slots \
  genoffice_xlsx_heavy_slots_available \
  genoffice_xlsx_workbook_sessions
do
  grep -q "^${metric} " "$TMP_DIR/metrics.txt" || fail "missing metric ${metric}"
done
printf '  metrics: ok\n'

curl -fsS -X POST \
  "$BASE_URL/v1/workbooks/blank?name=deployment-smoke.xlsx" \
  > "$TMP_DIR/blank.json"

SESSION_ID="$(node - "$TMP_DIR/blank.json" <<'NODE'
const fs = require('fs')
const value = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'))
if (!value.sessionId || !Array.isArray(value.sheets) || value.sheets.length === 0) process.exit(1)
process.stdout.write(value.sessionId)
NODE
)"

curl -fsS \
  -H "X-Xlsx-Session: $SESSION_ID" \
  "$BASE_URL/v1/sessions/$SESSION_ID" \
  > "$TMP_DIR/session.json"
printf '  blank workbook/session: ok\n'

STATUS="$(curl -sS \
  -o /dev/null \
  -w '%{http_code}' \
  -X DELETE \
  -H "X-Xlsx-Session: $SESSION_ID" \
  "$BASE_URL/v1/sessions/$SESSION_ID")"
[[ "$STATUS" == "204" ]] || fail "session cleanup returned HTTP ${STATUS}"
printf '  explicit session cleanup: ok\n'

STATUS="$(curl -sS \
  -o "$TMP_DIR/deleted.txt" \
  -w '%{http_code}' \
  -H "X-Xlsx-Session: $SESSION_ID" \
  "$BASE_URL/v1/sessions/$SESSION_ID")"
[[ "$STATUS" == "404" ]] || fail "deleted session remained reachable (HTTP ${STATUS})"
printf '  deleted-session isolation: ok\n'

printf 'PASS: XLSX Engine deployment smoke checks succeeded.\n'
