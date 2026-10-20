#!/usr/bin/env bash
set -euo pipefail

repo_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cache_dir="${VEFEYAT_CACHE_DIR:-$(dirname "$repo_dir")/.cache}"
cd "$repo_dir"
umask 077
mkdir -p .local "$cache_dir/uv" "$cache_dir/npm"
uv --cache-dir "$cache_dir/uv" sync --frozen
npm --cache "$cache_dir/npm" ci --no-audit --no-fund
node scripts/service-control.mjs start-db
.venv/bin/python scripts/wait-db.py
.venv/bin/python manage.py migrate --noinput
node scripts/service-control.mjs start-web
.venv/bin/python scripts/wait-web.py
printf 'Vefeyât hazır: http://127.0.0.1:8000\n'
