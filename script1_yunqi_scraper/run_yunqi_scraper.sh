#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

SCRIPT_DIR="$(pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

if [[ -f "$REPO_DIR/.env" ]]; then
  set -a
  # shellcheck disable=SC1091
  source "$REPO_DIR/.env"
  set +a
fi

export CODEX_PYTHON="${CODEX_PYTHON:-${PYTHON:-python3}}"
if [[ -z "${EXCEL_TEMPLATE:-}" ]]; then
  export EXCEL_TEMPLATE="$REPO_DIR/templates/选品表格-模板.xlsx"
elif [[ "$EXCEL_TEMPLATE" != /* ]]; then
  export EXCEL_TEMPLATE="$REPO_DIR/$EXCEL_TEMPLATE"
fi
NODE="${NODE:-node}"

"$NODE" yunqi_product_scraper.mjs
