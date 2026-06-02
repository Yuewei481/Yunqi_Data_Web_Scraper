#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

SCRIPT_DIR="$(pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
CALLER_OUTPUT_DIR="${OUTPUT_DIR:-}"

is_absolute_path() {
  [[ "$1" == /* || "$1" =~ ^[A-Za-z]:[/\\] ]]
}

if [[ -f "$REPO_DIR/.env" ]]; then
  set -a
  # shellcheck disable=SC1091
  source "$REPO_DIR/.env"
  set +a
fi
if [[ -n "$CALLER_OUTPUT_DIR" ]]; then
  export OUTPUT_DIR="$CALLER_OUTPUT_DIR"
fi

export CODEX_PYTHON="${CODEX_PYTHON:-${PYTHON:-python3}}"
if [[ -z "${EXCEL_TEMPLATE:-}" ]]; then
  export EXCEL_TEMPLATE="$REPO_DIR/templates/选品表格-模板.xlsx"
elif ! is_absolute_path "$EXCEL_TEMPLATE"; then
  export EXCEL_TEMPLATE="$REPO_DIR/$EXCEL_TEMPLATE"
fi
if [[ -n "${OUTPUT_DIR:-}" ]] && ! is_absolute_path "$OUTPUT_DIR"; then
  export OUTPUT_DIR="$SCRIPT_DIR/$OUTPUT_DIR"
fi
NODE="${NODE:-node}"

"$NODE" yunqi_product_scraper.mjs
