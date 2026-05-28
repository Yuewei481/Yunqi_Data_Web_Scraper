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

if [[ $# -ne 1 ]]; then
  echo "Usage: $0 base.xlsx" >&2
  exit 2
fi

BASE_XLSX="$1"
TODAY_DATE="$(date +%F)"
OUTPUT_DIR="${OUTPUT_DIR:-outputs/yunqi-pop-up-greeting-card}"
if [[ "$OUTPUT_DIR" != /* ]]; then
  OUTPUT_DIR="$SCRIPT_DIR/$OUTPUT_DIR"
fi
TODAY_XLSX="$OUTPUT_DIR/选品表格-pop-up-greeting-card-${TODAY_DATE}.xlsx"
EXISTING_IDS_JSON="$OUTPUT_DIR/existing-product-ids-for-append.json"
PYTHON="${CODEX_PYTHON:-${PYTHON:-python3}}"

mkdir -p "$OUTPUT_DIR"
"$PYTHON" append_yunqi_new_products.py --export-ids "$BASE_XLSX" "$EXISTING_IDS_JSON"
EXISTING_PRODUCT_IDS_PATH="$EXISTING_IDS_JSON" ./run_yunqi_scraper.sh
"$PYTHON" append_yunqi_new_products.py "$BASE_XLSX" "$TODAY_XLSX"
rm -f "$TODAY_XLSX" "$EXISTING_IDS_JSON"
