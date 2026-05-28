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

if [[ $# -lt 1 || $# -gt 2 ]]; then
  echo "Usage: $0 yesterday.xlsx [diff_output.xlsx]" >&2
  exit 2
fi

YESTERDAY_XLSX="$1"
TODAY_DATE="$(date +%F)"
OUTPUT_DIR="${OUTPUT_DIR:-outputs/yunqi-pop-up-greeting-card}"
if [[ "$OUTPUT_DIR" != /* ]]; then
  OUTPUT_DIR="$SCRIPT_DIR/$OUTPUT_DIR"
fi
TODAY_XLSX="$OUTPUT_DIR/选品表格-pop-up-greeting-card-${TODAY_DATE}.xlsx"
DIFF_XLSX="${2:-$OUTPUT_DIR/选品表格-pop-up-greeting-card-差异-${TODAY_DATE}.xlsx}"
PYTHON="${CODEX_PYTHON:-${PYTHON:-python3}}"

mkdir -p "$OUTPUT_DIR"
./run_yunqi_scraper.sh
"$PYTHON" compare_yunqi_daily_excel.py "$YESTERDAY_XLSX" "$TODAY_XLSX" "$DIFF_XLSX"
