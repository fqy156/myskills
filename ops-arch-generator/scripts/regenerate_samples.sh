#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
GENERATOR="$ROOT_DIR/scripts/generate_architecture_ppt.py"
OUTPUT_DIR="$ROOT_DIR/outputs"

workbooks=(
  "$ROOT_DIR/zyqd-test.xlsx"
  "$ROOT_DIR/zyqd2.xlsx"
  "$ROOT_DIR/zyqd05.xlsx"
)

for workbook in "${workbooks[@]}"; do
  python3 "$GENERATOR" --workbook "$workbook" --output-dir "$OUTPUT_DIR"
done
