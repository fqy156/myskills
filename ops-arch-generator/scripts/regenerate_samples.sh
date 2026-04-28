#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
REPO_ROOT="$(cd "$ROOT_DIR/.." && pwd)"
GENERATOR="$REPO_ROOT/skills/ops-arch-generator/scripts/generate_architecture_ppt.py"
OUTPUT_DIR="$REPO_ROOT/skills/ops-arch-generator/outputs"

workbooks=(
  "$ROOT_DIR/zyqd-test.xlsx"
  "$ROOT_DIR/zyqd2.xlsx"
  "$ROOT_DIR/zyqd05.xlsx"
)

for workbook in "${workbooks[@]}"; do
  python3 "$GENERATOR" --workbook "$workbook" --output-dir "$OUTPUT_DIR"
done
