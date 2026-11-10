#!/usr/bin/env bash
# Rebuilds the investor financial model from model/stage1.xlsx (the inputs and monthly engine).
# Needs: python3 + openpyxl, LibreOffice, and the xlsx skill's recalc.py (present in Claude sessions).
set -euo pipefail
cd "$(dirname "$0")"
RECALC_DIR="${RECALC_DIR:-/mnt/skills/public/xlsx/scripts}"
recalc() { (cd "$RECALC_DIR" && python3 recalc.py "$1" 180 | grep -E '"total_errors"|"total_formulas"'); }
mkdir -p sens output
echo "1/4 recalc inputs";           recalc "$PWD/stage1.xlsx"
echo "2/4 sensitivity snapshot";    python3 sens.py > /dev/null          # writes sens_results.json
echo "3/4 build sheets and text"
for s in stage3 post hybrid hybrid2 spaceplan clubsheet clubtext final_text; do python3 $s.py > /dev/null; done
echo "4/4 recalc final workbook";   recalc "$PWD/stage3b.xlsx"
cp stage3b.xlsx output/18_Golf_Performance_Club_Financial_Model.xlsx
python3 headline.py > ../docs/NUMBERS.md
echo "Done: model/output/18_Golf_Performance_Club_Financial_Model.xlsx and docs/NUMBERS.md"
