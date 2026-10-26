#!/usr/bin/env bash
# Recalculates the workbook headless in LibreOffice, then checks every output against the
# Python reference model.  Usage: tools/recalc_and_verify.sh
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PY="${PYTHON:-$ROOT/.venv/bin/python}"
SOFFICE="${SOFFICE:-$(command -v soffice || echo /Applications/LibreOffice.app/Contents/MacOS/soffice)}"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

# Throwaway LibreOffice profile that always recalculates .xlsx files on load
mkdir -p "$WORK/profile/user"
cat > "$WORK/profile/user/registrymodifications.xcu" <<'XCU'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
XCU

"$SOFFICE" -env:UserInstallation="file://$WORK/profile" --headless --norestore \
  --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WORK/out" \
  "$ROOT/model/FPA_Budget_Variance_Forecast_Model.xlsx" >/dev/null 2>&1

cd "$ROOT/tools"
"$PY" verify_workbook.py "$WORK/out/FPA_Budget_Variance_Forecast_Model.xlsx"
