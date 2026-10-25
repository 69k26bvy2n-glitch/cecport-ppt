#!/bin/bash
# QA 渲染器（PowerPoint 容器 PDF → PDFKit PNG），工程持久版
set -euo pipefail
INPUT="$1"; OUTDIR="$2"; WIDTH="${3:-1600}"
ABS_IN="$(cd "$(dirname "$INPUT")" && pwd)/$(basename "$INPUT")"
mkdir -p "$OUTDIR"; ABS_OUT="$(cd "$OUTDIR" && pwd)"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PPT_WORK=""; trap 'rm -rf "${PPT_WORK:-}" 2>/dev/null || true' EXIT
CT="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/tmp/qa-render.pdf"
T0=$(date +%s)
osascript <<EOF || true
with timeout of 300 seconds
  tell application "Microsoft PowerPoint"
      activate
      open POSIX file "$ABS_IN"
      delay 4
      save active presentation in (POSIX file "$CT") as save as PDF
      delay 3
      close active presentation saving no
  end tell
end timeout
EOF
if [ ! -s "$CT" ] || [ "$(stat -f %m "$CT" 2>/dev/null || echo 0)" -lt $((T0 - 3)) ]; then
  echo "PPT 未产出新 PDF"; exit 3
fi
mkdir -p "$ABS_OUT/slides"
PPT_WORK="$(mktemp /tmp/qa-render.XXXXXX)"
cp "$CT" "$PPT_WORK"
osascript -l JavaScript "$SCRIPT_DIR/pdf2png.js" "$PPT_WORK" "$ABS_OUT/slides" "$WIDTH" >/dev/null
echo "done: $ABS_OUT/slides/"
