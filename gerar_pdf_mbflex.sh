#!/usr/bin/env bash
# Gera o PDF leave-behind da proposta MBFlex (16:9, 8 páginas).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
HTML="$ROOT/proposta-mbflex-pdf.html"
OUT="$ROOT/MBFlex-Express_Assessoria-Growth_V4.pdf"
CHROME="${CHROME:-google-chrome}"

"$CHROME" --headless=new --disable-gpu --no-sandbox --disable-dev-shm-usage \
  --no-pdf-header-footer --hide-scrollbars --allow-file-access-from-files \
  --virtual-time-budget=8000 \
  --run-all-compositor-stages-before-draw \
  --print-to-pdf="$OUT" \
  "file://$HTML"

echo "wrote $OUT ($(wc -c < "$OUT") bytes)"
