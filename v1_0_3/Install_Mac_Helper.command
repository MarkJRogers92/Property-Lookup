#!/bin/zsh
set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEST="$HOME/Library/Application Scripts/com.microsoft.Excel"
HELPER="$SCRIPT_DIR/CookPropertyHTTP.applescript"
WORKBOOK="$SCRIPT_DIR/Cook_Property_Due_Diligence_v1_0_3_MAC_WINDOWS.xlsx"

echo ""
echo "Cook County Property Due Diligence - Mac helper installer"
echo "----------------------------------------------------------"
echo ""

if [[ ! -f "$HELPER" ]]; then
  echo "ERROR: CookPropertyHTTP.applescript is not next to this installer."
  echo "Keep the files together in the unzipped package and try again."
  read "?Press Return to close..."
  exit 1
fi

mkdir -p "$DEST"
cp "$HELPER" "$DEST/CookPropertyHTTP.applescript"
chmod 644 "$DEST/CookPropertyHTTP.applescript"

echo "Installed Mac Excel helper:"
echo "  $DEST/CookPropertyHTTP.applescript"
echo ""

if /usr/bin/curl --fail --silent --show-error --location --connect-timeout 8 --max-time 12 \
  'https://gis.cookcountyil.gov/traditional/rest/services/CookViewer3Parcels/MapServer/0?f=json' \
  >/dev/null 2>&1; then
  echo "Cook County connection test: OK"
else
  echo "Cook County connection test: could not confirm right now."
  echo "The helper is still installed; the workbook will report individual source failures."
fi

echo ""
echo "NEXT - one time only:"
echo "1. Excel will open the workbook."
echo "2. Save it as an Excel Macro-Enabled Workbook (.xlsm)."
echo "3. In Excel: Tools > Macro > Visual Basic Editor."
echo "4. Import CookPropertyDueDiligence_v1_0_3_CROSS_PLATFORM.bas."
echo "5. Debug > Compile VBAProject, save, close, and reopen with macros enabled."
echo "6. The blue RUN PROPERTY RESEARCH button will appear automatically."
echo ""

if [[ -f "$WORKBOOK" ]]; then
  open -a 'Microsoft Excel' "$WORKBOOK" || open "$WORKBOOK"
fi

read "?Press Return to close this installer..."
