#!/bin/bash
# McHenry Property Lookup -- one-click launcher (macOS)
#
# Double-click this file in Finder. First run builds the local database
# (10-30 seconds), every run after that just starts the app and opens
# it in your browser.

set -e
cd "$(dirname "$0")"

APP_DIR="$(pwd)"
DATA_DIR="$APP_DIR/data/Mchenry"
DB_PATH="$APP_DIR/mchenry_property.db"
PORT=8000
REQUIRED_CSVS=(
  "McHenry_County_TaxParcels.csv"
  "Planning_Landuse.csv"
  "Address_Points.csv"
  "Subdivisions.csv"
  "Road_Centerlines.csv"
)

echo "McHenry Property Lookup"
echo "========================"
echo

press_key_to_close() {
  read -n 1 -s -r -p "Press any key to close this window..."
  echo
}

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 is required but wasn't found on this Mac."
  echo "Install it from https://www.python.org/downloads/macos/ and then double-click this file again."
  press_key_to_close
  exit 1
fi

# If the app is already running (e.g. this was double-clicked twice),
# just reopen it in the browser instead of starting a second copy.
if curl -s -o /dev/null "http://localhost:$PORT/"; then
  echo "Already running -- opening in your browser."
  open "http://localhost:$PORT/"
  exit 0
fi

have_data() {
  for f in "${REQUIRED_CSVS[@]}"; do
    [ -f "$DATA_DIR/$f" ] || return 1
  done
  return 0
}

if ! have_data; then
  ZIP_FILE=$(find "$APP_DIR" -maxdepth 1 -iname "*.zip" | head -n 1)
  if [ -n "$ZIP_FILE" ]; then
    echo "Found data export: $(basename "$ZIP_FILE") -- unzipping..."
    TMP_EXTRACT=$(mktemp -d)
    unzip -q -o "$ZIP_FILE" -d "$TMP_EXTRACT"
    mkdir -p "$DATA_DIR"
    for f in "${REQUIRED_CSVS[@]}"; do
      SRC=$(find "$TMP_EXTRACT" -iname "$f" | head -n 1)
      if [ -n "$SRC" ]; then
        cp "$SRC" "$DATA_DIR/$f"
      fi
    done
    rm -rf "$TMP_EXTRACT"
    echo
  fi
fi

if ! have_data; then
  echo "Missing the McHenry County GIS data export."
  echo
  echo "Put the county's data export zip file in this folder:"
  echo "  $APP_DIR"
  echo "then double-click this file again -- it will unzip automatically."
  echo
  echo "(Or place the CSVs directly in: $DATA_DIR)"
  press_key_to_close
  exit 1
fi

if [ ! -f "$DB_PATH" ]; then
  echo "First run -- building the local database (about 10-30 seconds)..."
  python3 "$APP_DIR/build_database.py" --data-dir "$DATA_DIR" --out "$DB_PATH"
  echo
fi

echo "Starting the app at http://localhost:$PORT/ ..."
python3 "$APP_DIR/server.py" --db "$DB_PATH" --port "$PORT" &
SERVER_PID=$!

cleanup() {
  echo
  echo "Stopping McHenry Property Lookup..."
  kill "$SERVER_PID" 2>/dev/null
}
trap cleanup EXIT

sleep 1
open "http://localhost:$PORT/"

echo
echo "McHenry Property Lookup is running."
echo "Leave this window open while you use it. Close this window or press Ctrl+C to stop."
wait "$SERVER_PID"
