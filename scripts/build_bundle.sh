#!/usr/bin/env bash
# Build a clean, distributable .alfredworkflow from src/.
#
# Alfred's own "Export Workflow" zips the entire workflow folder, which sweeps
# in the gitignored config.json (real client_id/tenant_id) and token_cache.json
# (live OAuth tokens) — committing or sharing that leaks secrets. This script
# packages src/ WITHOUT those files (and without build junk), then refuses to
# emit a bundle that still contains any secret. Distribute the result by linking
# to it; a GitHub Release would be auto-published to the Alfred Gallery.
set -euo pipefail

cd "$(dirname "$0")/.."   # repo root
SRC="src"

VERSION=$(/usr/bin/python3 -c "import plistlib,sys;print(plistlib.load(open('$SRC/info.plist','rb'))['version'])")
OUT="releases/Outlook Suite_${VERSION}.alfredworkflow"

mkdir -p releases
rm -f "$OUT"   # zip -r appends; start fresh so nothing stale/secret carries over

( cd "$SRC" && zip -rq "../$OUT" . \
    -x '*__pycache__*' '*.pyc' 'config.json' 'token_cache.json' 'prefs.plist' '*.bak' '.DS_Store' '*/.DS_Store' )

if unzip -l "$OUT" | grep -qE '(config\.json|token_cache\.json|prefs\.plist)$'; then
    echo "ERROR: $OUT still contains a secret file — aborting." >&2
    rm -f "$OUT"
    exit 1
fi

echo "Built clean bundle: $OUT ($(du -h "$OUT" | cut -f1))"
