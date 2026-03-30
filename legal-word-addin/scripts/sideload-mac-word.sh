#!/usr/bin/env bash
# Copies manifest.xml into Word's sideload folder (macOS sandbox).
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DEST="${HOME}/Library/Containers/com.microsoft.Word/Data/Documents/wef"
MANIFEST="${ROOT}/manifest.xml"

if [[ ! -f "${MANIFEST}" ]]; then
  echo "Missing ${MANIFEST}" >&2
  exit 1
fi

mkdir -p "${DEST}"
cp "${MANIFEST}" "${DEST}/legal-ai-assistant.xml"
echo "Copied manifest to:"
echo "  ${DEST}/legal-ai-assistant.xml"

if [[ "${1:-}" == "--clear-cache" ]]; then
  CACHE="${HOME}/Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
  if [[ -d "${CACHE}" ]]; then
    rm -rf "${CACHE:?}/"*
    echo "Cleared Office add-in cache:"
    echo "  ${CACHE}"
  else
    echo "No cache folder yet (OK): ${CACHE}"
  fi
fi

echo ""
echo "1) Quit Word completely (Word → Quit Word or Cmd+Q)."
echo "2) Task pane URL must be reachable:"
echo "   • Hosted: deploy dist/ to HTTPS, ADDIN_PUBLIC_URL=… npm run build, sideload dist/manifest.xml"
echo "   • Local dev (proxy):  cd \"${ROOT}\" && npm run dev:all   (webpack :3000 + API proxy :3548)"
echo "   • Local dev (direct only):  npm run dev  — set Settings → Direct and expect possible CORS issues"
echo "3) Open Word, open a real document (not only the start screen)."
echo "4) On the Insert and Home tabs, look for ribbon group \"Legal AI\"."
echo ""
echo "If it still does not appear, Word may be using a different container. Candidates:"
ls -1 "${HOME}/Library/Containers" 2>/dev/null | grep -i word || echo "  (none found — open Word once, then run this script again)"
