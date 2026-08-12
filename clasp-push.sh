#!/usr/bin/env bash
# Push the current working tree to Apps Script (@HEAD) and verify it landed byte-for-byte.
# Works the same in any Claude Code cloud session for this repo — clasp-env.sh restores the
# OE credential from CLASP_AUTH_OE_B64 if it isn't already on disk.
set -euo pipefail

# shellcheck source=clasp-env.sh
source ./clasp-env.sh

clasp_oe push -f
clasp_verify_push || exit 1

echo
echo "Pushed to @HEAD. Hard-reload before testing — Apps Script caches hard."
echo "First-ever deploy: clasp_oe deploy -d \"description\" (then save the printed"
echo "deployment id as WEBAPP_DEPLOYMENT_ID so future deploys reuse the same URL)."
