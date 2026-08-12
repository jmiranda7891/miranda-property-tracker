#!/usr/bin/env bash
# Shared setup for ./clasp-push.sh — SOURCE this, don't run it.
#
# Does the four things the deploy script needs before it can talk to Apps Script:
#   1. refuses to run outside this repo (a stray push must never reach the wrong account)
#   2. restores the project-local credential from CLASP_AUTH_OE_B64 if it isn't on disk
#   3. finds a clasp to use without requiring a global install
#   4. defines clasp_oe(), which passes the right auth flags for wherever we're running
#
# Afterwards, call clasp as:  clasp_oe push -f
#
# Ported from the CL Social Media App repo's clasp-env.sh (same mechanism, OE account
# instead of CL — see that repo's CLAUDE.md "clasp multi-account setup" section for the
# full backstory on why cloud sessions use a project-local .clasprc.json).

if [[ "${BASH_SOURCE[0]}" == "$0" ]]; then
  echo "!!  clasp-env.sh is meant to be sourced, not executed. Use ./clasp-push.sh."
  exit 1
fi

if [[ "$(basename "$PWD")" != "miranda-property-tracker" ]]; then
  echo "!!  Not in miranda-property-tracker — aborting to avoid pushing to the wrong account."
  exit 1
fi

if [[ ! -f .clasp.json ]]; then
  echo "!!  No .clasp.json here — this doesn't look like the linked project."
  exit 1
fi

# Cloud sessions authenticate from a PROJECT-LOCAL .clasprc.json. If it's missing but the
# base64 credential is in the environment, restore it here. Never printed, never committed
# (.gitignore), never pushed to Apps Script (.claspignore is a whitelist).
if [[ ! -f .clasprc.json && -n "${CLASP_AUTH_OE_B64:-}" ]]; then
  printf '%s' "$CLASP_AUTH_OE_B64" | base64 -d > .clasprc.json
  chmod 600 .clasprc.json
  echo "..  restored .clasprc.json from CLASP_AUTH_OE_B64"
fi

# Find clasp without demanding a global install.
if [[ -x node_modules/.bin/clasp ]]; then
  CLASP_BIN=(node_modules/.bin/clasp)
  CLASP_HOW="node_modules/.bin/clasp"
elif command -v clasp >/dev/null 2>&1; then
  CLASP_BIN=(clasp)
  CLASP_HOW="clasp on PATH ($(command -v clasp))"
elif command -v npx >/dev/null 2>&1; then
  CLASP_BIN=(npx --yes @google/clasp@3)
  CLASP_HOW="npx @google/clasp@3 (fetched on demand, first call is slow)"
else
  echo "!!  No clasp and no npx. Install one: npm install -g @google/clasp"
  exit 1
fi

# Two auth styles, and they are NOT interchangeable:
#   cloud → project-local file, and -A must be passed on EVERY command; bare clasp looks
#           for a global ~/.clasprc.json that doesn't exist there.
#   Mac   → clasp 2.x named credential sets (clasp login --user OE). Pinned-v3 npx does not
#           understand --user, so refuse that combination instead of authenticating as
#           nobody (or the wrong account).
if [[ -f .clasprc.json ]]; then
  CLASP_AUTH=(-A .clasprc.json)
  CLASP_WHO="project-local credential (cloud session)"
elif [[ "${CLASP_BIN[0]}" == "npx" ]]; then
  echo "!!  Falling back to npx clasp v3, which has no --user flag, and there is no"
  echo "    project-local .clasprc.json to use instead."
  echo "    Either set CLASP_AUTH_OE_B64 in the environment"
  echo "    or install clasp and log in: npm install -g @google/clasp && clasp login --user OE"
  exit 1
else
  CLASP_AUTH=(--user OE)
  CLASP_WHO="named local login, --user OE (Mac)"
fi

clasp_oe() { "${CLASP_BIN[@]}" "${CLASP_AUTH[@]}" "$@"; }

# Prove the push actually landed, by pulling it back and diffing (see the CL Social Media
# App repo's CLAUDE.md for the exact 2026-08-06 incident this defends against: a push can
# report success and print "Pushed N files" while the deployed script stays stale).
DEPLOYED_FILES=(appsscript.json Code.js Index.html JavaScript.html Styles.html)
clasp_verify_push() {
  local tmp diff_count=0
  tmp="$(mktemp -d)"
  cp .clasp.json "$tmp"/ 2>/dev/null || true
  [[ -f .clasprc.json ]] && cp .clasprc.json "$tmp"/
  if ! ( cd "$tmp" && "${CLASP_BIN[@]}" "${CLASP_AUTH[@]}" pull >/dev/null 2>&1 ); then
    echo "!!  Could not pull the deployed copy back to verify the push."
    echo "    Check by hand before trusting this deploy."
    rm -rf "$tmp"; return 1
  fi
  local f
  for f in "${DEPLOYED_FILES[@]}"; do
    if ! diff -q "$tmp/$f" "$f" >/dev/null 2>&1; then
      echo "!!  $f on Apps Script does NOT match the local file."
      diff_count=$((diff_count + 1))
    fi
  done
  rm -rf "$tmp"
  if (( diff_count )); then
    echo "!!  THE PUSH DID NOT LAND, despite reporting success. $diff_count of ${#DEPLOYED_FILES[@]} files differ."
    echo "    Run the push again and watch this check. Do not test or release until it passes."
    return 1
  fi
  echo "..  verified: all ${#DEPLOYED_FILES[@]} deployed files match local, byte for byte"
  return 0
}

# The web app's single deployment id. Reusing this SAME id on every future
# `clasp deploy -i "$WEBAPP_DEPLOYMENT_ID" -d "description"` is what keeps the URL stable.
# URL: https://script.google.com/a/macros/orderexpress.com/s/AKfycbxx-MBnRft8Zp7xYQkCmsR7Yg0VxeojqLrcj5XjyNjiM95ZUT3Ayj9ALhhLfGqUdGA/exec
WEBAPP_DEPLOYMENT_ID="${WEBAPP_DEPLOYMENT_ID:-AKfycbxx-MBnRft8Zp7xYQkCmsR7Yg0VxeojqLrcj5XjyNjiM95ZUT3Ayj9ALhhLfGqUdGA}"

echo "..  clasp: $CLASP_HOW"
echo "..  auth:  $CLASP_WHO"
