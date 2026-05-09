#!/usr/bin/env bash
# pii-gauntlet.sh — verify no personal data leaked into TRACKED files in this
# repo before any public push.
#
# Portable: resolves its own repo via git rev-parse, runs from any cwd.
# Git-aware: scans only tracked files (`git ls-files`), so gitignored .env
# files, local secrets, and on-disk personal data don't trigger false fails.
#
# Run: ./scripts/pii-gauntlet.sh
# Exit 0 = clean. Non-zero = PII detected, see output for details.
#
# Self-exclusion: this script contains the patterns it searches for, so
# `pii-gauntlet.sh` is filtered out of the file list.

set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(git -C "$SCRIPT_DIR" rev-parse --show-toplevel 2>/dev/null)"
if [ -z "$REPO_ROOT" ]; then
  echo "ERROR: script is not inside a git repository ($SCRIPT_DIR)" >&2
  exit 2
fi
cd "$REPO_ROOT"

echo "=== PII Gauntlet ==="
echo "Repo: $REPO_ROOT"
echo

FAIL=0

# Build the list of TRACKED files to scan, with a few exclusions:
#   - the gauntlet itself (self-match)
#   - LICENSE files (copyright lines are intentional)
#   - lock files (huge, low-signal)
TRACKED_FILES=$(git ls-files \
  | grep -vE '(^|/)(pii-gauntlet\.sh|PII-GAUNTLET\.md|LICENSE|LICENSE\.md|package-lock\.json|uv\.lock|pnpm-lock\.yaml|yarn\.lock)$' \
  || true)

if [ -z "$TRACKED_FILES" ]; then
  echo "ERROR: no tracked files found" >&2
  exit 2
fi

check() {
  local label="$1"
  local pattern="$2"
  local exclude_substring="${3:-}"  # optional grep -v post-filter
  local hits
  # NUL-delimit to handle paths with spaces; -I skips binary files
  hits=$(echo "$TRACKED_FILES" | tr '\n' '\0' | xargs -0 grep -nIE "$pattern" 2>/dev/null || true)
  if [ -n "$exclude_substring" ] && [ -n "$hits" ]; then
    hits=$(echo "$hits" | grep -vE "$exclude_substring" || true)
  fi
  if [ -n "$hits" ]; then
    echo "FAIL [$label]:"
    echo "$hits" | head -20
    echo
    FAIL=1
  else
    echo "OK   [$label]"
  fi
}

# Personal name (full forms — single-word "plessas" is the brand name, OK)
check "Full personal name (EN)" "Dimitris[[:space:]]+Plessas|Dimitrios[[:space:]]+Plessas"
check "Full personal name (GR)" "Δημήτριος[[:space:]]+Πλέσσας|ΠΛΕΣΣΑΣ[[:space:]]+ΔΗΜΗΤΡΙΟΣ"

# Personal emails
check "Personal email" "dimitrios\.plessas@|plessasdimitrios@|plessas@nbg\.gr|plessas@gmail|plessas@yahoo"

# Personal phone / address
check "Personal phone" "694[[:space:]]?9200878|6949200878"
check "Personal address" "174[[:space:]]+Syggrou|Συγγρού[[:space:]]+174"

# Peer names (NBG colleagues / direct reports / managers)
check "Peer/colleague names" "Volioti|Bitrou|Sioutis|Theofilidi|Θεοφιλίδη|Χριστίνα|Lygeros|Oikonomou|Maraveas|Xona|Petropoulou|Laspas|Koutra|Giemelou"

# Family names
check "Family names" "Kitrilaki|Κιτριλάκη"

# NBG-internal project names
check "Internal projects" "Διπλή κάρτα|\bdual[- ]card\b|IRIS[[:space:]]+pilot|ECB[[:space:]]+Digital[[:space:]]+Euro[[:space:]]+CfEI"

# External partners discussed in NBG-internal context
check "Partner names" "\bWorldline\b|\bHelvia\b|\bWealthyhood\b|\bFeedzai\b|\bMellon\b|\b11FS\b|\bNCR\b"

# Tax authority refs
check "Tax authority" "ΑΑΔΕ|ΑΦΜ|ΑΔΤ|ΑΜΚΑ"

# Specific user-machine paths (Plessas-specific)
check "User-specific paths" "/Users/plessas|/SourceCode/claude-config|claude-config/shared-memory"

# Generic home-directory paths — catches /Users/<anyone>/ patterns. Filters
# out conventional placeholder names (me, you, user, USER, USERNAME) and
# macOS system paths (Shared, Public, Library, Guest).
check "Generic /Users/<user>/ paths" "/Users/[A-Za-z0-9_.-]+/" \
  "/Users/(me|you|user|USER|USERNAME|Shared|Public|Library|Guest|\.\.\.)/"

# Personal GCP project IDs (gen-lang-client-* is the Google AI Studio default
# project naming pattern — typically personal accounts)
check "Personal GCP project IDs" "gen-lang-client-[0-9]+"

echo
if [ $FAIL -eq 0 ]; then
  echo "=== GAUNTLET PASS ==="
  exit 0
else
  echo "=== GAUNTLET FAIL ==="
  echo "Fix the PII leaks above before any public push."
  exit 1
fi
