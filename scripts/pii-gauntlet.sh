#!/usr/bin/env bash
# pii-gauntlet.sh — verify no PII leaked into plessas-marketplace.
#
# TWO MODES:
#   --mode=ci      Scan only git-tracked files. Used by GitHub Actions to gate
#                  pushes. Any hit = FAIL = exit 1.
#   --mode=doctor  (default) Scan the entire working tree. Distinguishes tracked
#                  hits (FAIL — these would ship publicly) from gitignored hits
#                  (INFO — local-only, never pushed). Exit 1 only on tracked hits.
#
# Why two modes:
#   The CI mode is the actual safety gate.
#   The doctor mode helps the maintainer notice PII drift in their LOCAL files
#   before they accidentally `git add` something. It must NOT scare a teammate
#   running the script casually — "FAIL" on a gitignored file would teach them
#   to ignore the script entirely, defeating the point.
#
# Self-exclusion: this script contains the very patterns it searches for, so
# `--exclude=pii-gauntlet.sh` is essential to avoid self-match false positives.

set -u

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

# Default to CI mode on a runner. Three repositories invoke this script with no
# --mode flag at all, so they got doctor mode, which fails hard when the private
# denylist is absent -- and on a runner it is always absent. Keying off the
# environment rather than the caller means a workflow cannot get this wrong by
# omission. GitHub, GitLab and most others set CI=true.
if [ -n "${CI:-}" ] || [ -n "${GITHUB_ACTIONS:-}" ]; then
  MODE="ci"
else
  MODE="doctor"
fi
_MODE_FROM_ENV="$MODE"
for arg in "$@"; do
  case "$arg" in
    --mode=ci)     MODE="ci" ;;
    --mode=doctor) MODE="doctor" ;;
    -h|--help)
      sed -n '2,18p' "$0"
      exit 0
      ;;
    *) echo "Unknown arg: $arg" >&2; exit 2 ;;
  esac
done

echo "=== PII Gauntlet (mode: $MODE) ==="
echo "Repo: $REPO_ROOT"
echo

FAIL=0
INFO=0

# Build the file list once. CI mode = tracked only. Doctor mode = working tree.
if [ "$MODE" = "ci" ]; then
  # Exclude self + auto-generated lockfiles at any depth (lockfiles contain SHAs / hashes that
  # collide with the 9-digit-ID regex but carry no PII risk).
  # User-cleared public showcase assets (maintainer confirmed 2026-06-08): the
  # decks screenshot library is public-safe, and its INDEX.md captions legitimately
  # name NBG products (dual card, Skroutz, …). Exclude that subtree from scanning.
  # This script's path relative to the repo root. Derived, not hardcoded: the
  # same file lives in installers/ in some repos and scripts/ in others, and
  # six hand-maintained copies is what let them drift apart in the first place.
  SELF_REL=$(git ls-files --full-name -- "$0" 2>/dev/null | head -1)
  [ -z "$SELF_REL" ] && SELF_REL="installers/pii-gauntlet.sh"
  TRACKED=$(git ls-files \
    | grep -v "^$SELF_REL$" \
    | grep -vE '(^|/)LICENSE(\.md|\.txt)?$' \
    | grep -vE '(^|/)(package-lock\.json|yarn\.lock|pnpm-lock\.yaml|poetry\.lock|Pipfile\.lock)$' \
    | grep -vE '^plugins/decks/assets/screenshots/' \
    || true)
  TRACKED_TMP=$(mktemp)
  printf '%s\n' "$TRACKED" > "$TRACKED_TMP"
fi

# Helper: get the tracked-vs-untracked status of a file.
file_is_tracked() {
  git ls-files --error-unmatch "$1" >/dev/null 2>&1
}

scan_doctor() {
  local pattern="$1"
  grep -riE \
    --exclude-dir=.git \
    --exclude-dir=node_modules \
    --exclude-dir=dist \
    --exclude-dir=__pycache__ \
    --exclude-dir=.venv \
    --exclude-dir=venv \
    --exclude-dir=.remember \
    --exclude-dir=data \
    --exclude-dir=attachments \
    --exclude-dir=backups \
    --exclude-dir=archive \
    --exclude-dir=.next \
    --exclude-dir=dist \
    --exclude-dir=coverage \
    --exclude-dir=installers/deps \
    --exclude=pii-gauntlet.sh \
    --exclude=LICENSE \
    --exclude=PII-GAUNTLET.md \
    --exclude=package-lock.json \
    --binary-files=without-match \
    "$pattern" . 2>/dev/null \
    | grep -vE 'plugins/decks/assets/screenshots/' \
    || true
}

scan_ci() {
  local pattern="$1"
  # Search only git-tracked files. NUL-delimit the list and use `xargs -0`
  # (portable on BSD/macOS and GNU). The old `xargs -a FILE -d '\n'` form is
  # GNU-only: on macOS it errors "invalid option -- a", gets swallowed by
  # 2>/dev/null, and the gate silently PASSES while scanning nothing.
  if [ -s "$TRACKED_TMP" ]; then
    tr '\n' '\0' < "$TRACKED_TMP" | xargs -0 grep -nE --binary-files=without-match "$pattern" 2>/dev/null || true
  fi
}

# Drop hits whose PATH is a historical record rather than live configuration.
# Path-scoped only. Never extend this to filter on matched content: that would
# hide live hits and turn a working guardrail into a false green.
apply_exclusion() {
  local hits="$1"
  local exclude="$2"
  if [ -z "$exclude" ] || [ -z "$hits" ]; then
    printf '%s' "$hits"
    return
  fi
  # Copyright attribution names the author on purpose and is required by the
  # licence. Flagging it is noise, and noise is how a real hit gets ignored.
  printf '%s\n' "$hits" | grep -vE "$exclude" | grep -viE '\(c\)[[:space:]]*[0-9]{4}|copyright' || true
}

check() {
  local label="$1"
  local pattern="$2"
  local exclude="${3:-}"
  local hits

  if [ "$MODE" = "ci" ]; then
    hits=$(scan_ci "$pattern")
    hits=$(apply_exclusion "$hits" "$exclude")
    if [ -n "$hits" ]; then
      echo "FAIL [$label]:"
      echo "$hits" | head -20
      echo
      FAIL=1
    else
      echo "OK   [$label]"
    fi
    return
  fi

  # Doctor mode — separate tracked from gitignored.
  hits=$(scan_doctor "$pattern")
  hits=$(apply_exclusion "$hits" "$exclude")
  if [ -z "$hits" ]; then
    echo "OK   [$label]"
    return
  fi

  local tracked_hits=""
  local untracked_hits=""
  while IFS= read -r line; do
    [ -z "$line" ] && continue
    # line format: ./path/to/file:matched-text
    local path="${line%%:*}"
    path="${path#./}"
    if file_is_tracked "$path"; then
      tracked_hits+="$line"$'\n'
    else
      untracked_hits+="$line"$'\n'
    fi
  done <<< "$hits"

  if [ -n "$tracked_hits" ]; then
    echo "FAIL [$label]:                 (tracked — would ship publicly)"
    printf '%s' "$tracked_hits" | head -20
    echo
    FAIL=1
  fi
  if [ -n "$untracked_hits" ]; then
    echo "INFO [$label]:                 (gitignored / untracked — local-only)"
    printf '%s' "$untracked_hits" | head -10
    echo
    INFO=1
  fi
  if [ -z "$tracked_hits" ] && [ -z "$untracked_hits" ]; then
    echo "OK   [$label]"
  fi
}

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Generic organisation tells (safe to keep in a public repo)
# ---------------------------------------------------------------------------
#
# These are patterns, not names, so publishing them discloses nothing. They are
# also the checks that were missing entirely: the four exposures found in the
# 2026-08-24 audit all passed the gauntlet green, because nothing here looked
# for the employer's name, its mail domain, or a tenant hostname.

# Placeholders are the whole point of a good example, so they must not trip the
# check that exists to catch the real thing. Without these excludes the gauntlet
# flags `contoso.sharepoint.com` (the correct placeholder) and its own CHANGELOG
# entries describing the removal of the real host. A check that fires on
# deliberate, already-disclosed content teaches you to ignore it, which is how
# the four real exposures sat unnoticed next to a green gauntlet.
# Doc placeholders. Extend this when a new invented example host trips the check:
# being asked once "is this a real tenant?" is the check doing its job, and is a
# far better failure mode than the silence it replaces.
PLACEHOLDER='contoso|example|sample|template|your[-_.]?tenant|your-tenant|<[^>]+>|firstname\.lastname|your\.email|recipient\.name|user@|name@|(test|overridden|envvar|dummy|placeholder|foo|bar|[a-z])(-my)?\.sharepoint'

# Some repos name the employer on purpose: a marketplace written for colleagues
# says so in its README by design. Those opt out with a repo-root marker rather
# than carrying a permanent red light.
if [ ! -f ".pii-gauntlet-allow-employer-name" ]; then
  check "Employer name" '(^|[^A-Za-z0-9])(NBG|ΕΤΕ)([^A-Za-z0-9]|$)|Εθνική Τράπεζα|National Bank of Greece' "$PLACEHOLDER"
else
  echo "OK   [Employer name] (opted out via .pii-gauntlet-allow-employer-name)"
fi

check "Employer mail domain" '[A-Za-z0-9._%+-]+@nbg\.gr' "$PLACEHOLDER"
check "SharePoint tenant"    '[a-z0-9-]+\.sharepoint\.com' "$PLACEHOLDER"
# A bare UUID is not a finding: fixtures, generated filenames and message ids
# are full of them, and flagging all of them is how a check earns the right to
# be ignored. What matters is a GUID sitting where a TENANT id sits, so key on
# the surrounding context rather than on the shape alone.
check "Azure AD tenant id" \
  '(tenant[_-]?id|tenantId|\"tid\"|authority|login\.microsoftonline\.com/|realm)[^0-9a-f]{0,24}[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}' \
  '00000000-0000-0000-0000-000000000000|11111111-2222-3333-4444-555555555555|00000000-|11111111-|22222222-|33333333-|44444444-|55555555-|66666666-|77777777-|88888888-|99999999-|aaaaaaaa-|bbbbbbbb-|cccccccc-|dddddddd-|eeeeeeee-|ffffffff-|/common|/organizations|/consumers'

# ---------------------------------------------------------------------------
# Name-based checks, loaded from a private denylist
# ---------------------------------------------------------------------------
#
# The literal terms used to live in this file. That made the guard the
# disclosure: this script is anonymously readable, and it enumerated colleague
# names, a family name, a personal mobile and address, personal emails,
# unreleased internal project names, the partner list and every private repo
# path, complete with comment headers explaining which was which.
#
# They now live outside every public repo. Nothing here reveals what is checked.

PII_DENYLIST="${PII_DENYLIST:-$HOME/.claude/private/pii-denylist.conf}"

if [ -r "$PII_DENYLIST" ]; then
  # Tab-delimited: the patterns are full of regex alternation pipes, so "|"
  # cannot be the field separator.
  loaded=0
  while IFS=$'\t' read -r label pattern exclude; do
    case "$label" in ''|\#*) continue ;; esac
    [ -z "$pattern" ] && continue
    check "$label" "$pattern" "$exclude"
    loaded=$((loaded + 1))
  done < "$PII_DENYLIST"
  echo "     ($loaded name-based checks loaded from the private denylist)"
else
  # Absence is handled differently by mode, on purpose.
  #
  # In CI on a public repo the denylist legitimately does not exist and never
  # will, so failing here would just paint six repos red forever. Say plainly
  # that the name checks did not run, and let the generic ones stand.
  #
  # Locally the file should always be there. Its absence is a real
  # misconfiguration, and a guard that cannot evaluate its condition must
  # refuse rather than pass. That distinction is the whole point: the previous
  # version of this script reported OK on every check while, on macOS, scanning
  # exactly zero files.
  if [ "$MODE" = "ci" ]; then
    echo "SKIP [name-based checks]: no denylist at $PII_DENYLIST"
    echo "     Generic org-tell checks above still ran. Name, family, partner and"
    echo "     private-path checks did NOT. This is expected in public CI."
  else
    echo "FAIL [name-based checks]: no denylist at $PII_DENYLIST"
    echo "     Locally this file must exist. Without it the name, family, partner"
    echo "     and private-path checks are silently absent, which is exactly the"
    echo "     false green this rewrite removes."
    echo "     Fix: ensure ~/.claude/private -> claude-config/private is linked,"
    echo "     or set PII_DENYLIST to the file."
    FAIL=1
  fi
fi

if [ $FAIL -eq 0 ]; then
  if [ "$MODE" = "doctor" ] && [ $INFO -ne 0 ]; then
    echo "=== GAUNTLET PASS (with INFO on gitignored files — local-only, not in git) ==="
  else
    echo "=== GAUNTLET PASS ==="
  fi
  exit 0
else
  echo "=== GAUNTLET FAIL ==="
  if [ "$MODE" = "doctor" ]; then
    echo "Tracked PII detected. These files would ship publicly. Fix before committing."
  else
    echo "Fix the PII leaks above before any public push."
  fi
  exit 1
fi
