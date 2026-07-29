#!/usr/bin/env bash
# check-repository-layout.sh — validate repository layout conventions.
# Fails fast on the first violation.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

PASS=0
FAIL=0

ok()   { echo "  ok: $1"; PASS=$((PASS + 1)); }
fail() { echo "FAIL: $1" >&2; FAIL=$((FAIL + 1)); }

echo "=== Repository layout check ==="

# 1. .claude/settings.json must exist (correct filename)
if [ -f "${REPO_ROOT}/.claude/settings.json" ]; then
    ok ".claude/settings.json exists"
else
    fail ".claude/settings.json missing — Claude SessionStart hook will not fire"
fi

# 2. .claude/.settings.json must NOT exist (wrong filename)
if [ -f "${REPO_ROOT}/.claude/.settings.json" ]; then
    fail ".claude/.settings.json still exists — rename to settings.json"
else
    ok ".claude/.settings.json absent (correct)"
fi

# 3. All shell scripts under scripts/ must pass bash -n syntax check
while IFS= read -r -d '' script; do
    if bash -n "${script}" 2>/dev/null; then
        ok "bash -n $(basename "${script}")"
    else
        fail "syntax error in ${script}"
    fi
done < <(find "${REPO_ROOT}/scripts" -name '*.sh' -print0)

# 4. Skill-sync script must run successfully from outside the repository root
TMPDIR_CHECK="$(mktemp -d)"
if ( cd "${TMPDIR_CHECK}" && bash "${REPO_ROOT}/scripts/sync-agent-skills.sh" ) 2>/dev/null; then
    ok "sync-agent-skills.sh works from outside repo root"
else
    fail "sync-agent-skills.sh failed when run from outside the repository root"
fi
rm -rf "${TMPDIR_CHECK}"

echo ""
echo "Passed: ${PASS}  Failed: ${FAIL}"
[ "${FAIL}" -eq 0 ] || exit 1
echo "All layout checks passed."
