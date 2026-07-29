#!/usr/bin/env bash
# check-architecture.sh — validate Sources/ directory architecture conventions.
# Rejects patterns that violate the single-implementation rule:
#   - root outlook-ax.swift still present (must be removed after parity)
#   - duplicate L10n/parser declarations (outside the shared catalog)
#   - direct wins[0] usage in Sources/
#   - direct AX calls outside Sources/OutlookAX/
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
SOURCES="${REPO_ROOT}/Sources"

PASS=0
FAIL=0

ok()   { echo "  ok: $1"; PASS=$((PASS + 1)); }
fail() { echo "FAIL: $1" >&2; FAIL=$((FAIL + 1)); }

echo "=== Architecture check ==="

# 1. Root outlook-ax.swift must be absent (removed after parity was proven)
if [ -f "${REPO_ROOT}/outlook-ax.swift" ]; then
    fail "Root outlook-ax.swift still present — remove it after Sources/OutlookAX reaches parity"
else
    ok "Root outlook-ax.swift is absent (parity migration complete)"
fi

# 2. No wins[0] in Sources/ (use firstWindow, mainWindow, or semantic selectors)
if grep -r 'wins\[0\]' "${SOURCES}" --include='*.swift' --quiet 2>/dev/null; then
    fail "wins[0] found in Sources/ — use semantic window selectors instead:"
    grep -r 'wins\[0\]' "${SOURCES}" --include='*.swift' -n 2>/dev/null || true
else
    ok "No wins[0] in Sources/"
fi

# 3. No duplicate L10n struct/enum definitions in Sources/
# (The canonical L10n lives in Sources/OutlookAX/Localization/L10n.swift)
L10N_COUNT=$(grep -r '^private.*enum L10n\|^internal.*enum L10n\|^public.*enum L10n\|^enum L10n\|^struct L10n' \
    "${SOURCES}" --include='*.swift' 2>/dev/null | wc -l | tr -d ' ') || L10N_COUNT=0
if [ "${L10N_COUNT}" -gt 1 ]; then
    fail "Multiple L10n declarations in Sources/ (count: ${L10N_COUNT}). One shared catalog only."
    grep -r 'enum L10n\|struct L10n' "${SOURCES}" --include='*.swift' -n 2>/dev/null || true
else
    ok "Single L10n declaration in Sources/ (count: ${L10N_COUNT})"
fi

# 4. No direct AX calls outside Sources/OutlookAX/
for subdir in OutlookAXCLIKit OutlookAXCLI; do
    dir="${SOURCES}/${subdir}"
    if [ -d "${dir}" ]; then
        AX_COUNT=$(grep -r 'AXUIElement\|kAXRoleAttribute\|AXUIElementPerformAction' \
            "${dir}" --include='*.swift' 2>/dev/null | wc -l | tr -d ' ') || AX_COUNT=0
        if [ "${AX_COUNT}" -gt 0 ]; then
            fail "Direct AX calls in ${subdir}/ (${AX_COUNT} lines) — AX access must go through Sources/OutlookAX/"
            grep -r 'AXUIElement\|kAXRoleAttribute' "${dir}" --include='*.swift' -n 2>/dev/null || true
        else
            ok "No direct AX calls in ${subdir}/"
        fi
    else
        ok "${subdir}/ not present (skipped)"
    fi
done

# 5. Sources/OutlookAXCLI/main.swift must exist (executable entry point)
if [ -f "${SOURCES}/OutlookAXCLI/main.swift" ]; then
    ok "Sources/OutlookAXCLI/main.swift exists"
else
    fail "Sources/OutlookAXCLI/main.swift missing — executable entry point required"
fi

# 6. Package.swift must expose an executable product named 'outlook-ax'
if grep -q 'outlook-ax' "${REPO_ROOT}/Package.swift" 2>/dev/null; then
    ok "Package.swift references outlook-ax"
else
    fail "Package.swift does not expose an outlook-ax executable product"
fi

# 7. L10n catalog must be in the expected location
if [ -f "${SOURCES}/OutlookAX/Localization/L10n.swift" ]; then
    ok "Sources/OutlookAX/Localization/L10n.swift exists"
else
    fail "Sources/OutlookAX/Localization/L10n.swift missing — shared L10n catalog required"
fi

# 8. Error model must be in the expected location
if [ -f "${SOURCES}/OutlookAX/Models/Errors.swift" ]; then
    ok "Sources/OutlookAX/Models/Errors.swift exists"
else
    fail "Sources/OutlookAX/Models/Errors.swift missing — typed error model required"
fi

# 9. ConnectionManager must exist (passive/interactive connection policy)
if [ -f "${SOURCES}/OutlookAX/Connection/ConnectionManager.swift" ]; then
    ok "Sources/OutlookAX/Connection/ConnectionManager.swift exists"
else
    fail "Sources/OutlookAX/Connection/ConnectionManager.swift missing — connection policy required"
fi

# 10. Subprocess tests must exist
if [ -f "${REPO_ROOT}/Tests/OutlookAXCLISubprocessTests/OutlookAXCLISubprocessTests.swift" ]; then
    ok "Tests/OutlookAXCLISubprocessTests/ exists"
else
    fail "Tests/OutlookAXCLISubprocessTests/ missing — subprocess tests required"
fi

echo ""
echo "Passed: ${PASS}  Failed: ${FAIL}"
[ "${FAIL}" -eq 0 ] || exit 1
echo "All architecture checks passed."
