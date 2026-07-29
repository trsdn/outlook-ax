#!/usr/bin/env bash
# sync-agent-skills.sh — mirror .agents/skills/ into .claude/skills/
# Works correctly regardless of the current working directory.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

SRC="${REPO_ROOT}/.agents/skills"
DST="${REPO_ROOT}/.claude/skills"

[ -d "${SRC}" ] || exit 0
mkdir -p "${DST}"
rsync -a --delete "${SRC}/" "${DST}/"