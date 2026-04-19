#!/usr/bin/env bash
set -euo pipefail
[ -d .agents/skills ] || exit 0
mkdir -p .claude/skills
rsync -a --delete .agents/skills/ .claude/skills/