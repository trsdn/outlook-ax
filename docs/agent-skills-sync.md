# Agent Skills Sync Pattern

> How `.agents/skills/` (canonical, committed) gets mirrored to `.claude/skills/` (local, gitignored).

## Problem

Claude Code reads skills from `.claude/skills/`. But `.claude/` is tool-specific and
shouldn't be the source of truth. Other agents (Copilot, Cursor, custom) need the
same skills from a vendor-neutral location.

## Solution: `.agents/skills/` as Single Source of Truth

```
.agents/skills/          ← committed, vendor-neutral, canonical
.claude/skills/          ← gitignored, local mirror for Claude Code
```

## Three Tiers of Enforcement

### Tier 1 — SessionStart Hook (recommended)

Fires automatically when Claude Code opens a session. No manual steps.

**`.claude/settings.json`** in repo root:

```json
{
  "hooks": {
    "SessionStart": [
      {
        "hooks": [
          {
            "type": "command",
            "command": "bash scripts/sync-agent-skills.sh"
          }
        ]
      }
    ]
  }
}
```

**`scripts/sync-agent-skills.sh`**:

```bash
#!/usr/bin/env bash
set -euo pipefail
[ -d .agents/skills ] || exit 0
mkdir -p .claude/skills
rsync -a --delete .agents/skills/ .claude/skills/
```

**`.gitignore`** entry:

```
.claude/skills/
```

### Tier 2 — AGENTS.md / CLAUDE.md Instruction (fallback)

For repos where the hook isn't configured or on first run on a new machine:

```markdown
## Skills Path

Canonical skills live in `.agents/skills/`. Before using or listing skills,
run `bash scripts/sync-agent-skills.sh` to mirror them to `.claude/skills/`.
`.claude/skills/` is gitignored and only a local mirror.
```

Limitation: This is an instruction, not enforcement. Agents can forget it
in long sessions. That's why Hook > instruction.

### Tier 3 — Git Hooks (second defense line)

For non-Claude workflows (Copilot-only, manual branch switching):

- `post-checkout` and `post-merge` hooks run the sync script
- Set up via Husky or `core.hooksPath`
- Only worth it in repos with heavy branch-hopping

## Recommendation

**Tier 1 + Tier 2 combined**: Hook does the work automatically, AGENTS.md documents
the pattern and catches edge cases (hook not configured, first run, session without
repo root).

## Portability

For the same pattern across multiple repos, package it as a standalone skill
(`~/.agents/skills/agent-skills-sync/`) with the sync script and AGENTS.md snippet
as a portable artifact. Avoids copy-pasting the setup into every new repo.
