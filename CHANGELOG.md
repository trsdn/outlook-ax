# Changelog

All notable changes to outlook-ax will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [1.0.0] — 2026-04-18

### Added

- **42 CLI commands** covering mail, calendar, navigation, and system operations
- **Localization system** — `struct L10n` with 124 label arrays (de, en, fr, es, it) and matching helpers (`equalsAny`, `startsWithAny`, `matchesAny`, `endsWithAny`)
- **Mail commands** — `current`, `inbox`, `search`, `reply`, `reply-all`, `forward`, `delete`, `archive`, `compose`, `folders`, `folder`, `flag`, `read-unread`, `move`, `report`, `react`, `summarize`, `filter`
- **Calendar commands** — `today`, `create`, `view`, `navigate`, `calendars`, `toggle`, `timescale`, `filter`, `color`, `accept`, `tentative`, `decline`, `join`, `duplicate`, `categorize`, `private`, `show-as`
- **System commands** — `status`, `notifications`, `sync`, `auto-reply`, `my-day`, `account`
- **Auto-launch** — `ensureOutlookReady()` launches and unminimizes Outlook via AppleScript
- **JSON output** — all commands support `--json` with English-normalized values
- **Keyboard simulation** — CGEvent-based typing for date/time fields that reject AXValue writes
- **L10n-aware menu triggering** — `triggerMenuL10n` accepts multi-language path arrays
- **Agent documentation** — AGENTS.md, docs/architecture.md, docs/ax-paths.md
- **AX discovery skill** — `.agents/skills/ax-discovery/SKILL.md` for agent-driven UI exploration
- **Skills sync pattern** — `.agents/skills/` → `.claude/skills/` via SessionStart hook

[Unreleased]: https://github.com/trsdn/outlook-ax/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/trsdn/outlook-ax/releases/tag/v1.0.0
