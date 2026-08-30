# Changelog

All notable changes to outlook-ax will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased] — v2.0 (planned)

### Changed

- `--json` produces JSON v2 by default (v1 via `--json-version 1`)

## [1.1.1] — 2026-08-30

### Added

- Automated macOS release assets — version tags build, test, package, and
  smoke-test the `outlook-ax` executable before attaching it to the GitHub Release

## [1.1.0] — 2026-07-29

### Added

- **SwiftPM architecture** — `Sources/OutlookAXCLIKit`, `Sources/OutlookAXCLI` executable product
- **JSON v2 typed envelopes** — `JSONSuccess`/`JSONFailure` with native numbers and booleans
  (`notifications --json` now returns `"count": 3` as an integer, not `"3"` as a string)
- **Opt-in JSON v2** — `--json-version 2` selects v2 in v1.1; `--json-version 1` for legacy compat
- **Typed `CLICommand` enum** with `CLIParser` — argument parsing separated from execution
- **Typed `OutlookAXError`** with exit codes (77 permission, 69 unavailable, 1 failure, 64 usage)
- **Shared L10n catalog** — `Sources/OutlookAX/Localization/L10n.swift` with verified de/en/fr/es/it variants
- **`CalendarTemporal` model** — normalized temporal values with explicit resolution states
  (no silent year inference; unresolved = explicit `.unresolved` + reason)
- **`CalendarTemporalParser`** — 12-hour/24-hour, all-day, year-boundary, DST-safe parsing
- **Passive connection** — `status` and `notifications` no longer launch/activate Outlook
- **Stop-loss safety**:
  - `calendar create` no longer discards pre-existing event drafts
  - `calendar create` and `mail compose` block Send/Save after any field failure
  - `mail inbox --limit` strictly validates 1…100; exits 64 on invalid values
  - `mail inbox` verifies Inbox is selected; removes unsafe `AXRow` fallback
  - Window cleanup removed from `calendar today` / event detail reads
- **`scripts/check-repository-layout.sh`** — validates `.claude/settings.json` and script syntax
- **`scripts/check-architecture.sh`** — validates SwiftPM architecture conventions
- **GitHub Actions CI** — macOS hosted runner, build + test + help + layout + architecture checks
- **Docs** — `docs/json-schema-v2.md`, `docs/calendar-temporal-contract.md`

### Fixed

- `notifications --json` returned count as string `"3"`; now returns integer `3` (#26)
- `calendar create` discarded existing event drafts without confirmation (#2)
- Form commands (compose, calendar create) reported success after failed AX writes (#3)
- `mail inbox --limit 0` (and negative/non-integer values) did not return a usage error (#17)
- `mail inbox` read whichever folder was currently selected instead of verifying Inbox (#11)
- `mail inbox` treated every `AXRow` in the window as an email message (#12)
- `status` and `notifications` launched/activated Outlook when it was in the background (#4)
- Unknown commands and bare `mail`/`calendar` groups exited with status 0 instead of 64 (#18)
- `wins[0]` assumption in `OutlookAX` library replaced with semantic selectors (#20)
- Claude SessionStart hook was silently ignored because the settings filename was wrong (#23)

## [1.0.0] — 2026-04-18

### Added

- 42 CLI commands covering mail, calendar, navigation, and system operations
- Localization system — `struct L10n` with 124 label arrays (de, en, fr, es, it)
- Mail commands — current, inbox, search, reply, reply-all, forward, delete, archive,
  compose, folders, folder, flag, read-unread, move, report, react, summarize, filter
- Calendar commands — today, create, view, navigate, calendars, toggle, timescale,
  filter, color, accept, tentative, decline, join, duplicate, categorize, private, show-as
- System commands — status, notifications, sync, auto-reply, my-day, account
- Auto-launch — ensureOutlookReady() launches and unminimizes Outlook via AppleScript
- JSON output — all commands support `--json` with English-normalized values
- Keyboard simulation — CGEvent-based typing for date/time fields
- L10n-aware menu triggering — triggerMenuL10n accepts multi-language path arrays
- Agent documentation — AGENTS.md, docs/architecture.md, docs/ax-paths.md
- AX discovery skill — .agents/skills/ax-discovery/SKILL.md
- Skills sync pattern — .agents/skills/ → .claude/skills/ via SessionStart hook

[Unreleased]: https://github.com/trsdn/outlook-ax/compare/v1.1.0...HEAD
[1.1.1]: https://github.com/trsdn/outlook-ax/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/trsdn/outlook-ax/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/trsdn/outlook-ax/releases/tag/v1.0.0
