# AGENTS.md — outlook-ax

> This file is for AI agents. It describes the project structure, conventions,
> and rules so you can make changes correctly without reading all source files first.

## What This Is

A Swift CLI (`outlook-ax`) that controls Microsoft Outlook on macOS via the
Accessibility API (AXUIElement). Built as a Swift Package (SwiftPM).

```
swift build -c release --product outlook-ax
```

Produces a single distributable executable at `.build/release/outlook-ax`.
`make build` runs the above and copies the binary to `./outlook-ax`.

## Package Structure

```
Package.swift                    # SwiftPM manifest: OutlookAX + OutlookAXCLIKit + OutlookAXCLI
Makefile                         # build / install / clean / test / check
README.md                        # Human-readable usage docs
AGENTS.md                        # This file
CHANGELOG.md                     # Version history

Sources/OutlookAX/               # Library: AX, models, localization, parsers, services
  OutlookAX.swift                # Public API facade (backward-compatible library surface)
  AX/
    AXHelpers.swift              # Internal AX helpers: axRole, axTitle, axFind, axTypeText…
  Connection/
    ConnectionManager.swift      # passiveConnect(), interactiveConnect(), window selection
  Menu/
    MenuWalker.swift             # Menu bar traversal by L10n path arrays
  Mail/
    MailService.swift            # All mail commands: readInbox, compose, reply, etc.
  Calendar/
    CalendarService.swift        # All calendar commands: readToday, createEvent, navigate…
    CalendarTemporalParser.swift # Temporal parser: 12/24h, all-day, year, DST
  Models/
    Errors.swift                 # OutlookAXError with exit codes and machine codes
    MailModels.swift             # MailMessageSummary, MailMessage, FolderIdentity
    CalendarModels.swift         # CalendarEvent, Attendee, EventDetails, CalendarIdentity
    CalendarTemporal.swift       # CalendarTemporal, TemporalResolution, TimeZoneSource
  Localization/
    L10n.swift                   # Shared L10n catalog (de, en, fr, es, it)

Sources/OutlookAXCLIKit/         # CLI runtime: parsing, dispatch, JSON envelopes
  CLICommand.swift               # Typed CLICommand enum
  CLIParser.swift                # Argument parser → CLIParseResult (no exit())
  CLIRunner.swift                # Command runner → actual dispatch (no exit())
  JSONEnvelope.swift             # JSONSuccess, JSONFailure, typed payloads
  OutputRenderer.swift           # stdout/stderr output

Sources/OutlookAXCLI/            # Executable entry point
  main.swift                     # Only file that calls exit()

Tests/OutlookAXTests/            # Library tests: models, parsers, L10n
Tests/OutlookAXCLITests/         # CLI unit tests: parser, JSON envelope, exit codes
Tests/OutlookAXCLISubprocessTests/  # Subprocess tests: real binary exit codes, JSON shape
```

> **Important**: The root `outlook-ax.swift` has been removed. All AX, connection,
> mail, and calendar implementations live exclusively in `Sources/OutlookAX/`.
> Do NOT create a root-level Swift file as a shortcut — all additions go into
> the appropriate `Sources/` subdirectory.

## Rules for Making Changes

### Adding a new command

1. Add a case to `CLICommand` in `Sources/OutlookAXCLIKit/CLICommand.swift`
2. Add parsing logic to `CLIParser` in `Sources/OutlookAXCLIKit/CLIParser.swift`
3. Add a service method in `Sources/OutlookAX/Mail/MailService.swift` or
   `Sources/OutlookAX/Calendar/CalendarService.swift`
4. Add dispatch to `CLIRunner.run()` in `Sources/OutlookAXCLIKit/CLIRunner.swift`
5. Add tests in `Tests/OutlookAXCLITests/` and/or `Tests/OutlookAXCLISubprocessTests/`
6. Update `docs/` if the command adds new JSON output

### AX access

- All AX element access MUST go through `Sources/OutlookAX/` (library module)
- AXHelpers.swift provides internal helpers: `axRole()`, `axTitle()`, `axDesc()`,
  `axValue()`, `axChildren()`, `axFind()`, `axFindAll()`, `axTypeText()`, etc.
- `OutlookAXCLIKit` and `OutlookAXCLI` must NEVER import `ApplicationServices` directly

### Adding L10n labels

- Edit `Sources/OutlookAX/Localization/L10n.swift`
- Convention: `[de, en, fr, es, it]` — German first, then English, then others
- Each command-critical key MUST have variants for all five locales
- Use `L10n.equals()`, `L10n.startsWith()`, `L10n.matches()` (contains), `L10n.endsWith()`
- For AX matching use: `axEqualsAny()`, `axStartsWithAny()`, `axMatchesAny()`, `axEndsWithAny()`
- **Do NOT add localized strings to command files** — all UI labels go to L10n.swift

### Output conventions

- All JSON output values are normalized to **English** regardless of UI language
- Status values: `"Busy"`, `"Free"`, `"Tentative"`, `"Out of Office"`, `"Working Elsewhere"`
- Response values: `"accepted"`, `"declined"`, `"tentative"`, `"none"`, `"following"`
- View names: `"calendar"`, `"mail"`, `"people"`, `"unknown"`
- `--json` produces the v2 typed envelope; `--json-version 1` for legacy compat

### Connection policies

| Policy     | Use for                            | Effect                             |
|------------|------------------------------------|------------------------------------|
| passive    | Status, read, observational        | No launch/activation/unminimize    |
| interactive| Compose, create, navigate actions  | May launch/activate/focus          |

- `ConnectionManager.passiveConnect()` → throws if not running or no AX permission
- `ConnectionManager.tryPassiveConnect()` → returns nil instead of throwing (for status)
- `ConnectionManager.interactiveConnect()` → launches Outlook if needed, activates
- Never use an interactive connection for read-only commands.

### Window selection

- **Never** use `wins[0]` — use `OutlookConnection` semantic selectors instead
- `conn.mainMailWindow()` — semantic mail window (throws `.noAccessibleWindows`)
- `conn.calendarWindow()` — semantic calendar window (throws `.noAccessibleWindows`)
- `conn.uniqueNewWindow(before:matching:)` — before/after identity for new editors
- Fail with `OutlookAXError.ambiguousWindow` when multiple windows match

### Form safety

- Preflight every field before pressing Send/Save
- Check AX return codes after every write
- Verify postconditions (readable field value matches expected)
- **Never invoke Send/Save after any field failure**
- Fail with `OutlookAXError.sendBlockedByFieldFailure`
- Use before/after identity snapshots for new editor/compose window detection

### Error exit codes

| Code | Meaning                                      |
|------|----------------------------------------------|
| 0    | Success                                      |
| 1    | Operation failure                            |
| 64   | Usage / argument error                       |
| 69   | Outlook not running, no windows, wrong view  |
| 77   | Accessibility permission denied              |

### Safety

- Status/read commands MUST use passive connection (no activation)
- NEVER press Send, Delete, Discard, or broad Close in discovery or tests
- Date/time keyboard simulation is allowed only in an explicitly owned editor
- NEVER close a window that existed before the current operation
- Use before/after identity snapshots for new editor/compose window detection

### Architecture guards

Run `bash scripts/check-architecture.sh` to verify:
- Root `outlook-ax.swift` is absent (migration complete)
- No `wins[0]` in `Sources/`
- Single L10n definition in `Sources/`
- No direct AX calls outside `Sources/OutlookAX/`
- `ConnectionManager.swift` exists
- Subprocess tests exist

## Skills Path

Canonical skills live in `.agents/skills/`. Before using or listing skills,
run `bash scripts/sync-agent-skills.sh` to mirror them to `.claude/skills/`.
`.claude/skills/` is gitignored and only a local mirror.
The sync script works from any working directory.

## Claude Hook

`.claude/settings.json` (note: no dot prefix) configures the SessionStart hook
that runs `bash scripts/sync-agent-skills.sh` automatically.

## Build, Test, Check

```bash
make build                              # Release build + copy binary
swift test                              # Run all Swift tests (63 tests)
bash scripts/check-repository-layout.sh   # Layout checks
bash scripts/check-architecture.sh         # Architecture checks
make check                              # All of the above
```

> Note: `swift test` includes subprocess tests that run the real binary.
> `make build` must succeed before running `swift test` for subprocess tests to pass.

