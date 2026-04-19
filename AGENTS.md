# AGENTS.md — outlook-ax

> This file is for AI agents. It describes the project structure, conventions,
> and rules so you can make changes correctly without reading all 2000+ lines first.

## What This Is

A single-file Swift CLI (`outlook-ax.swift`) that controls Microsoft Outlook on macOS
via the Accessibility API (AXUIElement). No dependencies, no package manager.
Compiles with `swiftc -O outlook-ax.swift -o outlook-ax`.

## File Layout

```
outlook-ax.swift          # Everything — helpers, L10n, commands, argument parsing
Makefile                  # build / install / clean
README.md                 # Human-readable usage docs
LICENSE                   # MIT
skills/ax-discovery/      # Agent skill: how to explore Outlook's AX tree
docs/architecture.md      # Code structure, patterns, conventions (this companion)
docs/ax-paths.md          # Verified AX element paths by view
```

## Code Sections (in order)

| Line | MARK | Purpose |
|------|------|---------|
| ~7 | AX Helpers | `roleOf`, `titleOf`, `descOf`, `valueOf`, `childrenOf`, `findElement`, `findAll`, `collectText` |
| ~63 | — | `pressButton`, `pressButtonByTitle` — convenience wrappers |
| ~85 | Keyboard Simulation | `typeText`, `pressTab`, `pressReturn`, `pressEscape`, `pressCommandA` via CGEvent |
| ~177 | L10n | `matchesAny`, `startsWithAny`, `equalsAny`, `endsWithAny`, `pressButtonAny` + `struct L10n` with 124 label arrays |
| ~392 | Outlook Connection | `ensureOutlookReady()` (AppleScript launch/unminimize), `connectOutlook()`, `refreshWindows()`, `currentView()` |
| ~420 | JSON Output | `jsonFlag`, `detailsFlag`, `ok()`, `fail()`, `printJSON()` helpers |
| ~450 | JSON Types | Codable structs: `EmailJSON`, `InboxItemJSON`, `EventJSON`, `AttendeeJSON`, `CalendarInfoJSON` |
| ~530 | Commands: Status | `cmdStatus()` |
| ~550 | Commands: Notifications | `cmdNotifications()` |
| ~570 | Commands: Mail | `cmdMailCurrent`, `cmdMailInbox`, `cmdMailSearch`, `cmdMailReply`, `cmdMailForward`, `cmdMailDelete`, `cmdMailArchive`, `cmdMailCompose`, `cmdMailFolders`, `cmdMailFolder` |
| ~900 | Commands: Calendar | `cmdCalendarToday` (biggest function ~300 lines), `cmdCalendarCreate`, `cmdCalendarView`, `cmdCalendarNavigate`, `cmdCalendarCalendars`, `cmdCalendarToggle` |
| ~1530 | Commands: Navigate | `cmdNavigate(to:)` — switches Outlook views |
| ~1600 | Menu Bar Helper | `triggerMenuL10n` (L10n-aware), `triggerMenu` (legacy wrapper) |
| ~1640 | Mail Commands (batch 2) | `cmdMailReplyAll`, `cmdMailFlag`, `cmdMailReadUnread`, `cmdMailMove`, `cmdMailReport`, `cmdMailReact`, `cmdMailSummarize`, `cmdMailFilter` |
| ~1730 | Calendar Commands (batch 2) | `cmdCalendarTimescale`, `cmdCalendarFilter`, `cmdCalendarColor`, `cmdCalendarAccept/Tentative/Decline`, `cmdCalendarJoin`, `cmdCalendarDuplicate`, `cmdCalendarCategorize`, `cmdCalendarPrivate`, `cmdCalendarShowAs` |
| ~1880 | System Commands | `cmdSync`, `cmdAutoReply`, `cmdMyDay`, `cmdAccount` |
| ~1950 | Argument Parsing | `argValue()`, `usage()` |
| ~2070 | Main | Top-level `switch` on `CommandLine.arguments` |

## Rules for Making Changes

### Adding a new command

1. Write `func cmdNewThing()` — place it in the appropriate MARK section
2. Use L10n for ALL label matching (never hardcode a single language)
3. Add `case` to the `switch cmd` block in Main
4. Add line to `usage()`
5. Use `ok()` / `fail()` for output, support `--json` via `jsonFlag`

### Adding L10n labels

- Add to `struct L10n` (line ~206)
- Convention: `[de, en, fr, es, it]` — German first, then English, then others
- Use the matching function that fits: `equalsAny` (exact), `startsWithAny` (prefix), `matchesAny` (contains), `endsWithAny` (suffix)

### Output conventions

- All JSON output values are normalized to **English** regardless of UI language
- Status values: `"Busy"`, `"Free"`, `"Tentative"`, `"Out of Office"`, `"Working Elsewhere"`
- Response values: `"accepted"`, `"declined"`, `"tentative"`, `"none"`
- View names: `"calendar"`, `"mail"`, `"people"`, `"unknown"`
- `ok(message, extra:)` for success, `fail(message)` for errors (exits with code 1)

### Menu-based commands

- Use `triggerMenuL10n(app, path: [L10n.menuX, L10n.menuY])` — each path element is an array of L10n variants
- Legacy `triggerMenu(app, path: ["exact", "strings"])` wraps triggerMenuL10n with single-element arrays

### Safety

- `ensureOutlookReady()` auto-launches and unminimizes — no need to check manually
- `connectOutlook()` calls `ensureOutlookReady()` automatically
- Mail compose opens draft but does NOT send (unless explicit `--send` on calendar create)
- Date/time fields require keyboard simulation (CGEvent), not AXValue writes

## Do NOT

- Add dependencies or Package.swift — this must stay a single `swiftc` compilation
- Hardcode labels in any single language — always use L10n arrays
- Output non-English values in JSON — normalize everything
- Call `AXPressAction` on Send/Delete buttons in discovery/exploration code
- Add `import Foundation` — it's already implicit via ApplicationServices
