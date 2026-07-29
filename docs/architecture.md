# Architecture

> For AI agents modifying outlook-ax. Describes patterns, data flow, and design decisions.

## SwiftPM Architecture

The project uses a three-target SwiftPM architecture:

```
Package.swift
├── OutlookAX          (library) — AX, connection, services, models, L10n, parsers
├── OutlookAXCLIKit    (library) — CLI parsing, dispatch, JSON envelopes, output
└── OutlookAXCLI       (executable) — entry point only (main.swift → exit())
```

Build: `swift build -c release --product outlook-ax`

> **No root outlook-ax.swift**: The single-file implementation has been fully migrated.
> All implementations live in `Sources/OutlookAX/`.

## Data Flow

```
CLI args → main.swift
  → CLIParser.parse() → CLIParseResult (.command / .help / .usageError)
  → CLIRunner.run(command) → ConnectionManager.passiveConnect() or .interactiveConnect()
      ↓
  MailService.xxx(conn:) or CalendarService.xxx(conn:)
      ↓
  AXHelpers: axFind() / axFindAll() / axPressButtonAny()
  L10n: axEqualsAny() / axMatchesAny() / axStartsWithAny()
      ↓
  throws OutlookAXError or returns typed result
      ↓
  CLIRunner → OutputRenderer → stdout (JSON v2 or plain text)
  CLIRunner returns Int32 exit code → main.swift calls exit()
```

## Connection Policies

Two connection types from `ConnectionManager`:

| Method | Use for | AX trust check | Launch? |
|--------|---------|---------------|---------|
| `passiveConnect()` | Status, read commands | Yes (throws 77) | No |
| `tryPassiveConnect()` | `.status` only (never fails) | Yes (returns nil) | No |
| `interactiveConnect()` | Compose, create, navigate | Yes (throws 77) | Yes |

## AX Element Access Pattern

Every service method follows the same pattern:

```swift
// In MailService or CalendarService
public static func readInbox(limit: Int, conn: OutlookConnection) throws -> [MailMessageSummary] {
    // 1. Semantic window selection (never wins[0])
    let win = try conn.mainMailWindow()

    // 2. Precondition check
    guard isInbox else { throw OutlookAXError.inboxNotSelected }

    // 3. Find element using L10n-aware matching via AXHelpers
    guard let table = axFind(win, where: {
        axRole($0) == "AXTable" && axEqualsAny(axDesc($0), L10n.messageList)
    }) else { throw OutlookAXError.messageListNotFound }

    // 4. Parse and return typed result
    return parseRows(table, limit: limit)
}
```

## Form Safety Pattern

For compose/create commands — enforced in both MailService and CalendarService:

```swift
// 1. Snapshot windows before opening editor
let priorTitles = Set(conn.wins.map { axTitle($0) })

// 2. Open editor (press button)
// ...

// 3. Find new editor by before/after identity (throws .ambiguousWindow if multiple)
let editor = try conn.uniqueNewWindow(before: priorTitles, matching: { ... })

// 4. Fill each requested field — throw immediately on failure
guard let field = axFind(editor, where: { ... }) else {
    throw OutlookAXError.formFieldNotFound(field: "Subject")
}
let rc = AXUIElementSetAttributeValue(field, kAXValueAttribute as CFString, value as CFTypeRef)
guard rc == .success else { throw OutlookAXError.formFieldWriteFailed(field: "Subject") }

// 5. Send/Save ONLY reached if all fields succeeded
if send { AXUIElementPerformAction(sendBtn, kAXPressAction as CFString) }
```

## L10n System

### Why arrays, not dictionaries

A simple `[String]` array per label is the lightest approach:
- No key-value mapping needed — we just need "does any variant match?"
- Helper functions (`axEqualsAny`, `axMatchesAny`, etc.) iterate the array
- Adding a language = appending one string to each array
- No runtime locale detection needed — we try all variants

### Matching Functions

| Function | Use when |
|----------|----------|
| `axEqualsAny(text, variants)` | Exact match — button desc, section headers |
| `axStartsWithAny(text, variants)` | Prefix — "New notifications: 3", date prefixes |
| `axMatchesAny(text, variants)` | Contains — label anywhere in longer text |
| `axEndsWithAny(text, variants)` | Suffix — response counts "5 accepted." |
| `axPressButtonAny(win, prefixes:)` | Find + press first matching button |

These are also available as `L10n.equals()`, `L10n.startsWith()`, `L10n.matches()`, `L10n.endsWith()`.

### Output Normalization

The UI shows localized values. We normalize to English in output:

```
UI: "Gebucht" / "Occupé" / "Busy"  →  JSON: "Busy"
UI: "angenommen."                    →  JSON: "accepted"
UI: "Kalender"                       →  JSON: view: "calendar"
```
```

This happens in each command function, not centrally. The pattern is:
```swift
if startsWithAny(raw, L10n.statusBusy) { status = "Busy" }
```

## Menu Triggering

Two-tier system:

1. **`triggerMenuL10n(app, path: [[String]])`** — each path element is an L10n array.
   Walks menu bar → top menu → submenu → item, trying all variants at each level.

2. **`triggerMenu(app, path: [String])`** — legacy wrapper, wraps each string in `[x]`.
   Use for user-provided strings (e.g., category names).

Menu paths are max 3 levels: `[topMenu, item]` or `[topMenu, submenu, item]`.

## Keyboard Simulation

AXDateTimeArea and some text fields don't accept `AXUIElementSetAttributeValue`.
We use CGEvent-based keyboard simulation:

```
Focus element → Cmd+A (select all) → type replacement text → Tab to next field
```

Key functions:
- `typeText(_ text: String)` — types each character via CGEvent with Unicode
- `pressTab()`, `pressReturn()`, `pressEscape()` — virtual key events
- `pressCommandA()` — select-all shortcut

## JSON Output

Two modes controlled by `jsonFlag` (`--json`):

- **Text mode** (default): Human-readable, printed via `print()`
- **JSON mode**: Codable structs serialized via `JSONEncoder` + `JSONSerialization`

Helper functions:
- `ok(_ msg, extra:)` — success output. Text mode prints message, JSON mode outputs `{"ok": true, ...extra}`
- `fail(_ msg)` — error output + `exit(1)`. JSON mode outputs `{"ok": false, "error": msg}`
- `printJSON(_ value: Encodable)` — direct JSON serialization for data responses

## Argument Parsing

Manual `CommandLine.arguments` parsing — no ArgumentParser dependency.

```
args[0] = top-level command ("mail", "calendar", "navigate", "status", ...)
args[1] = subcommand ("current", "inbox", "today", "create", ...)
--flags  = parsed via argValue("--flag") or .contains("--flag")
```

The `switch` in Main dispatches to `cmdXxx()` functions. Nested switches handle
two-level commands like `mail current`, `calendar create`.
