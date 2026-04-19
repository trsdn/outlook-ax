# Architecture

> For AI agents modifying outlook-ax. Describes patterns, data flow, and design decisions.

## Single-File Design

Everything lives in `outlook-ax.swift` (~2100 lines). This is intentional:

- **No build system complexity** — `swiftc -O outlook-ax.swift -o outlook-ax` is the entire build
- **No dependency management** — only macOS system frameworks (ApplicationServices, AppKit, CoreGraphics)
- **Easy to audit** — one file, grep-able, no indirection layers

The file is organized top-to-bottom in dependency order: helpers → L10n → connection → types → commands → parsing → main.

## Data Flow

```
CLI args → switch dispatch → cmdXxx() → connectOutlook()
                                            ↓
                                    ensureOutlookReady()
                                    (AppleScript: launch + unminimize)
                                            ↓
                                    AXUIElement tree traversal
                                    (findElement / findAll with L10n matching)
                                            ↓
                                    ok() / fail() → stdout (JSON or text)
```

## AX Element Access Pattern

Every command follows the same pattern:

```swift
func cmdExample() {
    // 1. Connect (auto-launches Outlook)
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    // 2. Find element using L10n-aware matching
    if let elem = findElement(win, matching: {
        equalsAny(descOf($0), L10n.someLabel) && roleOf($0) == "AXButton"
    }) {
        // 3. Read or act
        let value = valueOf(elem)
        // 4. Output
        ok("Done", extra: ["key": value])
    } else {
        fail("Element not found")
    }
}
```

## L10n System

### Why arrays, not dictionaries

A simple `[String]` array per label is the lightest approach:
- No key-value mapping needed — we just need "does any variant match?"
- Helper functions (`matchesAny`, `equalsAny`, etc.) iterate the array
- Adding a language = appending one string to each array
- No runtime locale detection needed — we try all variants

### Matching Functions

| Function | Use when |
|----------|----------|
| `equalsAny(text, variants)` | Exact match — button desc, section headers |
| `startsWithAny(text, variants)` | Prefix — "Neue Benachrichtigungen: 3", date prefixes |
| `matchesAny(text, variants)` | Contains — label anywhere in longer text |
| `endsWithAny(text, variants)` | Suffix — response counts "5 accepted." |
| `pressButtonAny(win, descPrefixes:)` | Find + press first matching button |

### Output Normalization

The UI shows localized values. We normalize to English in output:

```
UI: "Gebucht" / "Occupé" / "Busy"  →  JSON: "Busy"
UI: "angenommen."                    →  JSON: "accepted"
UI: "Kalender"                       →  JSON: view: "calendar"
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
