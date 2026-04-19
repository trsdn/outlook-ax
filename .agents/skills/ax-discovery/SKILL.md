---
name: ax-discovery
description: >
  Discover and map Outlook UI elements via the macOS Accessibility API (AXUIElement).
  Use this skill to find new buttons, fields, or labels in Outlook's AX tree,
  identify missing L10n translations, and prototype new outlook-ax commands.
triggers:
  - ax discovery
  - outlook discovery
  - neue outlook elemente finden
  - find new outlook elements
  - l10n labels erweitern
  - extend l10n labels
---

# AX Discovery Skill for outlook-ax

## Purpose

Microsoft Outlook (new, WebView-based) exposes its UI via the macOS Accessibility API.
This skill teaches you how to explore that AX tree to:

1. **Find new UI elements** — buttons, fields, menus not yet mapped in outlook-ax
2. **Extend L10n** — discover labels in other languages to add to the `L10n` struct
3. **Prototype new commands** — write and test Swift code before adding to outlook-ax.swift

## Safety Rules

- **NEVER press Send, Delete, or destructive buttons** during discovery
- **Read-only exploration** — use `roleOf`, `descOf`, `titleOf`, `valueOf` only
- Only use `AXPressAction` on navigation elements (view switches, disclosure triangles)
- If unsure whether a button is destructive, **ask the user first**

## Prerequisites

- macOS with Accessibility permission granted to your terminal app
- Microsoft Outlook running and visible (not minimized)
- Xcode Command Line Tools (`swiftc` available)

## Discovery Script Template

Compile and run this to dump the AX tree of Outlook's front window.
Save it as a temporary `.swift` file, compile with `swiftc`, execute.

```swift
import ApplicationServices
import AppKit

func attr(_ e: AXUIElement, _ a: String) -> String {
    var ref: CFTypeRef?
    AXUIElementCopyAttributeValue(e, a as CFString, &ref)
    return ref as? String ?? ""
}
func children(_ e: AXUIElement) -> [AXUIElement] {
    var ref: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXChildrenAttribute as CFString, &ref)
    return ref as? [AXUIElement] ?? []
}

func dump(_ e: AXUIElement, depth: Int = 0, max: Int = 5) {
    let role  = attr(e, kAXRoleAttribute)
    let title = attr(e, kAXTitleAttribute)
    let desc  = attr(e, kAXDescriptionAttribute)
    let val   = attr(e, kAXValueAttribute)
    let pad   = String(repeating: "  ", count: depth)

    // Skip empty nodes
    if !title.isEmpty || !desc.isEmpty || !val.isEmpty || depth < 2 {
        var parts = [role]
        if !title.isEmpty { parts.append("title=\"\(title)\"") }
        if !desc.isEmpty  { parts.append("desc=\"\(desc)\"") }
        if !val.isEmpty && val.count < 80 { parts.append("val=\"\(val)\"") }
        print("\(pad)\(parts.joined(separator: " "))")
    }

    if depth < max {
        for child in children(e) { dump(child, depth: depth + 1, max: max) }
    }
}

// --- Configuration ---
let MAX_DEPTH = 5           // Increase for deeper exploration (6-8 for detail windows)
let FILTER_ROLE: String? = nil  // Set to "AXButton", "AXTextField", etc. to filter

guard let app = NSWorkspace.shared.runningApplications.first(where: {
    $0.bundleIdentifier == "com.microsoft.Outlook"
}) else { print("Outlook not running"); exit(1) }

let axApp = AXUIElementCreateApplication(app.processIdentifier)
var ref: CFTypeRef?
AXUIElementCopyAttributeValue(axApp, kAXWindowsAttribute as CFString, &ref)
guard let wins = ref as? [AXUIElement], let win = wins.first else {
    print("No Outlook windows"); exit(1)
}

print("=== Outlook AX Tree (depth \(MAX_DEPTH)) ===")
print("Window: \(attr(win, kAXTitleAttribute))")
dump(win, max: MAX_DEPTH)
```

### Usage

```bash
# Full tree dump (depth 5)
swiftc /tmp/ax-discover.swift -o /tmp/ax-discover && /tmp/ax-discover

# Deeper dive (edit MAX_DEPTH in script)
# Filter for specific roles (edit FILTER_ROLE in script)
```

### Filtered Discovery

To search for specific element types, add a filter before the print statement:

```swift
// Only show AXButton elements
if FILTER_ROLE != nil && role != FILTER_ROLE! { /* skip print, still recurse */ }
```

To search by description text:

```swift
// Only show elements whose desc contains a search term
let SEARCH = "calendar"  // case-insensitive
if !desc.lowercased().contains(SEARCH) && !title.lowercased().contains(SEARCH) {
    // skip print, still recurse children
}
```

## Understanding AX Elements

### Key Attributes

| Attribute | Function | Access in outlook-ax |
|-----------|----------|---------------------|
| `kAXRoleAttribute` | Element type | `roleOf(element)` |
| `kAXTitleAttribute` | Window/tab title | `titleOf(element)` |
| `kAXDescriptionAttribute` | Accessible label | `descOf(element)` |
| `kAXValueAttribute` | Current value/state | `valueOf(element)` |
| `kAXChildrenAttribute` | Child elements | `childrenOf(element)` |

### Common Outlook AX Roles

| Role | Where in Outlook | Example |
|------|-----------------|---------|
| `AXButton` | Toolbar, actions, navigation | Reply, Delete, Save, Send |
| `AXTextField` | Input fields | Search, Subject, Attendees |
| `AXStaticText` | Labels, content | Email body parts, header values |
| `AXTable` | Lists | Inbox message list, Calendar events |
| `AXRow` / `AXCell` | Table content | Individual messages, events |
| `AXWebArea` | Web content regions | Email body (Reading Pane) |
| `AXPopUpButton` | Dropdowns | Calendar view picker, Filter |
| `AXCheckBox` | Toggles | Calendar visibility in sidebar |
| `AXDateTimeArea` | Date/time pickers | Event start/end date and time |
| `AXDisclosureTriangle` | Expand/collapse | Folder groups, calendar groups |
| `AXOutline` | Tree structure | Folder sidebar, calendar sidebar |
| `AXRadioButton` | Navigation tabs | Mail/Calendar/People tabs |

### Important: Labels Come from the Web Layer

Outlook's AX labels (`desc`, `title`) are set by the **web/React layer**, not by native macOS `.strings` files.
This means:
- Labels change with Outlook's server-side deployments (no app update needed)
- Labels match the user's Outlook language setting, not macOS system language
- The `.lproj/*.strings` files in the app bundle are for **legacy native views only**

## Extending L10n

The `L10n` struct in `outlook-ax.swift` (line ~206) contains static arrays of localized label variants.

### Convention

Each array lists variants in order: `[de, en, fr, es, it, ...]`

```swift
static let exampleLabel = ["German", "English", "French", "Spanish", "Italian"]
```

### How to Add a New Language Variant

1. **Run discovery** with Outlook set to the target language
2. **Find the element** — note the `desc` or `title` value
3. **Add the variant** to the matching `L10n` array:

```swift
// Before:
static let save = ["Speichern", "Save", "Enregistrer", "Guardar"]
// After (added Italian):
static let save = ["Speichern", "Save", "Enregistrer", "Guardar", "Salva"]
```

4. **Compile and test**: `make build && ./outlook-ax status`

### How to Add a Completely New Label

1. Find the element via discovery, note the `desc`/`title` in multiple languages
2. Add a new `static let` to the `L10n` struct with all known variants
3. Use the appropriate matcher in command code:
   - `equalsAny(text, L10n.newLabel)` — exact match
   - `startsWithAny(text, L10n.newLabel)` — prefix match
   - `matchesAny(text, L10n.newLabel)` — contains match
   - `endsWithAny(text, L10n.newLabel)` — suffix match

### L10n Audit Technique

To check which L10n labels actually match in the current UI language, run discovery
and cross-reference with the L10n arrays. Look for:
- Buttons/fields that don't match any existing L10n entry → need new labels
- Elements that match partially → may need more variants
- Known L10n entries with no match → possible Outlook UI change

## Adding a New Command

### Step 1: Write the Function

Add before `// MARK: - Argument Parsing` (line ~1930):

```swift
func cmdNewFeature() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    // Use L10n for all label matching
    if let btn = findElement(win, matching: {
        equalsAny(descOf($0), L10n.newLabel) && roleOf($0) == "AXButton"
    }) {
        // ... action ...
        ok("Feature done", extra: ["key": "value"])
    } else {
        fail("Element not found")
    }
}
```

### Step 2: Add to Argument Parser

In the `switch cmd` block (line ~2051), add a case:

```swift
case "newfeature":
    cmdNewFeature()
```

Or for nested commands (e.g., `outlook-ax mail newfeature`), add inside the existing
`case "mail":` block's inner switch.

### Step 3: Update Usage

Add a line to the `usage()` function (line ~1940).

### Step 4: Build and Test

```bash
make build
./outlook-ax newfeature --json
```

## Verified AX Paths Reference

These paths have been tested and confirmed working.

### Mail View

| Element | Attribute | L10n Key |
|---------|-----------|----------|
| Message header container | `title` | `L10n.messageHeader` |
| Header details | `desc` | `L10n.headerDetails` |
| From field | `desc="messageHeaderFromContent"` | (stable, not localized) |
| Recipients | `desc="messageHeaderRecipientsContent"` | (stable, not localized) |
| Sent date | `title` starts with | `L10n.sentPrefix` |
| Reading pane (body) | `desc="Reading Pane"` | (stable, not localized) |
| Search field | `desc` | `L10n.search` |
| Message list table | `desc` | `L10n.messageList` |

### Calendar View

| Element | Attribute | L10n Key |
|---------|-----------|----------|
| Events table | `desc` contains | `L10n.calendarEventsTable` |
| New event button | `desc` starts with | `L10n.newEvent` |
| Subject field | `desc` | `L10n.subject` |
| Attendees field | `desc` | `L10n.addAttendees` |
| Start date | `title` | `L10n.startDate` |
| Start time | `title` | `L10n.startTime` |
| Save button | `desc` | `L10n.save` |
| Send button | `desc` | `L10n.send` |
| Discard button | `desc` | `L10n.discard` |
| Today button | `desc` starts with | `L10n.today` |
| Next day button | `desc` starts with | `L10n.nextDay` |
| Previous day button | `desc` starts with | `L10n.prevDay` |
| View picker | `desc` starts with | `L10n.calendarViewPicker` |
| Nav pane (sidebar) | `desc` | `L10n.navPane` |

### Calendar Event Detail Window

| Element | Attribute | L10n Key |
|---------|-----------|----------|
| Organizer section | `val` | `L10n.sectionOrganizer` |
| Required attendees | `val` | `L10n.sectionRequired` |
| Optional attendees | `val` | `L10n.sectionOptional` |
| Response counts | `val` ends with | `L10n.respAccepted`, `respDeclined`, etc. |
| Categories | `desc` contains | `L10n.category` |
| All-day indicator | cell `desc` contains | `L10n.allDay` |
| Show-as prefix | cell `desc` contains | `L10n.showAsPrefix` |

### Navigation

| Target | L10n Key | Element role |
|--------|----------|-------------|
| Calendar | `L10n.navCalendar` | AXRadioButton / AXButton |
| Mail | `L10n.navMail` | AXRadioButton / AXButton |
| People | `L10n.navPeople` | AXRadioButton / AXButton |
| Tasks | `L10n.navTasks` | AXRadioButton / AXButton |
| My Day | `L10n.myDay` | AXButton |

### Menu Bar (via `triggerMenuL10n`)

| Menu | L10n Key |
|------|----------|
| View menu | `L10n.menuView` |
| Event menu | `L10n.menuEvent` |
| Tools menu | `L10n.menuTools` |
| Switch to submenu | `L10n.menuSwitchTo` |
| Show As submenu | `L10n.menuShowAs` |
| Categorize submenu | `L10n.menuCategorize` |

## Keyboard Simulation (for Date/Time Fields)

`AXDateTimeArea` elements don't accept `AXValue` writes. Use keyboard simulation:

```swift
// 1. Focus the element
AXUIElementPerformAction(element, kAXPressAction as CFString)
Thread.sleep(forTimeInterval: 0.2)

// 2. Select all existing text
pressCommandA()  // Cmd+A

// 3. Type the new value
typeText("20.04.2026")  // Replaces selection

// 4. Tab to next field
pressTab()
```

These helper functions (`pressCommandA`, `typeText`, `pressTab`) are defined in outlook-ax.swift
using `CGEvent(keyboardEventSource:virtualKey:keyDown:)` with `keyboardSetUnicodeString`.
