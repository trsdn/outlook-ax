# outlook-ax

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Platform: macOS](https://img.shields.io/badge/Platform-macOS-blue.svg)](https://developer.apple.com/macos/)
[![Swift](https://img.shields.io/badge/Swift-5.9+-orange.svg)](https://swift.org)
[![GitHub release](https://img.shields.io/github/v/release/trsdn/outlook-ax)](https://github.com/trsdn/outlook-ax/releases)

Lightweight CLI to read and control Microsoft Outlook (new/Electron) via the macOS Accessibility API. No dependencies, no AppleScript — pure AXUIElement access.

## Build

```bash
make build        # → ./outlook-ax
make install      # → /usr/local/bin/outlook-ax
```

Requires Xcode Command Line Tools (`xcode-select --install`).

## Permissions

The terminal running `outlook-ax` needs **Accessibility** permission:
System Settings → Privacy & Security → Accessibility → enable your terminal app.

## Commands

### Status & Notifications

```bash
outlook-ax status                    # Outlook running? Which view? Notification count
outlook-ax status --json
outlook-ax notifications             # Current notification count
outlook-ax notifications --json
```

### Mail — Read

```bash
outlook-ax mail current              # Read currently open email
outlook-ax mail current --json
outlook-ax mail inbox --limit 5      # List inbox messages (default 10)
outlook-ax mail search "Budget Q2"   # Search and return results
```

### Mail — Actions

```bash
outlook-ax mail reply                # Reply to current email
outlook-ax mail reply-all            # Reply-all to current email
outlook-ax mail forward              # Forward current email
outlook-ax mail delete               # Delete current email
outlook-ax mail archive              # Archive current email
outlook-ax mail flag                 # Toggle flag on current email
outlook-ax mail read                 # Toggle read/unread status
outlook-ax mail move                 # Open move-to-folder dialog
outlook-ax mail report               # Report as spam/phishing
outlook-ax mail react                # Open emoji reaction picker
outlook-ax mail summarize            # Copilot summarize current email
outlook-ax mail filter               # Open filter/sort popup
```

### Mail — Compose

```bash
outlook-ax mail compose \
  --to "team@example.com" \
  --subject "Update" \
  --body "See attached."             # Opens draft, does NOT send
```

### Mail — Folders

```bash
outlook-ax mail folders              # List all mail folders
outlook-ax mail folders --json
outlook-ax mail folder "Archiv"      # Switch to a specific folder
```

### Calendar — Read

```bash
outlook-ax calendar today            # Today's events
outlook-ax calendar today --json
```

### Calendar — Create Event

```bash
outlook-ax calendar create \
  --subject "Standup" \
  --attendee "team@example.com" \
  --date "21.04.2026" \
  --time "09:00"                     # Opens form, does NOT send

outlook-ax calendar create \
  --subject "1:1 Review" \
  --attendee "boss@example.com" \
  --date "22.04.2026" \
  --time "14:00" \
  --send                             # Creates AND sends invite
```

### Calendar — View & Navigation

```bash
# Switch view mode
outlook-ax calendar view day
outlook-ax calendar view week
outlook-ax calendar view month
outlook-ax calendar view arbeitswoche  # Work week
outlook-ax calendar view dreitage      # Three days
outlook-ax calendar view liste         # List view

# Navigate
outlook-ax calendar navigate --direction today
outlook-ax calendar navigate --direction next
outlook-ax calendar navigate --direction prev
outlook-ax calendar navigate --date "25.04.2026"

# Time grid
outlook-ax calendar timescale 30       # 60|30|15|10|6|5 minutes

# Filter events
outlook-ax calendar filter all         # all|appointments|meetings|categories|recurring|declined

# Calendar color
outlook-ax calendar color blue         # blue|green|orange|platinum|yellow|cyan|magenta|brown|burgundy|teal|lilac
```

### Calendar — Event Actions

```bash
outlook-ax calendar accept             # Accept meeting invite
outlook-ax calendar tentative          # Tentatively accept
outlook-ax calendar decline            # Decline meeting invite
outlook-ax calendar join               # Join online meeting (Teams etc.)
outlook-ax calendar duplicate          # Duplicate selected event
outlook-ax calendar categorize         # Open category picker
outlook-ax calendar categorize "Work"  # Apply specific category
outlook-ax calendar private            # Toggle private flag
outlook-ax calendar show-as busy       # free|busy|tentative|oof|elsewhere
```

### Calendar — Manage Calendars

```bash
outlook-ax calendar calendars          # List calendars with visibility state
outlook-ax calendar calendars --json
outlook-ax calendar toggle "Birthdays" # Toggle calendar visibility on/off
```

### Navigation

```bash
outlook-ax navigate calendar           # Switch to Calendar view
outlook-ax navigate mail               # Switch to Mail view
outlook-ax navigate people             # Switch to People/Contacts
outlook-ax navigate todo               # Switch to To Do
outlook-ax navigate copilot            # Switch to Copilot
outlook-ax navigate onedrive           # Switch to OneDrive/Files
outlook-ax navigate favorites          # Switch to Favorites
outlook-ax navigate org-explorer       # Switch to Org Explorer
```

### System

```bash
outlook-ax sync                        # Trigger Outlook sync
outlook-ax auto-reply                  # Open auto-reply/OOF settings
outlook-ax myday                       # Toggle My Day panel
outlook-ax account "Work"              # Switch account/profile
outlook-ax account "Personal"
```

## JSON Output

All read commands support `--json` for machine-readable output.

```bash
outlook-ax mail current --json | jq .subject
outlook-ax calendar today --json | jq '.[].title'
outlook-ax status --json | jq .notifications
outlook-ax calendar calendars --json | jq '.[] | select(.visible)'
```

## How it works

- Finds Outlook via `NSWorkspace.shared.runningApplications` (`com.microsoft.Outlook`)
- Reads UI state through `AXUIElement` tree traversal
- Text fields (subject, attendee, to) are set via `AXUIElementSetAttributeValue`
- Date/time fields use CGEvent keyboard simulation (AXDateTimeArea ignores AXValue)
- Navigation between views via sidebar radio buttons + menu bar fallback
- Mail actions (reply, forward, delete, archive, flag, etc.) trigger toolbar buttons via `AXPress`
- Calendar view switching via AXPopUpButton + text filtering
- Event actions and system commands via menu bar item triggering
- Menu bar access via `kAXMenuBarAttribute` for deep menu paths

## Limitations

- Outlook is auto-launched and unminimized when needed (via AppleScript)
- Calendar event parsing depends on the current calendar view layout
- Date format for `--date` must match Outlook's locale (DD.MM.YYYY for German)
- Inbox parsing heuristic — field order may vary across Outlook versions
- Multi-language support via L10n label matching (de, en, fr, es, it) — see `skills/ax-discovery/` for how to add more
- Mail compose opens a draft but does not auto-send (safety by design)
- Event actions (accept/decline/join) require an event to be selected or open
- Some menu items are context-dependent and only enabled in specific views

## Skills (Agent Instructions)

The `skills/` folder contains structured instructions for AI agents working with outlook-ax:

| Skill | Purpose |
|-------|---------|
| [ax-discovery](skills/ax-discovery/SKILL.md) | Find new UI elements, extend L10n labels, add new commands |

Skills follow the SKILL.md format (YAML frontmatter + Markdown body) and travel with the component when published as a standalone repository.
