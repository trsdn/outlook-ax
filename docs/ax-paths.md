# Verified AX Paths

> For AI agents. Maps Outlook UI elements to their AX attributes and L10n keys.
> These paths were discovered and verified on Outlook for Mac (new, WebView-based).

## Important Context

- AX labels come from Outlook's **web/React layer**, not native `.strings` files
- Labels change with Outlook's server-side deployments (no app update needed)
- Labels match the user's **Outlook language setting**, not macOS system language
- The `.lproj/*.strings` files in the app bundle are for legacy native views only

## Mail View

### Message List

| Element | Role | Attribute | L10n Key | Notes |
|---------|------|-----------|----------|-------|
| Message list table | AXTable | `desc` | `L10n.messageList` | Contains AXRow children |
| Message row | AXRow | — | — | Children are AXCell with desc text |
| Message cell | AXCell | `desc` contains | `L10n.composeWindow` | Fallback: any AXRow in table |

### Reading Pane (Current Email)

| Element | Role | Attribute | L10n Key | Notes |
|---------|------|-----------|----------|-------|
| Header container | AXGroup | `title` | `L10n.messageHeader` | Contains from, to, date |
| Header details | AXGroup | `desc` | `L10n.headerDetails` | Expanded header area |
| From | AXStaticText | `desc="messageHeaderFromContent"` | — | **Not localized** — stable identifier |
| Recipients | AXStaticText | `desc="messageHeaderRecipientsContent"` | — | **Not localized** — stable identifier |
| Sent date | AXStaticText | `title` starts with | `L10n.sentPrefix` | e.g. "Gesendet am: 18.04.2026" |
| Body | AXWebArea | `desc="Reading Pane"` | — | **Not localized** — use `collectText()` to extract |

### Search

| Element | Role | Attribute | L10n Key |
|---------|------|-----------|----------|
| Search field | AXTextField | `desc` | `L10n.search` |

### Mail Actions (Toolbar)

| Action | Role | Attribute | L10n Key |
|--------|------|-----------|----------|
| Reply | AXButton | `desc` or `title` | `L10n.reply` |
| Reply All | AXButton | `desc` or `title` | `L10n.replyAll` |
| Forward | AXButton | `desc` | `L10n.forward` |
| Delete | AXButton | `desc` | `L10n.delete` |
| Archive | AXButton | `desc` | `L10n.archive` |
| Flag | AXButton | `desc` | `L10n.flag` |
| Read/Unread | AXButton | `desc` starts with | `L10n.markRead` |
| Move | AXButton | `desc` | `L10n.move` |
| Report | AXButton | `desc` | `L10n.report` |
| React | AXButton | `desc` | `L10n.react` |
| Summarize | AXButton | `desc` | `L10n.summarize` |
| Filter/Sort | AXPopUpButton | `desc` | `L10n.filterSort` |
| More items | AXButton | `desc` | `L10n.moreItems` |

### Compose Window

| Element | Role | Attribute | L10n Key |
|---------|------|-----------|----------|
| New mail button | AXButton | `desc` | `L10n.newEmail` |
| Compose window | AXWindow | `title` contains | `L10n.composeWindow` |
| To field | AXTextField | `desc` or `title` | `L10n.toField` |
| Subject field | AXTextField | `desc` | `L10n.subject` |
| Body field | AXWebArea | `desc` | `L10n.bodyField` |

### Folder Sidebar

| Element | Role | Attribute | L10n Key |
|---------|------|-----------|----------|
| Favorites section | — | `val` contains | `L10n.favorites` |
| All accounts section | — | `val` contains | `L10n.allAccounts` |
| Groups section | — | `val` contains | `L10n.groups` |

## Calendar View

### Event List/Table

| Element | Role | Attribute | L10n Key | Notes |
|---------|------|-----------|----------|-------|
| Events table | AXTable | `desc` contains | `L10n.calendarEventsTable` | Main event container |
| Calendar window | AXWindow | `title` | `L10n.calendarWindow` | |
| All-day indicator | — | cell `desc` contains | `L10n.allDay` | |
| Show-as prefix | — | cell `desc` contains | `L10n.showAsPrefix` | Followed by status value |

### Event Status Values (in cell desc)

| Status | L10n Key | Normalized Output |
|--------|----------|-------------------|
| Busy | `L10n.statusBusy` | `"Busy"` |
| Free | `L10n.statusFree` | `"Free"` |
| Tentative | `L10n.statusTentative` | `"Tentative"` |
| Out of Office | `L10n.statusOOF` | `"Out of Office"` |
| Working Elsewhere | `L10n.statusElsewhere` | `"Working Elsewhere"` |

### Event myResponse (from title prefix)

| Response | L10n Key | Normalized Output |
|----------|----------|-------------------|
| Declined | `L10n.declinedPrefix` | `"declined"` |
| Following | `L10n.followingPrefix` | `"following"` |
| (no prefix) | — | `"accepted"` |

### Event Detail Window (opened by clicking event)

| Element | Role | Attribute | L10n Key |
|---------|------|-----------|----------|
| Organizer section | AXStaticText | `val` | `L10n.sectionOrganizer` |
| Required attendees | AXStaticText | `val` | `L10n.sectionRequired` |
| Optional attendees | AXStaticText | `val` | `L10n.sectionOptional` |
| Accepted count | AXStaticText | `val` ends with | `L10n.respAccepted` |
| Declined count | AXStaticText | `val` ends with | `L10n.respDeclined` |
| Not responded | AXStaticText | `val` ends with | `L10n.respNotResponded` |
| Tentative count | AXStaticText | `val` ends with | `L10n.respTentative` |
| Categories | — | `desc` contains | `L10n.category` |
| Organizer name | — | check | `L10n.organizerPrefix` / `L10n.youAreOrganizer` |
| Noise filter | — | `val` matches | `L10n.detailNoise` |

### Create Event Form

| Element | Role | Attribute | L10n Key | Input Method |
|---------|------|-----------|----------|--------------|
| New event button | AXButton | `desc` starts with | `L10n.newEvent` | AXPress |
| Subject field | AXTextField | `desc` | `L10n.subject` | AXValue write |
| Attendees field | AXTextField | `desc` | `L10n.addAttendees` | AXValue write |
| Start date | AXDateTimeArea | `title` | `L10n.startDate` | **Keyboard sim** |
| Start time | AXDateTimeArea | `title` | `L10n.startTime` | **Keyboard sim** |
| Save button | AXButton | `desc` | `L10n.save` | AXPress |
| Send button | AXButton | `desc` | `L10n.send` | AXPress |
| Discard button | AXButton | `desc` | `L10n.discard` | AXPress |

### Calendar Navigation

| Element | Role | Attribute | L10n Key |
|---------|------|-----------|----------|
| Today button | AXButton | `desc` starts with | `L10n.today` |
| Next day button | AXButton | `desc` starts with | `L10n.nextDay` |
| Previous day button | AXButton | `desc` starts with | `L10n.prevDay` |
| View picker | AXPopUpButton | `desc` starts with | `L10n.calendarViewPicker` |

### Calendar Sidebar

| Element | Role | Attribute | L10n Key |
|---------|------|-----------|----------|
| Navigation pane | AXOutline | `desc` | `L10n.navPane` |
| My Calendars group | AXCheckBox | `desc` contains | `L10n.myCalendars` |
| Other Calendars | AXCheckBox | `desc` contains | `L10n.otherCalendars` |
| Shown indicator | — | `desc` contains | `L10n.calendarShown` |

## Navigation Bar

| Target | L10n Key | Roles to Try |
|--------|----------|-------------|
| Calendar | `L10n.navCalendar` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| Mail | `L10n.navMail` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| People | `L10n.navPeople` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| Tasks | `L10n.navTasks` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| Copilot | `L10n.navCopilot` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| OneDrive | `L10n.navOneDrive` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| Favorites | `L10n.navFavorites` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| Org Explorer | `L10n.navOrgExplorer` | AXRadioButton, AXButton, AXTab, AXMenuItem |
| My Day | `L10n.myDay` | AXButton |

## Menu Bar

### Top-Level Menus

| Menu | L10n Key |
|------|----------|
| View | `L10n.menuView` |
| Event | `L10n.menuEvent` |
| Tools | `L10n.menuTools` |

### Submenus

| Path | L10n Keys | Used By |
|------|-----------|---------|
| View → Switch to | `L10n.menuView` → `L10n.menuSwitchTo` | `cmdNavigate` fallback |
| View → Timescale → X min | `L10n.menuView` → `L10n.menuTimescale` → minutes | `cmdCalendarTimescale` |
| View → Filter → type | `L10n.menuView` → filter → `L10n.filterXxx` | `cmdCalendarFilter` |
| View → Color → color | `L10n.menuView` → `L10n.menuColor` → `L10n.colorXxx` | `cmdCalendarColor` |
| Event → Accept | `L10n.menuEvent` → `L10n.accept` | `cmdCalendarAccept` |
| Event → Tentative | `L10n.menuEvent` → `L10n.tentative` | `cmdCalendarTentative` |
| Event → Decline | `L10n.menuEvent` → `L10n.decline` | `cmdCalendarDecline` |
| Event → Join | `L10n.menuEvent` → `L10n.joinMeeting` | `cmdCalendarJoin` |
| Event → Duplicate | `L10n.menuEvent` → `L10n.duplicateEvent` | `cmdCalendarDuplicate` |
| Event → Categorize | `L10n.menuEvent` → `L10n.menuCategorize` | `cmdCalendarCategorize` |
| Event → Private | `L10n.menuEvent` → `L10n.menuPrivate` | `cmdCalendarPrivate` |
| Event → Show As → status | `L10n.menuEvent` → `L10n.menuShowAs` → `L10n.showAsXxx` | `cmdCalendarShowAs` |
| Tools → Sync | `L10n.menuTools` → `L10n.menuSync` | `cmdSync` |
| Tools → Auto-Reply | `L10n.menuTools` → `L10n.menuAutoReply` | `cmdAutoReply` |

## View Modes (Calendar View Picker)

| Mode | L10n Key |
|------|----------|
| Day | `L10n.viewDay` |
| Work Week | `L10n.viewWorkWeek` |
| Week | `L10n.viewWeek` |
| Month | `L10n.viewMonth` |
| Three Day | `L10n.viewThreeDay` |
| List | `L10n.viewList` |
