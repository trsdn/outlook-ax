import AppKit
import ApplicationServices
import Foundation

// MARK: - CalendarService
//
// Calendar commands not already covered by the OutlookAX enum's static methods.
// OutlookAX.readCalendarList(), OutlookAX.switchCalendarView(), and
// OutlookAX.navigateCalendar() are used for the read and view-switch paths.
//
// Form safety rules (issues #2, #3):
// - cmdCalendarCreate uses before/after window identity to find the editor.
// - All field failures abort BEFORE Send or Save is attempted.

public enum CalendarService {

    // MARK: - Create Event

    /// Open a new event editor, fill in the requested fields, and optionally send.
    /// FORM SAFETY: Any requested field failure throws before Send/Save is called.
    public static func createEvent(
        subject: String,
        attendee: String?,
        date: String?,
        time: String?,
        send: Bool,
        conn: OutlookConnection
    ) throws {
        let axApp = conn.app
        let calWin = try conn.calendarWindow()

        // Snapshot all existing window titles before opening the new editor.
        let priorTitles = Set(ConnectionManager.refreshWindows(axApp).map { axTitle($0) })

        // Open the new event using the button on the calendar window only.
        guard let newBtn = axFind(calWin, where: {
            axStartsWithAny(axDesc($0), L10n.newEvent) && axRole($0) == "AXButton" && !axDesc($0).isEmpty
        }) else { throw OutlookAXError.formFieldNotFound(field: "New Event button") }
        AXUIElementPerformAction(newBtn, kAXPressAction as CFString)
        Thread.sleep(forTimeInterval: 2.5)

        // Identify the new editor by before/after identity.
        // Exactly one new window matching event-editor titles must have appeared.
        let winsAfter = ConnectionManager.refreshWindows(axApp)
        let newWindows = winsAfter.filter {
            !priorTitles.contains(axTitle($0)) && axMatchesAny(axTitle($0), L10n.eventWindowTitles)
        }
        switch newWindows.count {
        case 0: throw OutlookAXError.detailWindowNotOpened
        case 1: break
        default: throw OutlookAXError.ambiguousWindow
        }
        let evWin = newWindows[0]
        AXUIElementPerformAction(evWin, kAXRaiseAction as CFString)
        Thread.sleep(forTimeInterval: 0.5)

        // Fill subject — verify postcondition.
        guard let subjField = axFind(evWin, where: {
            axEqualsAny(axDesc($0), L10n.subject) && axRole($0) == "AXTextField"
        }) else { throw OutlookAXError.formFieldNotFound(field: "Subject") }
        AXUIElementSetAttributeValue(subjField, kAXFocusedAttribute as CFString, true as CFTypeRef)
        Thread.sleep(forTimeInterval: 0.3)
        let subjectRC = AXUIElementSetAttributeValue(subjField, kAXValueAttribute as CFString, subject as CFTypeRef)
        Thread.sleep(forTimeInterval: 0.3)
        guard subjectRC == .success && axValue(subjField) == subject else {
            throw OutlookAXError.formFieldWriteFailed(field: "Subject")
        }

        // Fill attendee (optional but fail-closed if provided).
        if let attendeeEmail = attendee {
            guard let attField = axFind(evWin, where: {
                axEqualsAny(axDesc($0), L10n.addAttendees) && axRole($0) == "AXTextField"
            }) else { throw OutlookAXError.formFieldNotFound(field: "Attendees") }
            AXUIElementSetAttributeValue(attField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.5)
            axTypeText(attendeeEmail)
            Thread.sleep(forTimeInterval: 1.0)
            axPressTab()
            Thread.sleep(forTimeInterval: 1.0)
        }

        // Set date via keyboard simulation on AXDateTimeArea (AXValue write is ignored by Outlook).
        if let dateStr = date {
            guard let dateField = axFind(evWin, where: {
                axEqualsAny(axTitle($0), L10n.startDate) && axRole($0) == "AXDateTimeArea"
            }) else { throw OutlookAXError.formFieldNotFound(field: "Start date") }
            AXUIElementSetAttributeValue(dateField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            axSelectAll(); Thread.sleep(forTimeInterval: 0.1)
            axTypeText(dateStr); Thread.sleep(forTimeInterval: 0.5)
            axPressTab(); Thread.sleep(forTimeInterval: 0.3)
        }

        // Set start time.
        if let timeStr = time {
            guard let timeField = axFind(evWin, where: {
                axEqualsAny(axTitle($0), L10n.startTime) && axRole($0) == "AXDateTimeArea"
            }) else { throw OutlookAXError.formFieldNotFound(field: "Start time") }
            AXUIElementSetAttributeValue(timeField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            axSelectAll(); Thread.sleep(forTimeInterval: 0.1)
            axTypeText(timeStr); Thread.sleep(forTimeInterval: 0.5)
            axPressTab(); Thread.sleep(forTimeInterval: 0.3)
        }

        // Send or Save — reached only after all requested fields succeeded.
        if send {
            // Try Send first (meetings with attendees), then Save (single appointments).
            if let sendBtn = axFind(evWin, where: {
                axEqualsAny(axDesc($0), L10n.send) && axRole($0) == "AXButton"
            }) {
                AXUIElementPerformAction(sendBtn, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 1.0)
            } else if let saveBtn = axFind(evWin, where: {
                axEqualsAny(axDesc($0), L10n.save) && axRole($0) == "AXButton"
            }) {
                AXUIElementPerformAction(saveBtn, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 1.0)
            } else {
                throw OutlookAXError.formFieldNotFound(field: "Send/Save button")
            }
        }
        // If !send: form is left open for the user — do NOT save/discard.
    }

    // MARK: - Read Today

    /// Read today's calendar events from the calendar list view, filtered to the
    /// verified current date and sorted chronologically.
    ///
    /// - Parameter conn: Active Outlook connection.
    /// - Parameter referenceDate: The date used as "today" for filtering and year
    ///   context. Defaults to `Date()`. Pass a fixed date in tests for determinism.
    /// - Returns: Events for today only, all-day first then by start time.
    ///   An empty array is a valid result when today has no events.
    public static func readToday(conn: OutlookConnection, referenceDate: Date = Date()) throws -> [CalendarEvent] {
        let calWin = try conn.calendarWindow()

        // Best-effort navigate to today. Non-throwing: the calendar may already
        // show today, or the Today button may be absent in some view modes.
        if let todayBtn = axFind(calWin, where: {
            axStartsWithAny(axDesc($0), L10n.today) && axRole($0) == "AXButton"
        }) {
            AXUIElementPerformAction(todayBtn, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 0.8)
        }

        // Require calendar list view (AXTable with calendar events description).
        let tables = axFindAll(calWin, where: {
            axRole($0) == "AXTable" && axMatchesAny(axDesc($0), L10n.calendarEventsTable)
        }, maxDepth: 12)

        guard let table = tables.first else {
            throw OutlookAXError.eventsTableNotFound
        }

        let allEvents = parseEventsTable(table, referenceDate: referenceDate)
        let todayISO = isoDateString(from: referenceDate)
        // Filter to today's ISO date and sort chronologically.
        // Empty result is valid — today may have no scheduled events.
        return filterAndSort(allEvents, toISO: todayISO)
    }

    // MARK: - Filter and Sort (internal for testability)

    /// Filter `events` to those whose `temporal.isoDate` matches `todayISO`
    /// and sort the result chronologically (all-day first, then by start time).
    ///
    /// Falls back to the first date-header group when no event carries an ISO date
    /// (e.g. unsupported locale prevents extraction). Never silently infers a date
    /// that was resolved to a different day.
    static func filterAndSort(_ events: [CalendarEvent], toISO todayISO: String) -> [CalendarEvent] {
        // Primary: match by normalized ISO date extracted by CalendarTemporalParser.
        var todayEvents = events.filter { $0.temporal?.isoDate == todayISO }

        // Fallback: when the parser could not extract any ISO date at all (all temporal
        // values have isoDate == nil), the first date-header group shown after navigation
        // is today's group. This handles unsupported locale/date formats gracefully.
        if todayEvents.isEmpty && !events.isEmpty &&
            events.allSatisfy({ $0.temporal?.isoDate == nil }) {
            if let firstHeader = events.first(where: { !$0.date.isEmpty })?.date {
                todayEvents = events.filter { $0.date == firstHeader }
            }
        }

        return sortChronologically(todayEvents)
    }

    /// Sort events chronologically: all-day events first, then timed events by
    /// 24-hour start time. Pure function.
    static func sortChronologically(_ events: [CalendarEvent]) -> [CalendarEvent] {
        events.sorted { a, b in
            // All-day events before timed events.
            if a.isAllDay != b.isAllDay { return a.isAllDay }
            // Prefer temporal start time for precise ordering.
            if let ta = a.temporal?.localStartTime, let tb = b.temporal?.localStartTime {
                return ta < tb
            }
            // Fallback: localized start string lexicographic comparison.
            return a.start < b.start
        }
    }

    // MARK: - Parse Calendar Events Table

    /// Parse an AXTable of calendar events into structured `CalendarEvent` values.
    ///
    /// - Parameter table: The AXTable element from the calendar list view.
    /// - Parameter referenceDate: Used to supply the calendar year for temporal
    ///   parsing when the date header does not include a year. Defaults to `Date()`.
    static func parseEventsTable(_ table: AXUIElement, referenceDate: Date = Date()) -> [CalendarEvent] {
        let rows = axChildren(table).filter { axRole($0) == "AXRow" }
        var items: [CalendarEvent] = []
        var currentDate = ""
        let year = Calendar.current.component(.year, from: referenceDate)

        for row in rows {
            let cells = axChildren(row).filter { axRole($0) == "AXCell" }
            guard let cell = cells.first else { continue }

            let cellDesc = axDesc(cell)
            let textValues = axChildren(cell)
                .filter { axRole($0) == "AXStaticText" }
                .compactMap { v -> String? in let s = axValue(v); return s.isEmpty ? nil : s }

            // Date header: empty desc, single text
            if cellDesc.isEmpty && textValues.count == 1 {
                currentDate = textValues[0]
                continue
            }
            guard !cellDesc.isEmpty else { continue }

            let rawTitle = textValues.first ?? ""
            var title = rawTitle
            var myResponse = "accepted"
            if axStartsWithAny(rawTitle, L10n.declinedPrefix) {
                myResponse = "declined"
                title = String(rawTitle.drop(while: { $0 != ":" }).dropFirst(2))
            } else if axStartsWithAny(rawTitle, L10n.followingPrefix) {
                myResponse = "following"
                title = String(rawTitle.drop(while: { $0 != ":" }).dropFirst(2))
            }

            var start = ""
            var end = ""
            var isAllDay = axMatchesAny(cellDesc, L10n.allDay)
            // Raw time range string (e.g. "09:00 - 10:00") captured from a verified
            // AX child — never from the event title.
            var rawTimeRange = ""

            // Parse time range from AXStaticText children — never from the title.
            for val in textValues {
                guard val.contains(" - ") else { continue }
                let parts = val.components(separatedBy: " - ")
                guard parts.count == 2 else { continue }
                let left = parts[0].trimmingCharacters(in: .whitespaces)
                let right = parts[1].trimmingCharacters(in: .whitespaces)
                if left.count <= 5 && left.contains(":") {
                    start = left; end = right
                    rawTimeRange = val.trimmingCharacters(in: .whitespaces)
                } else {
                    start = left; end = right; isAllDay = true
                }
            }

            // Parse status from cellDesc
            var status = ""
            var statusStart: String.Index? = nil
            for prefix in L10n.showAsPrefix {
                if let r = cellDesc.range(of: prefix) { statusStart = r.upperBound; break }
            }
            if let idx = statusStart {
                let statusStr = String(cellDesc[idx...]).trimmingCharacters(in: .whitespaces)
                if axStartsWithAny(statusStr, L10n.statusBusy) { status = "Busy" }
                else if axStartsWithAny(statusStr, L10n.statusFree) { status = "Free" }
                else if axStartsWithAny(statusStr, L10n.statusTentative) { status = "Tentative" }
                else if axStartsWithAny(statusStr, L10n.statusOOF) { status = "Out of Office" }
                else if axStartsWithAny(statusStr, L10n.statusElsewhere) { status = "Working Elsewhere" }
                else { status = String(statusStr.prefix(20)) }
            }

            // Parse organizer from cellDesc
            var organizer = ""
            var orgStart: String.Index? = nil
            for prefix in L10n.organizerPrefix {
                if let r = cellDesc.range(of: prefix) { orgStart = r.upperBound; break }
            }
            if let idx = orgStart {
                organizer = String(cellDesc[idx...].prefix(while: { $0 != "," }))
            } else if axMatchesAny(cellDesc, L10n.youAreOrganizer) {
                organizer = "You"
            }

            // Categories are in AXUnknown children with "Kategorie"/"Category" in desc
            let catElements = axChildren(cell).filter { axMatchesAny(axDesc($0), L10n.category) }
            let categories: [String] = catElements.flatMap { cat in
                axChildren(cat).compactMap {
                    let v = axValue($0); return v.isEmpty ? nil : v
                }
            }

            // Build normalized temporal value via CalendarTemporalParser.
            // Year is supplied from referenceDate so we never silently infer it.
            let temporal = CalendarTemporalParser.parse(
                timeRange: rawTimeRange,
                dateHeader: currentDate,
                year: year,
                isAllDay: isAllDay
            )

            items.append(CalendarEvent(
                title: title.trimmingCharacters(in: .whitespaces),
                date: currentDate,
                start: start, end: end,
                isAllDay: isAllDay,
                myResponse: myResponse,
                organizer: organizer,
                status: status,
                calendar: "",
                categories: categories,
                temporal: temporal
            ))
        }
        return items
    }

    // MARK: - Private: isoDateString

    /// Format a Date as an ISO 8601 date string ("yyyy-MM-dd") in the current time zone.
    private static func isoDateString(from date: Date) -> String {
        let df = DateFormatter()
        df.locale = Locale(identifier: "en_US_POSIX")
        df.timeZone = TimeZone.current
        df.dateFormat = "yyyy-MM-dd"
        return df.string(from: date)
    }

    // MARK: - View

    /// Switch to the named calendar view via the menu bar.
    public static func switchView(mode: String, conn: OutlookConnection) throws {
        let variants = viewVariants(for: mode)

        if MenuWalker.trigger(app: conn.app, path: [L10n.menuView, variants]) {
            Thread.sleep(forTimeInterval: 0.4)
            return
        }

        // Fallback: popup button
        let win = try conn.calendarWindow()
        guard let viewPopup = axFind(win, where: {
            axStartsWithAny(axDesc($0), L10n.calendarViewPicker) && axRole($0) == "AXPopUpButton"
        }) else { throw OutlookAXError.viewSwitchFailed(target: mode) }
        AXUIElementPerformAction(viewPopup, kAXPressAction as CFString)
        Thread.sleep(forTimeInterval: 0.5)
        axTypeText(variants.first ?? mode)
        Thread.sleep(forTimeInterval: 0.3)
        axPressReturn()
        Thread.sleep(forTimeInterval: 0.5)
    }

    // MARK: - Navigate

    public static func navigate(direction: String?, date: String?, conn: OutlookConnection) throws {
        let win = try conn.calendarWindow()

        if let dir = direction {
            switch dir.lowercased() {
            case "today", "heute":
                guard axPressButtonAny(win, prefixes: L10n.today) else {
                    throw OutlookAXError.navigationFailed("Today button not found")
                }
                Thread.sleep(forTimeInterval: 0.5)
            case "next", "forward":
                guard axPressButtonAny(win, prefixes: L10n.nextDay) else {
                    throw OutlookAXError.navigationFailed("Next day button not found")
                }
                Thread.sleep(forTimeInterval: 0.3)
            case "prev", "previous", "back":
                guard axPressButtonAny(win, prefixes: L10n.prevDay) else {
                    throw OutlookAXError.navigationFailed("Previous day button not found")
                }
                Thread.sleep(forTimeInterval: 0.3)
            default:
                throw OutlookAXError.navigationFailed("Unknown direction '\(dir)'. Use: today, next, prev")
            }
        } else if let dateStr = date {
            if let dateBtn = axFind(win, where: {
                axRole($0) == "AXButton" && axDesc($0).contains(dateStr)
            }) {
                AXUIElementPerformAction(dateBtn, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 0.5)
            } else {
                throw OutlookAXError.navigationFailed("Date '\(dateStr)' not found in mini calendar")
            }
        } else {
            throw OutlookAXError.navigationFailed("Specify --direction or --date")
        }
    }

    // MARK: - Timescale

    public static func setTimescale(minutes: String, conn: OutlookConnection) throws {
        let valid = ["60", "30", "15", "10", "6", "5"]
        guard valid.contains(minutes) else {
            throw OutlookAXError.viewSwitchFailed(target: "timescale \(minutes)")
        }
        for suffix in L10n.minutesSuffix {
            if MenuWalker.trigger(app: conn.app, path: [L10n.menuView, ["\(minutes) \(suffix)"]]) {
                return
            }
        }
        throw OutlookAXError.viewSwitchFailed(target: "timescale \(minutes)")
    }

    // MARK: - Filter

    public static func setFilter(filter: String, conn: OutlookConnection) throws {
        let variants = filterVariants(for: filter)
        if MenuWalker.trigger(app: conn.app, path: [L10n.menuView, L10n.menuFilter, variants]) { return }
        throw OutlookAXError.viewSwitchFailed(target: "filter \(filter)")
    }

    // MARK: - Color

    public static func setColor(color: String, conn: OutlookConnection) throws {
        let variants = colorVariants(for: color)
        if MenuWalker.trigger(app: conn.app, path: [L10n.menuView, L10n.menuColor, variants]) { return }
        throw OutlookAXError.viewSwitchFailed(target: "color \(color)")
    }

    // MARK: - Calendars

    /// List all calendars in the navigation pane with visibility and group.
    public static func listCalendars(conn: OutlookConnection) throws -> [CalendarIdentity] {
        let win = try conn.calendarWindow()
        var calendars: [CalendarIdentity] = []
        var currentGroup = ""

        guard let outline = axFind(win, where: {
            axRole($0) == "AXOutline" && axEqualsAny(axDesc($0), L10n.navPane)
        }) else { return [] }

        let rows = axFindAll(outline, where: { axRole($0) == "AXRow" }, maxDepth: 3)
        for row in rows {
            let checkboxes = axFindAll(row, where: { axRole($0) == "AXCheckBox" }, maxDepth: 3)
            let triangles = axFindAll(row, where: { axRole($0) == "AXDisclosureTriangle" }, maxDepth: 3)

            for cb in checkboxes {
                let d = axDesc(cb); let v = axValue(cb)
                if !triangles.isEmpty && (axMatchesAny(d, L10n.myCalendars) ||
                    axMatchesAny(d, L10n.otherCalendars) || d.contains("@")) {
                    currentGroup = v
                    continue
                }
                if !v.isEmpty {
                    let visible = axMatchesAny(d, L10n.calendarShown)
                    calendars.append(CalendarIdentity(
                        name: v, account: "", group: currentGroup, isVisible: visible
                    ))
                }
            }
        }
        return calendars
    }

    /// Toggle calendar visibility by name.
    /// Throws `.folderAmbiguous` when multiple calendars share the same name.
    public static func toggleCalendar(name: String, conn: OutlookConnection) throws -> Bool {
        let win = try conn.calendarWindow()
        guard let outline = axFind(win, where: {
            axRole($0) == "AXOutline" && axEqualsAny(axDesc($0), L10n.navPane)
        }) else { throw OutlookAXError.folderAmbiguous(name: name) }

        let matches = axFindAll(outline, where: {
            axRole($0) == "AXCheckBox" && axValue($0).lowercased() == name.lowercased()
        })
        switch matches.count {
        case 0: throw OutlookAXError.folderAmbiguous(name: name)
        case 1:
            AXUIElementPerformAction(matches[0], kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 0.5)
            return axMatchesAny(axDesc(matches[0]), L10n.calendarShown)
        default:
            throw OutlookAXError.folderAmbiguous(name: name)
        }
    }

    // MARK: - Event Actions (via menu bar)

    public static func accept(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.accept]) else {
            throw OutlookAXError.formFieldNotFound(field: "Accept menu item")
        }
    }

    public static func tentative(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.tentative]) else {
            throw OutlookAXError.formFieldNotFound(field: "Tentative menu item")
        }
    }

    public static func decline(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.decline]) else {
            throw OutlookAXError.formFieldNotFound(field: "Decline menu item")
        }
    }

    public static func join(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.joinMeeting]) else {
            throw OutlookAXError.formFieldNotFound(field: "Join Meeting menu item")
        }
    }

    public static func duplicate(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.duplicateEvent]) else {
            throw OutlookAXError.formFieldNotFound(field: "Duplicate Event menu item")
        }
    }

    public static func categorize(category: String?, conn: OutlookConnection) throws {
        if let cat = category {
            guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.menuCategorize, [cat]]) else {
                throw OutlookAXError.formFieldNotFound(field: "Category '\(cat)'")
            }
        } else {
            guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.menuCategorize]) else {
                throw OutlookAXError.formFieldNotFound(field: "Categorize menu")
            }
        }
    }

    public static func setPrivate(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.menuPrivate]) else {
            throw OutlookAXError.formFieldNotFound(field: "Private menu item")
        }
    }

    public static func setShowAs(status: String, conn: OutlookConnection) throws {
        let variants = showAsVariants(for: status)
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuEvent, L10n.menuShowAs, variants]) else {
            throw OutlookAXError.viewSwitchFailed(target: "show-as \(status)")
        }
    }

    // MARK: - System Commands

    public static func sync(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuTools, L10n.menuSync]) else {
            throw OutlookAXError.navigationFailed("Sync menu item not found")
        }
    }

    public static func openAutoReply(conn: OutlookConnection) throws {
        guard MenuWalker.trigger(app: conn.app, path: [L10n.menuTools, L10n.menuAutoReply]) else {
            throw OutlookAXError.navigationFailed("Auto-reply menu item not found")
        }
    }

    public static func toggleMyDay(conn: OutlookConnection) throws {
        let win = try conn.calendarWindow()
        guard let btn = axFind(win, where: {
            axStartsWithAny(axDesc($0), L10n.myDay) && axRole($0) == "AXButton"
        }) else { throw OutlookAXError.formFieldNotFound(field: "My Day button") }
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
    }

    public static func switchAccount(name: String, conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        if let btn = axFind(win, where: {
            (axDesc($0).lowercased().contains(name.lowercased()) ||
             axValue($0).lowercased().contains(name.lowercased())) && axRole($0) == "AXButton"
        }) {
            AXUIElementPerformAction(btn, kAXPressAction as CFString)
            return
        }
        // Fallback: profile menu
        if MenuWalker.trigger(app: conn.app, path: [["Profile"], [name]]) { return }
        throw OutlookAXError.formFieldNotFound(field: "Account '\(name)'")
    }

    // MARK: - Navigation

    public static func navigateTo(target: String, conn: OutlookConnection) throws {
        let variants = navVariants(for: target)
        let win = try conn.mainMailWindow()

        for role in ["AXRadioButton", "AXButton", "AXTab"] {
            if let btn = axFind(win, where: {
                let d = axDesc($0); let t = axTitle($0)
                return (axEqualsAny(d, variants) || axEqualsAny(t, variants) || axStartsWithAny(d, variants))
                    && axRole($0) == role
            }) {
                AXUIElementPerformAction(btn, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 1.5)
                return
            }
        }

        // Menu bar fallback: View > Switch to > <target>
        if MenuWalker.trigger(app: conn.app, path: [L10n.menuView, L10n.menuSwitchTo, variants]) {
            Thread.sleep(forTimeInterval: 1.5)
            return
        }
        throw OutlookAXError.navigationFailed("Navigation to '\(target)' failed")
    }

    // MARK: - Private Helpers

    private static func viewVariants(for mode: String) -> [String] {
        switch mode.lowercased() {
        case "day", "tag", "jour": return L10n.viewDay
        case "workweek", "arbeitswoche": return L10n.viewWorkWeek
        case "week", "woche", "semaine": return L10n.viewWeek
        case "month", "monat", "mois": return L10n.viewMonth
        case "threeday", "3day", "drei": return L10n.viewThreeDay
        case "list", "liste": return L10n.viewList
        default: return [mode]
        }
    }

    private static func filterVariants(for filter: String) -> [String] {
        switch filter.lowercased() {
        case "all", "alle": return L10n.filterAll
        case "appointments", "termine": return L10n.filterAppointments
        case "meetings", "besprechungen": return L10n.filterMeetings
        case "categories", "kategorien": return L10n.filterCategories
        case "recurring", "wiederholung": return L10n.filterRecurring
        case "privacy", "datenschutz": return L10n.filterPrivacy
        case "declined", "abgelehnte": return L10n.filterDeclined
        default: return [filter]
        }
    }

    private static func colorVariants(for color: String) -> [String] {
        switch color.lowercased() {
        case "blue", "blau": return L10n.colorBlue
        case "green", "grün", "gruen": return L10n.colorGreen
        case "orange": return L10n.colorOrange
        case "platinum", "platin": return L10n.colorPlatinum
        case "yellow", "gelb": return L10n.colorYellow
        case "cyan", "zyan": return L10n.colorCyan
        case "magenta": return L10n.colorMagenta
        case "brown", "braun": return L10n.colorBrown
        case "burgundy": return L10n.colorBurgundy
        case "teal", "meeresgrün": return L10n.colorTeal
        case "lilac", "flieder": return L10n.colorLilac
        default: return [color]
        }
    }

    private static func showAsVariants(for status: String) -> [String] {
        switch status.lowercased() {
        case "free", "frei": return L10n.showAsFree
        case "tentative", "vorbehalt": return L10n.showAsTentative
        case "busy", "gebucht": return L10n.showAsBusy
        case "oof", "ooo", "away": return L10n.showAsOOF
        case "elsewhere", "woanders": return L10n.showAsElsewhere
        default: return [status]
        }
    }

    private static func navVariants(for target: String) -> [String] {
        switch target.lowercased() {
        case "calendar", "kalender": return L10n.navCalendar
        case "mail", "email", "e-mail": return L10n.navMail
        case "people", "personen", "contacts": return L10n.navPeople
        case "todo", "tasks", "aufgaben": return L10n.navTasks
        case "copilot": return L10n.navCopilot
        case "onedrive": return L10n.navOneDrive
        case "favorites", "favoriten": return L10n.navFavorites
        case "org-explorer", "org": return L10n.navOrgExplorer
        default: return [target]
        }
    }
}
