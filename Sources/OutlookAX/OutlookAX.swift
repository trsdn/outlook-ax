import AppKit
import ApplicationServices
import Foundation

/// Narrow library surface for reading Microsoft Outlook's calendar via the
/// macOS Accessibility API. Mirrors the subset of `outlook-ax` CLI functionality
/// that consuming apps (OpenOats, …) need, without subprocess/TCC overhead.
///
/// Design rules for this target:
/// - NEVER activate Outlook or steal focus. Reads must work with Outlook in
///   the background; callers control activation themselves.
/// - All failure modes surface as `OutlookAXError` — no `exit()`, no prints.
public enum OutlookAX {

    // MARK: - Public types

    public enum CalendarViewMode: String, Sendable {
        case day, workWeek, week, month, threeDay, list
    }

    public struct CalendarEvent: Codable, Sendable, Equatable {
        public var title: String
        /// Localised date header as shown in list view (e.g. "Mittwoch, 15. April").
        public var date: String
        /// "HH:MM" for timed events, or localised day label for all-day.
        public var start: String
        public var end: String
        public var isAllDay: Bool
        /// One of "accepted", "declined", "following".
        public var myResponse: String
        public var organizer: String
        public var status: String
        public var calendar: String

        public init(
            title: String,
            date: String,
            start: String,
            end: String,
            isAllDay: Bool,
            myResponse: String,
            organizer: String,
            status: String,
            calendar: String
        ) {
            self.title = title
            self.date = date
            self.start = start
            self.end = end
            self.isAllDay = isAllDay
            self.myResponse = myResponse
            self.organizer = organizer
            self.status = status
            self.calendar = calendar
        }
    }

    public struct Attendee: Codable, Sendable, Equatable {
        public var name: String
        /// "organizer" | "required" | "optional"
        public var type: String
        /// "accepted" | "declined" | "tentative" | "none"
        public var response: String

        public init(name: String, type: String, response: String) {
            self.name = name; self.type = type; self.response = response
        }
    }

    public struct EventDetails: Codable, Sendable, Equatable {
        public var attendees: [Attendee]
        public var location: String
        public var body: String
        public var calendar: String
        public var organizer: String

        public init(
            attendees: [Attendee], location: String, body: String,
            calendar: String, organizer: String
        ) {
            self.attendees = attendees; self.location = location
            self.body = body; self.calendar = calendar; self.organizer = organizer
        }
    }

    public enum OutlookAXError: LocalizedError {
        case outlookNotRunning
        case notInCalendarView
        case viewSwitchFailed(target: CalendarViewMode)
        case eventsTableNotFound
        case navigationFailed(String)
        case eventRowNotFound(title: String)
        case detailWindowNotOpened

        public var errorDescription: String? {
            switch self {
            case .outlookNotRunning:
                return "Microsoft Outlook is not running."
            case .notInCalendarView:
                return "Outlook is not on the calendar view."
            case .viewSwitchFailed(let mode):
                return "Could not switch Outlook to \(mode.rawValue) view."
            case .eventsTableNotFound:
                return "Calendar events table not found (Outlook probably not in list view)."
            case .navigationFailed(let msg):
                return "Could not navigate Outlook: \(msg)"
            case .eventRowNotFound(let title):
                return "Could not find '\(title)' in Outlook's list view."
            case .detailWindowNotOpened:
                return "Outlook did not open the event detail window."
            }
        }
    }

    // MARK: - Public API

    /// Whether this process has macOS Accessibility permission.
    public static func hasAccessibilityPermission() -> Bool {
        AXIsProcessTrusted()
    }

    /// Returns true if Outlook is currently running.
    public static func isOutlookRunning() -> Bool {
        connect() != nil
    }

    /// Read the currently-selected calendar view mode, or nil if we can't
    /// determine it (e.g. Outlook is not on the Calendar tab). Uses the
    /// "Kalenderansicht" popup button's title, which reflects live UI state.
    public static func currentCalendarViewMode() -> CalendarViewMode? {
        guard let conn = connect() else { return nil }
        for win in conn.wins {
            guard let picker = findElement(win, matching: { el in
                guard roleOf(el) == "AXPopUpButton" else { return false }
                let d = descOf(el)
                return L10n.calendarViewPicker.contains(where: { d.hasPrefix($0) })
            }) else { continue }
            let label = titleOf(picker)
            if L10n.viewDay.contains(label) { return .day }
            if L10n.viewWorkWeek.contains(label) { return .workWeek }
            if L10n.viewWeek.contains(label) { return .week }
            if L10n.viewMonth.contains(label) { return .month }
            if L10n.viewThreeDay.contains(label) { return .threeDay }
            if L10n.viewList.contains(label) { return .list }
        }
        return nil
    }

    /// Switch Outlook to the Calendar view (from Mail, People, etc.).
    /// Uses sidebar/menu-bar AXPress — Outlook does not need to come to
    /// the foreground.
    @discardableResult
    public static func switchToCalendar() throws -> Bool {
        guard let conn = connect() else { throw OutlookAXError.outlookNotRunning }
        // Already on calendar? Verify by checking the calendar view popup is
        // present (window title alone can false-positive for Mail builds that
        // show "Kalender" somewhere in the tree).
        if currentCalendarViewMode() != nil { return true }
        if currentView(conn.wins) == "calendar" { return true }
        let win = conn.wins[0]

        // Sidebar button — try the common roles in order. Outlook may expose
        // the Calendar nav entry as any of AXRadioButton/AXButton/AXTab.
        for role in ["AXRadioButton", "AXButton", "AXTab", "AXMenuItem"] {
            if let btn = findElement(win, matching: {
                let d = descOf($0); let t = titleOf($0)
                return (equalsAny(d, L10n.navCalendar) ||
                        equalsAny(t, L10n.navCalendar) ||
                        startsWithAny(d, L10n.navCalendar)) &&
                    roleOf($0) == role
            }) {
                AXUIElementPerformAction(btn, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 0.8)
                if currentView(refreshWindows(conn.app)) == "calendar" { return true }
            }
        }

        // Menu bar fallback: View > Switch to > Calendar.
        var mbRef: CFTypeRef?
        if AXUIElementCopyAttributeValue(conn.app, kAXMenuBarAttribute as CFString, &mbRef) == .success,
           let menuBar = mbRef as! AXUIElement? {
            for menu in childrenOf(menuBar) where equalsAny(titleOf(menu), L10n.menuView) {
                for child in childrenOf(menu) {
                    for item in childrenOf(child) where equalsAny(titleOf(item), L10n.menuSwitchTo) {
                        for sub in childrenOf(item) {
                            for subItem in childrenOf(sub) where equalsAny(titleOf(subItem), L10n.navCalendar) {
                                AXUIElementPerformAction(subItem, kAXPressAction as CFString)
                                Thread.sleep(forTimeInterval: 0.8)
                                if currentView(refreshWindows(conn.app)) == "calendar" { return true }
                            }
                        }
                    }
                }
            }
        }

        throw OutlookAXError.navigationFailed("Could not switch Outlook to calendar view.")
    }

    /// Switch Outlook's calendar to the given view via the menu bar
    /// (Anzeigen/View > <mode>). Uses AXPressAction — no keyboard focus
    /// needed, Outlook stays in the background.
    @discardableResult
    public static func switchCalendarView(to mode: CalendarViewMode) throws -> Bool {
        guard let conn = connect() else { throw OutlookAXError.outlookNotRunning }
        let variants = localizedVariants(for: mode)
        // Fast path: already there according to the popup OR (for .list) the
        // events table is visible — whichever is true. Popup title lags the
        // actual view state, so the table check is the more reliable signal.
        if mode == .list, findEventsTable(in: conn) != nil { return true }
        if currentCalendarViewMode() == mode { return true }

        guard triggerMenu(conn.app, path: [L10n.menuView, variants]) else {
            throw OutlookAXError.viewSwitchFailed(target: mode)
        }

        // Poll for Outlook to actually render the target view. For .list
        // we watch the events AXTable (definitive indicator). For other
        // modes we fall back to the popup title.
        let deadline = Date().addingTimeInterval(3.0)
        while Date() < deadline {
            if mode == .list {
                if findEventsTable(in: conn) != nil { return true }
            } else if currentCalendarViewMode() == mode {
                return true
            }
            Thread.sleep(forTimeInterval: 0.15)
        }
        // Press was accepted — return true; caller should verify by attempting
        // the read. Throwing here would mask the common case where Outlook
        // simply rendered late.
        return true
    }

    /// Find the list-view events AXTable, if one is currently rendered.
    /// Used as a render-state oracle for .list because the popup title is
    /// stale after AX press.
    private static func findEventsTable(in conn: (app: AXUIElement, wins: [AXUIElement])) -> AXUIElement? {
        let calWin = conn.wins.first(where: {
            equalsAny(titleOf($0), L10n.calendarWindow)
        }) ?? conn.wins.first!
        return findAll(calWin, matching: {
            roleOf($0) == "AXTable" && matchesAny(descOf($0), L10n.calendarEventsTable)
        }, maxDepth: 12).first
    }

    /// Open the detail window for the event whose title+date matches the
    /// given `event`, read attendees/location/body/calendar, then close the
    /// detail window.
    ///
    /// Double-click-based: Outlook comes to front briefly (~1 s). Use for
    /// explicit user actions (e.g. "Load details" / meeting confirmation),
    /// not on every selection.
    public static func readEventDetails(for event: CalendarEvent) throws -> EventDetails {
        guard let conn = connect() else { throw OutlookAXError.outlookNotRunning }
        let calWin = conn.wins.first(where: {
            equalsAny(titleOf($0), L10n.calendarWindow)
        }) ?? conn.wins.first!

        guard let table = findAll(calWin, matching: {
            roleOf($0) == "AXTable" && matchesAny(descOf($0), L10n.calendarEventsTable)
        }, maxDepth: 12).first else {
            throw OutlookAXError.eventsTableNotFound
        }

        let targetKey = normaliseTitle(event.title)
        // Walk the rows, track current date header, match by title+date.
        var currentDate = ""
        var matchedCell: AXUIElement?
        var matchedRow: AXUIElement?
        for row in childrenOf(table) where roleOf(row) == "AXRow" {
            guard let cell = childrenOf(row).first(where: { roleOf($0) == "AXCell" }) else { continue }
            let texts = childrenOf(cell)
                .filter { roleOf($0) == "AXStaticText" }
                .compactMap { v -> String? in let s = valueOf(v); return s.isEmpty ? nil : s }
            let cellDesc = descOf(cell)
            if cellDesc.isEmpty && texts.count == 1 {
                currentDate = texts[0]
                continue
            }
            guard let raw = texts.first else { continue }
            if currentDate == event.date && normaliseTitle(stripResponsePrefix(raw)) == targetKey {
                matchedCell = cell; matchedRow = row; break
            }
        }
        guard let cell = matchedCell else {
            throw OutlookAXError.eventRowNotFound(title: event.title)
        }

        // Close any stale detail windows first so we can reliably find the
        // newly-opened one below by "non-calendar window" heuristic.
        closeAuxiliaryWindows(in: conn.app)

        // Select the row (best-effort) and double-click its centre to open.
        if let row = matchedRow {
            AXUIElementSetAttributeValue(table, "AXSelectedRows" as CFString, [row] as CFArray)
            Thread.sleep(forTimeInterval: 0.15)
        }
        guard let pos = posOf(cell), let size = sizeOf(cell) else {
            throw OutlookAXError.detailWindowNotOpened
        }
        doubleClick(at: CGPoint(x: pos.x + size.width / 2, y: pos.y + size.height / 2))
        Thread.sleep(forTimeInterval: 1.5)

        let allWins = refreshWindows(conn.app)
        guard let detailWin = allWins.first(where: {
            let t = titleOf($0)
            return !t.isEmpty && !equalsAny(t, L10n.calendarWindow)
        }) else {
            throw OutlookAXError.detailWindowNotOpened
        }
        defer { closeWindow(detailWin) }

        return parseDetailWindow(detailWin, fallbackOrganizer: event.organizer, fallbackCalendar: event.calendar)
    }

    /// Read Outlook's calendar list view into structured events.
    /// Outlook must be on the Calendar tab in list view. Call
    /// `switchCalendarView(to: .list)` first if you are not sure.
    public static func readCalendarList() throws -> [CalendarEvent] {
        guard let conn = connect() else { throw OutlookAXError.outlookNotRunning }
        if currentView(conn.wins) != "calendar" {
            throw OutlookAXError.notInCalendarView
        }
        let calWin = conn.wins.first(where: {
            equalsAny(titleOf($0), L10n.calendarWindow)
        }) ?? conn.wins.first!

        let tables = findAll(calWin, matching: {
            roleOf($0) == "AXTable" && matchesAny(descOf($0), L10n.calendarEventsTable)
        }, maxDepth: 12)

        guard let table = tables.first else {
            throw OutlookAXError.eventsTableNotFound
        }
        return parseEventsTable(table)
    }

    /// Step the calendar N days relative to today. Negative offsets go back,
    /// positive offsets go forward. `0` clicks the Today button.
    ///
    /// Works reliably across any range because it uses Outlook's Previous
    /// day / Next day arrow buttons (always visible), unlike date-picker
    /// clicks which require the day to be inside the current mini-calendar
    /// month.
    public static func navigateCalendar(byDays offset: Int) throws {
        guard let conn = connect() else { throw OutlookAXError.outlookNotRunning }
        let win = conn.wins[0]

        if offset == 0 {
            guard pressButtonAny(win, descPrefixes: L10n.today) else {
                throw OutlookAXError.navigationFailed("Today button not found.")
            }
            Thread.sleep(forTimeInterval: 0.3)
            return
        }

        let (variants, label) = offset < 0
            ? (L10n.prevDay, "Previous day")
            : (L10n.nextDay, "Next day")
        for _ in 0..<abs(offset) {
            guard pressButtonAny(win, descPrefixes: variants) else {
                throw OutlookAXError.navigationFailed("\(label) button not found.")
            }
            // Small sleep so successive presses register; Outlook's nav is
            // synchronous enough that 120 ms is fine.
            Thread.sleep(forTimeInterval: 0.12)
        }
    }

    // MARK: - Private: connection

    private static func connect() -> (app: AXUIElement, wins: [AXUIElement])? {
        let bundleIDs = ["com.microsoft.Outlook", "com.microsoft.OneOutlook"]
        guard let proc = NSWorkspace.shared.runningApplications.first(where: {
            bundleIDs.contains($0.bundleIdentifier ?? "")
        }) else { return nil }
        let app = AXUIElementCreateApplication(proc.processIdentifier)
        var ref: CFTypeRef?
        AXUIElementCopyAttributeValue(app, kAXWindowsAttribute as CFString, &ref)
        guard let wins = ref as? [AXUIElement], !wins.isEmpty else { return nil }
        return (app, wins)
    }

    private static func currentView(_ wins: [AXUIElement]) -> String {
        for w in wins {
            let t = titleOf(w)
            if equalsAny(t, L10n.calendarWindow) { return "calendar" }
            if equalsAny(t, L10n.inboxWindow) || t.contains("Outlook") { return "mail" }
        }
        return "unknown"
    }

    private static func refreshWindows(_ app: AXUIElement) -> [AXUIElement] {
        var ref: CFTypeRef?
        AXUIElementCopyAttributeValue(app, kAXWindowsAttribute as CFString, &ref)
        return ref as? [AXUIElement] ?? []
    }

    // MARK: - Private: menu bar

    /// Set by clients that want internal diagnostics piped into their own log.
    public nonisolated(unsafe) static var logger: (@Sendable (String) -> Void)? = nil

    private static func trace(_ msg: String) { logger?(msg) }

    private static func triggerMenu(_ app: AXUIElement, path: [[String]]) -> Bool {
        var mbRef: CFTypeRef?
        guard AXUIElementCopyAttributeValue(app, kAXMenuBarAttribute as CFString, &mbRef) == .success,
              let menuBar = mbRef as! AXUIElement?
        else {
            trace("triggerMenu: kAXMenuBarAttribute failed")
            return false
        }
        guard path.count >= 2 else { return false }
        let topCount = childrenOf(menuBar).count
        trace("triggerMenu: menu bar children=\(topCount) looking for top=\(path[0]) item=\(path[1])")
        for topItem in childrenOf(menuBar) {
            let title = titleOf(topItem)
            guard equalsAny(title, path[0]) else { continue }
            let submenus = childrenOf(topItem)
            trace("triggerMenu: matched top='\(title)', submenus=\(submenus.count)")
            for submenu in submenus {
                let items = childrenOf(submenu)
                trace("triggerMenu: submenu children=\(items.count)")
                for item in items {
                    if equalsAny(titleOf(item), path[1]) {
                        AXUIElementPerformAction(item, kAXPressAction as CFString)
                        trace("triggerMenu: pressed '\(titleOf(item))'")
                        return true
                    }
                }
            }
        }
        trace("triggerMenu: no match for \(path[1])")
        return false
    }

    // MARK: - Private: detail-window parser

    private static func parseDetailWindow(
        _ win: AXUIElement,
        fallbackOrganizer: String,
        fallbackCalendar: String
    ) -> EventDetails {
        let winTitle = titleOf(win)
        // "Subject • Calendar • account"
        var calendar = fallbackCalendar
        let parts = winTitle.components(separatedBy: " • ")
        if parts.count >= 2 {
            calendar = parts.dropFirst().joined(separator: " • ")
        }

        // Location: first non-empty AXTextField.
        var location = ""
        for field in findAll(win, matching: {
            roleOf($0) == "AXTextField" && !valueOf($0).isEmpty
        }, maxDepth: 10) {
            let v = valueOf(field)
            if location.isEmpty { location = v }
        }

        // Attendees: walk AXStaticText, track section + response prefix.
        var attendees: [Attendee] = []
        var organizer = fallbackOrganizer
        var section = ""
        var response = ""
        let texts = findAll(win, matching: {
            roleOf($0) == "AXStaticText" && !valueOf($0).isEmpty
        }, maxDepth: 10)
        for t in texts {
            let v = valueOf(t)
            if equalsAny(v, L10n.sectionOrganizer) { section = "organizer"; response = ""; continue }
            if equalsAny(v, L10n.sectionRequired) { section = "required"; response = ""; continue }
            if equalsAny(v, L10n.sectionOptional) { section = "optional"; response = ""; continue }
            if endsWithAny(v, L10n.respAccepted) { response = "accepted"; continue }
            if endsWithAny(v, L10n.respDeclined) { response = "declined"; continue }
            if endsWithAny(v, L10n.respNotResponded) { response = "none"; continue }
            if endsWithAny(v, L10n.respTentative) { response = "tentative"; continue }
            if v == winTitle { continue }
            let isNoise = matchesAny(v, L10n.detailNoise) || v.contains("•") || v.count > 80 || v.count < 2
            if section == "organizer" && !isNoise && organizer.isEmpty {
                organizer = v
                continue
            }
            if (section == "required" || section == "optional") && !isNoise {
                attendees.append(Attendee(name: v, type: section, response: response))
            }
        }

        // Body: AXWebArea whose description is "Reading Pane".
        var body = ""
        if let web = findElement(win, matching: {
            descOf($0) == "Reading Pane" && roleOf($0) == "AXWebArea"
        }) {
            var parts: [String] = []
            collectText(web, into: &parts)
            body = parts.joined(separator: "\n").trimmingCharacters(in: .whitespacesAndNewlines)
        }

        return EventDetails(
            attendees: attendees, location: location, body: body,
            calendar: calendar, organizer: organizer
        )
    }

    private static func closeAuxiliaryWindows(in app: AXUIElement) {
        for w in refreshWindows(app) {
            let t = titleOf(w)
            if t.isEmpty || equalsAny(t, L10n.calendarWindow) { continue }
            closeWindow(w)
        }
    }

    private static func closeWindow(_ win: AXUIElement) {
        for child in childrenOf(win) where roleOf(child) == "AXButton" {
            var sr: CFTypeRef?
            AXUIElementCopyAttributeValue(child, kAXSubroleAttribute as CFString, &sr)
            if (sr as? String) == "AXCloseButton" {
                AXUIElementPerformAction(child, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 0.25)
                return
            }
        }
    }

    private static func stripResponsePrefix(_ raw: String) -> String {
        var s = raw
        if startsWithAny(s, L10n.declinedPrefix) {
            s = String(s.drop(while: { $0 != ":" }).dropFirst(2))
        } else if startsWithAny(s, L10n.followingPrefix) {
            s = String(s.drop(while: { $0 != ":" }).dropFirst(2))
        }
        return s.trimmingCharacters(in: .whitespaces)
    }

    private static func normaliseTitle(_ s: String) -> String {
        s.lowercased().trimmingCharacters(in: .whitespacesAndNewlines)
    }

    // MARK: - Private: list-view parser

    private static func parseEventsTable(_ table: AXUIElement) -> [CalendarEvent] {
        let rows = childrenOf(table).filter { roleOf($0) == "AXRow" }
        var items: [CalendarEvent] = []
        var currentDate = ""

        for row in rows {
            let cells = childrenOf(row).filter { roleOf($0) == "AXCell" }
            guard let cell = cells.first else { continue }

            let cellDesc = descOf(cell)
            let textValues = childrenOf(cell)
                .filter { roleOf($0) == "AXStaticText" }
                .compactMap { v -> String? in
                    let val = valueOf(v); return val.isEmpty ? nil : val
                }

            // Date header row: empty desc, single text like "Samstag, 18. April"
            if cellDesc.isEmpty && textValues.count == 1 {
                currentDate = textValues[0]
                continue
            }
            guard !cellDesc.isEmpty else { continue }

            let rawTitle = textValues.first ?? ""
            var title = rawTitle
            var myResponse = "accepted"
            if startsWithAny(rawTitle, L10n.declinedPrefix) {
                myResponse = "declined"
                title = String(rawTitle.drop(while: { $0 != ":" }).dropFirst(2))
            } else if startsWithAny(rawTitle, L10n.followingPrefix) {
                myResponse = "following"
                title = String(rawTitle.drop(while: { $0 != ":" }).dropFirst(2))
            }

            var start = ""
            var end = ""
            var isAllDay = matchesAny(cellDesc, L10n.allDay)
            for val in textValues {
                guard val.contains(" - ") else { continue }
                let parts = val.components(separatedBy: " - ")
                guard parts.count == 2 else { continue }
                let left = parts[0].trimmingCharacters(in: .whitespaces)
                let right = parts[1].trimmingCharacters(in: .whitespaces)
                if left.count <= 5 && left.contains(":") {
                    start = left; end = right
                } else {
                    start = left; end = right; isAllDay = true
                }
            }

            var status = ""
            var statusRange: Range<String.Index>?
            for prefix in L10n.showAsPrefix {
                if let r = cellDesc.range(of: prefix) { statusRange = r; break }
            }
            if let range = statusRange {
                let statusStr = String(cellDesc[range.upperBound...]).trimmingCharacters(in: .whitespaces)
                if startsWithAny(statusStr, L10n.statusBusy) { status = "Busy" }
                else if startsWithAny(statusStr, L10n.statusFree) { status = "Free" }
                else if startsWithAny(statusStr, L10n.statusTentative) { status = "Tentative" }
                else if startsWithAny(statusStr, L10n.statusOOF) { status = "Out of Office" }
                else if startsWithAny(statusStr, L10n.statusElsewhere) { status = "Working Elsewhere" }
                else { status = String(statusStr.prefix(20)) }
            }

            var organizer = ""
            var orgRange: Range<String.Index>?
            for prefix in L10n.organizerPrefix {
                if let r = cellDesc.range(of: prefix) { orgRange = r; break }
            }
            if let range = orgRange {
                organizer = String(cellDesc[range.upperBound...].prefix(while: { $0 != "," }))
            } else if matchesAny(cellDesc, L10n.youAreOrganizer) {
                organizer = "You"
            }

            let categories = childrenOf(cell)
                .filter { matchesAny(descOf($0), L10n.category) }
                .flatMap { cat -> [String] in
                    childrenOf(cat).compactMap {
                        let v = valueOf($0); return v.isEmpty ? nil : v
                    }
                }
            let calendar = categories.joined(separator: ", ")

            items.append(CalendarEvent(
                title: title.trimmingCharacters(in: .whitespaces),
                date: currentDate,
                start: start,
                end: end,
                isAllDay: isAllDay,
                myResponse: myResponse,
                organizer: organizer,
                status: status,
                calendar: calendar
            ))
        }
        return items
    }

    // MARK: - Private: localisation map

    private static func localizedVariants(for mode: CalendarViewMode) -> [String] {
        switch mode {
        case .day: return L10n.viewDay
        case .workWeek: return L10n.viewWorkWeek
        case .week: return L10n.viewWeek
        case .month: return L10n.viewMonth
        case .threeDay: return L10n.viewThreeDay
        case .list: return L10n.viewList
        }
    }
}

// MARK: - AX helpers (file-private)

private func roleOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXRoleAttribute as CFString, &r)
    return r as? String ?? ""
}
private func titleOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXTitleAttribute as CFString, &r)
    return r as? String ?? ""
}
private func descOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXDescriptionAttribute as CFString, &r)
    return r as? String ?? ""
}
private func valueOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXValueAttribute as CFString, &r)
    return r as? String ?? ""
}
private func childrenOf(_ e: AXUIElement) -> [AXUIElement] {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXChildrenAttribute as CFString, &r)
    return r as? [AXUIElement] ?? []
}

private func findElement(_ root: AXUIElement, matching: (AXUIElement) -> Bool, depth: Int = 0, maxDepth: Int = 14) -> AXUIElement? {
    if matching(root) { return root }
    if depth >= maxDepth { return nil }
    for child in childrenOf(root) {
        if let found = findElement(child, matching: matching, depth: depth + 1, maxDepth: maxDepth) {
            return found
        }
    }
    return nil
}
private func findAll(_ root: AXUIElement, matching: (AXUIElement) -> Bool, depth: Int = 0, maxDepth: Int = 12) -> [AXUIElement] {
    var results: [AXUIElement] = []
    if matching(root) { results.append(root) }
    if depth >= maxDepth { return results }
    for child in childrenOf(root) {
        results += findAll(child, matching: matching, depth: depth + 1, maxDepth: maxDepth)
    }
    return results
}

private func matchesAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(where: { text.contains($0) })
}
private func startsWithAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(where: { text.hasPrefix($0) })
}
private func equalsAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(text)
}
private func endsWithAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(where: { text.hasSuffix($0) })
}

private func posOf(_ e: AXUIElement) -> CGPoint? {
    var r: CFTypeRef?
    guard AXUIElementCopyAttributeValue(e, "AXPosition" as CFString, &r) == .success else { return nil }
    var p = CGPoint.zero
    AXValueGetValue(r as! AXValue, .cgPoint, &p)
    return p
}
private func sizeOf(_ e: AXUIElement) -> CGSize? {
    var r: CFTypeRef?
    guard AXUIElementCopyAttributeValue(e, "AXSize" as CFString, &r) == .success else { return nil }
    var s = CGSize.zero
    AXValueGetValue(r as! AXValue, .cgSize, &s)
    return s
}
private func doubleClick(at point: CGPoint) {
    let src = CGEventSource(stateID: .hidSystemState)
    for clickState: Int64 in [1, 2] {
        if let down = CGEvent(mouseEventSource: src, mouseType: .leftMouseDown,
                              mouseCursorPosition: point, mouseButton: .left) {
            down.setIntegerValueField(.mouseEventClickState, value: clickState)
            down.post(tap: .cghidEventTap)
        }
        if let up = CGEvent(mouseEventSource: src, mouseType: .leftMouseUp,
                            mouseCursorPosition: point, mouseButton: .left) {
            up.setIntegerValueField(.mouseEventClickState, value: clickState)
            up.post(tap: .cghidEventTap)
        }
        if clickState == 1 { Thread.sleep(forTimeInterval: 0.05) }
    }
}

/// Recursively collect AXStaticText values + http AXLink label+URL tuples
/// into `parts`, used to flatten a Reading Pane body into plain text.
private func collectText(_ e: AXUIElement, into parts: inout [String], depth: Int = 0, maxDepth: Int = 8) {
    if depth > maxDepth { return }
    let role = roleOf(e)
    if role == "AXStaticText" {
        let v = valueOf(e); if !v.isEmpty { parts.append(v) }
    }
    if role == "AXLink" {
        let t = titleOf(e); let d = descOf(e)
        if d.hasPrefix("http") && !t.isEmpty {
            parts.append("[\(t)](\(d))"); return
        }
    }
    for c in childrenOf(e) { collectText(c, into: &parts, depth: depth + 1, maxDepth: maxDepth) }
}

private func pressButtonAny(_ win: AXUIElement, descPrefixes: [String]) -> Bool {
    if let btn = findElement(win, matching: { el in
        guard roleOf(el) == "AXButton" else { return false }
        let d = descOf(el)
        return descPrefixes.contains(where: { d.hasPrefix($0) })
    }) {
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        return true
    }
    return false
}

// MARK: - L10n (file-private, subset used by the library)

private enum L10n {
    static let calendarWindow       = ["Kalender", "Calendar", "Calendrier", "Calendario"]
    static let inboxWindow          = ["Posteingang", "Inbox", "Boîte de réception", "Bandeja de entrada"]

    static let calendarEventsTable  = ["Kalenderereignisse", "Calendar events", "Événements du calendrier", "Eventos del calendario"]
    static let allDay               = ["ganztägig", "all day", "all-day", "toute la journée", "todo el día", "tutto il giorno"]

    static let showAsPrefix         = ["anzeigen als", "show as", "afficher comme", "mostrar como"]
    static let statusBusy           = ["Gebucht", "Busy", "Occupé", "Ocupado"]
    static let statusFree           = ["Frei", "Free", "Disponible", "Libre"]
    static let statusTentative      = ["Mit Vorbehalt", "Tentative", "Provisoire", "Provisional"]
    static let statusOOF            = ["Außer Haus", "Out of Office", "Absent(e)", "Fuera de la oficina"]
    static let statusElsewhere      = ["An anderem Ort", "Working Elsewhere", "Travaille ailleurs", "Trabajando en otro lugar"]

    static let organizerPrefix      = ["Organisator ", "Organizer ", "Organisateur ", "Organizador "]
    static let youAreOrganizer      = ["Sie sind der Organisator", "You are the organizer", "Vous êtes l'organisateur", "Usted es el organizador"]

    static let category             = ["Kategorie", "Category", "Catégorie", "Categoría"]

    static let declinedPrefix       = ["Declined: ", "Abgelehnt: "]
    static let followingPrefix      = ["Following: "]

    static let menuView             = ["Anzeigen", "View", "Affichage"]
    static let menuSwitchTo         = ["Wechseln zu", "Switch to", "Basculer vers"]
    static let navCalendar          = ["Kalender", "Calendar", "Calendrier"]

    static let calendarViewPicker   = ["Kalenderansicht", "Calendar view", "Vue du calendrier"]

    // Detail-window parsing
    static let sectionOrganizer     = ["Organisator", "Organizer", "Organisateur", "Organizador"]
    static let sectionRequired      = ["Erforderlich", "Required", "Obligatoire", "Obligatorio"]
    static let sectionOptional      = ["Optional", "Facultatif", "Opcional"]
    static let respAccepted         = ["angenommen.", "accepted."]
    static let respDeclined         = ["abgesagt.", "declined."]
    static let respNotResponded     = ["nicht geantwortet.", "not responded.", "haven't responded."]
    static let respTentative        = ["mit Vorbehalt.", "tentative.", "tentatively."]
    static let detailNoise          = ["statt.", "Findet am", "Takes place", "instead.", "Tiene lugar"]

    static let today                = ["Heute", "Today", "Aujourd'hui", "Hoy"]
    static let nextDay              = ["Nächster Tag", "Next day", "Next Day", "Jour suivant", "Día siguiente"]
    static let prevDay              = ["Vorheriger Tag", "Previous day", "Previous Day", "Jour précédent", "Día anterior"]

    static let viewDay              = ["Tag", "Day", "Jour"]
    static let viewWorkWeek         = ["Arbeitswoche", "Work Week", "Semaine de travail"]
    static let viewWeek             = ["Woche", "Week", "Semaine"]
    static let viewMonth            = ["Monat", "Month", "Mois"]
    static let viewThreeDay         = ["Drei Tage", "Three Day", "Trois jours"]
    static let viewList             = ["Liste", "List"]
}
