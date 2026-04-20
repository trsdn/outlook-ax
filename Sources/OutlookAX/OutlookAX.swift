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

    public enum OutlookAXError: LocalizedError {
        case outlookNotRunning
        case notInCalendarView
        case viewSwitchFailed(target: CalendarViewMode)
        case eventsTableNotFound
        case navigationFailed(String)

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

    /// Switch Outlook's calendar to the given view via the menu bar
    /// (Anzeigen/View > <mode>). Uses AXPressAction — no keyboard focus
    /// needed, Outlook stays in the background.
    @discardableResult
    public static func switchCalendarView(to mode: CalendarViewMode) throws -> Bool {
        guard let conn = connect() else { throw OutlookAXError.outlookNotRunning }
        let variants = localizedVariants(for: mode)
        if triggerMenu(conn.app, path: [L10n.menuView, variants]) {
            // Give Outlook a beat to re-render the AX tree.
            Thread.sleep(forTimeInterval: 0.4)
            return true
        }
        throw OutlookAXError.viewSwitchFailed(target: mode)
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

    /// Navigate Outlook's calendar to the given date via the mini calendar.
    /// Accepts Outlook's localised date format; callers must pass the label
    /// that Outlook uses (e.g. `"20. April"` in German).
    @discardableResult
    public static func navigate(toDateLabel label: String) throws -> Bool {
        guard let conn = connect() else { throw OutlookAXError.outlookNotRunning }
        // The mini calendar renders dates as AXButtons with their day label
        // as the title. Find one whose title contains the requested label.
        for win in conn.wins {
            if let btn = findElement(win, matching: {
                roleOf($0) == "AXButton" && titleOf($0).contains(label)
            }) {
                AXUIElementPerformAction(btn, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 0.3)
                return true
            }
        }
        throw OutlookAXError.navigationFailed("Date label '\(label)' not found in mini calendar.")
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

    // MARK: - Private: menu bar

    private static func triggerMenu(_ app: AXUIElement, path: [[String]]) -> Bool {
        var mbRef: CFTypeRef?
        guard AXUIElementCopyAttributeValue(app, kAXMenuBarAttribute as CFString, &mbRef) == .success,
              let menuBar = mbRef as! AXUIElement?
        else { return false }
        guard path.count >= 2 else { return false }
        for topItem in childrenOf(menuBar) {
            if equalsAny(titleOf(topItem), path[0]) {
                for submenu in childrenOf(topItem) {
                    for item in childrenOf(submenu) {
                        if equalsAny(titleOf(item), path[1]) {
                            AXUIElementPerformAction(item, kAXPressAction as CFString)
                            return true
                        }
                    }
                }
            }
        }
        return false
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

    static let viewDay              = ["Tag", "Day", "Jour"]
    static let viewWorkWeek         = ["Arbeitswoche", "Work Week", "Semaine de travail"]
    static let viewWeek             = ["Woche", "Week", "Semaine"]
    static let viewMonth            = ["Monat", "Month", "Mois"]
    static let viewThreeDay         = ["Drei Tage", "Three Day", "Trois jours"]
    static let viewList             = ["Liste", "List"]
}
