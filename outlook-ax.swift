#!/usr/bin/env swift
import AppKit
import Foundation

// MARK: - AX Helpers

func roleOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?; AXUIElementCopyAttributeValue(e, kAXRoleAttribute as CFString, &r); return r as? String ?? ""
}
func titleOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?; AXUIElementCopyAttributeValue(e, kAXTitleAttribute as CFString, &r); return r as? String ?? ""
}
func descOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?; AXUIElementCopyAttributeValue(e, kAXDescriptionAttribute as CFString, &r); return r as? String ?? ""
}
func valueOf(_ e: AXUIElement) -> String {
    var r: CFTypeRef?; AXUIElementCopyAttributeValue(e, kAXValueAttribute as CFString, &r); return r as? String ?? ""
}
func childrenOf(_ e: AXUIElement) -> [AXUIElement] {
    var r: CFTypeRef?; AXUIElementCopyAttributeValue(e, kAXChildrenAttribute as CFString, &r); return r as? [AXUIElement] ?? []
}

func findElement(_ root: AXUIElement, matching: (AXUIElement) -> Bool, depth: Int = 0, maxDepth: Int = 12) -> AXUIElement? {
    if matching(root) { return root }
    if depth >= maxDepth { return nil }
    for child in childrenOf(root) {
        if let found = findElement(child, matching: matching, depth: depth + 1, maxDepth: maxDepth) { return found }
    }
    return nil
}

func findAll(_ root: AXUIElement, matching: (AXUIElement) -> Bool, depth: Int = 0, maxDepth: Int = 10) -> [AXUIElement] {
    var results: [AXUIElement] = []
    if matching(root) { results.append(root) }
    if depth >= maxDepth { return results }
    for child in childrenOf(root) {
        results += findAll(child, matching: matching, depth: depth + 1, maxDepth: maxDepth)
    }
    return results
}

func collectText(_ elem: AXUIElement, into parts: inout [String], depth: Int = 0, maxDepth: Int = 8) {
    if depth > maxDepth { return }
    let role = roleOf(elem)
    if role == "AXStaticText" {
        let val = valueOf(elem)
        if !val.isEmpty { parts.append(val) }
    }
    if role == "AXLink" {
        let title = titleOf(elem)
        let desc = descOf(elem)
        if desc.hasPrefix("http") && !title.isEmpty {
            parts.append("[\(title)](\(desc))")
            return
        }
    }
    for child in childrenOf(elem) {
        collectText(child, into: &parts, depth: depth + 1, maxDepth: maxDepth)
    }
}

/// Press a toolbar button by matching its desc prefix
func pressButton(_ win: AXUIElement, descPrefix: String) -> Bool {
    if let btn = findElement(win, matching: {
        descOf($0).hasPrefix(descPrefix) && roleOf($0) == "AXButton"
    }) {
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        return true
    }
    return false
}

/// Press a toolbar button by matching its title
func pressButtonByTitle(_ win: AXUIElement, title: String) -> Bool {
    if let btn = findElement(win, matching: {
        titleOf($0) == title && roleOf($0) == "AXButton"
    }) {
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        return true
    }
    return false
}

// MARK: - Keyboard Simulation (for date/time fields and text input)

func typeText(_ text: String) {
    for char in text {
        let str = String(char)
        let src = CGEventSource(stateID: .hidSystemState)
        if let keyDown = CGEvent(keyboardEventSource: src, virtualKey: 0, keyDown: true) {
            let utf16 = Array(str.utf16)
            keyDown.keyboardSetUnicodeString(stringLength: utf16.count, unicodeString: utf16)
            keyDown.post(tap: .cghidEventTap)
        }
        if let keyUp = CGEvent(keyboardEventSource: src, virtualKey: 0, keyDown: false) {
            keyUp.post(tap: .cghidEventTap)
        }
        Thread.sleep(forTimeInterval: 0.03)
    }
}

func pressKey(_ keyCode: CGKeyCode, flags: CGEventFlags = []) {
    let src = CGEventSource(stateID: .hidSystemState)
    if let down = CGEvent(keyboardEventSource: src, virtualKey: keyCode, keyDown: true) {
        down.flags = flags
        down.post(tap: .cghidEventTap)
    }
    if let up = CGEvent(keyboardEventSource: src, virtualKey: keyCode, keyDown: false) {
        up.flags = flags
        up.post(tap: .cghidEventTap)
    }
    Thread.sleep(forTimeInterval: 0.05)
}

func pressTab() { pressKey(48) }
func pressReturn() { pressKey(36) }
func pressEscape() { pressKey(53) }
func selectAll() { pressKey(0, flags: .maskCommand) }

func posOf(_ e: AXUIElement) -> CGPoint? {
    var r: CFTypeRef?
    guard AXUIElementCopyAttributeValue(e, "AXPosition" as CFString, &r) == .success else { return nil }
    var pt = CGPoint.zero; AXValueGetValue(r as! AXValue, .cgPoint, &pt); return pt
}

func sizeOf(_ e: AXUIElement) -> CGSize? {
    var r: CFTypeRef?
    guard AXUIElementCopyAttributeValue(e, "AXSize" as CFString, &r) == .success else { return nil }
    var sz = CGSize.zero; AXValueGetValue(r as! AXValue, .cgSize, &sz); return sz
}

func doubleClick(at point: CGPoint) {
    let src = CGEventSource(stateID: .hidSystemState)
    for clickState: Int64 in [1, 2] {
        if let down = CGEvent(mouseEventSource: src, mouseType: .leftMouseDown, mouseCursorPosition: point, mouseButton: .left) {
            down.setIntegerValueField(.mouseEventClickState, value: clickState)
            down.post(tap: .cghidEventTap)
        }
        if let up = CGEvent(mouseEventSource: src, mouseType: .leftMouseUp, mouseCursorPosition: point, mouseButton: .left) {
            up.setIntegerValueField(.mouseEventClickState, value: clickState)
            up.post(tap: .cghidEventTap)
        }
        if clickState == 1 { Thread.sleep(forTimeInterval: 0.05) }
    }
}

/// Get all action names for an element (standard + Electron named actions)
func actionNamesOf(_ e: AXUIElement) -> [String] {
    var names: CFArray?
    guard AXUIElementCopyActionNames(e, &names) == .success, let arr = names as? [String] else { return [] }
    return arr
}

/// Perform a named action (e.g. "Name:Termin anzeigen\n...") by prefix match
func performNamedAction(_ e: AXUIElement, prefix: String) -> Bool {
    for action in actionNamesOf(e) {
        if action.contains(prefix) {
            return AXUIElementPerformAction(e, action as CFString) == .success
        }
    }
    return false
}

/// Get focused element from the application
func focusedElement(_ app: AXUIElement) -> AXUIElement? {
    var ref: CFTypeRef?
    guard AXUIElementCopyAttributeValue(app, kAXFocusedUIElementAttribute as CFString, &ref) == .success else { return nil }
    return (ref as! AXUIElement)
}

// MARK: - L10n (Localized Label Matching)
// AX labels come from Outlook's web layer and vary by UI language.
// Instead of detecting language, we match against all known variants.
// To add a new language: append its label to the relevant array.

/// Check if a string matches any of the given localized variants
func matchesAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(where: { text.contains($0) })
}

/// Check if a string starts with any of the given localized variants
func startsWithAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(where: { text.hasPrefix($0) })
}

/// Check if a string equals any of the given localized variants
func equalsAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(text)
}

/// Check if a string ends with any of the given localized variants
func endsWithAny(_ text: String, _ variants: [String]) -> Bool {
    variants.contains(where: { text.hasSuffix($0) })
}

/// Try to find an element matching any of the desc prefix variants
func pressButtonAny(_ win: AXUIElement, descPrefixes: [String]) -> Bool {
    for prefix in descPrefixes {
        if pressButton(win, descPrefix: prefix) { return true }
    }
    return false
}

// Localized AX label sets — each array contains [de, en, fr, es, it, ...] variants
// Add new languages by appending to these arrays.
struct L10n {
    // -- Window titles --
    static let calendarWindow       = ["Kalender", "Calendar", "Calendrier", "Calendario"]
    static let inboxWindow          = ["Posteingang", "Inbox", "Boîte de réception", "Bandeja de entrada"]

    // -- Calendar: table / list view --
    static let calendarEventsTable  = ["Kalenderereignisse", "Calendar events", "Événements du calendrier", "Eventos del calendario"]
    static let allDay               = ["ganztägig", "all day", "all-day", "toute la journée", "todo el día", "tutto il giorno"]

    // -- Calendar: status (cellDesc suffix after "anzeigen als" / "show as") --
    static let showAsPrefix         = ["anzeigen als", "show as", "afficher comme", "mostrar como"]
    static let statusBusy           = ["Gebucht", "Busy", "Occupé", "Ocupado"]
    static let statusFree           = ["Frei", "Free", "Disponible", "Libre"]
    static let statusTentative      = ["Mit Vorbehalt", "Tentative", "Provisoire", "Provisional"]
    static let statusOOF            = ["Außer Haus", "Out of Office", "Absent(e)", "Fuera de la oficina"]
    static let statusElsewhere      = ["An anderem Ort", "Working Elsewhere", "Travaille ailleurs", "Trabajando en otro lugar"]

    // -- Calendar: organizer (cellDesc) --
    static let organizerPrefix      = ["Organisator ", "Organizer ", "Organisateur ", "Organizador "]
    static let youAreOrganizer      = ["Sie sind der Organisator", "You are the organizer", "Vous êtes l'organisateur", "Usted es el organizador"]

    // -- Calendar: categories --
    static let category             = ["Kategorie", "Category", "Catégorie", "Categoría"]

    // -- Calendar: detail window sections --
    static let sectionOrganizer     = ["Organisator", "Organizer", "Organisateur", "Organizador"]
    static let sectionRequired      = ["Erforderlich", "Required", "Obligatoire", "Obligatorio"]
    static let sectionOptional      = ["Optional", "Facultatif", "Opcional"]

    // -- Calendar: attendee response counts (suffixes) --
    static let respAccepted         = ["angenommen.", "accepted."]
    static let respDeclined         = ["abgesagt.", "declined."]
    static let respNotResponded     = ["nicht geantwortet.", "not responded.", "haven't responded."]
    static let respTentative        = ["mit Vorbehalt.", "tentative.", "tentatively."]

    // -- Calendar: detail window noise filters --
    static let detailNoise          = ["statt.", "Findet am", "Takes place", "instead.", "Tiene lugar"]

    // -- Calendar: title prefixes for myResponse --
    static let declinedPrefix       = ["Declined: ", "Abgelehnt: "]
    static let followingPrefix      = ["Following: "]

    // -- Calendar: navigation buttons --
    static let today                = ["Heute", "Today", "Aujourd'hui", "Hoy"]
    static let nextDay              = ["Nächster Tag", "Next day", "Next Day", "Jour suivant", "Día siguiente"]
    static let prevDay              = ["Vorheriger Tag", "Previous day", "Previous Day", "Jour précédent", "Día anterior"]

    // -- Calendar: create event --
    static let newEvent             = ["Neuer Termin", "New event", "New Event", "Nouvel événement", "Nuevo evento"]
    static let subject              = ["Betreff", "Subject", "Objet", "Asunto"]
    static let addAttendees         = ["Erforderliche Personen hinzufügen", "Erforderliche Teilnehmer",
                                       "Add required attendees", "Required attendees",
                                       "Ajouter des participants obligatoires"]
    static let startDate            = ["Beginnt am", "Starts on", "Start date", "Commence le"]
    static let startTime            = ["Beginnt um", "Starts at", "Start time", "Commence à"]
    static let send                 = ["Senden", "Send", "Envoyer", "Enviar"]
    static let save                 = ["Speichern", "Save", "Enregistrer", "Guardar"]
    static let discard              = ["Verwerfen", "Discard", "Ignorer", "Descartar"]

    // -- Calendar: view picker --
    static let calendarViewPicker   = ["Kalenderansicht", "Calendar view", "Vue du calendrier"]

    // -- Calendar: sidebar / nav pane --
    static let navPane              = ["Navigationsbereich", "Navigation pane", "Volet de navigation"]
    static let myCalendars          = ["Meine Kalender", "My Calendars", "Mes calendriers"]
    static let otherCalendars       = ["Andere Kalender", "Other Calendars", "Autres calendriers"]
    static let calendarShown        = ["Angezeigt", "Shown", "Displayed", "Teilweise angezeigt", "Partially shown"]

    // -- Mail: header --
    static let messageHeader        = ["Nachrichtenkopfzeile", "Message header", "En-tête du message"]
    static let headerDetails        = ["Nachrichtenkopfdetails", "Message header details"]
    static let sentPrefix           = ["Gesendet am:", "Gesendet", "Sent on:", "Sent"]
    static let messageList          = ["Nachrichtenliste", "Message list", "Liste de messages"]

    // -- Mail: compose --
    static let newEmail             = ["Neue E-Mail", "New Email", "New email", "Nouveau courrier"]
    static let composeWindow        = ["Nachricht", "Message", "Neue", "New"]
    static let toField              = ["An", "To", "À", "Empfänger", "Recipients"]
    static let bodyField            = ["Nachrichtentext", "Message body"]
    static let fromPrefix           = ["Von: ", "From: ", "De : "]

    // -- Mail: search --
    static let search               = ["Suchen", "Search", "Rechercher", "Buscar"]

    // -- Mail: actions --
    static let reply                = ["Antworten", "Reply", "Répondre"]
    static let replyAll             = ["Allen antworten", "Reply All", "Répondre à tous"]
    static let forward              = ["Weiterleiten", "Forward", "Transférer"]
    static let delete               = ["Löschen", "Delete", "Supprimer"]
    static let archive              = ["Archivieren", "Archive", "Archiver"]
    static let flag                 = ["Kennzeichnen", "Flag", "Marquer"]
    static let markRead             = ["Markieren:", "Mark:"]
    static let move                 = ["Verschieben", "Move", "Déplacer"]
    static let report               = ["Melden", "Report", "Signaler"]
    static let react                = ["Reagieren", "React", "Réagir"]
    static let summarize            = ["Zusammenfassen", "Summarize", "Résumer"]
    static let filterSort           = ["Filtern und sortieren", "Filter and sort", "Filtrer et trier"]
    static let moreItems            = ["Weitere Elemente anzeigen", "Show more items"]

    // -- Navigation --
    static let navCalendar          = ["Kalender", "Calendar", "Calendrier"]
    static let navMail              = ["E-Mail", "Mail", "Courrier"]
    static let navPeople            = ["Personen", "People", "Contacts"]
    static let navTasks             = ["Aufgaben", "Tasks", "Tâches"]
    static let navCopilot           = ["Copilot"]
    static let navOneDrive          = ["OneDrive"]
    static let navFavorites         = ["Favoriten", "Favorites", "Favoris"]
    static let navOrgExplorer       = ["Organisations-Explorer", "Org Explorer"]

    // -- Notifications --
    static let newNotifications     = ["Neue Benachrichtigungen:", "New notifications:", "Nouvelles notifications:"]

    // -- My Day --
    static let myDay                = ["Mein Tag", "My Day", "Ma journée"]

    // -- Menu paths --
    static let menuView             = ["Anzeigen", "View", "Affichage"]
    static let menuSwitchTo         = ["Wechseln zu", "Switch to", "Basculer vers"]
    static let menuEvent            = ["Ereignis", "Event", "Événement"]
    static let menuTools            = ["Werkzeuge", "Tools", "Outils"]
    static let menuSync             = ["Synchronisieren", "Sync", "Synchroniser"]
    static let menuAutoReply        = ["Automatische Antworten...", "Automatic Replies...", "Réponses automatiques..."]
    static let menuShowAs           = ["Anzeigen als", "Show As"]
    static let menuCategorize       = ["Kategorisieren", "Categorize", "Catégoriser"]
    static let menuPrivate          = ["Privat", "Private", "Privé"]

    // -- Calendar: view modes --
    static let viewDay              = ["Tag", "Day", "Jour"]
    static let viewWorkWeek         = ["Arbeitswoche", "Work Week", "Semaine de travail"]
    static let viewWeek             = ["Woche", "Week", "Semaine"]
    static let viewMonth            = ["Monat", "Month", "Mois"]
    static let viewThreeDay         = ["Drei Tage", "Three Day", "Trois jours"]
    static let viewList             = ["Liste", "List", "Liste"]

    // -- Calendar: timescale --
    static let minutesSuffix        = ["Minuten", "Minutes", "minutes"]

    // -- Calendar: filter --
    static let filterAll            = ["Alle", "All", "Tous"]
    static let filterAppointments   = ["Termine", "Appointments", "Rendez-vous"]
    static let filterMeetings       = ["Besprechungen", "Meetings", "Réunions"]
    static let filterCategories     = ["Kategorien", "Categories", "Catégories"]
    static let filterShowAs         = ["Anzeigen als", "Show As"]
    static let filterRecurring      = ["Wiederholung", "Recurring", "Récurrence"]
    static let filterPrivacy        = ["Datenschutz", "Privacy", "Confidentialité"]
    static let filterDeclined       = ["Abgelehnte Ereignisse ausblenden", "Hide declined events"]

    // -- Calendar: event menu actions --
    static let accept               = ["Akzeptieren", "Zusagen", "Accept", "Accepter"]
    static let tentative            = ["Mit Vorbehalt", "Tentative", "Provisoire"]
    static let decline              = ["Ablehnen", "Decline", "Refuser"]
    static let joinMeeting          = ["An Onlinebesprechung teilnehmen", "Join Online Meeting", "Rejoindre la réunion"]
    static let duplicateEvent       = ["Ereignis duplizieren", "Duplicate Event", "Dupliquer l'événement"]
    static let cancelMeeting        = ["Besprechung absagen", "Cancel Meeting", "Annuler la réunion"]

    // -- Calendar: show-as menu values --
    static let showAsFree           = ["Frei", "Free"]
    static let showAsTentative      = ["Mit Vorbehalt", "Tentative"]
    static let showAsBusy           = ["Gebucht", "Busy"]
    static let showAsOOF            = ["Außer Haus", "Out of Office"]
    static let showAsElsewhere      = ["An anderem Ort tätig", "Working Elsewhere"]

    // -- Calendar: color menu --
    static let colorBlue            = ["Blau", "Blue", "Bleu", "Azul"]
    static let colorGreen           = ["Grün", "Green", "Vert", "Verde"]
    static let colorOrange          = ["Orange", "Naranja"]
    static let colorPlatinum        = ["Platingrau", "Platinum", "Platine", "Platino"]
    static let colorYellow          = ["Gelb", "Yellow", "Jaune", "Amarillo"]
    static let colorCyan            = ["Zyan", "Cyan"]
    static let colorMagenta         = ["Magenta"]
    static let colorBrown           = ["Braun", "Brown", "Marron", "Marrón"]
    static let colorBurgundy        = ["Burgunderrot", "Burgundy", "Bordeaux", "Burdeos"]
    static let colorTeal            = ["Meeresgrün", "Teal", "Sarcelle"]
    static let colorLilac           = ["Flieder", "Lilac", "Lilas", "Lila"]

    // -- Calendar: timescale menu --
    static let menuTimescale        = ["Zeitskala", "Timescale", "Échelle de temps"]
    static let menuFilter           = ["Filtern", "Filter", "Filtrer"]
    static let menuColor            = ["Farbe", "Color", "Couleur"]

    // -- Mail: folder sidebar --
    static let favorites            = ["Favoriten", "Favorites"]
    static let allAccounts          = ["Alle Konten", "All Accounts"]
    static let groups               = ["Gruppen", "Groups"]
}

// MARK: - Outlook Connection

/// Launch Outlook if not running, unminimize if minimized, bring to front.
func ensureOutlookReady() {
    let bundleID = "com.microsoft.Outlook"
    let isRunning = NSWorkspace.shared.runningApplications.contains(where: { $0.bundleIdentifier == bundleID })

    if !isRunning {
        // Launch via NSWorkspace
        let config = NSWorkspace.OpenConfiguration()
        config.activates = true
        if let url = NSWorkspace.shared.urlForApplication(withBundleIdentifier: bundleID) {
            let sema = DispatchSemaphore(value: 0)
            NSWorkspace.shared.openApplication(at: url, configuration: config) { _, _ in sema.signal() }
            _ = sema.wait(timeout: .now() + 10)
            // Wait for UI to be ready
            Thread.sleep(forTimeInterval: 3.0)
        }
    }

    // Unminimize and bring to front via AppleScript
    let script = """
    tell application "Microsoft Outlook"
        activate
        set miniaturized of every window to false
    end tell
    """
    if let appleScript = NSAppleScript(source: script) {
        var error: NSDictionary?
        appleScript.executeAndReturnError(&error)
    }
    // Give Outlook time to restore windows
    if !isRunning { Thread.sleep(forTimeInterval: 2.0) }
    else { Thread.sleep(forTimeInterval: 0.5) }
}

func connectOutlook() -> (app: AXUIElement, wins: [AXUIElement])? {
    // First ensure Outlook is running and visible
    ensureOutlookReady()

    guard let proc = NSWorkspace.shared.runningApplications.first(where: {
        $0.bundleIdentifier == "com.microsoft.Outlook"
    }) else { return nil }
    let axApp = AXUIElementCreateApplication(proc.processIdentifier)
    var ref: CFTypeRef?
    AXUIElementCopyAttributeValue(axApp, kAXWindowsAttribute as CFString, &ref)
    guard let wins = ref as? [AXUIElement], !wins.isEmpty else { return nil }
    return (axApp, wins)
}

func refreshWindows(_ app: AXUIElement) -> [AXUIElement] {
    var ref: CFTypeRef?
    AXUIElementCopyAttributeValue(app, kAXWindowsAttribute as CFString, &ref)
    return ref as? [AXUIElement] ?? []
}

func currentView(_ wins: [AXUIElement]) -> String {
    for w in wins {
        let t = titleOf(w)
        if equalsAny(t, L10n.calendarWindow) { return "calendar" }
        if equalsAny(t, L10n.inboxWindow) || t.contains("Outlook") { return "mail" }
        if equalsAny(t, L10n.navPeople) { return "people" }
    }
    return "unknown"
}

// MARK: - JSON Output

let jsonFlag = CommandLine.arguments.contains("--json")
let detailsFlag = CommandLine.arguments.contains("--details")

func printJSON<T: Encodable>(_ value: T) {
    let encoder = JSONEncoder()
    encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
    if let data = try? encoder.encode(value), let str = String(data: data, encoding: .utf8) {
        print(str)
    }
}

func fail(_ msg: String) -> Never {
    if jsonFlag {
        printJSON(["error": msg])
    } else {
        fputs("error: \(msg)\n", stderr)
    }
    exit(1)
}

func ok(_ action: String, extra: [String: String] = [:]) {
    if jsonFlag {
        var dict = ["status": "ok", "action": action]
        for (k, v) in extra { dict[k] = v }
        printJSON(dict)
    } else {
        print("\(action)")
        for (k, v) in extra { print("  \(k): \(v)") }
    }
}

// MARK: - JSON Types

struct StatusJSON: Codable {
    let running: Bool
    let window: String
    let view: String
    let notifications: Int
}

struct EmailJSON: Codable {
    let subject: String
    let from: String
    let to: String
    let date: String
    let body: String
}

struct InboxItemJSON: Codable {
    let subject: String
    let from: String
    let date: String
    let preview: String
}

struct AttendeeJSON: Codable {
    let name: String
    let type: String      // "required", "optional"
    let response: String  // "accepted", "declined", "none", "tentative", ""
}

struct CalendarEventJSON: Codable {
    let title: String
    let date: String
    let start: String
    let end: String
    let location: String
    let isAllDay: Bool
    let myResponse: String   // "accepted", "declined", "following", ""
    let status: String       // "Busy", "Free", "Tentative", "Out of Office", "Working Elsewhere"
    let calendar: String     // calendar name + account from detail window
    let attendees: [AttendeeJSON]  // from detail window (--details flag)
    let organizer: String    // from detail window (--details flag)
    let body: String         // from detail window (--details flag)
}

struct FolderJSON: Codable {
    let name: String
    let account: String
}

struct CalendarInfoJSON: Codable {
    let name: String
    let visible: Bool
    let group: String
}

// MARK: - Commands: Status

func cmdStatus() {
    if let conn = connectOutlook() {
        let view = currentView(conn.wins)
        let winTitle = titleOf(conn.wins[0])
        var notifCount = 0
        if let notifBtn = findElement(conn.wins[0], matching: {
            startsWithAny(descOf($0), L10n.newNotifications) && roleOf($0) == "AXButton"
        }) {
            let desc = descOf(notifBtn)
            let parts = desc.components(separatedBy: ": ")
            if parts.count >= 2 { notifCount = Int(parts[1]) ?? 0 }
        }
        if jsonFlag {
            printJSON(StatusJSON(running: true, window: winTitle, view: view, notifications: notifCount))
        } else {
            print("Outlook: running")
            print("Window:  \(winTitle)")
            print("View:    \(view)")
            if notifCount > 0 { print("Notifications: \(notifCount)") }
        }
    } else {
        if jsonFlag {
            printJSON(StatusJSON(running: false, window: "", view: "", notifications: 0))
        } else {
            print("Outlook: not running")
        }
    }
}

// MARK: - Commands: Notifications

func cmdNotifications() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if let notifBtn = findElement(conn.wins[0], matching: {
        startsWithAny(descOf($0), L10n.newNotifications) && roleOf($0) == "AXButton"
    }) {
        let desc = descOf(notifBtn)
        let parts = desc.components(separatedBy: ": ")
        let count = parts.count >= 2 ? (Int(parts[1]) ?? 0) : 0
        if jsonFlag {
            printJSON(["count": "\(count)"])
        } else {
            print("Notifications: \(count)")
        }
    } else {
        if jsonFlag {
            printJSON(["count": "0"])
        } else {
            print("Notifications: 0")
        }
    }
}

// MARK: - Commands: Mail

func cmdMailCurrent() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    guard let headerGroup = findElement(win, matching: { equalsAny(titleOf($0), L10n.messageHeader) }) else {
        fail("No email open (message header not found)")
    }

    var subject = "", from = "", date = "", recipients = ""

    for child in childrenOf(headerGroup) {
        if roleOf(child) == "AXStaticText" && subject.isEmpty {
            subject = valueOf(child)
        }
        if equalsAny(descOf(child), L10n.headerDetails) {
            for detail in childrenOf(child) {
                let desc = descOf(detail)
                if desc == "messageHeaderFromContent" {
                    from = valueOf(detail)
                    for p in L10n.fromPrefix { if from.hasPrefix(p) { from = String(from.dropFirst(p.count)); break } }
                }
                if desc == "messageHeaderRecipientsContent" {
                    recipients = valueOf(detail)
                }
                let t = titleOf(detail)
                if startsWithAny(t, L10n.sentPrefix) {
                    date = valueOf(detail)
                    if date.isEmpty { date = t }
                }
            }
        }
    }

    var bodyParts: [String] = []
    if let webArea = findElement(win, matching: { descOf($0) == "Reading Pane" }) {
        collectText(webArea, into: &bodyParts)
    }
    let body = bodyParts.filter { line in
        let l = line.lowercased()
        return !l.contains("unsubscribe") && !l.contains("subscribe") && line != "|" && line.count > 1
    }.joined(separator: "\n\n")

    if jsonFlag {
        printJSON(EmailJSON(subject: subject, from: from, to: recipients, date: date, body: body))
    } else {
        print("Subject: \(subject)")
        print("From:    \(from)")
        if !recipients.isEmpty { print("To:      \(recipients)") }
        print("Date:    \(date)")
        print("---")
        print(body)
    }
}

func cmdMailInbox(limit: Int) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    // Find message list table
    var items: [InboxItemJSON] = []

    if let table = findElement(win, matching: { roleOf($0) == "AXTable" && equalsAny(descOf($0), L10n.messageList) }) {
        let rows = childrenOf(table)
        for row in rows.prefix(limit) {
            let texts = findAll(row, matching: { roleOf($0) == "AXStaticText" }, maxDepth: 4)
            let values = texts.map { valueOf($0) }.filter { !$0.isEmpty }
            if values.isEmpty { continue }
            let subject = values.count > 2 ? values[2] : values[0]
            let from = values.count > 0 ? values[0] : ""
            let date = values.count > 1 ? values[1] : ""
            let preview = values.count > 3 ? values[3] : ""
            items.append(InboxItemJSON(subject: subject, from: from, date: date, preview: preview))
        }
    }

    if items.isEmpty {
        // Fallback: try AXCell or AXRow
        let rows = findAll(win, matching: {
            (roleOf($0) == "AXCell" && matchesAny(descOf($0), L10n.composeWindow)) || roleOf($0) == "AXRow"
        }, maxDepth: 8)
        for row in rows.prefix(limit) {
            let texts = findAll(row, matching: { roleOf($0) == "AXStaticText" }, maxDepth: 4)
            let values = texts.map { valueOf($0) }.filter { !$0.isEmpty }
            if values.isEmpty { continue }
            let subject = values.count > 2 ? values[2] : values[0]
            let from = values.count > 0 ? values[0] : ""
            let date = values.count > 1 ? values[1] : ""
            let preview = values.count > 3 ? values[3] : ""
            items.append(InboxItemJSON(subject: subject, from: from, date: date, preview: preview))
        }
    }

    if jsonFlag {
        printJSON(items)
    } else {
        if items.isEmpty { print("No messages found"); return }
        for (i, item) in items.enumerated() {
            print("\(i+1). \(item.subject)")
            print("   From: \(item.from)  Date: \(item.date)")
            if !item.preview.isEmpty { print("   \(item.preview.prefix(80))") }
        }
    }
}

func cmdMailSearch(query: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    guard let searchField = findElement(win, matching: {
        equalsAny(descOf($0), L10n.search) && roleOf($0) == "AXTextField"
    }) else { fail("Search field not found") }

    AXUIElementSetAttributeValue(searchField, kAXFocusedAttribute as CFString, true as CFTypeRef)
    Thread.sleep(forTimeInterval: 0.3)
    AXUIElementSetAttributeValue(searchField, kAXValueAttribute as CFString, query as CFTypeRef)
    Thread.sleep(forTimeInterval: 0.3)
    pressReturn()
    Thread.sleep(forTimeInterval: 2.0)

    ok("Search executed", extra: ["query": query])
}

func cmdMailReply() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    let found = L10n.reply.contains(where: { pressButtonByTitle(win, title: $0) || pressButton(win, descPrefix: $0) })
    if found {
        Thread.sleep(forTimeInterval: 1.0)
        ok("Reply opened")
    } else {
        fail("Reply button not found — is an email selected?")
    }
}

func cmdMailForward() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    // Forward might be in "Reagieren" submenu or directly available
    if pressButtonAny(win, descPrefixes: L10n.forward) {
        Thread.sleep(forTimeInterval: 1.0)
        ok("Forward opened")
    } else {
        // Try via "Show more items" first, then forward
        if pressButtonAny(win, descPrefixes: L10n.moreItems) {
            Thread.sleep(forTimeInterval: 0.5)
            if pressButtonAny(win, descPrefixes: L10n.forward) {
                Thread.sleep(forTimeInterval: 1.0)
                ok("Forward opened")
                return
            }
        }
        fail("Forward button not found — is an email selected?")
    }
}

func cmdMailDelete() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if pressButtonAny(conn.wins[0], descPrefixes: L10n.delete) {
        Thread.sleep(forTimeInterval: 0.5)
        ok("Email deleted")
    } else {
        fail("Delete button not found")
    }
}

func cmdMailArchive() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if pressButtonAny(conn.wins[0], descPrefixes: L10n.archive) {
        Thread.sleep(forTimeInterval: 0.5)
        ok("Email archived")
    } else {
        fail("Archive button not found")
    }
}

func cmdMailCompose(to: String?, subject: String?, body: String?) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    guard pressButtonAny(win, descPrefixes: L10n.newEmail) else {
        fail("'New Email' button not found — are you in mail view?")
    }
    Thread.sleep(forTimeInterval: 2.5)

    // Find the compose window
    let allWins = refreshWindows(conn.app)
    guard let composeWin = allWins.first(where: {
        matchesAny(titleOf($0), L10n.composeWindow)
    }) ?? allWins.first else {
        fail("Compose window not found")
    }
    AXUIElementPerformAction(composeWin, kAXRaiseAction as CFString)
    Thread.sleep(forTimeInterval: 0.5)

    // Fill To field
    if let toAddr = to {
        if let toField = findElement(composeWin, matching: {
            equalsAny(descOf($0), L10n.toField) && roleOf($0) == "AXTextField"
        }) {
            AXUIElementSetAttributeValue(toField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            typeText(toAddr)
            Thread.sleep(forTimeInterval: 0.5)
            pressTab()
            Thread.sleep(forTimeInterval: 0.5)
        }
    }

    // Fill Subject
    if let subj = subject {
        if let subjField = findElement(composeWin, matching: {
            equalsAny(descOf($0), L10n.subject) && roleOf($0) == "AXTextField"
        }) {
            AXUIElementSetAttributeValue(subjField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            AXUIElementSetAttributeValue(subjField, kAXValueAttribute as CFString, subj as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
        }
    }

    // Fill Body — type into the web area / body field
    if let bodyText = body {
        if let bodyArea = findElement(composeWin, matching: {
            equalsAny(descOf($0), L10n.bodyField) || roleOf($0) == "AXWebArea"
        }) {
            AXUIElementSetAttributeValue(bodyArea, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            typeText(bodyText)
            Thread.sleep(forTimeInterval: 0.3)
        }
    }

    ok("Compose window opened", extra: [
        "to": to ?? "",
        "subject": subject ?? ""
    ])
}

func cmdMailFolders() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    var folders: [FolderJSON] = []
    var currentAccount = ""

    if let outline = findElement(win, matching: { roleOf($0) == "AXOutline" }) {
        let rows = findAll(outline, matching: { roleOf($0) == "AXRow" }, maxDepth: 3)
        for row in rows {
            // Check for account header (has DisclosureTriangle)
            let triangles = findAll(row, matching: { roleOf($0) == "AXDisclosureTriangle" }, maxDepth: 3)
            let texts = findAll(row, matching: { roleOf($0) == "AXStaticText" }, maxDepth: 3)
            let checkboxes = findAll(row, matching: { roleOf($0) == "AXCheckBox" }, maxDepth: 3)

            if !triangles.isEmpty {
                // Account or category header
                for cb in checkboxes {
                    let d = descOf(cb)
                    if d.contains("@") { currentAccount = d.components(separatedBy: ";").first?.trimmingCharacters(in: .whitespaces) ?? d }
                }
                for t in texts {
                    let v = valueOf(t)
                    if v.contains("@") { currentAccount = v }
                }
            }

            // Folder entries are StaticText children
            for t in texts {
                let v = valueOf(t)
                if !v.isEmpty && !v.contains("@") && !equalsAny(v, L10n.favorites + L10n.allAccounts + L10n.groups) {
                    folders.append(FolderJSON(name: v, account: currentAccount))
                }
            }

            // Also check checkbox values (calendar sidebar uses these)
            for cb in checkboxes {
                let v = valueOf(cb)
                let d = descOf(cb)
                if !v.isEmpty && !d.contains("@") && triangles.isEmpty {
                    // Skip if already added via text
                    if !folders.contains(where: { $0.name == v }) {
                        folders.append(FolderJSON(name: v, account: currentAccount))
                    }
                }
            }
        }
    }

    if jsonFlag {
        printJSON(folders)
    } else {
        if folders.isEmpty { print("No folders found"); return }
        var lastAccount = ""
        for f in folders {
            if f.account != lastAccount && !f.account.isEmpty {
                print("\n[\(f.account)]")
                lastAccount = f.account
            }
            print("  \(f.name)")
        }
    }
}

func cmdMailFolder(name: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    if let outline = findElement(win, matching: { roleOf($0) == "AXOutline" }) {
        // Find a StaticText or Row with matching name
        if let target = findElement(outline, matching: {
            let v = valueOf($0)
            let role = roleOf($0)
            return role == "AXStaticText" && v.lowercased() == name.lowercased()
        }) {
            AXUIElementPerformAction(target, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 1.0)
            ok("Switched to folder", extra: ["folder": name])
            return
        }
        // Try clicking the row that contains the folder name
        let rows = findAll(outline, matching: { roleOf($0) == "AXRow" }, maxDepth: 3)
        for row in rows {
            let texts = findAll(row, matching: { roleOf($0) == "AXStaticText" && valueOf($0).lowercased() == name.lowercased() }, maxDepth: 3)
            if !texts.isEmpty {
                AXUIElementPerformAction(row, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 1.0)
                ok("Switched to folder", extra: ["folder": name])
                return
            }
        }
    }
    fail("Folder '\(name)' not found")
}

// MARK: - Commands: Calendar

func cmdCalendarToday() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let view = currentView(conn.wins)
    if view != "calendar" {
        fail("Not in calendar view. Run: outlook-ax navigate calendar")
    }

    let calWin = conn.wins.first(where: {
        equalsAny(titleOf($0), L10n.calendarWindow)
    }) ?? conn.wins.first!

    // List view: find AXTable with desc containing calendar events label
    let tables = findAll(calWin, matching: {
        roleOf($0) == "AXTable" && matchesAny(descOf($0), L10n.calendarEventsTable)
    }, maxDepth: 10)

    guard let table = tables.first else {
        fail("Calendar events table not found. Make sure you are in list view (View > List).")
    }

    // Parse rows: date headers vs event cells
    let rows = childrenOf(table).filter { roleOf($0) == "AXRow" }

    var items: [CalendarEventJSON] = []
    var eventCells: [AXUIElement] = []    // parallel to items — for --details double-click
    var eventRows: [AXUIElement?] = []    // parallel to items — for row selection
    var currentDate = ""

    for row in rows {
        let cells = childrenOf(row).filter { roleOf($0) == "AXCell" }
        guard let cell = cells.first else { continue }

        let cellDesc = descOf(cell)
        let textChildren = childrenOf(cell).filter { roleOf($0) == "AXStaticText" }
        let textValues = textChildren.compactMap { v -> String? in
            let val = valueOf(v); return val.isEmpty ? nil : val
        }

        // Date header row: single AXStaticText like "Samstag, 18. April"
        if cellDesc.isEmpty && textValues.count == 1 {
            currentDate = textValues[0]
            continue
        }

        // Event row: cellDesc contains structured info
        guard !cellDesc.isEmpty else { continue }

        // Extract title from first AXStaticText child, parse myResponse from prefix
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

        // Extract time from AXStaticText like "16:00 - 20:00" or "2.4.2026 - 28.4.2026"
        var start = ""
        var end = ""
        var isAllDay = false

        // Check for all-day event in cellDesc
        if matchesAny(cellDesc, L10n.allDay) {
            isAllDay = true
        }

        // Find time text (HH:MM - HH:MM pattern)
        for val in textValues {
            if val.contains(" - ") {
                let parts = val.components(separatedBy: " - ")
                if parts.count == 2 {
                    let left = parts[0].trimmingCharacters(in: .whitespaces)
                    let right = parts[1].trimmingCharacters(in: .whitespaces)
                    // Check if it's a time range (HH:MM) or date range
                    if left.count <= 5 && left.contains(":") {
                        start = left
                        end = right
                    } else {
                        // Date range for multi-day events
                        start = left
                        end = right
                        isAllDay = true
                    }
                }
            }
        }

        // Extract status from cellDesc: "anzeigen als<Status>" / "show as<Status>"
        var status = ""
        var statusRange: Range<String.Index>? = nil
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

        // Extract organizer from cellDesc
        var organizer = ""
        var orgRange: Range<String.Index>? = nil
        for prefix in L10n.organizerPrefix {
            if let r = cellDesc.range(of: prefix) { orgRange = r; break }
        }
        if let range = orgRange {
            organizer = String(cellDesc[range.upperBound...].prefix(while: { $0 != "," }))
        } else if matchesAny(cellDesc, L10n.youAreOrganizer) {
            organizer = "You"
        }

        // Extract categories from AXUnknown children with "Kategorie"/"Category" in desc
        let catElements = childrenOf(cell).filter {
            matchesAny(descOf($0), L10n.category)
        }
        var categories: [String] = []
        for cat in catElements {
            // Category text is in child AXStaticText
            for child in childrenOf(cat) {
                let v = valueOf(child)
                if !v.isEmpty { categories.append(v) }
            }
        }

        // Build calendar info from categories or leave empty (list view doesn't show calendar name)
        let calendar = categories.joined(separator: ", ")

        items.append(CalendarEventJSON(
            title: title.trimmingCharacters(in: .whitespaces),
            date: currentDate,
            start: start, end: end,
            location: "", isAllDay: isAllDay,
            myResponse: myResponse,
            status: status, calendar: calendar,
            attendees: [], organizer: organizer, body: ""
        ))
        eventCells.append(cell)
        eventRows.append(row)
    }

    // --details: open each event's detail window via double-click to read attendees, location, body
    if detailsFlag && !items.isEmpty {
        let axApp = conn.app

        // Close any stale detail windows first
        for w in refreshWindows(axApp) {
            let t = titleOf(w)
            if !t.isEmpty && !equalsAny(t, L10n.calendarWindow) {
                for child in childrenOf(w) {
                    if roleOf(child) == "AXButton" {
                        var sr: CFTypeRef?
                        AXUIElementCopyAttributeValue(child, kAXSubroleAttribute as CFString, &sr)
                        if (sr as? String) == "AXCloseButton" {
                            AXUIElementPerformAction(child, kAXPressAction as CFString)
                            Thread.sleep(forTimeInterval: 0.3)
                            break
                        }
                    }
                }
            }
        }

        for (idx, cell) in eventCells.enumerated() {
            guard let pos = posOf(cell), let sz = sizeOf(cell) else { continue }
            let clickPt = CGPoint(x: pos.x + sz.width / 2, y: pos.y + sz.height / 2)

            // Select row first
            if let parentRow = eventRows[idx] {
                AXUIElementSetAttributeValue(table, "AXSelectedRows" as CFString, [parentRow] as CFArray)
                Thread.sleep(forTimeInterval: 0.2)
            }

            doubleClick(at: clickPt)
            Thread.sleep(forTimeInterval: 1.5)

            // Find the detail window
            let allWins = refreshWindows(axApp)
            guard let detailWin = allWins.first(where: {
                let t = titleOf($0)
                return !t.isEmpty && !equalsAny(t, L10n.calendarWindow)
            }) else { continue }

            let winTitle = titleOf(detailWin)

            // Read calendar from window title: "Subject • Calendar • account"
            var detailCalendar = items[idx].calendar
            let titleParts = winTitle.components(separatedBy: " • ")
            if titleParts.count >= 2 {
                detailCalendar = titleParts.dropFirst().joined(separator: " • ")
            }

            // Read all static texts
            let texts = findAll(detailWin, matching: {
                roleOf($0) == "AXStaticText" && !valueOf($0).isEmpty
            }, maxDepth: 10)

            var location = ""
            var attendees: [AttendeeJSON] = []
            var body = ""
            var currentSection = "" // "Organisator", "Erforderlich", "Optional"

            // Read location from text field
            let locFields = findAll(detailWin, matching: {
                roleOf($0) == "AXTextField" && !valueOf($0).isEmpty
            }, maxDepth: 10)
            for lf in locFields {
                let v = valueOf(lf)
                if v.contains("Teams") || v.contains("Meeting") || v.contains("Room") || v.contains("Raum") || !v.isEmpty {
                    if location.isEmpty { location = v }
                }
            }

            // Parse texts for attendees with response status
            var currentResponse = ""
            for t in texts {
                let val = valueOf(t)
                if equalsAny(val, L10n.sectionOrganizer) { currentSection = "organizer"; currentResponse = ""; continue }
                if equalsAny(val, L10n.sectionRequired) { currentSection = "required"; currentResponse = ""; continue }
                if equalsAny(val, L10n.sectionOptional) { currentSection = "optional"; currentResponse = ""; continue }
                // Response status counts: "1 angenommen.", "3 declined.", etc.
                if endsWithAny(val, L10n.respAccepted) { currentResponse = "accepted"; continue }
                if endsWithAny(val, L10n.respDeclined) { currentResponse = "declined"; continue }
                if endsWithAny(val, L10n.respNotResponded) { currentResponse = "none"; continue }
                if endsWithAny(val, L10n.respTentative) { currentResponse = "tentative"; continue }
                // Skip the window title text
                if val == winTitle { continue }
                // Attendee names appear after section headers
                if (currentSection == "required" || currentSection == "optional") &&
                   !val.contains("•") && !matchesAny(val, L10n.detailNoise) &&
                   val.count < 80 && val.count > 1 {
                    attendees.append(AttendeeJSON(
                        name: val,
                        type: currentSection,
                        response: currentResponse
                    ))
                }
            }

            // Read body from WebArea
            if let webArea = findElement(detailWin, matching: {
                descOf($0) == "Reading Pane" && roleOf($0) == "AXWebArea"
            }) {
                var bodyParts: [String] = []
                collectText(webArea, into: &bodyParts)
                body = bodyParts.joined(separator: "\n").trimmingCharacters(in: .whitespacesAndNewlines)
            }

            // Update item
            items[idx] = CalendarEventJSON(
                title: items[idx].title,
                date: items[idx].date,
                start: items[idx].start, end: items[idx].end,
                location: location, isAllDay: items[idx].isAllDay,
                myResponse: items[idx].myResponse,
                status: items[idx].status,
                calendar: detailCalendar.isEmpty ? items[idx].calendar : detailCalendar,
                attendees: attendees,
                organizer: items[idx].organizer,
                body: body
            )

            // Close detail window
            var closed = false
            for child in childrenOf(detailWin) {
                if roleOf(child) == "AXButton" {
                    var sr: CFTypeRef?
                    AXUIElementCopyAttributeValue(child, kAXSubroleAttribute as CFString, &sr)
                    if (sr as? String) == "AXCloseButton" {
                        AXUIElementPerformAction(child, kAXPressAction as CFString)
                        closed = true
                        break
                    }
                }
            }
            if !closed { pressKey(13, flags: .maskCommand) }
            Thread.sleep(forTimeInterval: 0.5)
        }
    }

    // Sort: all-day first, then by start time
    items.sort { a, b in
        if a.isAllDay && !b.isAllDay { return true }
        if !a.isAllDay && b.isAllDay { return false }
        return a.start < b.start
    }

    if jsonFlag {
        printJSON(items)
    } else {
        if items.isEmpty { print("No events found"); return }
        var lastDate = ""
        for item in items {
            // Print date header when it changes
            if item.date != lastDate && !item.date.isEmpty {
                if !lastDate.isEmpty { print("") }
                print("── \(item.date) ──")
                lastDate = item.date
            }
            var line = ""
            if item.isAllDay {
                line += "[all day] "
            } else if !item.start.isEmpty {
                line += "\(item.start) - \(item.end)  "
            }
            line += item.title
            if item.myResponse == "declined" { line += " ✗" }
            else if item.myResponse == "following" { line += " 👁" }
            if !item.status.isEmpty { line += " [\(item.status)]" }
            if !item.calendar.isEmpty { line += " (\(item.calendar))" }
            if !item.organizer.isEmpty && item.organizer != "Sie" {
                line += " org: \(item.organizer)"
            }
            print(line)
            if detailsFlag {
                if !item.location.isEmpty { print("  Location: \(item.location)") }
                if !item.attendees.isEmpty {
                    let formatted = item.attendees.map { a -> String in
                        var s = a.type == "optional" ? "(opt) \(a.name)" : a.name
                        if !a.response.isEmpty { s += " [\(a.response)]" }
                        return s
                    }
                    print("  Attendees: \(formatted.joined(separator: ", "))")
                }
                if !item.body.isEmpty { print("  Body: \(String(item.body.prefix(200)))") }
            }
        }
    }
}

func cmdCalendarCreate(subject: String, attendee: String?, date: String?, time: String?, send: Bool) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let axApp = conn.app

    // Step 0: Discard existing draft
    for w in conn.wins {
        let t = titleOf(w)
        if matchesAny(t, ["Termin", "Event", "Événement"]) {
            if let btn = findElement(w, matching: { equalsAny(descOf($0), L10n.discard) && roleOf($0) == "AXButton" }) {
                AXUIElementPerformAction(btn, kAXPressAction as CFString)
                Thread.sleep(forTimeInterval: 1.5)
            }
        }
    }

    // Step 1: Find calendar window
    let wins = refreshWindows(axApp)
    guard let calWin = wins.first(where: { equalsAny(titleOf($0), L10n.calendarWindow) }) else {
        fail("Calendar window not found. Run: outlook-ax navigate calendar")
    }

    // Step 2: Navigate to date in mini calendar if provided
    if let dateStr = date {
        // Try to find the date button in the mini calendar
        // Date buttons have desc like "Montag, 20. April"
        let _ = findElement(calWin, matching: { descOf($0).contains(dateStr) })
    }

    // Step 3: Open new event
    guard let newBtn = findElement(calWin, matching: {
        startsWithAny(descOf($0), L10n.newEvent) && roleOf($0) == "AXButton" && !descOf($0).isEmpty
    }) else {
        fail("'New event' button not found")
    }
    AXUIElementPerformAction(newBtn, kAXPressAction as CFString)
    Thread.sleep(forTimeInterval: 2.5)

    // Step 4: Find event window
    let allWins = refreshWindows(axApp)
    guard let evWin = allWins.first(where: { matchesAny(titleOf($0), ["Termin", "Event", "Événement"]) }) else {
        fail("Event form window not found")
    }
    AXUIElementPerformAction(evWin, kAXRaiseAction as CFString)
    Thread.sleep(forTimeInterval: 0.5)

    // Step 5: Fill subject
    if let subj = findElement(evWin, matching: { equalsAny(descOf($0), L10n.subject) && roleOf($0) == "AXTextField" }) {
        AXUIElementSetAttributeValue(subj, kAXFocusedAttribute as CFString, true as CFTypeRef)
        Thread.sleep(forTimeInterval: 0.3)
        AXUIElementSetAttributeValue(subj, kAXValueAttribute as CFString, subject as CFTypeRef)
        Thread.sleep(forTimeInterval: 0.3)
    }

    // Step 6: Fill attendee
    if let attendeeEmail = attendee {
        if let att = findElement(evWin, matching: {
            equalsAny(descOf($0), L10n.addAttendees) && roleOf($0) == "AXTextField"
        }) {
            AXUIElementSetAttributeValue(att, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.5)
            typeText(attendeeEmail)
            Thread.sleep(forTimeInterval: 1.0)
            pressTab()
            Thread.sleep(forTimeInterval: 1.0)
        }
    }

    // Step 7: Set date (keyboard simulation — AXDateTimeArea ignores AXValue)
    if let dateStr = date {
        if let dateField = findElement(evWin, matching: { equalsAny(titleOf($0), L10n.startDate) && roleOf($0) == "AXDateTimeArea" }) {
            AXUIElementSetAttributeValue(dateField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            selectAll()
            Thread.sleep(forTimeInterval: 0.1)
            typeText(dateStr)
            Thread.sleep(forTimeInterval: 0.5)
            pressTab()
            Thread.sleep(forTimeInterval: 0.3)
        }
    }

    // Step 8: Set time
    if let timeStr = time {
        if let timeField = findElement(evWin, matching: { equalsAny(titleOf($0), L10n.startTime) && roleOf($0) == "AXDateTimeArea" }) {
            AXUIElementSetAttributeValue(timeField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            selectAll()
            Thread.sleep(forTimeInterval: 0.1)
            typeText(timeStr)
            Thread.sleep(forTimeInterval: 0.5)
            pressTab()
            Thread.sleep(forTimeInterval: 0.3)
        }
    }

    // Step 9: Send or leave open
    if send {
        if let sendBtn = findElement(evWin, matching: { equalsAny(descOf($0), L10n.send) && roleOf($0) == "AXButton" }) {
            AXUIElementPerformAction(sendBtn, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 1.0)
            ok("Event sent", extra: ["subject": subject])
        } else if let saveBtn = findElement(evWin, matching: { equalsAny(descOf($0), L10n.save) && roleOf($0) == "AXButton" }) {
            AXUIElementPerformAction(saveBtn, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 1.0)
            ok("Event saved (no attendee = no Send button)", extra: ["subject": subject])
        }
    } else {
        ok("Event form filled — NOT sent", extra: [
            "subject": subject,
            "attendee": attendee ?? "",
            "date": date ?? "",
            "time": time ?? ""
        ])
    }
}

func cmdCalendarView(mode: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }

    // Map mode to localized label variants
    let variants: [String]
    switch mode.lowercased() {
    case "day", "tag", "jour": variants = L10n.viewDay
    case "workweek", "arbeitswoche": variants = L10n.viewWorkWeek
    case "week", "woche", "semaine": variants = L10n.viewWeek
    case "month", "monat", "mois": variants = L10n.viewMonth
    case "threeday", "3day", "drei": variants = L10n.viewThreeDay
    case "list", "liste": variants = L10n.viewList
    default: variants = [mode]
    }

    // Prefer menu bar (View > <mode>): AXPress works without keyboard focus,
    // so the switch is reliable even when Outlook is in the background.
    if triggerMenuL10n(conn.app, path: [L10n.menuView, variants]) {
        Thread.sleep(forTimeInterval: 0.4)
        ok("Calendar view changed", extra: ["mode": mode])
        return
    }

    // Fallback: the calendar view popup + typed text (needs Outlook focused).
    let win = conn.wins[0]
    if let viewPopup = findElement(win, matching: {
        startsWithAny(descOf($0), L10n.calendarViewPicker) && roleOf($0) == "AXPopUpButton"
    }) {
        AXUIElementPerformAction(viewPopup, kAXPressAction as CFString)
        Thread.sleep(forTimeInterval: 0.5)
        typeText(variants[0])
        Thread.sleep(forTimeInterval: 0.3)
        pressReturn()
        Thread.sleep(forTimeInterval: 0.5)
        ok("Calendar view changed", extra: ["mode": mode, "via": "popup"])
    } else {
        fail("Calendar view picker not found — are you in calendar view?")
    }
}

func cmdCalendarNavigate(direction: String?, date: String?) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    if let dir = direction {
        switch dir.lowercased() {
        case "today", "heute":
            if pressButtonAny(win, descPrefixes: L10n.today) {
                Thread.sleep(forTimeInterval: 0.5)
                ok("Navigated to today")
            } else {
                fail("'Today' button not found")
            }
        case "next", "forward":
            if pressButtonAny(win, descPrefixes: L10n.nextDay) {
                Thread.sleep(forTimeInterval: 0.3)
                ok("Navigated to next day")
            } else {
                fail("'Next day' button not found")
            }
        case "prev", "previous", "back":
            if pressButtonAny(win, descPrefixes: L10n.prevDay) {
                Thread.sleep(forTimeInterval: 0.3)
                ok("Navigated to previous day")
            } else {
                fail("'Previous day' button not found")
            }
        default:
            fail("Unknown direction: \(dir). Use: today, next, prev")
        }
    } else if let dateStr = date {
        // Try clicking in mini calendar — buttons have desc like "Montag, 20. April"
        // Parse the date to find the right button
        if let dateBtn = findElement(win, matching: {
            let d = descOf($0)
            return roleOf($0) == "AXButton" && d.contains(dateStr)
        }) {
            AXUIElementPerformAction(dateBtn, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 0.5)
            ok("Navigated to date", extra: ["date": dateStr])
        } else {
            // Try navigating months first, then find the date
            // For now, provide a hint
            fail("Date '\(dateStr)' not found in mini calendar. Try a partial match like '20. April'")
        }
    } else {
        fail("Specify --direction or --date")
    }
}

func cmdCalendarCalendars() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    var calendars: [CalendarInfoJSON] = []
    var currentGroup = ""

    if let outline = findElement(win, matching: { roleOf($0) == "AXOutline" && equalsAny(descOf($0), L10n.navPane) }) {
        let rows = findAll(outline, matching: { roleOf($0) == "AXRow" }, maxDepth: 3)
        for row in rows {
            let checkboxes = findAll(row, matching: { roleOf($0) == "AXCheckBox" }, maxDepth: 3)
            let triangles = findAll(row, matching: { roleOf($0) == "AXDisclosureTriangle" }, maxDepth: 3)

            for cb in checkboxes {
                let d = descOf(cb)
                let v = valueOf(cb)

                // Group headers have DisclosureTriangles
                if !triangles.isEmpty && (matchesAny(d, L10n.myCalendars) || matchesAny(d, L10n.otherCalendars) ||
                    d.contains("@")) {
                    currentGroup = v
                    continue
                }

                if !v.isEmpty {
                    let visible = matchesAny(d, L10n.calendarShown)
                    calendars.append(CalendarInfoJSON(name: v, visible: visible, group: currentGroup))
                }
            }
        }
    }

    if jsonFlag {
        printJSON(calendars)
    } else {
        if calendars.isEmpty { print("No calendars found"); return }
        var lastGroup = ""
        for cal in calendars {
            if cal.group != lastGroup && !cal.group.isEmpty {
                print("\n[\(cal.group)]")
                lastGroup = cal.group
            }
            let marker = cal.visible ? "[x]" : "[ ]"
            print("  \(marker) \(cal.name)")
        }
    }
}

func cmdCalendarToggle(name: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    if let outline = findElement(win, matching: { roleOf($0) == "AXOutline" && equalsAny(descOf($0), L10n.navPane) }) {
        if let cb = findElement(outline, matching: {
            roleOf($0) == "AXCheckBox" && valueOf($0).lowercased() == name.lowercased()
        }) {
            AXUIElementPerformAction(cb, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 0.5)
            let newDesc = descOf(cb)
            let nowVisible = matchesAny(newDesc, L10n.calendarShown)
            ok("Calendar toggled", extra: ["name": name, "visible": nowVisible ? "true" : "false"])
            return
        }
    }
    fail("Calendar '\(name)' not found")
}

// MARK: - Commands: Navigate

func cmdNavigate(to target: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]

    // Map target to the L10n variants used in Outlook
    let buttonVariants: [String]
    switch target.lowercased() {
    case "calendar", "kalender":
        buttonVariants = L10n.navCalendar
    case "mail", "email", "e-mail", "post":
        buttonVariants = L10n.navMail
    case "people", "personen", "contacts":
        buttonVariants = L10n.navPeople
    case "todo", "tasks", "aufgaben":
        buttonVariants = L10n.navTasks
    case "copilot":
        buttonVariants = L10n.navCopilot
    case "onedrive":
        buttonVariants = L10n.navOneDrive
    case "favorites", "favoriten":
        buttonVariants = L10n.navFavorites
    case "org-explorer", "org", "organisations-explorer":
        buttonVariants = L10n.navOrgExplorer
    default:
        fail("Unknown target: \(target). Use: calendar, mail, people, todo, copilot, onedrive, favorites, org-explorer")
    }

    // Try multiple element types — Outlook uses different roles depending on context
    let roles = ["AXRadioButton", "AXButton", "AXTab", "AXMenuItem"]
    for role in roles {
        if let navBtn = findElement(win, matching: {
            let d = descOf($0); let t = titleOf($0)
            return (equalsAny(d, buttonVariants) || equalsAny(t, buttonVariants) || startsWithAny(d, buttonVariants)) &&
            roleOf($0) == role
        }) {
            AXUIElementPerformAction(navBtn, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 1.5)
            ok("Navigated to \(target)")
            return
        }
    }

    // Last resort: try via Menubar > View > Switch to
    var mbRef: CFTypeRef?
    let mbResult = AXUIElementCopyAttributeValue(conn.app, kAXMenuBarAttribute as CFString, &mbRef)
    if mbResult.rawValue == 0 {
        let mb = mbRef as! AXUIElement
        for menu in childrenOf(mb) {
            if equalsAny(titleOf(menu), L10n.menuView) {
                for child in childrenOf(menu) {
                    for item in childrenOf(child) {
                        if equalsAny(titleOf(item), L10n.menuSwitchTo) {
                            for sub in childrenOf(item) {
                                for subItem in childrenOf(sub) {
                                    if equalsAny(titleOf(subItem), buttonVariants) {
                                        AXUIElementPerformAction(subItem, kAXPressAction as CFString)
                                        Thread.sleep(forTimeInterval: 1.5)
                                        ok("Navigated to \(target) (via menu)")
                                        return
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    fail("Navigation to '\(target)' failed — button not found in UI or menu")
}

// MARK: - Menu Bar Helper

/// Trigger a menu item via the menu bar. Each path element is an array of L10n variants.
/// Example: [L10n.menuEvent, L10n.accept]
func triggerMenuL10n(_ app: AXUIElement, path: [[String]]) -> Bool {
    var mbRef: CFTypeRef?
    guard AXUIElementCopyAttributeValue(app, kAXMenuBarAttribute as CFString, &mbRef).rawValue == 0 else { return false }
    let mb = mbRef as! AXUIElement

    guard path.count >= 2 else { return false }
    for topItem in childrenOf(mb) {
        if equalsAny(titleOf(topItem), path[0]) {
            for submenu in childrenOf(topItem) {
                if path.count == 2 {
                    for item in childrenOf(submenu) {
                        if equalsAny(titleOf(item), path[1]) {
                            AXUIElementPerformAction(item, kAXPressAction as CFString)
                            return true
                        }
                    }
                } else if path.count == 3 {
                    for item in childrenOf(submenu) {
                        if equalsAny(titleOf(item), path[1]) {
                            for sub in childrenOf(item) {
                                for subItem in childrenOf(sub) {
                                    if equalsAny(titleOf(subItem), path[2]) {
                                        AXUIElementPerformAction(subItem, kAXPressAction as CFString)
                                        return true
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    return false
}

/// Legacy trigger with exact string paths (kept for backwards compat with user-provided strings)
func triggerMenu(_ app: AXUIElement, path: [String]) -> Bool {
    triggerMenuL10n(app, path: path.map { [$0] })
}

// MARK: - Mail Commands (new batch)

func cmdMailReplyAll() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if let btn = findElement(win, matching: {
        let d = descOf($0); let t = titleOf($0)
        return (equalsAny(d, L10n.replyAll) || equalsAny(t, L10n.replyAll)) && roleOf($0) == "AXButton"
    }) {
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        ok("Reply-all opened")
    } else {
        fail("Reply-all button not found — is an email open?")
    }
}

func cmdMailFlag() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if pressButtonAny(win, descPrefixes: L10n.flag) {
        ok("Flag toggled")
    } else {
        fail("Flag button not found — is an email selected?")
    }
}

func cmdMailReadUnread() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if let btn = findElement(win, matching: { startsWithAny(descOf($0), L10n.markRead) && roleOf($0) == "AXButton" }) {
        let current = descOf(btn)
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        // Detect whether we toggled to read or unread based on the label containing "unread"/"ungelesen"
        let wasUnread = current.lowercased().contains("ungelesen") || current.lowercased().contains("unread")
        if wasUnread {
            ok("Marked as unread")
        } else {
            ok("Marked as read")
        }
    } else {
        fail("Read/unread button not found — is an email selected?")
    }
}

func cmdMailMove() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if pressButtonAny(win, descPrefixes: L10n.move) {
        ok("Move dialog opened — select a folder")
    } else {
        fail("Move button not found — is an email selected?")
    }
}

func cmdMailReport() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if pressButtonAny(win, descPrefixes: L10n.report) {
        ok("Report dialog opened")
    } else {
        fail("Report button not found — is an email selected?")
    }
}

func cmdMailReact() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if pressButtonAny(win, descPrefixes: L10n.react) {
        ok("React picker opened")
    } else {
        fail("React button not found — is an email selected?")
    }
}

func cmdMailSummarize() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if pressButtonAny(win, descPrefixes: L10n.summarize) {
        ok("Copilot summarize triggered")
    } else {
        fail("Summarize button not found — Copilot may not be available")
    }
}

func cmdMailFilter() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if let popup = findElement(win, matching: { equalsAny(descOf($0), L10n.filterSort) && roleOf($0) == "AXPopUpButton" }) {
        AXUIElementPerformAction(popup, kAXShowMenuAction as CFString)
        ok("Filter/sort popup opened")
    } else {
        fail("Filter button not found — are you in mail view?")
    }
}

// MARK: - Calendar Commands (new batch)

func cmdCalendarTimescale(minutes: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    // Try localized "X Minutes" / "X Minuten" etc.
    for suffix in L10n.minutesSuffix {
        if triggerMenuL10n(conn.app, path: [L10n.menuView, ["\(minutes) \(suffix)"]]) {
            ok("Timescale set to \(minutes) minutes")
            return
        }
    }
    fail("Timescale '\(minutes)' not available — are you in calendar view?")
}

func cmdCalendarFilter(filter: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let variants: [String]
    switch filter.lowercased() {
    case "all", "alle": variants = L10n.filterAll
    case "appointments", "termine": variants = L10n.filterAppointments
    case "meetings", "besprechungen": variants = L10n.filterMeetings
    case "categories", "kategorien": variants = L10n.filterCategories
    case "showas", "anzeigen": variants = L10n.filterShowAs
    case "recurring", "wiederholung": variants = L10n.filterRecurring
    case "privacy", "datenschutz": variants = L10n.filterPrivacy
    case "declined", "abgelehnte": variants = L10n.filterDeclined
    default: variants = [filter]
    }
    if triggerMenuL10n(conn.app, path: [L10n.menuView, ["Filtern", "Filter", "Filtrer"], variants]) {
        ok("Calendar filter set to '\(filter)'")
    } else {
        fail("Calendar filter '\(filter)' not available — are you in calendar view?")
    }
}

func cmdCalendarColor(color: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let variants: [String]
    switch color.lowercased() {
    case "blue", "blau": variants = L10n.colorBlue
    case "green", "grün", "gruen": variants = L10n.colorGreen
    case "orange": variants = L10n.colorOrange
    case "platinum", "platin", "platingrau": variants = L10n.colorPlatinum
    case "yellow", "gelb": variants = L10n.colorYellow
    case "cyan", "zyan": variants = L10n.colorCyan
    case "magenta": variants = L10n.colorMagenta
    case "brown", "braun": variants = L10n.colorBrown
    case "burgundy", "burgunderrot": variants = L10n.colorBurgundy
    case "teal", "meeresgrün", "meeresgruen": variants = L10n.colorTeal
    case "lilac", "flieder": variants = L10n.colorLilac
    default: variants = [color]
    }
    if triggerMenuL10n(conn.app, path: [L10n.menuView, L10n.menuColor, variants]) {
        ok("Calendar color set to '\(color)'")
    } else {
        fail("Calendar color '\(color)' not available — are you in calendar view?")
    }
}

func cmdCalendarAccept() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.accept]) {
        ok("Meeting accepted")
    } else {
        fail("Accept not available — is a meeting event selected?")
    }
}

func cmdCalendarTentative() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.tentative]) {
        ok("Meeting tentatively accepted")
    } else {
        fail("Tentative not available — is a meeting event selected?")
    }
}

func cmdCalendarDecline() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.decline]) {
        ok("Meeting declined")
    } else {
        fail("Decline not available — is a meeting event selected?")
    }
}

func cmdCalendarJoin() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.joinMeeting]) {
        ok("Joining online meeting")
    } else {
        fail("Join not available — is an online meeting event selected?")
    }
}

func cmdCalendarDuplicate() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.duplicateEvent]) {
        ok("Event duplicated")
    } else {
        fail("Duplicate not available — is an event selected?")
    }
}

func cmdCalendarCategorize(category: String?) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if let cat = category {
        if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.menuCategorize, [cat]]) {
            ok("Category '\(cat)' applied")
        } else {
            fail("Category '\(cat)' not found")
        }
    } else {
        if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.menuCategorize]) {
            ok("Categorize menu opened")
        } else {
            fail("Categorize not available — is an event selected?")
        }
    }
}

func cmdCalendarPrivate() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.menuPrivate]) {
        ok("Private flag toggled")
    } else {
        fail("Private not available — is an event open for editing?")
    }
}

func cmdCalendarShowAs(status: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let variants: [String]
    switch status.lowercased() {
    case "free", "frei": variants = L10n.showAsFree
    case "tentative", "vorbehalt": variants = L10n.showAsTentative
    case "busy", "gebucht": variants = L10n.showAsBusy
    case "oof", "ooo", "away", "abwesend": variants = L10n.showAsOOF
    case "elsewhere", "woanders": variants = L10n.showAsElsewhere
    default: variants = [status]
    }
    if triggerMenuL10n(conn.app, path: [L10n.menuEvent, L10n.menuShowAs, variants]) {
        ok("Show-as set to '\(status)'")
    } else {
        fail("Show-as '\(status)' not available — is an event open for editing?")
    }
}

// MARK: - System Commands

func cmdSync() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuTools, L10n.menuSync]) {
        ok("Sync triggered")
    } else {
        fail("Sync failed")
    }
}

func cmdAutoReply() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    if triggerMenuL10n(conn.app, path: [L10n.menuTools, L10n.menuAutoReply]) {
        ok("Auto-reply settings opened")
    } else {
        fail("Auto-reply dialog failed to open")
    }
}

func cmdMyDay() {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    if let btn = findElement(win, matching: { startsWithAny(descOf($0), L10n.myDay) && roleOf($0) == "AXButton" }) {
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        ok("My Day panel toggled")
    } else {
        fail("My Day button not found")
    }
}

func cmdAccount(name: String) {
    guard let conn = connectOutlook() else { fail("Outlook not running") }
    let win = conn.wins[0]
    // Try toolbar account buttons first
    if let btn = findElement(win, matching: {
        (descOf($0).lowercased().contains(name.lowercased()) || valueOf($0).lowercased().contains(name.lowercased())) &&
        roleOf($0) == "AXButton"
    }) {
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        ok("Switched to account '\(name)'")
        return
    }
    // Fallback: Profile menu
    if triggerMenu(conn.app, path: ["Profile", name]) {
        ok("Switched to account '\(name)' (via menu)")
    } else {
        fail("Account '\(name)' not found")
    }
}

// MARK: - Argument Parsing

let args = Array(CommandLine.arguments.dropFirst())

func argValue(_ flag: String) -> String? {
    guard let idx = args.firstIndex(of: flag), idx + 1 < args.count else { return nil }
    return args[idx + 1]
}

func usage() -> Never {
    let text = """
    outlook-ax — Outlook Accessibility Bridge CLI

    Usage:
      outlook-ax status                          Show Outlook status + notification count
      outlook-ax notifications                   Show notification count

      MAIL — Read:
      outlook-ax mail current                    Read current email
      outlook-ax mail inbox [--limit N]          List inbox messages (default 10)
      outlook-ax mail search "query"             Execute search in Outlook

      MAIL — Actions:
      outlook-ax mail reply                      Reply to current email
      outlook-ax mail reply-all                  Reply-all to current email
      outlook-ax mail forward                    Forward current email
      outlook-ax mail delete                     Delete current email
      outlook-ax mail archive                    Archive current email
      outlook-ax mail flag                       Toggle flag on current email
      outlook-ax mail read                       Toggle read/unread on current email
      outlook-ax mail move                       Open move-to-folder dialog
      outlook-ax mail report                     Report spam/phishing
      outlook-ax mail react                      Open emoji reaction picker
      outlook-ax mail summarize                  Copilot summarize current email
      outlook-ax mail filter                     Open filter/sort popup

      MAIL — Compose:
      outlook-ax mail compose                    Compose new email
        [--to "email@example.com"]
        [--subject "Subject line"]
        [--body "Message text"]

      MAIL — Folders:
      outlook-ax mail folders                    List mail folders
      outlook-ax mail folder "name"              Switch to folder

      CALENDAR — Read:
      outlook-ax calendar today                  List today's events

      CALENDAR — Create:
      outlook-ax calendar create                 Create calendar event
        --subject "Title"
        [--attendee "email"]
        [--date "20.04.2026"]
        [--time "10:00"]
        [--send]

      CALENDAR — View & Navigation:
      outlook-ax calendar view <mode>            Switch view
        day|week|month|arbeitswoche|dreitage|liste
      outlook-ax calendar navigate               Navigate calendar
        --direction today|next|prev
        --date "20. April"
      outlook-ax calendar timescale <min>        Set time grid (60|30|15|10|6|5)
      outlook-ax calendar filter <type>          Filter events
        all|appointments|meetings|categories|recurring|declined
      outlook-ax calendar color <color>          Set calendar color
        blue|green|orange|platinum|yellow|cyan|magenta|brown|burgundy|teal|lilac

      CALENDAR — Manage:
      outlook-ax calendar calendars              List calendars with visibility
      outlook-ax calendar toggle "name"          Show/hide a calendar

      CALENDAR — Event Actions:
      outlook-ax calendar accept                 Accept meeting invite
      outlook-ax calendar tentative              Tentatively accept
      outlook-ax calendar decline                Decline meeting invite
      outlook-ax calendar join                   Join online meeting
      outlook-ax calendar duplicate              Duplicate event
      outlook-ax calendar categorize [name]      Set/open category
      outlook-ax calendar private                Toggle private flag
      outlook-ax calendar show-as <status>       Set availability
        free|busy|tentative|oof|elsewhere

      NAVIGATION:
      outlook-ax navigate <target>               Switch Outlook view
        calendar|mail|people|todo|copilot|onedrive|favorites|org-explorer

      SYSTEM:
      outlook-ax sync                            Trigger sync
      outlook-ax auto-reply                      Open auto-reply/OOF settings
      outlook-ax myday                           Toggle My Day panel
      outlook-ax account "name"                  Switch account/profile

    Flags:
      --json                                     Output as JSON (all read commands)

    Examples:
      outlook-ax status --json
      outlook-ax mail current --json
      outlook-ax mail reply-all
      outlook-ax mail flag
      outlook-ax mail summarize
      outlook-ax calendar create --subject "Standup" --time "09:00" --send
      outlook-ax calendar view arbeitswoche
      outlook-ax calendar accept
      outlook-ax calendar show-as busy
      outlook-ax calendar timescale 15
      outlook-ax sync
      outlook-ax navigate org-explorer
    """
    print(text)
    exit(0)
}

// MARK: - Main

guard args.count >= 1 else { usage() }

let cmd = args[0]

switch cmd {
case "status":
    cmdStatus()

case "notifications":
    cmdNotifications()

case "mail":
    guard args.count >= 2 else { usage() }
    switch args[1] {
    case "current":
        cmdMailCurrent()
    case "inbox":
        let limit = Int(argValue("--limit") ?? "10") ?? 10
        cmdMailInbox(limit: limit)
    case "search":
        guard args.count >= 3 else { fail("Usage: outlook-ax mail search \"query\"") }
        cmdMailSearch(query: args[2])
    case "reply":
        cmdMailReply()
    case "forward":
        cmdMailForward()
    case "delete":
        cmdMailDelete()
    case "archive":
        cmdMailArchive()
    case "compose":
        cmdMailCompose(to: argValue("--to"), subject: argValue("--subject"), body: argValue("--body"))
    case "folders":
        cmdMailFolders()
    case "folder":
        guard args.count >= 3 else { fail("Usage: outlook-ax mail folder \"name\"") }
        cmdMailFolder(name: args[2])
    case "reply-all", "replyall":
        cmdMailReplyAll()
    case "flag":
        cmdMailFlag()
    case "read", "unread", "mark":
        cmdMailReadUnread()
    case "move":
        cmdMailMove()
    case "report":
        cmdMailReport()
    case "react":
        cmdMailReact()
    case "summarize":
        cmdMailSummarize()
    case "filter":
        cmdMailFilter()
    default:
        usage()
    }

case "calendar":
    guard args.count >= 2 else { usage() }
    switch args[1] {
    case "today":
        cmdCalendarToday()
    case "create":
        guard let subject = argValue("--subject") else { fail("--subject required") }
        cmdCalendarCreate(
            subject: subject,
            attendee: argValue("--attendee"),
            date: argValue("--date"),
            time: argValue("--time"),
            send: args.contains("--send")
        )
    case "view":
        guard args.count >= 3 else { fail("Usage: outlook-ax calendar view day|week|month") }
        cmdCalendarView(mode: args[2])
    case "navigate":
        cmdCalendarNavigate(direction: argValue("--direction"), date: argValue("--date"))
    case "calendars":
        cmdCalendarCalendars()
    case "toggle":
        guard args.count >= 3 else { fail("Usage: outlook-ax calendar toggle \"name\"") }
        cmdCalendarToggle(name: args[2])
    case "timescale":
        guard args.count >= 3 else { fail("Usage: outlook-ax calendar timescale 30|15|10|6|5") }
        cmdCalendarTimescale(minutes: args[2])
    case "filter":
        guard args.count >= 3 else { fail("Usage: outlook-ax calendar filter all|appointments|meetings") }
        cmdCalendarFilter(filter: args[2])
    case "color":
        guard args.count >= 3 else { fail("Usage: outlook-ax calendar color blue|green|orange|...") }
        cmdCalendarColor(color: args[2])
    case "accept":
        cmdCalendarAccept()
    case "tentative":
        cmdCalendarTentative()
    case "decline":
        cmdCalendarDecline()
    case "join":
        cmdCalendarJoin()
    case "duplicate":
        cmdCalendarDuplicate()
    case "categorize":
        cmdCalendarCategorize(category: args.count >= 3 ? args[2] : nil)
    case "private":
        cmdCalendarPrivate()
    case "show-as", "showas":
        guard args.count >= 3 else { fail("Usage: outlook-ax calendar show-as free|busy|tentative|oof|elsewhere") }
        cmdCalendarShowAs(status: args[2])
    default:
        usage()
    }

case "navigate":
    guard args.count >= 2 else { fail("Usage: outlook-ax navigate calendar|mail|people|todo|copilot|onedrive|favorites|org-explorer") }
    cmdNavigate(to: args[1])

case "sync":
    cmdSync()

case "auto-reply", "autoreply", "oof":
    cmdAutoReply()

case "myday", "my-day":
    cmdMyDay()

case "account":
    guard args.count >= 2 else { fail("Usage: outlook-ax account \"name\"") }
    cmdAccount(name: args[1])

case "help", "--help", "-h":
    usage()

default:
    usage()
}
