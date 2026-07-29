import Foundation
import OutlookAX

// MARK: - CLIRunner
//
// Dispatches CLICommand values to their implementations in OutlookAX library.
// Does NOT call exit() — returns an exit code to main.swift.
//
// Exit code contract:
//  0   Success
//  1   Operation failure (AX action failed, form field missing, etc.)
//  64  Usage error (handled by CLIParser before reaching here)
//  69  Outlook not running or no accessible windows
//  77  Accessibility permission denied

public struct CLIRunner {

    private let renderer: OutputRenderer

    public init(renderer: OutputRenderer) {
        self.renderer = renderer
    }

    // MARK: - Main dispatch

    /// Run a parsed CLI command. Returns the process exit code.
    public func run(_ command: CLICommand) -> Int32 {
        switch command {

        // MARK: Help
        case .help:
            renderer.printHelp(usageText)
            return 0

        // MARK: Status / Notifications (passive reads — no launch)
        case .status:
            return runStatus()

        case .notifications:
            return runPassive(command: "notifications") { _ in
                let count = try MailService.readNotificationCount()
                renderer.printSuccess(NotificationData(count: count), command: "notifications")
                if !renderer.jsonMode { print("Notifications: \(count)") }
            }

        // MARK: Mail: read
        case .mailCurrent:
            return runInteractive(command: "mail.current") { conn in
                let msg = try MailService.readCurrent(conn: conn)
                renderer.printSuccess(msg, command: "mail.current")
                if !renderer.jsonMode { printMailMessage(msg) }
            }

        case .mailInbox(let limit):
            return runInteractive(command: "mail.inbox") { conn in
                let items = try MailService.readInbox(limit: limit, conn: conn)
                renderer.printSuccess(items, command: "mail.inbox")
                if !renderer.jsonMode { printInboxItems(items) }
            }

        case .mailSearch(let query):
            return runInteractive(command: "mail.search") { conn in
                let result = try MailService.search(query: query, conn: conn)
                switch result {
                case .results(let items):
                    renderer.printSuccess(items, command: "mail.search")
                    if !renderer.jsonMode { printInboxItems(items) }
                case .noResults:
                    renderer.printSuccess(["message": "No results for '\(query)'"], command: "mail.search")
                    if !renderer.jsonMode { print("No results for '\(query)'") }
                case .timeout:
                    throw OutlookAXError.searchTimeout
                }
            }

        // MARK: Mail: compose
        case .mailCompose(let to, let subject, let body):
            return runInteractive(command: "mail.compose") { conn in
                try MailService.compose(to: to, subject: subject, body: body, conn: conn)
                let result = ActionResult(action: "compose", details: [
                    "to": to ?? "", "subject": subject ?? ""
                ])
                renderer.printSuccess(result, command: "mail.compose")
                if !renderer.jsonMode { print("Compose window opened") }
            }

        // MARK: Mail: folders
        case .mailFolders:
            return runInteractive(command: "mail.folders") { conn in
                let folders = try MailService.listFolders(conn: conn)
                renderer.printSuccess(folders, command: "mail.folders")
                if !renderer.jsonMode { printFolders(folders) }
            }

        case .mailFolder(let name):
            return runInteractive(command: "mail.folder") { conn in
                try MailService.switchFolder(name: name, conn: conn)
                let result = ActionResult(action: "folder", details: ["folder": name])
                renderer.printSuccess(result, command: "mail.folder")
                if !renderer.jsonMode { print("Switched to folder: \(name)") }
            }

        // MARK: Mail: actions
        case .mailReply:
            return runInteractive(command: "mail.reply") { conn in
                try MailService.reply(conn: conn)
                printAction("Reply opened", command: "mail.reply")
            }

        case .mailReplyAll:
            return runInteractive(command: "mail.reply-all") { conn in
                try MailService.replyAll(conn: conn)
                printAction("Reply-all opened", command: "mail.reply-all")
            }

        case .mailForward:
            return runInteractive(command: "mail.forward") { conn in
                try MailService.forward(conn: conn)
                printAction("Forward opened", command: "mail.forward")
            }

        case .mailDelete:
            return runInteractive(command: "mail.delete") { conn in
                try MailService.delete(conn: conn)
                printAction("Email deleted", command: "mail.delete")
            }

        case .mailArchive:
            return runInteractive(command: "mail.archive") { conn in
                try MailService.archive(conn: conn)
                printAction("Email archived", command: "mail.archive")
            }

        case .mailFlag:
            return runInteractive(command: "mail.flag") { conn in
                try MailService.flag(conn: conn)
                printAction("Flag toggled", command: "mail.flag")
            }

        case .mailReadUnread:
            return runInteractive(command: "mail.read") { conn in
                let action = try MailService.toggleReadUnread(conn: conn)
                printAction(action, command: "mail.read")
            }

        case .mailMove:
            return runInteractive(command: "mail.move") { conn in
                try MailService.move(conn: conn)
                printAction("Move dialog opened", command: "mail.move")
            }

        case .mailReport:
            return runInteractive(command: "mail.report") { conn in
                try MailService.report(conn: conn)
                printAction("Report dialog opened", command: "mail.report")
            }

        case .mailReact:
            return runInteractive(command: "mail.react") { conn in
                try MailService.react(conn: conn)
                printAction("React picker opened", command: "mail.react")
            }

        case .mailSummarize:
            return runInteractive(command: "mail.summarize") { conn in
                try MailService.summarize(conn: conn)
                printAction("Copilot summarize triggered", command: "mail.summarize")
            }

        case .mailFilter:
            return runInteractive(command: "mail.filter") { conn in
                try MailService.filter(conn: conn)
                printAction("Filter popup opened", command: "mail.filter")
            }

        // MARK: Calendar: read
        case .calendarToday(let details):
            return runInteractive(command: "calendar.today") { conn in
                let events = try CalendarService.readToday(conn: conn)
                renderer.printSuccess(events, command: "calendar.today")
                if !renderer.jsonMode { printCalendarEvents(events, details: details) }
            }

        // MARK: Calendar: create
        case .calendarCreate(let subject, let attendee, let date, let time, let send):
            return runInteractive(command: "calendar.create") { conn in
                try CalendarService.createEvent(
                    subject: subject, attendee: attendee,
                    date: date, time: time, send: send, conn: conn
                )
                let result = ActionResult(action: "calendar.create", details: [
                    "subject": subject, "attendee": attendee ?? "",
                    "date": date ?? "", "time": time ?? "",
                    "sent": send ? "true" : "false"
                ])
                renderer.printSuccess(result, command: "calendar.create")
                if !renderer.jsonMode {
                    print(send ? "Event sent: \(subject)" : "Event form filled: \(subject)")
                }
            }

        // MARK: Calendar: view & navigation
        case .calendarView(let mode):
            return runInteractive(command: "calendar.view") { conn in
                try CalendarService.switchView(mode: mode, conn: conn)
                printAction("Calendar view: \(mode)", command: "calendar.view")
            }

        case .calendarNavigate(let direction, let date):
            return runInteractive(command: "calendar.navigate") { conn in
                try CalendarService.navigate(direction: direction, date: date, conn: conn)
                printAction("Calendar navigated", command: "calendar.navigate")
            }

        case .calendarTimescale(let minutes):
            return runInteractive(command: "calendar.timescale") { conn in
                try CalendarService.setTimescale(minutes: minutes, conn: conn)
                printAction("Timescale: \(minutes) min", command: "calendar.timescale")
            }

        case .calendarFilter(let filter):
            return runInteractive(command: "calendar.filter") { conn in
                try CalendarService.setFilter(filter: filter, conn: conn)
                printAction("Calendar filter: \(filter)", command: "calendar.filter")
            }

        case .calendarColor(let color):
            return runInteractive(command: "calendar.color") { conn in
                try CalendarService.setColor(color: color, conn: conn)
                printAction("Calendar color: \(color)", command: "calendar.color")
            }

        // MARK: Calendar: manage
        case .calendarCalendars:
            return runInteractive(command: "calendar.calendars") { conn in
                let cals = try CalendarService.listCalendars(conn: conn)
                renderer.printSuccess(cals, command: "calendar.calendars")
                if !renderer.jsonMode { printCalendars(cals) }
            }

        case .calendarToggle(let name):
            return runInteractive(command: "calendar.toggle") { conn in
                let visible = try CalendarService.toggleCalendar(name: name, conn: conn)
                let result = ActionResult(action: "toggle", details: [
                    "name": name, "visible": visible ? "true" : "false"
                ])
                renderer.printSuccess(result, command: "calendar.toggle")
                if !renderer.jsonMode { print("Calendar \(visible ? "shown" : "hidden"): \(name)") }
            }

        // MARK: Calendar: event actions
        case .calendarAccept:
            return runInteractive(command: "calendar.accept") { conn in
                try CalendarService.accept(conn: conn)
                printAction("Meeting accepted", command: "calendar.accept")
            }

        case .calendarTentative:
            return runInteractive(command: "calendar.tentative") { conn in
                try CalendarService.tentative(conn: conn)
                printAction("Meeting tentatively accepted", command: "calendar.tentative")
            }

        case .calendarDecline:
            return runInteractive(command: "calendar.decline") { conn in
                try CalendarService.decline(conn: conn)
                printAction("Meeting declined", command: "calendar.decline")
            }

        case .calendarJoin:
            return runInteractive(command: "calendar.join") { conn in
                try CalendarService.join(conn: conn)
                printAction("Joining online meeting", command: "calendar.join")
            }

        case .calendarDuplicate:
            return runInteractive(command: "calendar.duplicate") { conn in
                try CalendarService.duplicate(conn: conn)
                printAction("Event duplicated", command: "calendar.duplicate")
            }

        case .calendarCategorize(let category):
            return runInteractive(command: "calendar.categorize") { conn in
                try CalendarService.categorize(category: category, conn: conn)
                let label = category.map { "Category '\($0)' applied" } ?? "Categorize menu opened"
                printAction(label, command: "calendar.categorize")
            }

        case .calendarPrivate:
            return runInteractive(command: "calendar.private") { conn in
                try CalendarService.setPrivate(conn: conn)
                printAction("Private flag toggled", command: "calendar.private")
            }

        case .calendarShowAs(let status):
            return runInteractive(command: "calendar.show-as") { conn in
                try CalendarService.setShowAs(status: status, conn: conn)
                printAction("Show-as: \(status)", command: "calendar.show-as")
            }

        // MARK: Navigation
        case .navigate(let target):
            return runInteractive(command: "navigate") { conn in
                try CalendarService.navigateTo(target: target, conn: conn)
                printAction("Navigated to \(target)", command: "navigate")
            }

        // MARK: System
        case .sync:
            return runInteractive(command: "sync") { conn in
                try CalendarService.sync(conn: conn)
                printAction("Sync triggered", command: "sync")
            }

        case .autoReply:
            return runInteractive(command: "auto-reply") { conn in
                try CalendarService.openAutoReply(conn: conn)
                printAction("Auto-reply settings opened", command: "auto-reply")
            }

        case .myDay:
            return runInteractive(command: "myday") { conn in
                try CalendarService.toggleMyDay(conn: conn)
                printAction("My Day panel toggled", command: "myday")
            }

        case .account(let name):
            return runInteractive(command: "account") { conn in
                try CalendarService.switchAccount(name: name, conn: conn)
                printAction("Switched to account: \(name)", command: "account")
            }
        }
    }

    // MARK: - Usage

    private var usageText: String {
        """
        outlook-ax — Outlook Accessibility Bridge CLI

        Usage:
          outlook-ax status                          Show Outlook status + notification count
          outlook-ax notifications                   Show notification count

          MAIL — Read:
          outlook-ax mail current                    Read current email
          outlook-ax mail inbox [--limit N]          List inbox messages (1–100, default 10)
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
          outlook-ax calendar view <mode>            Switch view (day|week|month|workweek|threeday|list)
          outlook-ax calendar navigate               Navigate calendar
            --direction today|next|prev
            --date "YYYY-MM-DD"
          outlook-ax calendar timescale <min>        Set time grid (60|30|15|10|6|5)
          outlook-ax calendar filter <type>          Filter events
          outlook-ax calendar color <color>          Set calendar color

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

          NAVIGATION:
          outlook-ax navigate <target>               Switch Outlook view
            calendar|mail|people|todo|copilot|onedrive|favorites|org-explorer

          SYSTEM:
          outlook-ax sync                            Trigger sync
          outlook-ax auto-reply                      Open auto-reply/OOF settings
          outlook-ax myday                           Toggle My Day panel
          outlook-ax account "name"                  Switch account/profile

        Flags:
          --json                                     Output as JSON v2
          --json-version 1                           Force JSON v1 legacy output (with --json)
          --details                                  Include event details (calendar today)

        Exit codes:
          0   Success
          1   Operation failure
          64  Usage / argument error
          69  Outlook not running or no accessible windows
          77  Accessibility permission denied
        """
    }

    // MARK: - Connection helpers

    /// Run a command using a passive connection (no Outlook launch/activation).
    private func runPassive(command: String, action: (OutlookConnection) throws -> Void) -> Int32 {
        do {
            let conn = try ConnectionManager.passiveConnect()
            try action(conn)
            return 0
        } catch let err as OutlookAXError {
            renderer.printError(code: err.code, message: err.errorDescription ?? err.code,
                                command: command)
            return err.exitCode
        } catch {
            renderer.printError(code: "unexpectedError", message: error.localizedDescription, command: command)
            return 1
        }
    }

    /// Run a command using an interactive connection (launch + activate Outlook if needed).
    private func runInteractive(command: String, action: (OutlookConnection) throws -> Void) -> Int32 {
        do {
            let conn = try ConnectionManager.interactiveConnect()
            try action(conn)
            return 0
        } catch let err as OutlookAXError {
            renderer.printError(code: err.code, message: err.errorDescription ?? err.code,
                                command: command)
            return err.exitCode
        } catch {
            renderer.printError(code: "unexpectedError", message: error.localizedDescription, command: command)
            return 1
        }
    }

    // MARK: - Status (special: never fails, reports state)

    private func runStatus() -> Int32 {
        let s = MailService.readStatus()
        let data = StatusData(
            running: s.running,
            window: s.window,
            view: s.view,
            notifications: s.notifications
        )
        renderer.printSuccess(data, command: "status")
        if !renderer.jsonMode {
            if s.running {
                print("Outlook: running")
                if !s.window.isEmpty { print("Window:  \(s.window)") }
                print("View:    \(s.view)")
                if s.notifications > 0 { print("Notifications: \(s.notifications)") }
            } else if !s.hasAXPermission {
                print("error: Accessibility permission denied")
            } else {
                print("Outlook: not running")
            }
        }
        return 0
    }

    // MARK: - Action helper

    private func printAction(_ msg: String, command: String) {
        let result = ActionResult(action: command, details: [:])
        renderer.printSuccess(result, command: command)
        if !renderer.jsonMode { print(msg) }
    }

    // MARK: - Plain-text formatters

    private func printMailMessage(_ msg: MailMessage) {
        print("Subject: \(msg.subject)")
        print("From:    \(msg.from)")
        if !msg.to.isEmpty { print("To:      \(msg.to.joined(separator: ", "))") }
        print("Date:    \(msg.date)")
        print("---")
        print(msg.body)
    }

    private func printInboxItems(_ items: [MailMessageSummary]) {
        if items.isEmpty { print("No messages found"); return }
        for (i, item) in items.enumerated() {
            print("\(i+1). \(item.subject)")
            print("   From: \(item.from)  Date: \(item.date)")
            if !item.preview.isEmpty { print("   \(String(item.preview.prefix(80)))") }
        }
    }

    private func printFolders(_ folders: [FolderIdentity]) {
        var lastAccount = ""
        for f in folders {
            if f.account != lastAccount && !f.account.isEmpty {
                print("\n[\(f.account)]"); lastAccount = f.account
            }
            print("  \(f.name)")
        }
    }

    private func printCalendarEvents(_ events: [CalendarEvent], details: Bool) {
        if events.isEmpty { print("No events today"); return }
        var lastDate = ""
        for ev in events {
            if ev.date != lastDate && !ev.date.isEmpty {
                if !lastDate.isEmpty { print("") }
                print("── \(ev.date) ──")
                lastDate = ev.date
            }
            var line = ev.isAllDay ? "[all day] " : (ev.start.isEmpty ? "" : "\(ev.start)-\(ev.end)  ")
            line += ev.title
            if ev.myResponse == "declined" { line += " ✗" }
            else if ev.myResponse == "following" { line += " 👁" }
            if !ev.status.isEmpty { line += " [\(ev.status)]" }
            if !ev.categories.isEmpty { line += " (\(ev.categories.joined(separator: ", ")))" }
            print(line)
        }
    }

    private func printCalendars(_ cals: [CalendarIdentity]) {
        if cals.isEmpty { print("No calendars found"); return }
        var lastGroup = ""
        for c in cals {
            if c.group != lastGroup && !c.group.isEmpty { print("\n[\(c.group)]"); lastGroup = c.group }
            let marker = c.isVisible ? "[x]" : "[ ]"
            print("  \(marker) \(c.name)")
        }
    }
}

// MARK: - ActionResult (generic success payload for non-read commands)

public struct ActionResult: Encodable {
    public let action: String
    public let details: [String: String]
    public init(action: String, details: [String: String]) {
        self.action = action; self.details = details
    }
}
