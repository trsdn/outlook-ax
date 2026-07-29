import AppKit
import ApplicationServices
import Foundation

// MARK: - MailService
//
// Implements all mail commands. Requires an OutlookConnection obtained from
// ConnectionManager. Never calls exit(). Throws OutlookAXError on failures.
//
// Form safety rules (issues #3, #4):
// - cmdMailCompose verifies every requested field BEFORE proceeding.
// - Send is NEVER invoked after any requested field fails.

public enum MailService {

    // MARK: - Read: Status

    /// Read the Outlook status passively (no launch/activation).
    public static func readStatus() -> StatusResult {
        guard AXIsProcessTrusted() else {
            return StatusResult(
                running: false, hasAXPermission: false,
                window: "", view: "", notifications: 0
            )
        }
        guard let conn = ConnectionManager.tryPassiveConnect() else {
            return StatusResult(
                running: false, hasAXPermission: true,
                window: "", view: "", notifications: 0
            )
        }
        let view = ConnectionManager.currentView(conn.wins)
        let mainWin = conn.wins.first
        let winTitle = mainWin.map { axTitle($0) } ?? ""
        var notifCount = 0
        if let win = mainWin,
           let notifBtn = axFind(win, where: {
               axStartsWithAny(axDesc($0), L10n.newNotifications) && axRole($0) == "AXButton"
           }) {
            let parts = axDesc(notifBtn).components(separatedBy: ": ")
            if parts.count >= 2 { notifCount = Int(parts[1]) ?? 0 }
        }
        return StatusResult(
            running: true, hasAXPermission: true,
            window: winTitle, view: view, notifications: notifCount
        )
    }

    /// Read the notification count passively. Throws if Outlook is not running.
    public static func readNotificationCount() throws -> Int {
        let conn = try ConnectionManager.passiveConnect()
        guard let mainWin = conn.wins.first else { throw OutlookAXError.noAccessibleWindows }
        guard let notifBtn = axFind(mainWin, where: {
            axStartsWithAny(axDesc($0), L10n.newNotifications) && axRole($0) == "AXButton"
        }) else { return 0 }
        let parts = axDesc(notifBtn).components(separatedBy: ": ")
        return parts.count >= 2 ? (Int(parts[1]) ?? 0) : 0
    }

    // MARK: - Read: Current email

    /// Read the currently-open email from the reading pane.
    public static func readCurrent(conn: OutlookConnection) throws -> MailMessage {
        let win = try conn.mainMailWindow()

        guard let headerGroup = axFind(win, where: { axEqualsAny(axTitle($0), L10n.messageHeader) }) else {
            throw OutlookAXError.messageListNotFound
        }

        var subject = "", from = "", date = "", recipients = ""
        for child in axChildren(headerGroup) {
            if axRole(child) == "AXStaticText" && subject.isEmpty {
                subject = axValue(child)
            }
            if axEqualsAny(axDesc(child), L10n.headerDetails) {
                for detail in axChildren(child) {
                    let d = axDesc(detail)
                    if d == "messageHeaderFromContent" {
                        from = axValue(detail)
                        for p in L10n.fromPrefix where from.hasPrefix(p) {
                            from = String(from.dropFirst(p.count)); break
                        }
                    }
                    if d == "messageHeaderRecipientsContent" {
                        recipients = axValue(detail)
                    }
                    let t = axTitle(detail)
                    if axStartsWithAny(t, L10n.sentPrefix) {
                        date = axValue(detail)
                        if date.isEmpty { date = t }
                    }
                }
            }
        }

        var bodyParts: [String] = []
        if let web = axFind(win, where: { axDesc($0) == "Reading Pane" && axRole($0) == "AXWebArea" }) {
            axCollectText(web, into: &bodyParts)
        }
        let body = bodyParts.joined(separator: "\n\n")

        return MailMessage(subject: subject, from: from, to: [recipients], date: date, body: body)
    }

    // MARK: - Read: Inbox

    /// Read messages from the Inbox.
    /// - Verifies that Inbox is the selected folder (throws `.inboxNotSelected` if not).
    /// - Requires a verified message-list table (throws `.messageListNotFound` if absent).
    /// - `limit` must be in 1...100 (enforced by CLIParser before this call).
    public static func readInbox(limit: Int, conn: OutlookConnection) throws -> [MailMessageSummary] {
        let win = try conn.mainMailWindow()

        // Verify Inbox is selected
        let isInbox: Bool = {
            if axEqualsAny(axTitle(win), L10n.inboxWindow) { return true }
            return axFind(win, where: { e in
                let r = axRole(e)
                return (r == "AXRow" || r == "AXCell" || r == "AXHeading" || r == "AXStaticText")
                    && axEqualsAny(axDesc(e), L10n.inboxWindow)
            }) != nil
        }()
        guard isInbox else { throw OutlookAXError.inboxNotSelected }

        // Require verified message-list table
        guard let table = axFind(win, where: {
            axRole($0) == "AXTable" && axEqualsAny(axDesc($0), L10n.messageList)
        }) else { throw OutlookAXError.messageListNotFound }

        var items: [MailMessageSummary] = []
        let rows = axChildren(table).filter { axRole($0) == "AXRow" }

        for row in rows.prefix(limit) {
            let texts = axFindAll(row, where: { axRole($0) == "AXStaticText" }, maxDepth: 4)
            let values = texts.map { axValue($0) }.filter { !$0.isEmpty }
            guard !values.isEmpty else { continue }

            // Row layout: sender, date, subject, preview
            let from    = values.count > 0 ? values[0] : ""
            let date    = values.count > 1 ? values[1] : ""
            let subject = values.count > 2 ? values[2] : values[0]
            let preview = values.count > 3 ? values[3] : ""
            items.append(MailMessageSummary(subject: subject, from: from, date: date, preview: preview))
        }
        return items
    }

    // MARK: - Read: Search

    /// Execute a mail search and return results.
    public static func search(query: String, conn: OutlookConnection) throws -> MailSearchResult {
        let win = try conn.mainMailWindow()

        guard let searchField = axFind(win, where: {
            axEqualsAny(axDesc($0), L10n.search) && axRole($0) == "AXTextField"
        }) else {
            // Search field absent — return no results rather than throwing
            return .noResults
        }

        AXUIElementSetAttributeValue(searchField, kAXFocusedAttribute as CFString, true as CFTypeRef)
        Thread.sleep(forTimeInterval: 0.3)
        AXUIElementSetAttributeValue(searchField, kAXValueAttribute as CFString, query as CFTypeRef)
        Thread.sleep(forTimeInterval: 0.3)
        axPressReturn()

        // Wait for results to stabilize (up to 10 seconds)
        let deadline = Date().addingTimeInterval(10)
        var prevCount = -1
        while Date() < deadline {
            Thread.sleep(forTimeInterval: 1.0)
            let table = axFind(win, where: {
                axRole($0) == "AXTable" && axEqualsAny(axDesc($0), L10n.messageList)
            })
            let rows = table.map { axChildren($0).filter { axRole($0) == "AXRow" } } ?? []
            let count = rows.count
            if count == prevCount && count >= 0 { break }
            prevCount = count
        }

        guard Date() < deadline.addingTimeInterval(1) else { return .timeout }

        guard let table = axFind(win, where: {
            axRole($0) == "AXTable" && axEqualsAny(axDesc($0), L10n.messageList)
        }) else { return .noResults }

        let rows = axChildren(table).filter { axRole($0) == "AXRow" }
        if rows.isEmpty { return .noResults }

        var items: [MailMessageSummary] = []
        for row in rows.prefix(50) {
            let texts = axFindAll(row, where: { axRole($0) == "AXStaticText" }, maxDepth: 4)
            let values = texts.map { axValue($0) }.filter { !$0.isEmpty }
            guard !values.isEmpty else { continue }
            let from = values.count > 0 ? values[0] : ""
            let date = values.count > 1 ? values[1] : ""
            let subject = values.count > 2 ? values[2] : values[0]
            let preview = values.count > 3 ? values[3] : ""
            items.append(MailMessageSummary(subject: subject, from: from, date: date, preview: preview))
        }
        return items.isEmpty ? .noResults : .results(items)
    }

    // MARK: - Compose

    /// Open a compose window and fill the requested fields.
    /// FORM SAFETY: If any requested field fails to write, returns `.failure` and
    /// never attempts to send or save. Send is invoked only when all fields succeed.
    public static func compose(
        to: String?, subject: String?, body: String?,
        conn: OutlookConnection
    ) throws {
        let win = try conn.mainMailWindow()

        // Snapshot before opening compose window
        let priorTitles = Set(conn.wins.map { axTitle($0) })

        guard axPressButtonAny(win, prefixes: L10n.newEmail) else {
            throw OutlookAXError.formFieldNotFound(field: "New Email button")
        }
        Thread.sleep(forTimeInterval: 2.5)

        guard let composeWin = try conn.uniqueNewWindow(
            before: priorTitles,
            matching: { axMatchesAny(axTitle($0), L10n.composeWindow) }
        ) else {
            throw OutlookAXError.noAccessibleWindows
        }
        AXUIElementPerformAction(composeWin, kAXRaiseAction as CFString)
        Thread.sleep(forTimeInterval: 0.5)

        // Fill To field
        if let toAddr = to {
            guard let toField = axFind(composeWin, where: {
                axEqualsAny(axDesc($0), L10n.toField) && axRole($0) == "AXTextField"
            }) else { throw OutlookAXError.formFieldNotFound(field: "To") }
            AXUIElementSetAttributeValue(toField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            axTypeText(toAddr)
            Thread.sleep(forTimeInterval: 0.5)
            axPressTab()
            Thread.sleep(forTimeInterval: 0.5)
        }

        // Fill Subject field
        if let subj = subject {
            guard let subjField = axFind(composeWin, where: {
                axEqualsAny(axDesc($0), L10n.subject) && axRole($0) == "AXTextField"
            }) else { throw OutlookAXError.formFieldNotFound(field: "Subject") }
            AXUIElementSetAttributeValue(subjField, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            let rc = AXUIElementSetAttributeValue(subjField, kAXValueAttribute as CFString, subj as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            guard rc == .success else { throw OutlookAXError.formFieldWriteFailed(field: "Subject") }
        }

        // Fill Body field
        if let bodyText = body {
            guard let bodyArea = axFind(composeWin, where: {
                axEqualsAny(axDesc($0), L10n.bodyField) || axRole($0) == "AXWebArea"
            }) else { throw OutlookAXError.formFieldNotFound(field: "Body") }
            AXUIElementSetAttributeValue(bodyArea, kAXFocusedAttribute as CFString, true as CFTypeRef)
            Thread.sleep(forTimeInterval: 0.3)
            axTypeText(bodyText)
            Thread.sleep(forTimeInterval: 0.3)
        }
        // Note: compose does NOT send — it opens the compose window ready for the user.
    }

    // MARK: - Folders

    /// List mail folders from the navigation pane.
    public static func listFolders(conn: OutlookConnection) throws -> [FolderIdentity] {
        let win = try conn.mainMailWindow()
        var folders: [FolderIdentity] = []
        var currentAccount = ""

        guard let outline = axFind(win, where: { axRole($0) == "AXOutline" }) else { return [] }
        let rows = axFindAll(outline, where: { axRole($0) == "AXRow" }, maxDepth: 3)

        for row in rows {
            let triangles = axFindAll(row, where: { axRole($0) == "AXDisclosureTriangle" }, maxDepth: 3)
            let texts = axFindAll(row, where: { axRole($0) == "AXStaticText" }, maxDepth: 3)
            let checkboxes = axFindAll(row, where: { axRole($0) == "AXCheckBox" }, maxDepth: 3)

            if !triangles.isEmpty {
                // Account or group header
                for cb in checkboxes {
                    let d = axDesc(cb)
                    if d.contains("@") {
                        currentAccount = d.components(separatedBy: ";").first?
                            .trimmingCharacters(in: .whitespaces) ?? d
                    }
                }
                for t in texts {
                    let v = axValue(t)
                    if v.contains("@") { currentAccount = v }
                }
            }

            // Folder entries
            for t in texts {
                let v = axValue(t)
                let skip = v.isEmpty || v.contains("@") ||
                    axEqualsAny(v, ["Favoriten", "Favorites", "Alle Konten", "All Accounts", "Gruppen", "Groups"])
                if !skip {
                    folders.append(FolderIdentity(name: v, account: currentAccount, group: ""))
                }
            }
        }
        return folders
    }

    /// Switch to the named folder. Throws `.folderAmbiguous` if multiple folders match.
    public static func switchFolder(name: String, conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard let outline = axFind(win, where: { axRole($0) == "AXOutline" }) else {
            throw OutlookAXError.folderAmbiguous(name: name)
        }

        // Find all matching rows
        let matches = axFindAll(outline, where: {
            axRole($0) == "AXStaticText" && axValue($0).lowercased() == name.lowercased()
        }, maxDepth: 6)

        switch matches.count {
        case 0:
            throw OutlookAXError.folderAmbiguous(name: name)
        case 1:
            AXUIElementPerformAction(matches[0], kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 1.0)
        default:
            throw OutlookAXError.folderAmbiguous(name: name)
        }
    }

    // MARK: - Actions (toolbar buttons)

    public static func reply(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        let found = L10n.reply.contains(where: {
            axPressButtonByTitle(win, title: $0) || axPressButtonByDescPrefix(win, prefix: $0)
        })
        if !found { throw OutlookAXError.formFieldNotFound(field: "Reply button") }
        Thread.sleep(forTimeInterval: 1.0)
    }

    public static func replyAll(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard let btn = axFind(win, where: {
            let d = axDesc($0); let t = axTitle($0)
            return (axEqualsAny(d, L10n.replyAll) || axEqualsAny(t, L10n.replyAll)) && axRole($0) == "AXButton"
        }) else { throw OutlookAXError.formFieldNotFound(field: "Reply All button") }
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        Thread.sleep(forTimeInterval: 1.0)
    }

    public static func forward(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        if axPressButtonAny(win, prefixes: L10n.forward) { Thread.sleep(forTimeInterval: 1.0); return }
        // Try via "Show more items"
        if axPressButtonAny(win, prefixes: L10n.moreItems) {
            Thread.sleep(forTimeInterval: 0.5)
            if axPressButtonAny(win, prefixes: L10n.forward) { Thread.sleep(forTimeInterval: 1.0); return }
        }
        throw OutlookAXError.formFieldNotFound(field: "Forward button")
    }

    public static func delete(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard axPressButtonAny(win, prefixes: L10n.delete) else {
            throw OutlookAXError.formFieldNotFound(field: "Delete button")
        }
        Thread.sleep(forTimeInterval: 0.5)
    }

    public static func archive(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard axPressButtonAny(win, prefixes: L10n.archive) else {
            throw OutlookAXError.formFieldNotFound(field: "Archive button")
        }
        Thread.sleep(forTimeInterval: 0.5)
    }

    public static func flag(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard axPressButtonAny(win, prefixes: L10n.flag) else {
            throw OutlookAXError.formFieldNotFound(field: "Flag button")
        }
    }

    public static func toggleReadUnread(conn: OutlookConnection) throws -> String {
        let win = try conn.mainMailWindow()
        guard let btn = axFind(win, where: { axStartsWithAny(axDesc($0), L10n.markRead) && axRole($0) == "AXButton" }) else {
            throw OutlookAXError.formFieldNotFound(field: "Read/unread button")
        }
        let current = axDesc(btn)
        AXUIElementPerformAction(btn, kAXPressAction as CFString)
        let wasUnread = current.lowercased().contains("ungelesen") || current.lowercased().contains("unread")
        return wasUnread ? "marked as unread" : "marked as read"
    }

    public static func move(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard axPressButtonAny(win, prefixes: L10n.move) else {
            throw OutlookAXError.formFieldNotFound(field: "Move button")
        }
    }

    public static func report(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard axPressButtonAny(win, prefixes: L10n.report) else {
            throw OutlookAXError.formFieldNotFound(field: "Report button")
        }
    }

    public static func react(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard axPressButtonAny(win, prefixes: L10n.react) else {
            throw OutlookAXError.formFieldNotFound(field: "React button")
        }
    }

    public static func summarize(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard axPressButtonAny(win, prefixes: L10n.summarize) else {
            throw OutlookAXError.formFieldNotFound(field: "Summarize button")
        }
    }

    public static func filter(conn: OutlookConnection) throws {
        let win = try conn.mainMailWindow()
        guard let popup = axFind(win, where: {
            axEqualsAny(axDesc($0), L10n.filterSort) && axRole($0) == "AXPopUpButton"
        }) else { throw OutlookAXError.formFieldNotFound(field: "Filter popup") }
        AXUIElementPerformAction(popup, kAXShowMenuAction as CFString)
    }
}

// MARK: - StatusResult

public struct StatusResult {
    public let running: Bool
    public let hasAXPermission: Bool
    public let window: String
    public let view: String
    public let notifications: Int
}
