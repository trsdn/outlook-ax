import Foundation

// MARK: - CLICommand
//
// Typed representation of every supported CLI invocation.
// CLIParser maps raw arguments into these cases.

public enum CLICommand: Equatable {
    // Status / system
    case status
    case notifications
    case sync
    case autoReply
    case myDay
    case account(name: String)

    // Mail: read
    case mailCurrent
    case mailInbox(limit: Int)
    case mailSearch(query: String)

    // Mail: actions
    case mailReply
    case mailReplyAll
    case mailForward
    case mailDelete
    case mailArchive
    case mailFlag
    case mailReadUnread
    case mailMove
    case mailReport
    case mailReact
    case mailSummarize
    case mailFilter

    // Mail: compose
    case mailCompose(to: String?, subject: String?, body: String?)

    // Mail: folders
    case mailFolders
    case mailFolder(name: String)

    // Calendar: read
    case calendarToday(details: Bool)

    // Calendar: create
    case calendarCreate(subject: String, attendee: String?, date: String?, time: String?, send: Bool)

    // Calendar: view & navigation
    case calendarView(mode: String)
    case calendarNavigate(direction: String?, date: String?)
    case calendarTimescale(minutes: String)
    case calendarFilter(filter: String)
    case calendarColor(color: String)

    // Calendar: manage
    case calendarCalendars
    case calendarToggle(name: String)

    // Calendar: event actions
    case calendarAccept
    case calendarTentative
    case calendarDecline
    case calendarJoin
    case calendarDuplicate
    case calendarCategorize(category: String?)
    case calendarPrivate
    case calendarShowAs(status: String)

    // Navigation
    case navigate(to: String)

    // Help
    case help
}

// MARK: - CLIParseResult

public enum CLIParseResult {
    case command(CLICommand, json: Bool, details: Bool)
    case help
    case usageError(String)
}
