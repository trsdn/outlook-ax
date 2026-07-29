import Foundation

// MARK: - CLIParser
//
// Parses raw argument arrays into CLIParseResult.
// Never calls exit() — returns .usageError for bad input.

public struct CLIParser {

    private let args: [String]

    public init(args: [String]) {
        self.args = args
    }

    // MARK: - Public

    public func parse() -> CLIParseResult {
        let jsonFlag    = args.contains("--json")
        let detailsFlag = args.contains("--details")

        guard !args.isEmpty else {
            return .usageError("No command provided")
        }

        let cmd = args[0]

        switch cmd {
        case "help", "--help", "-h":
            return .help

        case "status":
            return .command(.status, json: jsonFlag, details: detailsFlag)

        case "notifications":
            return .command(.notifications, json: jsonFlag, details: detailsFlag)

        case "sync":
            return .command(.sync, json: jsonFlag, details: detailsFlag)

        case "auto-reply", "autoreply", "oof":
            return .command(.autoReply, json: jsonFlag, details: detailsFlag)

        case "myday", "my-day":
            return .command(.myDay, json: jsonFlag, details: detailsFlag)

        case "account":
            guard args.count >= 2 else {
                return .usageError("Usage: outlook-ax account \"name\"")
            }
            return .command(.account(name: args[1]), json: jsonFlag, details: detailsFlag)

        case "navigate":
            guard args.count >= 2 else {
                return .usageError("Usage: outlook-ax navigate calendar|mail|people|todo|copilot|onedrive|favorites|org-explorer")
            }
            return .command(.navigate(to: args[1]), json: jsonFlag, details: detailsFlag)

        case "mail":
            return parseMail(json: jsonFlag, details: detailsFlag)

        case "calendar":
            return parseCalendar(json: jsonFlag, details: detailsFlag)

        default:
            return .usageError("Unknown command '\(cmd)'. Run 'outlook-ax --help' for usage.")
        }
    }

    // MARK: - Private: mail

    private func parseMail(json: Bool, details: Bool) -> CLIParseResult {
        guard args.count >= 2 else {
            return .usageError("'mail' requires a subcommand. Run 'outlook-ax --help' for usage.")
        }
        let sub = args[1]
        switch sub {
        case "current":
            return .command(.mailCurrent, json: json, details: details)
        case "inbox":
            let limitStr = argValue("--limit") ?? "10"
            guard let limit = Int(limitStr), (1...100).contains(limit) else {
                return .usageError("--limit must be an integer in 1...100 (got: \(limitStr))")
            }
            return .command(.mailInbox(limit: limit), json: json, details: details)
        case "search":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax mail search \"query\"")
            }
            return .command(.mailSearch(query: args[2]), json: json, details: details)
        case "reply":
            return .command(.mailReply, json: json, details: details)
        case "reply-all", "replyall":
            return .command(.mailReplyAll, json: json, details: details)
        case "forward":
            return .command(.mailForward, json: json, details: details)
        case "delete":
            return .command(.mailDelete, json: json, details: details)
        case "archive":
            return .command(.mailArchive, json: json, details: details)
        case "flag":
            return .command(.mailFlag, json: json, details: details)
        case "read", "unread", "mark":
            return .command(.mailReadUnread, json: json, details: details)
        case "move":
            return .command(.mailMove, json: json, details: details)
        case "report":
            return .command(.mailReport, json: json, details: details)
        case "react":
            return .command(.mailReact, json: json, details: details)
        case "summarize":
            return .command(.mailSummarize, json: json, details: details)
        case "filter":
            return .command(.mailFilter, json: json, details: details)
        case "compose":
            return .command(
                .mailCompose(to: argValue("--to"), subject: argValue("--subject"), body: argValue("--body")),
                json: json, details: details
            )
        case "folders":
            return .command(.mailFolders, json: json, details: details)
        case "folder":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax mail folder \"name\"")
            }
            return .command(.mailFolder(name: args[2]), json: json, details: details)
        default:
            return .usageError("Unknown mail subcommand '\(sub)'. Run 'outlook-ax --help' for usage.")
        }
    }

    // MARK: - Private: calendar

    private func parseCalendar(json: Bool, details: Bool) -> CLIParseResult {
        guard args.count >= 2 else {
            return .usageError("'calendar' requires a subcommand. Run 'outlook-ax --help' for usage.")
        }
        let sub = args[1]
        switch sub {
        case "today":
            return .command(.calendarToday(details: details), json: json, details: details)
        case "create":
            guard let subject = argValue("--subject") else {
                return .usageError("--subject required for calendar create")
            }
            return .command(.calendarCreate(
                subject: subject,
                attendee: argValue("--attendee"),
                date: argValue("--date"),
                time: argValue("--time"),
                send: args.contains("--send")
            ), json: json, details: details)
        case "view":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax calendar view day|week|month|...")
            }
            return .command(.calendarView(mode: args[2]), json: json, details: details)
        case "navigate":
            return .command(
                .calendarNavigate(direction: argValue("--direction"), date: argValue("--date")),
                json: json, details: details
            )
        case "timescale":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax calendar timescale 60|30|15|10|6|5")
            }
            return .command(.calendarTimescale(minutes: args[2]), json: json, details: details)
        case "filter":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax calendar filter all|appointments|meetings")
            }
            return .command(.calendarFilter(filter: args[2]), json: json, details: details)
        case "color":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax calendar color blue|green|orange|...")
            }
            return .command(.calendarColor(color: args[2]), json: json, details: details)
        case "calendars":
            return .command(.calendarCalendars, json: json, details: details)
        case "toggle":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax calendar toggle \"name\"")
            }
            return .command(.calendarToggle(name: args[2]), json: json, details: details)
        case "accept":
            return .command(.calendarAccept, json: json, details: details)
        case "tentative":
            return .command(.calendarTentative, json: json, details: details)
        case "decline":
            return .command(.calendarDecline, json: json, details: details)
        case "join":
            return .command(.calendarJoin, json: json, details: details)
        case "duplicate":
            return .command(.calendarDuplicate, json: json, details: details)
        case "categorize":
            let cat = args.count >= 3 ? args[2] : nil
            return .command(.calendarCategorize(category: cat), json: json, details: details)
        case "private":
            return .command(.calendarPrivate, json: json, details: details)
        case "show-as", "showas":
            guard args.count >= 3 else {
                return .usageError("Usage: outlook-ax calendar show-as free|busy|tentative|oof|elsewhere")
            }
            return .command(.calendarShowAs(status: args[2]), json: json, details: details)
        default:
            return .usageError("Unknown calendar subcommand '\(sub)'. Run 'outlook-ax --help' for usage.")
        }
    }

    // MARK: - Helpers

    private func argValue(_ flag: String) -> String? {
        guard let idx = args.firstIndex(of: flag), idx + 1 < args.count else { return nil }
        return args[idx + 1]
    }
}
