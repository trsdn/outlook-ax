import Foundation

/// Typed errors for the OutlookAX library.
public enum OutlookAXError: LocalizedError, Equatable {
    // --- Connectivity ---
    case accessibilityPermissionDenied
    case outlookNotRunning
    case noAccessibleWindows

    // --- View / navigation ---
    case notInCalendarView
    case notInMailView
    case viewSwitchFailed(target: String)
    case navigationFailed(String)

    // --- Event table / parsing ---
    case eventsTableNotFound
    case eventRowNotFound(title: String)
    case ambiguousEvent(matches: Int)
    case detailWindowNotOpened
    case ambiguousWindow

    // --- Mail ---
    case messageListNotFound
    case inboxNotSelected
    case folderAmbiguous(name: String)
    case searchTimeout

    // --- Forms ---
    case formFieldNotFound(field: String)
    case formFieldWriteFailed(field: String)
    case sendBlockedByFieldFailure(reason: String)

    public var errorDescription: String? {
        switch self {
        case .accessibilityPermissionDenied:
            return "Accessibility permission is required. Grant access in System Settings > Privacy > Accessibility."
        case .outlookNotRunning:
            return "Microsoft Outlook is not running."
        case .noAccessibleWindows:
            return "Outlook is running but has no accessible windows."
        case .notInCalendarView:
            return "Outlook is not in calendar view."
        case .notInMailView:
            return "Outlook is not in mail view."
        case .viewSwitchFailed(let target):
            return "Could not switch Outlook to \(target) view."
        case .navigationFailed(let msg):
            return "Could not navigate Outlook: \(msg)"
        case .eventsTableNotFound:
            return "Calendar events table not found. Make sure Outlook is in list view."
        case .eventRowNotFound(let title):
            return "Could not find event '\(title)' in the calendar list."
        case .ambiguousEvent(let count):
            return "Event lookup returned \(count) matches; refine the search to get a unique result."
        case .detailWindowNotOpened:
            return "Outlook did not open the event detail window."
        case .ambiguousWindow:
            return "Multiple windows matched the expected identity; operation aborted."
        case .messageListNotFound:
            return "Message-list table not found."
        case .inboxNotSelected:
            return "Inbox is not the currently selected folder."
        case .folderAmbiguous(let name):
            return "Folder name '\(name)' is ambiguous across accounts; use --account to disambiguate."
        case .searchTimeout:
            return "Mail search timed out waiting for results to stabilize."
        case .formFieldNotFound(let field):
            return "Form field '\(field)' not found."
        case .formFieldWriteFailed(let field):
            return "Write to form field '\(field)' failed."
        case .sendBlockedByFieldFailure(let reason):
            return "Send/Save blocked: \(reason)"
        }
    }

    /// Machine-readable error code for JSON output.
    public var code: String {
        switch self {
        case .accessibilityPermissionDenied: return "accessibilityPermissionDenied"
        case .outlookNotRunning: return "outlookNotRunning"
        case .noAccessibleWindows: return "noAccessibleWindows"
        case .notInCalendarView: return "notInCalendarView"
        case .notInMailView: return "notInMailView"
        case .viewSwitchFailed: return "viewSwitchFailed"
        case .navigationFailed: return "navigationFailed"
        case .eventsTableNotFound: return "eventsTableNotFound"
        case .eventRowNotFound: return "eventRowNotFound"
        case .ambiguousEvent: return "ambiguousEvent"
        case .detailWindowNotOpened: return "detailWindowNotOpened"
        case .ambiguousWindow: return "ambiguousWindow"
        case .messageListNotFound: return "messageListNotFound"
        case .inboxNotSelected: return "inboxNotSelected"
        case .folderAmbiguous: return "folderAmbiguous"
        case .searchTimeout: return "searchTimeout"
        case .formFieldNotFound: return "formFieldNotFound"
        case .formFieldWriteFailed: return "formFieldWriteFailed"
        case .sendBlockedByFieldFailure: return "sendBlockedByFieldFailure"
        }
    }

    /// Recommended process exit code for CLI use.
    public var exitCode: Int32 {
        switch self {
        case .accessibilityPermissionDenied: return 77
        case .outlookNotRunning, .noAccessibleWindows,
             .notInCalendarView, .notInMailView,
             .eventsTableNotFound, .detailWindowNotOpened,
             .messageListNotFound, .inboxNotSelected: return 69
        default: return 1
        }
    }
}
