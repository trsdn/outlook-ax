import Foundation

// MARK: - Mail Models

/// A brief summary of a mail message as shown in the Inbox list.
public struct MailMessageSummary: Codable, Sendable, Equatable {
    public var subject: String
    public var from: String
    public var date: String
    public var preview: String

    public init(subject: String, from: String, date: String, preview: String) {
        self.subject = subject; self.from = from; self.date = date; self.preview = preview
    }
}

/// A complete mail message as read from the reading pane.
public struct MailMessage: Codable, Sendable, Equatable {
    public var subject: String
    public var from: String
    public var to: [String]
    public var date: String
    public var body: String

    public init(subject: String, from: String, to: [String], date: String, body: String) {
        self.subject = subject; self.from = from; self.to = to; self.date = date; self.body = body
    }
}

/// Identity of a mail folder.
public struct FolderIdentity: Codable, Sendable, Equatable {
    /// Display name of the folder.
    public var name: String
    /// The account this folder belongs to (e.g. "user@example.com").
    public var account: String
    /// Group path within the account (e.g. "Favorites", "All Accounts").
    public var group: String

    public init(name: String, account: String, group: String) {
        self.name = name; self.account = account; self.group = group
    }
}

/// Result of a mail search.
public enum MailSearchResult: Sendable {
    case results([MailMessageSummary])
    case noResults
    case timeout
}
