import Foundation

// MARK: - JSONEnvelope
//
// Typed JSON v2 success and error envelopes.
// Preserves native numbers and booleans instead of stringifying them.
//
// Migration rules:
//   v1.1: --json → v1 output; --json-version 2 → v2 output
//   v2.0: bare --json → v2 output; --json-version 1 → v1 output
//   v1 remains available through all 2.x releases

public struct JSONSuccess<T: Encodable>: Encodable {
    public let schemaVersion: Int
    public let ok: Bool
    public let command: String
    public let data: T

    public init(command: String, data: T) {
        self.schemaVersion = 2
        self.ok = true
        self.command = command
        self.data = data
    }
}

public struct JSONErrorDetails: Encodable {
    public let code: String
    public let message: String
    public let details: [String: String]

    public init(code: String, message: String, details: [String: String] = [:]) {
        self.code = code; self.message = message; self.details = details
    }
}

public struct JSONFailure: Encodable {
    public let schemaVersion: Int
    public let ok: Bool
    public let command: String
    public let error: JSONErrorDetails

    public init(command: String, error: JSONErrorDetails) {
        self.schemaVersion = 2
        self.ok = false
        self.command = command
        self.error = error
    }
}

// MARK: - v1 legacy envelope (string-value map)

public struct JSONLegacySuccess: Encodable {
    public let message: String
    public let extra: [String: String]

    public init(message: String, extra: [String: String]) {
        self.message = message; self.extra = extra
    }
}

// MARK: - Typed data payloads

public struct NotificationData: Codable {
    public let count: Int
    public init(count: Int) { self.count = count }
}

public struct StatusData: Codable {
    public let running: Bool
    public let window: String
    public let view: String
    public let notifications: Int
    public init(running: Bool, window: String, view: String, notifications: Int) {
        self.running = running; self.window = window; self.view = view; self.notifications = notifications
    }
}
