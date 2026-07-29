import Foundation

// MARK: - Calendar View Mode

public enum CalendarViewMode: String, Codable, Sendable, CaseIterable {
    case day, workWeek, week, month, threeDay, list
}

// MARK: - Calendar Identity

/// Stable identity for a calendar entry.
public struct CalendarIdentity: Codable, Sendable, Equatable {
    /// Display name of the calendar.
    public var name: String
    /// Account the calendar belongs to.
    public var account: String
    /// Group name (e.g. "My Calendars", "Other Calendars").
    public var group: String
    /// Whether the calendar is currently shown/visible.
    public var isVisible: Bool

    public init(name: String, account: String, group: String, isVisible: Bool) {
        self.name = name; self.account = account; self.group = group; self.isVisible = isVisible
    }
}

// MARK: - Attendee

public struct Attendee: Codable, Sendable, Equatable {
    /// "organizer" | "required" | "optional"
    public var type: String
    /// "accepted" | "declined" | "tentative" | "none"
    public var response: String
    public var name: String

    public init(name: String, type: String, response: String) {
        self.name = name; self.type = type; self.response = response
    }
}

// MARK: - Calendar Event

/// A calendar event as listed in the calendar list view.
public struct CalendarEvent: Codable, Sendable, Equatable {
    public var title: String
    /// Localized date header as shown in list view (e.g. "Mittwoch, 15. April").
    /// Preserved for display and legacy matching; use `temporal.isoDate` for machine logic.
    public var date: String
    /// Localized start time or all-day label. Use `temporal.localStartTime` for machine logic.
    public var start: String
    /// Localized end time or date. Use `temporal.localEndTime` for machine logic.
    public var end: String
    public var isAllDay: Bool
    /// "accepted" | "declined" | "tentative" | "following"
    public var myResponse: String
    public var organizer: String
    /// "Busy" | "Free" | "Tentative" | "Out of Office" | "Working Elsewhere"
    public var status: String
    /// Verified calendar name. Empty when unknown (not populated from categories).
    public var calendar: String
    /// Category labels (separate from calendar identity).
    public var categories: [String]
    /// Normalized machine-readable temporal values: ISO date, 24-hour times, RFC 3339
    /// timestamps, year, resolution, and time-zone source. Nil only if temporal data was
    /// not available at parse time. Decoded as nil from JSON that predates this field
    /// (backward-compatible optional key).
    public var temporal: CalendarTemporal?

    public init(
        title: String, date: String, start: String, end: String,
        isAllDay: Bool, myResponse: String, organizer: String,
        status: String, calendar: String, categories: [String] = [],
        temporal: CalendarTemporal? = nil
    ) {
        self.title = title; self.date = date; self.start = start; self.end = end
        self.isAllDay = isAllDay; self.myResponse = myResponse; self.organizer = organizer
        self.status = status; self.calendar = calendar; self.categories = categories
        self.temporal = temporal
    }
}

// MARK: - Event Details

public struct EventDetails: Codable, Sendable, Equatable {
    public var attendees: [Attendee]
    public var location: String
    public var body: String
    /// Verified calendar identity. Empty when unknown.
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
