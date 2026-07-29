import Foundation

// MARK: - Temporal Resolution

/// How confident we are in the resolved date/time value.
public enum TemporalResolution: String, Codable, Sendable {
    /// Both date and time-zone are known; RFC 3339 timestamp is available.
    case resolved
    /// Date and local time are known, but time-zone source is unconfirmed.
    /// RFC 3339 timestamp is omitted to avoid false precision.
    case localOnly
    /// Date or time could not be parsed completely (e.g. missing year).
    case unresolved
}

/// Reason an unresolved temporal value is unresolved.
public enum UnresolvedReason: String, Codable, Sendable {
    case missingYear
    case missingTime
    case ambiguousDST
    case parseError
}

// MARK: - Time Zone Source

/// Where the time zone used for resolution came from.
public enum TimeZoneSource: String, Codable, Sendable {
    /// Explicit time-zone label found in the AX tree (most reliable).
    case axLabel
    /// Inferred from the system's current time zone (fallback; may be wrong for remote events).
    case systemFallback
    /// No time-zone information was available.
    case absent
}

// MARK: - Calendar Display

/// Localized strings as shown in the Outlook UI (for display purposes only).
public struct CalendarDisplay: Codable, Sendable, Equatable {
    /// Localized date header, e.g. "Mittwoch, 15. April".
    public var dateHeader: String
    /// Localized start label, e.g. "09:00" or "ganztägig".
    public var startLabel: String
    /// Localized end label, e.g. "10:00".
    public var endLabel: String

    public init(dateHeader: String, startLabel: String, endLabel: String) {
        self.dateHeader = dateHeader; self.startLabel = startLabel; self.endLabel = endLabel
    }
}

// MARK: - Calendar Temporal

/// Normalized, machine-readable temporal values for a calendar event.
///
/// These values are separate from the localized display strings in `CalendarDisplay`.
/// A value of `.unresolved` means the data could not be determined reliably;
/// callers MUST NOT infer the missing value silently.
public struct CalendarTemporal: Codable, Sendable, Equatable {
    /// ISO 8601 date string (YYYY-MM-DD) for the event's start date, if resolved.
    public var isoDate: String?
    /// Local start time as "HH:MM" (24-hour), if available and not all-day.
    public var localStartTime: String?
    /// Local end time as "HH:MM" (24-hour), if available and not all-day.
    public var localEndTime: String?
    /// RFC 3339 / ISO 8601 timestamp for timed events with a known time zone.
    /// Absent for all-day events and `.localOnly` / `.unresolved` cases.
    public var startTimestamp: String?
    /// RFC 3339 end timestamp. For all-day events, this is the exclusive end date
    /// (start date + n days), not a synthetic midnight timestamp.
    public var endTimestamp: String?
    /// Whether this is an all-day event (no specific start/end time).
    public var isAllDay: Bool
    /// Explicit year. Nil means the year could not be determined.
    public var year: Int?
    /// How reliably the temporal values were resolved.
    public var resolution: TemporalResolution
    /// Why the value is unresolved (only set when resolution == .unresolved).
    public var unresolvedReason: UnresolvedReason?
    /// Where the time zone came from.
    public var timeZoneSource: TimeZoneSource
    /// IANA time-zone identifier (e.g. "Europe/Berlin"), if known.
    public var timeZoneIdentifier: String?

    public init(
        isoDate: String? = nil,
        localStartTime: String? = nil,
        localEndTime: String? = nil,
        startTimestamp: String? = nil,
        endTimestamp: String? = nil,
        isAllDay: Bool = false,
        year: Int? = nil,
        resolution: TemporalResolution = .unresolved,
        unresolvedReason: UnresolvedReason? = nil,
        timeZoneSource: TimeZoneSource = .absent,
        timeZoneIdentifier: String? = nil
    ) {
        self.isoDate = isoDate
        self.localStartTime = localStartTime
        self.localEndTime = localEndTime
        self.startTimestamp = startTimestamp
        self.endTimestamp = endTimestamp
        self.isAllDay = isAllDay
        self.year = year
        self.resolution = resolution
        self.unresolvedReason = unresolvedReason
        self.timeZoneSource = timeZoneSource
        self.timeZoneIdentifier = timeZoneIdentifier
    }

    /// Sentinel for an event where temporal data could not be parsed.
    public static func missing(reason: UnresolvedReason = .parseError) -> CalendarTemporal {
        CalendarTemporal(resolution: .unresolved, unresolvedReason: reason, timeZoneSource: .absent)
    }
}
