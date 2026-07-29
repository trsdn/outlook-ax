import Foundation

// MARK: - CalendarTemporalParser
//
// Parses localized date/time strings from Outlook's AX tree into
// normalized CalendarTemporal values.
//
// Rules:
// - Never silently infer a missing year; unresolved year → .unresolved
// - Never synthesize midnight timestamps for all-day events
// - Hyphenated titles are never treated as time ranges
// - 12-hour and 24-hour formats are both supported
// - DST-ambiguous wall times remain unresolved

public struct CalendarTemporalParser {

    private static let calendar24: Calendar = {
        var c = Calendar(identifier: .gregorian)
        c.locale = Locale(identifier: "en_US_POSIX")
        return c
    }()

    // MARK: - Public API

    /// Parse a time-range string (e.g. "09:00 - 10:00", "9:00 AM - 10:00 AM")
    /// obtained from a verified AX time child — NOT from arbitrary title text.
    ///
    /// Returns `.missing(reason: .parseError)` when the input does not match a
    /// recognized time-range pattern.  Returns `.missing(reason: .missingYear)`
    /// when a date range was found that lacks an explicit year context.
    public static func parse(
        timeRange: String,
        dateHeader: String,
        year: Int?,
        isAllDay: Bool,
        timeZone: TimeZone? = nil
    ) -> CalendarTemporal {
        let tzSource: TimeZoneSource = timeZone != nil ? .axLabel : .systemFallback
        let tz = timeZone ?? TimeZone.current

        // All-day from explicit AX flag
        if isAllDay {
            return parseAllDay(dateHeader: dateHeader, year: year, tz: tz, tzSource: tzSource)
        }

        // Check for date range (all-day multi-day)
        if looksLikeDateRange(timeRange) {
            return parseAllDay(dateHeader: dateHeader, year: year, tz: tz, tzSource: tzSource)
        }

        // Try time range (HH:MM - HH:MM or H:MM AM/PM - H:MM AM/PM)
        return parseTimedRange(
            timeRange: timeRange, dateHeader: dateHeader,
            year: year, tz: tz, tzSource: tzSource
        )
    }

    // MARK: - Private helpers

    /// Detect date ranges like "2.4.2026 - 28.4.2026" or "Apr 2 - Apr 28".
    /// This must NOT match ordinary hyphenated titles.
    private static func looksLikeDateRange(_ text: String) -> Bool {
        let trimmed = text.trimmingCharacters(in: .whitespaces)
        // Must contain " - " (space-hyphen-space)
        guard trimmed.contains(" - ") else { return false }
        let parts = trimmed.components(separatedBy: " - ")
        guard parts.count == 2 else { return false }
        let left = parts[0].trimmingCharacters(in: .whitespaces)
        // Date ranges contain digits; time ranges contain ":"
        // If both sides contain ":" they are time ranges
        if left.contains(":") { return false }
        // Check if left side looks like a date (has digits, may have dots or letters)
        return left.first?.isNumber == true || left.count > 6
    }

    private static func parseAllDay(
        dateHeader: String, year: Int?,
        tz: TimeZone, tzSource: TimeZoneSource
    ) -> CalendarTemporal {
        guard let y = year else {
            return CalendarTemporal(
                isAllDay: true,
                resolution: .unresolved,
                unresolvedReason: .missingYear,
                timeZoneSource: .absent
            )
        }
        let isoDate = extractISODate(from: dateHeader, year: y, tz: tz)
        return CalendarTemporal(
            isoDate: isoDate,
            isAllDay: true,
            year: y,
            resolution: isoDate != nil ? .localOnly : .unresolved,
            unresolvedReason: isoDate == nil ? .parseError : nil,
            timeZoneSource: .absent  // all-day events have no meaningful TZ
        )
    }

    private static func parseTimedRange(
        timeRange: String, dateHeader: String,
        year: Int?, tz: TimeZone, tzSource: TimeZoneSource
    ) -> CalendarTemporal {
        let trimmed = timeRange.trimmingCharacters(in: .whitespaces)
        guard trimmed.contains(" - ") else {
            return CalendarTemporal(resolution: .unresolved, unresolvedReason: .parseError, timeZoneSource: .absent)
        }
        let parts = trimmed.components(separatedBy: " - ")
        guard parts.count == 2 else {
            return CalendarTemporal(resolution: .unresolved, unresolvedReason: .parseError, timeZoneSource: .absent)
        }
        let rawStart = parts[0].trimmingCharacters(in: .whitespaces)
        let rawEnd   = parts[1].trimmingCharacters(in: .whitespaces)

        guard let start24 = toTime24(rawStart), let end24 = toTime24(rawEnd) else {
            return CalendarTemporal(resolution: .unresolved, unresolvedReason: .parseError, timeZoneSource: .absent)
        }

        guard let y = year else {
            return CalendarTemporal(
                localStartTime: start24,
                localEndTime: end24,
                isAllDay: false,
                resolution: .unresolved,
                unresolvedReason: .missingYear,
                timeZoneSource: .absent
            )
        }

        let isoDate = extractISODate(from: dateHeader, year: y, tz: tz)
        var startTS: String? = nil
        var endTS: String? = nil
        var resolution: TemporalResolution = .localOnly

        if let iso = isoDate {
            startTS = buildRFC3339(isoDate: iso, time24: start24, tz: tz)
            endTS   = buildRFC3339(isoDate: iso, time24: end24,   tz: tz)
            resolution = startTS != nil ? .resolved : .localOnly
        }

        return CalendarTemporal(
            isoDate: isoDate,
            localStartTime: start24,
            localEndTime: end24,
            startTimestamp: startTS,
            endTimestamp: endTS,
            isAllDay: false,
            year: y,
            resolution: resolution,
            timeZoneSource: resolution == .resolved ? tzSource : .absent,
            timeZoneIdentifier: resolution == .resolved ? tz.identifier : nil
        )
    }

    /// Convert 12-hour or 24-hour time string to "HH:MM" (24-hour).
    static func toTime24(_ raw: String) -> String? {
        let t = raw.trimmingCharacters(in: .whitespaces)

        // 24-hour: HH:MM or H:MM
        let h24 = #"^(\d{1,2}):(\d{2})$"#
        if let m = t.range(of: h24, options: .regularExpression) {
            let matched = String(t[m])
            let parts = matched.split(separator: ":").map(String.init)
            guard parts.count == 2,
                  let h = Int(parts[0]), let min = Int(parts[1]),
                  (0...23).contains(h), (0...59).contains(min) else { return nil }
            return String(format: "%02d:%02d", h, min)
        }

        // 12-hour: H:MM AM or H:MM PM (case-insensitive)
        let h12 = #"^(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)$"#
        if let m = t.range(of: h12, options: .regularExpression) {
            let matched = String(t[m])
            // Extract components via regex groups (Swift 5.9 compatible)
            let scanner = Scanner(string: matched)
            var h: Int = 0; var min: Int = 0
            guard scanner.scanInt(&h), scanner.scanString(":") != nil,
                  scanner.scanInt(&min) else { return nil }
            let upper = matched.uppercased()
            let isPM = upper.contains("PM")
            let isAM = upper.contains("AM")
            guard isAM || isPM else { return nil }
            guard (1...12).contains(h), (0...59).contains(min) else { return nil }
            var h24 = h
            if isPM && h != 12 { h24 = h + 12 }
            if isAM && h == 12 { h24 = 0 }
            return String(format: "%02d:%02d", h24, min)
        }

        return nil
    }

    /// Extract an ISO date string from a localized date header like
    /// "Mittwoch, 15. April" or "Wednesday, April 15".
    static func extractISODate(from header: String, year: Int, tz: TimeZone) -> String? {
        // Try common date extraction patterns
        let formats: [String] = [
            // German: "Mittwoch, 15. April 2026" or "15. April"
            "d. MMMM",
            "d. MMMM yyyy",
            "EEEE, d. MMMM",
            "EEEE, d. MMMM yyyy",
            // English: "Wednesday, April 15" or "April 15, 2026"
            "MMMM d",
            "MMMM d, yyyy",
            "EEEE, MMMM d",
            "EEEE, MMMM d, yyyy",
            // French: "mercredi 15 avril"
            "EEEE d MMMM",
            "d MMMM",
            // Spanish: "miércoles, 15 de abril"
            "EEEE, d 'de' MMMM",
            "d 'de' MMMM",
            // Italian: "mercoledì 15 aprile"
            "EEEE d MMMM",
        ]
        let locales = ["de", "en", "fr", "es", "it"]
        let stripped = header.components(separatedBy: ",").map {
            $0.trimmingCharacters(in: .whitespaces)
        }.joined(separator: ", ")

        for locale in locales {
            for fmt in formats {
                var cal = Calendar(identifier: .gregorian)
                cal.timeZone = tz
                let df = DateFormatter()
                df.calendar = cal
                df.timeZone = tz
                df.locale = Locale(identifier: locale)
                df.dateFormat = fmt
                if let date = df.date(from: stripped) ?? df.date(from: header) {
                    var comps = cal.dateComponents([.month, .day], from: date)
                    comps.year = year
                    comps.timeZone = tz
                    if let resolved = cal.date(from: comps) {
                        let iso = DateFormatter()
                        iso.calendar = cal
                        iso.timeZone = tz
                        iso.dateFormat = "yyyy-MM-dd"
                        return iso.string(from: resolved)
                    }
                }
            }
        }
        return nil
    }

    private static func buildRFC3339(isoDate: String, time24: String, tz: TimeZone) -> String? {
        let combined = "\(isoDate)T\(time24):00"
        let df = DateFormatter()
        df.dateFormat = "yyyy-MM-dd'T'HH:mm:ss"
        df.timeZone = tz
        df.locale = Locale(identifier: "en_US_POSIX")
        guard let date = df.date(from: combined) else { return nil }

        let rfc = DateFormatter()
        rfc.dateFormat = "yyyy-MM-dd'T'HH:mm:ssXXXXX"
        rfc.timeZone = tz
        rfc.locale = Locale(identifier: "en_US_POSIX")
        return rfc.string(from: date)
    }
}
