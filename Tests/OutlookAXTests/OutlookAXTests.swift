import XCTest
@testable import OutlookAX

final class OutlookAXTests: XCTestCase {

    // MARK: - Backward-compatibility: OutlookAX.CalendarEvent (library facade)

    func testCalendarEventRoundTripsAsJSON() throws {
        let event = OutlookAX.CalendarEvent(
            title: "Standup",
            date: "Dienstag, 21. April",
            start: "09:00",
            end: "09:30",
            isAllDay: false,
            myResponse: "accepted",
            organizer: "Alice",
            status: "Busy",
            calendar: ""
        )
        let data = try JSONEncoder().encode(event)
        let decoded = try JSONDecoder().decode(OutlookAX.CalendarEvent.self, from: data)
        XCTAssertEqual(event, decoded)
    }

    // MARK: - CalendarEvent with temporal field (production model)

    func testTopLevelCalendarEventRoundTripsAsJSON() throws {
        let event = CalendarEvent(
            title: "Team Meeting",
            date: "Wednesday, 15. April",
            start: "10:00",
            end: "11:00",
            isAllDay: false,
            myResponse: "accepted",
            organizer: "Bob",
            status: "Busy",
            calendar: "Work",
            categories: ["Important", "Team"]
        )
        let data = try JSONEncoder().encode(event)
        let decoded = try JSONDecoder().decode(CalendarEvent.self, from: data)
        XCTAssertEqual(event, decoded)
        XCTAssertEqual(decoded.categories, ["Important", "Team"])
        // temporal is nil → not present in JSON → still nil after decode
        XCTAssertNil(decoded.temporal)
    }

    func testCalendarEventTemporalFieldRoundTripsAsJSON() throws {
        // Build an event with a fully-resolved temporal value.
        let temporal = CalendarTemporal(
            isoDate: "2026-04-15",
            localStartTime: "09:00",
            localEndTime: "10:00",
            startTimestamp: "2026-04-15T09:00:00+02:00",
            endTimestamp: "2026-04-15T10:00:00+02:00",
            isAllDay: false,
            year: 2026,
            resolution: .resolved,
            timeZoneSource: .systemFallback,
            timeZoneIdentifier: "Europe/Berlin"
        )
        let event = CalendarEvent(
            title: "Standup", date: "Wednesday, April 15",
            start: "09:00", end: "10:00",
            isAllDay: false, myResponse: "accepted",
            organizer: "You", status: "Busy", calendar: "",
            temporal: temporal
        )
        let data = try JSONEncoder().encode(event)
        let decoded = try JSONDecoder().decode(CalendarEvent.self, from: data)
        XCTAssertEqual(decoded.temporal?.isoDate, "2026-04-15")
        XCTAssertEqual(decoded.temporal?.localStartTime, "09:00")
        XCTAssertEqual(decoded.temporal?.localEndTime, "10:00")
        XCTAssertEqual(decoded.temporal?.resolution, .resolved)
        XCTAssertEqual(decoded.temporal?.timeZoneIdentifier, "Europe/Berlin")
        XCTAssertFalse(decoded.temporal!.isAllDay)
    }

    func testCalendarEventBackwardCompatDecodeNilTemporal() throws {
        // JSON that predates the `temporal` field should decode with temporal == nil.
        let legacyJSON = """
        {
          "title": "Old Event",
          "date": "Monday, April 1",
          "start": "09:00",
          "end": "10:00",
          "isAllDay": false,
          "myResponse": "accepted",
          "organizer": "",
          "status": "Busy",
          "calendar": "",
          "categories": []
        }
        """.data(using: .utf8)!
        let decoded = try JSONDecoder().decode(CalendarEvent.self, from: legacyJSON)
        XCTAssertNil(decoded.temporal,
            "CalendarEvent decoded from legacy JSON (no `temporal` key) must have temporal == nil")
        XCTAssertEqual(decoded.title, "Old Event")
    }

    func testCalendarEventTemporalAppearsInJSONOutput() throws {
        // When temporal is non-nil, it must appear in the encoded JSON.
        let temporal = CalendarTemporalParser.parse(
            timeRange: "09:00 - 10:00",
            dateHeader: "Wednesday, April 15",
            year: 2026,
            isAllDay: false
        )
        let event = CalendarEvent(
            title: "Meeting", date: "Wednesday, April 15",
            start: "09:00", end: "10:00",
            isAllDay: false, myResponse: "accepted",
            organizer: "", status: "Busy", calendar: "",
            temporal: temporal
        )
        let data = try JSONEncoder().encode(event)
        let json = try JSONSerialization.jsonObject(with: data) as! [String: Any]
        XCTAssertNotNil(json["temporal"], "temporal field must appear in JSON when non-nil")
        let t = json["temporal"] as! [String: Any]
        XCTAssertEqual(t["localStartTime"] as? String, "09:00")
        XCTAssertEqual(t["localEndTime"] as? String, "10:00")
    }

    // MARK: - CalendarTemporal 12-hour and 24-hour value tests

    func testCalendarEventWith12HourTemporalValue() throws {
        // 12-hour input "9:00 AM - 10:30 AM" must normalize to 24-hour in temporal.
        let temporal = CalendarTemporalParser.parse(
            timeRange: "9:00 AM - 10:30 AM",
            dateHeader: "Wednesday, April 15",
            year: 2026,
            isAllDay: false
        )
        let event = CalendarEvent(
            title: "AM Meeting", date: "Wednesday, April 15",
            start: "9:00 AM", end: "10:30 AM",
            isAllDay: false, myResponse: "accepted",
            organizer: "", status: "Busy", calendar: "",
            temporal: temporal
        )
        // temporal must carry 24-hour normalized times
        XCTAssertEqual(event.temporal?.localStartTime, "09:00",
            "12-hour '9:00 AM' must normalize to '09:00' in temporal.localStartTime")
        XCTAssertEqual(event.temporal?.localEndTime, "10:30",
            "12-hour '10:30 AM' must normalize to '10:30' in temporal.localEndTime")
        XCTAssertFalse(event.temporal!.isAllDay)
        // JSON round-trip preserves the normalized times
        let data = try JSONEncoder().encode(event)
        let decoded = try JSONDecoder().decode(CalendarEvent.self, from: data)
        XCTAssertEqual(decoded.temporal?.localStartTime, "09:00")
    }

    func testCalendarEventWith24HourTemporalValue() throws {
        let temporal = CalendarTemporalParser.parse(
            timeRange: "14:30 - 15:45",
            dateHeader: "Wednesday, April 15",
            year: 2026,
            isAllDay: false
        )
        let event = CalendarEvent(
            title: "Afternoon Review", date: "Wednesday, April 15",
            start: "14:30", end: "15:45",
            isAllDay: false, myResponse: "accepted",
            organizer: "", status: "Busy", calendar: "",
            temporal: temporal
        )
        XCTAssertEqual(event.temporal?.localStartTime, "14:30")
        XCTAssertEqual(event.temporal?.localEndTime, "15:45")
    }

    func testCalendarEventWithUnresolvedTemporalSerializes() throws {
        // Events with unresolved temporal (e.g. missing year) must still serialize.
        let temporal = CalendarTemporal.missing(reason: .missingYear)
        let event = CalendarEvent(
            title: "Future Event", date: "Mittwoch, 15. April",
            start: "09:00", end: "10:00",
            isAllDay: false, myResponse: "accepted",
            organizer: "", status: "Busy", calendar: "",
            temporal: temporal
        )
        // Must not throw
        let data = try JSONEncoder().encode(event)
        let decoded = try JSONDecoder().decode(CalendarEvent.self, from: data)
        XCTAssertEqual(decoded.temporal?.resolution, .unresolved)
        XCTAssertEqual(decoded.temporal?.unresolvedReason, .missingYear)
        XCTAssertNil(decoded.temporal?.isoDate)
        XCTAssertNil(decoded.temporal?.startTimestamp)
    }

    // MARK: - filterAndSort (pure production-path tests, no AX required)

    /// Build a CalendarEvent with optional temporal for test use.
    private func makeEvent(
        title: String,
        date: String,
        start: String,
        end: String,
        isAllDay: Bool = false,
        temporal: CalendarTemporal? = nil
    ) -> CalendarEvent {
        CalendarEvent(
            title: title, date: date, start: start, end: end,
            isAllDay: isAllDay, myResponse: "accepted",
            organizer: "", status: "Busy", calendar: "",
            temporal: temporal
        )
    }

    /// Build a resolved temporal for test use.
    private func makeTemporal(isoDate: String, start24: String, end24: String) -> CalendarTemporal {
        CalendarTemporal(
            isoDate: isoDate,
            localStartTime: start24,
            localEndTime: end24,
            isAllDay: false,
            year: Int(isoDate.prefix(4)),
            resolution: .localOnly,
            timeZoneSource: .systemFallback
        )
    }

    func testFilterAndSortEmptyToday() {
        // When the table has no events at all, filterAndSort returns [].
        let result = CalendarService.filterAndSort([], toISO: "2026-04-15")
        XCTAssertTrue(result.isEmpty, "Empty input must produce empty output")
    }

    func testFilterAndSortTodayWithNoMatchingEvents() {
        // Events all belong to a different ISO date → today has no events → [].
        let tomorrow = makeTemporal(isoDate: "2026-04-16", start24: "09:00", end24: "10:00")
        let events = [
            makeEvent(title: "Tomorrow A", date: "Thu, Apr 16", start: "09:00", end: "10:00", temporal: tomorrow),
            makeEvent(title: "Tomorrow B", date: "Thu, Apr 16", start: "14:00", end: "15:00",
                      temporal: makeTemporal(isoDate: "2026-04-16", start24: "14:00", end24: "15:00"))
        ]
        let result = CalendarService.filterAndSort(events, toISO: "2026-04-15")
        XCTAssertTrue(result.isEmpty,
            "When all events resolve to a different ISO date, today must return empty (not fallback to wrong date)")
    }

    func testFilterAndSortMultipleDateHeaders() {
        // Table contains today's events followed by tomorrow's events.
        // Only today's events (2026-04-15) must be returned.
        let todayT1 = makeTemporal(isoDate: "2026-04-15", start24: "09:00", end24: "10:00")
        let todayT2 = makeTemporal(isoDate: "2026-04-15", start24: "14:00", end24: "15:00")
        let tomorrowT = makeTemporal(isoDate: "2026-04-16", start24: "09:00", end24: "10:00")

        let events = [
            makeEvent(title: "Today A",    date: "Wed, Apr 15", start: "09:00", end: "10:00", temporal: todayT1),
            makeEvent(title: "Today B",    date: "Wed, Apr 15", start: "14:00", end: "15:00", temporal: todayT2),
            makeEvent(title: "Tomorrow A", date: "Thu, Apr 16", start: "09:00", end: "10:00", temporal: tomorrowT),
        ]
        let result = CalendarService.filterAndSort(events, toISO: "2026-04-15")
        XCTAssertEqual(result.count, 2, "Only today's 2 events must be returned, not tomorrow's")
        XCTAssert(result.allSatisfy { $0.temporal?.isoDate == "2026-04-15" },
            "All returned events must have isoDate == today")
        XCTAssertFalse(result.contains { $0.title == "Tomorrow A" },
            "Tomorrow's events must be excluded")
    }

    func testFilterAndSortChronologicalOrder() {
        // Events with different start times must come out in ascending 24-hour order.
        let t1 = makeTemporal(isoDate: "2026-04-15", start24: "14:00", end24: "15:00")
        let t2 = makeTemporal(isoDate: "2026-04-15", start24: "09:00", end24: "10:00")
        let t3 = makeTemporal(isoDate: "2026-04-15", start24: "11:30", end24: "12:00")

        let events = [
            makeEvent(title: "Afternoon", date: "Wed, Apr 15", start: "14:00", end: "15:00", temporal: t1),
            makeEvent(title: "Morning",   date: "Wed, Apr 15", start: "09:00", end: "10:00", temporal: t2),
            makeEvent(title: "Midday",    date: "Wed, Apr 15", start: "11:30", end: "12:00", temporal: t3),
        ]
        let result = CalendarService.filterAndSort(events, toISO: "2026-04-15")
        XCTAssertEqual(result.count, 3)
        XCTAssertEqual(result[0].title, "Morning",   "First event must be 09:00 Morning")
        XCTAssertEqual(result[1].title, "Midday",    "Second event must be 11:30 Midday")
        XCTAssertEqual(result[2].title, "Afternoon", "Third event must be 14:00 Afternoon")
    }

    func testFilterAndSortAllDayEventsFirst() {
        // All-day events must sort before timed events regardless of title order.
        let allDayTemporal = CalendarTemporal(
            isoDate: "2026-04-15",
            isAllDay: true,
            year: 2026,
            resolution: .localOnly,
            timeZoneSource: .absent
        )
        let timedT = makeTemporal(isoDate: "2026-04-15", start24: "08:00", end24: "09:00")

        let events = [
            makeEvent(title: "Timed Early", date: "Wed, Apr 15", start: "08:00", end: "09:00",
                      isAllDay: false, temporal: timedT),
            makeEvent(title: "All Day",     date: "Wed, Apr 15", start: "", end: "",
                      isAllDay: true, temporal: allDayTemporal),
        ]
        let result = CalendarService.filterAndSort(events, toISO: "2026-04-15")
        XCTAssertEqual(result.count, 2)
        XCTAssertTrue(result[0].isAllDay, "All-day event must sort first")
        XCTAssertEqual(result[0].title, "All Day")
        XCTAssertFalse(result[1].isAllDay, "Timed event must sort after all-day")
    }

    func testFilterAndSortFallbackToFirstHeaderWhenNoISODate() {
        // When no event has an ISO date (all temporal.isoDate == nil),
        // fall back to the first date header group.
        let unresolvedTemporal = CalendarTemporal.missing(reason: .parseError)

        let events = [
            makeEvent(title: "Event A", date: "Mittwoch, 15. April", start: "09:00", end: "10:00",
                      temporal: unresolvedTemporal),
            makeEvent(title: "Event B", date: "Mittwoch, 15. April", start: "14:00", end: "15:00",
                      temporal: unresolvedTemporal),
            makeEvent(title: "Event C", date: "Donnerstag, 16. April", start: "09:00", end: "10:00",
                      temporal: unresolvedTemporal),
        ]
        // All isoDate == nil → fallback uses first date header "Mittwoch, 15. April"
        let result = CalendarService.filterAndSort(events, toISO: "2026-04-15")
        XCTAssertEqual(result.count, 2,
            "Fallback must return first date-header group (2 events), not third event")
        XCTAssert(result.allSatisfy { $0.date == "Mittwoch, 15. April" })
    }

    func testSortChronologicallyMixedTimes() {
        // sortChronologically is a pure function; test it directly.
        let t1 = makeTemporal(isoDate: "2026-04-15", start24: "16:00", end24: "17:00")
        let t2 = makeTemporal(isoDate: "2026-04-15", start24: "08:00", end24: "09:00")

        let events = [
            makeEvent(title: "Late",  date: "d", start: "16:00", end: "17:00", temporal: t1),
            makeEvent(title: "Early", date: "d", start: "08:00", end: "09:00", temporal: t2),
        ]
        let sorted = CalendarService.sortChronologically(events)
        XCTAssertEqual(sorted[0].title, "Early")
        XCTAssertEqual(sorted[1].title, "Late")
    }

    // MARK: - CalendarTemporal model tests

    func testCalendarTemporalUnresolvedMissingYear() {
        let t = CalendarTemporal.missing(reason: .missingYear)
        XCTAssertEqual(t.resolution, .unresolved)
        XCTAssertEqual(t.unresolvedReason, .missingYear)
        XCTAssertNil(t.isoDate)
        XCTAssertNil(t.startTimestamp)
    }

    func testCalendarTemporalAllDay() {
        let t = CalendarTemporal(
            isoDate: "2026-04-15",
            isAllDay: true,
            year: 2026,
            resolution: .localOnly,
            timeZoneSource: .absent
        )
        XCTAssertTrue(t.isAllDay)
        XCTAssertEqual(t.isoDate, "2026-04-15")
        XCTAssertNil(t.startTimestamp, "All-day events must not have synthetic timestamps")
    }

    // MARK: - OutlookAXError (top-level)

    func testOutlookAXErrorExitCodes() {
        XCTAssertEqual(OutlookAXError.accessibilityPermissionDenied.exitCode, 77)
        XCTAssertEqual(OutlookAXError.outlookNotRunning.exitCode, 69)
        XCTAssertEqual(OutlookAXError.noAccessibleWindows.exitCode, 69)
        XCTAssertEqual(OutlookAXError.formFieldNotFound(field: "subject").exitCode, 1)
    }

    func testOutlookAXErrorCodes() {
        XCTAssertEqual(OutlookAXError.accessibilityPermissionDenied.code, "accessibilityPermissionDenied")
        XCTAssertEqual(OutlookAXError.outlookNotRunning.code, "outlookNotRunning")
        XCTAssertEqual(OutlookAXError.folderAmbiguous(name: "Archive").code, "folderAmbiguous")
    }

    // MARK: - L10n shared catalog

    func testL10nContainsAllRequiredLocales() {
        // Every command-critical key must have variants for de, en, fr, es, it
        let requiredDE = "Posteingang"
        let requiredEN = "Inbox"
        let requiredFR = "Boîte de réception"
        let requiredES = "Bandeja de entrada"
        let requiredIT = "Posta in arrivo"
        XCTAssert(L10n.inboxWindow.contains(requiredDE), "de: Inbox label missing")
        XCTAssert(L10n.inboxWindow.contains(requiredEN), "en: Inbox label missing")
        XCTAssert(L10n.inboxWindow.contains(requiredFR), "fr: Inbox label missing")
        XCTAssert(L10n.inboxWindow.contains(requiredES), "es: Inbox label missing")
        XCTAssert(L10n.inboxWindow.contains(requiredIT), "it: Inbox label missing")
    }

    func testL10nHelpers() {
        XCTAssertTrue(L10n.equals("Inbox", L10n.inboxWindow))
        XCTAssertTrue(L10n.equals("Posteingang", L10n.inboxWindow))
        XCTAssertFalse(L10n.equals("Trash", L10n.inboxWindow))
        XCTAssertTrue(L10n.startsWith("Busy today", L10n.statusBusy))
        XCTAssertTrue(L10n.matches("show as Busy", L10n.statusBusy))
        XCTAssertTrue(L10n.endsWith("accepted.", L10n.respAccepted))
    }

    // MARK: - CalendarTemporalParser

    func test24HourTimeParsing() {
        XCTAssertEqual(CalendarTemporalParser.toTime24("09:00"), "09:00")
        XCTAssertEqual(CalendarTemporalParser.toTime24("9:30"), "09:30")
        XCTAssertEqual(CalendarTemporalParser.toTime24("16:45"), "16:45")
        XCTAssertNil(CalendarTemporalParser.toTime24("25:00"))
        XCTAssertNil(CalendarTemporalParser.toTime24("09:60"))
    }

    func test12HourTimeParsing() {
        XCTAssertEqual(CalendarTemporalParser.toTime24("9:00 AM"), "09:00")
        XCTAssertEqual(CalendarTemporalParser.toTime24("12:00 PM"), "12:00")
        XCTAssertEqual(CalendarTemporalParser.toTime24("12:00 AM"), "00:00")
        XCTAssertEqual(CalendarTemporalParser.toTime24("3:30 PM"), "15:30")
        XCTAssertEqual(CalendarTemporalParser.toTime24("11:59 pm"), "23:59")
    }

    func testHyphenatedTitleNotTreatedAsRange() {
        // Hyphenated titles like "Stand-up" must NOT be parsed as time ranges
        let result = CalendarTemporalParser.parse(
            timeRange: "Stand-up",
            dateHeader: "Wednesday, April 15",
            year: 2026,
            isAllDay: false
        )
        // "Stand-up" has no " - " so it should fail to parse as a range
        XCTAssertEqual(result.resolution, .unresolved)
    }

    func testMissingYearRemainsUnresolved() {
        let result = CalendarTemporalParser.parse(
            timeRange: "09:00 - 10:00",
            dateHeader: "Mittwoch, 15. April",
            year: nil,  // No year provided
            isAllDay: false
        )
        XCTAssertEqual(result.resolution, .unresolved)
        XCTAssertEqual(result.unresolvedReason, .missingYear)
        XCTAssertNil(result.startTimestamp, "Must not produce a timestamp when year is missing")
    }

    func testTimedEventWithYear() {
        let result = CalendarTemporalParser.parse(
            timeRange: "09:00 - 10:00",
            dateHeader: "Wednesday, April 15",
            year: 2026,
            isAllDay: false,
            timeZone: TimeZone(identifier: "Europe/Berlin")
        )
        XCTAssertEqual(result.localStartTime, "09:00")
        XCTAssertEqual(result.localEndTime, "10:00")
        XCTAssertNotNil(result.isoDate)
        XCTAssertFalse(result.isAllDay)
        // Resolution should be resolved since we have year and timezone
        XCTAssertTrue(result.resolution == .resolved || result.resolution == .localOnly)
    }

    func testAllDayEventNoTimestamp() {
        let result = CalendarTemporalParser.parse(
            timeRange: "",
            dateHeader: "Wednesday, April 15",
            year: 2026,
            isAllDay: true
        )
        XCTAssertTrue(result.isAllDay)
        XCTAssertNil(result.startTimestamp, "All-day events must not have RFC 3339 timestamps")
        XCTAssertNil(result.localStartTime)
    }

    func test12HourRangeParser() {
        let result = CalendarTemporalParser.parse(
            timeRange: "9:00 AM - 10:30 AM",
            dateHeader: "Wednesday, April 15",
            year: 2026,
            isAllDay: false
        )
        XCTAssertEqual(result.localStartTime, "09:00")
        XCTAssertEqual(result.localEndTime, "10:30")
        XCTAssertFalse(result.isAllDay)
    }
}
