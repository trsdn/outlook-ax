import XCTest
@testable import OutlookAX

final class OutlookAXTests: XCTestCase {
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
}
