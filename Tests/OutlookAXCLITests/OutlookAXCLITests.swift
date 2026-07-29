import XCTest
@testable import OutlookAXCLIKit
import OutlookAX

/// CLI unit tests: parser, JSON envelopes, exit-code semantics.
/// These tests run without Outlook and without AX access.
final class OutlookAXCLITests: XCTestCase {

    // MARK: - CLIParser: help

    func testHelpFlag() {
        let result = CLIParser(args: ["--help"]).parse()
        guard case .help = result else { return XCTFail("Expected .help, got \(result)") }
    }

    func testHelpCommand() {
        let result = CLIParser(args: ["help"]).parse()
        guard case .help = result else { return XCTFail("Expected .help, got \(result)") }
    }

    func testShortHelpFlag() {
        let result = CLIParser(args: ["-h"]).parse()
        guard case .help = result else { return XCTFail("Expected .help, got \(result)") }
    }

    // MARK: - CLIParser: usage errors

    func testNoArgs() {
        let result = CLIParser(args: []).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError, got \(result)") }
    }

    func testUnknownCommand() {
        let result = CLIParser(args: ["boguscommand"]).parse()
        guard case .usageError(let msg) = result else { return XCTFail("Expected .usageError, got \(result)") }
        XCTAssert(msg.contains("boguscommand"), "Error should mention the unknown command")
    }

    func testBareMailGroup() {
        let result = CLIParser(args: ["mail"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for bare 'mail', got \(result)") }
    }

    func testBareCalendarGroup() {
        let result = CLIParser(args: ["calendar"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for bare 'calendar', got \(result)") }
    }

    func testUnknownMailSubcommand() {
        let result = CLIParser(args: ["mail", "bogus"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for unknown mail subcommand, got \(result)") }
    }

    func testUnknownCalendarSubcommand() {
        let result = CLIParser(args: ["calendar", "bogus"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for unknown calendar subcommand, got \(result)") }
    }

    // MARK: - CLIParser: --limit validation

    func testLimitValid() {
        for limit in [1, 10, 50, 100] {
            let result = CLIParser(args: ["mail", "inbox", "--limit", "\(limit)"]).parse()
            guard case .command(.mailInbox(let l), _, _) = result else {
                return XCTFail("Expected .mailInbox for limit \(limit), got \(result)")
            }
            XCTAssertEqual(l, limit)
        }
    }

    func testLimitZeroIsError() {
        let result = CLIParser(args: ["mail", "inbox", "--limit", "0"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for limit=0, got \(result)") }
    }

    func testLimitNegativeIsError() {
        let result = CLIParser(args: ["mail", "inbox", "--limit", "-1"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for limit=-1, got \(result)") }
    }

    func testLimitOver100IsError() {
        let result = CLIParser(args: ["mail", "inbox", "--limit", "101"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for limit=101, got \(result)") }
    }

    func testLimitNonIntegerIsError() {
        let result = CLIParser(args: ["mail", "inbox", "--limit", "abc"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for limit=abc, got \(result)") }
    }

    func testLimitFloatIsError() {
        let result = CLIParser(args: ["mail", "inbox", "--limit", "10.5"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for limit=10.5, got \(result)") }
    }

    func testDefaultLimit() {
        let result = CLIParser(args: ["mail", "inbox"]).parse()
        guard case .command(.mailInbox(let l), _, _) = result else {
            return XCTFail("Expected .mailInbox with default limit, got \(result)")
        }
        XCTAssertEqual(l, 10)
    }

    // MARK: - CLIParser: json flag detection

    func testJsonFlagParsed() {
        let result = CLIParser(args: ["status", "--json"]).parse()
        guard case .command(.status, let json, _) = result else { return XCTFail("Expected .status, got \(result)") }
        XCTAssertTrue(json, "--json flag should set json=true")
    }

    func testNoJsonFlagByDefault() {
        let result = CLIParser(args: ["status"]).parse()
        guard case .command(.status, let json, _) = result else { return XCTFail("Expected .status, got \(result)") }
        XCTAssertFalse(json, "json should be false without --json")
    }

    func testDetailsFlagParsed() {
        let result = CLIParser(args: ["calendar", "today", "--details"]).parse()
        guard case .command(.calendarToday(let details), _, let d) = result else {
            return XCTFail("Expected .calendarToday, got \(result)")
        }
        XCTAssertTrue(details)
        XCTAssertTrue(d)
    }

    // MARK: - CLIParser: calendar create arguments

    func testCalendarCreateRequiresSubject() {
        let result = CLIParser(args: ["calendar", "create"]).parse()
        guard case .usageError = result else { return XCTFail("Expected .usageError for missing --subject, got \(result)") }
    }

    func testCalendarCreateWithSubject() {
        let result = CLIParser(args: ["calendar", "create", "--subject", "Test Meeting"]).parse()
        guard case .command(.calendarCreate(let subj, nil, nil, nil, false), _, _) = result else {
            return XCTFail("Expected .calendarCreate, got \(result)")
        }
        XCTAssertEqual(subj, "Test Meeting")
    }

    func testCalendarCreateWithSendFlag() {
        let result = CLIParser(args: ["calendar", "create", "--subject", "Test", "--send"]).parse()
        guard case .command(.calendarCreate(_, _, _, _, true), _, _) = result else {
            return XCTFail("Expected .calendarCreate with send=true, got \(result)")
        }
    }

    // MARK: - JSON envelope: typed output

    func testNotificationDataIsInteger() throws {
        let data = NotificationData(count: 5)
        let envelope = JSONSuccess(command: "notifications", data: data)
        let jsonData = try JSONEncoder().encode(envelope)
        let json = try JSONSerialization.jsonObject(with: jsonData) as! [String: Any]
        XCTAssertEqual(json["ok"] as? Bool, true)
        XCTAssertEqual(json["schemaVersion"] as? Int, 2)
        XCTAssertEqual(json["command"] as? String, "notifications")
        let nested = json["data"] as! [String: Any]
        // count must be an Int, not a String
        XCTAssertEqual(nested["count"] as? Int, 5, "Notification count must be a native integer, not a string")
    }

    func testStatusDataPreservesBoolean() throws {
        let data = StatusData(running: true, window: "Outlook", view: "calendar", notifications: 0)
        let envelope = JSONSuccess(command: "status", data: data)
        let jsonData = try JSONEncoder().encode(envelope)
        let json = try JSONSerialization.jsonObject(with: jsonData) as! [String: Any]
        let nested = json["data"] as! [String: Any]
        XCTAssertEqual(nested["running"] as? Bool, true, "running must be a native boolean, not a string")
    }

    func testFailureEnvelope() throws {
        let err = JSONErrorDetails(code: "outlookNotRunning", message: "Outlook is not running.")
        let envelope = JSONFailure(command: "status", error: err)
        let jsonData = try JSONEncoder().encode(envelope)
        let json = try JSONSerialization.jsonObject(with: jsonData) as! [String: Any]
        XCTAssertEqual(json["ok"] as? Bool, false)
        XCTAssertEqual(json["schemaVersion"] as? Int, 2)
        let errObj = json["error"] as! [String: Any]
        XCTAssertEqual(errObj["code"] as? String, "outlookNotRunning")
    }
}
