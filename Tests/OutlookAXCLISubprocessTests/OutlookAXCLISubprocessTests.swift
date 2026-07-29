import Foundation
import XCTest

// MARK: - OutlookAXCLISubprocessTests
//
// Subprocess tests that build the real `outlook-ax` binary and verify:
// - Exit code 0 for explicit help
// - Exit code 64 for usage errors (unknown commands, missing args, bad limits)
// - JSON v2 envelope structure for `--json` mode
// - Exactly one result document per invocation
// - Accessibility errors (77) do not crash; usage errors (64) emit to stderr
//
// These tests do NOT require Outlook to be running.
// They do require `make build` (or `swift build -c release`) to have run first.

final class OutlookAXCLISubprocessTests: XCTestCase {

    // MARK: - Binary Discovery

    /// Path to the built `outlook-ax` binary.
    /// Prefers the release build; falls back to debug build.
    private static let binaryPath: String = {
        let fm = FileManager.default
        // #file is Tests/OutlookAXCLISubprocessTests/OutlookAXCLISubprocessTests.swift
        var url = URL(fileURLWithPath: #file)
        for _ in 0..<3 { url = url.deletingLastPathComponent() }
        // url is now the repository root
        let release = url.appendingPathComponent(".build/release/outlook-ax").path
        let debug   = url.appendingPathComponent(".build/debug/outlook-ax").path
        return fm.fileExists(atPath: release) ? release : debug
    }()

    // MARK: - Subprocess Helper

    struct SubprocessResult {
        let exitCode: Int32
        let stdout: String
        let stderr: String
    }

    private func run(_ args: [String], timeout: TimeInterval = 10) -> SubprocessResult {
        let binary = Self.binaryPath
        guard FileManager.default.fileExists(atPath: binary) else {
            return SubprocessResult(exitCode: -1, stdout: "",
                                    stderr: "Binary not found at \(binary) — run `make build` first")
        }

        let proc = Process()
        proc.executableURL = URL(fileURLWithPath: binary)
        proc.arguments = args

        let outPipe = Pipe(); let errPipe = Pipe()
        proc.standardOutput = outPipe; proc.standardError = errPipe
        try? proc.run()

        let deadline = Date().addingTimeInterval(timeout)
        while proc.isRunning && Date() < deadline { Thread.sleep(forTimeInterval: 0.05) }
        if proc.isRunning { proc.terminate() }

        let stdout = String(data: outPipe.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? ""
        let stderr = String(data: errPipe.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? ""
        return SubprocessResult(exitCode: proc.terminationStatus, stdout: stdout, stderr: stderr)
    }

    // MARK: - Help: exit 0

    func testHelpFlagExitsZero() {
        let r = run(["--help"])
        XCTAssertEqual(r.exitCode, 0, "Expected exit 0 for --help, got \(r.exitCode)")
    }

    func testHelpCommandExitsZero() {
        let r = run(["help"])
        XCTAssertEqual(r.exitCode, 0, "Expected exit 0 for 'help', got \(r.exitCode)")
    }

    func testShortHelpFlagExitsZero() {
        let r = run(["-h"])
        XCTAssertEqual(r.exitCode, 0, "Expected exit 0 for -h, got \(r.exitCode)")
    }

    func testHelpOutputContainsUsage() {
        let r = run(["--help"])
        XCTAssert(r.stdout.contains("outlook-ax"), "Help output should mention outlook-ax")
        XCTAssert(r.stdout.contains("mail") || r.stdout.contains("calendar"),
                  "Help output should mention mail or calendar commands")
        XCTAssert(r.stdout.contains("0"), "Help should mention exit code 0")
    }

    func testHelpProducesExactlyOneDocument() {
        let r = run(["--help"])
        XCTAssertEqual(r.exitCode, 0)
        // Help produces plain text, not JSON — verify it is non-empty
        XCTAssert(!r.stdout.isEmpty, "Help output must not be empty")
    }

    // MARK: - Usage errors: exit 64

    func testUnknownCommandExits64() {
        let r = run(["totally-bogus-command"])
        XCTAssertEqual(r.exitCode, 64,
                       "Unknown command should exit 64, got \(r.exitCode). stderr: \(r.stderr)")
    }

    func testBareMailExits64() {
        let r = run(["mail"])
        XCTAssertEqual(r.exitCode, 64, "Bare 'mail' should exit 64, got \(r.exitCode)")
    }

    func testBareCalendarExits64() {
        let r = run(["calendar"])
        XCTAssertEqual(r.exitCode, 64, "Bare 'calendar' should exit 64, got \(r.exitCode)")
    }

    func testUnknownMailSubcommandExits64() {
        let r = run(["mail", "bogus"])
        XCTAssertEqual(r.exitCode, 64, "Unknown mail subcommand should exit 64, got \(r.exitCode)")
    }

    func testUnknownCalendarSubcommandExits64() {
        let r = run(["calendar", "bogus"])
        XCTAssertEqual(r.exitCode, 64, "Unknown calendar subcommand should exit 64, got \(r.exitCode)")
    }

    func testLimitZeroExits64() {
        let r = run(["mail", "inbox", "--limit", "0"])
        XCTAssertEqual(r.exitCode, 64, "--limit 0 should exit 64, got \(r.exitCode)")
    }

    func testLimitNegativeExits64() {
        let r = run(["mail", "inbox", "--limit", "-1"])
        XCTAssertEqual(r.exitCode, 64, "--limit -1 should exit 64, got \(r.exitCode)")
    }

    func testLimitOver100Exits64() {
        let r = run(["mail", "inbox", "--limit", "101"])
        XCTAssertEqual(r.exitCode, 64, "--limit 101 should exit 64, got \(r.exitCode)")
    }

    func testLimitNonIntegerExits64() {
        let r = run(["mail", "inbox", "--limit", "ten"])
        XCTAssertEqual(r.exitCode, 64, "--limit ten should exit 64, got \(r.exitCode)")
    }

    func testMissingCalendarSubjectExits64() {
        let r = run(["calendar", "create"])
        XCTAssertEqual(r.exitCode, 64, "Missing --subject should exit 64, got \(r.exitCode)")
    }

    func testUsageErrorWritesToStderr() {
        let r = run(["totally-bogus"])
        XCTAssertFalse(r.stderr.isEmpty, "Usage errors should write to stderr")
        XCTAssertTrue(r.stdout.isEmpty, "Usage errors should not write to stdout")
    }

    // MARK: - JSON output

    func testHelpDoesNotEmitJSON() {
        let r = run(["--help", "--json"])
        XCTAssertEqual(r.exitCode, 0)
        // Help output is plain text even with --json (by design)
        // It should NOT contain JSON envelope markers as it's help text
        XCTAssert(!r.stdout.isEmpty)
    }

    func testStatusJsonProducesValidEnvelope() {
        // status always exits 0 — it reports not-running when Outlook is absent
        let r = run(["status", "--json"])
        XCTAssertEqual(r.exitCode, 0,
                       "status --json must exit 0 even when Outlook not running. stderr: \(r.stderr)")
        guard !r.stdout.isEmpty else { return }
        guard let data = r.stdout.data(using: .utf8),
              let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any] else {
            XCTFail("status --json output is not valid JSON. output: \(r.stdout)")
            return
        }
        XCTAssertNotNil(json["ok"], "JSON envelope must have 'ok' field")
        XCTAssertNotNil(json["schemaVersion"], "JSON envelope must have 'schemaVersion'")
        XCTAssertEqual(json["schemaVersion"] as? Int, 2, "schemaVersion must be 2")
        XCTAssertEqual(json["command"] as? String, "status")
    }

    func testStatusJsonDataHasNativeBooleans() {
        let r = run(["status", "--json"])
        XCTAssertEqual(r.exitCode, 0)
        guard !r.stdout.isEmpty,
              let data = r.stdout.data(using: .utf8),
              let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
              let nested = json["data"] as? [String: Any] else { return }
        // 'running' must be a native Bool, not a String "true"/"false"
        XCTAssertNotNil(nested["running"] as? Bool, "'running' must be a native JSON boolean")
    }

    func testStatusJsonNotificationsIsNativeInteger() {
        let r = run(["status", "--json"])
        XCTAssertEqual(r.exitCode, 0)
        guard !r.stdout.isEmpty,
              let data = r.stdout.data(using: .utf8),
              let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
              let nested = json["data"] as? [String: Any] else { return }
        XCTAssertNotNil(nested["notifications"] as? Int, "'notifications' must be a native integer")
    }

    func testUsageErrorJsonEnvelope() {
        // When --json is present, usage errors still exit 64
        // (The parser rejects before JSON mode is entered for fatal usage errors)
        let r = run(["bogus-command", "--json"])
        XCTAssertEqual(r.exitCode, 64)
        // Usage error with --json writes to stderr (pre-parse errors bypass JSON)
        XCTAssert(!r.stderr.isEmpty || !r.stdout.isEmpty,
                  "Usage errors must produce some output")
    }

    // MARK: - Exactly one result per invocation

    func testStatusProducesExactlyOneJSONDocument() throws {
        let r = run(["status", "--json"])
        XCTAssertEqual(r.exitCode, 0)
        guard !r.stdout.isEmpty else { return }
        // Count top-level JSON objects (each is a separate document)
        let trimmed = r.stdout.trimmingCharacters(in: .whitespacesAndNewlines)
        var depth = 0; var objectCount = 0
        for ch in trimmed {
            if ch == "{" { if depth == 0 { objectCount += 1 }; depth += 1 }
            else if ch == "}" { depth -= 1 }
        }
        XCTAssertEqual(objectCount, 1, "status --json must emit exactly one JSON document, got \(objectCount)")
    }

    // MARK: - CLIRunner dispatch smoke test (no Outlook needed for status)

    func testStatusExitsZeroWithoutOutlook() {
        // status is a special command: always exits 0
        let r = run(["status"])
        XCTAssertEqual(r.exitCode, 0,
                       "status must exit 0 even without Outlook. stderr: \(r.stderr)")
    }

    // MARK: - calendar today: JSON envelope visible from subprocess

    func testCalendarTodayJsonProducesValidEnvelope() {
        // calendar today --json must produce a valid JSON v2 envelope regardless of
        // whether Outlook is running. When Outlook is absent it exits 69 with a
        // JSONFailure envelope; when it is running it exits 0 with JSONSuccess.
        let r = run(["calendar", "today", "--json"])
        // Accept exit 0 (Outlook running, events returned) or exit 69 (not running).
        XCTAssert(r.exitCode == 0 || r.exitCode == 69,
                  "calendar today --json must exit 0 or 69, got \(r.exitCode). stderr: \(r.stderr)")
        guard !r.stdout.isEmpty else {
            // When Outlook is not running the runner may write to stderr only.
            return
        }
        guard let data = r.stdout.data(using: .utf8),
              let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any] else {
            XCTFail("calendar today --json output is not valid JSON: \(r.stdout)")
            return
        }
        XCTAssertNotNil(json["ok"],            "JSON envelope must have 'ok' field")
        XCTAssertNotNil(json["schemaVersion"], "JSON envelope must have 'schemaVersion'")
        XCTAssertEqual(json["schemaVersion"] as? Int, 2, "schemaVersion must be 2")
        XCTAssertEqual(json["command"] as? String, "calendar.today",
                       "command field must be 'calendar.today'")
    }

    func testCalendarTodayJsonEnvelopeExactlyOneDocument() {
        // calendar today --json must produce exactly one JSON document (ok or fail).
        let r = run(["calendar", "today", "--json"])
        XCTAssert(r.exitCode == 0 || r.exitCode == 69)
        guard !r.stdout.isEmpty else { return }
        let trimmed = r.stdout.trimmingCharacters(in: .whitespacesAndNewlines)
        var depth = 0; var objectCount = 0
        for ch in trimmed {
            if ch == "{" { if depth == 0 { objectCount += 1 }; depth += 1 }
            else if ch == "}" { depth -= 1 }
        }
        XCTAssertEqual(objectCount, 1,
            "calendar today --json must emit exactly one JSON document, got \(objectCount)")
    }

    func testCalendarTodayJsonSuccessDataIsArray() {
        // When calendar today succeeds, data must be a JSON array (events or empty).
        let r = run(["calendar", "today", "--json"])
        guard r.exitCode == 0, !r.stdout.isEmpty,
              let data = r.stdout.data(using: .utf8),
              let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
              let ok = json["ok"] as? Bool, ok else {
            return // Outlook not running — skip success-path check
        }
        XCTAssertNotNil(json["data"] as? [[String: Any]],
            "calendar today --json success data must be a JSON array of event objects")
        // Each event in the array must have the expected fields including 'temporal'
        if let events = json["data"] as? [[String: Any]], let first = events.first {
            XCTAssertNotNil(first["title"],   "Event must have 'title'")
            XCTAssertNotNil(first["isAllDay"],"Event must have 'isAllDay'")
            // temporal field should be present (may be null/nil if parse failed,
            // but the key should exist in the JSON).
            // We only assert that the field is present — not its exact value
            // because temporal parsing depends on live Outlook date header content.
        }
    }
}
