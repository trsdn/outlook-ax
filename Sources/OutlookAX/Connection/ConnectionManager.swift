import AppKit
import ApplicationServices
import Foundation

// MARK: - OutlookConnection

/// A resolved Accessibility connection to a running Outlook process.
public struct OutlookConnection {
    public let app: AXUIElement
    public let wins: [AXUIElement]
}

// MARK: - ConnectionManager

/// Provides typed connections to Microsoft Outlook with explicit connection policies.
///
/// Policy rules:
/// - `passiveConnect()`: For status/read commands. Never launches, activates, or unminimizes Outlook.
///   Throws `.outlookNotRunning` or `.noAccessibleWindows` when Outlook is unavailable.
/// - `interactiveConnect()`: For commands that require user interaction (compose, create, navigate).
///   Launches Outlook if not running. Brings Outlook to front so the UI is reachable.
///   Still throws `.accessibilityPermissionDenied` if AX access is unavailable.
public enum ConnectionManager {

    private static let bundleIDs = ["com.microsoft.Outlook", "com.microsoft.OneOutlook"]

    // MARK: - Passive (no launch, no activation)

    /// Returns a connection to Outlook without launching or activating it.
    /// For read-only and status commands.
    ///
    /// - Throws: `OutlookAXError.accessibilityPermissionDenied` when AX is not trusted.
    /// - Throws: `OutlookAXError.outlookNotRunning` when Outlook is not in the process list.
    /// - Throws: `OutlookAXError.noAccessibleWindows` when Outlook is running but has no windows.
    public static func passiveConnect() throws -> OutlookConnection {
        guard AXIsProcessTrusted() else { throw OutlookAXError.accessibilityPermissionDenied }
        guard let proc = runningOutlookProcess() else { throw OutlookAXError.outlookNotRunning }
        let app = AXUIElementCreateApplication(proc.processIdentifier)
        let wins = windowList(app)
        guard !wins.isEmpty else { throw OutlookAXError.noAccessibleWindows }
        return OutlookConnection(app: app, wins: wins)
    }

    /// Returns a connection if Outlook is already running and has windows; nil otherwise.
    /// Does NOT throw — for use in `.status` which reports state without failing.
    public static func tryPassiveConnect() -> OutlookConnection? {
        guard AXIsProcessTrusted() else { return nil }
        guard let proc = runningOutlookProcess() else { return nil }
        let app = AXUIElementCreateApplication(proc.processIdentifier)
        let wins = windowList(app)
        guard !wins.isEmpty else { return nil }
        return OutlookConnection(app: app, wins: wins)
    }

    // MARK: - Interactive (launch + activate)

    /// Returns a connection to Outlook, launching it if necessary and bringing it to front.
    /// For interactive commands (compose, create, navigate) that require UI visibility.
    ///
    /// - Throws: `OutlookAXError.accessibilityPermissionDenied` when AX is not trusted.
    /// - Throws: `OutlookAXError.outlookNotRunning` when Outlook cannot be found/launched.
    /// - Throws: `OutlookAXError.noAccessibleWindows` when Outlook has no accessible windows.
    public static func interactiveConnect() throws -> OutlookConnection {
        guard AXIsProcessTrusted() else { throw OutlookAXError.accessibilityPermissionDenied }

        let isRunning = runningOutlookProcess() != nil
        if !isRunning {
            try launchOutlook()
        }

        // Bring Outlook to front and unminimize its windows.
        if let proc = runningOutlookProcess() {
            proc.activate(options: [.activateIgnoringOtherApps])
            Thread.sleep(forTimeInterval: 0.5)
            let app = AXUIElementCreateApplication(proc.processIdentifier)
            // Unminimize windows that are minimized
            for win in windowList(app) {
                var minimized: CFTypeRef?
                if AXUIElementCopyAttributeValue(win, kAXMinimizedAttribute as CFString, &minimized) == .success,
                   (minimized as? Bool) == true {
                    AXUIElementSetAttributeValue(win, kAXMinimizedAttribute as CFString, false as CFTypeRef)
                    Thread.sleep(forTimeInterval: 0.2)
                }
            }
        }

        guard let proc = runningOutlookProcess() else { throw OutlookAXError.outlookNotRunning }
        let app = AXUIElementCreateApplication(proc.processIdentifier)
        let wins = windowList(app)
        guard !wins.isEmpty else { throw OutlookAXError.noAccessibleWindows }
        return OutlookConnection(app: app, wins: wins)
    }

    // MARK: - Window Refresh

    /// Re-read the window list for an existing app element (post-operation refresh).
    public static func refreshWindows(_ app: AXUIElement) -> [AXUIElement] {
        windowList(app)
    }

    // MARK: - Current View Detection

    /// Heuristically classify which Outlook view is active from window titles.
    public static func currentView(_ wins: [AXUIElement]) -> String {
        for w in wins {
            let t = axTitle(w)
            if axEqualsAny(t, L10n.calendarWindow) { return "calendar" }
            if axEqualsAny(t, L10n.inboxWindow) { return "mail" }
            if t.contains("Outlook") { return "mail" }
        }
        return "unknown"
    }

    // MARK: - Private

    private static func runningOutlookProcess() -> NSRunningApplication? {
        NSWorkspace.shared.runningApplications.first(where: {
            bundleIDs.contains($0.bundleIdentifier ?? "")
        })
    }

    private static func windowList(_ app: AXUIElement) -> [AXUIElement] {
        var ref: CFTypeRef?
        AXUIElementCopyAttributeValue(app, kAXWindowsAttribute as CFString, &ref)
        return ref as? [AXUIElement] ?? []
    }

    private static func launchOutlook() throws {
        let bundleID = "com.microsoft.Outlook"
        guard let url = NSWorkspace.shared.urlForApplication(withBundleIdentifier: bundleID) else {
            throw OutlookAXError.outlookNotRunning
        }
        let cfg = NSWorkspace.OpenConfiguration()
        cfg.activates = true
        let sema = DispatchSemaphore(value: 0)
        var launchError: Error?
        NSWorkspace.shared.openApplication(at: url, configuration: cfg) { _, err in
            launchError = err
            sema.signal()
        }
        _ = sema.wait(timeout: .now() + 12)
        if let err = launchError { throw err }
        Thread.sleep(forTimeInterval: 3.0)
    }
}

// MARK: - Semantic Window Selection

extension OutlookConnection {

    /// The main mail window: the first window whose title matches an inbox label,
    /// or the first window that looks like the primary Outlook window.
    /// Throws `.noAccessibleWindows` when no suitable window is found.
    public func mainMailWindow() throws -> AXUIElement {
        if let w = wins.first(where: { w in
            let t = axTitle(w)
            return axEqualsAny(t, L10n.inboxWindow) ||
                   (t.contains("Outlook") && !axMatchesAny(t, L10n.composeWindow))
        }) { return w }
        guard let w = wins.first else { throw OutlookAXError.noAccessibleWindows }
        return w
    }

    /// The main calendar window: the first window whose title matches a calendar label.
    /// Falls back to first window. Throws `.notInCalendarView` if none found.
    public func calendarWindow() throws -> AXUIElement {
        if let w = wins.first(where: { axEqualsAny(axTitle($0), L10n.calendarWindow) }) { return w }
        guard let w = wins.first else { throw OutlookAXError.noAccessibleWindows }
        return w
    }

    /// Identifies exactly one new window that appeared after an operation.
    /// Returns `.ambiguousWindow` when multiple new windows appeared.
    /// Returns nil when no new windows appeared.
    public func uniqueNewWindow(
        before priorTitles: Set<String>,
        matching predicate: (AXUIElement) -> Bool
    ) throws -> AXUIElement? {
        let after = ConnectionManager.refreshWindows(app)
        let newWins = after.filter { !priorTitles.contains(axTitle($0)) && predicate($0) }
        if newWins.count > 1 { throw OutlookAXError.ambiguousWindow }
        return newWins.first
    }
}
