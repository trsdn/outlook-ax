import ApplicationServices
import Foundation

// MARK: - MenuWalker
//
// Walks Outlook's menu bar to trigger localized menu items by path.
// Each path element is an array of L10n label variants to try.
// Never calls exit() — returns Bool indicating success.

public enum MenuWalker {

    /// Trigger a menu item described by a path of L10n variant arrays.
    ///
    /// Example (View > Day):
    ///   `trigger(app: app, path: [L10n.menuView, L10n.viewDay])`
    ///
    /// Example (View > Timescale > 30 Minutes):
    ///   `trigger(app: app, path: [L10n.menuView, L10n.menuTimescale, ["30 Minuten", "30 Minutes"]])`
    ///
    /// - Returns: true when the deepest menu item was found and pressed.
    @discardableResult
    public static func trigger(app: AXUIElement, path: [[String]]) -> Bool {
        guard path.count >= 2 else { return false }

        var mbRef: CFTypeRef?
        guard AXUIElementCopyAttributeValue(app, kAXMenuBarAttribute as CFString, &mbRef) == .success,
              let menuBar = mbRef as! AXUIElement? else { return false }

        // Walk path[0] == top-level menu title
        for topItem in axChildren(menuBar) {
            guard axEqualsAny(axTitle(topItem), path[0]) else { continue }

            // Expand top-level menu
            AXUIElementPerformAction(topItem, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 0.15)

            let result = walkSubmenu(topItem, path: path, depth: 1)
            if !result {
                // Dismiss opened menu on failure
                AXUIElementPerformAction(topItem, kAXCancelAction as CFString)
            }
            return result
        }
        return false
    }

    // MARK: - Private

    private static func walkSubmenu(_ item: AXUIElement, path: [[String]], depth: Int) -> Bool {
        guard depth < path.count else { return false }

        let variants = path[depth]
        let isLeaf = depth == path.count - 1

        // Iterate through submenus and their children
        for submenu in axChildren(item) {
            for candidate in axChildren(submenu) {
                let t = axTitle(candidate)
                guard axEqualsAny(t, variants) else { continue }

                if isLeaf {
                    // This is the target menu item — press it.
                    AXUIElementPerformAction(candidate, kAXPressAction as CFString)
                    return true
                } else {
                    // Intermediate item: expand and recurse.
                    AXUIElementPerformAction(candidate, kAXPressAction as CFString)
                    Thread.sleep(forTimeInterval: 0.12)
                    if walkSubmenu(candidate, path: path, depth: depth + 1) { return true }
                }
            }
        }
        return false
    }
}
