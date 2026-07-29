import AppKit
import ApplicationServices
import Foundation

// MARK: - AX Element Attribute Readers
// Internal to the OutlookAX module — not exported to consumers.

func axRole(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXRoleAttribute as CFString, &r)
    return r as? String ?? ""
}

func axTitle(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXTitleAttribute as CFString, &r)
    return r as? String ?? ""
}

func axDesc(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXDescriptionAttribute as CFString, &r)
    return r as? String ?? ""
}

func axValue(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXValueAttribute as CFString, &r)
    return r as? String ?? ""
}

func axChildren(_ e: AXUIElement) -> [AXUIElement] {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXChildrenAttribute as CFString, &r)
    return r as? [AXUIElement] ?? []
}

func axSubrole(_ e: AXUIElement) -> String {
    var r: CFTypeRef?
    AXUIElementCopyAttributeValue(e, kAXSubroleAttribute as CFString, &r)
    return r as? String ?? ""
}

func axPosition(_ e: AXUIElement) -> CGPoint? {
    var r: CFTypeRef?
    guard AXUIElementCopyAttributeValue(e, "AXPosition" as CFString, &r) == .success else { return nil }
    var pt = CGPoint.zero
    AXValueGetValue(r as! AXValue, .cgPoint, &pt)
    return pt
}

func axSize(_ e: AXUIElement) -> CGSize? {
    var r: CFTypeRef?
    guard AXUIElementCopyAttributeValue(e, "AXSize" as CFString, &r) == .success else { return nil }
    var sz = CGSize.zero
    AXValueGetValue(r as! AXValue, .cgSize, &sz)
    return sz
}

// MARK: - AX Tree Search

/// DFS search for the first element matching the predicate.
func axFind(
    _ root: AXUIElement,
    where predicate: (AXUIElement) -> Bool,
    depth: Int = 0,
    maxDepth: Int = 14
) -> AXUIElement? {
    if predicate(root) { return root }
    if depth >= maxDepth { return nil }
    for child in axChildren(root) {
        if let found = axFind(child, where: predicate, depth: depth + 1, maxDepth: maxDepth) {
            return found
        }
    }
    return nil
}

/// DFS search collecting all elements matching the predicate.
func axFindAll(
    _ root: AXUIElement,
    where predicate: (AXUIElement) -> Bool,
    depth: Int = 0,
    maxDepth: Int = 12
) -> [AXUIElement] {
    var results: [AXUIElement] = []
    if predicate(root) { results.append(root) }
    if depth >= maxDepth { return results }
    for child in axChildren(root) {
        results += axFindAll(child, where: predicate, depth: depth + 1, maxDepth: maxDepth)
    }
    return results
}

// MARK: - Text Extraction

/// Recursively collect visible text from an AX subtree.
func axCollectText(_ elem: AXUIElement, into parts: inout [String], depth: Int = 0, maxDepth: Int = 8) {
    if depth > maxDepth { return }
    let role = axRole(elem)
    if role == "AXStaticText" {
        let v = axValue(elem)
        if !v.isEmpty { parts.append(v) }
        return
    }
    if role == "AXLink" {
        let t = axTitle(elem); let d = axDesc(elem)
        if d.hasPrefix("http") && !t.isEmpty { parts.append("[\(t)](\(d))"); return }
    }
    for child in axChildren(elem) { axCollectText(child, into: &parts, depth: depth + 1, maxDepth: maxDepth) }
}

// MARK: - Button Helpers

/// Press a button whose `kAXDescriptionAttribute` starts with the given prefix.
@discardableResult
func axPressButtonByDescPrefix(_ win: AXUIElement, prefix: String) -> Bool {
    guard let btn = axFind(win, where: { axDesc($0).hasPrefix(prefix) && axRole($0) == "AXButton" }) else { return false }
    return AXUIElementPerformAction(btn, kAXPressAction as CFString) == .success
}

/// Press a button whose `kAXTitleAttribute` exactly matches the given title.
@discardableResult
func axPressButtonByTitle(_ win: AXUIElement, title: String) -> Bool {
    guard let btn = axFind(win, where: { axTitle($0) == title && axRole($0) == "AXButton" }) else { return false }
    return AXUIElementPerformAction(btn, kAXPressAction as CFString) == .success
}

/// Try each prefix in `prefixes` in order; return true on first success.
@discardableResult
func axPressButtonAny(_ win: AXUIElement, prefixes: [String]) -> Bool {
    prefixes.contains(where: { axPressButtonByDescPrefix(win, prefix: $0) })
}

/// Close a window through its AXCloseButton.
func axCloseWindow(_ win: AXUIElement) {
    for child in axChildren(win) where axRole(child) == "AXButton" {
        if axSubrole(child) == "AXCloseButton" {
            AXUIElementPerformAction(child, kAXPressAction as CFString)
            Thread.sleep(forTimeInterval: 0.25)
            return
        }
    }
}

// MARK: - Keyboard / Mouse Simulation

/// Type a Unicode string character by character.
func axTypeText(_ text: String) {
    for char in text {
        let str = String(char)
        let src = CGEventSource(stateID: .hidSystemState)
        if let keyDown = CGEvent(keyboardEventSource: src, virtualKey: 0, keyDown: true) {
            let utf16 = Array(str.utf16)
            keyDown.keyboardSetUnicodeString(stringLength: utf16.count, unicodeString: utf16)
            keyDown.post(tap: .cghidEventTap)
        }
        if let keyUp = CGEvent(keyboardEventSource: src, virtualKey: 0, keyDown: false) {
            keyUp.post(tap: .cghidEventTap)
        }
        Thread.sleep(forTimeInterval: 0.03)
    }
}

func axPressKey(_ keyCode: CGKeyCode, flags: CGEventFlags = []) {
    let src = CGEventSource(stateID: .hidSystemState)
    if let dn = CGEvent(keyboardEventSource: src, virtualKey: keyCode, keyDown: true) {
        dn.flags = flags; dn.post(tap: .cghidEventTap)
    }
    if let up = CGEvent(keyboardEventSource: src, virtualKey: keyCode, keyDown: false) {
        up.flags = flags; up.post(tap: .cghidEventTap)
    }
    Thread.sleep(forTimeInterval: 0.05)
}

func axPressTab()    { axPressKey(48) }
func axPressReturn() { axPressKey(36) }
func axSelectAll()   { axPressKey(0, flags: .maskCommand) }

/// Double-click at the centre of an AX element.
func axDoubleClick(at point: CGPoint) {
    let src = CGEventSource(stateID: .hidSystemState)
    for clickState: Int64 in [1, 2] {
        let typeDown = CGEventType.leftMouseDown; let typeUp = CGEventType.leftMouseUp
        if let dn = CGEvent(mouseEventSource: src, mouseType: typeDown, mouseCursorPosition: point, mouseButton: .left) {
            dn.setIntegerValueField(.mouseEventClickState, value: clickState); dn.post(tap: .cghidEventTap)
        }
        if let up = CGEvent(mouseEventSource: src, mouseType: typeUp, mouseCursorPosition: point, mouseButton: .left) {
            up.setIntegerValueField(.mouseEventClickState, value: clickState); up.post(tap: .cghidEventTap)
        }
        if clickState == 1 { Thread.sleep(forTimeInterval: 0.05) }
    }
}

// MARK: - L10n Matching Helpers
// These delegate to the shared L10n catalog helpers.

/// True if `text` equals any variant in `variants` (case-sensitive exact match).
func axEqualsAny(_ text: String, _ variants: [String]) -> Bool {
    L10n.equals(text, variants)
}

/// True if `text` starts with any variant in `variants`.
func axStartsWithAny(_ text: String, _ variants: [String]) -> Bool {
    L10n.startsWith(text, variants)
}

/// True if `text` contains any variant in `variants`.
func axMatchesAny(_ text: String, _ variants: [String]) -> Bool {
    L10n.matches(text, variants)
}

/// True if `text` ends with any variant in `variants`.
func axEndsWithAny(_ text: String, _ variants: [String]) -> Bool {
    L10n.endsWith(text, variants)
}

// MARK: - String Helpers

/// Lowercase and trim whitespace for stable title matching.
func axNormalise(_ s: String) -> String {
    s.lowercased().trimmingCharacters(in: .whitespacesAndNewlines)
}

/// Strip "Declined: " / "Following: " response prefixes from event titles.
func axStripResponsePrefix(_ raw: String) -> String {
    var s = raw
    if axStartsWithAny(s, L10n.declinedPrefix) {
        s = String(s.drop(while: { $0 != ":" }).dropFirst(2))
    } else if axStartsWithAny(s, L10n.followingPrefix) {
        s = String(s.drop(while: { $0 != ":" }).dropFirst(2))
    }
    return s.trimmingCharacters(in: .whitespaces)
}
