import Foundation

// MARK: - OutputRenderer
//
// Renders CLI results to stdout/stderr without calling exit().
// main.swift reads the returned exit code and calls exit().

public struct OutputRenderer {

    public let jsonMode: Bool
    private let encoder: JSONEncoder

    public init(jsonMode: Bool) {
        self.jsonMode = jsonMode
        let enc = JSONEncoder()
        enc.outputFormatting = [.prettyPrinted, .sortedKeys]
        self.encoder = enc
    }

    // MARK: - Success output

    public func printSuccess<T: Encodable>(_ data: T, command: String) {
        if jsonMode {
            let envelope = JSONSuccess(command: command, data: data)
            printEncoded(envelope)
        }
        // For non-JSON mode, callers format their own human-readable output.
    }

    public func printError(code: String, message: String, command: String, details: [String: String] = [:]) {
        if jsonMode {
            let envelope = JSONFailure(command: command, error: JSONErrorDetails(code: code, message: message, details: details))
            printEncoded(envelope)
        } else {
            fputs("error: \(message)\n", stderr)
        }
    }

    public func printHelp(_ text: String) {
        print(text)
    }

    // MARK: - Private

    private func printEncoded<T: Encodable>(_ value: T) {
        guard let data = try? encoder.encode(value),
              let str = String(data: data, encoding: .utf8) else { return }
        print(str)
    }
}
