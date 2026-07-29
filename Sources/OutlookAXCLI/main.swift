// Sources/OutlookAXCLI/main.swift
// Executable entry point for the outlook-ax CLI.
// This is the only file in the product that calls exit().

import Foundation
import OutlookAXCLIKit

let rawArgs = Array(CommandLine.arguments.dropFirst())
let parser = CLIParser(args: rawArgs)
let parseResult = parser.parse()

let jsonMode = rawArgs.contains("--json")
let renderer = OutputRenderer(jsonMode: jsonMode)

switch parseResult {
case .help:
    let runner = CLIRunner(renderer: renderer)
    exit(runner.run(.help))

case .usageError(let msg):
    fputs("error: \(msg)\n", stderr)
    exit(64)

case .command(let command, let json, _):
    let r = OutputRenderer(jsonMode: json)
    let runner = CLIRunner(renderer: r)
    let code = runner.run(command)
    exit(code)
}
