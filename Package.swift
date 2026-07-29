// swift-tools-version:5.9
import PackageDescription

let package = Package(
    name: "OutlookAX",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "OutlookAX", targets: ["OutlookAX"]),
        .executable(name: "outlook-ax", targets: ["OutlookAXCLI"]),
    ],
    targets: [
        // Core library: AX, models, localization, parsers, services
        .target(
            name: "OutlookAX",
            path: "Sources/OutlookAX"
        ),
        // CLI runtime: command parsing, dispatch, JSON envelopes, output
        .target(
            name: "OutlookAXCLIKit",
            dependencies: ["OutlookAX"],
            path: "Sources/OutlookAXCLIKit"
        ),
        // Executable entry point
        .executableTarget(
            name: "OutlookAXCLI",
            dependencies: ["OutlookAXCLIKit"],
            path: "Sources/OutlookAXCLI"
        ),
        // Library unit tests
        .testTarget(
            name: "OutlookAXTests",
            dependencies: ["OutlookAX"],
            path: "Tests/OutlookAXTests"
        ),
        // CLI unit tests (parser, JSON envelope)
        .testTarget(
            name: "OutlookAXCLITests",
            dependencies: ["OutlookAXCLIKit"],
            path: "Tests/OutlookAXCLITests"
        ),
        // CLI subprocess tests: build and exec the real binary, verify exit codes + JSON
        .testTarget(
            name: "OutlookAXCLISubprocessTests",
            path: "Tests/OutlookAXCLISubprocessTests"
        ),
    ]
)

