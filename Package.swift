// swift-tools-version:5.9
import PackageDescription

let package = Package(
    name: "OutlookAX",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "OutlookAX", targets: ["OutlookAX"]),
    ],
    targets: [
        .target(
            name: "OutlookAX",
            path: "Sources/OutlookAX"
        ),
        .testTarget(
            name: "OutlookAXTests",
            dependencies: ["OutlookAX"],
            path: "Tests/OutlookAXTests"
        ),
    ]
)
