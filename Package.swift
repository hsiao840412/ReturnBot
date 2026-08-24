// swift-tools-version: 6.2
import PackageDescription

let package = Package(
    name: "ReturnBotMac",
    platforms: [.macOS(.v26)],
    products: [
        .executable(name: "ReturnBotMac", targets: ["ReturnBotMac"])
    ],
    targets: [
        .executableTarget(
            name: "ReturnBotMac",
            path: "Sources/ReturnBotMac"
        )
    ]
)
