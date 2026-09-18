// swift-tools-version: 6.2
import PackageDescription

let package = Package(
    name: "ReturnBotMac",
    platforms: [.macOS(.v26)],
    products: [
        .executable(name: "ReturnBotMac", targets: ["ReturnBotMac"])
    ],
    dependencies: [
        .package(url: "https://github.com/sparkle-project/Sparkle", exact: "2.10.0")
    ],
    targets: [
        .executableTarget(
            name: "ReturnBotMac",
            dependencies: [.product(name: "Sparkle", package: "Sparkle")],
            path: "Sources/ReturnBotMac",
            linkerSettings: [.unsafeFlags(["-Xlinker", "-rpath", "-Xlinker", "@executable_path/../Frameworks"])]
        ),
        .testTarget(name: "ReturnBotMacTests", dependencies: ["ReturnBotMac"])
    ]
)
