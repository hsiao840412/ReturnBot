import SwiftUI

@main
struct ReturnBotMacApp: App {
    var body: some Scene {
        WindowGroup {
            ContentView()
                .frame(minWidth: 700, minHeight: 640)
        }
        .windowResizability(.contentSize)
        .defaultSize(width: 760, height: 700)
    }
}
