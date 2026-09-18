import SwiftUI

@main
struct ReturnBotMacApp: App {
    var body: some Scene {
        WindowGroup("退料機器人") {
            ReturnBotHome()
                .frame(minWidth: 820, minHeight: 740)
        }
        .windowResizability(.contentSize)
        .defaultSize(width: 960, height: 850)
    }
}

struct ReturnBotHome: View {
    @Environment(\.colorScheme) private var colorScheme
    @State private var recall = false
    @StateObject private var recallModel = RecallModel()
    @StateObject private var runner = ReturnBotRunner()
    @State private var returnType: ReturnType = .mailIn
    @State private var csvURL: URL?

    var body: some View {
        VStack(spacing: 0) {
            Picker("作業", selection: $recall) {
                Text("一般退料").tag(false)
                Text("寄銷召回 Beta").tag(true)
            }
            .pickerStyle(.segmented)
            .padding(.horizontal, 24)
            .padding(.vertical, 12)
            Divider()
            if recall { RecallView(model: recallModel) } else { ContentView(runner: runner, returnType: $returnType, csvURL: $csvURL) }
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity)
        .background(AppSurface.window(colorScheme).ignoresSafeArea())
        .background(TahoeWindowBackground().ignoresSafeArea())
    }
}
