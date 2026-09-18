import SwiftUI

@main
struct ReturnBotMacApp: App {
    @NSApplicationDelegateAdaptor(AppUpdateController.self) private var updates
    var body: some Scene {
        WindowGroup("退料機器人") {
            ReturnBotHome()
                .environmentObject(updates)
                .frame(minWidth: 820, minHeight: 740)
        }
        .windowResizability(.contentSize)
        .defaultSize(width: 960, height: 850)
        .commands {
            CommandGroup(after: .appInfo) {
                Button("檢查更新…") { updates.checkForUpdates() }
                    .disabled(!updates.canCheckForUpdates)
                Toggle("自動檢查更新", isOn: Binding(get: { updates.automaticallyChecks }, set: { updates.automaticallyChecks = $0 }))
            }
        }
    }
}

struct ReturnBotHome: View {
    @EnvironmentObject private var updates: AppUpdateController
    @State private var workspaceID = UUID()
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
        .onAppear {
            updates.register(workspaceID) { [weak runner = runner, weak recallModel = recallModel] in
                UpdateWorkState(
                    busy: (runner?.isRunning ?? false) || (recallModel?.busy ?? false),
                    hasUnsavedRecall: recallModel?.hasUnexportedWork ?? false
                )
            }
        }
        .onDisappear { updates.unregister(workspaceID) }
    }
}
