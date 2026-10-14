import SwiftUI

@main
struct ReturnBotMacApp: App {
    @NSApplicationDelegateAdaptor(AppUpdateController.self) private var updates
    var body: some Scene {
        WindowGroup("退料機器人") {
            ReturnBotHome()
                .environmentObject(updates)
                .frame(minWidth: 1100, minHeight: 700)
        }
        .windowResizability(.contentSize)
        .defaultSize(width: 1280, height: 800)
        .commands {
            UpdateCommands(updates: updates)
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
        HStack(spacing: 0) {
            sidebar
            Divider()
            VStack(spacing: 0) {
                WorkflowHeader(
                    title: recall ? "寄銷召回" : "一般退料",
                    subtitle: recall ? "整理零件、核對價格，完成分單與匯出" : "選擇退料類型與 ePacking List，製作 Excel 文件"
                )
                Divider()
                if recall {
                    RecallView(model: recallModel)
                } else {
                    ContentView(runner: runner, returnType: $returnType, csvURL: $csvURL)
                }
                Divider()
                statusBar
            }
            .frame(maxWidth: .infinity, maxHeight: .infinity)
        }
        .frame(maxWidth: .infinity, maxHeight: .infinity)
        .background(WorkflowBackground())
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

    private var sidebar: some View {
        VStack(alignment: .leading, spacing: 24) {
            HStack(spacing: 10) {
                Image(systemName: "arrow.uturn.backward")
                    .font(.system(size: 18, weight: .semibold)).foregroundStyle(.white)
                    .frame(width: 38, height: 38)
                    .glassEffect(.regular.tint(.blue), in: .circle)
                VStack(alignment: .leading, spacing: 3) {
                    Text("ReturnBot").font(.system(size: 20, weight: .bold, design: .rounded))
                    Text("自動生成Excel文件").font(.caption).foregroundStyle(.secondary)
                }
            }
            VStack(alignment: .leading, spacing: 8) {
                Text("作業").font(.caption.weight(.semibold)).foregroundStyle(.secondary)
                    .padding(.horizontal, 12).padding(.bottom, 4)
                navigationButton("一般退料", symbol: "arrow.uturn.backward.circle", selected: !recall) { recall = false }
                navigationButton("寄銷召回", symbol: "shippingbox", selected: recall) { recall = true }
            }
            Spacer()
            VStack(alignment: .leading, spacing: 12) {
                Label(runner.isRunning || recallModel.busy ? "作業處理中" : "",
                      systemImage: runner.isRunning || recallModel.busy ? "hourglass" : "desktopcomputer")
                    .font(.caption).foregroundStyle(.secondary)
                Divider()
                Button(updates.lastUpdateError == nil ? "檢查更新…" : "重試檢查更新…", systemImage: "arrow.triangle.2.circlepath") {
                    updates.checkForUpdates()
                }
                .buttonStyle(.glass).disabled(!updates.canCheckForUpdates)
                Toggle("自動檢查更新", isOn: Binding(
                    get: { updates.automaticallyChecks },
                    set: { updates.setAutomaticallyChecks($0) }
                ))
                .font(.caption).toggleStyle(.checkbox)
                Text("v\(Bundle.main.object(forInfoDictionaryKey: "CFBundleShortVersionString") as? String ?? "3.2.6")")
                    .font(.caption.monospacedDigit()).foregroundStyle(.tertiary)
            }
        }
        .padding(.horizontal, 16).padding(.vertical, 24)
        .frame(width: 228)
        .frame(maxHeight: .infinity)
        .background(AppSurface.panel(colorScheme).opacity(0.65))
    }

    private func navigationButton(_ title: String, symbol: String, selected: Bool, action: @escaping () -> Void) -> some View {
        Button {
            withAnimation(.snappy(duration: 0.22)) { action() }
        } label: {
            HStack(spacing: 10) {
                Image(systemName: symbol).frame(width: 22)
                Text(title).fontWeight(selected ? .semibold : .regular)
                Spacer()
            }
            .padding(.horizontal, 12).frame(height: 44)
            .foregroundStyle(selected ? Color.white : Color.primary)
            .background(selected ? Color.blue : Color.clear, in: .rect(cornerRadius: 12))
            .contentShape(.rect)
        }
        .buttonStyle(.plain).help(title)
        .accessibilityAddTraits(selected ? .isSelected : [])
    }

    @ViewBuilder private var statusBar: some View {
        if recall {
            WorkflowStatusBar(
                text: recallModel.error ?? (recallModel.message.isEmpty ? "就緒，請匯入價格表並加入零件" : recallModel.message),
                symbol: recallModel.error != nil ? "exclamationmark.triangle.fill" : recallModel.outputPath != nil ? "checkmark.circle.fill" : "info.circle",
                color: recallModel.error != nil ? .red : recallModel.outputPath != nil ? .green : .secondary,
                busy: recallModel.busy,
                detail: "\(recallModel.rows.count) 筆明細"
            )
        } else {
            WorkflowStatusBar(
                text: runner.errorMessage ?? runner.status,
                symbol: runner.errorMessage != nil ? "exclamationmark.triangle.fill" : runner.outputPath != nil ? "checkmark.circle.fill" : "info.circle",
                color: runner.errorMessage != nil ? .red : runner.outputPath != nil ? .green : .secondary,
                busy: runner.isRunning,
                detail: "輸出至下載項目"
            )
        }
    }
}

/// Commands must observe the model themselves; delegate-adaptor access alone
/// does not subscribe the menu to its published settings/session changes.
struct UpdateCommands: Commands {
    @ObservedObject var updates: AppUpdateController

    var body: some Commands {
        CommandGroup(after: .appInfo) {
            Button(updates.lastUpdateError == nil ? "檢查更新…" : "重試檢查更新…") {
                updates.checkForUpdates()
            }
            .disabled(!updates.canCheckForUpdates)
            Toggle("自動檢查更新", isOn: Binding(
                get: { updates.automaticallyChecks },
                set: { updates.setAutomaticallyChecks($0) }
            ))
        }
    }
}
