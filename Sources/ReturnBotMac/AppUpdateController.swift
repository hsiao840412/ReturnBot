import AppKit
import Combine
import Sparkle

struct UpdateWorkState {
    var busy: Bool
    var hasUnsavedRecall: Bool

    static func combined(_ states: [UpdateWorkState]) -> UpdateWorkState {
        UpdateWorkState(busy: states.contains { $0.busy }, hasUnsavedRecall: states.contains { $0.hasUnsavedRecall })
    }
}

/// Sparkle owns download verification, installation, rollback and relaunch.
/// The app only supplies user controls and protects active work at termination.
@MainActor
final class AppUpdateController: NSObject, ObservableObject, NSApplicationDelegate, SPUUpdaterDelegate {
    @Published private(set) var canCheckForUpdates = false
    @Published var automaticallyChecks = true {
        didSet { controller?.updater.automaticallyChecksForUpdates = automaticallyChecks }
    }
    private var controller: SPUStandardUpdaterController?
    private var observation: NSKeyValueObservation?
    private var workspaces: [UUID: () -> UpdateWorkState] = [:]

    var workState: UpdateWorkState { .combined(workspaces.values.map { $0() }) }

    func register(_ id: UUID, state: @escaping () -> UpdateWorkState) { workspaces[id] = state }
    func unregister(_ id: UUID) { workspaces.removeValue(forKey: id) }

    func applicationDidFinishLaunching(_ notification: Notification) {
        // Unbundled development executables have no signed feed configuration.
        guard Bundle.main.object(forInfoDictionaryKey: "SUPublicEDKey") != nil else { return }
        let controller = SPUStandardUpdaterController(startingUpdater: false, updaterDelegate: self, userDriverDelegate: nil)
        self.controller = controller
        observation = controller.updater.observe(\.canCheckForUpdates, options: [.initial, .new]) { [weak self] updater, _ in
            Task { @MainActor in self?.canCheckForUpdates = updater.canCheckForUpdates }
        }
        controller.startUpdater()
        automaticallyChecks = controller.updater.automaticallyChecksForUpdates
        if automaticallyChecks { controller.updater.checkForUpdatesInBackground() }
    }

    func checkForUpdates() { controller?.checkForUpdates(nil) }

    func updater(_ updater: SPUUpdater, mayPerform updateCheck: SPUUpdateCheck) throws {
        if workState.busy {
            throw NSError(domain: "ReturnBot.Update", code: 1,
                          userInfo: [NSLocalizedDescriptionKey: "正在辨識或製作文件，完成後再檢查更新。"])
        }
    }

    func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply {
        let state = workState
        if state.busy {
            let alert = NSAlert()
            alert.messageText = "文件還在處理中"
            alert.informativeText = "請等辨識或文件製作完成後，再退出或安裝更新。"
            alert.addButton(withTitle: "繼續處理")
            alert.runModal()
            return .terminateCancel
        }
        if state.hasUnsavedRecall {
            let alert = NSAlert()
            alert.messageText = "寄銷召回還有未匯出的資料"
            alert.informativeText = "退出或安裝更新會關閉目前案件。請先匯出，或確認放棄本次資料。"
            alert.addButton(withTitle: "返回案件")
            alert.addButton(withTitle: "放棄資料並繼續")
            return alert.runModal() == .alertSecondButtonReturn ? .terminateNow : .terminateCancel
        }
        return .terminateNow
    }
}
