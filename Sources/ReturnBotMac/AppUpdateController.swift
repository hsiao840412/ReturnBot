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
    @Published private(set) var automaticallyChecks = true
    @Published private(set) var lastUpdateError: String?
    private let preferences: UserDefaults
    private var controller: SPUStandardUpdaterController?
    private var observations: [NSKeyValueObservation] = []

    override convenience init() { self.init(preferences: .standard) }

    init(preferences: UserDefaults) {
        self.preferences = preferences
        super.init()
        automaticallyChecks = preferences.object(forKey: "SUEnableAutomaticChecks") as? Bool ?? true
        lastUpdateError = preferences.string(forKey: "ReturnBotLastUpdateError")
    }

    func setAutomaticallyChecks(_ enabled: Bool) {
        automaticallyChecks = enabled
        preferences.set(enabled, forKey: "SUEnableAutomaticChecks")
        controller?.updater.automaticallyChecksForUpdates = enabled
    }

    private var workspaces: [UUID: () -> UpdateWorkState] = [:]

    var workState: UpdateWorkState { .combined(workspaces.values.map { $0() }) }

    func register(_ id: UUID, state: @escaping () -> UpdateWorkState) { workspaces[id] = state }
    func unregister(_ id: UUID) { workspaces.removeValue(forKey: id) }

    func applicationDidFinishLaunching(_ notification: Notification) {
        // Unbundled development executables have no signed feed configuration.
        guard Bundle.main.object(forInfoDictionaryKey: "SUPublicEDKey") != nil else { return }
        let controller = SPUStandardUpdaterController(startingUpdater: false, updaterDelegate: self, userDriverDelegate: nil)
        self.controller = controller
        observations = [
            controller.updater.observe(\.canCheckForUpdates, options: [.initial, .new]) { [weak self] updater, _ in
                Task { @MainActor in self?.canCheckForUpdates = updater.canCheckForUpdates }
            },
            controller.updater.observe(\.automaticallyChecksForUpdates, options: [.initial, .new]) { [weak self] updater, _ in
                Task { @MainActor in self?.automaticallyChecks = updater.automaticallyChecksForUpdates }
            }
        ]
        controller.startUpdater()
        automaticallyChecks = controller.updater.automaticallyChecksForUpdates
        // Retry on every launch when enabled, including after a failed download.
        // A user's explicit opt-out always takes precedence over retrying.
        if automaticallyChecks { controller.updater.checkForUpdatesInBackground() }
    }

    func checkForUpdates() {
        guard let updater = controller?.updater, updater.canCheckForUpdates else { return }
        updater.checkForUpdates()
    }

    func recordUpdateResult(_ error: Error?) {
        let nsError = error as NSError?
        // No update is a successful check; cancellation is not a failure to retry.
        if let nsError, nsError.domain == SUSparkleErrorDomain,
           [SUError.noUpdateError.rawValue, SUError.installationCanceledError.rawValue].contains(OSStatus(nsError.code)) {
            lastUpdateError = nil
        } else {
            lastUpdateError = error?.localizedDescription
        }
        preferences.set(lastUpdateError, forKey: "ReturnBotLastUpdateError")
    }

    func updater(_ updater: SPUUpdater, didFinishUpdateCycleFor updateCheck: SPUUpdateCheck, error: Error?) {
        recordUpdateResult(error)
        canCheckForUpdates = updater.canCheckForUpdates
    }

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
