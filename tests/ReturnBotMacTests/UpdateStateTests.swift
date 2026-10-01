import XCTest
import Sparkle
@testable import ReturnBotMac

final class UpdateStateTests: XCTestCase {
    @MainActor
    func testAutomaticPreferenceAndRetrySurviveRestart() {
        let suite = "ReturnBotTests." + UUID().uuidString
        let defaults = UserDefaults(suiteName: suite)!
        defer { defaults.removePersistentDomain(forName: suite) }
        let model = AppUpdateController(preferences: defaults)
        XCTAssertTrue(model.automaticallyChecks)
        model.setAutomaticallyChecks(false)
        model.recordUpdateResult(NSError(domain: NSURLErrorDomain, code: NSURLErrorNetworkConnectionLost))
        let restarted = AppUpdateController(preferences: defaults)
        XCTAssertFalse(restarted.automaticallyChecks)
        XCTAssertNotNil(restarted.lastUpdateError)
        restarted.setAutomaticallyChecks(true)
        restarted.recordUpdateResult(NSError(domain: SUSparkleErrorDomain,
            code: Int(SUError.noUpdateError.rawValue)))
        XCTAssertNil(restarted.lastUpdateError)
        restarted.recordUpdateResult(NSError(domain: SUSparkleErrorDomain,
            code: Int(SUError.installationCanceledError.rawValue)))
        XCTAssertNil(restarted.lastUpdateError)
        restarted.recordUpdateResult(nil)
        let recovered = AppUpdateController(preferences: defaults)
        XCTAssertTrue(recovered.automaticallyChecks)
        XCTAssertNil(recovered.lastUpdateError)
    }

    @MainActor
    func testAllWindowsProtectWorkAndUnregister() {
        let controller = AppUpdateController()
        let first = UUID(), second = UUID()
        controller.register(first) { .init(busy: true, hasUnsavedRecall: false) }
        controller.register(second) { .init(busy: false, hasUnsavedRecall: true) }
        XCTAssertTrue(controller.workState.busy)
        XCTAssertTrue(controller.workState.hasUnsavedRecall)
        controller.unregister(first)
        XCTAssertFalse(controller.workState.busy)
        XCTAssertTrue(controller.workState.hasUnsavedRecall)
        controller.unregister(second)
        XCTAssertFalse(controller.workState.hasUnsavedRecall)
    }

    @MainActor
    func testRecallExportThenEditRequiresProtection() {
        let model = RecallModel()
        XCTAssertFalse(model.hasUnexportedWork)
        model.input = "661-12345 1"
        XCTAssertTrue(model.hasUnexportedWork)
        model.input = ""
        model.caseNumber = "TEST"
        XCTAssertTrue(model.hasUnexportedWork)
        model.outputPath = "/tmp/test.xlsx"
        XCTAssertFalse(model.hasUnexportedWork)
        model.invalidate()
        XCTAssertTrue(model.hasUnexportedWork)
    }
}
