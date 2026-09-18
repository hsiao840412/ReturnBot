import XCTest
@testable import ReturnBotMac

final class UpdateStateTests: XCTestCase {
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
