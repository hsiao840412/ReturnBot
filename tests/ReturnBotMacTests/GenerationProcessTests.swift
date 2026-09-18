import XCTest
@testable import ReturnBotMac

final class GenerationProcessTests: XCTestCase {
    @MainActor
    func testFinalResultIsDrainedBeforeCompletion() async throws {
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }
        let script = directory.appendingPathComponent("helper.sh")
        try """
        echo '{"type":"progress","message":"正在製作文件"}'
        printf '%s' '{"type":"result","success":true,"outputPath":"/tmp/測試.xlsx"}'
        """.write(to: script, atomically: true, encoding: .utf8)
        let runner = ReturnBotRunner(runtimeOverride: (URL(fileURLWithPath: "/bin/sh"), [script.path], directory))
        runner.run(returnType: .kbb, csvURL: directory.appendingPathComponent("input.csv"))
        let deadline = Date().addingTimeInterval(5)
        while runner.isRunning && Date() < deadline { try await Task.sleep(for: .milliseconds(10)) }
        XCTAssertFalse(runner.isRunning)
        XCTAssertNil(runner.errorMessage)
        XCTAssertEqual(runner.status, "生成完成")
        XCTAssertEqual(runner.outputPath, "/tmp/測試.xlsx")
    }
}
