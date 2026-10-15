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
        XCTAssertNotNil(runner.generatedAt)
        XCTAssertEqual(runner.generatedFor?.returnType, .kbb)
        XCTAssertEqual(runner.generatedFor?.csvURL, directory.appendingPathComponent("input.csv"))
        XCTAssertTrue(runner.isCurrentResult(returnType: .kbb, csvURL: directory.appendingPathComponent("input.csv")))
        XCTAssertFalse(runner.isCurrentResult(returnType: .kbbBattery, csvURL: directory.appendingPathComponent("input.csv")))
        XCTAssertFalse(runner.isCurrentResult(returnType: .kbb, csvURL: directory.appendingPathComponent("another.csv")))
        XCTAssertFalse(runner.isCurrentResult(returnType: .kbb, csvURL: nil))
        // Changing the current selection doesn't discard the previous document.
        XCTAssertEqual(runner.outputPath, "/tmp/測試.xlsx")
    }

    @MainActor
    func testFailedRegenerationCannotLookLikeAnEarlierSuccess() async throws {
        let directory = FileManager.default.temporaryDirectory.appendingPathComponent(UUID().uuidString)
        try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true)
        defer { try? FileManager.default.removeItem(at: directory) }
        let script = directory.appendingPathComponent("helper.sh")
        try """
        if [ "$2" = "kbb-battery" ]; then
            echo '{"type":"result","success":false,"message":"測試失敗"}'
            exit 1
        fi
        echo '{"type":"result","success":true,"outputPath":"/tmp/測試.xlsx"}'
        """.write(to: script, atomically: true, encoding: .utf8)
        let runner = ReturnBotRunner(runtimeOverride: (URL(fileURLWithPath: "/bin/sh"), [script.path], directory))
        let input = directory.appendingPathComponent("input.csv")
        runner.run(returnType: .kbb, csvURL: input)
        var deadline = Date().addingTimeInterval(5)
        while runner.isRunning && Date() < deadline { try await Task.sleep(for: .milliseconds(10)) }
        XCTAssertTrue(runner.isCurrentResult(returnType: .kbb, csvURL: input))

        runner.run(returnType: .kbbBattery, csvURL: input)
        XCTAssertNil(runner.generatedFor)
        XCTAssertNil(runner.generatedAt)
        XCTAssertNil(runner.outputPath)
        deadline = Date().addingTimeInterval(5)
        while runner.isRunning && Date() < deadline { try await Task.sleep(for: .milliseconds(10)) }
        XCTAssertFalse(runner.isRunning)
        XCTAssertNotNil(runner.errorMessage)
        XCTAssertFalse(runner.isCurrentResult(returnType: .kbb, csvURL: input))
        XCTAssertNil(runner.generatedFor)
        XCTAssertNil(runner.outputPath)
    }
}
