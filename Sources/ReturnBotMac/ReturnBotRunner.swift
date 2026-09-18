import AppKit
import Foundation

@MainActor
final class ReturnBotRunner: ObservableObject {
    @Published var isRunning = false
    @Published var status = "就緒"
    @Published var outputPath: String?
    @Published var errorMessage: String?
    @Published var warnings: [String] = []

    private var process: Process?
    private var finalResult: [String: Any]?
    private var bufferedOutput = ""
    private var bufferedError = ""
    private var securityScopedURL: URL?

    private let runtimeOverride: (executable: URL, argumentsPrefix: [String], workingDirectory: URL)?

    init(runtimeOverride: (executable: URL, argumentsPrefix: [String], workingDirectory: URL)? = nil) {
        self.runtimeOverride = runtimeOverride
    }

    private var runtime: (executable: URL, argumentsPrefix: [String], workingDirectory: URL)? {
        if let runtimeOverride { return runtimeOverride }
        if let resources = Bundle.main.resourceURL {
            let bundledHelper = resources.appendingPathComponent("ReturnBotHelper/ReturnBotHelper")
            if FileManager.default.isExecutableFile(atPath: bundledHelper.path) {
                return (bundledHelper, [], resources)
            }
        }

        let root = URL(fileURLWithPath: FileManager.default.currentDirectoryPath)
        let helper = root.appendingPathComponent("returnbot_cli.py")
        let virtualPython = root.appendingPathComponent(".venv/bin/python")
        guard FileManager.default.fileExists(atPath: helper.path) else { return nil }
        let python = FileManager.default.fileExists(atPath: virtualPython.path)
            ? virtualPython
            : URL(fileURLWithPath: "/usr/bin/python3")
        return (python, [helper.path], root)
    }

    func run(returnType: ReturnType, csvURL: URL) {
        guard !isRunning else { return }

        isRunning = true
        status = "正在準備文件..."
        outputPath = nil
        errorMessage = nil
        warnings = []
        bufferedOutput = ""
        bufferedError = ""
        finalResult = nil

        if csvURL.startAccessingSecurityScopedResource() {
            securityScopedURL = csvURL
        }

        guard let runtime else {
            finishWithError("找不到 ReturnBot Python helper。")
            return
        }

        let task = Process()
        task.currentDirectoryURL = runtime.workingDirectory
        task.executableURL = runtime.executable
        task.arguments = runtime.argumentsPrefix + ["--type", returnType.cliValue, "--csv", csvURL.path]

        let output = Pipe()
        task.standardOutput = output
        task.standardError = output
        do {
            try task.run()
            process = task
            Task.detached { [weak self] in
                var pending = Data()
                while true {
                    let data = output.fileHandleForReading.availableData
                    if data.isEmpty { break }
                    pending.append(data)
                    while let newline = pending.firstIndex(of: 10) {
                        await self?.consume(String(decoding: pending.prefix(through: newline), as: UTF8.self))
                        pending.removeSubrange(...newline)
                    }
                }
                if !pending.isEmpty {
                    await self?.consume(String(decoding: pending, as: UTF8.self) + "\n")
                }
                task.waitUntilExit()
                await self?.processEnded(task.terminationStatus)
            }
        } catch {
            finishWithError("無法啟動文件製作：\(error.localizedDescription)")
        }
    }

    private func processEnded(_ exitCode: Int32) {
        guard isRunning else { return }
        if let payload = finalResult {
            warnings = payload["warnings"] as? [String] ?? []
            if payload["success"] as? Bool == true && exitCode == 0 {
                outputPath = payload["outputPath"] as? String
                errorMessage = nil
                status = "生成完成"
                finish()
            } else {
                finishWithError(payload["message"] as? String ?? "文件製作失敗。")
            }
            return
        }
        let diagnostic = bufferedError.trimmingCharacters(in: .whitespacesAndNewlines)
        finishWithError(diagnostic.isEmpty ? "文件製作未回傳結果（\(exitCode)）。" : diagnostic)
    }

    func openOutput() {
        guard let outputPath else { return }
        NSWorkspace.shared.open(URL(fileURLWithPath: outputPath))
    }

    private func consume(_ text: String) {
        bufferedOutput += text
        let lines = bufferedOutput.split(separator: "\n", omittingEmptySubsequences: false)
        bufferedOutput = String(lines.last ?? "")
        for line in lines.dropLast() where !line.isEmpty {
            consumeLine(String(line))
        }
    }

    private func consumeLine(_ line: String) {
        guard
            let data = line.data(using: .utf8),
            let payload = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
            let type = payload["type"] as? String
        else { bufferedError += line + "\n"; return }

        if type == "progress" {
            status = payload["message"] as? String ?? "處理中..."
            return
        }

        guard type == "result" else { return }
        finalResult = payload
    }

    private func finishWithError(_ message: String) {
        errorMessage = message
        status = "生成失敗"
        finish()
    }

    private func finish() {
        isRunning = false
        process = nil
        securityScopedURL?.stopAccessingSecurityScopedResource()
        securityScopedURL = nil
    }
}
