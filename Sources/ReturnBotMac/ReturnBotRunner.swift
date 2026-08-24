import AppKit
import Foundation

@MainActor
final class ReturnBotRunner: ObservableObject {
    @Published var isRunning = false
    @Published var isPreparing = false
    @Published var status = "就緒"
    @Published var outputPath: String?
    @Published var errorMessage: String?
    @Published var warnings: [String] = []

    private var process: Process?
    private var bufferedOutput = ""
    private var bufferedError = ""
    private var preflightOutput = ""
    private var preflightError = ""
    private var securityScopedURL: URL?

    private var runtime: (executable: URL, argumentsPrefix: [String], workingDirectory: URL)? {
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

    func prepareExcelAccess() {
        guard !isPreparing, !isRunning else { return }

        guard let runtime else {
            finishWithError("找不到 ReturnBot Python helper。")
            return
        }

        isPreparing = true
        status = "正在準備 Excel..."
        errorMessage = nil
        preflightOutput = ""
        preflightError = ""

        let task = Process()
        task.currentDirectoryURL = runtime.workingDirectory
        task.executableURL = runtime.executable
        task.arguments = runtime.argumentsPrefix + ["--preflight"]

        let stdout = Pipe()
        let stderr = Pipe()
        task.standardOutput = stdout
        task.standardError = stderr
        stdout.fileHandleForReading.readabilityHandler = { [weak self] handle in
            let data = handle.availableData
            guard !data.isEmpty, let text = String(data: data, encoding: .utf8) else { return }
            Task { @MainActor in self?.consumePreflight(text) }
        }
        stderr.fileHandleForReading.readabilityHandler = { [weak self] handle in
            let data = handle.availableData
            guard !data.isEmpty, let text = String(data: data, encoding: .utf8) else { return }
            Task { @MainActor in self?.preflightError += text }
        }
        task.terminationHandler = { [weak self] process in
            Task { @MainActor in
                guard let self, self.isPreparing else { return }
                let diagnostic = self.preflightError.trimmingCharacters(in: .whitespacesAndNewlines)
                self.finishPreflight(
                    success: false,
                    message: diagnostic.isEmpty
                        ? "Excel 權限預檢已結束（\(process.terminationStatus)）。"
                        : diagnostic
                )
            }
        }

        do {
            try task.run()
            process = task
        } catch {
            finishPreflight(success: false, message: "無法啟動 Excel 權限預檢：\(error.localizedDescription)")
        }
    }

    func run(returnType: ReturnType, csvURL: URL) {
        guard !isRunning else { return }

        isRunning = true
        status = "正在啟動 Python..."
        outputPath = nil
        errorMessage = nil
        warnings = []
        bufferedOutput = ""
        bufferedError = ""

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

        let stdout = Pipe()
        let stderr = Pipe()
        task.standardOutput = stdout
        task.standardError = stderr

        stdout.fileHandleForReading.readabilityHandler = { [weak self] handle in
            let data = handle.availableData
            guard !data.isEmpty, let text = String(data: data, encoding: .utf8) else { return }
            Task { @MainActor in self?.consume(text) }
        }

        stderr.fileHandleForReading.readabilityHandler = { [weak self] handle in
            let data = handle.availableData
            guard !data.isEmpty, let text = String(data: data, encoding: .utf8) else { return }
            Task { @MainActor in
                // xlwings may emit harmless diagnostics; only surface stderr if the task fails.
                self?.bufferedError += text
            }
        }

        task.terminationHandler = { [weak self] process in
            Task { @MainActor in
                guard let self else { return }
                if self.isRunning {
                    let diagnostic = self.bufferedError.trimmingCharacters(in: .whitespacesAndNewlines)
                    self.finishWithError(diagnostic.isEmpty
                        ? "Python 程序已結束（\(process.terminationStatus)）。"
                        : diagnostic)
                }
            }
        }

        do {
            try task.run()
            process = task
        } catch {
            finishWithError("無法啟動 Python：\(error.localizedDescription)")
        }
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

    private func consumePreflight(_ text: String) {
        preflightOutput += text
        let lines = preflightOutput.split(separator: "\n", omittingEmptySubsequences: false)
        preflightOutput = String(lines.last ?? "")
        for line in lines.dropLast() where !line.isEmpty {
            guard
                let data = String(line).data(using: .utf8),
                let payload = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
                let type = payload["type"] as? String
            else { continue }

            if type == "progress" {
                status = payload["message"] as? String ?? "正在準備 Excel..."
            } else if type == "result" {
                finishPreflight(
                    success: payload["success"] as? Bool ?? false,
                    message: payload["message"] as? String ?? "Excel 權限預檢失敗。"
                )
            }
        }
    }

    private func finishPreflight(success: Bool, message: String) {
        isPreparing = false
        process = nil
        status = success ? "Excel 已就緒" : "Excel 需要授權"
        if !success { errorMessage = message }
    }

    private func consumeLine(_ line: String) {
        guard
            let data = line.data(using: .utf8),
            let payload = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
            let type = payload["type"] as? String
        else { return }

        if type == "progress" {
            status = payload["message"] as? String ?? "處理中..."
            return
        }

        guard type == "result" else { return }
        let success = payload["success"] as? Bool ?? false
        warnings = payload["warnings"] as? [String] ?? []
        if success {
            errorMessage = nil
            outputPath = payload["outputPath"] as? String
            status = "生成完成"
            finish()
        } else {
            finishWithError(payload["message"] as? String ?? "發生未知錯誤。")
        }
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
