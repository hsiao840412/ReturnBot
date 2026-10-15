import SwiftUI
import UniformTypeIdentifiers

enum ReturnType: String, CaseIterable, Identifiable {
    case mailIn = "Mail-in KBB"
    case mailInBattery = "Mail-in 電池膨脹"
    case kbb = "一般 KBB"
    case kbbBattery = "單獨鋰電池 KBB"

    var id: Self { self }

    var cliValue: String {
        switch self {
        case .mailIn: "mail-in"
        case .mailInBattery: "mail-in-battery"
        case .kbb: "kbb"
        case .kbbBattery: "kbb-battery"
        }
    }

    var symbol: String {
        switch self {
        case .mailIn, .mailInBattery: "shippingbox"
        case .kbb, .kbbBattery: "arrow.uturn.backward.circle"
        }
    }
}

struct ContentView: View {
    @ObservedObject var runner: ReturnBotRunner
    @Binding var returnType: ReturnType
    @Binding var csvURL: URL?
    @State private var showingImporter = false
    @State private var showingInstructions = false
    @State private var fileError: String?
    @State private var isDropTargeted = false

    private var currentResult: Bool { runner.isCurrentResult(returnType: returnType, csvURL: csvURL) }
    private var previousResult: Bool { runner.outputPath != nil && !currentResult }

    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 18) {
                nextStep
                HStack(alignment: .top, spacing: 18) {
                    typePicker.frame(width: 300)
                    VStack(spacing: 18) {
                        filePicker
                        actionArea
                    }.frame(maxWidth: .infinity)
                }
                DisclosureGroup("第一次使用？查看操作說明", isExpanded: $showingInstructions) {
                    VStack(alignment: .leading, spacing: 10) {
                        Text("1. 依本次作業選擇退料方式；不確定分類時，請先向資深夥伴確認。")
                        Text("2. 加入從 ePacking List 下載的 CSV 清單，保留原始欄位與內容。")
                        Text("3. 按「製作 Excel 文件」，完成後開啟文件，核對料號、數量與退料資料。")
                        Text("文件存放在 Mac 的「\(AppFileLocations.outputFolderLabel)」。製作完成表示檔案已建立，內容仍需人工核對。")
                            .foregroundStyle(.secondary)
                    }
                    .font(.callout).padding(.top, 10).textSelection(.enabled)
                }
                .font(.callout).padding(16)
            }
            .padding(.horizontal, 30).padding(.top, 24).padding(.bottom, 30)
        }
        .groupBoxStyle(WorkflowCardStyle())
        .fileImporter(isPresented: $showingImporter,
                      allowedContentTypes: [.commaSeparatedText, .plainText], allowsMultipleSelection: false) { result in
            switch result {
            case .success(let urls): if let url = urls.first { csvURL = url }
            case .failure(let error):
                if (error as NSError).code != NSUserCancelledError { fileError = "無法開啟清單，請重新選擇檔案。\n\(error.localizedDescription)" }
            }
        }
        .alert("未能製作文件", isPresented: errorBinding) {
            Button("返回檢查") { runner.errorMessage = nil }
        } message: { Text(runner.errorMessage ?? "") }
        .alert("無法加入清單", isPresented: Binding(get: { fileError != nil }, set: { if !$0 { fileError = nil } })) {
            Button("好") { fileError = nil }
        } message: { Text(fileError ?? "") }
    }

    @ViewBuilder private var nextStep: some View {
        if runner.isRunning {
            WorkflowNotice(title: "正在製作文件，請稍候", detail: runner.status, symbol: "hourglass")
        } else if currentResult {
            WorkflowNotice(title: "下一步：開啟 Excel，人工核對內容",
                           detail: "文件已儲存。請核對料號、數量與退料資料，再依原作業流程使用。",
                           symbol: "doc.text.magnifyingglass", color: .orange)
        } else if previousResult {
            WorkflowNotice(title: "退料方式或清單已變更，請重新製作文件",
                           detail: "下方保留上一份文件供查閱；它不是目前選擇的製作結果。",
                           symbol: "arrow.triangle.2.circlepath", color: .orange)
        } else {
            WorkflowNotice(title: csvURL == nil ? "下一步：加入本次退料清單" : "下一步：確認選擇，再製作 Excel",
                           detail: csvURL == nil ? "先確認左側退料方式，再選擇或拖入 ePacking List 的 CSV 檔。" : "請確認退料方式和檔名，文件將儲存到本機「\(AppFileLocations.outputFolderLabel)」。")
        }
    }

    private var typePicker: some View {
        GroupBox("1  選擇退料方式") {
            VStack(alignment: .leading, spacing: 10) {
                Text("請依本次退料作業選擇表單。")
                    .font(.callout).foregroundStyle(.secondary)
                ForEach(ReturnType.allCases) { type in
                    Button {
                        withAnimation(.snappy(duration: 0.22)) { returnType = type }
                    } label: {
                        HStack(spacing: 10) {
                            Image(systemName: type.symbol).frame(width: 22)
                            VStack(alignment: .leading, spacing: 4) {
                                Text(type.rawValue).fontWeight(returnType == type ? .semibold : .regular)
                                Text(type.formDescription).font(.caption)
                                    .foregroundStyle(returnType == type ? Color.white.opacity(0.85) : Color.secondary)
                            }
                            Spacer(minLength: 0)
                            if returnType == type { Image(systemName: "checkmark.circle.fill") }
                        }
                        .padding(.horizontal, 12).padding(.vertical, 12)
                        .frame(maxWidth: .infinity, alignment: .leading)
                        .foregroundStyle(returnType == type ? Color.white : Color.primary)
                        .background(returnType == type ? Color.blue : Color.clear, in: .rect(cornerRadius: 12))
                        .contentShape(.rect)
                    }
                    .buttonStyle(.plain).help(type.rawValue)
                    .accessibilityAddTraits(returnType == type ? .isSelected : [])
                }
            }
        }
        .disabled(runner.isRunning)
    }

    private var filePicker: some View {
        GroupBox("2  加入退料清單（CSV）") {
            VStack(alignment: .leading, spacing: 14) {
                HStack(alignment: .top, spacing: 12) {
                    Image(systemName: csvURL == nil ? "doc.badge.plus" : "doc.text.fill")
                        .font(.system(size: 26, weight: .medium)).foregroundStyle(.blue)
                    VStack(alignment: .leading, spacing: 6) {
                        Text(csvURL?.lastPathComponent ?? "將清單拖到這裡，或按下方按鈕選擇")
                            .font(.body.weight(.medium)).lineLimit(2).textSelection(.enabled)
                            .help(csvURL?.path ?? "ePacking List 的 CSV 清單")
                        Text(csvURL == nil ? "使用從 ePacking List 下載的 CSV 檔。" : "已選擇檔案；製作前請確認檔名是否正確。")
                            .font(.callout).foregroundStyle(.secondary)
                    }
                    Spacer(minLength: 0)
                }
                Button(csvURL == nil ? "選擇退料清單…" : "更換清單…", systemImage: "folder") { showingImporter = true }
                    .buttonStyle(.glass)
            }
            .padding(4).frame(maxWidth: .infinity, minHeight: 96, alignment: .leading)
        }
        .overlay(RoundedRectangle(cornerRadius: 19).stroke(isDropTargeted ? Color.blue : Color.clear, style: StrokeStyle(lineWidth: 2, dash: [6])))
        .dropDestination(for: URL.self) { urls, _ in
            guard !runner.isRunning else { return false }
            guard urls.count == 1, let url = urls.first, url.isFileURL,
                  ["csv", "txt"].contains(url.pathExtension.lowercased()) else {
                fileError = "請一次加入一份 CSV 退料清單。Excel 活頁簿（.xlsx）請先匯出為 CSV。"
                return false
            }
            csvURL = url
            return true
        } isTargeted: { isDropTargeted = $0 }
        .disabled(runner.isRunning)
    }

    private var actionArea: some View {
        GroupBox("3  製作文件，再人工核對") {
            VStack(alignment: .leading, spacing: 14) {
                if let path = runner.outputPath, let input = runner.generatedFor {
                    Label(currentResult ? "文件已製作，內容待人工核對" : "上一份文件（與目前選擇不同）",
                          systemImage: currentResult ? "doc.text.magnifyingglass" : "clock.arrow.circlepath")
                        .font(.headline).foregroundStyle(currentResult ? Color.orange : Color.secondary)
                    VStack(alignment: .leading, spacing: 7) {
                        LabeledContent("退料方式", value: input.returnType.rawValue)
                        LabeledContent("使用的清單", value: input.csvURL.lastPathComponent)
                        if let date = runner.generatedAt {
                            LabeledContent("製作時間") { Text(date, format: .dateTime.month().day().hour().minute()).monospacedDigit() }
                        }
                        LabeledContent("Excel 檔案", value: URL(fileURLWithPath: path).lastPathComponent)
                    }
                    .font(.callout).textSelection(.enabled)
                    Divider()
                } else {
                    LabeledContent("本次退料方式", value: returnType.rawValue).font(.callout)
                    Text(csvURL == nil ? "請先加入清單，才能製作文件。" : "請確認上方退料方式與清單，再開始製作。")
                        .font(.callout).foregroundStyle(.secondary)
                }
                if currentResult {
                    Button { runner.openOutput() } label: {
                        primaryLabel("開啟 Excel，核對內容", symbol: "arrow.up.forward.app")
                    }.buttonStyle(.glassProminent).controlSize(.large)
                    HStack(spacing: 12) {
                        Button("製作另一份…") { showingImporter = true }
                        Button("重新製作本份文件") { generate() }
                    }.buttonStyle(.glass)
                } else {
                    Button { generate() } label: {
                        primaryLabel(runner.isRunning ? "正在製作，請稍候…" : "製作 Excel 文件", symbol: runner.isRunning ? "hourglass" : "arrow.right")
                    }
                    .buttonStyle(.glassProminent).controlSize(.large)
                    .disabled(csvURL == nil || runner.isRunning)
                    if previousResult {
                        Button("開啟上一份文件", systemImage: "clock.arrow.circlepath") { runner.openOutput() }.buttonStyle(.glass)
                    }
                }
                if runner.outputPath != nil {
                    ForEach(runner.warnings, id: \.self) { warning in
                        Label(warning, systemImage: "exclamationmark.triangle.fill")
                            .font(.callout).foregroundStyle(.orange).textSelection(.enabled)
                    }
                }
                Text("儲存位置：Mac 的「\(AppFileLocations.outputFolderLabel)」。")
                    .font(.caption).foregroundStyle(.secondary)
            }
        }
    }

    private func primaryLabel(_ title: String, symbol: String) -> some View {
        HStack {
            if runner.isRunning { ProgressView().controlSize(.small) }
            Text(title)
            Spacer()
            Image(systemName: symbol)
        }.frame(maxWidth: .infinity).padding(.vertical, 5)
    }

    private func generate() {
        guard let csvURL else { return }
        runner.run(returnType: returnType, csvURL: csvURL)
    }

    private var errorBinding: Binding<Bool> {
        Binding(get: { runner.errorMessage != nil }, set: { if !$0 { runner.errorMessage = nil } })
    }
}

private extension ReturnType {
    var formDescription: String {
        switch self {
        case .mailIn: "Mail-in KBB 退料表單"
        case .mailInBattery: "電池膨脹專用表單"
        case .kbb: "一般 KBB 退料表單"
        case .kbbBattery: "單獨鋰電池專用表單"
        }
    }
}
