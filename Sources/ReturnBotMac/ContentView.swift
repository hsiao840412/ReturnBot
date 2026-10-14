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

    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 18) {
                HStack(alignment: .top, spacing: 18) {
                    typePicker.frame(width: 300)
                    VStack(spacing: 18) {
                        filePicker
                        actionArea
                    }
                    .frame(maxWidth: .infinity)
                }
                GroupBox("作業流程") {
                    HStack(alignment: .top, spacing: 20) {
                        step("01", title: "選擇類型", detail: "依退回零件選擇對應作業。")
                        step("02", title: "匯入 CSV", detail: "使用 ePacking List 的 CSV 檔案。")
                        step("03", title: "產出後核對", detail: "開啟 Excel，人工確認內容後再使用。")
                    }
                }
            }
            .padding(.horizontal, 30).padding(.top, 24).padding(.bottom, 30)
        }
        .groupBoxStyle(WorkflowCardStyle())
        .fileImporter(
            isPresented: $showingImporter,
            allowedContentTypes: [.commaSeparatedText, .plainText],
            allowsMultipleSelection: false
        ) { result in
            if case let .success(urls) = result { csvURL = urls.first }
        }
        .alert("生成失敗", isPresented: errorBinding) {
            Button("好") { runner.errorMessage = nil }
        } message: {
            Text(runner.errorMessage ?? "")
        }
    }

    private var typePicker: some View {
        GroupBox("01  退料類型") {
            VStack(spacing: 8) {
                ForEach(ReturnType.allCases) { type in
                    Button {
                        withAnimation(.snappy(duration: 0.22)) { returnType = type }
                    } label: {
                        HStack(spacing: 11) {
                            Image(systemName: type.symbol).frame(width: 22)
                            Text(type.rawValue).fontWeight(returnType == type ? .semibold : .regular)
                            Spacer(minLength: 4)
                            if returnType == type { Image(systemName: "checkmark.circle.fill") }
                        }
                        .padding(.horizontal, 15).frame(height: 43)
                        .foregroundStyle(returnType == type ? Color.white : Color.primary)
                        .background(returnType == type ? Color.blue : Color.clear, in: .rect(cornerRadius: 15))
                        .contentShape(.rect)
                    }
                    .buttonStyle(.plain).help(type.rawValue)
                    .accessibilityAddTraits(returnType == type ? .isSelected : [])
                }
            }
        }
    }

    private var filePicker: some View {
        GroupBox("02  ePacking List") {
            HStack(spacing: 12) {
                Image(systemName: csvURL == nil ? "doc.text" : "doc.text.fill")
                    .font(.system(size: 24, weight: .medium))
                    .frame(width: 36).foregroundStyle(.blue)
                VStack(alignment: .leading, spacing: 6) {
                    Text(csvURL?.lastPathComponent ?? "尚未選擇 CSV")
                        .font(.body.weight(.medium)).lineLimit(1)
                        .help(csvURL?.path ?? "請選擇 ePacking List CSV")
                        .textSelection(.enabled)
                    Text(csvURL == nil ? "匯入 ePacking List 的 CSV 檔案" : "已就緒，可製作退料文件")
                        .font(.caption).foregroundStyle(.secondary)
                }
                Spacer(minLength: 10)
                Button(csvURL == nil ? "選擇檔案" : "更換檔案", systemImage: "folder") { showingImporter = true }
                    .buttonStyle(.glass)
            }
            .padding(.vertical, 8)
        }
    }

    private var actionArea: some View {
        GroupBox("03  製作與核對") {
            VStack(alignment: .leading, spacing: 14) {
                if let path = runner.outputPath {
                    HStack(alignment: .top, spacing: 12) {
                        Image(systemName: "checkmark.circle.fill").font(.title2).foregroundStyle(.green)
                        VStack(alignment: .leading, spacing: 6) {
                            Text("退料文件已產出").font(.headline)
                            Text(URL(fileURLWithPath: path).lastPathComponent)
                                .font(.callout).textSelection(.enabled).help(path)
                            Text("請開啟 Excel，人工核對文件內容。")
                                .font(.callout).foregroundStyle(.secondary)
                        }
                    }
                } else {
                    Text(returnType.rawValue).font(.headline)
                    Text("文件儲存至本機下載項目，產出後仍需人工核對。")
                        .font(.callout).foregroundStyle(.secondary)
                }
                Button {
                    guard let csvURL else { return }
                    runner.run(returnType: returnType, csvURL: csvURL)
                } label: {
                    HStack {
                        if runner.isRunning { ProgressView().controlSize(.small) }
                        Text(runner.isRunning ? runner.status : "生成 Excel 退料文件")
                        Spacer()
                        Image(systemName: runner.isRunning ? "hourglass" : "arrow.right")
                    }
                    .frame(maxWidth: .infinity).padding(.vertical, 4)
                }
                .buttonStyle(.glassProminent).controlSize(.large)
                .disabled(csvURL == nil || runner.isRunning)
                if runner.outputPath != nil {
                    Button("打開結果", systemImage: "arrow.up.forward.app") { runner.openOutput() }
                        .buttonStyle(.glass)
                }
                ForEach(runner.warnings, id: \.self) { warning in
                    Label(warning, systemImage: "exclamationmark.triangle.fill")
                        .font(.callout).foregroundStyle(.orange).textSelection(.enabled)
                }
            }
        }
    }

    private func step(_ number: String, title: String, detail: String) -> some View {
        HStack(alignment: .top, spacing: 10) {
            Text(number).font(.caption.weight(.semibold).monospacedDigit())
                .foregroundStyle(.blue).padding(8)
                .background(Color.blue.opacity(0.10), in: .circle)
            VStack(alignment: .leading, spacing: 6) {
                Text(title).font(.callout.weight(.semibold))
                Text(detail).font(.caption).foregroundStyle(.secondary)
            }
        }
        .frame(maxWidth: .infinity, alignment: .leading)
    }

    private var errorBinding: Binding<Bool> {
        Binding(get: { runner.errorMessage != nil }, set: { if !$0 { runner.errorMessage = nil } })
    }
}
