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
        ZStack {
            background

            VStack(spacing: 17) {
                header
                typePicker
                filePicker
                actionArea
            }
            .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .top)
            .padding(.horizontal, 30)
            .padding(.top, 24)
            .padding(.bottom, 20)
        }
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

    private var background: some View { WorkflowBackground() }
    private var header: some View {
        WorkflowHeader(title: "退料機器人", subtitle: "ReturnBot · 退料文件製作", symbol: "arrow.uturn.backward")
    }

    private var typePicker: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("退料類型")
                .font(.subheadline.weight(.semibold))
                .foregroundStyle(.secondary)

            GlassEffectContainer(spacing: 7) {
                VStack(spacing: 7) {
                ForEach(ReturnType.allCases) { type in
                    Button {
                        withAnimation(.snappy(duration: 0.22)) {
                        returnType = type
                        }
                    } label: {
                        HStack(spacing: 11) {
                            Image(systemName: type.symbol)
                                .frame(width: 22)
                            Text(type.rawValue)
                            Spacer()
                            if returnType == type {
                                Image(systemName: "checkmark.circle.fill")
                                    .foregroundStyle(.white)
                            }
                        }
                        .frame(maxWidth: .infinity)
                        .frame(height: 43)
                        .padding(.horizontal, 15)
                        .contentShape(.rect)
                        .foregroundStyle(returnType == type ? Color.white : Color.primary)
                        .glassEffect(
                            returnType == type
                                ? .regular.tint(.blue).interactive()
                                : .regular.interactive(),
                            in: .rect(cornerRadius: 15)
                        )
                    }
                    .buttonStyle(.plain)
                    .frame(maxWidth: .infinity)
                    .contentShape(.rect)
                }
            }
            }
        }
    }

    private var filePicker: some View {
        VStack(alignment: .leading, spacing: 12) {
            Text("ePacking List")
                .font(.subheadline.weight(.semibold))
                .foregroundStyle(.secondary)

            GlassEffectContainer {
            HStack(spacing: 12) {
                Image(systemName: csvURL == nil ? "doc.text" : "doc.text.fill")
                    .font(.system(size: 20, weight: .medium))
                    .frame(width: 32)
                    .foregroundStyle(csvURL == nil ? Color.secondary : Color.blue)
                VStack(alignment: .leading, spacing: 2) {
                    Text(csvURL?.lastPathComponent ?? "尚未選擇 CSV")
                        .lineLimit(1)
                    if csvURL != nil {
                        Text("已就緒")
                            .font(.caption)
                            .foregroundStyle(.secondary)
                    }
                }
                Spacer()
                Button("選擇檔案", systemImage: "folder") { showingImporter = true }
                    .buttonStyle(.glass)
            }
            .padding(.leading, 17)
            .padding(.trailing, 10)
            .frame(height: 66)
            .glassEffect(.regular, in: .rect(cornerRadius: 19))
            }
        }
    }

    private var actionArea: some View {
        VStack(spacing: 13) {
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
                .frame(maxWidth: .infinity)
                .padding(.vertical, 4)
            }
            .buttonStyle(.glassProminent)
            .controlSize(.large)
            .disabled(csvURL == nil || runner.isRunning)

            HStack {
                Label(runner.status, systemImage: statusSymbol)
                    .foregroundStyle(statusColor)
                Spacer()
                if runner.outputPath != nil {
                    Button("打開結果", systemImage: "arrow.up.forward.app") { runner.openOutput() }
                        .buttonStyle(.glass)
                } else {
                    Label("儲存至下載項目", systemImage: "arrow.down.circle")
                        .foregroundStyle(.secondary)
                }
            }
            .font(.callout)

            ForEach(runner.warnings, id: \.self) { warning in
                Label(warning, systemImage: "exclamationmark.triangle.fill")
                    .font(.caption)
                    .foregroundStyle(.yellow)
                    .frame(maxWidth: .infinity, alignment: .leading)
            }
        }
    }

    private var errorBinding: Binding<Bool> {
        Binding(get: { runner.errorMessage != nil }, set: { if !$0 { runner.errorMessage = nil } })
    }

    private var statusSymbol: String {
        if runner.isRunning { return "circle.dotted" }
        if runner.outputPath != nil { return "checkmark.circle.fill" }
        return "circle.fill"
    }

    private var statusColor: Color {
        if runner.isRunning { return .blue }
        if runner.outputPath != nil { return .green }
        return .secondary
    }
}
