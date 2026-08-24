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
    @StateObject private var runner = ReturnBotRunner()
    @State private var returnType: ReturnType = .mailIn
    @State private var csvURL: URL?
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
        .task {
            runner.prepareExcelAccess()
        }
    }

    private var background: some View {
        TahoeWindowBackground()
        .overlay {
            ZStack {
                Circle()
                    .fill(Color.blue.opacity(0.13))
                    .frame(width: 390, height: 390)
                    .blur(radius: 115)
                    .offset(x: 310, y: -220)
                Circle()
                    .fill(Color.purple.opacity(0.08))
                    .frame(width: 330, height: 330)
                    .blur(radius: 120)
                    .offset(x: -330, y: 250)
                LinearGradient(
                    colors: [Color.white.opacity(0.025), Color.black.opacity(0.08)],
                    startPoint: .top,
                    endPoint: .bottom
                )
            }
        }
        .ignoresSafeArea()
    }

    private var header: some View {
        HStack(spacing: 13) {
            Image(systemName: "arrow.uturn.backward")
                .font(.system(size: 18, weight: .semibold))
                .frame(width: 40, height: 40)
                .foregroundStyle(.white)
                .glassEffect(.regular.tint(.blue).interactive(), in: .circle)

            VStack(alignment: .leading, spacing: 3) {
                Text("退料機器人")
                    .font(.system(size: 23, weight: .bold, design: .rounded))
                Text("ReturnBot · Excel 自動化")
                    .foregroundStyle(.secondary)
            }
            Spacer()
            HStack(spacing: 7) {
                Circle().fill(.green).frame(width: 7, height: 7)
                Text("v3.0")
            }
            .font(.callout.weight(.semibold))
            .padding(.horizontal, 13)
            .padding(.vertical, 7)
            .glassEffect(.regular, in: .capsule)
        }
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
            .disabled(csvURL == nil || runner.isRunning || runner.isPreparing)

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
        if runner.isRunning || runner.isPreparing { return "circle.dotted" }
        if runner.outputPath != nil { return "checkmark.circle.fill" }
        return "circle.fill"
    }

    private var statusColor: Color {
        if runner.isRunning || runner.isPreparing { return .blue }
        if runner.outputPath != nil { return .green }
        return .secondary
    }
}
