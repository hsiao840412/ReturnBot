import AppKit
import SwiftUI
import UniformTypeIdentifiers
import Vision

struct RecallPrice: Codable, Sendable {
    var part: String
    var twd: String
    var description: String
}

struct RecallLibrary: Codable, Sendable {
    var entries: [RecallPrice]
    var source: String
    var sheet: String
    var importedAt: String
    var sourceRows: Int?
    var matchedRows: Int?
    var mergedRows: Int?
    var zeroPriceCount: Int?
}

struct RecallLine: Identifiable, Codable, Equatable, Sendable {
    var id = UUID()
    var part: String
    var quantity: String
    var description: String
    var twd: String
    var weight: String = "0.2"
    var country: String = "CN"
    var zeroPriceConfirmed = false
}

struct RecallOutputLine: Decodable, Sendable {
    let part: String
    let description: String
    let quantity: Int
    let usd: Int
    let total: Int
}

struct RecallGroup: Decodable, Sendable {
    let rows: [RecallOutputLine]
    let total: Int
}

struct RecallResult: Decodable, Sendable {
    let groups: [RecallGroup]
    let quantity: Int
    let total: Int
    var outputPath: String?
}

struct RecallEnvelope<T: Decodable & Sendable>: Decodable, Sendable {
    let success: Bool
    let data: T?
    let message: String?
}

enum RecallFailure: LocalizedError {
    case message(String)
    var errorDescription: String? { if case .message(let text) = self { text } else { nil } }
}

enum RecallBridge {
    static func call<T: Decodable & Sendable>(_ operation: String, payload: Data, as: T.Type) async throws -> T {
        try await Task.detached {
            let process = Process()
            if let resources = Bundle.main.resourceURL,
               FileManager.default.isExecutableFile(atPath: resources.appendingPathComponent("ReturnBotHelper/ReturnBotHelper").path) {
                process.executableURL = resources.appendingPathComponent("ReturnBotHelper/ReturnBotHelper")
                process.arguments = ["--recall", operation]
            } else {
                let root = URL(fileURLWithPath: FileManager.default.currentDirectoryPath)
                let python = root.appendingPathComponent(".venv/bin/python")
                process.executableURL = FileManager.default.isExecutableFile(atPath: python.path) ? python : URL(fileURLWithPath: "/usr/bin/python3")
                process.arguments = [root.appendingPathComponent("returnbot_cli.py").path, "--recall", operation]
            }
            let input = Pipe(), output = Pipe()
            process.standardInput = input
            process.standardOutput = output
            process.standardError = output
            try process.run()
            try input.fileHandleForWriting.write(contentsOf: payload)
            try input.fileHandleForWriting.close()
            let data = output.fileHandleForReading.readDataToEndOfFile()
            process.waitUntilExit()
            // Ignore unrelated diagnostics, but require a complete structured result.
            for line in String(decoding: data, as: UTF8.self).split(separator: "\n").reversed() {
                if let response = try? JSONDecoder().decode(RecallEnvelope<T>.self, from: Data(line.utf8)) {
                    guard response.success, let value = response.data else {
                        throw RecallFailure.message(response.message ?? "處理失敗")
                    }
                    return value
                }
            }
            throw RecallFailure.message("無法取得處理結果：\(String(decoding: data, as: UTF8.self).prefix(500))")
        }.value
    }


}

@MainActor
final class RecallModel: ObservableObject {
    @Published var library: RecallLibrary?
    @Published var rows: [RecallLine] = []
    @Published var input = ""
    @Published var preciseOCR = false
    @Published var caseNumber = ""
    @Published var rate = ""
    @Published var weight = "0.2"
    @Published var country = "CN"
    @Published var busy = false
    @Published var message = ""
    @Published var error: String?
    @Published var result: RecallResult?
    @Published var reviewed = false
    @Published var image: NSImage?
    @Published var outputPath: String?

    private var libraryURL: URL {
        FileManager.default.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
            .appendingPathComponent("ReturnBot/recall-prices.json")
    }

    init() {
        if let data = try? Data(contentsOf: libraryURL) {
            library = try? JSONDecoder().decode(RecallLibrary.self, from: data)
        }
    }

    func invalidate() { result = nil; reviewed = false; outputPath = nil }

    func lookup(_ id: UUID) {
        guard let index = rows.firstIndex(where: { $0.id == id }) else { return }
        let part = rows[index].part.trimmingCharacters(in: .whitespacesAndNewlines).uppercased()
        guard let price = library?.entries.first(where: { $0.part == part }) else {
            error = "價格表查不到 \(part)，請確認料號或手動補齊價格與原文。"; return
        }
        rows[index].part = part
        rows[index].twd = price.twd
        rows[index].description = price.description
        rows[index].zeroPriceConfirmed = false
        invalidate()
    }

    func importPrices(_ url: URL) async {
        busy = true
        defer { busy = false }
        let access = url.startAccessingSecurityScopedResource()
        defer { if access { url.stopAccessingSecurityScopedResource() } }
        do {
            let payload = try JSONSerialization.data(withJSONObject: ["path": url.path])
            let imported = try await RecallBridge.call("prices", payload: payload, as: RecallLibrary.self)
            try FileManager.default.createDirectory(at: libraryURL.deletingLastPathComponent(), withIntermediateDirectories: true)
            try JSONEncoder().encode(imported).write(to: libraryURL, options: .atomic)
            library = imported
            message = "已匯入 \(imported.entries.count) 個料號，合併 \(imported.mergedRows ?? 0) 筆相同資料；含 \(imported.zeroPriceCount ?? 0) 個零價料號。"
        } catch { self.error = error.localizedDescription }
    }

    func addInput() {
        guard let library else { error = "請先匯入價格表"; return }
        let pattern = #"(?i)\b[A-Z0-9]*661-[A-Z0-9]+\b"#
        let regex = try! NSRegularExpression(pattern: pattern)
        var added = 0
        var unmatched: [String] = []
        for line in input.components(separatedBy: .newlines) where !line.trimmingCharacters(in: .whitespaces).isEmpty {
            let ns = line as NSString
            let matches = regex.matches(in: line, range: NSRange(location: 0, length: ns.length))
            let cells = line.split(whereSeparator: { $0 == "\t" || $0 == "," || $0 == " " }).map(String.init)
            let part: String
            if matches.count == 1 { part = ns.substring(with: matches[0].range).uppercased() }
            else if matches.isEmpty, cells.count <= 2, let first = cells.first { part = first.uppercased() }
            else { unmatched.append(line); continue }
            let price = library.entries.first { $0.part == part }
            // Only infer quantity from an unambiguous two-field input. OCR rows
            // containing other columns require explicit quantity review.
            let quantity = cells.count == 2 && cells[0].uppercased() == part && Int(cells[1]) != nil ? cells[1] : ""
            rows.append(RecallLine(part: part, quantity: quantity, description: price?.description ?? "", twd: price?.twd ?? "", weight: weight, country: country))
            added += 1
        }
        input = unmatched.joined(separator: "\n")
        invalidate()
        message = "新增 \(added) 筆；請確認數量與原文。\(unmatched.isEmpty ? "" : "尚有 \(unmatched.count) 行未加入，保留於輸入區。")"
    }

    func pasteImage() async {
        guard !busy else { return }
        guard let data = NSPasteboard.general.data(forType: .png) ?? NSPasteboard.general.data(forType: .tiff) else {
            error = "剪貼簿沒有圖片；文字可直接貼進輸入區。"; return
        }
        await recognize(data)
    }

    func recognize(_ data: Data) async {
        guard !busy else { return }
        busy = true
        message = preciseOCR ? "正在精確辨識截圖…" : "正在快速辨識截圖…"
        defer { busy = false }
        do {
            let result = try await RecallOCR.shared.recognize(data, precise: preciseOCR)
            let text = result.text
            guard !text.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty else {
                throw RecallFailure.message("沒有辨識到文字。可勾選精確辨識，或截取較清晰的料號與數量區域再試。")
            }
            image = NSImage(data: data)
            input += (input.isEmpty ? "" : "\n") + text
            message = "已辨識截圖（\(String(format: "%.1f", result.seconds)) 秒\(result.cached ? "，使用快取" : "")）。先核對料號與數量，再按「加入明細」；辨識不準時可勾選精確辨識。"
        } catch { self.error = error.localizedDescription }
    }

    func calculate(exportTo directory: URL? = nil) async {
        busy = true
        defer { busy = false }
        do {
            let rowData = try JSONSerialization.jsonObject(with: JSONEncoder().encode(rows))
            var payload: [String: Any] = ["case": caseNumber, "rate": rate, "rows": rowData,
                                          "priceSource": library?.source ?? "", "priceImportedAt": library?.importedAt ?? ""]
            if let directory { payload["outputDirectory"] = directory.path }
            let data = try JSONSerialization.data(withJSONObject: payload)
            let value = try await RecallBridge.call(directory == nil ? "preview" : "export", payload: data, as: RecallResult.self)
            result = value
            outputPath = value.outputPath
            message = directory == nil ? "分單預覽已更新，請核對後匯出。" : "已輸出 \(value.groups.count) 組 KGB Invoice、外箱標籤與 DHL CSV。"
        } catch { self.error = error.localizedDescription }
    }
}

struct RecallView: View {
    @Environment(\.colorScheme) private var colorScheme
    @ObservedObject var model: RecallModel
    @State private var importingPrices = false
    @State private var importingImage = false

    var body: some View {
        ScrollView {
            VStack(alignment: .leading, spacing: 18) {
                WorkflowHeader(title: "寄銷召回", subtitle: "整理零件、核對價格，完成分單與匯出", symbol: "shippingbox", beta: true)
                HStack(spacing: 18) {
                    Label("每單上限 USD 5,000", systemImage: "shippingbox")
                    Label("美金單價四捨五入", systemImage: "dollarsign.circle")
                    Spacer()
                    if model.busy { ProgressView().controlSize(.small) }
                }.font(.caption).foregroundStyle(.secondary)
                GroupBox("01  價格表與案件") {
                    VStack(alignment: .leading, spacing: 12) {
                        HStack {
                            Text(model.library.map { "\($0.source) · \($0.entries.count) 筆 · \($0.importedAt.prefix(10))" } ?? "請匯入自己的價格表（Excel / CSV）")
                            Spacer()
                            Button("匯入價格表", systemImage: "folder") { importingPrices = true }
                        }
                        Text("可直接匯入原始零件價格表；自動篩選 TWD 交換價格並合併相同料號。資料只保存在此 Mac。")
                            .font(.caption).foregroundStyle(.secondary)
                        HStack {
                            VStack(alignment: .leading, spacing: 6) {
                                Text("召回單號").font(.caption.weight(.semibold)).foregroundStyle(.secondary)
                                TextField("例如 NAR9521123", text: $model.caseNumber)
                            }
                            VStack(alignment: .leading, spacing: 6) {
                                Text("本次匯率 · TWD / USD").font(.caption.weight(.semibold)).foregroundStyle(.secondary)
                                TextField("1 美金等於多少台幣", text: $model.rate)
                            }
                        }.textFieldStyle(.roundedBorder)
                        HStack {
                            Text("新列預設重量")
                            TextField("0.2", text: $model.weight).frame(width: 65)
                            Text("原產地")
                            TextField("CN", text: $model.country).frame(width: 65)
                            Text("匯率請輸入本次採用值；重量單位沿用 DHL 貨件設定。")
                                .font(.caption).foregroundStyle(.secondary)
                        }
                    }.padding(8)
                }
                GroupBox("02  加入退回零件") {
                    VStack(alignment: .leading, spacing: 10) {
                        HStack {
                            Button("貼上截圖", systemImage: "doc.on.clipboard") { Task { await model.pasteImage() } }
                            Button("選擇截圖", systemImage: "photo") { importingImage = true }
                            Toggle("精確辨識（較慢）", isOn: $model.preciseOCR).toggleStyle(.checkbox)
                            Spacer()
                            Button("新增空白列", systemImage: "plus") {
                                model.rows.append(RecallLine(part: "", quantity: "", description: "", twd: "", weight: model.weight, country: model.country))
                                model.invalidate()
                            }
                        }
                        HStack(alignment: .top) {
                            if let image = model.image {
                                Image(nsImage: image).resizable().scaledToFit().frame(maxWidth: 300, maxHeight: 180)
                            }
                            TextEditor(text: $model.input).font(.system(.body, design: .monospaced)).frame(height: 120)
                                .scrollContentBackground(.hidden)
                                .padding(8)
                                .background(AppSurface.input(colorScheme), in: RoundedRectangle(cornerRadius: 12))
                                .overlay(RoundedRectangle(cornerRadius: 12).stroke(Color(nsColor: .separatorColor)))
                        }
                        HStack {
                            Text("一行一筆，例如 TA661-42901 2。截圖可分次加入；不會自動合併重複料號。")
                                .font(.caption).foregroundStyle(.secondary)
                            Spacer()
                            Button("加入明細", systemImage: "plus.circle") { model.addInput() }.disabled(model.input.isEmpty || model.library == nil)
                        }
                    }.padding(8)
                }
                GroupBox("03  核對明細") {
                    VStack(alignment: .leading, spacing: 8) {
                        Text("商品描述原文會同時用於 Invoice 與 DHL。").font(.caption).foregroundStyle(.secondary)
                        if model.rows.isEmpty {
                            VStack(spacing: 10) {
                                Image(systemName: "list.bullet.rectangle").font(.system(size: 28)).foregroundStyle(.blue)
                                Text("尚未加入零件").font(.headline)
                                Text("貼上截圖或輸入料號後，價格與描述會自動帶入。")
                                    .font(.callout).foregroundStyle(.secondary)
                            }.frame(maxWidth: .infinity).padding(.vertical, 24)
                        } else {
                            HStack {
                                Text("料號 / 查價").frame(width: 215, alignment: .leading)
                                Text("數量").frame(width: 55)
                                Text("台幣單價").frame(width: 90)
                                Text("重量").frame(width: 60)
                                Text("原產地").frame(width: 55)
                                Spacer()
                            }.font(.caption).foregroundStyle(.secondary)
                        }
                        ForEach($model.rows) { $row in
                            VStack(alignment: .leading, spacing: 6) {
                                HStack {
                                    TextField("料號", text: Binding(get: { row.part }, set: {
                                        row.part = $0
                                        row.twd = ""
                                        row.description = ""
                                        row.zeroPriceConfirmed = false
                                    })).frame(width: 160)
                                    Button("查價") { model.lookup(row.id) }
                                    TextField("數量", text: $row.quantity).frame(width: 55)
                                    TextField("台幣單價", text: Binding(get: { row.twd }, set: {
                                        row.twd = $0
                                        row.zeroPriceConfirmed = false
                                    })).frame(width: 90)
                                    TextField("重量", text: $row.weight).frame(width: 60)
                                    TextField("原產地", text: $row.country).frame(width: 55)
                                    if model.rows.filter({ $0.part == row.part }).count > 1 {
                                        Text("重複料號").font(.caption).foregroundStyle(.orange)
                                    }
                                    Spacer()
                                    Button(role: .destructive) { model.rows.removeAll { $0.id == row.id }; model.invalidate() } label: {
                                        Image(systemName: "minus.circle")
                                    }
                                }
                                TextField("商品描述原文（必填）", text: $row.description, axis: .vertical)
                                if row.quantity.isEmpty || row.twd.isEmpty || row.description.isEmpty {
                                    Text("待補齊：數量、價格或商品描述").font(.caption).foregroundStyle(.orange)
                                }
                                if let price = Double(row.twd), let rate = Double(model.rate), rate > 0, (price / rate).rounded() == 0 {
                                    Toggle("此筆為 0 美金：確認以 0 美金申報", isOn: $row.zeroPriceConfirmed)
                                        .font(.caption).foregroundStyle(.orange)
                                }
                            }.textFieldStyle(.roundedBorder).padding(.vertical, 5)
                            Divider()
                        }
                    }.padding(8)
                }
                HStack {
                    Button { Task { await model.calculate() } } label: {
                        HStack {
                            Text("預覽分單")
                            Spacer()
                            Image(systemName: "arrow.right")
                        }.frame(maxWidth: .infinity).padding(.vertical, 4)
                    }.buttonStyle(.glassProminent).controlSize(.large)
                        .disabled(model.rows.isEmpty)
                }
                if !model.message.isEmpty {
                    Label(model.message, systemImage: model.busy ? "hourglass" : "info.circle")
                        .font(.callout).foregroundStyle(.secondary)
                }
                if let result = model.result {
                    GroupBox("04  分單與匯出") {
                        VStack(alignment: .leading, spacing: 10) {
                            ForEach(Array(result.groups.enumerated()), id: \.offset) { index, group in
                                DisclosureGroup("第 \(index + 1) 單 · \(group.rows.count) 筆 · USD \(group.total.formatted())") {
                                    ForEach(Array(group.rows.enumerated()), id: \.offset) { _, row in
                                        Text("\(row.part)  × \(row.quantity)  × USD \(row.usd)  =  \(row.total)").font(.callout.monospacedDigit())
                                    }
                                }
                            }
                            Text("共 \(result.quantity) 件，USD \(result.total.formatted())。每單套用 KGB 範本，含 Invoice 與外箱標籤；預設每單一箱，毛重與尺寸請裝箱後填寫。")
                                .font(.caption).foregroundStyle(.secondary)
                            Toggle("已核對數量、價格、描述、重量與原產地", isOn: $model.reviewed)
                            HStack {
                                Button("選擇資料夾並匯出") { chooseOutput() }.buttonStyle(.glassProminent).disabled(!model.reviewed)
                                if let path = model.outputPath {
                                    Button("打開結果") { NSWorkspace.shared.open(URL(fileURLWithPath: path)) }
                                }
                            }
                        }.padding(8)
                    }
                }
            }.padding(.horizontal, 30).padding(.top, 24).padding(.bottom, 30)
        }
        .background(WorkflowBackground())
        .buttonStyle(.glass)
        .controlSize(.regular)
        .groupBoxStyle(RecallSectionStyle())
        .textFieldStyle(.roundedBorder)
        .disabled(model.busy)
        .onChange(of: model.caseNumber) { model.invalidate() }
        .onChange(of: model.rate) { model.invalidate() }
        .onChange(of: model.rows) { model.invalidate() }
        .fileImporter(isPresented: $importingPrices, allowedContentTypes: [.commaSeparatedText, UTType(filenameExtension: "xlsx")!]) { result in
            if case .success(let url) = result { Task { await model.importPrices(url) } }
        }
        .fileImporter(isPresented: $importingImage, allowedContentTypes: [.image]) { result in
            if case .success(let url) = result {
                Task {
                    let access = url.startAccessingSecurityScopedResource()
                    defer { if access { url.stopAccessingSecurityScopedResource() } }
                    do {
                        let data = try await Task.detached(priority: .userInitiated) { try Data(contentsOf: url) }.value
                        await model.recognize(data)
                    }
                    catch { model.error = error.localizedDescription }
                }
            }
        }
        .alert("請確認", isPresented: Binding(get: { model.error != nil }, set: { if !$0 { model.error = nil } })) {
            Button("好") { model.error = nil }
        } message: { Text(model.error ?? "") }
    }

    private func chooseOutput() {
        let panel = NSOpenPanel()
        panel.canChooseFiles = false
        panel.canChooseDirectories = true
        panel.canCreateDirectories = true
        panel.prompt = "匯出至此"
        if panel.runModal() == .OK, let url = panel.url {
            Task {
                let access = url.startAccessingSecurityScopedResource()
                defer { if access { url.stopAccessingSecurityScopedResource() } }
                await model.calculate(exportTo: url)
            }
        }
    }
}

private struct RecallSectionStyle: GroupBoxStyle {
    @Environment(\.colorScheme) private var colorScheme
    func makeBody(configuration: Configuration) -> some View {
        VStack(alignment: .leading, spacing: 12) {
            configuration.label.font(.subheadline.weight(.semibold)).foregroundStyle(.secondary)
            configuration.content.frame(maxWidth: .infinity, alignment: .leading)
        }
        .padding(16)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(AppSurface.panel(colorScheme), in: RoundedRectangle(cornerRadius: 19))
        .glassEffect(.regular, in: .rect(cornerRadius: 19))
    }
}
