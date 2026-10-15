import AppKit
import SwiftUI
import UniformTypeIdentifiers
import Vision

struct RecallPrice: Codable, Sendable {
    var part: String
    var twd: String
    var description: String
    var usesStockPrice: Bool?
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
    var usesStockPrice: Bool?
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

    var hasUnexportedWork: Bool {
        !input.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty ||
        (outputPath == nil && (!rows.isEmpty || !caseNumber.isEmpty))
    }

    private var libraryURL: URL {
        AppFileLocations.recallLibraryURL
    }

    init() {
        if let data = try? Data(contentsOf: libraryURL) {
            library = try? JSONDecoder().decode(RecallLibrary.self, from: data)
            if library?.entries.contains(where: { Double($0.twd) == 0 && $0.usesStockPrice == nil }) == true {
                library = nil
                message = "舊價格表含交換價格為 0 的料號，請重新匯入價格表，以取得庫存價格。"
            }
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
        rows[index].usesStockPrice = price.usesStockPrice
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
            rows.append(RecallLine(part: part, quantity: quantity, description: price?.description ?? "", twd: price?.twd ?? "", weight: weight, country: country, usesStockPrice: price?.usesStockPrice))
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
        message = "正在辨識截圖…"
        defer { busy = false }
        do {
            let result = try await RecallOCR.shared.recognize(data)
            let text = result.text
            guard !text.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty else {
                throw RecallFailure.message("沒有辨識到文字。請截取較清晰的料號與數量區域再試。")
            }
            image = NSImage(data: data)
            input += (input.isEmpty ? "" : "\n") + text
            message = "已辨識截圖（\(String(format: "%.1f", result.seconds)) 秒\(result.cached ? "，使用快取" : "")）。先核對料號與數量，再按「加入明細」。"
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

    @State private var step: RecallStep = .prepare

    private enum RecallStep: Int, CaseIterable {
        case prepare = 1, review, save
        var title: String {
            switch self {
            case .prepare: "準備資料"
            case .review: "核對零件明細"
            case .save: "確認分單與儲存"
            }
        }
    }

    var body: some View {
        VStack(spacing: 0) {
            HStack(spacing: 8) {
                ForEach(RecallStep.allCases, id: \.rawValue) { item in
                    Button { step = item } label: {
                        WorkflowStepLabel(number: item.rawValue, title: item.title, selected: step == item,
                                          complete: item == .prepare && !model.rows.isEmpty || item == .review && model.result != nil)
                    }
                    .buttonStyle(.plain)
                    .disabled(item == .review && model.rows.isEmpty || item == .save && model.result == nil)
                    .help(item == .save && model.result == nil ? "核對明細後，先按「查看分單結果」" : item.title)
                }
            }
            .padding(.horizontal, 30).padding(.vertical, 12)
            Divider()
            workspace.frame(maxWidth: .infinity, maxHeight: .infinity)
            Divider()
            actionFooter
        }
        .buttonStyle(.glass).controlSize(.regular)
        .groupBoxStyle(WorkflowCardStyle()).textFieldStyle(.roundedBorder)
        .disabled(model.busy)
        .onAppear {
            if model.result != nil { step = .save }
            else if !model.rows.isEmpty { step = .review }
        }
        .onChange(of: model.caseNumber) { invalidatePreview() }
        .onChange(of: model.rate) { invalidatePreview() }
        .onChange(of: model.rows) { invalidatePreview() }
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
        .alert("請檢查資料", isPresented: Binding(get: { model.error != nil }, set: { if !$0 { model.error = nil } })) {
            Button("返回檢查") { model.error = nil }
        } message: { Text(model.error ?? "") }
    }

    @ViewBuilder private var workspace: some View {
        switch step {
        case .prepare:
            ScrollView {
                VStack(alignment: .leading, spacing: 18) {
                    WorkflowNotice(title: "先準備價格表與本次要退回的零件",
                                   detail: "填寫案件資料，再貼上料號與數量或匯入截圖。按「加入明細」後，到下一步逐項核對。")
                    HStack(alignment: .top, spacing: 18) { priceCard; inputCard }
                    if !model.message.isEmpty {
                        Label(model.message, systemImage: "info.circle")
                            .font(.callout).foregroundStyle(.secondary).textSelection(.enabled)
                    }
                }
                .padding(.horizontal, 30).padding(.top, 18).padding(.bottom, 24)
            }
        case .review:
            VStack(alignment: .leading, spacing: 14) {
                HStack(spacing: 12) {
                    VStack(alignment: .leading, spacing: 5) {
                        Text("請逐項核對，再查看分單結果").font(.headline)
                        Text("案件：\(model.caseNumber.isEmpty ? "尚未填寫" : model.caseNumber) · 匯率：\(model.rate.isEmpty ? "尚未填寫" : model.rate)")
                            .font(.callout).foregroundStyle(.secondary).textSelection(.enabled)
                    }
                    Spacer()
                    Button("修改案件／加入零件", systemImage: "pencil") { step = .prepare }
                }
                detailCard.frame(maxHeight: .infinity)
            }
            .padding(.horizontal, 30).padding(.vertical, 18)
        case .save:
            ScrollView {
                VStack(alignment: .leading, spacing: 18) {
                    WorkflowNotice(title: model.outputPath == nil ? "確認分單內容後，再儲存文件" : "文件已儲存，請開啟檢查",
                                   detail: model.outputPath == nil ? "系統依每單 USD 7,500 上限分單。請展開各單核對，確認下方勾選項目，再選擇儲存位置。" : "已建立 Excel 與 DHL 寄件清單。請依原作業流程檢查文件及裝箱資料。",
                                   symbol: model.outputPath == nil ? "doc.text.magnifyingglass" : "folder.badge.checkmark",
                                   color: model.outputPath == nil ? .orange : .green)
                    if let result = model.result { resultCard(result) }
                }
                .padding(.horizontal, 30).padding(.top, 18).padding(.bottom, 24)
            }
        }
    }

    private var actionFooter: some View {
        HStack(spacing: 16) {
            VStack(alignment: .leading, spacing: 4) {
                Text(footerTitle).font(.callout.weight(.semibold))
                Text(footerDetail).font(.caption).foregroundStyle(.secondary)
                    .lineLimit(2).help(footerDetail)
            }
            Spacer(minLength: 12)
            switch step {
            case .prepare:
                Button("前往核對明細", systemImage: "arrow.right") { step = .review }
                    .buttonStyle(.glassProminent).disabled(model.rows.isEmpty)
            case .review:
                Button("上一步") { step = .prepare }
                Button("查看分單結果", systemImage: "arrow.right") {
                    Task {
                        await model.calculate()
                        if model.result != nil { step = .save }
                    }
                }
                .buttonStyle(.glassProminent).disabled(model.rows.isEmpty)
            case .save:
                Button("返回核對明細") { step = .review }
                if let path = model.outputPath {
                    Button("再次儲存…") { chooseOutput() }.disabled(!model.reviewed)
                    Button("開啟已儲存的文件", systemImage: "arrow.up.forward.app") {
                        NSWorkspace.shared.open(URL(fileURLWithPath: path))
                    }.buttonStyle(.glassProminent)
                } else {
                    Button("選擇資料夾並儲存…", systemImage: "folder") { chooseOutput() }
                        .buttonStyle(.glassProminent).disabled(!model.reviewed)
                }
            }
        }
        .controlSize(.large).padding(.horizontal, 30).padding(.vertical, 16)
    }

    private var footerTitle: String {
        if model.busy { return "正在處理，請稍候…" }
        switch step {
        case .prepare: return model.rows.isEmpty ? "尚未加入零件明細" : "已加入 \(model.rows.count) 筆零件明細"
        case .review: return incompleteCount > 0 ? "還有 \(incompleteCount) 筆資料需要補齊" : "共 \(model.rows.count) 筆，請確認內容"
        case .save: return model.outputPath != nil ? "文件已儲存" : model.reviewed ? "已勾選核對項目，可選擇儲存位置" : "請先核對內容並勾選確認項目"
        }
    }

    private var footerDetail: String {
        switch step {
        case .prepare:
            return !model.input.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty ? "輸入區仍有文字，請確認是否已按「加入明細」。" : "加入明細後，按右側按鈕進行核對。"
        case .review: return "可直接修改表格；需要補資料時，返回上一步。"
        case .save: return "每單含 Excel 文件、外箱標籤與 DHL 寄件清單（CSV）。"
        }
    }

    private func invalidatePreview() {
        model.invalidate()
        if step == .save { step = .review }
    }

    private var priceCard: some View {
        GroupBox("價格表與案件資料") {
            VStack(alignment: .leading, spacing: 12) {
                HStack(alignment: .top, spacing: 10) {
                    VStack(alignment: .leading, spacing: 4) {
                        Text(model.library?.source ?? "尚未匯入價格表")
                            .font(.callout.weight(.medium)).lineLimit(1)
                            .help(model.library?.source ?? "請匯入 Excel 或 CSV 價格表")
                        if let library = model.library {
                            Text("\(library.entries.count) 個料號 · \(library.importedAt.prefix(10))")
                                .font(.caption.monospacedDigit()).foregroundStyle(.secondary)
                        } else {
                            Text("Excel / CSV · 資料只保存在此 Mac")
                                .font(.caption).foregroundStyle(.secondary)
                        }
                    }
                    Spacer(minLength: 8)
                    Button("匯入價格表", systemImage: "folder") { importingPrices = true }
                }
                Divider()
                HStack(spacing: 12) {
                    WorkflowField(title: "召回單號") {
                        TextField("例如 NAR9521123", text: $model.caseNumber).accessibilityLabel("召回單號")
                    }
                    WorkflowField(title: "匯率 · 1 美金等於多少台幣") {
                        TextField("1 美金等於多少台幣", text: $model.rate).accessibilityLabel("本次匯率")
                            .monospacedDigit()
                    }
                }
                HStack(spacing: 12) {
                    WorkflowField(title: "新列預設重量") {
                        TextField("0.2", text: $model.weight).accessibilityLabel("新列預設重量").monospacedDigit()
                    }
                    WorkflowField(title: "新列原產地") {
                        TextField("CN", text: $model.country).accessibilityLabel("新列原產地")
                    }
                }
                Text("自動篩選 TWD 價格並合併相同料號；交換價格為 0 時改用庫存價格。")
                    .font(.caption).foregroundStyle(.secondary)
                Text("匯率請輸入本次採用值；重量單位沿用 DHL 貨件設定。")
                    .font(.caption).foregroundStyle(.secondary)
            }
        }
    }

    private var inputCard: some View {
        GroupBox("加入零件清單") {
            VStack(alignment: .leading, spacing: 12) {
                ViewThatFits(in: .horizontal) {
                    HStack(spacing: 12) { imageButtons; Spacer(minLength: 0); addBlankButton }
                    VStack(alignment: .leading, spacing: 10) {
                        HStack(spacing: 12) { imageButtons }
                        addBlankButton
                    }
                }
                if let image = model.image {
                    Image(nsImage: image).resizable().scaledToFit()
                        .frame(maxWidth: .infinity, maxHeight: 140)
                        .accessibilityLabel("待核對的零件截圖")
                }
                TextEditor(text: $model.input)
                    .font(.system(.body, design: .monospaced)).frame(height: 120)
                    .scrollContentBackground(.hidden).padding(8)
                    .background(AppSurface.input(colorScheme), in: .rect(cornerRadius: 12))
                    .overlay(RoundedRectangle(cornerRadius: 12).stroke(Color(nsColor: .separatorColor)))
                    .accessibilityLabel("待加入的料號與數量")
                    .overlay(alignment: .topLeading) {
                        if model.input.isEmpty {
                            Text("貼上料號與數量，或使用上方截圖辨識。\n例如：TA661-42901 2")
                                .font(.callout).foregroundStyle(.tertiary).padding(12).allowsHitTesting(false)
                        }
                    }
                Text("一行一筆，例如 TA661-42901 2。截圖可分次加入，不會自動合併重複料號。")
                    .font(.caption).foregroundStyle(.secondary)
                HStack {
                    Spacer()
                    Button("加入明細", systemImage: "plus.circle") { model.addInput() }
                        .disabled(model.input.isEmpty || model.library == nil)
                }
            }
        }
    }

    @ViewBuilder private var imageButtons: some View {
        Button("貼上截圖", systemImage: "doc.on.clipboard") { Task { await model.pasteImage() } }
        Button("選擇截圖", systemImage: "photo") { importingImage = true }
    }

    private var addBlankButton: some View {
        Button("手動新增一筆", systemImage: "plus") {
            model.rows.append(RecallLine(part: "", quantity: "", description: "", twd: "", weight: model.weight, country: model.country))
            model.invalidate()
        }
    }

    private var detailCard: some View {
        GroupBox("核對要退回的零件") {
            VStack(alignment: .leading, spacing: 12) {
                HStack {
                    Text("請核對料號、數量、單價、重量、原產地及商品描述。")
                        .font(.caption).foregroundStyle(.secondary)
                    Spacer()
                    Text("\(model.rows.count) 筆").font(.caption.monospacedDigit()).foregroundStyle(.secondary)
                }
                if model.rows.isEmpty {
                    WorkflowEmptyState(symbol: "list.bullet.rectangle", title: "尚未加入零件",
                                       detail: "貼上截圖或輸入料號後，核對數量、價格與商品描述。")
                } else {
                    detailTable
                    if incompleteCount > 0 {
                        Label("\(incompleteCount) 筆待補齊數量、價格或商品描述", systemImage: "exclamationmark.triangle.fill")
                            .font(.callout).foregroundStyle(.orange)
                    }
                    if !zeroDollarRows.isEmpty {
                        Text("以下零件的美金單價為 0，請逐筆確認：")
                            .font(.callout.weight(.semibold)).foregroundStyle(.orange)
                        ScrollView {
                            LazyVStack(alignment: .leading, spacing: 10) {
                                ForEach(zeroDollarRows) { row in
                                    Toggle("\(row.part.isEmpty ? "未填料號" : row.part)：確認以 0 美金申報", isOn: line(row).zeroPriceConfirmed)
                                        .font(.callout).foregroundStyle(.orange)
                                }
                            }
                        }.frame(height: min(60, CGFloat(zeroDollarRows.count) * 30))
                    }
                }
            }
        }
    }

    private var zeroDollarRows: [RecallLine] { model.rows.filter(isZeroDollar) }

    private var incompleteCount: Int {
        model.rows.filter { $0.quantity.isEmpty || $0.twd.isEmpty || $0.description.isEmpty }.count
    }

    private func isZeroDollar(_ row: RecallLine) -> Bool {
        guard let price = Double(row.twd), let rate = Double(model.rate), rate > 0 else { return false }
        return (price / rate).rounded() == 0
    }

    // Bind by stable ID so table reuse and deleting a row don't edit a different line.
    private func line(_ row: RecallLine) -> Binding<RecallLine> {
        Binding(
            get: { model.rows.first { $0.id == row.id } ?? row },
            set: { updated in
                if let index = model.rows.firstIndex(where: { $0.id == row.id }) { model.rows[index] = updated }
            }
        )
    }

    private var detailTable: some View {
        let duplicates = Set(Dictionary(grouping: model.rows, by: \.part).filter { $0.value.count > 1 }.keys)
        return Table(model.rows) {
            TableColumn("料號 / 查價") { row in
                HStack(spacing: 6) {
                    TextField("料號", text: Binding(get: { line(row).wrappedValue.part }, set: {
                        var updated = line(row).wrappedValue
                        updated.part = $0
                        updated.twd = ""
                        updated.usesStockPrice = nil
                        updated.description = ""
                        updated.zeroPriceConfirmed = false
                        line(row).wrappedValue = updated
                    }))
                    .accessibilityLabel("料號").help(row.part)
                    Button("查價") { model.lookup(row.id) }.help("查詢此料號的價格與描述")
                    if duplicates.contains(row.part) {
                        Image(systemName: "exclamationmark.triangle.fill").foregroundStyle(.orange)
                            .help("重複料號").accessibilityLabel("重複料號")
                    }
                }.frame(minHeight: 42)
            }.width(min: 180, ideal: 220, max: 280)
            TableColumn("數量") { row in
                TextField("數量", text: line(row).quantity).monospacedDigit().accessibilityLabel("數量")
            }.width(48)
            TableColumn("台幣單價") { row in
                VStack(alignment: .leading, spacing: 3) {
                    TextField("單價", text: Binding(get: { line(row).wrappedValue.twd }, set: {
                        var updated = line(row).wrappedValue
                        updated.twd = $0
                        updated.usesStockPrice = nil
                        updated.zeroPriceConfirmed = false
                        line(row).wrappedValue = updated
                    })).monospacedDigit().accessibilityLabel("台幣單價")
                    if row.usesStockPrice == true {
                        Text("庫存價格").font(.caption2).foregroundStyle(.secondary)
                            .help("交換價格為 0，故使用庫存價格")
                    }
                }
            }.width(80)
            TableColumn("重量") { row in
                TextField("重量", text: line(row).weight).monospacedDigit().accessibilityLabel("重量")
            }.width(56)
            TableColumn("原產地") { row in
                TextField("原產地", text: line(row).country).accessibilityLabel("原產地")
            }.width(52)
            TableColumn("商品描述原文（必填）") { row in
                TextField("商品描述原文", text: line(row).description)
                    .accessibilityLabel("商品描述原文").help(row.description)
            }.width(min: 150, ideal: 250, max: .infinity)
            TableColumn("操作") { row in
                Button(role: .destructive) {
                    model.rows.removeAll { $0.id == row.id }; model.invalidate()
                } label: { Image(systemName: "minus.circle") }
                .buttonStyle(.plain).foregroundStyle(.red)
                .help("移除此筆明細").accessibilityLabel("移除此筆明細")
            }.width(32)
        }
        .tableStyle(.inset)
        .frame(minHeight: 100, maxHeight: .infinity)
        .background(AppSurface.panel(colorScheme), in: .rect(cornerRadius: 12))
        .clipShape(.rect(cornerRadius: 12))
    }

    private func resultCard(_ result: RecallResult) -> some View {
        GroupBox("分單結果與人工核對") {
            VStack(alignment: .leading, spacing: 12) {
                HStack {
                    Label("\(result.groups.count) 組分單", systemImage: "shippingbox")
                    Spacer()
                    Text("\(result.quantity) 件 · USD \(result.total.formatted())")
                        .fontWeight(.semibold).monospacedDigit()
                }.textSelection(.enabled)
                ForEach(Array(result.groups.enumerated()), id: \.offset) { index, group in
                    DisclosureGroup("第 \(index + 1) 單 · \(group.rows.count) 筆 · USD \(group.total.formatted())") {
                        ForEach(Array(group.rows.enumerated()), id: \.offset) { _, row in
                            Text("\(row.part)  × \(row.quantity)  × USD \(row.usd)  =  \(row.total)")
                                .font(.callout.monospacedDigit()).textSelection(.enabled)
                                .frame(maxWidth: .infinity, alignment: .leading)
                        }
                    }
                }
                Text("每單含 KGB Invoice、外箱標籤與 DHL CSV；預設每單一箱，毛重與尺寸請裝箱後填寫。")
                    .font(.caption).foregroundStyle(.secondary)
                Divider()
                Toggle("我已逐項核對數量、價格、商品描述、重量與原產地", isOn: $model.reviewed)

            }
        }
    }

    private func chooseOutput() {
        let panel = NSOpenPanel()
        panel.canChooseFiles = false
        panel.canChooseDirectories = true
        panel.canCreateDirectories = true
        panel.prompt = "儲存至此"
        if panel.runModal() == .OK, let url = panel.url {
            Task {
                let access = url.startAccessingSecurityScopedResource()
                defer { if access { url.stopAccessingSecurityScopedResource() } }
                await model.calculate(exportTo: url)
            }
        }
    }
}
