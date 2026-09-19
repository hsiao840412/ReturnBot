import Foundation
import Vision

actor RecallOCR {
    static let shared = RecallOCR()
    struct Result: Sendable {
        let text: String
        let seconds: Double
        let cached: Bool
    }
    private var previous: (data: Data, text: String)?

    func recognize(_ data: Data) throws -> Result {
        let start = Date()
        if let previous, previous.data == data {
            return Result(text: previous.text, seconds: 0, cached: true)
        }
        let request = VNRecognizeTextRequest()
        request.recognitionLevel = .fast
        request.usesLanguageCorrection = false
        // Keep original pixels: shrinking tall screenshots can erase small part numbers.
        try VNImageRequestHandler(data: data).perform([request])
        let cells = (request.results ?? []).compactMap { observation -> Cell? in
            guard let text = observation.topCandidates(1).first?.string else { return nil }
            return Cell(rect: observation.boundingBox, text: text)
        }
        let text = Self.joinRows(cells)
        // Bound the cache to one screenshot; never cache failed or empty recognition.
        if !text.isEmpty { previous = (data, text) }
        return Result(text: text, seconds: Date().timeIntervalSince(start), cached: false)
    }

    struct Cell: Sendable {
        let rect: CGRect
        let text: String
    }

    static func joinRows(_ cells: [Cell]) -> String {
        var rows: [(y: CGFloat, height: CGFloat, cells: [Cell])] = []
        let sorted = cells.sorted { $0.rect.midY > $1.rect.midY }
        let maxHeight = cells.map(\.rect.height).max() ?? 0
        for cell in sorted {
            let rect = cell.rect
            var match: Int?
            // Sorted baselines let us stop once earlier rows are too far away.
            for index in rows.indices.reversed() {
                let distance = rows[index].y - rect.midY
                if distance > maxHeight * 0.55 { break }
                if distance < min(rows[index].height, rect.height) * 0.55 {
                    match = index
                }
            }
            if let match { rows[match].cells.append(cell) }
            else { rows.append((rect.midY, rect.height, [cell])) }
        }
        return rows.map { $0.cells.sorted { $0.rect.minX < $1.rect.minX }.map(\.text).joined(separator: "\t") }
            .joined(separator: "\n")
    }
}
