import XCTest
import AppKit
@testable import ReturnBotMac

final class PerformanceTests: XCTestCase {
    func testColumnsRemainOnTheirOwnRows() {
        let cells = [
            RecallOCR.Cell(rect: CGRect(x: 0.8, y: 0.79, width: 0.05, height: 0.04), text: "2"),
            RecallOCR.Cell(rect: CGRect(x: 0.1, y: 0.8, width: 0.3, height: 0.04), text: "TA661-42901"),
            RecallOCR.Cell(rect: CGRect(x: 0.1, y: 0.6, width: 0.3, height: 0.04), text: "661-12345"),
            RecallOCR.Cell(rect: CGRect(x: 0.8, y: 0.6, width: 0.05, height: 0.04), text: "1")
        ]
        XCTAssertEqual(RecallOCR.joinRows(cells), "TA661-42901\t2\n661-12345\t1")
        XCTAssertEqual(RecallOCR.joinRows([]), "")
    }

    @MainActor
    func testScreenshotRecognitionAndCache() async throws {
        let bitmap = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: 2200, pixelsHigh: 1600,
                                      bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true,
                                      isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
        NSGraphicsContext.saveGraphicsState()
        NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: bitmap)
        NSColor.white.setFill()
        NSRect(x: 0, y: 0, width: 2200, height: 1600).fill()
        for row in 0..<28 {
            let attributes: [NSAttributedString.Key: Any] = [.font: NSFont.monospacedSystemFont(ofSize: 30, weight: .regular), .foregroundColor: NSColor.black]
            ("TA661-\(42901 + row)" as NSString).draw(at: NSPoint(x: 90, y: 1500 - row * 50), withAttributes: attributes)
            ("\(row % 3 + 1)" as NSString).draw(at: NSPoint(x: 1800, y: 1500 - row * 50), withAttributes: attributes)
        }
        NSGraphicsContext.restoreGraphicsState()
        let data = bitmap.representation(using: .png, properties: [:])!
        let service = RecallOCR()
        let fast = try await service.recognize(data)
        let cached = try await service.recognize(data)
        XCTAssertFalse(fast.cached)
        XCTAssertTrue(cached.cached)
        XCTAssertEqual(cached.text, fast.text)
        for row in 0..<28 {
            XCTAssertTrue(fast.text.contains("TA661-\(42901 + row)"), fast.text)
        }
    }
}
