import XCTest
@testable import ReturnBotMac

final class RecallPriceTests: XCTestCase {
    @MainActor
    func testFallbackMarkerFollowsLookupAndInput() {
        let model = RecallModel()
        model.library = RecallLibrary(entries: [
            RecallPrice(part: "TA661-1", twd: "300", description: "Original", usesStockPrice: true),
            RecallPrice(part: "TA661-2", twd: "100", description: "Normal")
        ], source: "test", sheet: "test", importedAt: "test")
        model.input = "TA661-1 2"
        model.addInput()
        XCTAssertEqual(model.rows.first?.twd, "300")
        XCTAssertEqual(model.rows.first?.usesStockPrice, true)
        model.rows[0].part = "TA661-2"
        model.lookup(model.rows[0].id)
        XCTAssertEqual(model.rows[0].twd, "100")
        XCTAssertNil(model.rows[0].usesStockPrice)
    }

    func testExistingPriceEntryDecodesWithoutMarker() throws {
        let data = Data(#"{"part":"TA661-1","twd":"100","description":"Original"}"#.utf8)
        let price = try JSONDecoder().decode(RecallPrice.self, from: data)
        XCTAssertNil(price.usesStockPrice)
    }
}
