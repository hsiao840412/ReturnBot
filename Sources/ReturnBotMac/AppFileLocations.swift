import Foundation

/// Trial bundles keep their working files separate from the installed app.
/// Release bundles don't set ReturnBotPreview and retain the existing locations.
enum AppFileLocations {
    static var isPreview: Bool { Bundle.main.object(forInfoDictionaryKey: "ReturnBotPreview") as? Bool == true }

    static var recallLibraryURL: URL {
        let folder = isPreview ? "ReturnBot-AdminPreview" : "ReturnBot"
        return FileManager.default.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
            .appendingPathComponent(folder).appendingPathComponent("recall-prices.json")
    }

    static var generationOutputDirectory: URL? {
        guard isPreview else { return nil }
        return FileManager.default.urls(for: .downloadsDirectory, in: .userDomainMask)[0]
            .appendingPathComponent("ReturnBot行政試用", isDirectory: true)
    }

    static var outputFolderLabel: String { isPreview ? "下載項目／ReturnBot行政試用" : "下載項目" }
}
