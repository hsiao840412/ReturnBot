import AppKit
import SwiftUI

enum AppSurface {
    static func window(_ scheme: ColorScheme) -> Color {
        scheme == .dark ? Color(red: 0.10, green: 0.11, blue: 0.13) : Color(red: 0.94, green: 0.95, blue: 0.97)
    }

    static func panel(_ scheme: ColorScheme) -> Color {
        scheme == .dark ? Color(red: 0.15, green: 0.16, blue: 0.18) : .white
    }

    static func input(_ scheme: ColorScheme) -> Color {
        scheme == .dark ? Color(red: 0.08, green: 0.09, blue: 0.11) : .white
    }
}

/// An opaque adaptive background shared by both workflows.
struct TahoeWindowBackground: NSViewRepresentable {
    func makeNSView(context: Context) -> NSView {
        let view = OpaqueWindowView()
        view.wantsLayer = true
        view.refreshBackground()
        return view
    }

    func updateNSView(_ nsView: NSView, context: Context) {
        (nsView as? OpaqueWindowView)?.refreshBackground()
    }
}

private final class OpaqueWindowView: NSView {
    override var isOpaque: Bool { true }

    private var solidColor: NSColor {
        effectiveAppearance.bestMatch(from: [.darkAqua, .aqua]) == .darkAqua
            ? NSColor(srgbRed: 0.10, green: 0.11, blue: 0.13, alpha: 1)
            : NSColor(srgbRed: 0.94, green: 0.95, blue: 0.97, alpha: 1)
    }

    override func draw(_ dirtyRect: NSRect) {
        solidColor.setFill()
        bounds.fill()
    }

    override func viewDidMoveToWindow() {
        super.viewDidMoveToWindow()
        refreshBackground()
    }

    override func viewDidChangeEffectiveAppearance() {
        super.viewDidChangeEffectiveAppearance()
        refreshBackground()
    }

    func refreshBackground() {
        layer?.backgroundColor = solidColor.cgColor
        needsDisplay = true
        guard let window else { return }
        window.isOpaque = true
        window.alphaValue = 1
        window.backgroundColor = solidColor
        window.titlebarAppearsTransparent = false
    }
}
