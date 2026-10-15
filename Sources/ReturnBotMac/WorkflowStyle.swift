import SwiftUI

struct WorkflowBackground: View {
    @Environment(\.colorScheme) private var colorScheme

    var body: some View {
        AppSurface.window(colorScheme)
            .overlay {
                ZStack {
                    Circle().fill(Color.blue.opacity(0.13))
                        .frame(width: 390, height: 390).blur(radius: 115)
                        .offset(x: 310, y: -220)
                    Circle().fill(Color.purple.opacity(0.08))
                        .frame(width: 330, height: 330).blur(radius: 120)
                        .offset(x: -330, y: 250)
                    LinearGradient(colors: [.white.opacity(0.025), .black.opacity(0.08)],
                                   startPoint: .top, endPoint: .bottom)
                }
                .allowsHitTesting(false)
                .accessibilityHidden(true)
            }
            .background(TahoeWindowBackground())
            .ignoresSafeArea()
    }
}

struct WorkflowHeader: View {
    let title: String
    let subtitle: String

    var body: some View {
        VStack(alignment: .leading, spacing: 4) {
            Text(title).font(.system(size: 23, weight: .bold, design: .rounded))
            Text(subtitle).font(.callout).foregroundStyle(.secondary)
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding(.horizontal, 30).padding(.vertical, 20)
    }
}

struct WorkflowCardStyle: GroupBoxStyle {
    @Environment(\.colorScheme) private var colorScheme

    func makeBody(configuration: Configuration) -> some View {
        VStack(alignment: .leading, spacing: 12) {
            configuration.label.font(.subheadline.weight(.semibold)).foregroundStyle(.secondary)
            configuration.content.frame(maxWidth: .infinity, alignment: .leading)
        }
        .padding(16)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(AppSurface.panel(colorScheme), in: .rect(cornerRadius: 19))
        .glassEffect(.regular, in: .rect(cornerRadius: 19))
    }
}

struct WorkflowField<Content: View>: View {
    let title: String
    @ViewBuilder let content: Content

    var body: some View {
        VStack(alignment: .leading, spacing: 6) {
            Text(title).font(.caption.weight(.semibold)).foregroundStyle(.secondary)
            content
        }
        .frame(maxWidth: .infinity, alignment: .leading)
    }
}

struct WorkflowEmptyState: View {
    let symbol: String
    let title: String
    let detail: String

    var body: some View {
        VStack(spacing: 10) {
            Image(systemName: symbol).font(.system(size: 28)).foregroundStyle(.blue)
            Text(title).font(.headline)
            Text(detail).font(.callout).foregroundStyle(.secondary).multilineTextAlignment(.center)
        }
        .frame(maxWidth: .infinity).padding(.vertical, 24)
    }
}

struct WorkflowStatusBar: View {
    let text: String
    let symbol: String
    var color: Color = .secondary
    var busy = false
    let detail: String

    var body: some View {
        HStack(spacing: 10) {
            if busy { ProgressView().controlSize(.small) }
            Label(text, systemImage: symbol)
                .foregroundStyle(color).lineLimit(1).help(text)
            Spacer(minLength: 12)
            Text(detail).foregroundStyle(.secondary).lineLimit(1).help(detail)
        }
        .font(.callout).textSelection(.enabled)
        .padding(.horizontal, 30).padding(.vertical, 12)
    }
}

/// One sentence explains what to do next; the icon supplements the wording.
struct WorkflowNotice: View {
    let title: String
    let detail: String
    var symbol = "arrow.right.circle.fill"
    var color: Color = .blue

    var body: some View {
        HStack(alignment: .top, spacing: 12) {
            Image(systemName: symbol).font(.title3).foregroundStyle(color)
            VStack(alignment: .leading, spacing: 5) {
                Text(title).font(.callout.weight(.semibold))
                Text(detail).font(.callout).foregroundStyle(.secondary).fixedSize(horizontal: false, vertical: true)
            }
            Spacer(minLength: 0)
        }
        .padding(14).frame(maxWidth: .infinity, alignment: .leading)
        .background(color.opacity(0.08), in: .rect(cornerRadius: 12))
    }
}

struct WorkflowStepLabel: View {
    let number: Int
    let title: String
    let selected: Bool
    var complete = false

    var body: some View {
        HStack(spacing: 9) {
            Image(systemName: complete ? "checkmark.circle.fill" : "\(number).circle.fill")
                .font(.title3)
            Text(title).font(.callout.weight(selected ? .semibold : .regular))
        }
        .foregroundStyle(selected ? Color.blue : complete ? Color.green : Color.secondary)
        .padding(.horizontal, 12).padding(.vertical, 10)
        .frame(maxWidth: .infinity, alignment: .leading)
        .background(selected ? Color.blue.opacity(0.09) : Color.clear, in: .rect(cornerRadius: 12))
        .accessibilityElement(children: .combine)
    }
}
