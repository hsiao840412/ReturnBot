import SwiftUI

struct WorkflowBackground: View {
    var body: some View {
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
}

struct WorkflowHeader: View {
    let title: String
    let subtitle: String
    let symbol: String
    var beta = false
    var body: some View {
        HStack(spacing: 13) {
            Image(systemName: symbol)
                .font(.system(size: 18, weight: .semibold))
                .frame(width: 40, height: 40)
                .foregroundStyle(.white)
                .glassEffect(.regular.tint(.blue).interactive(), in: .circle)
            VStack(alignment: .leading, spacing: 3) {
                HStack(spacing: 9) {
                    Text(title).font(.system(size: 23, weight: .bold, design: .rounded))
                    if beta {
                        Text("Beta").font(.caption.weight(.semibold)).foregroundStyle(.blue)
                            .padding(.horizontal, 9).padding(.vertical, 4)
                            .glassEffect(.regular.tint(.blue.opacity(0.12)), in: .capsule)
                    }
                }
                Text(subtitle).foregroundStyle(.secondary)
            }
            Spacer()
            HStack(spacing: 7) {
                Circle().fill(.green).frame(width: 7, height: 7)
                Text("v\(Bundle.main.object(forInfoDictionaryKey: "CFBundleShortVersionString") as? String ?? "3.2")")
            }
            .font(.callout.weight(.semibold))
            .padding(.horizontal, 13).padding(.vertical, 7)
            .glassEffect(.regular, in: .capsule)
        }
    }
}
