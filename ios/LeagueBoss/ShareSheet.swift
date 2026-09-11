import SwiftUI
import UIKit

/// Wraps UIActivityViewController so we can offer a "Save to Files" /
/// "Share" sheet for the Excel exports the app downloads.
struct ShareSheet: UIViewControllerRepresentable {
    let activityItems: [Any]

    func makeUIViewController(context: Context) -> UIActivityViewController {
        UIActivityViewController(activityItems: activityItems, applicationActivities: nil)
    }

    func updateUIViewController(_ uiViewController: UIActivityViewController, context: Context) {}
}
