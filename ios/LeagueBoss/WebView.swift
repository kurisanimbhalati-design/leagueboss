import SwiftUI
import WebKit

/// SwiftUI wrapper around the WKWebView owned by WebViewModel, with
/// pull-to-refresh wired up.
struct WebView: UIViewRepresentable {
    @ObservedObject var model: WebViewModel

    func makeUIView(context: Context) -> WKWebView {
        let webView = model.webView
        webView.scrollView.bounces = true

        let refreshControl = UIRefreshControl()
        refreshControl.addTarget(context.coordinator, action: #selector(Coordinator.refresh), for: .valueChanged)
        webView.scrollView.refreshControl = refreshControl

        return webView
    }

    func updateUIView(_ uiView: WKWebView, context: Context) {
        // Nothing to sync every render — WebViewModel drives all state.
    }

    func makeCoordinator() -> Coordinator {
        Coordinator(model: model)
    }

    final class Coordinator: NSObject {
        let model: WebViewModel
        init(model: WebViewModel) { self.model = model }

        @objc func refresh(_ sender: UIRefreshControl) {
            model.reload()
            sender.endRefreshing()
        }
    }
}
