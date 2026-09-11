import Foundation
import WebKit
import Combine

struct IdentifiableURL: Identifiable {
    let id = UUID()
    let url: URL
}

/// Owns the WKWebView instance and drives loading state, navigation and
/// file downloads (used by the app's "Export to Excel" buttons).
final class WebViewModel: NSObject, ObservableObject {
    let webView: WKWebView

    @Published var isLoading = false
    @Published var canGoBack = false
    @Published var loadError: String?
    @Published var downloadedFile: IdentifiableURL?

    private var pendingDownloadURL: URL?

    override init() {
        let config = WKWebViewConfiguration()
        // Persist cookies (so your login session survives closing the app).
        config.websiteDataStore = .default()
        webView = WKWebView(frame: .zero, configuration: config)
        super.init()
        webView.navigationDelegate = self
        webView.uiDelegate = self
        webView.allowsBackForwardNavigationGestures = true
    }

    func load(_ url: URL) {
        webView.load(URLRequest(url: url))
    }

    func reload() {
        webView.reload()
    }

    func goBack() {
        webView.goBack()
    }
}

// MARK: - Navigation

extension WebViewModel: WKNavigationDelegate {
    func webView(_ webView: WKWebView, didStartProvisionalNavigation navigation: WKNavigation!) {
        isLoading = true
        loadError = nil
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        isLoading = false
        canGoBack = webView.canGoBack
    }

    func webView(_ webView: WKWebView, didFail navigation: WKNavigation!, withError error: Error) {
        isLoading = false
        loadError = error.localizedDescription
    }

    func webView(_ webView: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
        isLoading = false
        loadError = error.localizedDescription
    }

    /// Detect the Excel export downloads (Content-Disposition: attachment) and
    /// hand them to WKDownload instead of trying to render them as a page.
    func webView(
        _ webView: WKWebView,
        decidePolicyFor navigationResponse: WKNavigationResponse,
        decisionHandler: @escaping (WKNavigationResponsePolicy) -> Void
    ) {
        if isAttachment(navigationResponse.response) {
            decisionHandler(.download)
        } else {
            decisionHandler(.allow)
        }
    }

    func webView(_ webView: WKWebView, navigationResponse: WKNavigationResponse, didBecome download: WKDownload) {
        download.delegate = self
    }

    func webView(_ webView: WKWebView, navigationAction: WKNavigationAction, didBecome download: WKDownload) {
        download.delegate = self
    }

    private func isAttachment(_ response: URLResponse) -> Bool {
        guard let http = response as? HTTPURLResponse,
              let disposition = http.value(forHTTPHeaderField: "Content-Disposition") else {
            return false
        }
        return disposition.lowercased().contains("attachment")
    }
}

// MARK: - UI (target=_blank links, JS alerts)

extension WebViewModel: WKUIDelegate {
    func webView(
        _ webView: WKWebView,
        createWebViewWith configuration: WKWebViewConfiguration,
        for navigationAction: WKNavigationAction,
        windowFeatures: WKWindowFeatures
    ) -> WKWebView? {
        // Open links that try to open a new window/tab in the same webview.
        if navigationAction.targetFrame == nil {
            webView.load(navigationAction.request)
        }
        return nil
    }
}

// MARK: - Downloads (Excel export)

extension WebViewModel: WKDownloadDelegate {
    func download(
        _ download: WKDownload,
        decideDestinationUsingResponse response: URLResponse,
        suggestedFilename: String,
        completionHandler: @escaping (URL?) -> Void
    ) {
        let destination = FileManager.default.temporaryDirectory.appendingPathComponent(suggestedFilename)
        try? FileManager.default.removeItem(at: destination)
        pendingDownloadURL = destination
        completionHandler(destination)
    }

    func downloadDidFinish(_ download: WKDownload) {
        DispatchQueue.main.async {
            if let url = self.pendingDownloadURL {
                self.downloadedFile = IdentifiableURL(url: url)
            }
            self.pendingDownloadURL = nil
        }
    }

    func download(_ download: WKDownload, didFailWithError error: Error, resumeData: Data?) {
        DispatchQueue.main.async {
            self.loadError = error.localizedDescription
            self.pendingDownloadURL = nil
        }
    }
}
