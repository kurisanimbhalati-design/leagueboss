import Foundation

/// App-wide configuration.
enum AppConfig {
    /// The URL of your deployed LeagueBoss backend (see ../README.md for how to deploy it).
    /// Replace this with your real Render/Railway URL before running the app.
    static let baseURLString = "https://YOUR-APP-NAME.onrender.com"

    static var siteURL: URL {
        guard let url = URL(string: baseURLString) else {
            fatalError("AppConfig.baseURLString is not a valid URL: \(baseURLString)")
        }
        return url
    }
}
