import SwiftUI

struct ContentView: View {
    @StateObject private var model = WebViewModel()

    var body: some View {
        NavigationStack {
            ZStack {
                WebView(model: model)
                    .ignoresSafeArea(edges: .bottom)

                if model.isLoading {
                    ProgressView()
                        .scaleEffect(1.2)
                }
            }
            .navigationTitle("LeagueBoss")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .navigationBarLeading) {
                    Button {
                        model.goBack()
                    } label: {
                        Image(systemName: "chevron.left")
                    }
                    .disabled(!model.canGoBack)
                }
                ToolbarItem(placement: .navigationBarTrailing) {
                    Button {
                        model.reload()
                    } label: {
                        Image(systemName: "arrow.clockwise")
                    }
                }
            }
            .alert(
                "Couldn't Load LeagueBoss",
                isPresented: Binding(
                    get: { model.loadError != nil },
                    set: { newValue in if !newValue { model.loadError = nil } }
                )
            ) {
                Button("Retry") { model.load(AppConfig.siteURL) }
                Button("OK", role: .cancel) {}
            } message: {
                Text(model.loadError ?? "")
            }
        }
        .sheet(item: $model.downloadedFile) { file in
            ShareSheet(activityItems: [file.url])
        }
        .onAppear {
            model.load(AppConfig.siteURL)
        }
    }
}

#Preview {
    ContentView()
}
