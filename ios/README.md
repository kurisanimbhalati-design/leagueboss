# Getting LeagueBoss on your iPhone

LeagueBoss is a Python (Flask) web app — it needs a server to run, so it can't
run natively on your phone. The setup here does two things:

1. **Deploy the Flask app** to a free host, so it's reachable at a URL like
   `https://leagueboss-xxxx.onrender.com` from anywhere (Wi-Fi or cellular).
2. **Wrap that URL in a real iOS app** built in Xcode, so it has its own icon
   and opens full-screen like any other app on your phone.

You'll need a Mac with Xcode installed for step 2. Step 1 can be done from
any browser.

---

## Step 1 — Deploy the backend to Render

1. Make sure this repo (with `app.py`, `templates/`, `requirements.txt`,
   `Procfile`, `render.yaml`) is pushed to GitHub — it already is, on the
   `claude/web-app-iphone-logging-eajerp` branch of your `leagueboss` repo.
2. Go to [render.com](https://render.com) and sign up (free).
3. Click **New +** → **Blueprint**, and point it at your GitHub repo. Render
   will read `render.yaml` and set everything up automatically (build command,
   start command, and a `SECRET_KEY`).
4. When it asks for the `ADMIN_PASSWORD` environment variable, set it to a
   password you'll actually use to log in — don't leave the app's default
   (`safa2026`) since anyone with the URL could log in otherwise.
5. Click **Apply** / **Create**. The first deploy takes a couple of minutes.
6. Once it's live, Render gives you a URL like
   `https://leagueboss-xxxx.onrender.com`. Open it in Safari to confirm the
   app loads and you can log in with `admin` / the password you set.

**Data persistence caveat:** on Render's free plan, `league_data.json` lives
on disk that resets whenever you redeploy (pushing new code) — it survives
the app going to sleep from inactivity, just not a redeploy. For personal use
this is usually fine, but if you want results to survive redeploys too,
upgrade the service to a paid plan with a persistent disk, mount it (e.g. at
`/var/data`), and set an environment variable `DATA_FILE=/var/data/league_data.json`
(the app already reads that env var — see `app.py`).

If you'd rather use Railway instead of Render, the same `Procfile` works
there too — just skip `render.yaml` and set the start command to
`gunicorn app:app` manually in Railway's dashboard.

---

## Step 2 — Build the Xcode app

### One-time setup

1. Open **Xcode** → **File → New → Project**.
2. Choose **iOS → App**, click Next.
3. Fill in:
   - **Product Name:** `LeagueBoss`
   - **Interface:** SwiftUI
   - **Language:** Swift
   - Leave "Use Core Data" and "Include Tests" unchecked.
4. Save it anywhere (e.g. right next to this repo, or inside `ios/` — the
   generated `.xcodeproj` doesn't need to be in this repo, but you can add it
   later if you want it version-controlled).
5. Xcode creates `LeagueBossApp.swift`, `ContentView.swift`, and an
   `Assets.xcassets` for you. **Delete** the generated `ContentView.swift`
   and `LeagueBossApp.swift` — you'll use the ones from this folder instead.
6. In Finder, drag these files from `ios/LeagueBoss/` in this repo into your
   Xcode project's file list (drop them into the main `LeagueBoss` group):
   - `AppConfig.swift`
   - `LeagueBossApp.swift`
   - `WebViewModel.swift`
   - `WebView.swift`
   - `ShareSheet.swift`
   - `ContentView.swift`
   - When prompted, check **"Copy items if needed"** and make sure the
     **LeagueBoss target** checkbox is ticked.

### Point it at your deployed app

7. Open `AppConfig.swift` in Xcode and replace the placeholder with your real
   Render URL from Step 1:

   ```swift
   static let baseURLString = "https://leagueboss-xxxx.onrender.com"
   ```

### Run it on your iPhone

8. Plug your iPhone into your Mac (or use wireless debugging), and select it
   as the run destination in Xcode's toolbar (next to the Play/Stop buttons).
9. Click the project name in the file navigator → your target → **Signing &
   Capabilities**:
   - Set **Team** to your Apple ID (Xcode → Settings → Accounts to add one if
     you haven't — a free personal Apple ID works, no paid account needed to
     just run it on your own phone).
   - Change the **Bundle Identifier** to something unique, e.g.
     `com.yourname.leagueboss`.
10. Press **Run (▶)**. Xcode builds, installs, and launches the app on your
    phone.
11. First launch only: on your iPhone, go to **Settings → General → VPN &
    Device Management**, tap your Apple ID under "Developer App", and tap
    **Trust**. Then relaunch LeagueBoss from your home screen.

You now have a **LeagueBoss** icon on your phone that opens your league app
full-screen, with pull-to-refresh, back navigation, and Excel exports handled
via the Share Sheet.

**Free Apple ID limitation:** apps signed with a free account stop opening
after **7 days** — just reopen the project in Xcode and hit Run again to
re-sign it (your data isn't affected, it lives on the server). A paid Apple
Developer account ($99/year) removes that limit and lets you install via
TestFlight without a cable.

---

## Quicker alternative (no Xcode at all)

If you just want something on your home screen today: open your Render URL
in **Safari** on your iPhone, tap the **Share** button, then **Add to Home
Screen**. It behaves almost identically to the native wrapper above (its own
icon, full-screen, no browser chrome) — the only things you lose are the
Excel-export share sheet and offline error handling this Xcode project adds.
