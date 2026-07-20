# BGK Events — Setup

The site is now split in two:
- **Frontend** (`index.html`) — lives in this GitHub repo, hosted for free on GitHub Pages. This is the page people actually visit.
- **Backend** (`Code.gs`) — lives in Google Apps Script, bound to the Google Sheet. It's a JSON API only now (no HTML), used only to log in and read/write the Sheet.

## 1. Create the Google Sheet
1. Go to sheets.google.com → Blank spreadsheet.
2. Rename it "BGK Events DB" (top left).

## 2. Attach the script
1. In the Sheet, click **Extensions → Apps Script**.
2. Delete anything in the default `Code.gs` file, then paste in the full contents of `Code.gs` from this repo.
3. If there's still an `Index.html` file in the project from before, delete it — it's no longer used.
4. Click the save icon (or Ctrl+S).

## 3. Seed the database
1. In the Apps Script editor, find the function dropdown at the top (next to Debug) and select **setup**.
2. Click **Run**.
3. First time only: it'll ask for permissions — click through "Review permissions" → choose your Google account → "Advanced" → "Go to BGK Events (unsafe)" → Allow. (This warning is normal for scripts you write yourself.)
4. Check the Sheet — you should now see three tabs: `Users`, `Events`, `Enrollments`, with Kevin seeded as the first admin.

## 4. Deploy the backend as a web app (JSON API)
1. In the Apps Script editor, click **Deploy → New deployment** (or, if a deployment already exists, **Manage deployments → pencil icon → Version: New version**).
2. Click the gear icon next to "Select type" → **Web app**.
3. Set "Execute as" to **Me**, and "Who has access" to **Anyone**.
4. Click **Deploy**, authorize again if asked.
5. Copy the **Web app URL** — it ends in `/exec`.
6. Open `index.html` in this repo and check the `EXEC_URL` line near the top of the `<script>` — make sure it matches that URL exactly. If it changed, update it and push the change to GitHub.

## 5. Turn on GitHub Pages
1. In the GitHub repo, go to **Settings → Pages**.
2. Under "Source," pick the branch (`main`) and root folder (`/`).
3. Save. GitHub will give you a URL like `https://bearhopsolutions.github.io/bgk-events/` — that's the link everyone uses to reach BGK Events.

## 6. Log in
- Go to the GitHub Pages URL.
- Log in as `kevin412l@hotmail.com` / PIN `041295`.
- You'll see the Admin link in the top right — that's where you add events (name, status, visibility) for now.

## Notes
- To add Jared, Nichole, Abe, Christina etc. as users for now: open the `Users` tab in the Sheet and add a row per person (Email, Name, PIN, IsAdmin=FALSE). We'll build a nicer admin UI for this later if you want.
- Changing `index.html`: just push the change to GitHub — Pages updates automatically within a minute or two.
- Changing `Code.gs`: paste the updated version into the Apps Script editor, then **Deploy → Manage deployments → edit (pencil) → New version → Deploy** to push it live. Just saving doesn't update the live URL.
- The event pages themselves (Biggest Loser Challenge, Camping, etc.) aren't built yet — clicking an event right now shows a "coming soon" placeholder. That's next.
