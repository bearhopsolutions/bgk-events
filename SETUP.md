# BGK Events — Architecture

This site is now split across three pieces:

- **Frontend** (`index.html`, this repo) — hosted on GitHub Pages. What people actually visit: login, home page, event registration, Weight Loss Challenge signup/weigh-ins.
- **Backend API** (`BGK Events Bot` project on bearhop) — a small Express server. Handles login, event listing, registration, and weigh-in submissions. Talks to Google Sheets via a service account (no OAuth prompts, no Apps Script). Reachable at `https://bearhop.taileae8f4.ts.net` via Tailscale Funnel.
- **Discord bot** (same project on bearhop, one process) — handles *all* admin work: `/create-event`, `/edit-event`, `/events` to list them, `/mark-paid`, and a "Verify" button posted alongside every weigh-in photo (which land in a private `#weigh-in-photos` channel in the BGK Events Discord server).

## Admin workflow (Discord)
- `/create-event name:<name> module:<Weight Loss Challenge|Camping|...> visibility:<All|Enrolled|Hidden>`
- `/edit-event event:<pick from list> regstart:<YYYY-MM-DD> regend:<YYYY-MM-DD> enddate:<YYYY-MM-DD> ...` — sets the registration window and archive date. Status (Enrolling/Active/Archived) is computed automatically from these, same as before.
- `/events` — lists every event and its current computed status.
- `/mark-paid event:<pick> user:<pick> paid:<true/false>`
- Weigh-in photos post automatically to `#weigh-in-photos` with a **Verify** button — click it to mark that weigh-in verified (feeds the leaderboard).

## Adding users
Open the Sheet's `Users` tab and add a row: Email, Name, PIN, IsAdmin (TRUE only for admins — admin status only matters for who can see "Hidden"/"Enrolled"-only events, since all create/edit/verify actions now happen in Discord, not the website).

## Forgot PIN
No self-service reset — if someone forgets their PIN, an admin looks it up/changes it directly in the `Users` tab.

## Deploying changes
- **Frontend** (`index.html`): edit, commit, push — GitHub Pages picks it up within a minute or two.
- **Backend/bot** (on bearhop, `/home/bearhop/Desktop/BGKEventsBot`): edit locally, `scp` the changed files over, then:
  ```
  ssh bearhop
  cd /home/bearhop/Desktop/BGKEventsBot
  sudo docker compose up -d --build
  ```

## Monitoring
- The bot writes a heartbeat row to the Sheet's `BotStatus` tab every 30 seconds.
- A separate `bgk-events-watchdog` container checks that heartbeat every minute and sends a Telegram DOWN/RECOVERED alert (same Telegram bot/chat as UO Outlands, prefixed `[BGK Events]`) if it goes stale for more than 2 minutes.

## What's not built yet
Camping, Blind Date with a Book, Christmas Party, and BGK Creations are selectable as event modules but only show a "coming soon" placeholder when someone views them — same as before, just Weight Loss Challenge is live so far.
