# Optimus (warehouse scanner)

Barcode/location scanner for Canvasback's warehouse inventory. Owner: Steve (GitHub `dewdad1000000`). Styled to match Harry (warm cream, near-black text, gold accent; Fraunces + IBM Plex).

## How it ships
- Hosted on **GitHub Pages** from `main`: https://dewdad1000000.github.io/new-warehouse-scanner/ . Pushing to `main` publishes in a minute or two.
- Unless the person asks otherwise, test, then commit and push straight to `main`. Always `git pull` first; never force-push.
- This repository is **public**: never commit passwords, API keys or personal data.

## Layout
- `index.html` is the live app: one self-contained file (HTML, CSS, JS). Make changes here.
- `icons/` holds the app icons; `manifest.json` makes it installable on phones. `icon/` is an older duplicate.
- `index_1.html` … `index_10.html` and `index.html.html` are old saved versions from manual uploads. Leave them alone unless Steve asks to clean them up.

## Data
- Scans are sent to a **Google Apps Script web app** (`config.webAppUrl`, saved in the browser as `wh_config_url`), which writes to a Google Sheet. Actions used: `?action=recent`, `?action=recentProducts`, and POSTs for new scans.
- Per-device settings live in localStorage: `wh_config_url`, `wh_config_roster`, `wh_scanner_name`, `wh_active_location`, `wh_camera_id`.
- The Apps Script code itself is not in this repo; changes to it are made in Google Apps Script by Steve. If an app change needs a script change, say exactly what to change there.

## Testing before you push
- Open `index.html` with Playwright (pre-installed) at phone width (390px) and desktop; check for console errors. The camera won't work in tests; check the rest of the UI.
- Don't send test scans to the real web app URL.

## Rules
- It's used on phones in the warehouse: big tap targets, readable in bright light, works one-handed.
- Plain wording for whoever is scanning.
