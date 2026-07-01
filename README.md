# Zensi UGC Analytics Tool

Recurring public-account analytics tool for Zensi creator reporting across Instagram, TikTok, and YouTube.

It scrapes creator accounts directly from the roster in `data/creators.zensi.json`. Analytics CSVs can be imported as optional supplements, but the tool does not depend on them.

## What It Exports

- Streamlit dashboard
- latest scheduled snapshot for the dashboard
- Word report
- Excel workbook
- post-level CSV
- creator summary CSV
- platform summary CSV

## Creator Roster

Creator accounts live in:

```text
data/creators.zensi.json
```

Each creator looks like:

```json
{
  "name": "Creator Name",
  "active": true,
  "instagram": ["@handle"],
  "tiktok": ["@handle"],
  "youtube": ["@handle"]
}
```

Set `"active": false` to pause a creator without deleting them.

## Local Setup

```powershell
cd "C:\Users\Meso\Desktop\Studio Cores\Zensi\Tools\zensi-creator-scraper"
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy config.example.json config.json
```

Make sure `config.json` points to the roster:

```json
"roster_json": "C:\\Users\\Meso\\Desktop\\Studio Cores\\Zensi\\Tools\\zensi-creator-scraper\\data\\creators.zensi.json"
```

## Run The Streamlit App

Set an app password first:

```powershell
$env:APP_PASSWORD = "use-a-long-random-password"
```

```powershell
streamlit run app.py
```

The app opens with the latest scheduled scrape if `data/latest_snapshot.json` exists. It also lets the team run a manual scrape and download Word, Excel, or CSV exports.

## Run Reports Locally

Run the last complete 7-day reporting window:

```powershell
python -m zensi_scraper run-weekly --config config.json
```

Refresh only the Streamlit snapshot:

```powershell
python -m zensi_scraper refresh-cache --config config.json --snapshot data/latest_snapshot.json
```

Run a custom date window:

```powershell
python -m zensi_scraper run `
  --campaign-name Zensi `
  --start-date 2026-06-18 `
  --end-date 2026-06-30 `
  --output-dir "C:\Users\Meso\Desktop\Studio Cores\Zensi\Reporting\weekly" `
  --roster-json "C:\Users\Meso\Desktop\Studio Cores\Zensi\Tools\zensi-creator-scraper\data\creators.zensi.json"
```

## Add Or Import Creators

Add one creator:

```powershell
python -m zensi_scraper add-creator `
  --roster-json data\creators.zensi.json `
  --name "New Creator" `
  --instagram "@ig_handle" `
  --tiktok "@tt_handle" `
  --youtube "@yt_handle"
```

Import handles from a roster CSV:

```powershell
python -m zensi_scraper import-roster `
  --from-csv "C:\path\to\creators.csv" `
  --output data\creators.zensi.json
```

That CSV is only for creator/account ingestion. It is not an analytics dependency.

## Optional Analytics Supplements

If a platform blocks anonymous public scraping or you want to compare against internal exports, add supplemental analytics CSVs in `config.json`:

```json
"supplemental_analytics_csvs": [
  "C:\\path\\to\\extra-platform-export.csv"
]
```

Direct public scrape remains primary during dedupe; supplement rows fill gaps.

## Online Setup: Streamlit + GitHub Actions

Use this setup for a free hosted dashboard and recurring refreshes:

1. Push this folder to a private GitHub repo.
2. In Streamlit Community Cloud, deploy `app.py` from that repo.
3. In Streamlit secrets, set `APP_PASSWORD` to a long random password.
4. In GitHub repo settings, allow Actions to read and write repository contents.
5. Keep `.github/workflows/weekly-report.yml` enabled.

The workflow refreshes `data/latest_snapshot.json` every 2 hours:

```yaml
cron: "17 */2 * * *"
```

It also builds formal Friday report artifacts:

```yaml
cron: "47 12 * * 5"
```

When GitHub Actions commits a refreshed `data/latest_snapshot.json`, Streamlit redeploys the dashboard from the latest repo state.

## Data Verification And Integrity

Every post row carries:

- `verification_status`
- `verified_metrics`
- `metric_sources`
- `verified_at`
- `integrity_notes`

Platform verification rules:

- YouTube: RSS discovers recent videos, then `yt-dlp` checks the live public page for views and likes. Comments stay `N/A` unless YouTube exposes a comment count.
- TikTok: TikWM discovers account/post rows and returns views, likes, comments, saves, and shares. `yt-dlp` checks the live public post page as a second source when available.
- Instagram: public feed rows are used when anonymous access is available. oEmbed/page meta verifies owner, caption, date, likes, and comments.

The tool does not convert missing metrics into zero. A metric is treated as verified only when a public source returned that metric.

## Security Defaults

- The Streamlit app is password-gated and will not run without `APP_PASSWORD`.
- Keep the GitHub repo private.
- `.streamlit/secrets.toml`, `.env`, `config.json`, report files, logs, caches, and local exports are ignored.
- `data/latest_snapshot.json` contains campaign analytics and should be treated as private.
- Do not make this repo public unless you remove campaign data first.

## Windows Friday Schedule

For a laptop/desktop recurring setup:

```powershell
.\scripts\install_windows_task.ps1
```

Run it manually:

```powershell
Start-ScheduledTask -TaskName "Zensi Weekly Creator Analytics"
```

## Sharing The Tool

Create a clean zip:

```powershell
Compress-Archive -Path "C:\Users\Meso\Desktop\Studio Cores\Zensi\Tools\zensi-creator-scraper" -DestinationPath "$env:USERPROFILE\Downloads\zensi-creator-scraper.zip" -Force
```

Your teammate can run it locally with Streamlit, or push it to GitHub and connect it to Streamlit Community Cloud.

## Reality Checks

- No creator login is used.
- Instagram anonymous scraping is rate-limited. The tool can scrape public profile and post meta, but feed rows may sometimes return zero.
- TikTok public endpoints can change. Validate against platform exports periodically.
- YouTube RSS is stable for recent videos, but likes can be unavailable on some pages.
- GitHub scheduled workflows are cron-based but can be delayed on shared infrastructure, so avoid treating the exact minute as guaranteed.
