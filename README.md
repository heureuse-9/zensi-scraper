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
python -m zensi_scraper refresh-cache --config config.json --snapshot data/latest_snapshot.json --lookback-days 7
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

Creator/platform exports can include extra owner-only fields such as:

```text
saves, shares, reposts, remixes, remix_views
```

Those fields are carried through to Word, Excel, CSV, and the dashboard when supplied. They are not guessed from public pages.

## Owner API Metrics Without Studio Logging In

The team should not log into creator accounts. For owner-only metrics, use one of these safe paths:

- Creator authorizes the app/OAuth connection themselves.
- Creator sends a platform export CSV/screenshot-derived CSV.
- A manager/business account that already has authorized asset access provides a token.

If you do not want any login or OAuth flow at all, use the creator export path. In the Streamlit app, open `No-Login Owner Metrics Import`, download the CSV template, fill it from the creator's own analytics export/screenshot, and upload it during a manual refresh. The scraper still discovers public posts itself; the import only fills private fields such as IG saves/shares/reposts and YouTube shares/save-to-playlist/remix fields.

Supported secret/config inputs:

```text
INSTAGRAM_GRAPH_ACCESS_TOKEN
INSTAGRAM_GRAPH_TOKENS_JSON
INSTAGRAM_GRAPH_USER_IDS_JSON
INSTAGRAM_GRAPH_VERSION
YOUTUBE_ANALYTICS_ACCESS_TOKEN
YOUTUBE_ANALYTICS_TOKENS_JSON
```

`INSTAGRAM_GRAPH_USER_IDS_JSON` maps handles to IG user IDs, for example:

```json
{"creator_handle": "17841400000000000"}
```

`INSTAGRAM_GRAPH_TOKENS_JSON` and `YOUTUBE_ANALYTICS_TOKENS_JSON` can map handles to per-creator tokens. If no per-handle token exists, the global token is used.

## Online Setup: Streamlit + GitHub Actions

Use this setup for a free hosted dashboard and recurring refreshes:

1. Push this folder to a private GitHub repo.
2. In Streamlit Community Cloud, deploy `app.py` from that repo.
3. In Streamlit secrets, set `APP_PASSWORD` to a long random password.
4. In GitHub repo settings, allow Actions to read and write repository contents.
5. Keep `.github/workflows/weekly-report.yml` enabled.

The workflow refreshes the last complete 7-day `data/latest_snapshot.json` every 2 hours:

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

- YouTube: RSS discovers recent videos when available. If RSS fails, `yt-dlp` reads the public Shorts tab and each discovered Short is checked against the live public page. If `YOUTUBE_API_KEY` is set, the official YouTube Data API is also used for public statistics, especially comments.
- YouTube owner metrics: if `YOUTUBE_ANALYTICS_ACCESS_TOKEN` or `YOUTUBE_ANALYTICS_TOKENS_JSON` is set, the YouTube Analytics API fills shares and net save-to-playlist activity. It also fills `remix_views`, which means views referred from remix links in Shorts. Exact remix count still needs a Studio export if YouTube does not expose it through the authorized API.
- TikTok: TikWM discovers account/post rows and returns views, likes, comments, saves, and shares. `yt-dlp` checks the live public post page as a second source when available.
- Instagram: public feed rows are used when anonymous access is available. oEmbed/page meta verifies owner, caption, date, likes, and comments.
- Instagram owner metrics: if `INSTAGRAM_GRAPH_ACCESS_TOKEN` plus `INSTAGRAM_GRAPH_USER_IDS_JSON` are set, the Meta Graph API fills available saved, shares, reposts, views, likes, and comments counts from owner-authorized media fields.

The tool does not convert missing metrics into zero. A metric is treated as verified only when a public source returned that metric.

For better YouTube comment accuracy, set `YOUTUBE_API_KEY` in GitHub Actions secrets and Streamlit secrets. For owner-only YouTube metrics, use OAuth access tokens, not an API key. The app reads secrets without showing them in the frontend.

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
- YouTube public pages expose views and likes consistently, but comment counts may be unavailable without the official API. Shares, saves, and remixes are YouTube Studio/Analytics metrics, not public page metrics.
- GitHub scheduled workflows are cron-based but can be delayed on shared infrastructure, so avoid treating the exact minute as guaranteed.
