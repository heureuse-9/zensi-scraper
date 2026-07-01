# Security Notes

This repo is intended to stay private.

## Required For Hosted Use

- Deploy from a private GitHub repo.
- Set `APP_PASSWORD` in Streamlit Community Cloud secrets.
- Do not commit `.streamlit/secrets.toml`, `.env`, `config.json`, report exports, logs, or browser caches.
- Keep GitHub Actions permissions scoped to repository contents only.

## Data Handling

- The scraper does not use creator logins.
- The app reads public account/post data and the local roster in `data/creators.zensi.json`.
- `data/latest_snapshot.json` contains reportable creator analytics and should be treated as private campaign data.
- Optional CSV imports are supplements only; they should not contain credentials or account access tokens.

## Before Making The Repo Public

Delete `data/latest_snapshot.json` and any campaign-specific creator roster/data first. Public repos should contain tool code only.
