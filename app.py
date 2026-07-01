import hmac
import os
import tempfile
from pathlib import Path

import streamlit as st

from zensi_scraper.core import (
    RunConfig,
    build_exports_from_payload,
    collect_data,
    deserialize_payload,
    load_snapshot,
    parse_date,
    serialize_payload,
    weekly_window,
)


ROOT = Path(__file__).parent
DEFAULT_ROSTER = ROOT / "data" / "creators.zensi.json"
DEFAULT_SNAPSHOT = ROOT / "data" / "latest_snapshot.json"


def get_secret(name):
    try:
        return st.secrets.get(name)
    except Exception:
        return os.environ.get(name)


def require_password():
    expected = get_secret("APP_PASSWORD")
    if not expected:
        st.error("App is locked. Set APP_PASSWORD in Streamlit secrets or the environment.")
        st.stop()
    if st.session_state.get("authenticated"):
        return
    st.title("Zensi UGC Analytics")
    st.caption("Private reporting workspace")
    password = st.text_input("Password", type="password")
    if st.button("Unlock", type="primary"):
        if hmac.compare_digest(password, expected):
            st.session_state["authenticated"] = True
            st.rerun()
        st.error("Wrong password.")
    st.stop()


def platform_counts(posts):
    return {platform: sum(1 for p in posts if p["platform"] == platform) for platform in ["Instagram", "TikTok", "YouTube"]}


def download_exports(payload, key_prefix):
    with tempfile.TemporaryDirectory() as tmp:
        result = build_exports_from_payload(payload, Path(tmp))
        cols = st.columns(3)
        if result.get("docx"):
            path = Path(result["docx"])
            cols[0].download_button("Download Word", path.read_bytes(), file_name=path.name, key=f"{key_prefix}_docx")
        if result.get("xlsx"):
            path = Path(result["xlsx"])
            cols[1].download_button("Download Excel", path.read_bytes(), file_name=path.name, key=f"{key_prefix}_xlsx")
        csv_bytes = []
        for csv_path in result.get("csvs", []):
            path = Path(csv_path)
            csv_bytes.append((path.name, path.read_bytes()))
        if csv_bytes:
            name, data = csv_bytes[0]
            cols[2].download_button("Download posts CSV", data, file_name=name, key=f"{key_prefix}_csv_0")
            for idx, (name, data) in enumerate(csv_bytes[1:], start=1):
                st.download_button(f"Download {name}", data, file_name=name, key=f"{key_prefix}_csv_{idx}")


st.set_page_config(page_title="Zensi UGC Analytics", layout="wide")
require_password()
st.title("Zensi UGC Analytics")
st.caption("Direct public-account scraper for Instagram, TikTok, and YouTube. CSV imports are optional supplements, never the source of truth.")

snapshot = load_snapshot(DEFAULT_SNAPSHOT)
if snapshot:
    restored = deserialize_payload(snapshot)
    posts = restored["posts"]
    counts = platform_counts(posts)
    st.subheader("Latest scheduled scrape")
    st.caption(f"{snapshot.get('start_date')} to {snapshot.get('end_date')} | Updated {snapshot.get('updated_at', 'N/A')}")
    metric_cols = st.columns(5)
    metric_cols[0].metric("Posts", len(posts))
    metric_cols[1].metric("Instagram", counts["Instagram"])
    metric_cols[2].metric("TikTok", counts["TikTok"])
    metric_cols[3].metric("YouTube", counts["YouTube"])
    metric_cols[4].metric("Creators", len({p["creator"] for p in posts}))

    st.bar_chart(counts)
    top_posts = sorted(posts, key=lambda p: p.get("views") or 0, reverse=True)[:40]
    st.dataframe(
        [
            {
                "Creator": p["creator"],
                "Platform": p["platform"],
                "Date": p["date"].isoformat() if p.get("date") else "",
                "Views": p.get("views"),
                "Likes": p.get("likes"),
                "Comments": p.get("comments"),
                "Saves": p.get("saves"),
                "Shares": p.get("shares"),
                "Post": p.get("url"),
            }
            for p in top_posts
        ],
        use_container_width=True,
    )
    download_exports(snapshot, "latest")
else:
    st.warning("No scheduled snapshot found yet. Run the GitHub Action or use the manual refresh below.")

st.divider()
st.subheader("Manual refresh")
default_start, default_end = weekly_window()

with st.sidebar:
    campaign_name = st.text_input("Campaign name", "Zensi")
    start_date = st.date_input("Start date", value=default_start)
    end_date = st.date_input("End date", value=default_end)
    verify_instagram = st.checkbox("Verify Instagram oEmbed/page meta", value=True)
    google_sheet_csv_url = st.text_input("Optional Google Sheet CSV URL for roster import", "")

roster_file = st.file_uploader("Optional creator roster JSON override", type=["json"])
creators_csv = st.file_uploader("Optional creator roster CSV import", type=["csv"])
supplemental_csvs = st.file_uploader("Optional analytics CSV supplements", type=["csv"], accept_multiple_files=True)

st.info("By default this uses data/creators.zensi.json. Uploads only override or supplement the roster for this run.")

if st.button("Refresh from public accounts", type="primary"):
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        roster_path = DEFAULT_ROSTER if DEFAULT_ROSTER.exists() else None
        creators_path = None
        supplemental_paths = []

        if roster_file:
            roster_path = work / "creators.json"
            roster_path.write_bytes(roster_file.read())
        if creators_csv:
            creators_path = work / "creators.csv"
            creators_path.write_bytes(creators_csv.read())
        for idx, upload in enumerate(supplemental_csvs or [], start=1):
            path = work / f"supplemental_{idx}.csv"
            path.write_bytes(upload.read())
            supplemental_paths.append(path)

        config = RunConfig(
            campaign_name=campaign_name,
            start_date=parse_date(str(start_date)),
            end_date=parse_date(str(end_date)),
            output_dir=work / "outputs",
            roster_json=roster_path,
            creators_csv=creators_path,
            supplemental_analytics_csvs=supplemental_paths,
            google_sheet_csv_url=google_sheet_csv_url or None,
            scrape_public=True,
            verify_instagram=verify_instagram,
            export_csv=True,
            export_xlsx=True,
            export_docx=True,
        )
        with st.spinner("Scraping public accounts and preparing exports..."):
            data = collect_data(config)
            payload = serialize_payload(data, config)

        counts = platform_counts(data["posts"])
        st.success(f"Refreshed {len(data['posts'])} posts: {counts}")
        download_exports(payload, "manual")
