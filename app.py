import hmac
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from zensi_scraper.core import (
    RunConfig,
    build_exports_from_payload,
    build_integrity_report,
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


def hydrate_optional_env_secrets():
    for name in [
        "YOUTUBE_API_KEY",
        "YOUTUBE_ANALYTICS_ACCESS_TOKEN",
        "YOUTUBE_ANALYTICS_TOKENS_JSON",
        "INSTAGRAM_GRAPH_ACCESS_TOKEN",
        "INSTAGRAM_GRAPH_TOKENS_JSON",
        "INSTAGRAM_GRAPH_USER_IDS_JSON",
        "INSTAGRAM_GRAPH_VERSION",
    ]:
        value = get_secret(name)
        if value and not os.environ.get(name):
            os.environ[name] = str(value)


def platform_counts(posts):
    return {platform: sum(1 for p in posts if p["platform"] == platform) for platform in ["Instagram", "TikTok", "YouTube"]}


def post_rows(posts):
    rows = []
    for p in posts:
        rows.append(
            {
                "Creator": p.get("creator"),
                "Platform": p.get("platform"),
                "Date": p["date"].isoformat() if p.get("date") else "",
                "Views": p.get("views"),
                "Likes": p.get("likes"),
                "Comments": p.get("comments"),
                "Saves": p.get("saves"),
                "Shares": p.get("shares"),
                "Reposts": p.get("reposts"),
                "Remixes": p.get("remixes"),
                "Remix Views": p.get("remix_views"),
                "Category": p.get("category"),
                "Verification": p.get("verification_status", "unverified"),
                "Verified Metrics": p.get("verified_metrics", ""),
                "Unavailable Metrics": p.get("unavailable_metrics", ""),
                "Metric Sources": p.get("metric_sources", ""),
                "Verified At": p.get("verified_at", ""),
                "Integrity Notes": p.get("integrity_notes", ""),
                "Post": p.get("url"),
            }
        )
    return pd.DataFrame(rows)


def apply_filters(df):
    with st.sidebar:
        st.header("Filters")
        creators = st.multiselect("Creators", sorted(df["Creator"].dropna().unique()))
        platforms = st.multiselect("Platforms", sorted(df["Platform"].dropna().unique()))
        statuses = st.multiselect("Verification", sorted(df["Verification"].dropna().unique()))
        parsed_dates = pd.to_datetime(df["Date"], errors="coerce").dropna().dt.date
        default_range = (parsed_dates.min(), parsed_dates.max()) if not parsed_dates.empty else []
        date_range = st.date_input("Date range", value=default_range)
        min_views = st.number_input("Minimum views", min_value=0, value=0, step=100)
        search = st.text_input("Search caption/link", "")
    filtered = df.copy()
    if creators:
        filtered = filtered[filtered["Creator"].isin(creators)]
    if platforms:
        filtered = filtered[filtered["Platform"].isin(platforms)]
    if statuses:
        filtered = filtered[filtered["Verification"].isin(statuses)]
    if len(date_range) == 2:
        start, end = date_range
        dates = pd.to_datetime(filtered["Date"], errors="coerce").dt.date
        filtered = filtered[(dates >= start) & (dates <= end)]
    if min_views:
        filtered = filtered[(pd.to_numeric(filtered["Views"], errors="coerce").fillna(0) >= min_views)]
    if search:
        haystack = (
            filtered["Category"].fillna("")
            + " "
            + filtered["Post"].fillna("")
            + " "
            + filtered["Unavailable Metrics"].fillna("")
            + " "
            + filtered["Integrity Notes"].fillna("")
        ).str.lower()
        filtered = filtered[haystack.str.contains(search.lower(), regex=False)]
    return filtered


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


def render_dashboard(payload, key_prefix):
    restored = deserialize_payload(payload)
    posts = restored["posts"]
    df = post_rows(posts)
    counts = platform_counts(posts)
    integrity = payload.get("integrity") or build_integrity_report(posts)

    st.subheader("Latest Verified Snapshot")
    st.caption(f"{payload.get('start_date')} to {payload.get('end_date')} | Updated {payload.get('updated_at', 'N/A')}")

    cols = st.columns(6)
    cols[0].metric("Posts", len(posts))
    cols[1].metric("Instagram", counts["Instagram"])
    cols[2].metric("TikTok", counts["TikTok"])
    cols[3].metric("YouTube", counts["YouTube"])
    cols[4].metric("Creators", len({p["creator"] for p in posts}))
    cols[5].metric("Verified", integrity.get("statuses", {}).get("verified", 0))

    filtered = apply_filters(df)
    metric_df = filtered.copy()
    for col in ["Views", "Likes", "Comments", "Saves", "Shares", "Reposts", "Remixes", "Remix Views"]:
        metric_df[col] = pd.to_numeric(metric_df[col], errors="coerce").fillna(0)

    chart_cols = st.columns(2)
    if not metric_df.empty:
        platform_chart = metric_df.groupby("Platform", as_index=True)[["Views", "Likes"]].sum()
        creator_chart = metric_df.groupby("Creator", as_index=True)["Views"].sum().sort_values(ascending=False).head(12)
        chart_cols[0].bar_chart(platform_chart)
        chart_cols[1].bar_chart(creator_chart)

    status_cols = st.columns(3)
    statuses = integrity.get("statuses", {})
    status_cols[0].metric("Verified rows", statuses.get("verified", 0))
    status_cols[1].metric("Partial rows", statuses.get("partial", 0))
    status_cols[2].metric("Unverified rows", statuses.get("unverified", 0))

    with st.expander("Data Verification And Integrity System", expanded=False):
        st.markdown(
            """
            - YouTube and TikTok posts are discovered from public account feeds, then re-checked with `yt-dlp` against the live public post page.
            - TikTok rows keep the TikWM public API metrics for saves/shares and use `yt-dlp` as an independent public-page check for views/likes/comments.
            - YouTube rows use RSS for discovery, `yt-dlp` for live public page metrics, and the official YouTube Data API for comments when `YOUTUBE_API_KEY` is configured.
            - YouTube Analytics OAuth can verify shares, comments, and save-to-playlist activity. It can also report `remix_views`, meaning views referred from remix links in Shorts.
            - Exact YouTube remix-count and some Studio-only metrics stay `N/A` unless a creator export/API source supplies them.
            - Instagram rows use public feed rows when anonymous access is available, then oEmbed/page meta for owner, caption, date, likes, and comments.
            - Instagram saves, shares/reposts, and other insight-only metrics stay `N/A` unless a creator insights/API export supplies them.
            - A metric is counted as verified only when a public source returned that metric. Missing extraction is shown as partial/unverified, never silently converted to zero.
            """
        )

    display_df = filtered.copy()
    display_df["_ViewsSort"] = pd.to_numeric(display_df["Views"], errors="coerce").fillna(-1)
    display_df["_DateSort"] = pd.to_datetime(display_df["Date"], errors="coerce")
    display_df = display_df.sort_values(["_ViewsSort", "_DateSort"], ascending=[False, False]).drop(columns=["_ViewsSort", "_DateSort"])
    st.dataframe(
        display_df,
        use_container_width=True,
        column_config={"Post": st.column_config.LinkColumn("Post")},
        hide_index=True,
    )
    download_exports(payload, key_prefix)


st.set_page_config(page_title="Zensi UGC Analytics", layout="wide")
require_password()
hydrate_optional_env_secrets()
st.title("Zensi UGC Analytics")
st.caption("Direct public-account analytics with row-level source verification.")

snapshot = load_snapshot(DEFAULT_SNAPSHOT)
if snapshot:
    render_dashboard(snapshot, "latest")
else:
    st.warning("No scheduled snapshot found yet. Run the GitHub Action or use the manual refresh below.")

st.divider()
st.subheader("Manual Refresh")
default_start, default_end = weekly_window()

with st.sidebar:
    st.header("Manual Run")
    campaign_name = st.text_input("Campaign name", "Zensi")
    start_date = st.date_input("Start date", value=default_start, key="manual_start")
    end_date = st.date_input("End date", value=default_end, key="manual_end")
    verify_metrics = st.checkbox("Verify public metrics", value=True)
    google_sheet_csv_url = st.text_input("Optional Google Sheet CSV URL for roster import", "")

roster_file = st.file_uploader("Optional creator roster JSON override", type=["json"])
creators_csv = st.file_uploader("Optional creator roster CSV import", type=["csv"])
supplemental_csvs = st.file_uploader("Optional analytics CSV supplements", type=["csv"], accept_multiple_files=True)

if st.button("Refresh From Public Accounts", type="primary"):
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
            verify_instagram=verify_metrics,
            export_csv=True,
            export_xlsx=True,
            export_docx=True,
        )
        with st.spinner("Scraping public accounts and verifying metrics..."):
            data = collect_data(config)
            payload = serialize_payload(data, config)

        st.success(f"Refreshed {len(data['posts'])} posts.")
        render_dashboard(payload, "manual")
