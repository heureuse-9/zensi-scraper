import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import Mock, patch

from zensi_scraper.core import (
    RunConfig,
    apply_metric_record,
    build_exports_from_payload,
    build_integrity_report,
    fetch_youtube_api_video_stats,
    load_creator_registry,
    load_csv_posts,
    mark_public_unavailable_metrics,
    metrics_from_ytdlp_info,
    owner_metrics_csv_template,
    records_from_instagram_owner_media,
    serialize_payload,
    split_handles,
    weekly_window,
    write_snapshot,
    youtube_analytics_records_from_rows,
    youtube_flat_entries_from_info,
)


ROOT = Path(__file__).resolve().parents[1]


class CoreTests(unittest.TestCase):
    def test_split_handles_accepts_strings_lists_and_urls(self):
        self.assertEqual(split_handles("@alpha, https://www.tiktok.com/@beta/gamma"), ["alpha", "beta/gamma"])
        self.assertEqual(split_handles(["@alpha", "alpha", "youtube.com/@Beta/"]), ["alpha", "Beta"])

    def test_weekly_window_uses_last_complete_days(self):
        start, end = weekly_window(date(2026, 7, 3), 7)
        self.assertEqual(start, date(2026, 6, 26))
        self.assertEqual(end, date(2026, 7, 2))

    def test_seed_roster_loads_known_creators(self):
        roster = load_creator_registry(ROOT / "data" / "creators.zensi.json")
        self.assertGreaterEqual(len(roster), 15)
        self.assertIn("Bryson haywood", roster)
        self.assertIn("lockinwithdrain", roster["Bryson haywood"]["tt"])

    def test_ytdlp_metrics_are_normalized_without_fake_zeroes(self):
        metrics = metrics_from_ytdlp_info(
            "YouTube",
            {
                "id": "tLcsrdjJJG4",
                "title": "Well ok so my phone just betrayed me",
                "view_count": 2279,
                "like_count": 27,
                "comment_count": None,
                "upload_date": "20260624",
                "webpage_url": "https://www.youtube.com/watch?v=tLcsrdjJJG4",
            },
        )
        self.assertEqual(metrics["views"], 2279)
        self.assertEqual(metrics["likes"], 27)
        self.assertIsNone(metrics["comments"])
        self.assertEqual(metrics["post_id"], "tLcsrdjJJG4")

    def test_metric_integrity_keeps_missing_metrics_unverified(self):
        post = {
            "creator": "Creator One",
            "platform": "YouTube",
            "handle": "creatorone",
            "date": date(2026, 6, 24),
            "post_type": "Short",
            "post_id": "tLcsrdjJJG4",
            "title": "Old title",
            "views": 0,
            "likes": 0,
            "comments": None,
            "saves": None,
            "shares": None,
            "url": "https://www.youtube.com/shorts/tLcsrdjJJG4",
            "_source": "youtube_rss_public_page",
        }
        apply_metric_record(
            post,
            {
                "views": 2279,
                "likes": 27,
                "comments": None,
                "title": "Well ok so my phone just betrayed me",
                "source": "yt_dlp_public_page",
            },
        )
        self.assertEqual(post["views"], 2279)
        self.assertEqual(post["likes"], 27)
        self.assertIsNone(post["comments"])
        self.assertEqual(post["verification_status"], "partial")
        self.assertIn("comments", post["integrity_notes"])
        self.assertIn("yt_dlp_public_page", post["metric_sources"])

    def test_youtube_data_api_stats_verify_comments_without_fake_private_metrics(self):
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {
            "items": [
                {
                    "id": "tLcsrdjJJG4",
                    "snippet": {"title": "Well ok so my phone just betrayed me", "publishedAt": "2026-06-24T12:00:00Z"},
                    "statistics": {"viewCount": "2200", "likeCount": "27", "commentCount": "0", "favoriteCount": "0"},
                }
            ]
        }
        with patch("zensi_scraper.core.requests.get", return_value=response) as get:
            stats = fetch_youtube_api_video_stats("tLcsrdjJJG4", "api-key")

        self.assertEqual(stats["views"], 2200)
        self.assertEqual(stats["likes"], 27)
        self.assertEqual(stats["comments"], 0)
        self.assertNotIn("shares", stats)
        self.assertNotIn("saves", stats)
        self.assertNotIn("remixes", stats)
        get.assert_called_once()

    def test_youtube_flat_entries_from_info_normalizes_shorts_tab_results(self):
        entries = youtube_flat_entries_from_info(
            {
                "entries": [
                    {"id": "abc123", "title": "First short", "url": "https://www.youtube.com/shorts/abc123", "view_count": 1700},
                    {"title": "Missing id", "url": "https://www.youtube.com/shorts/def456"},
                    {"id": "", "url": ""},
                ]
            }
        )
        self.assertEqual(len(entries), 2)
        self.assertEqual(entries[0]["post_id"], "abc123")
        self.assertEqual(entries[1]["post_id"], "def456")
        self.assertEqual(entries[0]["views"], 1700)

    def test_instagram_owner_media_records_include_private_counts(self):
        records = records_from_instagram_owner_media(
            {
                "data": [
                    {
                        "id": "1789",
                        "permalink": "https://www.instagram.com/reel/abc123/",
                        "caption": "Study reset",
                        "timestamp": "2026-07-06T12:00:00+0000",
                        "total_views_count": 1000,
                        "total_like_count": 100,
                        "total_comments_count": 10,
                        "saved_count": 25,
                        "shares_count": 4,
                        "reposts_count": 2,
                    }
                ]
            }
        )
        self.assertEqual(records["abc123"]["views"], 1000)
        self.assertEqual(records["abc123"]["saves"], 25)
        self.assertEqual(records["abc123"]["shares"], 4)
        self.assertEqual(records["abc123"]["reposts"], 2)
        self.assertIn("instagram_graph_owner_media", records["abc123"]["source"])

    def test_youtube_owner_analytics_records_include_shares_saves_and_remix_views(self):
        records = youtube_analytics_records_from_rows(
            ["video", "views", "likes", "comments", "shares", "videosAddedToPlaylists", "videosRemovedFromPlaylists"],
            [["tLcsrdjJJG4", 2200, 27, 0, 9, 11, 3]],
        )
        remix_records = youtube_analytics_records_from_rows(
            ["video", "insightTrafficSourceType", "views"],
            [["tLcsrdjJJG4", "VIDEO_REMIXES", 42]],
        )
        self.assertEqual(records["tLcsrdjJJG4"]["shares"], 9)
        self.assertEqual(records["tLcsrdjJJG4"]["saves"], 8)
        self.assertEqual(records["tLcsrdjJJG4"]["playlist_adds"], 11)
        self.assertEqual(records["tLcsrdjJJG4"]["playlist_removes"], 3)
        self.assertEqual(remix_records["tLcsrdjJJG4"]["remix_views"], 42)

    def test_private_platform_metrics_are_marked_unavailable_not_zero(self):
        post = {
            "creator": "Creator One",
            "platform": "Instagram",
            "handle": "creatorone",
            "date": date(2026, 7, 6),
            "post_type": "Reel",
            "post_id": "abc123",
            "title": "Study reset",
            "views": 1000,
            "likes": 100,
            "comments": 10,
            "saves": None,
            "shares": None,
            "reposts": None,
            "remixes": None,
            "url": "https://www.instagram.com/reel/abc123/",
        }
        mark_public_unavailable_metrics(post)
        self.assertIsNone(post["saves"])
        self.assertIsNone(post["shares"])
        self.assertIsNone(post["reposts"])
        self.assertIn("saves", post["unavailable_metrics"])
        self.assertIn("creator insights/API export", post["integrity_notes"])

    def test_creator_export_csv_can_supply_ig_and_youtube_private_metrics(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "analytics.csv"
            path.write_text(
                "\n".join(
                    [
                        "creator,platform,date,url,views,likes,comments,saves,shares,reposts,remixes",
                        "Creator One,Instagram,2026-07-06,https://www.instagram.com/reel/abc123/,1000,100,10,25,4,2,",
                        "Creator Two,YouTube,2026-07-06,https://www.youtube.com/shorts/tLcsrdjJJG4,2200,27,0,8,3,,1",
                    ]
                ),
                encoding="utf-8",
            )
            posts = load_csv_posts(path, date(2026, 7, 1), date(2026, 7, 7))

        ig = next(p for p in posts if p["platform"] == "Instagram")
        yt = next(p for p in posts if p["platform"] == "YouTube")
        self.assertEqual(ig["saves"], 25)
        self.assertEqual(ig["shares"], 4)
        self.assertEqual(ig["reposts"], 2)
        self.assertEqual(yt["comments"], 0)
        self.assertEqual(yt["saves"], 8)
        self.assertEqual(yt["shares"], 3)
        self.assertEqual(yt["remixes"], 1)
        self.assertEqual(ig["verification_status"], "verified")
        self.assertIn("creator_export_csv", ig["metric_sources"])

    def test_owner_metrics_template_has_no_login_private_metric_columns(self):
        template = owner_metrics_csv_template()
        header = template.splitlines()[0].split(",")
        for column in ["creator", "platform", "date", "url", "saves", "shares", "reposts", "remixes", "remix_views"]:
            self.assertIn(column, header)
        self.assertIn("Instagram", template)
        self.assertIn("YouTube", template)

    def test_integrity_report_counts_statuses(self):
        report = build_integrity_report(
            [
                {"platform": "YouTube", "verification_status": "partial"},
                {"platform": "TikTok", "verification_status": "verified"},
                {"platform": "Instagram", "verification_status": "unverified"},
            ]
        )
        self.assertEqual(report["total_posts"], 3)
        self.assertEqual(report["statuses"]["partial"], 1)
        self.assertEqual(report["platforms"]["TikTok"]["verified"], 1)

    def test_snapshot_can_rebuild_exports(self):
        config = RunConfig(
            campaign_name="Zensi",
            start_date=date(2026, 6, 26),
            end_date=date(2026, 7, 2),
            output_dir=ROOT / "runs" / "test",
            export_csv=True,
            export_xlsx=True,
            export_docx=True,
        )
        data = {
            "roster": {"Creator One": {"ig": ["creatorone"], "tt": [], "yt": []}},
            "profiles_tt": [],
            "profiles_ig": [],
            "profiles_yt": [],
            "posts": [
                {
                    "creator": "Creator One",
                    "platform": "Instagram",
                    "handle": "creatorone",
                    "date": date(2026, 6, 30),
                    "post_type": "Reel",
                    "post_id": "abc123",
                    "title": "Study reset",
                    "views": 1000,
                    "likes": 100,
                    "comments": 10,
                    "saves": None,
                    "shares": None,
                    "category": "Study / academics",
                    "url": "https://www.instagram.com/reel/abc123/",
                    "_source": "public_scrape",
                }
            ],
            "summaries": [
                {
                    "creator": "Creator One",
                    "handles": "IG @creatorone",
                    "tt_followers": None,
                    "tt_total_likes": None,
                    "tt_videos": 0,
                    "ig_followers": None,
                    "ig_posts": 1,
                    "ig_agg_likes": 100,
                    "ig_agg_comments": 10,
                    "ig_agg_views": 1000,
                    "ig_avg_likes": 100,
                    "ig_eng_rate": None,
                    "yt_subscribers": None,
                    "yt_total_views": 0,
                    "yt_total_likes": 0,
                    "yt_videos": 0,
                    "yt_avg_views": None,
                    "yt_avg_likes": None,
                    "yt_eng_rate": None,
                    "total_audience": 0,
                    "top_platform": "Instagram",
                }
            ],
        }
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            snapshot = write_snapshot(tmp_path / "latest_snapshot.json", data, config)
            payload = serialize_payload(data, config)
            self.assertTrue(snapshot.exists())
            result = build_exports_from_payload(payload, tmp_path)
            self.assertTrue(Path(result["docx"]).exists())
            self.assertTrue(Path(result["xlsx"]).exists())
            self.assertEqual(len(result["csvs"]), 3)


if __name__ == "__main__":
    unittest.main()
