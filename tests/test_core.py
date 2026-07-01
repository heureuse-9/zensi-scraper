import tempfile
import unittest
from datetime import date
from pathlib import Path

from zensi_scraper.core import (
    RunConfig,
    apply_metric_record,
    build_exports_from_payload,
    build_integrity_report,
    load_creator_registry,
    metrics_from_ytdlp_info,
    serialize_payload,
    split_handles,
    weekly_window,
    write_snapshot,
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
