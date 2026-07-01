import tempfile
import unittest
from datetime import date
from pathlib import Path

from zensi_scraper.core import (
    RunConfig,
    build_exports_from_payload,
    load_creator_registry,
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
