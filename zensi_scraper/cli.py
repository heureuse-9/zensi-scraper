import argparse
import json
from pathlib import Path

from .core import (
    RunConfig,
    collect_data,
    config_from_json,
    load_creator_registry,
    parse_date,
    run_report,
    save_roster_json,
    write_snapshot,
    split_handles,
    weekly_window,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build recurring Zensi creator analytics reports.")
    sub = parser.add_subparsers(dest="command")

    run = sub.add_parser("run", help="Run from a config JSON file or direct CLI args.")
    run.add_argument("--config", help="Path to config.json.")
    run.add_argument("--campaign-name", default="Zensi")
    run.add_argument("--start-date")
    run.add_argument("--end-date")
    run.add_argument("--output-dir")
    run.add_argument("--roster-json")
    run.add_argument("--creators-csv", help="Optional handle roster import only, not analytics.")
    run.add_argument("--supplemental-analytics-csv", action="append", default=[], help="Optional analytics CSV supplement. Can be repeated.")
    run.add_argument("--supplemental-docx", help="Optional report DOCX supplement.")
    run.add_argument("--google-sheet-csv-url")
    run.add_argument("--chrome-path")
    run.add_argument("--output-stem")
    run.add_argument("--no-public-scrape", action="store_true")
    run.add_argument("--no-instagram-verify", action="store_true")
    run.add_argument("--no-csv", action="store_true")
    run.add_argument("--no-xlsx", action="store_true")
    run.add_argument("--no-docx", action="store_true")

    weekly = sub.add_parser("run-weekly", help="Run the last complete weekly reporting window.")
    weekly.add_argument("--config", required=True)
    weekly.add_argument("--date", help="Override run date, YYYY-MM-DD. Defaults to today.")
    weekly.add_argument("--lookback-days", type=int, default=7)

    refresh = sub.add_parser("refresh-cache", help="Refresh the latest public-scrape snapshot for the Streamlit app.")
    refresh.add_argument("--config", required=True)
    refresh.add_argument("--snapshot", default="data/latest_snapshot.json")
    refresh.add_argument("--date", help="Override run date, YYYY-MM-DD. Defaults to today.")
    refresh.add_argument("--lookback-days", type=int, default=14)
    refresh.add_argument("--no-instagram-verify", action="store_true")

    init = sub.add_parser("init-config", help="Write a starter config file.")
    init.add_argument("--path", default="config.json")

    roster = sub.add_parser("import-roster", help="Import creator handles into a roster JSON.")
    roster.add_argument("--from-csv", required=True)
    roster.add_argument("--output", required=True)
    roster.add_argument("--google-sheet-csv-url")

    add = sub.add_parser("add-creator", help="Add one creator to a roster JSON.")
    add.add_argument("--roster-json", required=True)
    add.add_argument("--name", required=True)
    add.add_argument("--instagram", default="")
    add.add_argument("--tiktok", default="")
    add.add_argument("--youtube", default="")
    return parser


def config_from_args(args) -> RunConfig:
    if args.config:
        config = config_from_json(args.config)
        if args.no_public_scrape:
            config.scrape_public = False
        if args.no_instagram_verify:
            config.verify_instagram = False
        if args.no_csv:
            config.export_csv = False
        if args.no_xlsx:
            config.export_xlsx = False
        if args.no_docx:
            config.export_docx = False
        return config
    missing = [name for name in ["start_date", "end_date", "output_dir"] if not getattr(args, name)]
    if missing:
        raise SystemExit(f"Missing required args without --config: {', '.join('--' + m.replace('_', '-') for m in missing)}")
    return RunConfig(
        campaign_name=args.campaign_name,
        start_date=parse_date(args.start_date),
        end_date=parse_date(args.end_date),
        output_dir=Path(args.output_dir),
        roster_json=Path(args.roster_json) if args.roster_json else None,
        creators_csv=Path(args.creators_csv) if args.creators_csv else None,
        supplemental_analytics_csvs=[Path(p) for p in args.supplemental_analytics_csv],
        supplemental_docx=Path(args.supplemental_docx) if args.supplemental_docx else None,
        google_sheet_csv_url=args.google_sheet_csv_url,
        scrape_public=not args.no_public_scrape,
        verify_instagram=not args.no_instagram_verify,
        chrome_path=args.chrome_path,
        output_stem=args.output_stem,
        export_csv=not args.no_csv,
        export_xlsx=not args.no_xlsx,
        export_docx=not args.no_docx,
    )


def write_config(path: Path):
    sample = {
        "campaign_name": "Zensi",
        "start_date": "",
        "end_date": "",
        "output_dir": r"C:\Users\Meso\Desktop\Studio Cores\Zensi\Reporting",
        "roster_json": r"C:\Users\Meso\Desktop\Studio Cores\Zensi\Tools\zensi-creator-scraper\data\creators.zensi.json",
        "creators_csv": "",
        "supplemental_analytics_csvs": [],
        "supplemental_docx": "",
        "google_sheet_csv_url": "",
        "scrape_public": True,
        "verify_instagram": True,
        "chrome_path": "",
        "output_stem": "",
        "export_csv": True,
        "export_xlsx": True,
        "export_docx": True,
    }
    path.write_text(json.dumps(sample, indent=2), encoding="utf-8")
    print(f"Wrote {path}")


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "init-config":
        write_config(Path(args.path))
        return
    if args.command == "import-roster":
        roster = load_creator_registry(None, Path(args.from_csv), args.google_sheet_csv_url)
        save_roster_json(Path(args.output), roster)
        print(f"Imported {len(roster)} creator(s) to {args.output}")
        return
    if args.command == "add-creator":
        roster_path = Path(args.roster_json)
        roster = load_creator_registry(roster_path)
        entry = roster.setdefault(args.name, {"ig": [], "tt": [], "yt": []})
        for handle in split_handles(args.instagram):
            if handle.lower() not in [h.lower() for h in entry["ig"]]:
                entry["ig"].append(handle)
        for handle in split_handles(args.tiktok):
            if handle.lower() not in [h.lower() for h in entry["tt"]]:
                entry["tt"].append(handle)
        for handle in split_handles(args.youtube):
            if handle.lower() not in [h.lower() for h in entry["yt"]]:
                entry["yt"].append(handle)
        save_roster_json(roster_path, roster)
        print(f"Saved {args.name} to {roster_path}")
        return
    if args.command == "run-weekly":
        config = config_from_json(args.config, require_dates=False)
        run_date = parse_date(args.date) if args.date else None
        config.start_date, config.end_date = weekly_window(run_date, args.lookback_days)
        result = run_report(config)
        print(f"Done weekly run: {config.start_date} to {config.end_date}")
        print(f"Posts: {result['posts']}")
        print(f"Platform counts: {result['platform_counts']}")
        print(f"Word: {result['docx']}")
        print(f"Excel: {result['xlsx']}")
        print(f"CSVs: {', '.join(str(p) for p in result['csvs'])}")
        return
    if args.command == "refresh-cache":
        config = config_from_json(args.config, require_dates=False)
        run_date = parse_date(args.date) if args.date else None
        config.start_date, config.end_date = weekly_window(run_date, args.lookback_days)
        config.export_csv = False
        config.export_xlsx = False
        config.export_docx = False
        if args.no_instagram_verify:
            config.verify_instagram = False
        data = collect_data(config)
        snapshot = write_snapshot(Path(args.snapshot), data, config)
        counts = {platform: sum(1 for p in data["posts"] if p["platform"] == platform) for platform in ["Instagram", "TikTok", "YouTube"]}
        print(f"Refreshed snapshot: {snapshot}")
        print(f"Window: {config.start_date} to {config.end_date}")
        print(f"Posts: {len(data['posts'])}")
        print(f"Platform counts: {counts}")
        return
    if args.command != "run":
        parser.print_help()
        return
    result = run_report(config_from_args(args))
    print(f"Done: {result['posts']} posts")
    print(f"Platform counts: {result['platform_counts']}")
    print(f"Word: {result['docx']}")
    print(f"Excel: {result['xlsx']}")
    print(f"CSVs: {', '.join(str(p) for p in result['csvs'])}")


if __name__ == "__main__":
    main()
