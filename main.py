"""Import one semester of the IO timetable into Google Calendar.

  python main.py 1                 # dry run for semester 1: report + preview_sem1.csv (no Google access)
  python main.py 1 --push          # create/update semester 1 events in Google Calendar
  python main.py 1 --push --prune  # ...and delete tagged events no longer in the plan
  python main.py 1 --delete-all    # remove every semester 1 event this script created
"""
from __future__ import annotations

import argparse
import csv
import os
import tomllib
from collections import Counter

from parse_plan import GOOGLE_COLORS, NOTICE_KEY, PlanParser

HERE = os.path.dirname(os.path.abspath(__file__))
COLOR_NAMES = {None: "default", "1": "Lavender", "2": "Sage", "3": "Grape", "4": "Flamingo", "5": "Banana",
               "6": "Tangerine", "7": "Peacock", "8": "Graphite", "9": "Blueberry",
               "10": "Basil", "11": "Tomato"}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("semester", type=int, choices=[1, 2, 3], help="semester to import (1, 2 or 3)")
    ap.add_argument("--config", default=os.path.join(HERE, "config.toml"))
    ap.add_argument("--excel", help="override excel_file from config")
    ap.add_argument("--push", action="store_true", help="write events to Google Calendar")
    ap.add_argument("--prune", action="store_true", help="with --push: delete stale tagged events")
    ap.add_argument("--delete-all", action="store_true", help="delete all events created by this script")
    args = ap.parse_args()

    with open(args.config, "rb") as f:
        cfg = tomllib.load(f)
    cfg["semester"] = args.semester
    for key in ("semester_header", "calendar_name", "source_tag"):
        cfg[key] = (cfg.get(key) or "").replace("{semester}", str(args.semester))
    if args.excel:
        cfg["excel_file"] = args.excel
    excel = cfg["excel_file"]
    if not os.path.isabs(excel):
        excel = os.path.join(os.path.dirname(os.path.abspath(args.config)), excel)

    if args.delete_all:
        import gcal
        svc = gcal.get_service()
        cal = gcal.resolve_calendar(svc, cfg, create=False)
        ids = [e["id"] for e in gcal.list_tagged(svc, cal, cfg["source_tag"])]
        if input(f"Delete {len(ids)} event(s)? [y/N] ").lower() == "y":
            gcal.delete_ids(svc, cal, ids)
        return

    parser = PlanParser(cfg, excel)
    events = parser.parse()
    report(events, parser)

    preview = os.path.join(HERE, f"preview_sem{args.semester}.csv")
    with open(preview, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["date", "start", "end", "code", "title", "lecturer/topic", "room", "color", "excel_row"])
        for e in sorted(events, key=lambda e: (e.date, e.start or e.date)):
            w.writerow([e.date, e.start or "all-day", e.end or (e.end_date or ""), e.code, e.title,
                        e.lecturer, e.room, COLOR_NAMES[e.color_id], e.row])
    print(f"\nPreview written to {preview}")

    if not args.push:
        print("Dry run only. Re-run with --push to write to Google Calendar.")
        return

    import gcal
    svc = gcal.get_service()
    cal = gcal.resolve_calendar(svc, cfg)
    gcal.upsert(svc, cal, events, cfg)
    if args.prune:
        keep = {e.event_id for e in events}
        stale = [x["id"] for x in gcal.list_tagged(svc, cal, cfg["source_tag"]) if x["id"] not in keep]
        if stale:
            gcal.delete_ids(svc, cal, stale)


def report(events, parser):
    counts = Counter(e.code for e in events)
    print(f"{len(events)} events parsed for {parser.cfg['semester_header']}.\n")
    print(f"{'code':<10} {'n':>3}  {'colour':<10} title")
    for code, n in counts.items():
        label = "(notices)" if code == NOTICE_KEY else code
        title = "" if code == NOTICE_KEY else next(e.title for e in events if e.code == code)
        print(f"{label:<10} {n:>3}  {COLOR_NAMES[parser.color_map[code]]:<10} {title}")
    if parser.warnings:
        print("\nWARNINGS:")
        for w in dict.fromkeys(parser.warnings):
            print("  -", w)
    if parser.skipped:
        print("\nSKIPPED (date anomaly – fix via [date_fixes] in config.toml):")
        for s in parser.skipped:
            print("  -", s)


if __name__ == "__main__":
    main()
