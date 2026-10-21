#!/usr/bin/env python3
"""
User-wise Netskope events for a fixed list of apps / websites.

Put the apps and websites you care about in targets.txt (next to this
script), one per line -- app names exactly as Netskope shows them
(e.g. ChatGPT) or domains / hostnames (e.g. chat.openai.com). Lines
starting with # are ignored. The list never leaves this machine except
as the API filter.

Each entry is matched against app, site, domain and hostname, on both
    /api/v2/events/datasearch/application   activities, objects, files
    /api/v2/events/datasearch/page          plain website visits
and every matching event is written, one row per event, sorted by user:
    output/UserWise_<from>_to_<to>.csv

The HTTP, retry, paging and time-window logic is imported from
../Netskope-API/export_genai_page_bytes.py, same as bulk_upload_uba.py.

Usage:
    python export_userwise.py 2026-08-01 2026-09-30   -> both days inclusive
    python export_userwise.py --selftest              -> offline check
"""

import os
import sys
import time
from datetime import datetime, timedelta
from typing import Any, Dict, List, Tuple

import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "Netskope-API"))
import export_genai_page_bytes as ns  # noqa: E402

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
# SECURITY: the token lives in this file -- keep it out of git / shares.
API_TOKEN: str = "PUT_YOUR_API_TOKEN_HERE"
BASE_URL: str = "https://YOUR_TENANT.goskope.com"
AUTH_MODE: str = "bearer"                 # flip to "netskope" on 401s

TARGETS_FILE: str = os.path.join(SCRIPT_DIR, "targets.txt")

# Event type label -> endpoint. Drop "Page" if you only want app activity.
ENDPOINTS: Dict[str, str] = {
    "Application": "/api/v2/events/datasearch/application",
    "Page": "/api/v2/events/datasearch/page",
}

# Every target is tried against each of these fields (exact match).
MATCH_FIELDS: List[str] = ["app", "site", "domain", "hostname"]

# Targets per query, keeps the URL a sane length.
TARGET_BATCH: int = 50

OUTPUT_DIR: str = os.path.join(SCRIPT_DIR, "output")
LOG_DIR: str = os.path.join(SCRIPT_DIR, "logs")

# API field -> CSV header, in output order. Missing fields export blank.
COLUMNS: Dict[str, str] = {
    "user": "User",
    "event_type": "Event Type",
    "app": "Application",
    "site": "Site",
    "domain": "Domain",
    "hostname": "Hostname",
    "url": "URL",
    "activity": "Activity",
    "object_id": "Object ID",
    "object": "Object Name",
    "object_type": "Object Type",
    "file_path": "File Path",
    "file_size": "File Size (Bytes)",
    "timestamp": "Event Date",
    "organization_unit": "Organization Unit",
}

# ---------------------------------------------------------------------------


def read_targets(path: str) -> List[str]:
    """Non-blank, non-# lines, de-duplicated, order kept."""
    with open(path, encoding="utf-8-sig") as fh:
        lines = (line.strip() for line in fh)
        return list(dict.fromkeys(l for l in lines if l and not l.startswith("#")))


def build_query(targets: List[str]) -> str:
    """app in [...] or site in [...] or ... -- quoted here, not via
    ns.build_term, so names containing commas survive."""
    values = "[" + ", ".join(ns.quote_value(t) for t in targets) + "]"
    return " or ".join(f"{field} in {values}" for field in MATCH_FIELDS)


def parse_range(first: str, last: str) -> Tuple[int, int, str]:
    """Local midnight of `first` .. 23:59:59 of `last` (endtime is inclusive)."""
    start = datetime.strptime(first, "%Y-%m-%d")
    end = datetime.strptime(last, "%Y-%m-%d") + timedelta(days=1)
    if end <= start:
        raise ValueError("end date is before start date")
    return int(start.timestamp()), int(end.timestamp()) - 1, f"{first}_to_{last}"


def to_frame(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    frame = pd.DataFrame(rows).reindex(columns=list(COLUMNS))
    frame = frame.sort_values(["user", "timestamp"], kind="stable")
    frame["timestamp"] = frame["timestamp"].map(ns.format_epoch)
    return frame.rename(columns=COLUMNS).fillna("")


def main() -> int:
    for name in ("API_TOKEN", "BASE_URL", "AUTH_MODE"):
        setattr(ns, name, globals()[name])
    ns.LOG_DIR, ns.LOG_BASENAME = LOG_DIR, "export_userwise.log"
    ns.SEND_FIELDS_PARAM = False      # ns's field list is for its own export
    ns.setup_logging(time.strftime("%H%M%S"))

    if "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL:
        ns.log.error("CONFIG ERROR: set API_TOKEN / BASE_URL at the top of %s",
                     os.path.basename(__file__))
        return 1
    try:
        start, end, label = parse_range(sys.argv[1], sys.argv[2])
        targets = read_targets(TARGETS_FILE)
    except (IndexError, ValueError) as exc:
        ns.log.error("BAD DATES: %s (use: export_userwise.py 2026-08-01 2026-09-30)", exc)
        return 1
    except OSError as exc:
        ns.log.error("TARGETS FILE: %s", exc)
        return 1
    if not targets:
        ns.log.error("targets.txt is empty")
        return 1

    session = ns.build_session()
    stats: Dict[str, int] = {"calls": 0}
    rows: List[Dict[str, Any]] = []
    ns.log.info("range      %s", label)
    ns.log.info("targets    %d", len(targets))
    try:
        for event_type, endpoint in ENDPOINTS.items():
            ns.ENDPOINT_PATH = endpoint
            for i in range(0, len(targets), TARGET_BATCH):
                query = build_query(targets[i:i + TARGET_BATCH])
                for record in ns.fetch_all(session, query, start, end, stats):
                    rows.append({**ns.flatten_record(record), "event_type": event_type})
            ns.log.info("%-10s %d events so far", event_type, len(rows))
    except ns.ApiError as exc:
        ns.log.error("NETSKOPE FAILED: %s", exc)
        return 2

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, f"UserWise_{label}.csv")
    frame = to_frame(rows)
    frame.to_csv(path, index=False, encoding="utf-8")
    ns.log.info("wrote      %s (%d rows, %d users)", path, len(frame),
                frame["User"].nunique())
    ns.log.info("API calls  %d", stats["calls"])
    return 0 if rows else 3


def selftest() -> None:
    assert build_query(["ChatGPT", "a, b"]) == (
        "app in ['ChatGPT', 'a, b'] or site in ['ChatGPT', 'a, b'] or "
        "domain in ['ChatGPT', 'a, b'] or hostname in ['ChatGPT', 'a, b']")

    start, end, label = parse_range("2026-08-01", "2026-09-30")
    assert label == "2026-08-01_to_2026-09-30"
    assert time.localtime(start)[:6] == (2026, 8, 1, 0, 0, 0)
    assert time.localtime(end)[:6] == (2026, 9, 30, 23, 59, 59)
    try:
        parse_range("2026-09-30", "2026-08-01")
        raise AssertionError("reversed range accepted")
    except ValueError:
        pass

    frame = to_frame([
        {"user": "b@x.com", "app": "ChatGPT", "activity": "Upload",
         "object": "a.pdf", "object_id": "1", "file_size": 10,
         "timestamp": 1754000000, "event_type": "Application"},
        {"user": "a@x.com", "domain": "x.com", "timestamp": 1754000001,
         "event_type": "Page"},
    ])
    assert list(frame.columns) == list(COLUMNS.values())
    assert list(frame["User"]) == ["a@x.com", "b@x.com"]
    assert frame.iloc[1]["Object ID"] == "1" and frame.iloc[0]["Object ID"] == ""
    assert frame.iloc[0]["Event Date"]
    print("selftest OK")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        sys.exit(0)
    sys.exit(main())
