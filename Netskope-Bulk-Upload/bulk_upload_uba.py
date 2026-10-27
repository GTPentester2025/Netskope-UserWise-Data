#!/usr/bin/env python3
"""
Netskope UBA bulk upload / download -> application events mapping.

  1. Pull UBA alerts                  (/api/v2/events/datasearch/alert,
                                        query: alert_type eq 'uba' and
                                        activity in [Upload, Download,
                                        Download All])
     -- the same filter as Skope IT > Events & Alerts.
  2. Take the distinct user emails from those alerts and pull the
     application events for just those users, same activities
                                       (/api/v2/events/datasearch/application).
  3. One row per UBA alert, with that user's events summarised onto it.

Outputs (in ./output):
    UBA_BulkTransfer_<day>.csv     every matching UBA alert, all fields
    UBA_AppEvents_<day>.csv        application events for those users
    UBA_Mapped_<day>.csv           UBA alert columns + app event columns

The HTTP, retry, paging and time-window logic is NOT duplicated here: it is
imported from ../Netskope-API/export_genai_page_bytes.py, so a fix made there
applies to this script too. Keep both folders side by side.

Usage:
    python bulk_upload_uba.py               -> yesterday
    python bulk_upload_uba.py 2026-10-18    -> that day
    python bulk_upload_uba.py --days 7      -> past 7 full days (to yesterday)
    python bulk_upload_uba.py --selftest    -> offline check of filter + join
"""

import os
import sys
import time
from typing import Any, Dict, List

import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "Netskope-API"))
import export_genai_page_bytes as ns  # noqa: E402

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
# SECURITY: the token lives in this file -- keep it out of git (chmod 600).
API_TOKEN: str = "PUT_YOUR_API_TOKEN_HERE"
BASE_URL: str = "https://YOUR_TENANT.goskope.com"
AUTH_MODE: str = "bearer"                 # flip to "netskope" on 401s

ALERT_ENDPOINT: str = "/api/v2/events/datasearch/alert"
APP_ENDPOINT: str = "/api/v2/events/datasearch/application"

# Same filter as Skope IT > Events & Alerts: Alert Type = UBA and Activity in
# these. Applied server-side to BOTH the UBA alerts and the application
# events, so Browse, Login, mail etc. never come back.
ACTIVITIES: List[str] = ["Upload", "Download", "Download All"]
UBA_QUERY: str = ("alert_type eq 'uba' and "
                  + ns.build_term("activity", "in", ",".join(ACTIVITIES)))

# Users per application-events query. Keeps the `user in [...]` filter
# (and the URL) a sane length.
EMAIL_BATCH: int = 50

OUTPUT_DIR: str = os.path.join(SCRIPT_DIR, "output")
LOG_DIR: str = os.path.join(SCRIPT_DIR, "logs")

# Columns kept in the mapped file. Missing fields export blank.
UBA_COLUMNS: List[str] = ["user", "alert_name", "activity", "app",
                          "timestamp", "severity", "count", "_id"]
COUNT_COLUMNS: List[str] = ["Matched Events", "Total Size (MB)"]

# Excel refuses cells over 32,767 characters; long file lists are cut here.
CELL_LIMIT: int = 32000

# ---------------------------------------------------------------------------


def norm(value: Any) -> str:
    """'User@X.com ' == 'user@x.com', 'Upload' == 'upload'. NaN -> ''."""
    return "" if pd.isna(value) else str(value).strip().casefold()


def match_key(users: pd.Series, activities: pd.Series) -> pd.Series:
    """Join key: email + activity, so an Upload alert gets uploaded files."""
    # astype(str): an empty Series maps to float64 and "+" would then fail.
    return users.map(norm).astype(str) + "|" + activities.map(norm).astype(str)


def join_unique(values: pd.Series) -> str:
    """Distinct non-blank values, sorted, '; '-joined, capped for Excel."""
    text = "; ".join(sorted({str(v).strip() for v in values.dropna()} - {""}))
    return text[:CELL_LIMIT]


def summarise_events(app_rows: List[Dict[str, Any]]) -> pd.DataFrame:
    """One row per email + activity: object names, count, MB, apps, dates."""
    ev = pd.DataFrame(app_rows).reindex(
        columns=["user", "activity", "app", "object", "file_size", "timestamp"])
    ev["_key"] = match_key(ev["user"], ev["activity"])
    ev["mb"] = pd.to_numeric(ev["file_size"], errors="coerce").fillna(0) / (1024 * 1024)
    ev["ts"] = pd.to_numeric(ev["timestamp"], errors="coerce")

    g = ev.groupby("_key")
    return pd.DataFrame({
        "Object Names": g["object"].agg(join_unique),
        "Matched Events": g.size(),
        "Total Size (MB)": g["mb"].sum().round(2),
        "Event Apps": g["app"].agg(join_unique),
        "First Event": g["ts"].min().map(ns.format_epoch),
        "Last Event": g["ts"].max().map(ns.format_epoch),
    })


def map_events(uba_rows: List[Dict[str, Any]],
               app_rows: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    Exactly one row per UBA alert. The application events with the same
    email AND activity are summarised onto it -- that is where the object
    (file) names come from, since UBA alerts don't carry them.
    No matching events -> blank names, zero counts.
    """
    uba = pd.DataFrame(uba_rows).reindex(columns=UBA_COLUMNS)
    uba["_key"] = match_key(uba["user"], uba["activity"])
    uba["timestamp"] = uba["timestamp"].map(ns.format_epoch)
    uba = uba.rename(columns={c: f"UBA {c}" for c in UBA_COLUMNS})

    mapped = uba.join(summarise_events(app_rows), on="_key").drop(columns="_key")
    mapped[COUNT_COLUMNS] = mapped[COUNT_COLUMNS].fillna(0)
    mapped["Matched Events"] = mapped["Matched Events"].astype(int)
    return mapped.fillna("")


def resolve_range():
    """
    --days N -> the N full days ending yesterday (today is excluded, it's
    still filling up). Otherwise the usual single day: CLI date or yesterday.
    """
    if "--days" not in sys.argv:
        return ns.prompt_for_date_range()
    days = int(sys.argv[sys.argv.index("--days") + 1])
    if days < 1:
        raise ValueError("--days must be 1 or more")
    end = ns.local_midnight(time.time())
    start = end
    for _ in range(days):                 # DST-safe, unlike end - days*86400
        start = ns.previous_local_midnight(start)
    day = lambda t: time.strftime("%Y-%m-%d", time.localtime(t))
    last = ns.previous_local_midnight(end)
    label = day(start) if days == 1 else f"{day(start)}_to_{day(last)}"
    return start, end - ns.BOUNDARY_TRIM, label


def fetch(session, endpoint: str, query: str, start: int, end: int,
          stats: Dict[str, int]) -> List[Dict[str, Any]]:
    ns.ENDPOINT_PATH = endpoint
    return [ns.flatten_record(r)
            for r in ns.fetch_all(session, query, start, end, stats)]


def write(rows, path: str) -> None:
    frame = rows if isinstance(rows, pd.DataFrame) else pd.DataFrame(rows)
    frame.to_csv(path, index=False, encoding="utf-8")
    ns.log.info("wrote      %s (%d rows)", path, len(frame))


def main() -> int:
    for name in ("API_TOKEN", "BASE_URL", "AUTH_MODE"):
        setattr(ns, name, globals()[name])
    ns.LOG_DIR, ns.LOG_BASENAME = LOG_DIR, "bulk_upload_uba.log"
    ns.SEND_FIELDS_PARAM = False      # ns's field list is for page events
    ns.setup_logging(time.strftime("%H%M%S"))

    if "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL:
        ns.log.error("CONFIG ERROR: set API_TOKEN / BASE_URL at the top of %s",
                     os.path.basename(__file__))
        return 1

    try:
        start, end, label = resolve_range()
    except (ValueError, IndexError) as exc:
        ns.log.error("BAD DATE / --days: %s (use --days 7 or YYYY-MM-DD)", exc)
        return 1

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out = lambda kind: os.path.join(OUTPUT_DIR, f"UBA_{kind}_{label}.csv")
    session = ns.build_session()
    stats: Dict[str, int] = {"calls": 0}

    try:
        ns.log.info("day        %s", label)
        ns.log.info("UBA query  %s", UBA_QUERY)
        alerts = fetch(session, ALERT_ENDPOINT, UBA_QUERY, start, end, stats)
        ns.log.info("UBA alerts %d", len(alerts))
        if not alerts:
            return 3
        write(alerts, out("BulkTransfer"))

        emails = sorted({str(a["user"]).strip() for a in alerts if a.get("user")})
        ns.log.info("users      %d", len(emails))
        events: List[Dict[str, Any]] = []
        activity = ns.build_term("activity", "in", ",".join(ACTIVITIES))
        for i in range(0, len(emails), EMAIL_BATCH):
            users = ns.build_term("user", "in", ",".join(emails[i:i + EMAIL_BATCH]))
            query = f"{users} and {activity}"
            events += fetch(session, APP_ENDPOINT, query, start, end, stats)
        write(events, out("AppEvents"))
    except ns.ApiError as exc:
        ns.log.error("NETSKOPE FAILED: %s", exc)
        return 2

    write(map_events(alerts, events), out("Mapped"))
    ns.log.info("API calls  %d", stats["calls"])
    return 0


def selftest() -> None:
    assert UBA_QUERY == ("alert_type eq 'uba' and activity in "
                         "['Upload', 'Download', 'Download All']")
    alerts = [
        {"user": "A@x.com", "activity": "Upload", "app": "Box"},
        {"user": "a@x.com", "activity": "Download All", "app": "Box"},
        {"user": "b@x.com", "activity": "Download", "app": "Drive"},
    ]

    mb = 1024 * 1024
    mapped = map_events(alerts, [
        {"user": "a@x.com ", "app": "ChatGPT", "activity": "upload",
         "object": "a.xlsx", "file_size": 2 * mb, "timestamp": 1760000000},
        {"user": "a@x.com", "app": "Gemini", "activity": "Upload",
         "object": "b.csv", "file_size": mb, "timestamp": 1760086400},
        {"user": "A@X.COM", "app": "Box", "activity": "Download All",
         "object": "c.pdf", "file_size": mb},
        {"user": "b@x.com", "app": "Drive", "activity": "Upload",
         "object": "wrong-activity.txt"},
        {"user": "z@x.com", "app": "Box", "activity": "Upload"},
    ])
    assert len(mapped) == len(alerts)             # one row per alert, always
    up, dl_all, b = mapped.iloc[0], mapped.iloc[1], mapped.iloc[2]
    # Upload alert gets only uploaded files, not the Download All one.
    assert up["Object Names"] == "a.xlsx; b.csv"
    assert (up["Matched Events"], up["Total Size (MB)"]) == (2, 3.0)
    assert up["Event Apps"] == "ChatGPT; Gemini"
    assert up["First Event"] and up["Last Event"]
    assert (dl_all["Object Names"], dl_all["Matched Events"]) == ("c.pdf", 1)
    # b only uploaded, so their Download alert matches nothing.
    assert (b["Object Names"], b["Matched Events"]) == ("", 0)
    assert len(map_events(alerts, [])) == len(alerts)
    print("selftest OK")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        sys.exit(0)
    sys.exit(main())
