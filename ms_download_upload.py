#!/usr/bin/env python3
"""
Files a user downloaded from a Microsoft app and then uploaded to any other
app -- one alert per such upload.

Scope is the users with a UBA alert whose activity is one of UBA_ACTIVITIES
(or --activity) in the date range -- the Skope IT filter that
Netskope-Bulk-Upload/bulk_upload_uba.py uses. For each of them every
Download from MICROSOFT_APPS and every Upload is pulled, and an Upload to an
app outside MICROSOFT_APPS alerts when the same user downloaded the same
file from a Microsoft app at most MAX_GAP_HOURS (--hours) before it.
Downloads are pulled from that many hours before the first day too:
    same file   equal MD5, or equal file name ignoring case, spacing and a
                " (1)" / " - Copy" suffix (GENERIC_NAME names such as
                image.png pair on MD5 only)
The alert names the latest such download.

    UBA alerts   /api/v2/events/datasearch/alert         alert_type eq 'uba' and activity in UBA_ACTIVITIES
    Downloads    /api/v2/events/datasearch/application   those users, MICROSOFT_APPS
    Uploads      /api/v2/events/datasearch/application   those users, any app

Output output/MS_Download_Upload_<from>_to_<to>.xlsx, a filter on every
header row:
    Alerts       one row per download -> upload
    Events       every download / upload pulled, in the Sheet1 layout
                 (Time (GMT), Action, Activity, User, Application, ...) plus
                 the UBA activity that alerted on it and its Alert IDs
    UBA Alerts   the UBA alerts that put each user in scope
A sheet past Excel's row limit continues on "<sheet> 2", "<sheet> 3", ...

Everything this script needs is configured below, starting with API_TOKEN
and BASE_URL; export_userwise.py and Netskope-API only lend it code.

Usage:
    python ms_download_upload.py 2026-10-01 2026-10-27   -> both days inclusive
    python ms_download_upload.py 2026-10-01 2026-10-27 --activity "Bulk Upload All" --activity Upload
    python ms_download_upload.py 2026-10-01 2026-10-27 --hours 8
    python ms_download_upload.py --selftest              -> offline check
"""

import argparse
import math
import os
import re
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Any, Callable, Dict, Iterable, List, Set, Tuple

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import export_userwise as eu

ns = eu.ns

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
# SECURITY: the token lives in this file -- keep it out of git / shares
# (this repo is public).
API_TOKEN: str = "PUT_YOUR_API_TOKEN_HERE"
BASE_URL: str = "https://YOUR_TENANT.goskope.com"
AUTH_MODE: str = "bearer"                 # flip to "netskope" on 401s

# A download from one of these can start an alert; an upload to one never
# ends one. Netskope app names, compared ignoring case.
MICROSOFT_APPS: List[str] = [
    "Microsoft Office 365 OneDrive for Business",
    "Microsoft Office 365 Sharepoint Online",
    "Microsoft Teams",
    "Microsoft Office 365 Outlook.com",
]

# UBA alert activities that put a user in scope, exact Netskope spelling
# (Skope IT: Alert Type = UBA, Activity in these). --activity picks some (or
# others, e.g. "Download All") for one run.
UBA_ACTIVITIES: List[str] = ["Bulk Upload All", "Bulk Download All", "Upload", "Download"]

DOWNLOAD_ACTIVITIES: List[str] = ["Download"]
UPLOAD_ACTIVITIES: List[str] = ["Upload"]      # add "Attach" for webmail attachments

# Most hours from a download to an upload that still alerts; --hours
# overrides it per run. 0 = any time within the range.
MAX_GAP_HOURS: float = 24

# File names too common to pair on by name alone (image.png, image001.png,
# blob, ...), matched after lower-casing; they still pair on MD5.
GENERIC_NAME = re.compile(r"^(image\d*|img\d*|blob|file|untitled)(\.\w+)?$")
# " (1)", "(2)", " - Copy", " - Copy (2)" at the end of a file name.
COPY_SUFFIX = re.compile(r"(\s*\(\d+\)|\s+-\s+copy)+$")

ALERT_ENDPOINT: str = "/api/v2/events/datasearch/alert"
EVENT_ENDPOINT: str = "/api/v2/events/datasearch/application"

# Users per query, keeps the URL a sane length.
USER_BATCH: int = 50

# API field -> Excel header, the Sheet1 layout. Fields an event lacks
# export blank.
COLUMNS: Dict[str, str] = {
    "timestamp": "Time (GMT)",     # Excel date, yyyy-mm-dd h:mm:ss
    "action": "Action",
    "activity": "Activity",
    "user": "User",
    "app": "Application",
    "userip": "Source IP",
    "hostname": "Host Name",
    "url": "URL",
    "object": "Object",
    "object_type": "Object Type",
    "srcip": "Source IP (Egress)",
    "src_country": "Source Country",
    "file_path": "File_path",
    "app_activity": "App_activity",
}
# Further fields per endpoint, after COLUMNS.
ALERT_COLUMNS: Dict[str, str] = {
    "alert_name": "Alert Name",
    "severity": "Severity",
    "count": "Count",
}
EVENT_COLUMNS: Dict[str, str] = {"md5": "MD5"}

ALERT_HEADER: List[str] = [
    "Alert ID", "User", "File",
    "Downloaded (GMT)", "Download App", "Download App_activity", "Download File_path",
    "Uploaded (GMT)", "Upload App", "Upload Host Name", "Upload URL",
    "Upload Source IP (Egress)", "Upload Source Country",
    "Gap (Hours)", "Matching MS Downloads", "Matched On",
    "UBA Activity (Download)", "UBA Activity (Upload)", "UBA Activities (User)",
]
EVENT_HEADER: List[str] = (list(COLUMNS.values()) + list(EVENT_COLUMNS.values())
                           + ["UBA Activity", "Alert IDs"])
UBA_HEADER: List[str] = list(COLUMNS.values()) + list(ALERT_COLUMNS.values())

# Column widths; anything else gets its header length.
WIDE_HEADERS = {"URL", "Upload URL", "Object", "File", "File_path",
                "Download File_path", "Alert Name", "Application",
                "Download App", "Upload App", "UBA Activities (User)"}

OUTPUT_DIR: str = eu.OUTPUT_DIR
LOG_DIR: str = eu.LOG_DIR

# ---------------------------------------------------------------------------


def lower(value: Any) -> str:
    return str(value).strip().lower()


def seconds(event: Dict[str, Any]) -> int:
    return int(float(event["timestamp"] or 0))


def event_order(event: Dict[str, Any]) -> Tuple[str, int, bool]:
    """User, then time; a download before an upload in the same second."""
    return (lower(event["user"]), seconds(event),
            lower(event["activity"]) in {lower(a) for a in UPLOAD_ACTIVITIES})


def direction(activity: Any) -> str:
    """"upload" / "download" for any activity naming one (Bulk Upload All,
    Download All, ...), else the activity itself, lower-cased."""
    name = lower(activity)
    return next((d for d in ("upload", "download") if d in name), name)


def alert_fields() -> List[str]:
    return list(COLUMNS) + list(ALERT_COLUMNS)


def event_fields() -> List[str]:
    return list(COLUMNS) + list(EVENT_COLUMNS)


def uba_query(activities: List[str]) -> str:
    """Same filter as bulk_upload_uba.py: UBA alerts with these activities."""
    return f"alert_type eq 'uba' and {eu.build_query(['activity'], activities)}"


def event_queries(users: List[str]) -> List[Tuple[str, str]]:
    """(label, query) for one batch of users: their Microsoft downloads,
    then all their uploads."""
    who = eu.build_query(["user"], users)
    return [
        ("Downloads", f"{who} and {eu.build_query(['activity'], DOWNLOAD_ACTIVITIES)}"
                      f" and {eu.build_query(['app'], MICROSOFT_APPS)}"),
        ("Uploads", f"{who} and {eu.build_query(['activity'], UPLOAD_ACTIVITIES)}"),
    ]


def file_name(event: Dict[str, Any]) -> str:
    """The object, else the file path, without any folders."""
    name = event.get("object") or event.get("file_path") or ""
    return re.split(r"[\\/]", str(name))[-1].strip()


def name_key(name: str) -> str:
    """Comparable file name: lower case, single spaces, no COPY_SUFFIX;
    "" for a GENERIC_NAME."""
    stem, ext = os.path.splitext(" ".join(name.lower().split()))
    key = COPY_SUFFIX.sub("", stem) + ext
    return "" if GENERIC_NAME.match(key) else key


def file_keys(event: Dict[str, Any]) -> List[str]:
    """"md5:<hash>" and "name:<name_key>", each when known."""
    keys = []
    if event.get("md5"):
        keys.append("md5:" + lower(event["md5"]))
    name = name_key(file_name(event))
    if name:
        keys.append("name:" + name)
    return keys


def download_windows(first: str, last: str, hours: float) -> List[Tuple[int, int]]:
    """Day windows for downloads: `first`..`last` plus the days before
    `first` an upload on `first` can reach back to within `hours`."""
    day = datetime.strptime(first, "%Y-%m-%d") - timedelta(days=math.ceil(hours / 24))
    return eu.day_windows(day.strftime("%Y-%m-%d"), last)


def correlate(events: List[Dict[str, Any]], max_gap_hours: float) -> List[Dict[str, Any]]:
    """One alert per upload to a non-Microsoft app of a file the same user
    downloaded from a Microsoft app at most `max_gap_hours` before it (0 =
    any time), naming the latest such download. Sorts `events` by user then
    time and lists each event's Alert IDs under "alert_ids"."""
    microsoft = {lower(a) for a in MICROSOFT_APPS}
    downloads = {lower(a) for a in DOWNLOAD_ACTIVITIES}
    uploads = {lower(a) for a in UPLOAD_ACTIVITIES}
    max_gap = max_gap_hours * 3600 or float("inf")

    events.sort(key=event_order)
    found: List[Dict[str, Any]] = []
    user = None
    by_file: Dict[str, List[Dict[str, Any]]] = {}   # file key -> MS downloads so far
    for event in events:
        if lower(event["user"]) != user:
            user, by_file = lower(event["user"]), defaultdict(list)
        activity = lower(event["activity"])
        from_microsoft = lower(event["app"]) in microsoft
        keys = file_keys(event)
        if activity in downloads and from_microsoft:
            for key in keys:
                by_file[key].append(event)
        elif activity in uploads and not from_microsoft:
            earlier = {id(d): d for key in keys for d in by_file.get(key, ())
                       if seconds(event) - seconds(d) <= max_gap}
            if not earlier:
                continue
            source = max(earlier.values(), key=seconds)
            shared = set(file_keys(source)) & set(keys)
            alert_id = f"A{len(found) + 1:04d}"
            found.append({
                "id": alert_id, "download": source, "upload": event,
                "earlier": len(earlier),
                "matched": " + ".join(label for prefix, label in
                                      (("md5:", "MD5"), ("name:", "File name"))
                                      if any(k.startswith(prefix) for k in shared)),
            })
            for side in (source, event):
                side.setdefault("alert_ids", []).append(alert_id)
    return found


def uba_labels(alerts: List[Dict[str, Any]]) -> Tuple[Callable[[Dict[str, Any]], str],
                                                       Dict[str, Set[str]]]:
    """(UBA activities that alerted on an event's user, direction and app --
    or user and direction, for an alert with no app; every UBA activity per
    user). "Bulk Upload All" labels uploads, "Bulk Download All" downloads."""
    by_event: Dict[Tuple[str, str, str], Set[str]] = defaultdict(set)
    by_user: Dict[str, Set[str]] = defaultdict(set)
    for alert in alerts:
        user = lower(alert["user"])
        by_event[(user, direction(alert["activity"]), lower(alert["app"]))].add(
            alert["activity"])
        by_user[user].add(alert["activity"])

    def event_label(event: Dict[str, Any]) -> str:
        user, way = lower(event["user"]), direction(event["activity"])
        names = (by_event.get((user, way, lower(event["app"])), set())
                 | by_event.get((user, way, ""), set()))
        return ", ".join(sorted(names))

    return event_label, by_user


def sheet1_row(record: Dict[str, Any]) -> List[Any]:
    return [eu.to_utc(record[f]) if f == "timestamp" else record[f] for f in COLUMNS]


def alert_row(found: Dict[str, Any], event_label: Callable[[Dict[str, Any]], str],
              by_user: Dict[str, Set[str]]) -> List[Any]:
    down, up = found["download"], found["upload"]
    return [
        found["id"], up["user"], file_name(up) or file_name(down),
        eu.to_utc(down["timestamp"]), down["app"], down["app_activity"], down["file_path"],
        eu.to_utc(up["timestamp"]), up["app"], up["hostname"], up["url"],
        up["srcip"], up["src_country"],
        round((seconds(up) - seconds(down)) / 3600, 2), found["earlier"], found["matched"],
        event_label(down), event_label(up),
        ", ".join(sorted(by_user.get(lower(up["user"]), ()))),
    ]


def column_width(header: str) -> int:
    if header in WIDE_HEADERS:
        return 40
    if "(GMT)" in header:
        return 20
    return max(12, len(header) + 2)


def add_sheet(book: Workbook, title: str, header: List[str],
              rows: Iterable[List[Any]]) -> int:
    """Header frozen with a filter on it, rolling over to "<title> 2", ...
    at the Excel row limit. Returns the row count."""
    sheets, count = [], 0

    def new_sheet() -> None:
        sheet = book.create_sheet(title if not sheets else f"{title} {len(sheets) + 1}")
        sheet.freeze_panes = "A2"
        for i, name in enumerate(header, 1):
            sheet.column_dimensions[get_column_letter(i)].width = column_width(name)
        sheet.append(header)
        sheets.append(sheet)

    for row in rows:
        if count % eu.SHEET_ROWS == 0:
            new_sheet()
        sheets[-1].append(row)
        count += 1
    if not sheets:
        new_sheet()
    last_column = get_column_letter(len(header))
    for i, sheet in enumerate(sheets):
        filled = min(eu.SHEET_ROWS, count - i * eu.SHEET_ROWS)
        sheet.auto_filter.ref = f"A1:{last_column}{filled + 1}"
    return count


def write_report(path: str, found: List[Dict[str, Any]], events: List[Dict[str, Any]],
                 alerts: List[Dict[str, Any]]) -> None:
    event_label, by_user = uba_labels(alerts)
    book = Workbook(write_only=True)
    add_sheet(book, "Alerts", ALERT_HEADER,
              (alert_row(f, event_label, by_user) for f in found))
    add_sheet(book, "Events", EVENT_HEADER,
              (sheet1_row(e) + [e[f] for f in EVENT_COLUMNS]
               + [event_label(e), ", ".join(e.get("alert_ids", []))] for e in events))
    add_sheet(book, "UBA Alerts", UBA_HEADER,
              (sheet1_row(a) + [a[f] for f in ALERT_COLUMNS]
               for a in sorted(alerts, key=lambda a: (lower(a["user"]), seconds(a)))))
    book.save(path)


def fetch(session, stats: Dict[str, int], endpoint: str, fields: List[str],
          query: str, windows: List[Tuple[int, int]], label: str) -> List[Dict[str, Any]]:
    """Every event for `query`, a day at a time, as {field: value} for
    `fields` only; rows identical in every field dropped."""
    ns.ENDPOINT_PATH = endpoint
    ns.DESIRED_COLUMNS = fields
    rows: Dict[Tuple[Any, ...], Dict[str, Any]] = {}
    for start, end in windows:
        before = len(rows)
        for record in ns.fetch_all(session, query, start, end, stats):
            flat = ns.flatten_record(record)
            row = {f: "" if flat.get(f) is None else flat[f] for f in fields}
            rows.setdefault(tuple(row.values()), row)
        ns.log.info("%s %-10s %d events (total %d)",
                    time.strftime("%Y-%m-%d", time.localtime(start)),
                    label, len(rows) - before, len(rows))
    return list(rows.values())


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Files downloaded from a Microsoft app, then uploaded elsewhere.")
    parser.add_argument("start", help="first day, e.g. 2026-10-01")
    parser.add_argument("end", help="last day, inclusive, e.g. 2026-10-27")
    parser.add_argument("--activity", action="append", metavar="NAME",
                        help="UBA alert activity that puts a user in scope, repeatable "
                             f"(default: {', '.join(UBA_ACTIVITIES)})")
    parser.add_argument("--hours", type=float, default=MAX_GAP_HOURS,
                        help="most hours from download to upload, 0 = any "
                             f"(default: {MAX_GAP_HOURS:g})")
    args = parser.parse_args(argv)
    if args.hours < 0:
        parser.error("--hours must be 0 or more")
    return args


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    for name in ("API_TOKEN", "BASE_URL", "AUTH_MODE"):
        setattr(ns, name, globals()[name])
    ns.LOG_DIR, ns.LOG_BASENAME = LOG_DIR, "ms_download_upload.log"
    # Ask only for the exported fields. Unknown names are ignored by the API.
    ns.SEND_FIELDS_PARAM = True
    ns.PLACEHOLDER_COLUMNS = set()
    ns.setup_logging(time.strftime("%H%M%S"))

    if "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL:
        ns.log.error("CONFIG ERROR: set API_TOKEN / BASE_URL at the top of %s",
                     os.path.basename(__file__))
        return 1
    try:
        windows = eu.day_windows(args.start, args.end)
        reach_back = download_windows(args.start, args.end, args.hours)
    except ValueError as exc:
        ns.log.error("BAD DATES: %s (use: %s 2026-10-01 2026-10-27)",
                     exc, os.path.basename(__file__))
        return 1
    activities = args.activity or UBA_ACTIVITIES
    label = f"{args.start}_to_{args.end}"
    session = ns.build_session()
    stats: Dict[str, int] = {"calls": 0}
    ns.log.info("range      %s (%d days)", label, len(windows))
    ns.log.info("UBA        %s", ", ".join(activities))
    ns.log.info("window     upload within %s of the download",
                f"{args.hours:g} hours" if args.hours else "any time")

    try:
        alerts = fetch(session, stats, ALERT_ENDPOINT, alert_fields(),
                       uba_query(activities), windows, "UBA alerts")
        users = list(dict.fromkeys(str(a["user"]) for a in alerts if a["user"] != ""))
        if not users:
            ns.log.warning("no UBA alerts with activity %s in this range -- names "
                           "must match Netskope exactly", ", ".join(activities))
            return 3
        ns.log.info("users      %d with UBA alerts", len(users))

        events: List[Dict[str, Any]] = []
        for i in range(0, len(users), USER_BATCH):
            for kind, query in event_queries(users[i:i + USER_BATCH]):
                events += fetch(session, stats, EVENT_ENDPOINT, event_fields(), query,
                                reach_back if kind == "Downloads" else windows, kind)
    except ns.ApiError as exc:
        ns.log.error("NETSKOPE FAILED: %s", exc)
        return 2

    found = correlate(events, args.hours)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    xlsx = os.path.join(OUTPUT_DIR, f"MS_Download_Upload_{label}.xlsx")
    write_report(xlsx, found, events, alerts)
    ns.log.info("alerts     %d (%d users)", len(found),
                len({lower(f["upload"]["user"]) for f in found}))
    ns.log.info("wrote      %s", xlsx)
    ns.log.info("API calls  %d", stats["calls"])
    return 0


def selftest() -> None:
    import tempfile
    from openpyxl import load_workbook

    od, sp = MICROSOFT_APPS[0], MICROSOFT_APPS[1]

    def ev(user, ts, activity, app, obj="", md5="", path=""):
        row = dict.fromkeys(event_fields(), "")
        row.update(user=user, timestamp=ts, activity=activity, app=app,
                   object=obj, md5=md5, file_path=path)
        return row

    def sample():
        return [
            ev("a", 100, "Download", od, "Report.xlsx"),
            ev("a", 200, "Download", sp, "report (1).xlsx"),
            ev("a", 300, "Upload", "Notion", "Report.XLSX"),      # -> 200
            ev("a", 400, "Upload", sp, "Report.xlsx"),            # Microsoft, no
            ev("a", 50, "Upload", "Gmail", "Report.xlsx"),        # before, no
            ev("a", 450, "Download", "Microsoft Teams", "image.png"),
            ev("a", 500, "Upload", "WhatsApp", "image.png"),      # generic, no
            ev("b", 600, "Upload", "Notion", "Report.xlsx"),      # other user, no
            ev("a", 650, "Download", "Microsoft Office 365 Outlook.com", "orig.pdf", md5="M1"),
            ev("a", 700, "Upload", "Dropbox", "renamed.pdf", md5="m1"),  # MD5 -> 650
            ev("a", 800, "Download", "Box", "x.doc"),
            ev("a", 900, "Upload", "Notion", "x.doc"),            # not from Microsoft, no
        ]

    assert name_key("Budget v2 (1).xlsx") == "budget v2.xlsx"
    assert name_key("Doc - Copy (2).docx") == name_key("doc(3).DOCX") == "doc.docx"
    assert name_key("  Two   Spaces.PDF") == "two spaces.pdf"
    assert name_key("image001.png") == name_key("IMAGE (2).PNG") == name_key("") == ""
    assert file_name(ev("a", 1, "Download", od, path="Documents/Team/Budget.xlsx")) == "Budget.xlsx"

    events = sample()
    found = correlate(events, 0)
    assert [(seconds(f["download"]), seconds(f["upload"]), f["earlier"], f["matched"], f["id"])
            for f in found] == [(200, 300, 2, "File name", "A0001"), (650, 700, 1, "MD5", "A0002")]
    assert [e["user"] for e in events] == ["a"] * 11 + ["b"]
    assert [seconds(e) for e in events if e.get("alert_ids")] == [200, 300, 650, 700]

    pair = [ev("a", 1, "Download", od, "image.png", md5="m2"),
            ev("a", 2, "Upload", "Notion", "image.png", md5="m2")]
    assert [f["matched"] for f in correlate(pair, 0)] == ["MD5"]     # generic name, same MD5
    same = [ev("a", 5, "Upload", "Notion", "r.pdf"), ev("a", 5, "Download", od, "r.pdf")]
    assert len(correlate(same, 24)) == 1                              # same second counts

    found = correlate(sample(), 0.05)                                 # 180 s
    assert [(seconds(f["download"]), f["earlier"]) for f in found] == [(200, 1), (650, 1)]
    assert [seconds(f["upload"]) for f in correlate(sample(), 0.02)] == [700]   # 72 s
    day = 24 * 3600
    late = [ev("a", 0, "Download", od, "r.pdf"), ev("a", day + 1, "Upload", "Notion", "r.pdf")]
    assert correlate(late, 24) == [] and len(correlate(late, 25)) == 1

    assert len(download_windows("2026-10-16", "2026-10-17", 0)) == 2
    reach = download_windows("2026-10-16", "2026-10-17", 24)
    assert len(reach) == 3 and time.localtime(reach[0][0])[:3] == (2026, 10, 15)
    assert len(download_windows("2026-10-16", "2026-10-17", 25)) == 4
    args = parse_args(["2026-10-16", "2026-10-17"])
    assert (args.hours, args.activity) == (MAX_GAP_HOURS, None) and MAX_GAP_HOURS == 24
    args = parse_args(["2026-10-16", "2026-10-17", "--hours", "8", "--activity", "Upload"])
    assert (args.hours, args.activity) == (8, ["Upload"])

    assert uba_query(UBA_ACTIVITIES) == (
        "alert_type eq 'uba' and activity in "
        "['Bulk Upload All', 'Bulk Download All', 'Upload', 'Download']")
    (dl_kind, dl_query), (up_kind, up_query) = event_queries(["a@x.com", "b"])
    assert (dl_kind, up_kind) == ("Downloads", "Uploads")
    assert dl_query == ("user in ['a@x.com', 'b'] and activity in ['Download'] and app in "
                        "['Microsoft Office 365 OneDrive for Business', 'Microsoft Office 365 "
                        "Sharepoint Online', 'Microsoft Teams', 'Microsoft Office 365 Outlook.com']")
    assert up_query == "user in ['a@x.com', 'b'] and activity in ['Upload']"

    assert [direction(a) for a in ("Bulk Upload All", "Download All", "upload", "Login")] == [
        "upload", "download", "upload", "login"]

    def al(user, ts, activity, app):
        row = dict.fromkeys(alert_fields(), "")
        row.update(user=user, timestamp=ts, activity=activity, app=app)
        return row

    alerts = [al("A", 900, "Bulk Upload All", "Notion"),
              al("a", 10, "Download", "")]
    event_label, by_user = uba_labels(alerts)
    assert event_label(ev("a", 1, "Upload", "notion")) == "Bulk Upload All"
    assert event_label(ev("a", 1, "Download", od)) == "Download"
    assert event_label(ev("a", 1, "Upload", "Box")) == ""
    assert by_user["a"] == {"Bulk Upload All", "Download"}

    # fetch keeps only the asked fields and drops rows that only differ in _id.
    saved_fetch = ns.fetch_all
    ns.fetch_all = lambda session, query, start, end, stats: [
        {"_id": 1, "user": "a", "app": "X", "extra": 5}, {"_id": 2, "user": "a", "app": "X"}]
    try:
        rows = fetch(None, {"calls": 0}, EVENT_ENDPOINT, ["user", "app", "md5"], "q", [(0, 1)], "t")
        assert rows == [{"user": "a", "app": "X", "md5": ""}]
        assert ns.ENDPOINT_PATH == EVENT_ENDPOINT
    finally:
        ns.fetch_all = saved_fetch

    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    events = sample()
    found = correlate(events, 24)
    write_report(path, found, events, alerts)
    book = load_workbook(path)
    assert book.sheetnames == ["Alerts", "Events", "UBA Alerts"]
    sheet = book["Alerts"]
    assert [c.value for c in sheet[1]] == ALERT_HEADER
    assert sheet.auto_filter.ref == "A1:S3" and sheet.freeze_panes == "A2"
    first = dict(zip(ALERT_HEADER, [c.value for c in sheet[2]]))
    assert (first["Alert ID"], first["User"], first["File"], first["Upload App"],
            first["Download App"], first["Gap (Hours)"], first["Matched On"],
            first["UBA Activity (Download)"], first["UBA Activity (Upload)"],
            first["UBA Activities (User)"]) == (
        "A0001", "a", "Report.XLSX", "Notion", sp, 0.03, "File name",
        "Download", "Bulk Upload All", "Bulk Upload All, Download")
    assert first["Uploaded (GMT)"].year == 1970                         # an Excel date
    sheet = book["Events"]
    assert [c.value for c in sheet[1]][:14] == [
        "Time (GMT)", "Action", "Activity", "User", "Application", "Source IP",
        "Host Name", "URL", "Object", "Object Type", "Source IP (Egress)",
        "Source Country", "File_path", "App_activity"]               # Sheet1 layout
    assert sheet.max_row == 13 and sheet.auto_filter.ref == "A1:Q13"
    epoch = datetime(1970, 1, 1)
    assert [(int((r[0].value - epoch).total_seconds()), r[16].value)
            for r in sheet.iter_rows(min_row=2) if r[16].value] == [
        (200, "A0001"), (300, "A0001"), (650, "A0002"), (700, "A0002")]
    sheet = book["UBA Alerts"]
    assert [c.value for c in sheet[1]] == UBA_HEADER
    assert [(r[2].value, r[3].value) for r in sheet.iter_rows(min_row=2)] == [
        ("Download", "a"), ("Bulk Upload All", "A")]

    saved_rows, eu.SHEET_ROWS = eu.SHEET_ROWS, 2
    try:
        book = Workbook(write_only=True)
        assert add_sheet(book, "T", ["x", "y"], [[1, 2], [3, 4], [5, 6]]) == 3
        assert add_sheet(book, "E", ["x"], []) == 0
        book.save(path)
        book = load_workbook(path)
        assert book.sheetnames == ["T", "T 2", "E"]
        assert [book[n].auto_filter.ref for n in book.sheetnames] == ["A1:B3", "A1:B2", "A1:A1"]
    finally:
        eu.SHEET_ROWS = saved_rows
    print("selftest OK")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        sys.exit(0)
    sys.exit(main(sys.argv[1:]))
