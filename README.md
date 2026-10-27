# Netskope User-wise Data

Netskope REST API v2 exports, everything needed in this one repo:

| Path | What it does |
|---|---|
| `export_userwise.py` | User-wise events for your own list of apps and of websites, for a date range, into two Excel files |
| `ms_download_upload.py` | Alerts on files downloaded from a Microsoft app and then uploaded to any other app |
| `Netskope-Bulk-Upload/bulk_upload_uba.py` | UBA bulk upload / download alerts with the users' application events mapped on ([its README](Netskope-Bulk-Upload/README.md)) |
| `Netskope-API/` | Shared HTTP, retry, paging and time-window code (`export_genai_page_bytes.py`) imported by all three, plus the GenAI exports ([its README](Netskope-API/README.md)) |

## Setup

    pip install requests pandas openpyxl

1. Set `API_TOKEN` and `BASE_URL` at the top of `export_userwise.py`
   (do not commit them); `ms_download_upload.py` reads them from there.
   `bulk_upload_uba.py` has its own copy at its top.
2. For `export_userwise.py`, create the lists (one entry per line, `#` for
   comments, both git-ignored; a missing or empty file skips that report):
   - `apps.txt` - app names in exact Netskope spelling, e.g. `ChatGPT`
     (matched against `app`)
   - `websites.txt` - the Website value from a Netskope export, e.g. `sharecad`,
     or a domain / hostname
     (matched against `site`, `domain`, `hostname`)

Exports (`output/`), logs and any CSV / XLSX are git-ignored, so user data
is never pushed.

## User-wise apps and websites

    python export_userwise.py 2026-08-01 2026-09-30   # both days inclusive
    python export_userwise.py --selftest              # offline check

One Excel file per list, `Apps_<from>_to_<to>.xlsx` and
`Websites_<from>_to_<to>.xlsx`, each with a single **Events** sheet: one row
per event, sorted by user then time, rows identical in every column dropped.

Columns: Timestamp (UTC) (Excel date, `yyyy-mm-dd h:mm:ss`), Browser, User, Source IP, OS, Device,
User Group, Organization Unit, Website, Total Bytes, Bytes Uploaded,
Bytes Downloaded (same as the Skope IT export), then Event Type, Application,
URL, Activity, Object ID, Object Name, Object Type, File Path,
File Size (Bytes).

Event Type is `Traffic` (`/api/v2/events/datasearch/page`: bytes, browser...)
or `Activity` (`/api/v2/events/datasearch/application`: activity, object,
file). Columns an event type doesn't have are blank.

A sheet past Excel's 1,048,576-row limit continues on `Events 2`, `Events 3`, ...

## Microsoft download -> other upload alerts

`ms_download_upload.py` flags a file a user downloaded from a whitelisted
Microsoft app (OneDrive for Business, SharePoint Online, Teams,
Outlook.com - `MICROSOFT_APPS`) and then uploaded to any other app. The
token needs the `alert` and `application` datasearch endpoints.

    python ms_download_upload.py 2026-10-01 2026-10-27
    python ms_download_upload.py 2026-10-01 2026-10-27 --activity "Bulk Upload All" --activity Upload
    python ms_download_upload.py 2026-10-01 2026-10-27 --hours 8
    python ms_download_upload.py --selftest

1. Users in scope: those with a UBA alert (`alert_type eq 'uba'`) whose
   activity is `Bulk Upload All`, `Bulk Download All`, `Upload` or
   `Download` (`UBA_ACTIVITIES`; `--activity` picks some) - the same Skope IT
   filter `bulk_upload_uba.py` uses.
2. For them: every Download from a Microsoft app and every Upload
   (`/api/v2/events/datasearch/application`).
3. An Upload to a non-Microsoft app alerts when the same user downloaded the
   same file from a Microsoft app at most 24 hours before it
   (`MAX_GAP_HOURS`, or `--hours`; 0 = any time) - same MD5, or same file
   name ignoring case, spacing and a ` (1)` / ` - Copy` suffix (names like
   `image.png` match on MD5 only). Downloads are also pulled from that many
   hours before the first day, so an upload early on day one still finds
   the previous night's download.

Output `output/MS_Download_Upload_<from>_to_<to>.xlsx`, filters on every header:

- **Alerts** - one row per upload: user, file, the latest matching Microsoft
  download (time, app, App_activity, File_path), the upload (time, app, host,
  URL, egress IP, country), gap in hours, how it matched, and the UBA
  activities that alerted on the download, on the upload and on the user
  (`Bulk Upload All` labels uploads, `Bulk Download All` downloads).
- **Events** - every download and upload pulled, in the Sheet1 layout (Time
  (GMT), Action, Activity, User, Application, Source IP, Host Name, URL,
  Object, Object Type, Source IP (Egress), Source Country, File_path,
  App_activity) plus MD5, UBA Activity and the Alert IDs the row belongs to.
- **UBA Alerts** - the UBA alerts that put each user in scope.

Source IP is the `userip` field and Source IP (Egress) is `srcip`.
