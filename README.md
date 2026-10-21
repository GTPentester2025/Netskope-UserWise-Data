# Netskope User-wise Data

Pulls user-wise Netskope events for your own list of apps and of websites,
for a date range, into two Excel files.

## Setup

Keep this folder next to [Netskope-API](https://github.com/black0405/Netskope-API);
the HTTP, retry and paging logic is imported from
`../Netskope-API/export_genai_page_bytes.py`.

    pip install requests pandas openpyxl

1. Set `API_TOKEN` and `BASE_URL` at the top of `export_userwise.py`
   (do not commit them).
2. Create the lists (one entry per line, `#` for comments, both git-ignored;
   a missing or empty file skips that report):
   - `apps.txt` - app names in exact Netskope spelling, e.g. `ChatGPT`
     (matched against `app`)
   - `websites.txt` - domains / hostnames, e.g. `chat.openai.com`
     (matched against `site`, `domain`, `hostname`)

## Run

    python export_userwise.py 2026-08-01 2026-09-30   # both days inclusive
    python export_userwise.py --selftest              # offline check

## Output (`output/`, one row per event, sorted by user)

`Apps_<from>_to_<to>.xlsx`
- **Activities** (`/api/v2/events/datasearch/application`): User, Application,
  URL, Activity, Object ID, Object Name, Object Type, File Path,
  File Size (Bytes), Event Date, Organization Unit
- **Traffic** (`/api/v2/events/datasearch/page`): see below

`Websites_<from>_to_<to>.xlsx`
- **Traffic**: Timestamp (UTC) Seconds, Browser, User, Source IP, OS, Device,
  User Group, Organization Unit, Website, Total Bytes, Bytes Uploaded,
  Bytes Downloaded

A sheet past Excel's 1,048,576-row limit continues on `Traffic 2`, `Traffic 3`, ...
The raw per-sheet CSVs are kept in `output/` as well.
