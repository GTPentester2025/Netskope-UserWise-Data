# Netskope User-wise Data

Pulls user-wise Netskope events (application activity + website visits) for a
fixed list of apps and websites, for a date range.

## Setup

Keep this folder next to [Netskope-API](https://github.com/black0405/Netskope-API);
the HTTP, retry and paging logic is imported from
`../Netskope-API/export_genai_page_bytes.py`.

1. Set `API_TOKEN` and `BASE_URL` at the top of `export_userwise.py`
   (do not commit them).
2. Create `targets.txt`: one app name (exact Netskope spelling, e.g. `ChatGPT`)
   or website domain/hostname per line, `#` for comments. It is git-ignored.

## Run

    python export_userwise.py 2026-08-01 2026-09-30   # both days inclusive
    python export_userwise.py --selftest              # offline check

Output: `output/UserWise_<from>_to_<to>.csv`, one row per event, sorted by user:
User, Event Type, Application, Site, Domain, Hostname, URL, Activity,
Object ID, Object Name, Object Type, File Path, File Size (Bytes),
Event Date, Organization Unit.

Each target is matched exactly against `app`, `site`, `domain` and `hostname`
on `/api/v2/events/datasearch/application` and `/api/v2/events/datasearch/page`.
