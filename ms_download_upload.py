#!/usr/bin/env python3
"""
Files a user downloaded from a Microsoft app and then uploaded to any other
app -- one alert per such upload.

One self-contained script: it calls the Netskope REST API v2 itself and
needs only `pip install requests openpyxl`. The token and tenant come from
NETSKOPE_API_TOKEN / NETSKOPE_BASE_URL (/ NETSKOPE_AUTH_MODE): the
environment first, then the KEY=value file named by MSDU_CONFIG_FILE (a
secrets manager can render it), then config.env beside this script, then
API_TOKEN / BASE_URL below. server.py is the dashboard that
runs it on demand and daily.

Scope is the users with a UBA alert whose activity is one of UBA_ACTIVITIES
(or --activity) in the date range -- the Skope IT filter Alert Type = UBA,
Activity in [...]. For each of them every Download from MICROSOFT_APPS and
every Upload is pulled, and an Upload to an app outside MICROSOFT_APPS
alerts when the same user downloaded the same file from a Microsoft app at
most MAX_GAP_HOURS (--hours) before it. Downloads are pulled from that many
hours before the first day too:
    same file   equal MD5, or equal file name ignoring case, spacing and a
                " (1)" / " - Copy" suffix (GENERIC_NAME names such as
                image.png pair on MD5 only)
The alert names the latest such download.

    UBA alerts   /api/v2/events/datasearch/alert         alert_type eq 'uba' and activity in UBA_ACTIVITIES
    Downloads    /api/v2/events/datasearch/application   those users, MICROSOFT_APPS
    Uploads      /api/v2/events/datasearch/application   those users, any app
Every filter goes to Netskope in the query; each query covers whole days
and as many users as fit in a safe URL, and WORKERS queries run at once.

Output output/MS_Download_Upload_<day>.xlsx (or _<from>_to_<to> for a
range; --out picks another folder), a filter on every header row:
    Alerts       one row per download -> upload
    Events       every download / upload pulled, in the Sheet1 layout
                 (Time (GMT), Action, Activity, User, Application, ...) plus
                 the UBA activity that alerted on it and its Alert IDs
    UBA Alerts   the UBA alerts that put each user in scope
A sheet past Excel's row limit continues on "<sheet> 2", "<sheet> 3", ...
Log: logs/ms_download_upload.log

Usage:
    python ms_download_upload.py                         -> yesterday
    python ms_download_upload.py 2026-10-18              -> that day
    python ms_download_upload.py 2026-10-01 2026-10-27   -> both days inclusive
    python ms_download_upload.py --activity "Bulk Upload All" --activity Upload
    python ms_download_upload.py --hours 8
    python ms_download_upload.py --out runs/today        -> workbook there
    python ms_download_upload.py --selftest              -> offline check

Exit codes: 0 ok, 1 config / date error, 2 Netskope API failure,
3 no matching UBA alerts.
"""

import argparse
import logging
import os
import re
import sys
import threading
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from typing import Any, Callable, Dict, Iterable, List, Set, Tuple

import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
# Last resort only -- load_config() replaces these from the environment or a
# config file. SECURITY: a token typed here lives in this file; keep it out
# of git / shares.
API_TOKEN: str = "PUT_YOUR_API_TOKEN_HERE"
BASE_URL: str = "https://YOUR_TENANT.goskope.com"
AUTH_MODE: str = "bearer"                 # flip to "netskope" on 401s
# Names the KEY=value file the credentials are read from; unset means
# config.env beside this script.
CONFIG_FILE_ENV: str = "MSDU_CONFIG_FILE"

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

# App_activity a Microsoft download needs to start an alert, and a UBA alert
# needs to go on the UBA N+ Files sheet. Blank always qualifies; compared
# ignoring case.
APP_ACTIVITIES: List[str] = ["FileDownloaded"]

# Most hours from a download to an upload that still alerts; --hours
# overrides it per run. 0 = any time within the range.
MAX_GAP_HOURS: float = 24

# UBA alerts whose name reports at least this many files ("Spike of 312
# files uploaded to ...") also get a sheet of their own. The API's `count`
# is not the number of files, so it comes from the name.
BULK_FILES: int = 300
FILE_COUNT = re.compile(r"(\d[\d,]*)\s+files?\b", re.IGNORECASE)

# File names too common to pair on by name alone (image.png, image001.png,
# blob, ...), matched after lower-casing; they still pair on MD5.
GENERIC_NAME = re.compile(r"^(image\d*|img\d*|blob|file|untitled)(\.\w+)?$")
# " (1)", "(2)", " - Copy", " - Copy (2)" at the end of a file name.
COPY_SUFFIX = re.compile(r"(\s*\(\d+\)|\s+-\s+copy)+$")

ALERT_ENDPOINT: str = "/api/v2/events/datasearch/alert"
EVENT_ENDPOINT: str = "/api/v2/events/datasearch/application"

# Users per query: as many as fit in this many characters of the
# `user in [...]` list, which keeps the request URL a safe length.
USER_LIST_CHARS: int = 4000
# Queries run at once. Netskope answers too many with HTTP 429, which is
# waited out and retried.
WORKERS: int = 4

# Events per request, and the longest span one request covers: a day.
# Lower it (e.g. 3600) only if a very busy tenant pages unreliably.
PAGE_SIZE: int = 5000
TIME_WINDOW_SECONDS: int = 86400
REQUEST_TIMEOUT: int = 60       # seconds per HTTP request
QUERY_TIMEOUT: int = 180        # datasearch's server-side timeout
MAX_RETRIES: int = 5            # for 429 / 5xx / network errors
BACKOFF_BASE: float = 2.0       # seconds; 2, 4, 8, ... unless Retry-After says

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
    "Gap (Hours)", "UBA alert_name",
]
EVENT_HEADER: List[str] = (list(COLUMNS.values()) + list(EVENT_COLUMNS.values())
                           + ["UBA Activity", "Alert IDs"])
UBA_HEADER: List[str] = list(COLUMNS.values()) + list(ALERT_COLUMNS.values())
# The UBA N+ Files sheet's columns; "files" is the count from the alert name.
BULK_FIELDS: List[str] = ["timestamp", "action", "activity", "user", "app", "srcip",
                          "src_country", "alert_name", "severity", "count", "files",
                          "app_activity"]
BULK_HEADER: List[str] = [{**COLUMNS, **ALERT_COLUMNS, "files": "Files"}[f]
                          for f in BULK_FIELDS]

# Column widths; anything else gets its header length.
WIDE_HEADERS = {"URL", "Upload URL", "Object", "File", "File_path",
                "Download File_path", "Alert Name", "Application",
                "Download App", "Upload App", "UBA alert_name"}

# Excel's 1,048,576 rows minus the header.
SHEET_ROWS: int = 1_048_575

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR: str = os.path.join(SCRIPT_DIR, "output")
LOG_DIR: str = os.path.join(SCRIPT_DIR, "logs")

# ---------------------------------------------------------------------------

log = logging.getLogger("ms_download_upload")
_stats_lock = threading.Lock()
_thread = threading.local()


class ApiError(RuntimeError):
    """A Netskope failure that retrying won't fix."""


def read_config_file(path: str) -> Dict[str, str]:
    """KEY=value lines (# comments and blanks skipped, one pair of matching
    quotes around a value dropped) -- a hand-made config.env, or the file a
    secrets manager renders."""
    values: Dict[str, str] = {}
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = (part.strip() for part in line.split("=", 1))
            if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
                value = value[1:-1]
            values[key] = value
    return values


def load_config() -> str:
    """Set API_TOKEN, BASE_URL and AUTH_MODE: environment, then the config
    file, then the values typed into this file. Returns the config file path
    looked at ("" if there was none)."""
    global API_TOKEN, BASE_URL, AUTH_MODE
    path = os.environ.get(CONFIG_FILE_ENV) or os.path.join(SCRIPT_DIR, "config.env")
    values = read_config_file(path) if os.path.isfile(path) else {}

    def pick(name: str, current: str) -> str:
        return os.environ.get(name) or values.get(name) or current

    API_TOKEN = pick("NETSKOPE_API_TOKEN", API_TOKEN)
    BASE_URL = pick("NETSKOPE_BASE_URL", BASE_URL)
    AUTH_MODE = pick("NETSKOPE_AUTH_MODE", AUTH_MODE).lower()
    return path if values else ""


def config_missing() -> bool:
    """True while the token or tenant is blank or still the placeholder."""
    return (not API_TOKEN or not BASE_URL
            or "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL)


def setup_logging() -> None:
    """Messages to the console, timestamped to logs/ms_download_upload.log."""
    os.makedirs(LOG_DIR, exist_ok=True)
    console = logging.StreamHandler()
    console.setFormatter(logging.Formatter("%(message)s"))
    logfile = logging.FileHandler(os.path.join(LOG_DIR, "ms_download_upload.log"),
                                  encoding="utf-8")
    logfile.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s"))
    log.handlers[:] = [console, logfile]
    log.setLevel(logging.INFO)


# --- Netskope API ----------------------------------------------------------

def quote(value: Any) -> str:
    """A Netskope query string literal, embedded quotes escaped."""
    return "'" + str(value).replace("'", "\\'") + "'"


def in_list(field: str, values: Iterable[Any]) -> str:
    """field in ['a', 'b']"""
    return f"{field} in [" + ", ".join(quote(v) for v in values) + "]"


def build_session() -> requests.Session:
    session = requests.Session()
    if AUTH_MODE == "netskope":
        session.headers["Netskope-Api-Token"] = API_TOKEN
    else:
        session.headers["Authorization"] = f"Bearer {API_TOKEN}"
    session.headers["Accept"] = "application/json"
    return session


def get_json(session, url: str, params: Dict[str, Any], stats: Dict[str, int]) -> Any:
    """One GET, retried on 429, 5xx, timeouts and dropped connections."""
    for attempt in range(1, MAX_RETRIES + 1):
        with _stats_lock:
            stats["calls"] += 1
        try:
            response = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
        except (requests.Timeout, requests.ConnectionError) as exc:
            wait = BACKOFF_BASE ** attempt
            log.warning("Netskope %s (attempt %d/%d) -- retry in %.0fs",
                        type(exc).__name__, attempt, MAX_RETRIES, wait)
            time.sleep(wait)
            continue
        code = response.status_code
        if code == 429 or code >= 500:
            retry_after = str(response.headers.get("Retry-After", ""))
            wait = float(retry_after) if retry_after.isdigit() else BACKOFF_BASE ** attempt
            log.warning("Netskope HTTP %d (attempt %d/%d) -- retry in %.0fs",
                        code, attempt, MAX_RETRIES, wait)
            time.sleep(wait)
            continue
        if code in (401, 403):
            raise ApiError(f"HTTP {code}: check API_TOKEN, its endpoint scopes and "
                           f"AUTH_MODE (now '{AUTH_MODE}', try the other one). "
                           f"{response.text[:300]}")
        if not response.ok:
            # 400 is nearly always the query: a field or value Netskope rejects.
            raise ApiError(f"HTTP {code} for {url}: {response.text[:500]}")
        if not response.text.strip():
            return {}
        try:
            return response.json()
        except ValueError:
            raise ApiError(f"not JSON from {url}: {response.text[:300]}")
    raise ApiError(f"gave up after {MAX_RETRIES} attempts: {url}")


def records_of(payload: Any) -> List[Dict[str, Any]]:
    """The events in a datasearch reply, {"result": [...], "status": {...}}."""
    if isinstance(payload, list):
        return payload
    status = payload.get("status")
    if isinstance(status, dict) and status.get("execution") not in (None, "SUCCESS"):
        log.warning("Netskope query %s: %s", status.get("execution"), status.get("message", ""))
    for key in ("result", "data", "events"):
        if isinstance(payload.get(key), list):
            return payload[key]
    return []


def fetch_window(session, stats: Dict[str, int], endpoint: str, query: str,
                 fields: List[str], start: int, end: int) -> List[Dict[str, Any]]:
    """Every event in start..end, PAGE_SIZE at a time."""
    url = BASE_URL.rstrip("/") + endpoint
    params = {"query": query, "starttime": start, "endtime": end, "limit": PAGE_SIZE,
              "timeout": QUERY_TIMEOUT, "fields": ",".join(fields)}
    records: List[Dict[str, Any]] = []
    offset = 0
    while True:
        page = records_of(get_json(session, url, dict(params, offset=offset), stats))
        records += page
        if len(page) < PAGE_SIZE:
            return records
        offset += PAGE_SIZE


def day_range(first: str, last: str) -> Tuple[int, int]:
    """(local midnight starting `first`, 23:59:59 ending `last`)."""
    start = datetime.strptime(first, "%Y-%m-%d")
    stop = datetime.strptime(last, "%Y-%m-%d") + timedelta(days=1)
    if stop <= start:
        raise ValueError("end date is before start date")
    return int(start.timestamp()), int(stop.timestamp()) - 1


def slices(start: int, end: int) -> List[Tuple[int, int]]:
    """start..end in TIME_WINDOW_SECONDS pieces."""
    bounds = list(range(start, end, TIME_WINDOW_SECONDS)) + [end]
    return list(zip(bounds[:-1], bounds[1:]))


def plain(value: Any) -> Any:
    """An API value as one cell: None -> "", a list -> "a, b"."""
    if value is None:
        return ""
    if isinstance(value, list):
        return ", ".join("" if v is None else str(v) for v in value)
    if isinstance(value, dict):
        return str(value)
    return value


def thread_session() -> requests.Session:
    """This thread's session; worker threads don't share one."""
    if getattr(_thread, "session", None) is None:
        _thread.session = build_session()
    return _thread.session


def fetch(session, stats: Dict[str, int], endpoint: str, fields: List[str],
          query: str, start: int, end: int) -> List[Dict[str, Any]]:
    """Every event for `query` in start..end, as {field: value} for `fields`
    only; rows identical in every field dropped (that also covers an event
    on the shared edge of two slices). No session = this thread's."""
    rows: Dict[Tuple[Any, ...], Dict[str, Any]] = {}
    for piece_start, piece_end in slices(start, end):
        for record in fetch_window(session or thread_session(), stats, endpoint, query,
                                   fields, piece_start, piece_end):
            row = {f: plain(record.get(f)) for f in fields}
            rows.setdefault(tuple(row.values()), row)
    return list(rows.values())


def fetch_parallel(stats: Dict[str, int],
                   tasks: List[Tuple[str, str, int, int]]) -> List[Dict[str, Any]]:
    """Every application event for each (label, query, start, end), WORKERS
    queries at a time; the first failure stops the rest."""
    rows: List[Dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(fetch, None, stats, EVENT_ENDPOINT, event_fields(),
                               query, start, end): label
                   for label, query, start, end in tasks}
        try:
            for future in as_completed(futures):
                got = future.result()
                rows += got
                log.info("%-16s %d events", futures[future], len(got))
        except BaseException:
            pool.shutdown(cancel_futures=True)
            raise
    return rows


def user_batches(users: List[str]) -> List[List[str]]:
    """`users` in order, split so each `user in [...]` list stays within
    USER_LIST_CHARS."""
    batches: List[List[str]] = [[]]
    size = 0
    for user in users:
        cost = len(quote(user)) + 2              # ", " between names
        if batches[-1] and size + cost > USER_LIST_CHARS:
            batches.append([])
            size = 0
        batches[-1].append(user)
        size += cost
    return [b for b in batches if b]


# --- Correlation -----------------------------------------------------------

def lower(value: Any) -> str:
    return str(value).strip().lower()


def seconds(event: Dict[str, Any]) -> int:
    return int(float(event["timestamp"] or 0))


def to_utc(value: Any) -> Any:
    """Epoch seconds -> naive UTC datetime (Excel has no time zones);
    blank -> empty cell."""
    if value in ("", None):
        return None
    try:
        return datetime.fromtimestamp(float(value), timezone.utc).replace(tzinfo=None)
    except (TypeError, ValueError):
        return value


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
    """Skope IT's Alert Type = UBA, Activity in [...]."""
    return f"alert_type eq 'uba' and {in_list('activity', activities)}"


def event_queries(users: List[str]) -> List[Tuple[str, str]]:
    """(label, query) for one batch of users: their Microsoft downloads,
    then all their uploads."""
    who = in_list("user", users)
    return [
        ("Downloads", f"{who} and {in_list('activity', DOWNLOAD_ACTIVITIES)}"
                      f" and {in_list('app', MICROSOFT_APPS)}"),
        ("Uploads", f"{who} and {in_list('activity', UPLOAD_ACTIVITIES)}"),
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


def wanted_app_activity(record: Dict[str, Any]) -> bool:
    """App_activity blank or one of APP_ACTIVITIES."""
    value = lower(record["app_activity"])
    return not value or value in {lower(a) for a in APP_ACTIVITIES}


def correlate(events: List[Dict[str, Any]], max_gap_hours: float) -> List[Dict[str, Any]]:
    """One alert per upload to a non-Microsoft app of a file the same user
    downloaded from a Microsoft app (App_activity blank or in APP_ACTIVITIES)
    at most `max_gap_hours` before it (0 = any time), naming the latest such
    download. Sorts `events` by user then time and lists each event's Alert
    IDs under "alert_ids"."""
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
        if activity in downloads and from_microsoft and wanted_app_activity(event):
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


def uba_labels(alerts: List[Dict[str, Any]]) -> Callable[[Dict[str, Any]], str]:
    """UBA activities that alerted on an event's user, direction and app --
    or user and direction, for an alert with no app. "Bulk Upload All"
    labels uploads, "Bulk Download All" downloads."""
    by_event: Dict[Tuple[str, str, str], Set[str]] = defaultdict(set)
    for alert in alerts:
        by_event[(lower(alert["user"]), direction(alert["activity"]),
                  lower(alert["app"]))].add(alert["activity"])

    def event_label(event: Dict[str, Any]) -> str:
        user, way = lower(event["user"]), direction(event["activity"])
        names = (by_event.get((user, way, lower(event["app"])), set())
                 | by_event.get((user, way, ""), set()))
        return ", ".join(sorted(names))

    return event_label


def uba_alert_names(alerts: List[Dict[str, Any]]) -> Dict[str, str]:
    """Each user's UBA alert names, "a, b" -- the alerts that put them in scope."""
    names: Dict[str, Set[str]] = defaultdict(set)
    for alert in alerts:
        if alert["alert_name"] != "":
            names[lower(alert["user"])].add(str(alert["alert_name"]))
    return {user: ", ".join(sorted(n)) for user, n in names.items()}


def file_count(alert: Dict[str, Any]) -> int:
    """The number of files a UBA alert's name reports, 0 if it names none."""
    match = FILE_COUNT.search(str(alert["alert_name"]))
    return int(match.group(1).replace(",", "")) if match else 0


def bulk_alerts(alerts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """UBA alerts of BULK_FILES+ files with a wanted App_activity, most files first."""
    return sorted((a for a in alerts if file_count(a) >= BULK_FILES and wanted_app_activity(a)),
                  key=lambda a: (-file_count(a), lower(a["user"]), seconds(a)))


# --- Excel -----------------------------------------------------------------

def sheet1_row(record: Dict[str, Any]) -> List[Any]:
    return [to_utc(record[f]) if f == "timestamp" else record[f] for f in COLUMNS]


def bulk_row(alert: Dict[str, Any]) -> List[Any]:
    cells = {"timestamp": to_utc(alert["timestamp"]), "files": file_count(alert)}
    return [cells[f] if f in cells else alert[f] for f in BULK_FIELDS]


def alert_row(found: Dict[str, Any], alert_names: Dict[str, str]) -> List[Any]:
    down, up = found["download"], found["upload"]
    return [
        found["id"], up["user"], file_name(up) or file_name(down),
        to_utc(down["timestamp"]), down["app"], down["app_activity"], down["file_path"],
        to_utc(up["timestamp"]), up["app"], up["hostname"], up["url"],
        up["srcip"], up["src_country"],
        round((seconds(up) - seconds(down)) / 3600, 2),
        alert_names.get(lower(up["user"]), ""),
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
        if count % SHEET_ROWS == 0:
            new_sheet()
        sheets[-1].append(row)
        count += 1
    if not sheets:
        new_sheet()
    last_column = get_column_letter(len(header))
    for i, sheet in enumerate(sheets):
        filled = min(SHEET_ROWS, count - i * SHEET_ROWS)
        sheet.auto_filter.ref = f"A1:{last_column}{filled + 1}"
    return count


def write_report(path: str, found: List[Dict[str, Any]], events: List[Dict[str, Any]],
                 alerts: List[Dict[str, Any]]) -> None:
    event_label = uba_labels(alerts)
    book = Workbook(write_only=True)
    alert_names = uba_alert_names(alerts)
    add_sheet(book, "Alerts", ALERT_HEADER, (alert_row(f, alert_names) for f in found))
    add_sheet(book, f"UBA {BULK_FILES}+ Files", BULK_HEADER,
              (bulk_row(a) for a in bulk_alerts(alerts)))
    add_sheet(book, "Events", EVENT_HEADER,
              (sheet1_row(e) + [e[f] for f in EVENT_COLUMNS]
               + [event_label(e), ", ".join(e.get("alert_ids", []))] for e in events))
    add_sheet(book, "UBA Alerts", UBA_HEADER,
              (sheet1_row(a) + [a[f] for f in ALERT_COLUMNS]
               for a in sorted(alerts, key=lambda a: (lower(a["user"]), seconds(a)))))
    book.save(path)


# --- Run -------------------------------------------------------------------

def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Files downloaded from a Microsoft app, then uploaded elsewhere.")
    parser.add_argument("start", nargs="?",
                        help="first day, e.g. 2026-10-01 (default: yesterday)")
    parser.add_argument("end", nargs="?",
                        help="last day, inclusive, e.g. 2026-10-27 (default: the first day)")
    parser.add_argument("--activity", action="append", metavar="NAME",
                        help="UBA alert activity that puts a user in scope, repeatable "
                             f"(default: {', '.join(UBA_ACTIVITIES)})")
    parser.add_argument("--hours", type=float, default=MAX_GAP_HOURS,
                        help="most hours from download to upload, 0 = any "
                             f"(default: {MAX_GAP_HOURS:g})")
    parser.add_argument("--out", metavar="DIR",
                        help="folder for the workbook (default: output/ beside this script)")
    args = parser.parse_args(argv)
    if args.hours < 0:
        parser.error("--hours must be 0 or more")
    args.start = args.start or (date.today() - timedelta(days=1)).isoformat()
    args.end = args.end or args.start
    return args


def main(argv: List[str]) -> int:
    args = parse_args(argv)
    setup_logging()
    load_config()
    if config_missing():
        log.error("CONFIG ERROR: no Netskope token / tenant -- set NETSKOPE_API_TOKEN and "
                  "NETSKOPE_BASE_URL (the environment, MSDU_CONFIG_FILE or config.env "
                  "beside %s)", os.path.basename(__file__))
        return 1
    try:
        start, end = day_range(args.start, args.end)
    except ValueError as exc:
        log.error("BAD DATES: %s (use: %s 2026-10-01 2026-10-27)",
                  exc, os.path.basename(__file__))
        return 1
    activities = args.activity or UBA_ACTIVITIES
    label = args.start if args.start == args.end else f"{args.start}_to_{args.end}"
    stats: Dict[str, int] = {"calls": 0}
    began = time.time()
    days = round((end + 1 - start) / 86400)
    log.info("range      %s (%d day%s)", label, days, "" if days == 1 else "s")
    log.info("UBA        %s", ", ".join(activities))
    log.info("window     upload within %s of the download",
             f"{args.hours:g} hours" if args.hours else "any time")

    try:
        alerts = fetch(None, stats, ALERT_ENDPOINT, alert_fields(),
                       uba_query(activities), start, end)
        log.info("UBA alerts %d", len(alerts))
        users = list(dict.fromkeys(str(a["user"]) for a in alerts if a["user"] != ""))
        if not users:
            log.warning("no UBA alerts with activity %s in this range -- names "
                        "must match Netskope exactly", ", ".join(activities))
            return 3
        groups = user_batches(users)
        log.info("users      %d with UBA alerts, %d per query at most, %d quer%s each "
                 "for downloads and uploads", len(users), max(map(len, groups)),
                 len(groups), "y" if len(groups) == 1 else "ies")

        # Downloads also from `hours` before the first day, so an upload
        # early on it still finds the download.
        since = start - int(args.hours * 3600)
        tasks = [(f"{kind} {n}/{len(groups)}", query,
                  since if kind == "Downloads" else start, end)
                 for n, group in enumerate(groups, 1)
                 for kind, query in event_queries(group)]
        events = fetch_parallel(stats, tasks)
    except ApiError as exc:
        log.error("NETSKOPE FAILED: %s", exc)
        return 2

    found = correlate(events, args.hours)
    out_dir = args.out or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    xlsx = os.path.join(out_dir, f"MS_Download_Upload_{label}.xlsx")
    write_report(xlsx, found, events, alerts)
    log.info("alerts     %d (%d users)", len(found),
             len({lower(f["upload"]["user"]) for f in found}))
    log.info("UBA %d+   %d alerts", BULK_FILES, len(bulk_alerts(alerts)))
    log.info("wrote      %s", xlsx)
    log.info("API calls  %d in %.0fs", stats["calls"], time.time() - began)
    return 0


def selftest() -> None:
    import json
    import tempfile
    from openpyxl import load_workbook

    global AUTH_MODE, BACKOFF_BASE, PAGE_SIZE, SHEET_ROWS, TIME_WINDOW_SECONDS
    global USER_LIST_CHARS, get_json, API_TOKEN, BASE_URL
    log.handlers[:] = [logging.NullHandler()]      # the retry tests warn on purpose
    od, sp = MICROSOFT_APPS[0], MICROSOFT_APPS[1]

    def ev(user, ts, activity, app, obj="", md5="", path="", act=""):
        row = dict.fromkeys(event_fields(), "")
        row.update(user=user, timestamp=ts, activity=activity, app=app,
                   object=obj, md5=md5, file_path=path, app_activity=act)
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

    # --- Netskope API, against a fake session -----------------------------
    class Reply:
        def __init__(self, status_code, body=None, headers=None):
            self.status_code, self.headers = status_code, headers or {}
            self.text = "" if body is None else json.dumps(body)
            self.ok = status_code < 400

        def json(self):
            return json.loads(self.text)

    class FakeSession:
        def __init__(self, *replies):
            self.replies, self.params = list(replies), []

        def get(self, url, params=None, timeout=None):
            self.params.append(params)
            reply = self.replies.pop(0)
            if isinstance(reply, Exception):
                raise reply
            return reply

    saved = AUTH_MODE, BACKOFF_BASE, PAGE_SIZE, TIME_WINDOW_SECONDS, USER_LIST_CHARS, get_json
    try:
        AUTH_MODE = "bearer"
        assert build_session().headers["Authorization"] == f"Bearer {API_TOKEN}"
        AUTH_MODE = "netskope"
        assert build_session().headers["Netskope-Api-Token"] == API_TOKEN

        BACKOFF_BASE = 0                                              # no real waits
        stats = {"calls": 0}
        session = FakeSession(Reply(429, headers={"Retry-After": "0"}),
                              requests.ConnectionError("down"), Reply(503),
                              Reply(200, {"result": [{"user": "a"}],
                                          "status": {"execution": "SUCCESS"}}))
        assert records_of(get_json(session, "u", {}, stats)) == [{"user": "a"}]
        assert stats["calls"] == 4
        for code in (401, 400):
            try:
                get_json(FakeSession(Reply(code, {"error": "x"})), "u", {}, stats)
                raise AssertionError(f"HTTP {code} accepted")
            except ApiError:
                pass
        assert records_of({}) == [] and records_of([{"a": 1}]) == [{"a": 1}]

        PAGE_SIZE = 2
        session = FakeSession(Reply(200, {"result": [{"user": "a"}, {"user": "b"}]}),
                              Reply(200, {"result": [{"user": "c"}]}))
        got = fetch_window(session, {"calls": 0}, EVENT_ENDPOINT, "q", ["user", "app"], 5, 9)
        assert [r["user"] for r in got] == ["a", "b", "c"]
        assert [p["offset"] for p in session.params] == [0, 2]
        assert session.params[0] == {"query": "q", "starttime": 5, "endtime": 9, "limit": 2,
                                     "timeout": QUERY_TIMEOUT, "fields": "user,app", "offset": 0}

        # A day is one request; rows that only differ in _id or a field not
        # asked for are one row.
        PAGE_SIZE = 5
        day = 86400
        assert slices(0, day - 1) == [(0, day - 1)]
        assert slices(-3600, day - 1) == [(-3600, day - 3600), (day - 3600, day - 1)]
        session = FakeSession(
            Reply(200, {"result": [{"_id": 1, "user": "a", "app": ["X", "Y"], "extra": 5},
                                   {"_id": 2, "user": "a", "app": ["X", "Y"]}]}),
            Reply(200, {"result": [{"_id": 3, "user": "b", "app": None}]}))
        TIME_WINDOW_SECONDS = 3600
        rows = fetch(session, {"calls": 0}, EVENT_ENDPOINT, ["user", "app", "md5"], "q", 0, 7200)
        assert rows == [{"user": "a", "app": "X, Y", "md5": ""},
                        {"user": "b", "app": "", "md5": ""}]
        assert [(p["starttime"], p["endtime"]) for p in session.params] == [(0, 3600), (3600, 7200)]
        assert slices(0, 7201) == [(0, 3600), (3600, 7200), (7200, 7201)]
        TIME_WINDOW_SECONDS = day

        # Users packed into as few queries as fit USER_LIST_CHARS, in order.
        users = [f"user{i:04d}@example.com" for i in range(500)]
        groups = user_batches(users)
        assert sum(groups, []) == users and len(groups) == 4
        assert all(len(in_list("user", g)) - len("user in []") <= USER_LIST_CHARS for g in groups)
        USER_LIST_CHARS = 25
        assert user_batches(["a@x.com", "b@x.com", "c@x.com"]) == [["a@x.com", "b@x.com"], ["c@x.com"]]
        assert user_batches([]) == []

        # Parallel queries: every task's rows; one failing query fails the run.
        def fake_get_json(session, url, params, stats):
            with _stats_lock:
                stats["calls"] += 1
            if params["query"] == "bad":
                raise ApiError("HTTP 400")
            return {"result": [] if params["offset"] else
                    [{"user": params["query"], "activity": "Upload", "timestamp": params["starttime"]}]}

        get_json = fake_get_json
        stats = {"calls": 0}
        rows = fetch_parallel(stats, [("U 1/2", "u1", 0, 10), ("U 2/2", "u2", 5, 10),
                                      ("D 1/1", "d1", -50, 10)])
        assert sorted((r["user"], r["timestamp"]) for r in rows) == [("d1", -50), ("u1", 0), ("u2", 5)]
        assert stats["calls"] == 3
        try:
            fetch_parallel(stats, [("U", "u1", 0, 10), ("X", "bad", 0, 10)])
            raise AssertionError("failed query ignored")
        except ApiError:
            pass
    finally:
        (AUTH_MODE, BACKOFF_BASE, PAGE_SIZE, TIME_WINDOW_SECONDS, USER_LIST_CHARS,
         get_json) = saved

    start, end = day_range("2026-08-01", "2026-09-30")
    assert time.localtime(start)[:6] == (2026, 8, 1, 0, 0, 0)
    assert time.localtime(end)[:6] == (2026, 9, 30, 23, 59, 59)
    start, end = day_range("2026-10-26", "2026-10-26")
    assert time.localtime(start)[:6] == (2026, 10, 26, 0, 0, 0)
    assert time.localtime(end)[:6] == (2026, 10, 26, 23, 59, 59)
    try:
        day_range("2026-09-30", "2026-08-01")
        raise AssertionError("reversed range accepted")
    except ValueError:
        pass
    assert to_utc(1784808436) == datetime(2026, 7, 23, 12, 7, 16) and to_utc("") is None

    # --- Correlation ------------------------------------------------------
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
    acts = [ev("a", 1, "Download", od, "r.pdf", act="filedownloaded"),
            ev("a", 2, "Download", od, "r.pdf", act="FileSyncDownloadedFull"),
            ev("a", 3, "Upload", "Notion", "r.pdf")]
    assert [seconds(f["download"]) for f in correlate(acts, 0)] == [1]   # not the later 2
    assert correlate(acts[1:], 0) == []

    args = parse_args(["2026-10-16", "2026-10-17"])
    assert (args.start, args.end) == ("2026-10-16", "2026-10-17")
    assert (args.hours, args.activity) == (MAX_GAP_HOURS, None) and MAX_GAP_HOURS == 24
    args = parse_args(["--hours", "8", "--activity", "Upload"])
    yesterday = (date.today() - timedelta(days=1)).isoformat()
    assert (args.start, args.end, args.hours, args.activity) == (yesterday, yesterday, 8, ["Upload"])
    args = parse_args(["2026-10-18"])
    assert (args.start, args.end, args.out) == ("2026-10-18", "2026-10-18", None)
    assert parse_args(["--out", "runs/x"]).out == "runs/x"

    # Credentials: the environment, then the config file, then the typed-in values.
    cfg_dir = tempfile.mkdtemp()
    cfg = os.path.join(cfg_dir, "config.env")
    with open(cfg, "w", encoding="utf-8") as fh:
        fh.write("# rendered\nNETSKOPE_BASE_URL=https://t.goskope.com\n"
                 "NETSKOPE_API_TOKEN='tok=en'\n\nNETSKOPE_AUTH_MODE=Netskope\n")
    assert read_config_file(cfg) == {"NETSKOPE_BASE_URL": "https://t.goskope.com",
                                     "NETSKOPE_API_TOKEN": "tok=en",
                                     "NETSKOPE_AUTH_MODE": "Netskope"}
    names = (CONFIG_FILE_ENV, "NETSKOPE_API_TOKEN", "NETSKOPE_BASE_URL", "NETSKOPE_AUTH_MODE")
    saved_cfg = API_TOKEN, BASE_URL, AUTH_MODE
    saved_env = {n: os.environ.pop(n, None) for n in names}
    try:
        os.environ[CONFIG_FILE_ENV] = cfg
        assert load_config() == cfg and not config_missing()
        assert (API_TOKEN, BASE_URL, AUTH_MODE) == ("tok=en", "https://t.goskope.com", "netskope")
        os.environ["NETSKOPE_API_TOKEN"] = "from-env"
        load_config()
        assert API_TOKEN == "from-env"                                 # environment wins
        del os.environ["NETSKOPE_API_TOKEN"]
        os.environ[CONFIG_FILE_ENV] = os.path.join(cfg_dir, "absent.env")
        API_TOKEN, BASE_URL = "PUT_YOUR_API_TOKEN_HERE", "https://YOUR_TENANT.goskope.com"
        assert load_config() == "" and config_missing()                # placeholders stay
    finally:
        API_TOKEN, BASE_URL, AUTH_MODE = saved_cfg
        for n, v in saved_env.items():
            os.environ.pop(n, None)
            if v is not None:
                os.environ[n] = v

    assert uba_query(UBA_ACTIVITIES) == (
        "alert_type eq 'uba' and activity in "
        "['Bulk Upload All', 'Bulk Download All', 'Upload', 'Download']")
    (dl_kind, dl_query), (up_kind, up_query) = event_queries(["a@x.com", "o'neil@x.com"])
    assert (dl_kind, up_kind) == ("Downloads", "Uploads")
    assert dl_query == ("user in ['a@x.com', 'o\\'neil@x.com'] and activity in ['Download'] "
                        "and app in ['Microsoft Office 365 OneDrive for Business', "
                        "'Microsoft Office 365 Sharepoint Online', 'Microsoft Teams', "
                        "'Microsoft Office 365 Outlook.com']")
    assert up_query == "user in ['a@x.com', 'o\\'neil@x.com'] and activity in ['Upload']"

    assert [direction(a) for a in ("Bulk Upload All", "Download All", "upload", "Login")] == [
        "upload", "download", "upload", "login"]

    def al(user, ts, activity, app, name="", act=""):
        row = dict.fromkeys(alert_fields(), "")
        row.update(user=user, timestamp=ts, activity=activity, app=app, alert_name=name,
                   app_activity=act)
        return row

    alerts = [al("A", 900, "Bulk Upload All", "Notion", "Upload spike"),
              al("a", 10, "Download", "", "Download spike")]
    event_label = uba_labels(alerts)
    assert event_label(ev("a", 1, "Upload", "notion")) == "Bulk Upload All"
    assert event_label(ev("a", 1, "Download", od)) == "Download"
    assert event_label(ev("a", 1, "Upload", "Box")) == ""
    assert uba_alert_names(alerts + [al("a", 20, "Upload", "Box"),
                                     al("b", 30, "Upload", "Box", "Upload spike")]) == {
        "a": "Download spike, Upload spike", "b": "Upload spike"}

    assert [file_count(al("a", 1, "Upload", "", name)) for name in (
        "Spike of 26 files downloaded from Box", "Spike of 1,250 Files uploaded",
        "1 file uploaded", "Upload spike", "")] == [26, 1250, 1, 0, 0]

    # --- Excel ------------------------------------------------------------
    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    events = sample()
    found = correlate(events, 24)
    bulk = [al("b", 50, "Upload", "Box", "Spike of 300 files uploaded to Box"),
            al("c", 40, "Bulk Upload All", "Box", "Spike of 1,250 files uploaded to Box",
               "FileDownloaded"),
            al("d", 60, "Upload", "Box", "Spike of 299 files uploaded to Box"),
            al("e", 70, "Upload", "Box", "Spike of 500 files uploaded to Box", "FileUploaded")]
    write_report(path, found, events, alerts + bulk)
    book = load_workbook(path)
    assert book.sheetnames == ["Alerts", "UBA 300+ Files", "Events", "UBA Alerts"]
    sheet = book["Alerts"]
    assert [c.value for c in sheet[1]] == ALERT_HEADER
    assert ALERT_HEADER[-2:] == ["Gap (Hours)", "UBA alert_name"] and sheet.max_column == 15
    assert sheet.auto_filter.ref == "A1:O3" and sheet.freeze_panes == "A2"
    first = dict(zip(ALERT_HEADER, [c.value for c in sheet[2]]))
    assert (first["Alert ID"], first["User"], first["File"], first["Upload App"],
            first["Download App"], first["Gap (Hours)"], first["UBA alert_name"]) == (
        "A0001", "a", "Report.XLSX", "Notion", sp, 0.03, "Download spike, Upload spike")
    assert first["Uploaded (GMT)"].year == 1970                         # an Excel date
    sheet = book["UBA 300+ Files"]
    assert [c.value for c in sheet[1]] == BULK_HEADER == [
        "Time (GMT)", "Action", "Activity", "User", "Application", "Source IP (Egress)",
        "Source Country", "Alert Name", "Severity", "Count", "Files", "App_activity"]
    assert [(r[3].value, r[7].value, r[10].value, r[11].value or "")
            for r in sheet.iter_rows(min_row=2)] == [
        ("c", "Spike of 1,250 files uploaded to Box", 1250, "FileDownloaded"),
        ("b", "Spike of 300 files uploaded to Box", 300, "")]          # most files first
    assert sheet.auto_filter.ref == "A1:L3"
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
        ("Download", "a"), ("Bulk Upload All", "A"), ("Upload", "b"),
        ("Bulk Upload All", "c"), ("Upload", "d"), ("Upload", "e")]

    saved_rows, SHEET_ROWS = SHEET_ROWS, 2
    try:
        book = Workbook(write_only=True)
        assert add_sheet(book, "T", ["x", "y"], [[1, 2], [3, 4], [5, 6]]) == 3
        assert add_sheet(book, "E", ["x"], []) == 0
        book.save(path)
        book = load_workbook(path)
        assert book.sheetnames == ["T", "T 2", "E"]
        assert [book[n].auto_filter.ref for n in book.sheetnames] == ["A1:B3", "A1:B2", "A1:A1"]
    finally:
        SHEET_ROWS = saved_rows
    print("selftest OK")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        sys.exit(0)
    sys.exit(main(sys.argv[1:]))
