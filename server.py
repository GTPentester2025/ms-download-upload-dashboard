#!/usr/bin/env python3
"""
MS Download -> Upload dashboard: a small web page for ms_download_upload.py.

It runs ms_download_upload.py -- on demand from the page (date range, hours
window, UBA activities) or once a day for yesterday (--run-daily, from cron
or a systemd timer) -- keeps the last KEEP_RUNS runs
under runs/, and shows each run's alerts, its UBA 300+ files alerts and its
log, with the workbook to download.

Stdlib + openpyxl (which the script needs anyway). Every URL the page uses is
relative, so it also works under a subpath behind a reverse proxy.
Behind a sign-in proxy that passes X-Auth-Name / X-Auth-Email / X-Auth-User,
that name is recorded as who started a run.

    python server.py [--port 8160] [--host 127.0.0.1]    the dashboard
    python server.py --run-daily                          yesterday's run
    python server.py --selftest                           offline check

One run at a time: runs/.lock names the running run and the process that owns
it. Each run is a folder runs/<id>/ holding
    run.json      parameters, status, timings, counts
    run.log       everything the script printed
    summary.json  the Alerts and UBA 300+ Files rows the page shows
    MS_Download_Upload_<range>.xlsx
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from datetime import date, datetime, timedelta, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, List, Optional
from urllib.parse import urlsplit

import ms_download_upload as msdu

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
APP_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(APP_DIR, "ms_download_upload.py")
RUNS_DIR = os.environ.get("MSDU_RUNS_DIR") or os.path.join(APP_DIR, "runs")
STATIC_DIR = os.path.join(APP_DIR, "static")

DEFAULT_PORT = 8160
KEEP_RUNS = 30            # run folders kept, newest first (a month of daily runs)
MAX_RANGE_DAYS = 31       # longest date range the page may ask for
MAX_HOURS = 168           # widest download -> upload window the page may ask for
MAX_ROWS = 2000           # rows per table sent to the page; the workbook has all
LOG_TAIL_LINES = 400      # log lines sent with a run
DAILY_WAIT_SECONDS = 3600 # how long the daily run waits for a run already going
DAILY_TIMER = os.environ.get("MSDU_DAILY_TIMER", "ms-download-upload-daily.timer")

RUN_ID = re.compile(r"^\d{8}-\d{6}-(daily|manual)(-\d+)?$")
DAY = re.compile(r"^\d{4}-\d{2}-\d{2}$")
STATIC_FILES = {
    "dashboard.css": "text/css; charset=utf-8",
    "dashboard.js": "text/javascript; charset=utf-8",
}
# The script's exit codes (see its docstring) -> run status.
STATUS_BY_EXIT = {0: "ok", 1: "failed", 2: "failed", 3: "no-uba"}

# ---------------------------------------------------------------------------

_lock = threading.Lock()   # guards starting a run within this process


class Busy(RuntimeError):
    """Another run is in progress."""


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def pid_alive(pid: Any) -> bool:
    if not isinstance(pid, int) or pid <= 0:
        return False
    if os.name == "nt":       # os.kill(pid, 0) would TERMINATE the process on Windows
        import ctypes
        kernel = ctypes.windll.kernel32
        handle = kernel.OpenProcess(0x1000, False, pid)   # PROCESS_QUERY_LIMITED_INFORMATION
        if not handle:
            return False
        code = ctypes.c_ulong()
        got = kernel.GetExitCodeProcess(handle, ctypes.byref(code))
        kernel.CloseHandle(handle)
        return bool(got) and code.value == 259            # STILL_ACTIVE
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def read_json(path: str, default: Any = None) -> Any:
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def write_json(path: str, data: Any) -> None:
    """Atomic: a reader never sees half a file."""
    tmp = f"{path}.{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=1, default=str)
    os.replace(tmp, path)


# --- Run storage -----------------------------------------------------------

def run_dir(run_id: str) -> str:
    if not RUN_ID.match(run_id or ""):
        raise ValueError(f"bad run id {run_id!r}")
    return os.path.join(RUNS_DIR, run_id)


def load_run(run_id: str) -> Optional[Dict[str, Any]]:
    try:
        return read_json(os.path.join(run_dir(run_id), "run.json"))
    except ValueError:
        return None


def save_run(run: Dict[str, Any]) -> None:
    write_json(os.path.join(run_dir(run["id"]), "run.json"), run)


def list_runs() -> List[Dict[str, Any]]:
    """Every kept run, newest first."""
    try:
        names = sorted((n for n in os.listdir(RUNS_DIR) if RUN_ID.match(n)), reverse=True)
    except FileNotFoundError:
        return []
    runs = [load_run(n) for n in names]
    return [r for r in runs if r]


def new_run_id(trigger: str) -> str:
    stem = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + trigger
    run_id, n = stem, 1
    while os.path.exists(os.path.join(RUNS_DIR, run_id)):
        n += 1
        run_id = f"{stem}-{n}"
    return run_id


def prune_runs(keep: Optional[int] = None) -> None:
    """Delete all but the newest `keep` (KEEP_RUNS) run folders; never a running one."""
    for run in list_runs()[KEEP_RUNS if keep is None else keep:]:
        if run.get("status") != "running":
            shutil.rmtree(run_dir(run["id"]), ignore_errors=True)


# --- The lock --------------------------------------------------------------

def lock_path() -> str:
    return os.path.join(RUNS_DIR, ".lock")


def current_lock() -> Optional[Dict[str, Any]]:
    """The live lock, or None. A lock whose owner process is gone is removed."""
    held = read_json(lock_path())
    if held is None:
        return None
    if pid_alive(held.get("owner")):
        return held
    try:
        os.remove(lock_path())
    except FileNotFoundError:
        pass
    return None


def take_lock(run_id: str) -> None:
    os.makedirs(RUNS_DIR, exist_ok=True)
    for _ in range(2):
        try:
            fd = os.open(lock_path(), os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
        except FileExistsError:
            held = current_lock()
            if held:
                raise Busy(f"run {held.get('run')} is in progress")
            continue          # the stale lock was just removed: try again
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump({"run": run_id, "owner": os.getpid(), "since": now_iso()}, fh)
        return
    raise Busy("could not take the run lock")


def release_lock(run_id: str) -> None:
    held = read_json(lock_path())
    if held and held.get("run") == run_id:
        try:
            os.remove(lock_path())
        except FileNotFoundError:
            pass


def recover_interrupted() -> None:
    """A run still marked running whose processes are gone -- the service was
    restarted mid-run -- is marked interrupted, so the page stops waiting."""
    held = current_lock()
    for run in list_runs():
        if run.get("status") != "running":
            continue
        if held and held.get("run") == run["id"]:
            continue
        if pid_alive(run.get("pid")):
            continue
        run.update(status="interrupted", finished_at=run.get("finished_at") or now_iso(),
                   error="the dashboard was restarted while this run was going")
        save_run(run)


# --- Validation ------------------------------------------------------------

def validate_params(body: Dict[str, Any], today: Optional[date] = None) -> Dict[str, Any]:
    """The run parameters from the page, checked. Raises ValueError with a
    message meant for the page."""
    today = today or date.today()
    start, end = str(body.get("start", "")), str(body.get("end", "") or body.get("start", ""))
    for value in (start, end):
        if not DAY.match(value):
            raise ValueError("Dates must be YYYY-MM-DD.")
    try:
        first, last = date.fromisoformat(start), date.fromisoformat(end)
    except ValueError:
        raise ValueError("That is not a real date.")
    if last < first:
        raise ValueError("The end date is before the start date.")
    if last > today:
        raise ValueError("The end date is in the future.")
    if (last - first).days + 1 > MAX_RANGE_DAYS:
        raise ValueError(f"Pick at most {MAX_RANGE_DAYS} days.")

    try:
        hours = float(body.get("hours", msdu.MAX_GAP_HOURS))
    except (TypeError, ValueError):
        raise ValueError("Hours must be a number.")
    if not 0 <= hours <= MAX_HOURS:
        raise ValueError(f"Hours must be between 0 and {MAX_HOURS}.")

    activities = body.get("activities") or list(msdu.UBA_ACTIVITIES)
    if not isinstance(activities, list) or not all(isinstance(a, str) for a in activities):
        raise ValueError("Activities must be a list of names.")
    unknown = [a for a in activities if a not in msdu.UBA_ACTIVITIES]
    if unknown:
        raise ValueError(f"Unknown UBA activity: {', '.join(unknown)}.")
    return {"start": start, "end": end, "hours": hours,
            "activities": [a for a in msdu.UBA_ACTIVITIES if a in activities]}


# --- Running ---------------------------------------------------------------

def script_command(run: Dict[str, Any]) -> List[str]:
    cmd = [sys.executable, SCRIPT, run["start"], run["end"], "--hours", f"{run['hours']:g}",
           "--out", run_dir(run["id"])]
    for activity in run["activities"]:
        cmd += ["--activity", activity]
    return cmd


def start_run(params: Dict[str, Any], trigger: str, by: str) -> Dict[str, Any]:
    """Create the run, take the lock and start the script. Returns the run
    and the Popen; the caller waits on it and calls finish_run."""
    with _lock:
        run_id = new_run_id(trigger)
        take_lock(run_id)
        try:
            os.makedirs(run_dir(run_id))
            run = dict(params, id=run_id, trigger=trigger, by=by, status="running",
                       started_at=now_iso(), finished_at=None, duration_s=None,
                       exit_code=None, pid=None, workbook=None, counts=None, error=None)
            log = open(os.path.join(run_dir(run_id), "run.log"), "wb")
            try:
                proc = subprocess.Popen(script_command(run), cwd=APP_DIR, stdout=log,
                                        stderr=subprocess.STDOUT,
                                        env=dict(os.environ, PYTHONUNBUFFERED="1",
                                                 PYTHONIOENCODING="utf-8"))
            finally:
                log.close()
            run["pid"] = proc.pid
            save_run(run)
        except BaseException:
            release_lock(run_id)
            raise
    return {"run": run, "proc": proc}


def finish_run(run: Dict[str, Any], exit_code: int) -> Dict[str, Any]:
    """Record how the run ended, read its workbook, prune, release the lock."""
    folder = run_dir(run["id"])
    try:
        run["exit_code"] = exit_code
        run["status"] = STATUS_BY_EXIT.get(exit_code, "failed")
        run["finished_at"] = now_iso()
        started = datetime.fromisoformat(run["started_at"].replace("Z", "+00:00"))
        run["duration_s"] = round((datetime.now(timezone.utc) - started).total_seconds(), 1)
        books = sorted(n for n in os.listdir(folder) if n.endswith(".xlsx"))
        run["workbook"] = books[-1] if books else None
        if run["workbook"]:
            try:
                summary, counts = read_workbook(os.path.join(folder, run["workbook"]))
                write_json(os.path.join(folder, "summary.json"), summary)
                run["counts"] = counts
            except Exception as exc:          # a broken workbook must not wedge the run
                run["error"] = f"could not read the workbook: {exc}"
        if run["status"] == "failed" and not run.get("error"):
            run["error"] = last_error(os.path.join(folder, "run.log"))
        save_run(run)
    finally:
        release_lock(run["id"])
        prune_runs()
    return run


def run_in_background(params: Dict[str, Any], by: str) -> Dict[str, Any]:
    started = start_run(params, "manual", by)
    run, proc = started["run"], started["proc"]
    threading.Thread(target=lambda: finish_run(run, proc.wait()), daemon=True,
                     name=f"run-{run['id']}").start()
    return run


def run_daily() -> int:
    """Yesterday, with the script's defaults -- what the timer runs. Waits up
    to DAILY_WAIT_SECONDS for a run already in progress."""
    yesterday = (date.today() - timedelta(days=1)).isoformat()
    params = {"start": yesterday, "end": yesterday, "hours": float(msdu.MAX_GAP_HOURS),
              "activities": list(msdu.UBA_ACTIVITIES)}
    recover_interrupted()
    deadline = time.monotonic() + DAILY_WAIT_SECONDS
    while True:
        try:
            started = start_run(params, "daily", "daily timer")
            break
        except Busy as exc:
            if time.monotonic() > deadline:
                print(f"daily run skipped: {exc} for over {DAILY_WAIT_SECONDS}s", file=sys.stderr)
                return 1
            print(f"waiting: {exc}", file=sys.stderr)
            time.sleep(30)
    run = finish_run(started["run"], started["proc"].wait())
    print(f"daily run {run['id']}: {run['status']} {json.dumps(run.get('counts'))}")
    return 0 if run["status"] in ("ok", "no-uba") else 1


# --- Reading results -------------------------------------------------------

def cell(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return "" if value is None else value


def read_workbook(path: str):
    """(summary for the page, counts) from a finished run's workbook. Sheets
    past Excel's row limit continue on "<name> 2", ... and are read too."""
    from openpyxl import load_workbook
    book = load_workbook(path, read_only=True)
    try:
        def sheet_rows(title: str):
            header, rows, total = [], [], 0
            for name in book.sheetnames:
                if name != title and not re.fullmatch(re.escape(title) + r" \d+", name):
                    continue
                for i, row in enumerate(book[name].iter_rows(values_only=True)):
                    if i == 0:
                        header = header or [cell(v) for v in row]
                        continue
                    total += 1
                    if len(rows) < MAX_ROWS:
                        rows.append([cell(v) for v in row])
            return header, rows, total

        alerts_h, alerts, n_alerts = sheet_rows("Alerts")
        bulk_title = f"UBA {msdu.BULK_FILES}+ Files"
        bulk_h, bulk, n_bulk = sheet_rows(bulk_title)
        _, _, n_events = sheet_rows("Events")
        _, _, n_uba = sheet_rows("UBA Alerts")
        user_col = alerts_h.index("User") if "User" in alerts_h else None
        users = len({str(r[user_col]).lower() for r in alerts if user_col is not None})
    finally:
        book.close()
    summary = {
        "alerts": {"header": alerts_h, "rows": alerts, "total": n_alerts},
        "bulk": {"title": bulk_title, "header": bulk_h, "rows": bulk, "total": n_bulk},
    }
    counts = {"alerts": n_alerts, "users": users, "users_partial": n_alerts > len(alerts),
              "bulk": n_bulk, "events": n_events, "uba_alerts": n_uba}
    return summary, counts


def last_error(log_file: str) -> Optional[str]:
    """The script's last ERROR line, else the log's last line."""
    try:
        with open(log_file, encoding="utf-8", errors="replace") as fh:
            lines = [ln.strip() for ln in fh if ln.strip()]
    except OSError:
        return None
    errors = [ln for ln in lines if "ERROR" in ln or "FAILED" in ln or "Traceback" in ln]
    picked = (errors or lines or [None])[-1]
    return mask(picked)[:500] if picked else None


def secrets_to_mask() -> List[str]:
    msdu.load_config()
    return [s for s in (msdu.API_TOKEN,) if s and len(s) >= 6 and not msdu.config_missing()]


def mask(text: str) -> str:
    for secret in secrets_to_mask():
        text = text.replace(secret, "********")
    return text


def read_log(run_id: str, tail: Optional[int] = None) -> str:
    try:
        with open(os.path.join(run_dir(run_id), "run.log"), encoding="utf-8",
                  errors="replace") as fh:
            lines = fh.readlines()
    except OSError:
        return ""
    if tail is not None:
        lines = lines[-tail:]
    return mask("".join(lines))


def config_status() -> Dict[str, Any]:
    msdu.load_config()
    missing = msdu.config_missing()
    host = urlsplit(msdu.BASE_URL).hostname if not missing else None
    return {"ready": not missing, "tenant": host, "auth_mode": msdu.AUTH_MODE}


def schedule_status() -> Dict[str, Any]:
    """The daily timer's next and last trigger, from systemd when there is one."""
    out: Dict[str, Any] = {"timer": DAILY_TIMER, "next": None, "last": None, "active": None}
    if not shutil.which("systemctl"):
        return out
    try:
        r = subprocess.run(["systemctl", "show", DAILY_TIMER, "-p", "NextElapseUSecRealtime",
                            "-p", "LastTriggerUSec", "-p", "ActiveState"],
                           capture_output=True, text=True, timeout=5)
    except (OSError, subprocess.SubprocessError):
        return out
    values = dict(line.split("=", 1) for line in r.stdout.splitlines() if "=" in line)
    for key, name in (("next", "NextElapseUSecRealtime"), ("last", "LastTriggerUSec")):
        value = values.get(name, "").strip()
        out[key] = value if value and value != "n/a" else None
    out["active"] = values.get("ActiveState") == "active"
    return out


def state() -> Dict[str, Any]:
    runs = list_runs()
    held = current_lock()
    return {
        "running": next((r for r in runs if held and r["id"] == held.get("run")), None),
        "config": config_status(),
        "schedule": schedule_status(),
        "defaults": {"hours": msdu.MAX_GAP_HOURS, "activities": msdu.UBA_ACTIVITIES,
                     "microsoft_apps": msdu.MICROSOFT_APPS, "bulk_files": msdu.BULK_FILES,
                     "max_range_days": MAX_RANGE_DAYS, "max_hours": MAX_HOURS,
                     "keep_runs": KEEP_RUNS, "yesterday": (date.today() - timedelta(days=1)).isoformat()},
        "runs": runs,
    }


def run_detail(run_id: str) -> Optional[Dict[str, Any]]:
    run = load_run(run_id)
    if not run:
        return None
    summary = read_json(os.path.join(run_dir(run_id), "summary.json"), {})
    return {"run": run, "summary": summary, "log_tail": read_log(run_id, LOG_TAIL_LINES),
            "max_rows": MAX_ROWS}


# --- HTTP ------------------------------------------------------------------

PAGE_CSP = ("default-src 'self'; script-src 'self'; "
            "style-src 'self' https://fonts.googleapis.com; font-src https://fonts.gstatic.com; "
            "img-src 'self' data:; connect-src 'self'; frame-ancestors 'self'; base-uri 'none'; "
            "form-action 'none'")


class Handler(BaseHTTPRequestHandler):
    server_version = "ms-download-upload"
    sys_version = ""

    # -- plumbing
    def send_body(self, status: int, body: bytes, ctype: str,
                  extra: Optional[Dict[str, str]] = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "same-origin")
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def send_json(self, status: int, data: Any) -> None:
        self.send_body(status, json.dumps(data, default=str).encode("utf-8"),
                       "application/json; charset=utf-8")

    def send_error_json(self, status: int, message: str) -> None:
        self.send_json(status, {"error": message})

    def log_message(self, fmt: str, *args: Any) -> None:
        sys.stderr.write(f"{self.address_string()} {fmt % args}\n")

    def log_request(self, code: Any = "-", size: Any = "-") -> None:
        # The page polls; only POSTs and failures are worth a journal line.
        if self.command == "GET" and str(code)[:1] in ("2", "3"):
            return
        super().log_request(code, size)

    def who(self) -> str:
        for header in ("X-Auth-Name", "X-Auth-Email", "X-Auth-User"):
            value = (self.headers.get(header) or "").strip()
            if value:
                return value[:120]
        return "web user"

    # -- routes
    def do_HEAD(self) -> None:
        self.do_GET()

    def do_GET(self) -> None:
        path = urlsplit(self.path).path
        try:
            if path in ("/", "/index.html"):
                with open(os.path.join(APP_DIR, "dashboard.html"), "rb") as fh:
                    self.send_body(200, fh.read(), "text/html; charset=utf-8",
                                   {"Content-Security-Policy": PAGE_CSP})
            elif path.startswith("/static/") and path[8:] in STATIC_FILES:
                with open(os.path.join(STATIC_DIR, path[8:]), "rb") as fh:
                    self.send_body(200, fh.read(), STATIC_FILES[path[8:]])
            elif path == "/healthz":
                self.send_body(200, b"ok\n", "text/plain; charset=utf-8")
            elif path == "/api/state":
                self.send_json(200, state())
            elif m := re.fullmatch(r"/api/runs/([^/]+)", path):
                detail = run_detail(m.group(1)) if RUN_ID.match(m.group(1)) else None
                if detail:
                    self.send_json(200, detail)
                else:
                    self.send_error_json(404, "No such run.")
            elif m := re.fullmatch(r"/api/runs/([^/]+)/log", path):
                if RUN_ID.match(m.group(1)) and load_run(m.group(1)):
                    self.send_body(200, read_log(m.group(1)).encode("utf-8"),
                                   "text/plain; charset=utf-8")
                else:
                    self.send_error_json(404, "No such run.")
            elif m := re.fullmatch(r"/api/runs/([^/]+)/download", path):
                self.send_workbook(m.group(1))
            else:
                self.send_error_json(404, "Not found.")
        except BrokenPipeError:
            pass
        except Exception as exc:                  # noqa: BLE001 -- never kill the server
            self.log_message("error on %s: %r", path, exc)
            self.send_error_json(500, "Server error -- see the service log.")

    def send_workbook(self, run_id: str) -> None:
        run = load_run(run_id) if RUN_ID.match(run_id) else None
        name = run and run.get("workbook")
        if not name:
            self.send_error_json(404, "This run has no workbook.")
            return
        with open(os.path.join(run_dir(run_id), os.path.basename(name)), "rb") as fh:
            body = fh.read()
        self.send_body(200, body,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       {"Content-Disposition": f'attachment; filename="{os.path.basename(name)}"'})

    def same_origin(self) -> bool:
        """A cross-site page cannot start runs: browsers send Origin on a POST,
        and it must name this host."""
        origin = self.headers.get("Origin")
        if not origin:
            return True
        host = self.headers.get("X-Forwarded-Host") or self.headers.get("Host") or ""
        host = host.split(",")[0].strip()
        # Hostnames only: nginx passes $host, which drops a non-default port.
        return (urlsplit(origin).hostname or "") == (urlsplit("//" + host).hostname or "")

    def do_POST(self) -> None:
        path = urlsplit(self.path).path
        if path != "/api/runs":
            self.send_error_json(404, "Not found.")
            return
        if not (self.headers.get("Content-Type") or "").startswith("application/json"):
            self.send_error_json(415, "Send JSON.")
            return
        if not self.same_origin():
            self.send_error_json(403, "Cross-site request refused.")
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = -1
        if not 0 < length <= 10_000:
            self.send_error_json(400, "Bad request body.")
            return
        try:
            body = json.loads(self.rfile.read(length))
            if not isinstance(body, dict):
                raise ValueError("Bad request body.")
            params = validate_params(body)
        except ValueError as exc:
            self.send_error_json(400, str(exc) if str(exc) else "Bad request body.")
            return
        if not config_status()["ready"]:
            self.send_error_json(409, "No Netskope token yet: set NETSKOPE_API_TOKEN and "
                                      "NETSKOPE_BASE_URL (config.env or the environment).")
            return
        try:
            run = run_in_background(params, self.who())
        except Busy as exc:
            self.send_error_json(409, f"A run is already going ({exc}).")
            return
        self.send_json(202, {"run": run})


def serve(host: str, port: int) -> None:
    os.makedirs(RUNS_DIR, exist_ok=True)
    recover_interrupted()
    httpd = ThreadingHTTPServer((host, port), Handler)
    httpd.daemon_threads = True
    print(f"MS Download -> Upload dashboard on http://{host}:{port}/  (runs: {RUNS_DIR})",
          flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()


# --- Self-test -------------------------------------------------------------

def selftest() -> None:
    import tempfile
    import urllib.error
    import urllib.request
    from openpyxl import Workbook

    global RUNS_DIR, SCRIPT, KEEP_RUNS
    saved = RUNS_DIR, SCRIPT, KEEP_RUNS
    tmp = tempfile.mkdtemp()
    RUNS_DIR = os.path.join(tmp, "runs")

    # A stand-in for the script: writes a workbook like the real one and exits
    # with the code its first date asks for (2026-01-0N -> exit N - 1).
    fake = os.path.join(tmp, "fake_script.py")
    with open(fake, "w", encoding="utf-8") as fh:
        fh.write(
            "import os, sys\n"
            "from openpyxl import Workbook\n"
            "args = sys.argv[1:]\n"
            "out = args[args.index('--out') + 1]\n"
            "code = int(args[0][-1]) - 1\n"
            "print('range', args[0], 'activities', args.count('--activity'))\n"
            "if code == 1: print('CONFIG ERROR: no token'); sys.exit(1)\n"
            "if code == 3: print('no UBA alerts'); sys.exit(3)\n"
            "book = Workbook(write_only=True)\n"
            "a = book.create_sheet('Alerts'); a.append(['Alert ID', 'User', 'File'])\n"
            "for i in range(3): a.append([f'A{i}', 'u1' if i else 'U2', 'f.pdf'])\n"
            "a2 = book.create_sheet('Alerts 2'); a2.append(['Alert ID', 'User', 'File'])\n"
            "a2.append(['A9', 'u3', 'g.pdf'])\n"
            "b = book.create_sheet('UBA 300+ Files'); b.append(['User', 'Files'])\n"
            "b.append(['u1', 312])\n"
            "e = book.create_sheet('Events'); e.append(['x'])\n"
            "for i in range(5): e.append([i])\n"
            "u = book.create_sheet('UBA Alerts'); u.append(['x']); u.append([1]); u.append([2])\n"
            "book.save(os.path.join(out, 'MS_Download_Upload_' + args[0] + '.xlsx'))\n"
            "sys.exit(code)\n")
    SCRIPT = fake
    try:
        # Validation.
        today = date(2026, 11, 4)
        ok = validate_params({"start": "2026-11-01", "end": "2026-11-03", "hours": "8",
                              "activities": ["Upload", "Bulk Upload All"]}, today)
        assert ok == {"start": "2026-11-01", "end": "2026-11-03", "hours": 8.0,
                      "activities": ["Bulk Upload All", "Upload"]}     # script's order
        assert validate_params({"start": "2026-11-03"}, today)["activities"] == msdu.UBA_ACTIVITIES
        for bad in ({"start": "3/11/2026"}, {"start": "2026-11-03", "end": "2026-11-01"},
                    {"start": "2026-11-05"}, {"start": "2026-09-01", "end": "2026-11-03"},
                    {"start": "2026-11-03", "hours": 500}, {"start": "2026-11-03", "hours": "x"},
                    {"start": "2026-11-03", "activities": ["Login"]},
                    {"start": "2026-02-30"}):
            try:
                validate_params(bad, today)
                raise AssertionError(f"accepted {bad}")
            except ValueError:
                pass
        for bad_id in ("../etc", "20261104-030000-x", "", "20261104-030000-daily/../x"):
            assert not RUN_ID.match(bad_id)

        # A run end to end: lock, workbook, counts, summary, log.
        params = validate_params({"start": "2026-01-01", "hours": 24}, today)
        started = start_run(params, "manual", "tester")
        run = started["run"]
        assert current_lock()["run"] == run["id"]
        try:
            start_run(params, "manual", "tester")
            raise AssertionError("second run started while one was going")
        except Busy:
            pass
        run = finish_run(run, started["proc"].wait())
        assert run["status"] == "ok" and run["workbook"] == "MS_Download_Upload_2026-01-01.xlsx"
        assert run["counts"] == {"alerts": 4, "users": 3, "users_partial": False, "bulk": 1,
                                 "events": 5, "uba_alerts": 2}, run["counts"]
        assert current_lock() is None
        detail = run_detail(run["id"])
        assert detail["summary"]["alerts"]["header"] == ["Alert ID", "User", "File"]
        assert [r[0] for r in detail["summary"]["alerts"]["rows"]] == ["A0", "A1", "A2", "A9"]
        assert detail["summary"]["bulk"]["rows"] == [["u1", 312]]
        assert "activities 4" in detail["log_tail"]

        # Exit codes -> status, with the error line kept.
        for day, status in (("2026-01-02", "failed"), ("2026-01-04", "no-uba")):
            p = dict(params, start=day, end=day)
            s = start_run(p, "manual", "tester")
            r = finish_run(s["run"], s["proc"].wait())
            assert r["status"] == status, (day, r)
            assert r["workbook"] is None
        assert "CONFIG ERROR" in list_runs()[1]["error"]

        # A run left "running" by a restart is marked interrupted.
        ghost = dict(run, id="20200101-000000-manual", status="running", pid=999999999)
        os.makedirs(run_dir(ghost["id"]))
        save_run(ghost)
        recover_interrupted()
        assert load_run(ghost["id"])["status"] == "interrupted"
        # A stale lock (owner gone) does not block the next run.
        with open(lock_path(), "w", encoding="utf-8") as fh:
            json.dump({"run": "x", "owner": 999999999}, fh)
        assert current_lock() is None and not os.path.exists(lock_path())

        # Retention: newest KEEP_RUNS kept.
        KEEP_RUNS = 2
        prune_runs(KEEP_RUNS)
        assert len(list_runs()) == 2

        # The HTTP side, on a free port.
        httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        port = httpd.server_address[1]
        threading.Thread(target=httpd.serve_forever, daemon=True).start()
        base = f"http://127.0.0.1:{port}"

        def call(path, data=None, headers=None):
            req = urllib.request.Request(base + path, data=data, headers=headers or {},
                                         method="POST" if data is not None else "GET")
            try:
                with urllib.request.urlopen(req, timeout=10) as resp:
                    return resp.status, resp.headers, resp.read()
            except urllib.error.HTTPError as err:
                return err.code, err.headers, err.read()

        status, headers, body = call("/")
        assert status == 200 and b"dashboard.js" in body
        assert "script-src 'self'" in headers["Content-Security-Policy"]
        assert call("/static/dashboard.js")[0] == 200
        assert call("/static/../server.py")[0] == 404
        status, _, body = call("/api/state")
        st = json.loads(body)
        assert status == 200 and st["running"] is None and len(st["runs"]) == 2
        assert st["defaults"]["activities"] == msdu.UBA_ACTIVITIES
        latest = st["runs"][0]["id"]
        assert call(f"/api/runs/{latest}")[0] == 200
        assert call("/api/runs/..%2f..%2fetc")[0] == 404
        assert call("/api/runs/20200101-000000-manual/download")[0] == 404
        jhdr = {"Content-Type": "application/json"}
        assert call("/api/runs", b"start=x", {"Content-Type": "text/plain"})[0] == 415
        assert call("/api/runs", b'{"start":"2026-01-01"}',
                    dict(jhdr, Origin="https://evil.example"))[0] == 403
        assert call("/api/runs", b'{"start":"nope"}', jhdr)[0] == 400
        httpd.shutdown()
        httpd.server_close()
    finally:
        RUNS_DIR, SCRIPT, KEEP_RUNS = saved
        shutil.rmtree(tmp, ignore_errors=True)
    print("selftest OK")


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description="MS Download -> Upload dashboard")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--run-daily", action="store_true",
                        help="run yesterday with the defaults, then exit (the daily timer)")
    parser.add_argument("--selftest", action="store_true", help="offline check, then exit")
    args = parser.parse_args(argv)
    if args.selftest:
        selftest()
        return 0
    if args.run_daily:
        return run_daily()
    serve(args.host, args.port)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
