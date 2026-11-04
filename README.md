# MS Download → Upload Alerts

Finds files a user downloaded from a Microsoft 365 app (OneDrive for Business,
SharePoint Online, Teams, Outlook.com) and then uploaded to any other app within
24 hours, using the Netskope REST API v2. One alert per upload, written to an
Excel workbook, with a small dashboard to run it and browse the results.

| File | What it is |
|---|---|
| `ms_download_upload.py` | The detection script. Runs on its own from a shell. |
| `server.py` | The dashboard: runs the script on demand or daily and keeps the run history. |
| `dashboard.html`, `static/` | The dashboard page. |
| `config.env.example` | Where the Netskope tenant and token go. |

## How it works

1. **Scope**: users with a UBA alert whose activity is Bulk Upload All, Bulk
   Download All, Upload or Download in the date range.
2. **Collect**: their downloads from the four Microsoft apps (starting 24 hours
   before the range) and all their uploads.
3. **Match**: same user, same file. Either the MD5 matches, or the file name
   matches ignoring case, spacing and “(1)” / “ - Copy”. Generic names like
   `image.png` match on MD5 only.
4. **Time-box**: the upload goes to a non-Microsoft app within 24 hours of the
   download (`--hours`), and the download's App_activity is FileDownloaded or blank.
5. **Report**: `MS_Download_Upload_<range>.xlsx` with the Alerts, UBA 300+ Files,
   Events and UBA Alerts sheets.

## Setup

Python 3.9+.

```bash
pip install -r requirements.txt
cp config.env.example config.env        # fill in NETSKOPE_BASE_URL and NETSKOPE_API_TOKEN
```

The token needs the `events/datasearch` alert and application endpoints. The values
are read from the environment first, then the file `MSDU_CONFIG_FILE` names, then
`config.env` beside the script. `config.env` is git-ignored.

## The dashboard

```bash
python server.py                        # http://127.0.0.1:8160/
python server.py --port 9000 --host 0.0.0.0
```

- **Run now**: pick a date range (up to 31 days), the hours window and the UBA
  activities. Only one run happens at a time.
- **Results**: counts, the alerts table (filter, sort), the UBA 300+ files alerts,
  the run log and the workbook to download.
- **History**: the last 30 runs are kept in `runs/`; older ones are deleted.

The dashboard has **no login of its own** and binds to `127.0.0.1` by default.
To share it, put it behind a reverse proxy that handles sign-in. Every URL it uses
is relative, so it works under a subpath. If the proxy passes `X-Auth-Name`,
`X-Auth-Email` or `X-Auth-User`, that name is recorded as who started a run.

## Daily run

`python server.py --run-daily` runs yesterday with the defaults, records it in the
dashboard's history and exits (non-zero if the run failed). Schedule it with cron:

```cron
0 4 * * *  cd /path/to/ms-download-upload-dashboard && python3 server.py --run-daily
```

Or with a systemd timer. The dashboard shows the next trigger of
`ms-download-upload-daily.timer`, and `MSDU_DAILY_TIMER` names another one:

```ini
# /etc/systemd/system/ms-download-upload-daily.service
[Service]
Type=oneshot
WorkingDirectory=/path/to/ms-download-upload-dashboard
ExecStart=/usr/bin/python3 server.py --run-daily

# /etc/systemd/system/ms-download-upload-daily.timer
[Timer]
OnCalendar=*-*-* 04:00:00
Persistent=true

[Install]
WantedBy=timers.target
```

## The script on its own

```bash
python ms_download_upload.py                         # yesterday
python ms_download_upload.py 2026-10-18              # that day
python ms_download_upload.py 2026-10-01 2026-10-27   # both days inclusive
python ms_download_upload.py --activity "Bulk Upload All" --activity Upload
python ms_download_upload.py --hours 8               # 0 = any gap
python ms_download_upload.py --out runs/today        # workbook there
```

Exit codes: 0 ok, 1 config or date error, 2 Netskope API failure, 3 no matching UBA alerts.

## Checks

Neither needs a tenant:

```bash
python ms_download_upload.py --selftest
python server.py --selftest
```
