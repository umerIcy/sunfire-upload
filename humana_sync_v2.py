"""
Humana call recordings: Google Sheet -> CallGrid -> Sunfire, then external API.
"""

import argparse
import csv
import glob
import json
import logging
import os
import re
import sys
import time
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote

import openpyxl
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(HERE, "config.json")
ENV_FILE = os.path.join(HERE, ".env")
LOG_FILE = os.path.join(HERE, "sent_log_v2.csv")
PENDING_FILE = os.path.join(HERE, "pending_v2.json")
LOCK_FILE = os.path.join(HERE, "humana_sync_v2.lock")
REC_DIR = os.path.join(HERE, "recordings")
LOG_COLUMNS = ["time", "enrollment_code", "client", "caller_id", "sale_date",
               "files", "callRecordingId", "status", "error", "method"]
DONE_STATUSES = {"finalized", "already_finalized", "already_exists"}
log = logging.getLogger("humana_sync_v2")
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
]


FIELD_CANDIDATES = {
    "id": ["id", "callId", "call_id", "uuid", "_id"],
    "caller": ["CallerId", "callerId", "caller_id", "callerNumber", "from", "ani", "phone"],
    "recording": ["callRecordingUrl", "recordingUrl", "recording_url", "recording"],
    "start": ["createdAt", "startTime", "start_time", "startedAt", "date", "timestamp"],
    "end": ["endTime", "end_time", "endedAt", "callEnd", "end"],
    "duration": ["callDuration", "duration", "durationSeconds", "totalDuration"],
}

AUDIO_EXT = {
    "audio/mpeg": ".mp3", "audio/mp3": ".mp3", "audio/wav": ".wav",
    "audio/x-wav": ".wav", "audio/wave": ".wav", "audio/mp4": ".m4a",
    "audio/x-m4a": ".m4a", "audio/ogg": ".ogg", "audio/webm": ".webm",
}
CONTENT_TYPE = {
    ".mp3": "audio/mpeg", ".wav": "audio/wav", ".m4a": "audio/mp4",
    ".ogg": "audio/ogg", ".webm": "audio/webm",
}
SAMPLE_FIELDS = {
    "agent_name": "Sample Agent",
    "agent_email": "agent@example.com",
    "agent_npn": "12345678",
    "team_name": "Sales Team A",
    "enrollee_name": "Sample Enrollee",
    "enrollee_phone": "+15555550100",
    "enrollee_id": "ENR-00000",
    "zip_code": "90210",
    "policy_name": "Sample Plan",
    "policy_number": "H0000-000-000",
    "affiliate_key": "sample-affiliate",
}
OPTIONAL_COLUMNS = [
    ("agent_name", "agent_column", ["agent name"]),
    ("agent_email", "agent_email_column", ["agent email"]),
    ("agent_npn", "agent_npn_column", ["agent npn", "npn"]),
    ("team_name", "team_column", ["team name"]),
    ("zip_code", "zip_column", ["zip code", "zip"]),
    ("policy_name", "policy_column", ["policy name", "plan name"]),
    ("policy_number", "policy_number_column", ["policy number"]),
]


# start the console log
def setup_logging():
    log.setLevel(logging.INFO)
    log.handlers.clear()
    log.propagate = False
    fmt = logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S")
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    log.addHandler(console)


# turn a byte count into kb or mb
def fmt_bytes(n):
    n = float(n or 0)
    if n >= 1_000_000:
        return f"{n / 1_000_000:.1f} MB"
    if n >= 1000:
        return f"{n / 1000:.0f} KB"
    return f"{int(n)} B"


# build the sent-row id from code, caller, and date
def row_key(row):
    sale = row["date"].isoformat() if row.get("date") else ""
    return f"{row.get('code') or ''}|{row.get('caller') or ''}|{sale}"


# --------------------------------------------------------------------------- #
# Config / helpers
# --------------------------------------------------------------------------- #
# read key=value lines from .env
def load_env():
    if not os.path.exists(ENV_FILE):
        sys.exit(".env not found next to this script.")
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            env[key.strip()] = value.strip()
    return env


# return one .env value or stop if it is missing
def env_value(env, key):
    value = (env.get(key) or "").strip()
    if not value:
        sys.exit(f".env is missing {key}")
    return value


# load config.json and fill secrets from .env
def load_config():
    if not os.path.exists(CONFIG_FILE):
        sys.exit("config.json not found next to this script.")
    with open(CONFIG_FILE, encoding="utf-8-sig") as f:
        cfg = json.load(f)
    if "FILL" in json.dumps(cfg):
        sys.exit("config.json still has FILL placeholders - fill them in first.")
    env = load_env()
    cfg.setdefault("sunfire", {})
    cfg["sunfire"]["username"] = env_value(env, "SUNFIRE_USERNAME")
    cfg["sunfire"]["password"] = env_value(env, "SUNFIRE_PASSWORD")
    cfg.setdefault("callgrid", {})
    cfg["callgrid"]["api_key"] = env_value(env, "CALLGRID_API_KEY")
    cfg.setdefault("ix", {})
    cfg["ix"]["organization_id"] = env_value(env, "IX_ORGANIZATION_ID")
    cfg["ix"]["client_id"] = env_value(env, "IX_CLIENT_ID")
    cfg["ix"]["client_secret"] = env_value(env, "IX_CLIENT_SECRET")
    try:
        cfg["google_service_account"] = json.loads(env_value(env, "GOOGLE_SERVICE_ACCOUNT_JSON"))
    except json.JSONDecodeError as e:
        sys.exit(f".env GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON: {e}")
    return cfg


# keep the last 10 digits of a phone number
def phone10(value):
    """Normalise any phone format (17578794589.0, '(757) 879-4589', '+1757...') to 10 digits."""
    if value is None:
        return ""
    if isinstance(value, float):
        value = int(value)
    digits = re.sub(r"\D", "", str(value))
    return digits[-10:] if len(digits) >= 10 else ""


# turn a sheet cell into a date
def as_date(value):
    """date, datetime, Google/Excel serial, 2026-10-17, or 10/17/2026 -> date, or None."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, bool) or value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return serial_to_date(value)
    text = str(value).strip()
    if re.fullmatch(r"\d+(\.0+)?", text):
        return serial_to_date(float(text))
    head = text.split("T", 1)[0].split(" ", 1)[0]
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", head)
    if m:
        year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", head)
        if not m:
            return None
        month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(year, month, day)
    except ValueError:
        return None


# turn a google or excel day number into a date
def serial_to_date(value):
    """Google/Excel day count since 1899-12-30. The fraction is the time of day."""
    serial = float(value)
    if serial <= 0 or serial > 100000:
        return None
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(serial))).date()
    except (OverflowError, ValueError):
        return None


# turn a cell into cleaned text
def cell_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value).split())


# turn an enrollment code cell into text
def as_code(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


# turn a callgrid time into utc
def to_utc(value):
    """ISO string or epoch (s or ms) -> aware UTC datetime, or None."""
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)) or (isinstance(value, str) and value.isdigit()):
        n = float(value)
        return datetime.fromtimestamp(n / 1000 if n > 1e11 else n, tz=timezone.utc)
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    return (dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)).astimezone(timezone.utc)


# format a datetime as utc text
def iso_z(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


# format a datetime for the log
def stamp(dt):
    return dt.isoformat(timespec="seconds")


# load rows that were already sent
def read_log():
    """Sent rows, keyed by enrollment code + caller + sale date."""
    keys = set()
    if not os.path.exists(LOG_FILE):
        return keys
    with open(LOG_FILE, newline="", encoding="utf-8") as f:
        rows = csv.DictReader(f)
        if not rows.fieldnames or "status" not in rows.fieldnames:
            return keys
        for r in rows:
            if r.get("status") not in DONE_STATUSES:
                continue
            code = (r.get("enrollment_code") or "").strip()
            caller = phone10(r.get("caller_id"))
            sale = (r.get("sale_date") or "").strip()
            if code and caller and sale:
                keys.add(f"{code}|{caller}|{sale}")
    return keys


# append one result row to the sent log
def write_log(row):
    new = (not os.path.exists(LOG_FILE)) or os.path.getsize(LOG_FILE) == 0
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LOG_COLUMNS)
        if new:
            w.writeheader()
        w.writerow({"time": stamp(datetime.now()), **row})


# load rows waiting to be tried again
def read_pending():
    if not os.path.exists(PENDING_FILE):
        return {"modified_time": "", "codes": {}}
    with open(PENDING_FILE, encoding="utf-8") as f:
        data = json.load(f)
    data.setdefault("modified_time", "")
    data.setdefault("codes", {})
    return data


# save the retry list
def write_pending(data):
    tmp = PENDING_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, PENDING_FILE)


# true when a waiting row may be tried now
def retry_due(entry, now):
    """True when this pending code should be tried on this pass."""
    if not entry:
        return True
    if entry.get("status") == "gave_up":
        return False
    nxt = entry.get("next_retry") or ""
    if not nxt:
        return True
    try:
        return now >= datetime.fromisoformat(nxt)
    except ValueError:
        return True


# true when any waiting row may be tried now
def any_due(pending, now):
    return any(retry_due(entry, now) for entry in pending.get("codes", {}).values())


# pick the sheet rows that should be processed now
def select_work(rows, done, pending, now, force=False):
    """Split Humana rows into counts plus the list that should hit CallGrid now."""
    ready = []
    seen = set()
    counts = {"no_code": 0, "sent": 0, "missing": 0, "not_due": 0, "gave_up": 0, "ready": 0}
    for row in rows:
        key = row_key(row) if row["code"] and row["caller"] and row["date"] else ""
        entry = pending["codes"].get(key) if key else None
        if not row["code"]:
            counts["no_code"] += 1
            continue
        if key and key in done:
            counts["sent"] += 1
            continue
        if not row["caller"] or not row["date"]:
            counts["missing"] += 1
            continue
        if key in seen:
            log.info("  row %s repeats %s %s %s - using the earlier row",
                     row["row"], row["code"], row["caller"], row["date"])
            continue
        seen.add(key)
        if not force and entry and entry.get("status") == "gave_up":
            counts["gave_up"] += 1
            continue
        if not force and entry and not retry_due(entry, now):
            counts["not_due"] += 1
            continue
        counts["ready"] += 1
        ready.append(row)
    return ready, counts


# record a miss and set the next try time
def note_retry(pending, row, status, reason, cfg, now, force=False):
    """Record a miss and set the next time this code may be tried."""
    steps = cfg["sheet"].get("retry_minutes") or [10, 30, 120]
    give_up_hours = float(cfg["sheet"].get("give_up_hours", 24))
    key = row_key(row)
    entry = pending["codes"].get(key) or {
        "attempts": 0,
        "first_seen": stamp(now),
    }
    if force and entry.get("status") == "gave_up":
        entry["attempts"] = 0
        entry["first_seen"] = stamp(now)
    entry["attempts"] = int(entry.get("attempts") or 0) + 1
    entry["client"] = row["client"]
    entry["caller"] = row["caller"]
    entry["sale_date"] = row["date"].isoformat()
    entry["reason"] = reason
    entry["last_attempt"] = stamp(now)
    first = datetime.fromisoformat(entry["first_seen"])
    if now - first >= timedelta(hours=give_up_hours):
        entry["status"] = "gave_up"
        entry["next_retry"] = ""
    else:
        delay = steps[min(entry["attempts"] - 1, len(steps) - 1)]
        entry["status"] = status
        entry["next_retry"] = stamp(now + timedelta(minutes=float(delay)))
    pending["codes"][key] = entry
    return entry


# drop pending rows that are sent or gone from the sheet
def prune_pending(pending, rows, done):
    live = {row_key(r) for r in rows if r["code"] and r["caller"] and r["date"]}
    for key in list(pending["codes"]):
        if key in done or key not in live:
            pending["codes"].pop(key, None)


# start an empty result row
def blank_log(row):
    return {"enrollment_code": row["code"], "client": row["client"], "caller_id": row["caller"],
            "sale_date": row["date"].isoformat(), "files": "", "callRecordingId": "",
            "status": "", "error": "", "method": ""}


# --------------------------------------------------------------------------- #
# Google Sheet
# --------------------------------------------------------------------------- #
# import the google auth libraries
def google_modules():
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.service_account import Credentials
    except ImportError:
        sys.exit("Install Google auth first: py -m pip install -r requirements.txt")
    return Request, Credentials


# get an access token for the service account
def google_creds():
    """Access token for the service account stored in .env."""
    info = load_config()["google_service_account"]
    Request, Credentials = google_modules()
    try:
        creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
        creds.refresh(Request())
    except Exception as e:
        raise RuntimeError(f"Google service account login failed: {e}") from e
    if not creds.token:
        raise RuntimeError("Google service account returned no access token.")
    log.info("  Google service account %s", creds.service_account_email)
    return creds


# call a google api and refresh the token on 401
def google_get(url, params):
    creds = google_creds()
    headers = {"Authorization": "Bearer " + creds.token}
    short = url.split("googleapis.com", 1)[-1]
    log.info("  Google GET %s", short)
    started = time.monotonic()
    try:
        response = requests.get(url, params=params, headers=headers, timeout=180)
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Google request failed: {e}") from e
    log.info("  Google %s in %.1fs", response.status_code, time.monotonic() - started)
    if response.status_code == 401:
        Request, _ = google_modules()
        try:
            creds.refresh(Request())
        except Exception as e:
            raise RuntimeError(f"Google service account token refresh failed: {e}") from e
        log.info("  Google token refreshed, retrying")
        headers["Authorization"] = "Bearer " + creds.token
        started = time.monotonic()
        response = requests.get(url, params=params, headers=headers, timeout=180)
        log.info("  Google %s in %.1fs", response.status_code, time.monotonic() - started)
    if response.status_code == 403:
        email = getattr(creds, "service_account_email", "the service account")
        raise RuntimeError(
            f"Google 403. Share the sheet with {email} as a viewer. {response.text[:200]}"
        )
    if not response.ok:
        raise RuntimeError(f"Google {response.status_code}: {response.text[:300]}")
    return response.json()


# return the spreadsheet id from config
def sheet_id_of(cfg):
    return (cfg["sheet"].get("spreadsheet_id") or "").strip()


# ask drive when the sheet was last changed
def source_stamp(cfg):
    """Opaque 'has the sheet changed' token. Empty when it cannot be known."""
    sheet_id = sheet_id_of(cfg)
    if sheet_id:
        data = google_get(
            f"https://www.googleapis.com/drive/v3/files/{sheet_id}",
            {"fields": "modifiedTime", "supportsAllDrives": "true"},
        )
        return data.get("modifiedTime") or ""
    path = cfg["sheet"].get("path") or ""
    if not path:
        return ""
    full = path if os.path.isabs(path) else os.path.join(HERE, path)
    if not os.path.exists(full):
        raise RuntimeError(f"Sales file not found: {full}")
    return str(os.path.getmtime(full))


# download one tab from the google sheet
def fetch_grid(sheet_id, tab):
    """Current cell values for one tab. Dates come back as Excel serial numbers."""
    a1 = "'" + tab.replace("'", "''") + "'"
    data = google_get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{quote(a1)}",
        {"valueRenderOption": "UNFORMATTED_VALUE", "dateTimeRenderOption": "SERIAL_NUMBER"},
    )
    return data.get("values") or []


# read one tab from a local excel file
def fetch_xlsx_grid(cfg):
    s = cfg["sheet"]
    path = s.get("path") or ""
    full = path if os.path.isabs(path) else os.path.join(HERE, path)
    if not path or not os.path.exists(full):
        raise RuntimeError(f"Sales file not found: {full or '(no path configured)'}")
    wb = openpyxl.load_workbook(full, data_only=True, read_only=True)
    try:
        if s["tab"] not in wb.sheetnames:
            raise RuntimeError(f"Tab '{s['tab']}' not found. Tabs: {wb.sheetnames}")
        return [list(r) for r in wb[s["tab"]].iter_rows(values_only=True)]
    finally:
        wb.close()


# return one cell, or nothing if it is missing
def cell(row, idx):
    if idx is None or not row or idx >= len(row):
        return None
    return row[idx]


# turn sheet rows into humana sales
def rows_from_table(table, cfg):
    s = cfg["sheet"]
    if not table:
        raise RuntimeError(f"Tab '{s['tab']}' is empty.")
    header = [str(h).strip().lower() if h not in (None, "") else "" for h in table[0]]

    # find a required column by name
    def col(name):
        key = name.strip().lower()
        if key not in header:
            raise RuntimeError(
                f"Column '{name}' not found in tab '{s['tab']}'. "
                f"Add it to row 1 (or fix the name in config.json -> sheet)."
            )
        return header.index(key)

    c_date, c_caller = col(s["date_column"]), col(s["caller_column"])
    c_carrier, c_client = col(s["carrier_column"]), col(s["client_column"])
    code_name = (s.get("code_column") or "").strip()
    c_code = col(code_name) if code_name else None
    carrier = s["carrier"].lower()

    # find an optional column by name
    def optional_col(config_key, names):
        configured = (s.get(config_key) or "").strip().lower()
        if configured:
            return header.index(configured) if configured in header else None
        for name in names:
            if name in header:
                return header.index(name)
        return None

    extras = [(field, optional_col(config_key, names))
              for field, config_key, names in OPTIONAL_COLUMNS]

    out = []
    for i, raw in enumerate(table[1:], start=2):
        if str(cell(raw, c_carrier) or "").strip().lower() != carrier:
            continue
        item = {
            "row": i,
            "code": as_code(cell(raw, c_code)) if c_code is not None else "",
            "caller": phone10(cell(raw, c_caller)),
            "date": as_date(cell(raw, c_date)),
            "client": " ".join(str(cell(raw, c_client) or "").split()),
        }
        for field, idx in extras:
            text = cell_text(cell(raw, idx)) if idx is not None else ""
            if text:
                item[field] = text
        out.append(item)
    apply_codes_file(out, cfg)
    return out


# load humana rows from the google sheet or excel file
def read_sheet(cfg):
    tab = cfg["sheet"]["tab"]
    log.info("  reading tab %s", tab)
    started = time.monotonic()
    sheet_id = sheet_id_of(cfg)
    if sheet_id:
        table = fetch_grid(sheet_id, tab)
    elif cfg["sheet"].get("path"):
        table = fetch_xlsx_grid(cfg)
    else:
        raise RuntimeError("config.json sheet needs spreadsheet_id or path.")
    rows = rows_from_table(table, cfg)
    log.info("  %s Humana row(s) from %s in %.1fs", len(rows), tab, time.monotonic() - started)
    return rows


# lowercase a client name for matching
def name_key(name):
    return " ".join(str(name).lower().split())


# load client name to enrollment code pairs
def read_codes_file(cfg):
    path = cfg["sheet"].get("codes_file")
    if not path or not os.path.exists(os.path.join(HERE, path)):
        return {}
    codes = {}
    with open(os.path.join(HERE, path), newline="", encoding="utf-8-sig") as f:
        for raw in csv.DictReader(f):
            raw = {k.strip().lower(): (v or "").strip() for k, v in raw.items() if k}
            if raw.get("client_name") and raw.get("enrollment_code"):
                codes[name_key(raw["client_name"])] = raw["enrollment_code"]
    return codes


# fill a blank enrollment code when the name matches one row
def apply_codes_file(rows, cfg):
    """Fill a blank Enrollment Code when the client name matches exactly one row."""
    codes = read_codes_file(cfg)
    if not codes:
        return
    carrier = cfg["sheet"]["carrier"]
    by_name = {}
    for row in rows:
        by_name.setdefault(name_key(row["client"]), []).append(row)
    for name, code in codes.items():
        matches = by_name.get(name, [])
        if len(matches) == 1 and not matches[0]["code"]:
            matches[0]["code"] = code
        elif len(matches) > 1:
            log.info("  '%s' matches %s %s rows (%s) - skipped, too ambiguous",
                     name, len(matches), carrier, ", ".join(str(m["row"]) for m in matches))
        elif not matches:
            log.info("  '%s' (%s) not found among %s rows in the sheet", name, code, carrier)


# --------------------------------------------------------------------------- #
# Single-instance lock
# --------------------------------------------------------------------------- #
# true when a process id is still running
def pid_alive(pid):
    if pid <= 0:
        return False
    if sys.platform != "win32":
        try:
            os.kill(pid, 0)
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        return True
    import ctypes
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    handle = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if handle:
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    return ctypes.windll.kernel32.GetLastError() == 5


class Lock:
    # stop a second copy from running
    def __init__(self):
        self.fd = None
        for _ in range(2):
            try:
                self.fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                os.write(self.fd, str(os.getpid()).encode())
                return
            except FileExistsError:
                if not self._stale():
                    sys.exit("Another humana_sync_v2 is already running.")
                try:
                    os.remove(LOCK_FILE)
                except OSError:
                    pass
        sys.exit("Could not take humana_sync_v2.lock.")

    # true when the lock file belongs to a dead process
    def _stale(self):
        try:
            with open(LOCK_FILE, encoding="utf-8") as f:
                pid = int((f.read() or "0").strip())
        except (OSError, ValueError):
            return True
        return not pid_alive(pid)

    # remove the lock file
    def release(self):
        if self.fd is None:
            return
        try:
            os.close(self.fd)
        except OSError:
            pass
        self.fd = None
        try:
            os.remove(LOCK_FILE)
        except OSError:
            pass


# --------------------------------------------------------------------------- #
# CallGrid
# --------------------------------------------------------------------------- #
class CallGrid:
    # set up the callgrid client
    def __init__(self, cfg):
        c = cfg["callgrid"]
        self.c = c
        self.base = c.get("base_url", "https://api.callgrid.com").rstrip("/")
        self.calls_path = c.get("calls_path", "/api/call")
        self.s = requests.Session()
        self.s.headers[c.get("auth_header", "Authorization")] = \
            c.get("auth_prefix", "Bearer ") + c["api_key"]
        fm = c.get("field_map", {})
        self.fields = {k: ([fm[k]] if fm.get(k) else []) + v for k, v in FIELD_CANDIDATES.items()}
        self._cache = {}
        self.last_reason = ""

    # ---- listing -------------------------------------------------------- #
    # post one callgrid request with no retries
    def _post(self, url, body):
        try:
            return self.s.post(url, json=body, timeout=60)
        except (requests.exceptions.ReadTimeout, requests.exceptions.ConnectionError) as e:
            raise RuntimeError(f"CallGrid {type(e).__name__}") from e

    # list this caller's calls around the sale date
    def raw_calls(self, caller, day):
        """Ask CallGrid for just this caller's calls around the sale date (server-side filter)."""
        key = (caller, day)
        if key in self._cache:
            return self._cache[key]
        body = {
            "startDate": (day - timedelta(days=1)).isoformat(),
            "endDate": (day + timedelta(days=2)).isoformat(),
            "maxItems": 100,
            "filters": {"items": [{"rules": [
                {"tagName": "CallerId", "condition": "contains", "values": [caller]}]}]},
        }
        log.info("  CallGrid lookup %s (%s to %s)", caller, body["startDate"], body["endDate"])
        started = time.monotonic()
        r = self._post(self.base + self.calls_path, body)
        log.info("  CallGrid replied %s in %.1fs", r.status_code, time.monotonic() - started)
        if not r.ok:
            raise RuntimeError(f"CallGrid {r.status_code}: {r.text[:300]}")
        data = r.json()
        rows = data.get("data", []) if isinstance(data, dict) else data
        rows = [c for c in rows if phone10(self.get(c, "caller")) == caller]
        self._cache[key] = rows
        return rows

    # read the first filled field on a call
    def get(self, call, field):
        for key in self.fields[field]:
            if key in call and call[key] not in (None, ""):
                return call[key]
        return None

    # fetch one call when the list row has no recording url
    def detail(self, call_id):
        r = self.s.get(f"{self.base}{self.calls_path}/{call_id}", timeout=60)
        if not r.ok:
            return {}
        data = r.json()
        return data.get("data", data) if isinstance(data, dict) else {}

    # ---- matching ------------------------------------------------------- #
    # keep recorded calls near the sale date, longest first
    def matching_calls(self, caller, day):
        """Recorded calls from this caller within a day of the sale, longest first."""
        min_sec = self.c.get("min_call_seconds", 120)
        mine = self.raw_calls(caller, day)
        n_date = n_short = n_norec = 0
        found = []
        for call in mine:
            start = to_utc(self.get(call, "start"))
            if not start or abs((start.date() - day).days) > 1:
                n_date += 1
                continue
            end = to_utc(self.get(call, "end"))
            dur = self.get(call, "duration")
            if not end and dur is not None:
                end = start + timedelta(seconds=float(dur))
            if not end or (end - start).total_seconds() < min_sec:
                n_short += 1
                continue
            url = self.get(call, "recording")
            if isinstance(url, dict):
                url = url.get("url") or url.get("href")
            if not url and self.get(call, "id"):            # list row had no URL: ask for detail
                url = self.get(self.detail(self.get(call, "id")), "recording")
            if not url:
                n_norec += 1
                continue
            found.append({"id": str(self.get(call, "id") or start.timestamp()),
                          "url": url, "start": start, "end": end,
                          "seconds": (end - start).total_seconds()})
        found.sort(key=lambda x: x["seconds"], reverse=True)

        if not found:
            if not mine:
                self.last_reason = (f"CallGrid has NO calls from {caller} between "
                                    f"{day - timedelta(days=1)} and {day + timedelta(days=2)}. "
                                    f"Check the Caller ID / date in the sheet.")
            else:
                self.last_reason = (f"{len(mine)} call(s) from {caller} found but none usable: "
                                    f"{n_date} outside date range, {n_short} shorter than "
                                    f"{min_sec}s, {n_norec} with no recording")
        return found if self.c.get("send_all_calls_that_day") else found[:1]

    # ---- download ------------------------------------------------------- #
    # save a recording to disk
    def download(self, url, dest_base):
        """Download with the API key, follow redirects, return the saved file path."""
        log.info("  downloading recording")
        started = time.monotonic()
        r = self.s.get(url, timeout=300, stream=True)
        ctype = r.headers.get("Content-Type", "").lower()
        if r.ok and "json" in ctype:
            try:
                signed = r.json().get("url")
            except ValueError:
                signed = None
            r.close()
            if signed:
                log.info("  following signed recording link")
                r = requests.get(signed, timeout=300, stream=True)
                ctype = r.headers.get("Content-Type", "").lower()
        if not r.ok:
            detail = ""
            try:
                detail = r.text[:200]
            except Exception:
                pass
            raise RuntimeError(f"Recording download failed {r.status_code}: {detail}")
        if "json" in ctype or "html" in ctype or ctype.startswith("text/"):
            raise RuntimeError(f"Recording download returned text, not audio: {r.text[:200]}")
        ext = AUDIO_EXT.get(ctype.split(";")[0].strip(), ".mp3")
        path = dest_base + ext
        total = int(r.headers.get("Content-Length") or 0)
        got = 0
        last = time.monotonic()
        with open(path, "wb") as f:
            for chunk in r.iter_content(chunk_size=256 * 1024):
                if not chunk:
                    continue
                f.write(chunk)
                got += len(chunk)
                now = time.monotonic()
                if now - last >= 5:
                    if total:
                        log.info("  downloading %s / %s", fmt_bytes(got), fmt_bytes(total))
                    else:
                        log.info("  downloading %s so far", fmt_bytes(got))
                    last = now
        if got < 1000:
            raise RuntimeError(f"Recording download too small ({got} bytes)")
        log.info("  saved %s (%s) in %.1fs", os.path.basename(path), fmt_bytes(got),
                 time.monotonic() - started)
        return path


# --------------------------------------------------------------------------- #
# Sunfire
# --------------------------------------------------------------------------- #
class Sunfire:
    # log in to the sunfire enrollment api
    def __init__(self, cfg):
        c = cfg["sunfire"]
        self.host = c["host"].rstrip("/")
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "curl/8.9.1"
        log.info("  Sunfire login")
        started = time.monotonic()
        r = self.s.post(f"{self.host}/v2/authenticate",
                        json={"username": c["username"], "password": c["password"]}, timeout=30)
        log.info("  Sunfire login -> %s in %.1fs", r.status_code, time.monotonic() - started)
        if r.status_code == 403 and "<html" in r.text.lower():
            raise RuntimeError("Sunfire blocked the connection - is the US VPN on?")
        if not r.ok:
            raise RuntimeError(f"Sunfire login failed {r.status_code}: {r.text[:300]}")
        self.s.headers["Authorization"] = "Bearer " + r.json()["token"]

    # register, upload, and finalize one enrollment
    def send(self, code, files):
        """files: list of {path, start, end}. Returns (status, callRecordingId)."""
        names = {os.path.basename(f["path"]): f["path"] for f in files}
        body = {"sunfireEnrollmentCode": code, "callMetadata": {"recordings": [
            {"filename": os.path.basename(f["path"]), "startTime": iso_z(f["start"]),
             "endTime": iso_z(f["end"])} for f in files]}}
        log.info("  Sunfire register %s (%s file(s))", code, len(files))
        started = time.monotonic()
        r = self.s.post(f"{self.host}/api/ix/callrecordings/internalenrollment", json=body, timeout=60)
        log.info("  Sunfire register -> %s in %.1fs", r.status_code, time.monotonic() - started)
        if r.status_code == 409:
            return "already_exists", ""
        if not r.ok:
            raise RuntimeError(f"Register {r.status_code}: {r.text[:300]}")
        call_id = r.json()["callRecordingId"]
        log.info("  Sunfire call %s", call_id)
        for t in r.json()["recordingUploads"]:
            path = names[t["filename"]]
            log.info("  Sunfire upload %s (%s)", t["filename"], fmt_bytes(os.path.getsize(path)))
            started = time.monotonic()
            with open(path, "rb") as f:
                u = requests.put(t["presignedUrl"], data=f, timeout=900)
            log.info("  Sunfire PUT %s -> %s in %.1fs", t["filename"], u.status_code,
                     time.monotonic() - started)
            if not u.ok:
                raise RuntimeError(f"Upload {t['filename']} {u.status_code} (callRecordingId {call_id})")
        log.info("  Sunfire finalize %s", call_id)
        r = self.s.post(f"{self.host}/api/ix/callrecordings/{call_id}/finalize", timeout=60)
        log.info("  Sunfire finalize -> %s", r.status_code)
        if not r.ok:
            raise RuntimeError(f"Finalize {r.status_code}: {r.text[:300]} (callRecordingId {call_id})")
        return r.json().get("status", "finalized"), call_id


class IntelligenceX:
    """Auth0 external call API. Used only after the Sunfire enrollment upload fails."""

    # set up the auth0 external api client
    def __init__(self, cfg):
        c = cfg.get("ix") or {}
        missing = [k for k in ("organization_id", "client_id", "client_secret")
                   if not str(c.get(k) or "").strip()]
        if missing:
            raise RuntimeError(
                "config.json ix is missing: " + ", ".join(missing) + ". "
                "Copy them from Admin > Integrations."
            )
        self.org_id = str(c["organization_id"]).strip()
        self.client_id = str(c["client_id"]).strip()
        self.client_secret = str(c["client_secret"]).strip()
        self.host = c.get("host", "https://api.compliancex.ai").rstrip("/")
        self.auth_url = c.get("auth_url", "https://sunfire.sunfirematrix.auth0.com/oauth/token")
        self.audience = c.get("audience", "https://api.compliancex.ai")
        self.call_type = c.get("call_type") or "enrollment"
        self.disposition = c.get("disposition") or "sold"
        self.samples = dict(SAMPLE_FIELDS)
        for key in SAMPLE_FIELDS:
            if str(c.get(key) or "").strip():
                self.samples[key] = str(c[key]).strip()
        self.carrier = cfg["sheet"].get("carrier") or "Humana"
        self.s = requests.Session()
        self.token = ""
        self.token_exp = 0.0

    # get an auth0 token if the current one is old
    def ensure_token(self):
        now = time.time()
        if self.token and now < self.token_exp - 120:
            return
        log.info("  requesting Auth0 token")
        r = requests.post(self.auth_url, json={
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "audience": self.audience,
            "scope": "access:external",
        }, timeout=30)
        if not r.ok:
            raise RuntimeError(f"Auth0 {r.status_code}: {r.text[:300]}")
        data = r.json()
        if not data.get("access_token"):
            raise RuntimeError(f"Auth0 returned no access_token: {r.text[:300]}")
        self.token = data["access_token"]
        self.token_exp = now + int(data.get("expires_in") or 86400)
        self.s.headers["Authorization"] = "Bearer " + self.token
        log.info("  Auth0 token ok for %.0f min", (self.token_exp - now) / 60)

    # call the external api and refresh the token on 401
    def request(self, method, path, **kwargs):
        kwargs.setdefault("timeout", 60)
        self.ensure_token()
        started = time.monotonic()
        r = self.s.request(method, self.host + path, **kwargs)
        log.info("  %s %s -> %s in %.1fs", method, path, r.status_code, time.monotonic() - started)
        if r.status_code == 401:
            log.info("  token rejected, requesting a new one")
            self.token = ""
            self.ensure_token()
            started = time.monotonic()
            r = self.s.request(method, self.host + path, **kwargs)
            log.info("  %s %s -> %s in %.1fs", method, path, r.status_code, time.monotonic() - started)
        return r

    # use the sheet value, or a sample if it is blank
    def sample(self, row, key, real=""):
        return real or row.get(key) or self.samples[key]

    # create a call, upload the audio, and finalize it
    def send(self, row, files):
        """files: list of {path, start, end}. Returns (status, call id)."""
        sold = files[0]["start"]
        effective = row["date"].replace(day=1)
        if row["date"].month == 12:
            effective = effective.replace(year=row["date"].year + 1, month=1)
        else:
            effective = effective.replace(month=row["date"].month + 1)
        call = {
            "uploaded_at": row["date"].isoformat(),
            "enrollee_name": self.sample(row, "enrollee_name", row.get("client")),
            "enrollee_phone": self.sample(row, "enrollee_phone",
                                          "+1" + row["caller"] if row.get("caller") else ""),
            "enrollee_id": self.sample(row, "enrollee_id"),
            "call_type": self.call_type,
            "disposition": self.disposition,
            "carrier_name": self.carrier,
            "sunfire_enrollment_code": row["code"],
            "zip_code": self.sample(row, "zip_code"),
            "date_sold": iso_z(sold),
            "date_effective": iso_z(datetime(effective.year, effective.month, effective.day,
                                             tzinfo=timezone.utc)),
            "agent_name": self.sample(row, "agent_name"),
            "agent_email": self.sample(row, "agent_email"),
            "agent_npn": self.sample(row, "agent_npn"),
            "team_name": self.sample(row, "team_name"),
            "policy_name": self.sample(row, "policy_name"),
            "policy_number": self.sample(row, "policy_number"),
            "affiliate_key": self.sample(row, "affiliate_key"),
        }
        log.info("  external create %s (%s, %s)", row["code"], call["enrollee_name"],
                 call["enrollee_phone"])
        r = self.request("POST", "/api/external/calls",
                         json={"organization_id": self.org_id, "call": call})
        if not r.ok:
            raise RuntimeError(f"Create call {r.status_code}: {r.text[:400]}")
        call_id = r.json().get("id") or ""
        if not call_id:
            raise RuntimeError(f"Create call returned no id: {r.text[:400]}")
        log.info("  created %s", call_id)

        used_names = set()
        for f in files:
            filename = os.path.basename(f["path"])
            stem, ext = os.path.splitext(filename)
            n = 2
            while filename in used_names:
                filename = f"{stem}_{n}{ext}"
                n += 1
            used_names.add(filename)
            nbytes = os.path.getsize(f["path"])
            log.info("  registering %s (%s)", filename, fmt_bytes(nbytes))
            r = self.request("POST", f"/api/external/calls/{call_id}/recordings", json={
                "recording": {
                    "filename": filename,
                    "start_time": iso_z(f["start"]),
                    "end_time": iso_z(f["end"]),
                }
            })
            if not r.ok:
                raise RuntimeError(f"Register {filename} {r.status_code}: {r.text[:400]} ({call_id})")
            data = r.json()
            upload_url = data.get("presigned_url") or data.get("presignedUrl")
            if not upload_url:
                raise RuntimeError(f"No presigned URL for {filename} ({call_id})")
            ctype = CONTENT_TYPE.get(ext.lower(), "application/octet-stream")
            log.info("  uploading %s (%s) to S3", filename, fmt_bytes(nbytes))
            started = time.monotonic()
            with open(f["path"], "rb") as fh:
                u = requests.put(upload_url, data=fh, headers={"Content-Type": ctype}, timeout=900)
            log.info("  S3 PUT %s -> %s in %.1fs", filename, u.status_code, time.monotonic() - started)
            if not u.ok:
                raise RuntimeError(f"Upload {filename} {u.status_code}: {u.text[:200]} ({call_id})")

        log.info("  finalizing %s", call_id)
        r = self.request("POST", f"/api/external/calls/{call_id}/finalize")
        if not r.ok:
            raise RuntimeError(f"Finalize {r.status_code}: {r.text[:400]} ({call_id})")
        return (r.json() or {}).get("status") or "finalized", call_id


# --------------------------------------------------------------------------- #
# Commands
# --------------------------------------------------------------------------- #
# format one line of sheet counts
def summary_line(cfg, rows, counts, due):
    return (f"{cfg['sheet']['carrier']} rows: {len(rows)} | "
            f"sent: {counts['sent']} | no code: {counts['no_code']} | "
            f"missing fields: {counts['missing']} | not due yet: {counts['not_due']} | "
            f"gave up: {counts['gave_up']} | due now: {due}")


# true when the sheet should be read again on the timer
def read_is_stale(pending, now, cfg):
    """Re-read on a timer even when Drive still reports the same modifiedTime."""
    minutes = float(cfg["sheet"].get("full_read_minutes", 15))
    last = pending.get("checked_at") or ""
    if not last:
        return True
    try:
        return now - datetime.fromisoformat(last) >= timedelta(minutes=minutes)
    except ValueError:
        return True


# true when the sheet is unchanged and nothing is due
def sheet_is_idle(cfg, pending, now, force):
    if force:
        log.info("  forced read, ignoring the last sheet check")
        return False, ""
    log.info("  checking if the Google Sheet changed")
    modified = source_stamp(cfg)
    busy = any_due(pending, now) or read_is_stale(pending, now, cfg)
    if modified and modified == pending.get("modified_time") and not busy:
        return True, modified
    return False, modified


# save the sheet check time and drop finished pending rows
def finish_pass(pending, rows, done, modified, dry_run):
    if dry_run:
        return
    if modified:
        pending["modified_time"] = modified
    pending["checked_at"] = stamp(datetime.now())
    prune_pending(pending, rows, done)
    write_pending(pending)


# check that the google service account can get a token
def cmd_auth(cfg, args):
    creds = google_creds()
    log.info("Service account is working. Share the sheet with %s as a viewer.",
             creds.service_account_email)


# print what is due without calling callgrid
def cmd_check_sheet(cfg, args):
    now = datetime.now().replace(microsecond=0)
    rows = read_sheet(cfg)
    done, pending = read_log(), read_pending()
    ready, counts = select_work(rows, done, pending, now, force=False)
    log.info(summary_line(cfg, rows, counts, len(ready)))
    for row in rows:
        key = row_key(row) if row["caller"] and row["date"] else ""
        if not row["code"] or (key and key in done):
            continue
        entry = pending["codes"].get(key) or {}
        if not row["caller"] or not row["date"]:
            log.info("  row %s (%s): missing/invalid Caller ID or Date", row["row"], row["client"])
        elif entry.get("status") == "gave_up":
            log.info("  row %s %s (%s): gave up - %s",
                     row["row"], row["client"], row["code"], entry.get("reason", ""))
        elif entry.get("next_retry"):
            log.info("  row %s %s (%s): %s until %s",
                     row["row"], row["client"], row["code"], entry.get("status"), entry["next_retry"])


# download one caller's recording and do not upload it
def cmd_test_callgrid(cfg, args):
    day = as_date(args.date)
    if not day:
        sys.exit(f"Date must look like 2026-10-17 (got {args.date!r})")
    caller = phone10(args.caller)
    cg = CallGrid(cfg)
    raw = cg.raw_calls(caller, day)
    log.info("CallGrid returned %s call(s) from %s.", len(raw), caller)
    for c in raw:
        log.info("  %s", {k: c.get(k) for k in ("createdAt", "CallerId", "callDuration",
                                                 "CampaignName", "callRecordingUrl")})
    matches = cg.matching_calls(caller, day)
    log.info("Matched %s usable recorded call(s).", len(matches))
    if not matches:
        log.info("Why not: %s", cg.last_reason)
        return
    os.makedirs(REC_DIR, exist_ok=True)
    for i, m in enumerate(matches, 1):
        log.info("  %s -> %s  (%.0fs)", iso_z(m["start"]), iso_z(m["end"]), m["seconds"])
        base = os.path.join(REC_DIR, f"{caller}_{day.isoformat()}_{i}")
        path = cg.download(m["url"], base)
        log.info("  saved %s (%s)", path, fmt_bytes(os.path.getsize(path)))


# pick the local filename for this recording
def file_base(row, index, code_rows):
    rows_for_code = code_rows.get(row["code"]) or [row["row"]]
    if len(rows_for_code) == 1 or row["row"] == min(rows_for_code):
        return os.path.join(REC_DIR, f"{row['code']}_{index}")
    return os.path.join(REC_DIR, f"{row['code']}_row{row['row']}_{index}")


class LazyClient:
    """Build an API client on first use and remember a login failure for this run."""

    # remember how to build an api client
    def __init__(self, factory):
        self.factory = factory
        self.obj = None
        self.err = None

    # build the client on first use and reuse a login failure
    def get(self):
        if self.err:
            raise self.err
        if self.obj is None:
            try:
                self.obj = self.factory()
            except Exception as e:
                self.err = e
                raise
        return self.obj


# try sunfire, then the external api if that fails
def upload_row(sunfire, ix, row, files):
    """Sunfire enrollment first. The external API runs only when that upload fails."""
    try:
        status, call_id = sunfire.get().send(row["code"], files)
        return status, call_id, "sunfire"
    except Exception as e:
        log.info("  Sunfire failed: %s", e)
        log.info("  falling back to external API")
        status, call_id = ix.get().send(row, files)
        return status, call_id, "external"


# delete local audio after a successful upload
def delete_recordings(files):
    """Remove local audio after a successful upload so recordings/ does not keep growing."""
    for item in files:
        path = item["path"]
        try:
            os.remove(path)
        except FileNotFoundError:
            continue
        except OSError as e:
            log.info("  could not delete %s: %s", os.path.basename(path), e)
            continue
        log.info("  deleted %s", os.path.basename(path))


# save a miss and schedule the next try
def record_miss(cfg, args, pending, row, status, reason, now):
    if args.dry_run:
        log.info("  would mark %s: %s", status, reason)
        return
    entry = note_retry(pending, row, status, reason, cfg, now, force=args.force)
    write_pending(pending)
    log_row = blank_log(row)
    log_row.update(status=entry["status"], error=reason)
    write_log(log_row)
    when = entry["next_retry"] or "gave up"
    log.info("  %s - next try %s", entry["status"], when)


# process every due row once
def cmd_run(cfg, args):
    now = datetime.now().replace(microsecond=0)
    pending = read_pending()
    idle, modified = sheet_is_idle(cfg, pending, now, args.force)
    if idle:
        log.info("Google Sheet unchanged (%s) and nothing is due.", modified)
        return

    log.info("reading Google Sheet")
    rows = read_sheet(cfg)
    done = read_log()
    ready, counts = select_work(rows, done, pending, now, force=args.force)
    if args.limit:
        ready = ready[:args.limit]
    dry = " (DRY RUN)" if args.dry_run else ""
    log.info("%s%s", summary_line(cfg, rows, counts, len(ready)), dry)
    if not ready:
        finish_pass(pending, rows, done, modified, args.dry_run)
        return

    os.makedirs(REC_DIR, exist_ok=True)
    cg = CallGrid(cfg)
    code_rows = {}
    for row in rows:
        if row["code"]:
            code_rows.setdefault(row["code"], []).append(row["row"])
    sunfire = LazyClient(lambda: Sunfire(cfg))
    ix = LazyClient(lambda: IntelligenceX(cfg))
    ok = waiting = fail = 0

    for n, row in enumerate(ready, 1):
        key = row_key(row)
        if key in done:
            continue
        label = f"[{n}/{len(ready)}] row {row['row']} {row['client']} ({row['code']})"
        log.info("%s caller %s sale %s", label, row["caller"], row["date"].isoformat())
        log_row = blank_log(row)
        try:
            calls = cg.matching_calls(row["caller"], row["date"])
            if not calls:
                reason = cg.last_reason or "no recorded CallGrid call found"
                log.info("  no recording yet - %s", reason)
                record_miss(cfg, args, pending, row, "waiting", reason, datetime.now())
                waiting += 1
                continue
            log.info("  matched %s recording(s)", len(calls))
            files = []
            for i, c in enumerate(calls, 1):
                base = file_base(row, i, code_rows)
                existing = [p for p in glob.glob(base + ".*") if os.path.getsize(p) > 1000]
                if existing:
                    path = existing[0]
                    log.info("  reusing %s (%s, %.0fs)", os.path.basename(path),
                             fmt_bytes(os.path.getsize(path)), c["seconds"])
                else:
                    log.info("  need recording %s (%.0fs)", c["id"], c["seconds"])
                    path = cg.download(c["url"], base)
                files.append({"path": path, "start": c["start"], "end": c["end"]})
            log_row["files"] = "; ".join(os.path.basename(f["path"]) for f in files)

            if args.dry_run:
                log.info("  dry run, not sent (%s)", log_row["files"])
                continue
            status, call_id, method = upload_row(sunfire, ix, row, files)
            log_row.update(status=status, callRecordingId=call_id, method=method)
            write_log(log_row)
            if status in DONE_STATUSES:
                done.add(key)
                pending["codes"].pop(key, None)
                write_pending(pending)
                delete_recordings(files)
            log.info("  done %s via %s %s", status, method, call_id)
            ok += 1
        except Exception as e:
            log.info("  FAILED %s", e)
            record_miss(cfg, args, pending, row, "failed", str(e), datetime.now())
            fail += 1

    finish_pass(pending, rows, done, modified, args.dry_run)
    if not args.dry_run:
        log.info("Done: %s sent, %s waiting on a recording, %s failed.", ok, waiting, fail)


# keep reading the sheet and uploading on a timer
def cmd_poll(cfg, args):
    minutes = args.minutes if args.minutes is not None else float(cfg["sheet"].get("poll_minutes", 3))
    if minutes <= 0:
        raise RuntimeError("poll interval must be greater than 0 minutes.")
    log.info("Polling the Google Sheet every %s minutes. Ctrl+C to stop.", f"{minutes:g}")
    while True:
        started = time.monotonic()
        try:
            cmd_run(cfg, args)
        except KeyboardInterrupt:
            log.info("Stopped.")
            return
        except Exception as e:
            log.info("Poll error: %s", e)
        remaining = minutes * 60 - (time.monotonic() - started)
        if remaining > 0:
            log.info("next Google Sheet check in %.0fs", remaining)
            try:
                time.sleep(remaining)
            except KeyboardInterrupt:
                log.info("Stopped.")
                return


# add --dry-run, --limit, and --force
def add_work_args(parser, polling=False):
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--limit", type=int)
    parser.add_argument("--force", action="store_true",
                        help="Ignore the retry schedule and try every unsent code now")
    if polling:
        parser.add_argument("--minutes", type=float, help="Override config sheet.poll_minutes")


# parse the command and run it
def main():
    p = argparse.ArgumentParser(
        description="Send Humana recordings from the Google Sheet. Sunfire first, external API if that fails.")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("auth", help="Check the Google service account in .env")
    sub.add_parser("check-sheet")
    t = sub.add_parser("test-callgrid")
    t.add_argument("--caller", required=True)
    t.add_argument("--date", required=True, help="YYYY-MM-DD (2026-10-1 is ok)")
    add_work_args(sub.add_parser("run"))
    add_work_args(sub.add_parser("poll"), polling=True)
    args = p.parse_args()

    setup_logging()
    cfg = load_config()
    needs_lock = args.cmd in ("run", "poll")
    lock = Lock() if needs_lock else None
    try:
        {"auth": cmd_auth, "check-sheet": cmd_check_sheet, "test-callgrid": cmd_test_callgrid,
         "run": cmd_run, "poll": cmd_poll}[args.cmd](cfg, args)
    except RuntimeError as e:
        sys.exit(str(e))
    finally:
        if lock:
            lock.release()


if __name__ == "__main__":
    main()
