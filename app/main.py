import contextlib
import gzip
import json
import io
import os
import secrets
import sqlite3
import threading
from functools import lru_cache
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from uuid import uuid4
import requests
from dotenv import load_dotenv
from fitparse import FitFile
from fastapi import Body, FastAPI, HTTPException, Query, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

load_dotenv()

app = FastAPI()
app.mount("/icons", StaticFiles(directory="icons"), name="icons")
app.mount("/static", StaticFiles(directory="app/static"), name="static")


@app.on_event("startup")
def on_startup() -> None:
    init_db()
    migrate_from_json()

TOKEN_FILE = Path("data/strava_tokens.json")
CALENDAR_FILE = Path("data/calendar_items.json")
PAIRS_FILE = Path("data/workout_pairs.json")
ACTIVITY_OVERRIDES_FILE = Path("data/activity_overrides.json")
IMPORTED_ACTIVITIES_FILE = Path("data/imported_activities.json")
FIT_PARSED_DIR = Path("data/fit_parsed")
SETTINGS_FILE = Path("data/settings.json")
PLANNED_FILE = Path("data/planned_workouts.json")
TEMPLATES_DIR = Path("app/templates")
DB_PATH = Path("data/trainingfreaks.db")
TP_EXPORT_WORKOUT_ROOTS = (
    Path("tp_export/athlete_4211127/full_history/workouts"),
    Path("tp_export/athlete_4211127/manual_test/workouts"),
)
STRAVA_TOKEN_URL = "https://www.strava.com/oauth/token"
STRAVA_ACTIVITIES_URL = "https://www.strava.com/api/v3/athlete/activities"
FILE_LOCK = threading.Lock()
_pending_oauth_state: str = ""


def read_json_file(path: Path, default: Any) -> Any:
    with FILE_LOCK:
        if not path.exists():
            return default
        try:
            return json.loads(path.read_text())
        except json.JSONDecodeError:
            return default


def write_json_file(path: Path, payload: Any) -> None:
    with FILE_LOCK:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, indent=2))


# ---------------------------------------------------------------------------
# SQLite helpers
# ---------------------------------------------------------------------------

@contextlib.contextmanager
def get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    with get_db() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS activities (
                id TEXT PRIMARY KEY,
                source TEXT NOT NULL DEFAULT 'fit',
                name TEXT,
                type TEXT,
                start_date_local TEXT,
                distance REAL,
                moving_time REAL,
                description TEXT,
                comments TEXT,
                comments_feed TEXT,
                feel INTEGER,
                rpe INTEGER,
                tss_override REAL,
                tss_source TEXT,
                if_value REAL,
                np_value REAL,
                hr_tss REAL,
                work_kj REAL,
                calories REAL,
                avg_speed REAL,
                avg_power REAL,
                avg_hr REAL,
                min_hr REAL,
                max_hr REAL,
                min_power REAL,
                max_power REAL,
                elev_gain_m REAL,
                fit_id TEXT,
                fit_filename TEXT,
                fit_data BLOB,
                fit_parsed_json TEXT,
                duration_min REAL,
                distance_km REAL,
                distance_m REAL,
                elevation_m REAL,
                distance_unit TEXT,
                elevation_unit TEXT,
                planned_tss REAL,
                planned_if REAL,
                planned_avg_speed REAL,
                planned_calories REAL,
                planned_work_kj REAL,
                completed_duration_min REAL,
                analysis_edits TEXT,
                hidden INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT (datetime('now'))
            )
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS activity_overrides (
                id TEXT PRIMARY KEY,
                title TEXT,
                date TEXT,
                type TEXT,
                description TEXT,
                comments TEXT,
                comments_feed TEXT,
                feel INTEGER,
                rpe INTEGER,
                tss_override REAL,
                if_value REAL,
                tss_source TEXT,
                duration_min REAL,
                distance_km REAL,
                distance_m REAL,
                elevation_m REAL,
                distance_unit TEXT,
                elevation_unit TEXT,
                planned_tss REAL,
                planned_if REAL,
                planned_avg_speed REAL,
                planned_calories REAL,
                planned_work_kj REAL,
                completed_duration_min REAL,
                analysis_edits TEXT,
                hidden INTEGER NOT NULL DEFAULT 0
            )
        """)


def migrate_from_json() -> None:
    """One-time migration from legacy JSON files to SQLite. Safe to call repeatedly."""
    if IMPORTED_ACTIVITIES_FILE.exists():
        items = read_json_file(IMPORTED_ACTIVITIES_FILE, [])
        with get_db() as db:
            for item in items:
                if db.execute("SELECT id FROM activities WHERE id = ?", (item["id"],)).fetchone():
                    continue
                fit_data = None
                fit_parsed_json = None
                fit_id = item.get("fit_id")
                if fit_id:
                    fit_path = Path("data/imports") / f"{fit_id}.fit"
                    if fit_path.exists():
                        fit_data = fit_path.read_bytes()
                    parsed = read_json_file(FIT_PARSED_DIR / f"{fit_id}.json", {})
                    if parsed:
                        fit_parsed_json = json.dumps(parsed)
                cf = item.get("comments_feed", [])
                ae = item.get("analysis_edits", {})
                db.execute(
                    _activity_insert_sql(),
                    _activity_insert_params(item, fit_data, fit_parsed_json, cf, ae),
                )

    if ACTIVITY_OVERRIDES_FILE.exists():
        overrides = read_json_file(ACTIVITY_OVERRIDES_FILE, {})
        with get_db() as db:
            for aid, override in overrides.items():
                if db.execute("SELECT id FROM activity_overrides WHERE id = ?", (aid,)).fetchone():
                    continue
                _upsert_override(db, aid, override)


def _activity_insert_sql() -> str:
    return """
        INSERT OR IGNORE INTO activities (
            id, source, name, type, start_date_local, distance, moving_time,
            description, comments, comments_feed, feel, rpe,
            tss_override, tss_source, if_value, np_value, hr_tss,
            work_kj, calories, avg_speed, avg_power, avg_hr, min_hr, max_hr,
            min_power, max_power, elev_gain_m,
            fit_id, fit_filename, fit_data, fit_parsed_json,
            duration_min, distance_km, distance_m, elevation_m,
            distance_unit, elevation_unit,
            analysis_edits, hidden, created_at
        ) VALUES (
            ?,?,?,?,?,?,?,
            ?,?,?,?,?,
            ?,?,?,?,?,
            ?,?,?,?,?,?,?,
            ?,?,?,
            ?,?,?,?,
            ?,?,?,?,
            ?,?,
            ?,?,?
        )
    """


def _activity_insert_params(
    item: dict[str, Any],
    fit_data: bytes | None,
    fit_parsed_json: str | None,
    cf: list,
    ae: dict,
) -> tuple:
    return (
        item["id"],
        item.get("source", "fit"),
        item.get("name"),
        item.get("type"),
        item.get("start_date_local"),
        item.get("distance"),
        item.get("moving_time"),
        item.get("description", ""),
        item.get("comments", ""),
        json.dumps(cf if isinstance(cf, list) else []),
        item.get("feel"),
        item.get("rpe"),
        item.get("tss_override"),
        item.get("tss_source"),
        item.get("if_value"),
        item.get("np_value"),
        item.get("hr_tss"),
        item.get("work_kj"),
        item.get("calories"),
        item.get("avg_speed"),
        item.get("avg_power"),
        item.get("avg_hr"),
        item.get("min_hr"),
        item.get("max_hr"),
        item.get("min_power"),
        item.get("max_power"),
        item.get("elev_gain_m"),
        item.get("fit_id"),
        item.get("fit_filename"),
        fit_data,
        fit_parsed_json,
        item.get("duration_min"),
        item.get("distance_km"),
        item.get("distance_m"),
        item.get("elevation_m"),
        item.get("distance_unit"),
        item.get("elevation_unit"),
        json.dumps(ae if isinstance(ae, dict) else {}),
        1 if item.get("hidden") else 0,
        item.get("created_at", datetime.utcnow().isoformat(timespec="seconds") + "Z"),
    )


def _upsert_override(db: sqlite3.Connection, aid: str, override: dict[str, Any]) -> None:
    cf = override.get("comments_feed", [])
    ae = override.get("analysis_edits", {})
    db.execute(
        """
        INSERT OR REPLACE INTO activity_overrides (
            id, title, date, type, description, comments, comments_feed,
            feel, rpe, tss_override, if_value, tss_source,
            duration_min, distance_km, distance_m, elevation_m,
            distance_unit, elevation_unit,
            planned_tss, planned_if, planned_avg_speed, planned_calories, planned_work_kj,
            completed_duration_min, analysis_edits, hidden
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        (
            aid,
            override.get("title"),
            override.get("date"),
            override.get("type"),
            override.get("description"),
            override.get("comments"),
            json.dumps(cf if isinstance(cf, list) else []),
            override.get("feel"),
            override.get("rpe"),
            override.get("tss_override"),
            override.get("if_value"),
            override.get("tss_source"),
            override.get("duration_min"),
            override.get("distance_km"),
            override.get("distance_m"),
            override.get("elevation_m"),
            override.get("distance_unit"),
            override.get("elevation_unit"),
            override.get("planned_tss"),
            override.get("planned_if"),
            override.get("planned_avg_speed"),
            override.get("planned_calories"),
            override.get("planned_work_kj"),
            override.get("completed_duration_min"),
            json.dumps(ae if isinstance(ae, dict) else {}),
            1 if override.get("hidden") else 0,
        ),
    )


def row_to_activity(row: sqlite3.Row) -> dict[str, Any]:
    d = dict(row)
    d.pop("fit_data", None)
    d.pop("fit_parsed_json", None)
    for col in ("comments_feed", "analysis_edits"):
        raw = d.get(col)
        if raw and isinstance(raw, str):
            try:
                d[col] = json.loads(raw)
            except json.JSONDecodeError:
                d[col] = [] if col == "comments_feed" else {}
        elif raw is None:
            d[col] = [] if col == "comments_feed" else {}
    return d


def override_to_dict(row: sqlite3.Row) -> dict[str, Any]:
    d = dict(row)
    for col in ("comments_feed", "analysis_edits"):
        raw = d.get(col)
        if raw and isinstance(raw, str):
            try:
                d[col] = json.loads(raw)
            except json.JSONDecodeError:
                d[col] = [] if col == "comments_feed" else {}
        elif raw is None:
            d[col] = [] if col == "comments_feed" else {}
    return d


def save_tokens(token_data: dict) -> None:
    write_json_file(TOKEN_FILE, token_data)


def load_tokens() -> dict:
    data = read_json_file(TOKEN_FILE, {})
    if not data:
        raise HTTPException(status_code=400, detail="No saved Strava tokens found.")
    return data


def refresh_access_token(refresh_token: str) -> dict:
    client_id = os.getenv("STRAVA_CLIENT_ID")
    client_secret = os.getenv("STRAVA_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise HTTPException(status_code=500, detail="Missing Strava client credentials.")

    resp = requests.post(
        STRAVA_TOKEN_URL,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        },
        timeout=30,
    )
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail=resp.text)

    token_data = resp.json()
    save_tokens(token_data)
    return token_data


def fetch_activities(after: int | None = None, before: int | None = None, per_page: int = 100) -> list[dict[str, Any]]:
    token_data = load_tokens()
    access_token = token_data.get("access_token")
    if not access_token:
        raise HTTPException(status_code=400, detail="Saved token file is missing access_token.")

    def do_fetch(token: str) -> list[dict[str, Any]] | None:
        all_items: list[dict[str, Any]] = []
        page = 1
        max_pages = 10
        while page <= max_pages:
            params: dict[str, Any] = {"per_page": per_page, "page": page}
            if after is not None:
                params["after"] = after
            if before is not None:
                params["before"] = before
            resp = requests.get(
                STRAVA_ACTIVITIES_URL,
                headers={"Authorization": f"Bearer {token}"},
                params=params,
                timeout=30,
            )
            if resp.status_code == 401:
                return None
            if resp.status_code != 200:
                raise HTTPException(status_code=resp.status_code, detail=resp.text)
            batch = resp.json()
            if not isinstance(batch, list):
                break
            all_items.extend(batch)
            if len(batch) < per_page:
                break
            page += 1
        return all_items

    fetched = do_fetch(access_token)
    if fetched is None:
        refresh_token = token_data.get("refresh_token")
        if not refresh_token:
            raise HTTPException(status_code=401, detail="Access token expired and no refresh_token available.")
        token_data = refresh_access_token(refresh_token)
        fetched = do_fetch(token_data.get("access_token", ""))
        if fetched is None:
            raise HTTPException(status_code=401, detail="Failed to refresh Strava token.")
    return fetched


def load_calendar_items() -> list[dict[str, Any]]:
    if not CALENDAR_FILE.exists():
        items: list[dict[str, Any]] = []
        if PLANNED_FILE.exists():
            try:
                legacy = read_json_file(PLANNED_FILE, [])
                if isinstance(legacy, list):
                    for row in legacy:
                        items.append(
                            {
                                "id": row.get("id") or str(uuid4()),
                                "kind": "workout",
                                "workout_type": row.get("workout_type", "Other"),
                                "date": row.get("date"),
                                "title": row.get("title", "Untitled Workout"),
                                "duration_min": row.get("planned_duration_min", 0),
                                "distance_km": row.get("planned_distance_km", 0),
                                "intensity": row.get("planned_intensity", 6),
                                "description": row.get("description", ""),
                                "created_at": row.get("created_at")
                                or datetime.utcnow().isoformat(timespec="seconds") + "Z",
                            }
                        )
            except json.JSONDecodeError:
                items = []
        save_calendar_items(items)
        return items

    raw = read_json_file(CALENDAR_FILE, [])
    if isinstance(raw, list):
        return raw
    return []

def save_calendar_items(items: list[dict[str, Any]]) -> None:
    write_json_file(CALENDAR_FILE, items)

def load_pairs() -> list[dict[str, Any]]:
    raw = read_json_file(PAIRS_FILE, [])
    if isinstance(raw, list):
        return raw
    return []

def save_pairs(items: list[dict[str, Any]]) -> None:
    write_json_file(PAIRS_FILE, items)

def load_activity_overrides() -> dict[str, dict[str, Any]]:
    with get_db() as db:
        rows = db.execute("SELECT * FROM activity_overrides").fetchall()
    return {r["id"]: override_to_dict(r) for r in rows}


def save_activity_overrides(items: dict[str, dict[str, Any]]) -> None:
    with get_db() as db:
        for aid, override in items.items():
            _upsert_override(db, aid, override)


def load_imported_activities() -> list[dict[str, Any]]:
    with get_db() as db:
        rows = db.execute(
            "SELECT * FROM activities WHERE source = 'fit' AND hidden = 0"
        ).fetchall()
    return [row_to_activity(r) for r in rows]


def get_imported_activity(activity_id: str) -> dict[str, Any] | None:
    with get_db() as db:
        row = db.execute(
            "SELECT * FROM activities WHERE id = ?", (activity_id,)
        ).fetchone()
    return row_to_activity(row) if row else None


def get_imported_activity_raw(activity_id: str) -> sqlite3.Row | None:
    with get_db() as db:
        return db.execute(
            "SELECT * FROM activities WHERE id = ?", (activity_id,)
        ).fetchone()


def apply_parsed_fit_to_activity(item: dict[str, Any], parsed: dict[str, Any], file_id: str, filename: str) -> dict[str, Any]:
    summary = parsed.get("summary", {})
    laps = parsed.get("laps", []) if isinstance(parsed.get("laps"), list) else []
    lap_duration_s = 0.0
    for lap in laps:
        if isinstance(lap, dict):
            try:
                lap_duration_s += float(lap.get("duration_s") or 0)
            except (TypeError, ValueError):
                pass
    duration_s = float(summary.get("duration_s") or 0)
    if duration_s <= 0 and lap_duration_s > 0:
        duration_s = lap_duration_s
    if duration_s <= 0:
        series = parsed.get("series", [])
        if isinstance(series, list) and len(series) >= 2:
            try:
                start_ts = datetime.fromisoformat(str(series[0].get("timestamp")))
                end_ts = datetime.fromisoformat(str(series[-1].get("timestamp")))
                duration_s = max(0.0, (end_ts - start_ts).total_seconds())
            except Exception:
                pass

    distance_m = float(summary.get("distance_m") or 0)
    if distance_m <= 0 and laps:
        lap_dist = 0.0
        for lap in laps:
            if isinstance(lap, dict):
                try:
                    lap_dist += float(lap.get("distance_m") or 0)
                except (TypeError, ValueError):
                    pass
        if lap_dist > 0:
            distance_m = lap_dist

    item["fit_id"] = file_id
    item["fit_filename"] = Path(filename).name
    item["distance"] = float(distance_m or item.get("distance") or 0)
    item["moving_time"] = float(duration_s or item.get("moving_time") or 0)
    if summary.get("start"):
        item["start_date_local"] = str(summary.get("start"))
    sport = str(summary.get("sport") or item.get("type") or "Ride").title()
    item["type"] = sport
    item["if_value"] = summary.get("if")
    item["np_value"] = summary.get("normalized_power")
    item["tss_override"] = summary.get("tss")
    item["work_kj"] = summary.get("work_kj")
    item["calories"] = summary.get("calories")
    item["avg_speed"] = summary.get("avg_speed")
    item["avg_power"] = summary.get("avg_power")
    item["avg_hr"] = summary.get("avg_hr")
    item["min_hr"] = summary.get("min_hr")
    item["max_hr"] = summary.get("max_hr")
    item["min_power"] = summary.get("min_power")
    item["max_power"] = summary.get("max_power")
    item["elev_gain_m"] = summary.get("elev_gain_m")
    item["hr_tss"] = summary.get("hr_tss")
    return item


def default_settings() -> dict[str, Any]:
    return {
        "unit_system": "metric",
        "units": {"distance": "km", "elevation": "m"},
        "ftp": {
            "ride": None,
            "run": None,
            "swim": None,
            "row": None,
            "strength": None,
            "other": None,
        },
        "lthr": {"ride": None, "run": None, "row": None, "swim": None, "strength": None, "other": None, "global": None},
    }


def sanitize_ftp_value(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        n = float(value)
    except (TypeError, ValueError):
        return None
    if n <= 0:
        return None
    return max(50.0, min(600.0, n))


def load_settings() -> dict[str, Any]:
    raw = read_json_file(SETTINGS_FILE, {})
    settings = default_settings()
    if isinstance(raw, dict):
        unit_system = str(raw.get("unit_system", "")).strip().lower()
        if unit_system in {"metric", "imperial"}:
            settings["unit_system"] = unit_system
        units = raw.get("units", {})
        if isinstance(units, dict):
            if units.get("distance") in {"km", "mi", "m"}:
                settings["units"]["distance"] = units.get("distance")
            if units.get("elevation") in {"m", "ft"}:
                settings["units"]["elevation"] = units.get("elevation")
        ftp = raw.get("ftp", {})
        if isinstance(ftp, dict):
            for key in settings["ftp"].keys():
                settings["ftp"][key] = sanitize_ftp_value(ftp.get(key))
        lthr = raw.get("lthr", {})
        if isinstance(lthr, dict):
            for key in settings["lthr"].keys():
                settings["lthr"][key] = sanitize_lthr_value(lthr.get(key))
    return settings


def save_settings(settings: dict[str, Any]) -> dict[str, Any]:
    merged = default_settings()
    unit_system = str(settings.get("unit_system", "")).strip().lower()
    if unit_system in {"metric", "imperial"}:
        merged["unit_system"] = unit_system
    units = settings.get("units", {})
    if isinstance(units, dict):
        if units.get("distance") in {"km", "mi", "m"}:
            merged["units"]["distance"] = units.get("distance")
        if units.get("elevation") in {"m", "ft"}:
            merged["units"]["elevation"] = units.get("elevation")
    ftp = settings.get("ftp", {})
    if isinstance(ftp, dict):
        for key in merged["ftp"].keys():
            merged["ftp"][key] = sanitize_ftp_value(ftp.get(key))
    lthr = settings.get("lthr", {})
    if isinstance(lthr, dict):
        for key in merged["lthr"].keys():
            merged["lthr"][key] = sanitize_lthr_value(lthr.get(key))
    write_json_file(SETTINGS_FILE, merged)
    return merged


def sport_to_ftp_key(sport: str) -> str:
    s = str(sport or "").lower()
    if "ride" in s or "cycle" in s or "bike" in s:
        return "ride"
    if "run" in s or "walk" in s:
        return "run"
    if "swim" in s:
        return "swim"
    if "row" in s:
        return "row"
    if "strength" in s or "weight" in s:
        return "strength"
    return "other"


def _as_float(v: Any) -> float | None:
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _normalize_comments_feed(raw_feed: Any) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    if not isinstance(raw_feed, list):
        return out
    for entry in raw_feed:
        if isinstance(entry, dict):
            text = str(entry.get("text") or entry.get("comment") or entry.get("body") or "").strip()
            if not text:
                continue
            author = str(
                entry.get("author")
                or entry.get("name")
                or entry.get("commenterName")
                or entry.get("user")
                or "Athlete"
            ).strip() or "Athlete"
            at = str(
                entry.get("at")
                or entry.get("created_at")
                or entry.get("dateCreated")
                or entry.get("timestamp")
                or ""
            ).strip()
            out.append({"author": author, "at": at, "text": text})
            continue
        text = str(entry or "").strip()
        if not text:
            continue
        out.append({"author": "Athlete", "at": "", "text": text})
    return out


def _iso(v: Any) -> str | None:
    if isinstance(v, datetime):
        return v.isoformat()
    if v is None:
        return None
    text = str(v).strip()
    return text or None


def _normalize_time_of_day(raw: Any) -> str:
    text = str(raw or "").strip()
    if not text:
        return ""
    hhmm = text[:5]
    if len(hhmm) == 5 and hhmm[2] == ":":
        try:
            h = int(hhmm[:2])
            m = int(hhmm[3:5])
            if 0 <= h <= 23 and 0 <= m <= 59:
                return f"{h:02d}:{m:02d}"
        except (TypeError, ValueError):
            pass
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return f"{dt.hour:02d}:{dt.minute:02d}"
    except ValueError:
        return ""


@lru_cache(maxsize=8)
def _tp_export_workout_file_index(filename: str) -> dict[str, Path]:
    out: dict[str, Path] = {}
    for root in TP_EXPORT_WORKOUT_ROOTS:
        if not root.exists():
            continue
        for path in root.glob(f"*/{filename}"):
            workout_id = path.parent.name.strip()
            if workout_id and workout_id not in out:
                out[workout_id] = path
    return out


def _tp_export_workout_file(workout_id: str, filename: str) -> Path | None:
    return _tp_export_workout_file_index(filename).get(str(workout_id))


@lru_cache(maxsize=1)
def _tp_export_start_time_map() -> dict[str, str]:
    out: dict[str, str] = {}
    for workout_id, path in _tp_export_workout_file_index("workout.json").items():
        try:
            raw = json.loads(path.read_text())
        except (OSError, json.JSONDecodeError):
            continue
        start_raw = str(raw.get("startTime") or "").strip()
        if not start_raw:
            continue
        try:
            dt = datetime.fromisoformat(start_raw.replace("Z", "+00:00"))
            out[workout_id] = dt.strftime("%H:%M:%S")
        except ValueError:
            continue
    return out


def _merge_tp_start_time(current_start: Any, workout_id: str) -> str | None:
    tp_time = _tp_export_start_time_map().get(str(workout_id))
    if not tp_time:
        return _iso(current_start)

    current_iso = _iso(current_start)
    if not current_iso:
        return None
    try:
        cur_dt = datetime.fromisoformat(current_iso.replace("Z", "+00:00"))
    except ValueError:
        return current_iso

    # Imported rows used 08:00 as a placeholder. Replace only placeholders.
    if not ((cur_dt.hour == 8 and cur_dt.minute == 0 and cur_dt.second == 0) or (cur_dt.hour == 0 and cur_dt.minute == 0 and cur_dt.second == 0)):
        return current_iso
    return f"{cur_dt.date().isoformat()}T{tp_time}"


def _mean(values: list[float]) -> float | None:
    if not values:
        return None
    return sum(values) / len(values)


def _max(values: list[float]) -> float | None:
    if not values:
        return None
    return max(values)


def sanitize_lthr_value(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        n = float(value)
    except (TypeError, ValueError):
        return None
    if n <= 0:
        return None
    return max(100.0, min(220.0, n))


# Coggan HR zone upper-bound percentages of LTHR (Z1–Z4; Z5 = above last bound)
_HR_ZONE_BOUNDS = [68.0, 84.0, 95.0, 106.0]
_HR_ZONE_RATES = [30.0, 55.0, 70.0, 90.0, 110.0]  # TSS/hr for zones 1–5


def _hr_tss(points: list[dict[str, Any]], lthr: float) -> float | None:
    """Calculate hrTSS from per-second HR series using Coggan zones."""
    if not lthr or lthr <= 0:
        return None
    time_in_zone = [0.0] * 5
    for p in points:
        hr = _as_float(p.get("heart_rate"))
        if hr is None or hr <= 0:
            continue
        pct = hr / lthr * 100.0
        zone = next((i for i, bound in enumerate(_HR_ZONE_BOUNDS) if pct < bound), 4)
        time_in_zone[zone] += 1.0  # one record ≈ one second
    total = sum(t / 3600.0 * _HR_ZONE_RATES[z] for z, t in enumerate(time_in_zone))
    return total if total > 0 else None


def _normalized_power(points: list[dict[str, Any]]) -> float | None:
    samples: list[tuple[float, float]] = []
    for row in points:
        p = _as_float(row.get("power"))
        ts = row.get("timestamp")
        if p is None or p <= 0 or not isinstance(ts, str):
            continue
        try:
            t = datetime.fromisoformat(ts).timestamp()
        except ValueError:
            continue
        samples.append((t, p))
    if not samples:
        return None
    window: list[tuple[float, float]] = []
    rolling_avgs: list[float] = []
    sum_power = 0.0
    for t, p in samples:
        window.append((t, p))
        sum_power += p
        while window and (t - window[0][0]) > 30.0:
            sum_power -= window[0][1]
            window.pop(0)
        if window:
            rolling_avgs.append(sum_power / len(window))
    if not rolling_avgs:
        return None
    mean_p4 = sum(v ** 4 for v in rolling_avgs) / len(rolling_avgs)
    return mean_p4 ** 0.25


def parse_fit_file_to_json(path: Path, settings: dict[str, Any] | None = None) -> dict[str, Any]:
    with path.open("rb") as handle:
        return parse_fit_stream_to_json(handle, settings=settings)


def parse_fit_bytes_to_json(content: bytes, settings: dict[str, Any] | None = None) -> dict[str, Any]:
    return parse_fit_stream_to_json(io.BytesIO(content), settings=settings)


def parse_fit_stream_to_json(stream: Any, settings: dict[str, Any] | None = None) -> dict[str, Any]:
    # Analysis pipeline source of truth:
    # records -> chart series points, laps -> lap table and lap-range selection.
    fit = FitFile(stream)
    points: list[dict[str, Any]] = []
    laps: list[dict[str, Any]] = []
    session_values: dict[str, Any] = {}
    sport = "Ride"

    for msg in fit.get_messages():
        vals = msg.get_values()
        name = msg.name
        if name == "record":
            ts = vals.get("timestamp")
            if not isinstance(ts, datetime):
                continue
            row = {
                "timestamp": ts.isoformat(),
                "heart_rate": _as_float(vals.get("heart_rate")),
                "speed": _as_float(vals.get("speed")),
                "distance": _as_float(vals.get("distance")),
                "cadence": _as_float(vals.get("cadence")),
                "power": _as_float(vals.get("power")),
                "altitude": _as_float(vals.get("altitude")),
            }
            points.append(row)
        elif name == "lap":
            start_ts = vals.get("start_time") or vals.get("timestamp")
            start_iso = _iso(start_ts)
            dur_s = _as_float(vals.get("total_timer_time")) or _as_float(vals.get("total_elapsed_time")) or 0.0
            end_iso = None
            if isinstance(start_ts, datetime):
                end_iso = (start_ts.timestamp() + dur_s)
                end_iso = datetime.fromtimestamp(end_iso).isoformat()
            laps.append(
                {
                    "name": f"Lap {len(laps) + 1}",
                    "start": start_iso,
                    "end": end_iso,
                    "duration_s": dur_s if dur_s > 0 else None,
                    "distance_m": _as_float(vals.get("total_distance")),
                    "avg_hr": _as_float(vals.get("avg_heart_rate")),
                    "max_hr": _as_float(vals.get("max_heart_rate")),
                    "avg_speed": _as_float(vals.get("avg_speed")),
                    "max_speed": _as_float(vals.get("max_speed")),
                    "avg_power": _as_float(vals.get("avg_power")),
                    "max_power": _as_float(vals.get("max_power")),
                    "normalized_power": _as_float(vals.get("normalized_power")),
                    "avg_cadence": _as_float(vals.get("avg_cadence")),
                    "max_cadence": _as_float(vals.get("max_cadence")),
                    "moving_duration_s": _as_float(vals.get("total_timer_time")),
                    "work_kj": (_as_float(vals.get("total_work")) / 1000.0) if _as_float(vals.get("total_work")) else None,
                    "calories": _as_float(vals.get("total_calories")),
                }
            )
        elif name == "session":
            session_values = vals
            s = str(vals.get("sport") or "").strip()
            if s:
                sport = s.title()
        elif name == "sport":
            s = str(vals.get("sport") or "").strip()
            if s:
                sport = s.title()

    if not points:
        raise HTTPException(status_code=400, detail="No record points found in FIT file.")

    first_ts = datetime.fromisoformat(points[0]["timestamp"])
    last_ts = datetime.fromisoformat(points[-1]["timestamp"])
    duration_s = max(1.0, (last_ts - first_ts).total_seconds())

    distances = [p["distance"] for p in points if p.get("distance") is not None]
    distance_m = 0.0
    if distances:
        distance_m = max(0.0, distances[-1] - distances[0]) if len(distances) > 1 else max(0.0, distances[0])
    session_distance = _as_float(session_values.get("total_distance"))
    if session_distance and session_distance > 0:
        distance_m = session_distance

    hr_values = [p["heart_rate"] for p in points if p.get("heart_rate") is not None]
    speed_values = [p["speed"] for p in points if p.get("speed") is not None]
    power_values = [p["power"] for p in points if p.get("power") is not None]
    cadence_values = [p["cadence"] for p in points if p.get("cadence") is not None]
    alt_values = [p["altitude"] for p in points if p.get("altitude") is not None]

    avg_speed = _mean(speed_values)
    max_speed = _max(speed_values)
    session_timer = _as_float(session_values.get("total_timer_time")) or _as_float(session_values.get("total_elapsed_time"))
    if session_timer and session_timer > 0:
        duration_s = session_timer

    if not laps:
        laps.append(
            {
                "name": "Lap 1",
                "start": first_ts.isoformat(),
                "end": last_ts.isoformat(),
                "duration_s": duration_s,
                "distance_m": distance_m,
                "avg_hr": _mean(hr_values),
                "max_hr": _max(hr_values),
                "avg_speed": avg_speed,
                "max_speed": max_speed,
                "avg_power": _mean(power_values),
                "max_power": _max(power_values),
                "avg_cadence": _mean(cadence_values),
                "max_cadence": _max(cadence_values),
            }
        )

    ftp_key = sport_to_ftp_key(sport)
    ftp_value = None
    lthr_value = None
    if settings:
        ftp_value = sanitize_ftp_value((settings.get("ftp") or {}).get(ftp_key))
        lthr_raw = settings.get("lthr") or {}
        lthr_value = sanitize_lthr_value(lthr_raw.get(ftp_key) or lthr_raw.get("global"))
    np_value = _normalized_power(points)
    if_value = None
    if ftp_value and np_value and np_value > 0:
        if_value = np_value / ftp_value
    tss_value = None
    if if_value and np_value and ftp_value and duration_s > 0:
        tss_value = (duration_s * np_value * if_value) / (ftp_value * 3600.0) * 100.0
    hr_tss_value = _hr_tss(points, lthr_value) if lthr_value else None

    return {
        "summary": {
            "start": first_ts.isoformat(),
            "end": last_ts.isoformat(),
            "duration_s": duration_s,
            "distance_m": distance_m,
            "avg_hr": _as_float(session_values.get("avg_heart_rate")) or _mean(hr_values),
            "max_hr": _as_float(session_values.get("max_heart_rate")) or _max(hr_values),
            "avg_speed": _as_float(session_values.get("avg_speed")) or avg_speed,
            "max_speed": _as_float(session_values.get("max_speed")) or max_speed,
            "avg_power": _as_float(session_values.get("avg_power")) or _mean(power_values),
            "max_power": _as_float(session_values.get("max_power")) or _max(power_values),
            "avg_cadence": _as_float(session_values.get("avg_cadence")) or _mean(cadence_values),
            "max_cadence": _as_float(session_values.get("max_cadence")) or _max(cadence_values),
            "elev_gain_m": _as_float(session_values.get("total_ascent")),
            "work_kj": (_as_float(session_values.get("total_work")) / 1000.0) if _as_float(session_values.get("total_work")) else None,
            "calories": _as_float(session_values.get("total_calories")),
            "sport": sport,
            "sport_key": ftp_key,
            "ftp": ftp_value,
            "lthr": lthr_value,
            "if": if_value,
            "tss": tss_value,
            "hr_tss": hr_tss_value,
            "normalized_power": np_value,
        },
        "series": points,
        "laps": laps,
    }


def save_fit_parsed(fit_id: str, data: dict[str, Any]) -> None:
    with get_db() as db:
        db.execute(
            "UPDATE activities SET fit_parsed_json = ? WHERE fit_id = ?",
            (json.dumps(data), fit_id),
        )


def _tp_channel_index(channel_set: Any) -> dict[str, int]:
    out: dict[str, int] = {}
    if not isinstance(channel_set, list):
        return out
    for idx, name in enumerate(channel_set):
        key = str(name or "")
        if not key:
            continue
        out[key] = idx
        out[key.lower()] = idx
    return out


def _tp_channel_value(values: list[Any], idx_map: dict[str, int], *names: str) -> float | None:
    for name in names:
        idx = idx_map.get(name)
        if idx is None:
            idx = idx_map.get(name.lower())
        if idx is None or idx < 0 or idx >= len(values):
            continue
        value = _as_float(values[idx])
        if value is not None:
            return value
    return None


def _tp_export_laps_for_workout(workout_id: str, start_dt: datetime) -> list[dict[str, Any]]:
    path = _tp_export_workout_file(str(workout_id), "detaildata.json")
    if not path or not path.exists():
        return []
    try:
        raw = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError):
        return []
    laps_stats = raw.get("lapsStats")
    if not isinstance(laps_stats, list):
        return []

    laps: list[dict[str, Any]] = []
    for idx, lap in enumerate(laps_stats):
        if not isinstance(lap, dict):
            continue
        begin_ms = _as_float(lap.get("begin"))
        end_ms = _as_float(lap.get("end"))
        elapsed_ms = _as_float(lap.get("elapsedTime"))
        if begin_ms is None:
            begin_ms = 0.0
        if end_ms is None and elapsed_ms is not None:
            end_ms = begin_ms + elapsed_ms
        if end_ms is None:
            continue
        start_iso = (start_dt + timedelta(milliseconds=max(0.0, begin_ms))).isoformat()
        end_iso = (start_dt + timedelta(milliseconds=max(begin_ms, end_ms))).isoformat()
        duration_s = max(0.0, (end_ms - begin_ms) / 1000.0)
        laps.append(
            {
                "name": str(lap.get("name") or f"Lap {idx + 1}"),
                "start": start_iso,
                "end": end_iso,
                "duration_s": duration_s if duration_s > 0 else None,
                "distance_m": _as_float(lap.get("distance")),
                "avg_hr": _as_float(lap.get("averageHeartRate")),
                "max_hr": _as_float(lap.get("maxHeartRate")),
                "avg_speed": _as_float(lap.get("averageSpeed")),
                "max_speed": _as_float(lap.get("maxSpeed")),
                "avg_power": _as_float(lap.get("averagePower")),
                "max_power": _as_float(lap.get("maxPower")),
                "normalized_power": _as_float(lap.get("normalizedPowerActual")),
                "avg_cadence": _as_float(lap.get("averageCadence")),
                "max_cadence": _as_float(lap.get("maxCadence")),
                "moving_duration_s": duration_s if duration_s > 0 else None,
                "work_kj": _as_float(lap.get("work")),
                "calories": _as_float(lap.get("calories")),
            }
        )
    return laps


def _tp_export_start_dt(workout_id: str, fallback: datetime) -> datetime:
    path = _tp_export_workout_file(str(workout_id), "workout.json")
    if not path or not path.exists():
        return fallback
    try:
        raw = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError):
        return fallback
    start_raw = str(raw.get("startTime") or "").strip()
    if not start_raw:
        return fallback
    try:
        return datetime.fromisoformat(start_raw.replace("Z", "+00:00"))
    except ValueError:
        return fallback


def _apply_tp_lap_timing(parsed: dict[str, Any], fit_id: str) -> dict[str, Any]:
    if not isinstance(parsed, dict):
        return parsed
    summary_raw = parsed.get("summary")
    summary = summary_raw if isinstance(summary_raw, dict) else {}
    fallback_start = datetime.utcnow()
    start_raw = _iso(summary.get("start"))
    if start_raw:
        try:
            fallback_start = datetime.fromisoformat(start_raw.replace("Z", "+00:00"))
        except ValueError:
            pass
    start_dt = _tp_export_start_dt(fit_id, fallback_start)
    tp_laps = _tp_export_laps_for_workout(fit_id, start_dt)
    if not tp_laps:
        return parsed

    out = dict(parsed)
    out["laps"] = tp_laps

    s = dict(summary)
    first_lap_start = tp_laps[0].get("start")
    last_lap_end = tp_laps[-1].get("end")
    if first_lap_start:
        s["start"] = first_lap_start
    if last_lap_end:
        s["end"] = last_lap_end
    total_lap_s = 0.0
    for lap in tp_laps:
        dur = _as_float((lap or {}).get("duration_s"))
        if dur and dur > 0:
            total_lap_s += dur
    if total_lap_s > 0:
        s["duration_s"] = total_lap_s
    out["summary"] = s
    return out


def _build_fit_from_tp_stream(fit_id: str) -> dict[str, Any]:
    try:
        with get_db() as db:
            activity_row = db.execute(
                """
                SELECT start_date_local, type, distance, moving_time, if_value, tss_override,
                       np_value, hr_tss, work_kj, calories, avg_speed, avg_power, avg_hr,
                       max_hr, max_power
                FROM activities
                WHERE fit_id = ?
                """,
                (fit_id,),
            ).fetchone()
            stream_row = db.execute(
                """
                SELECT channel_set_json, samples_gzip, encoding
                FROM tp_streams
                WHERE workout_id = ?
                """,
                (fit_id,),
            ).fetchone()
    except sqlite3.OperationalError as err:
        raise HTTPException(status_code=404, detail=f"TP stream table unavailable: {err}") from err

    if not activity_row or not stream_row:
        raise HTTPException(status_code=404, detail="Parsed FIT data not found.")

    raw_blob = stream_row["samples_gzip"]
    if raw_blob is None:
        raise HTTPException(status_code=404, detail="Parsed FIT data not found.")

    channel_set_raw = stream_row["channel_set_json"]
    try:
        channel_set = json.loads(channel_set_raw) if channel_set_raw else []
    except json.JSONDecodeError:
        channel_set = []

    raw_bytes = bytes(raw_blob)
    encoding = str(stream_row["encoding"] or "")
    try:
        payload_bytes = gzip.decompress(raw_bytes) if "gzip" in encoding else raw_bytes
        samples = json.loads(payload_bytes.decode("utf-8"))
    except Exception as err:
        raise HTTPException(status_code=500, detail=f"Corrupted TP stream payload: {err}") from err

    if not isinstance(samples, list) or not samples:
        raise HTTPException(status_code=404, detail="No TP stream samples found.")

    start_raw = _merge_tp_start_time(_iso(activity_row["start_date_local"]) or datetime.utcnow().isoformat(), fit_id) or datetime.utcnow().isoformat()
    try:
        start_dt = datetime.fromisoformat(start_raw.replace("Z", "+00:00"))
    except ValueError:
        start_dt = datetime.utcnow()

    idx_map = _tp_channel_index(channel_set)
    points: list[dict[str, Any]] = []
    for row in samples:
        if not isinstance(row, dict):
            continue
        ms = _as_float(row.get("ms"))
        values = row.get("values")
        if ms is None or not isinstance(values, list):
            continue
        ts = (start_dt + timedelta(milliseconds=ms)).isoformat()
        point: dict[str, Any] = {"timestamp": ts}

        hr = _tp_channel_value(values, idx_map, "heartRate", "heart_rate", "heartrate")
        speed = _tp_channel_value(values, idx_map, "speed")
        distance = _tp_channel_value(values, idx_map, "distance")
        cadence = _tp_channel_value(values, idx_map, "cadence")
        power = _tp_channel_value(values, idx_map, "power")
        lat = _tp_channel_value(values, idx_map, "positionLat", "lat", "latitude")
        lng = _tp_channel_value(values, idx_map, "positionLong", "positionLng", "lng", "longitude")

        if hr is not None:
            point["heart_rate"] = hr
        if speed is not None:
            point["speed"] = speed
        if distance is not None:
            point["distance"] = distance
        if cadence is not None:
            point["cadence"] = cadence
        if power is not None:
            point["power"] = power
        if lat is not None:
            point["lat"] = lat
        if lng is not None:
            point["lng"] = lng

        if len(point) > 1:
            points.append(point)

    if not points:
        raise HTTPException(status_code=404, detail="No supported TP stream channels available.")

    first_ts = datetime.fromisoformat(points[0]["timestamp"].replace("Z", "+00:00"))
    last_ts = datetime.fromisoformat(points[-1]["timestamp"].replace("Z", "+00:00"))
    duration_s = _as_float(activity_row["moving_time"]) or max(1.0, (last_ts - first_ts).total_seconds())
    distance_series = [p["distance"] for p in points if p.get("distance") is not None]
    distance_m = _as_float(activity_row["distance"]) or (distance_series[-1] if distance_series else 0.0)

    hr_values = [p["heart_rate"] for p in points if p.get("heart_rate") is not None]
    speed_values = [p["speed"] for p in points if p.get("speed") is not None]
    power_values = [p["power"] for p in points if p.get("power") is not None]
    cadence_values = [p["cadence"] for p in points if p.get("cadence") is not None]

    sport = str(activity_row["type"] or "Workout")
    sport_key = sport_to_ftp_key(sport)

    summary = {
        "start": first_ts.isoformat(),
        "end": last_ts.isoformat(),
        "duration_s": duration_s,
        "distance_m": distance_m,
        "avg_hr": _as_float(activity_row["avg_hr"]) or _mean(hr_values),
        "max_hr": _as_float(activity_row["max_hr"]) or _max(hr_values),
        "avg_speed": _as_float(activity_row["avg_speed"]) or _mean(speed_values),
        "max_speed": _max(speed_values),
        "avg_power": _as_float(activity_row["avg_power"]) or _mean(power_values),
        "max_power": _as_float(activity_row["max_power"]) or _max(power_values),
        "avg_cadence": _mean(cadence_values),
        "max_cadence": _max(cadence_values),
        "elev_gain_m": None,
        "work_kj": _as_float(activity_row["work_kj"]),
        "calories": _as_float(activity_row["calories"]),
        "sport": sport,
        "sport_key": sport_key,
        "ftp": None,
        "lthr": None,
        "if": _as_float(activity_row["if_value"]),
        "tss": _as_float(activity_row["tss_override"]),
        "hr_tss": _as_float(activity_row["hr_tss"]),
        "normalized_power": _as_float(activity_row["np_value"]),
    }
    laps = _tp_export_laps_for_workout(fit_id, start_dt)
    if not laps:
        laps = [
            {
                "name": "Lap 1",
                "start": first_ts.isoformat(),
                "end": last_ts.isoformat(),
                "duration_s": duration_s,
                "distance_m": distance_m,
                "avg_hr": summary["avg_hr"],
                "max_hr": summary["max_hr"],
                "avg_speed": summary["avg_speed"],
                "max_speed": summary["max_speed"],
                "avg_power": summary["avg_power"],
                "max_power": summary["max_power"],
                "avg_cadence": summary["avg_cadence"],
                "max_cadence": summary["max_cadence"],
            }
        ]
    has_gps = any(c in ("positionLat", "positionLong", "lat", "lng", "latitude", "longitude") for c in channel_set)
    return {"summary": summary, "series": points, "laps": laps, "has_gps": has_gps}


def load_fit_parsed(fit_id: str) -> dict[str, Any]:
    with get_db() as db:
        row = db.execute(
            "SELECT fit_parsed_json FROM activities WHERE fit_id = ?", (fit_id,)
        ).fetchone()
    if row and row["fit_parsed_json"]:
        try:
            parsed = json.loads(row["fit_parsed_json"])
            return _apply_tp_lap_timing(parsed, fit_id)
        except json.JSONDecodeError as err:
            raise HTTPException(status_code=500, detail="Corrupted FIT parsed data.") from err
    return _apply_tp_lap_timing(_build_fit_from_tp_stream(fit_id), fit_id)


def demo_activities() -> list[dict[str, Any]]:
    return []


def ensure_seed_calendar_items() -> None:
    return


def normalize_item(payload: dict[str, Any]) -> dict[str, Any]:
    kind = str(payload.get("kind", "")).strip().lower()
    if not kind:
        raise HTTPException(status_code=400, detail="kind is required.")

    date_str = str(payload.get("date", "")).strip()
    if not date_str:
        raise HTTPException(status_code=400, detail="date is required (YYYY-MM-DD).")

    try:
        parsed_date = date.fromisoformat(date_str)
    except ValueError as err:
        raise HTTPException(status_code=400, detail="Invalid date format, use YYYY-MM-DD.") from err

    title = str(payload.get("title", "")).strip()
    if not title:
        title = {
            "workout": "Untitled Workout",
            "event": "Untitled Event",
            "goal": "Untitled Goal",
            "note": "Untitled Note",
            "metrics": "Daily Metrics",
            "availability": "Availability",
        }.get(kind, "Untitled Item")

    description = str(payload.get("description", "")).strip()

    item: dict[str, Any] = {
        "id": str(uuid4()),
        "kind": kind,
        "date": parsed_date.isoformat(),
        "title": title,
        "description": description,
        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
    }

    if kind == "workout":
        workout_type = str(payload.get("workout_type", "Other")).strip() or "Other"
        start_time = _normalize_time_of_day(payload.get("start_time"))
        if not start_time:
            start_time = _normalize_time_of_day(payload.get("start_date_local"))
        try:
            duration = float(payload.get("duration_min", 0) or 0)
            distance = float(payload.get("distance_km", 0) or 0)
            distance_m = float(payload.get("distance_m", distance * 1000) or 0)
            elevation_m = float(payload.get("elevation_m", 0) or 0)
            intensity = float(payload.get("intensity", 6) or 6)
            completed_duration = float(payload.get("completed_duration_min", 0) or 0)
            completed_distance = float(payload.get("completed_distance_km", 0) or 0)
            completed_distance_m = float(payload.get("completed_distance_m", completed_distance * 1000) or 0)
            completed_elevation_m = float(payload.get("completed_elevation_m", 0) or 0)
            completed_tss = float(payload.get("completed_tss", 0) or 0)
            completed_if = float(payload.get("completed_if", 0) or 0)
            completed_np = float(payload.get("completed_np", 0) or 0)
            completed_work_kj = float(payload.get("completed_work_kj", 0) or 0)
            completed_calories = float(payload.get("completed_calories", 0) or 0)
            completed_avg_speed = float(payload.get("completed_avg_speed", 0) or 0)
            completed_hr_min = float(payload.get("completed_hr_min", 0) or 0)
            completed_hr_avg = float(payload.get("completed_hr_avg", 0) or 0)
            completed_hr_max = float(payload.get("completed_hr_max", 0) or 0)
            completed_power_min = float(payload.get("completed_power_min", 0) or 0)
            completed_power_avg = float(payload.get("completed_power_avg", 0) or 0)
            completed_power_max = float(payload.get("completed_power_max", 0) or 0)
            planned_if = float(payload.get("planned_if", 0) or 0)
            planned_tss = float(payload.get("planned_tss", 0) or 0)
            planned_avg_speed = float(payload.get("planned_avg_speed", 0) or 0)
            planned_calories = float(payload.get("planned_calories", 0) or 0)
            planned_work_kj = float(payload.get("planned_work_kj", 0) or 0)
        except (TypeError, ValueError) as err:
            raise HTTPException(status_code=400, detail="Workout values must be numeric.") from err

        item["workout_type"] = workout_type
        item["start_time"] = start_time
        item["duration_min"] = max(0.0, duration)
        item["distance_km"] = max(0.0, distance)
        item["distance_m"] = max(0.0, distance_m)
        item["elevation_m"] = max(0.0, elevation_m)
        d_unit = str(payload.get("distance_unit", "km"))
        e_unit = str(payload.get("elevation_unit", "m"))
        item["distance_unit"] = d_unit if d_unit in {"km", "mi", "m", "yd"} else "km"
        item["elevation_unit"] = e_unit if e_unit in {"m", "ft"} else "m"
        item["intensity"] = max(1.0, min(10.0, intensity))
        item["completed_duration_min"] = max(0.0, completed_duration)
        item["completed_distance_km"] = max(0.0, completed_distance)
        item["completed_distance_m"] = max(0.0, completed_distance_m)
        item["completed_elevation_m"] = max(0.0, completed_elevation_m)
        item["completed_tss"] = max(0.0, completed_tss)
        item["completed_if"] = max(0.0, completed_if)
        item["completed_np"] = max(0.0, completed_np)
        item["completed_work_kj"] = max(0.0, completed_work_kj)
        item["completed_calories"] = max(0.0, completed_calories)
        item["completed_avg_speed"] = max(0.0, completed_avg_speed)
        item["completed_hr_min"] = max(0.0, completed_hr_min)
        item["completed_hr_avg"] = max(0.0, completed_hr_avg)
        item["completed_hr_max"] = max(0.0, completed_hr_max)
        item["completed_power_min"] = max(0.0, completed_power_min)
        item["completed_power_avg"] = max(0.0, completed_power_avg)
        item["completed_power_max"] = max(0.0, completed_power_max)
        item["planned_if"] = max(0.0, planned_if)
        item["planned_tss"] = max(0.0, planned_tss)
        item["planned_avg_speed"] = max(0.0, planned_avg_speed)
        item["planned_calories"] = max(0.0, planned_calories)
        item["planned_work_kj"] = max(0.0, planned_work_kj)
        item["comments"] = str(payload.get("comments", "")).strip()
        raw_feed = payload.get("comments_feed", [])
        item["comments_feed"] = _normalize_comments_feed(raw_feed)
        feel = payload.get("feel")
        try:
            feel_val = int(feel) if feel is not None and str(feel).strip() else 0
        except (TypeError, ValueError):
            feel_val = 0
        try:
            rpe_val = int(payload.get("rpe", 0) or 0)
        except (TypeError, ValueError):
            rpe_val = 0
        item["feel"] = max(0, min(5, feel_val))
        item["rpe"] = max(0, min(10, rpe_val))
        raw_analysis = payload.get("analysis_edits", {})
        if isinstance(raw_analysis, dict):
            deleted = raw_analysis.get("deletedChannels", [])
            cuts = raw_analysis.get("cuts", [])
            item["analysis_edits"] = {
                "deletedChannels": [str(x) for x in deleted if isinstance(x, str)],
                "cuts": [
                    {"startSec": max(0.0, float(c.get("startSec", 0) or 0)), "endSec": max(0.0, float(c.get("endSec", 0) or 0))}
                    for c in cuts
                    if isinstance(c, dict)
                ],
            }
        else:
            item["analysis_edits"] = {"deletedChannels": [], "cuts": []}

    if kind == "event":
        item["event_type"] = str(payload.get("event_type", "Race")).strip() or "Race"
        raw_priority = str(payload.get("priority", "C")).strip().upper()
        item["priority"] = raw_priority if raw_priority in ("A", "B", "C") else "C"

    if kind == "goal":
        item["completed"] = bool(payload.get("completed", False))
        item["sort_order"] = int(payload.get("sort_order", 0))

    if kind == "availability":
        item["availability"] = str(payload.get("availability", "Unavailable")).strip() or "Unavailable"

    return item


@app.get("/", response_class=HTMLResponse)
def page() -> FileResponse:
    return FileResponse(path=str(TEMPLATES_DIR / "index.html"), media_type="text/html")


@app.get("/health")
def health() -> dict[str, bool]:
    return {"ok": True}


@app.get("/settings")
def get_settings() -> dict[str, Any]:
    settings = load_settings()
    if not SETTINGS_FILE.exists():
        save_settings(settings)
    return settings


@app.put("/settings")
def put_settings(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    return save_settings(payload)


@app.get("/strava-status")
def strava_status() -> dict:
    data = read_json_file(TOKEN_FILE, {})
    connected = bool(data.get("access_token"))
    athlete = data.get("athlete", {})
    return {
        "connected": connected,
        "athlete_name": f"{athlete.get('firstname', '')} {athlete.get('lastname', '')}".strip() if connected else None,
    }


@app.get("/connect")
def connect() -> RedirectResponse:
    global _pending_oauth_state
    client_id = os.getenv("STRAVA_CLIENT_ID")
    redirect_uri = os.getenv("STRAVA_REDIRECT_URI")
    if not client_id or not redirect_uri:
        raise HTTPException(status_code=500, detail="Missing STRAVA_CLIENT_ID or STRAVA_REDIRECT_URI.")
    state = secrets.token_urlsafe(24)
    _pending_oauth_state = state

    auth_url = (
        "https://www.strava.com/oauth/authorize"
        f"?client_id={client_id}"
        "&response_type=code"
        f"&redirect_uri={redirect_uri}"
        f"&state={state}"
        "&approval_prompt=auto"
        "&scope=read,activity:read"
    )
    return RedirectResponse(url=auth_url)


@app.get("/callback")
def callback(code: str = Query(...), state: str = Query(...)) -> RedirectResponse:
    global _pending_oauth_state
    client_id = os.getenv("STRAVA_CLIENT_ID")
    client_secret = os.getenv("STRAVA_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise HTTPException(status_code=500, detail="Missing STRAVA_CLIENT_ID or STRAVA_CLIENT_SECRET.")
    if not _pending_oauth_state or state != _pending_oauth_state:
        raise HTTPException(status_code=400, detail="Invalid OAuth state.")
    _pending_oauth_state = ""

    resp = requests.post(
        STRAVA_TOKEN_URL,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "code": code,
            "grant_type": "authorization_code",
        },
        timeout=30,
    )
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail=resp.text)

    token_data = resp.json()
    save_tokens(token_data)
    return RedirectResponse(url="/")


@app.get("/activities")
def activities(
    after: int | None = Query(default=None),
    before: int | None = Query(default=None),
) -> list[dict[str, Any]]:
    return fetch_activities(after=after, before=before)


@app.get("/ui/activities")
def ui_activities() -> list[dict[str, Any]]:
    ensure_seed_calendar_items()
    demo = demo_activities()
    imported = load_imported_activities()
    overrides = load_activity_overrides()

    try:
        live = fetch_activities()
        by_id = {str(x.get("id")): x for x in [*demo, *imported]}
        for row in live:
            by_id[str(row.get("id"))] = row
        merged = list(by_id.values())
    except HTTPException:
        merged = [*demo, *imported]

    override_fields = [
        "description", "comments", "comments_feed", "feel", "rpe",
        "tss_override", "if_value", "tss_source", "analysis_edits",
        "duration_min", "distance_km", "distance_m", "elevation_m",
        "distance_unit", "elevation_unit", "planned_tss", "planned_if",
        "planned_avg_speed", "planned_calories", "planned_work_kj",
        "completed_duration_min",
    ]

    out: list[dict[str, Any]] = []
    for row in merged:
        rid = str(row.get("id"))
        override = overrides.get(rid, {})
        if override.get("hidden"):
            continue
        updated = {**row}
        if override.get("date"):
            old = str(updated.get("start_date_local", ""))
            time_part = old[10:] if len(old) > 10 else "T08:00:00"
            updated["start_date_local"] = f"{override['date']}{time_part}"
        if override.get("title"):
            updated["name"] = str(override["title"])
        if override.get("type"):
            updated["type"] = str(override["type"])
        for k in override_fields:
            if k in override and override[k] is not None:
                updated[k] = override[k]
        merged_start = _merge_tp_start_time(updated.get("start_date_local"), rid)
        if merged_start:
            updated["start_date_local"] = merged_start
        out.append(updated)
    return out


@app.delete("/activities/{activity_id}")
def delete_activity_local(activity_id: str) -> dict[str, bool]:
    with get_db() as db:
        # Mark FIT-imported activity as hidden (don't DELETE so history is preserved)
        db.execute(
            "UPDATE activities SET hidden = 1 WHERE id = ?", (activity_id,)
        )
        # Upsert override to mark Strava activities hidden too
        existing = db.execute(
            "SELECT id FROM activity_overrides WHERE id = ?", (activity_id,)
        ).fetchone()
        if existing:
            db.execute(
                "UPDATE activity_overrides SET hidden = 1 WHERE id = ?", (activity_id,)
            )
        else:
            db.execute(
                "INSERT INTO activity_overrides (id, hidden) VALUES (?, 1)", (activity_id,)
            )

    pairs = load_pairs()
    pairs = [p for p in pairs if str(p.get("strava_id")) != activity_id]
    save_pairs(pairs)
    return {"ok": True}


@app.post("/import-fit")
async def import_fit(request: Request, filename: str = Query(default="workout.fit")) -> dict[str, Any]:
    ext = Path(filename).suffix.lower()
    if ext != ".fit":
        raise HTTPException(status_code=400, detail="Only .fit files are supported.")

    content = await request.body()
    if not content:
        raise HTTPException(status_code=400, detail="Empty file.")

    file_id = str(uuid4())
    settings = load_settings()
    try:
        parsed = parse_fit_bytes_to_json(content, settings=settings)
    except HTTPException:
        raise
    except Exception as err:
        raise HTTPException(status_code=400, detail=f"Failed to parse FIT: {err}") from err

    safe_name = Path(filename).name
    name = Path(safe_name).stem.replace("_", " ").replace("-", " ").strip() or "Imported Workout"
    item: dict[str, Any] = {
        "id": f"imported-{file_id}",
        "name": name.title(),
        "type": "Ride",
        "distance": 0,
        "moving_time": 0,
        "start_date_local": f"{date.today().isoformat()}T08:00:00",
        "description": "",
        "source": "fit",
        "fit_id": file_id,
        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
    }
    item = apply_parsed_fit_to_activity(item, parsed, file_id, safe_name)

    with get_db() as db:
        db.execute(
            _activity_insert_sql().replace("INSERT OR IGNORE", "INSERT OR REPLACE"),
            _activity_insert_params(
                item,
                content,
                json.dumps(parsed),
                item.get("comments_feed", []),
                item.get("analysis_edits", {}),
            ),
        )
    return item


@app.get("/fit/{fit_id}")
def get_fit_parsed(fit_id: str) -> dict[str, Any]:
    return load_fit_parsed(fit_id)


@app.post("/activities/{activity_id}/fit/upload")
async def upload_fit_for_activity(
    activity_id: str, request: Request, filename: str = Query(default="workout.fit")
) -> dict[str, Any]:
    content = await request.body()
    if not content:
        raise HTTPException(status_code=400, detail="Empty file.")
    if Path(filename).suffix.lower() != ".fit":
        raise HTTPException(status_code=400, detail="Only .fit files are supported.")

    item = get_imported_activity(activity_id)
    if item is None:
        raise HTTPException(status_code=404, detail="Activity not found.")

    file_id = str(uuid4())
    settings = load_settings()
    parsed = parse_fit_bytes_to_json(content, settings=settings)
    item = apply_parsed_fit_to_activity(item, parsed, file_id, filename)

    with get_db() as db:
        db.execute(
            """UPDATE activities SET
                fit_id=?, fit_filename=?, fit_data=?, fit_parsed_json=?,
                distance=?, moving_time=?, start_date_local=?, type=?,
                if_value=?, np_value=?, tss_override=?, work_kj=?, calories=?,
                avg_speed=?, avg_power=?, avg_hr=?, min_hr=?, max_hr=?,
                min_power=?, max_power=?, elev_gain_m=?, hr_tss=?
            WHERE id=?""",
            (
                file_id, Path(filename).name, content, json.dumps(parsed),
                item.get("distance"), item.get("moving_time"), item.get("start_date_local"),
                item.get("type"), item.get("if_value"), item.get("np_value"),
                item.get("tss_override"), item.get("work_kj"), item.get("calories"),
                item.get("avg_speed"), item.get("avg_power"), item.get("avg_hr"),
                item.get("min_hr"), item.get("max_hr"), item.get("min_power"),
                item.get("max_power"), item.get("elev_gain_m"), item.get("hr_tss"),
                activity_id,
            ),
        )
    return item


@app.post("/activities/{activity_id}/fit/recalculate")
def recalculate_fit_for_activity(activity_id: str) -> dict[str, Any]:
    with get_db() as db:
        row = db.execute(
            "SELECT * FROM activities WHERE id = ?", (activity_id,)
        ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Activity not found.")
    fit_id = str(row["fit_id"] or "").strip()
    if not fit_id:
        raise HTTPException(status_code=400, detail="No FIT attached.")
    fit_data = row["fit_data"]
    if not fit_data:
        raise HTTPException(status_code=404, detail="FIT file data missing.")

    parsed = parse_fit_bytes_to_json(bytes(fit_data), settings=load_settings())
    item = row_to_activity(row)
    filename = str(row["fit_filename"] or f"{fit_id}.fit")
    item = apply_parsed_fit_to_activity(item, parsed, fit_id, filename)

    with get_db() as db:
        db.execute(
            """UPDATE activities SET
                fit_parsed_json=?,
                distance=?, moving_time=?, start_date_local=?, type=?,
                if_value=?, np_value=?, tss_override=?, work_kj=?, calories=?,
                avg_speed=?, avg_power=?, avg_hr=?, min_hr=?, max_hr=?,
                min_power=?, max_power=?, elev_gain_m=?, hr_tss=?
            WHERE id=?""",
            (
                json.dumps(parsed),
                item.get("distance"), item.get("moving_time"), item.get("start_date_local"),
                item.get("type"), item.get("if_value"), item.get("np_value"),
                item.get("tss_override"), item.get("work_kj"), item.get("calories"),
                item.get("avg_speed"), item.get("avg_power"), item.get("avg_hr"),
                item.get("min_hr"), item.get("max_hr"), item.get("min_power"),
                item.get("max_power"), item.get("elev_gain_m"), item.get("hr_tss"),
                activity_id,
            ),
        )
    return item


@app.delete("/activities/{activity_id}/fit")
def delete_fit_for_activity(activity_id: str) -> dict[str, Any]:
    item = get_imported_activity(activity_id)
    if item is None:
        raise HTTPException(status_code=404, detail="Activity not found.")
    with get_db() as db:
        db.execute(
            """UPDATE activities SET
                fit_id=NULL, fit_filename=NULL, fit_data=NULL, fit_parsed_json=NULL,
                if_value=NULL, tss_override=NULL, avg_power=NULL,
                avg_hr=NULL, min_hr=NULL, max_hr=NULL,
                min_power=NULL, max_power=NULL, elev_gain_m=NULL
            WHERE id=?""",
            (activity_id,),
        )
    for key in ("fit_id", "fit_filename", "if_value", "tss_override",
                "avg_power", "avg_hr", "min_hr", "max_hr", "min_power", "max_power", "elev_gain_m"):
        item.pop(key, None)
    return item


@app.post("/activities/{activity_id}/fit/restore")
def restore_fit_for_activity(activity_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    item = get_imported_activity(activity_id)
    if item is None:
        raise HTTPException(status_code=404, detail="Activity not found.")

    fields = ["fit_id", "fit_filename", "distance", "moving_time", "avg_power",
              "avg_hr", "min_hr", "max_hr", "min_power", "max_power", "elev_gain_m",
              "if_value", "tss_override"]
    updates = {k: payload[k] for k in fields if k in payload}
    if updates:
        set_clause = ", ".join(f"{k}=?" for k in updates)
        with get_db() as db:
            db.execute(
                f"UPDATE activities SET {set_clause} WHERE id=?",
                (*updates.values(), activity_id),
            )
        item.update(updates)
    return item


@app.get("/activities/{activity_id}/fit/download")
async def download_fit_for_activity(activity_id: str):
    from fastapi.responses import Response
    with get_db() as db:
        row = db.execute(
            "SELECT fit_data, fit_filename, fit_id FROM activities WHERE id = ?", (activity_id,)
        ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Activity not found.")
    if not row["fit_data"]:
        raise HTTPException(status_code=404, detail="No FIT attached.")
    filename = str(row["fit_filename"] or f"{row['fit_id']}.fit")
    return Response(
        content=bytes(row["fit_data"]),
        media_type="application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.put("/activities/{activity_id}/meta")
def update_activity_meta(activity_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, bool]:
    updates: dict[str, Any] = {}

    if "description" in payload:
        updates["description"] = str(payload.get("description", ""))
    if "comments" in payload:
        updates["comments"] = str(payload.get("comments", ""))
    if "comments_feed" in payload:
        raw_feed = payload.get("comments_feed")
        if isinstance(raw_feed, list):
            updates["comments_feed"] = json.dumps(_normalize_comments_feed(raw_feed))
    if "title" in payload:
        updates["title"] = str(payload.get("title", "")).strip()
    if "type" in payload:
        updates["type"] = str(payload.get("type", "")).strip()
    if "start_date_local" in payload:
        raw_start = str(payload.get("start_date_local") or "").strip()
        if raw_start:
            try:
                _ = datetime.fromisoformat(raw_start.replace("Z", "+00:00"))
                updates["start_date_local"] = raw_start
            except ValueError:
                pass
    if "feel" in payload:
        try:
            updates["feel"] = max(0, min(5, int(payload.get("feel") or 0)))
        except (TypeError, ValueError):
            updates["feel"] = 0
    if "rpe" in payload:
        try:
            updates["rpe"] = max(0, min(10, int(payload.get("rpe") or 0)))
        except (TypeError, ValueError):
            updates["rpe"] = 0
    if "if_value" in payload:
        updates["if_value"] = _as_float(payload.get("if_value"))
    if "tss_override" in payload:
        updates["tss_override"] = _as_float(payload.get("tss_override"))
    for key in ("duration_min", "distance_km", "distance_m", "elevation_m",
                "planned_tss", "planned_if", "planned_avg_speed", "planned_calories",
                "planned_work_kj", "completed_duration_min"):
        if key in payload:
            updates[key] = _as_float(payload.get(key))
    if "distance_unit" in payload:
        unit = str(payload.get("distance_unit") or "").strip().lower()
        if unit in {"km", "mi", "m", "yd"}:
            updates["distance_unit"] = unit
    if "elevation_unit" in payload:
        unit = str(payload.get("elevation_unit") or "").strip().lower()
        if unit in {"m", "ft"}:
            updates["elevation_unit"] = unit
    if "tss_source" in payload:
        src = str(payload.get("tss_source") or "").strip()
        if src in {"power", "hr"}:
            updates["tss_source"] = src
        elif src == "":
            updates["tss_source"] = None
    if "analysis_edits" in payload:
        raw_analysis = payload.get("analysis_edits")
        if isinstance(raw_analysis, dict):
            deleted = raw_analysis.get("deletedChannels", [])
            cuts = raw_analysis.get("cuts", [])
            updates["analysis_edits"] = json.dumps({
                "deletedChannels": [str(x) for x in deleted if isinstance(x, str)],
                "cuts": [
                    {"startSec": max(0.0, float(c.get("startSec", 0) or 0)), "endSec": max(0.0, float(c.get("endSec", 0) or 0))}
                    for c in cuts if isinstance(c, dict)
                ],
            })

    if not updates:
        return {"ok": True}

    with get_db() as db:
        # Check if this is a FIT-imported activity
        is_fit = db.execute(
            "SELECT id FROM activities WHERE id = ?", (activity_id,)
        ).fetchone()

        if is_fit:
            activity_updates = dict(updates)
            if "title" in activity_updates:
                activity_updates["name"] = activity_updates.pop("title")
            set_clause = ", ".join(f"{k}=?" for k in activity_updates)
            db.execute(
                f"UPDATE activities SET {set_clause} WHERE id=?",
                (*activity_updates.values(), activity_id),
            )
            # Keep override table aligned so stale pair overrides do not mask user edits.
            if any(k in updates for k in ("type", "title", "start_date_local")):
                existing_override = db.execute(
                    "SELECT * FROM activity_overrides WHERE id = ?", (activity_id,)
                ).fetchone()
                if existing_override:
                    if "type" in updates:
                        db.execute(
                            "UPDATE activity_overrides SET type = ? WHERE id = ?",
                            (updates.get("type"), activity_id),
                        )
                    if "title" in updates:
                        db.execute(
                            "UPDATE activity_overrides SET title = ? WHERE id = ?",
                            (updates.get("title"), activity_id),
                        )
                    if "start_date_local" in updates:
                        date_part = str(updates.get("start_date_local") or "").split("T")[0].strip()
                        if date_part:
                            db.execute(
                                "UPDATE activity_overrides SET date = ? WHERE id = ?",
                                (date_part, activity_id),
                            )
        else:
            # Strava/external activity — use overrides table
            override_updates = dict(updates)
            if "start_date_local" in override_updates:
                date_part = str(override_updates.pop("start_date_local") or "").split("T")[0].strip()
                if date_part:
                    override_updates["date"] = date_part
            if not override_updates:
                return {"ok": True}
            existing = db.execute(
                "SELECT * FROM activity_overrides WHERE id = ?", (activity_id,)
            ).fetchone()
            if existing:
                set_clause = ", ".join(f"{k}=?" for k in override_updates)
                db.execute(
                    f"UPDATE activity_overrides SET {set_clause} WHERE id=?",
                    (*override_updates.values(), activity_id),
                )
            else:
                cols = ["id"] + list(override_updates.keys())
                placeholders = ", ".join("?" for _ in cols)
                db.execute(
                    f"INSERT INTO activity_overrides ({', '.join(cols)}) VALUES ({placeholders})",
                    (activity_id, *override_updates.values()),
                )
    return {"ok": True}


@app.get("/calendar-items")
def get_calendar_items() -> list[dict[str, Any]]:
    ensure_seed_calendar_items()
    items = load_calendar_items()
    return sorted(items, key=lambda x: (str(x.get("date", "")), str(x.get("created_at", ""))))


@app.post("/calendar-items")
def create_calendar_item(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    item = normalize_item(payload)
    items = load_calendar_items()
    items.append(item)
    save_calendar_items(items)
    return item


@app.put("/calendar-items/{item_id}")
def update_calendar_item(item_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    items = load_calendar_items()
    idx = next((i for i, row in enumerate(items) if row.get("id") == item_id), -1)
    if idx < 0:
        raise HTTPException(status_code=404, detail="Item not found.")

    existing = items[idx]
    merged = {**existing, **payload, "id": existing.get("id"), "created_at": existing.get("created_at")}
    normalized = normalize_item(merged)
    normalized["id"] = existing.get("id")
    normalized["created_at"] = existing.get("created_at")
    items[idx] = normalized
    save_calendar_items(items)
    return normalized


@app.put("/calendar-items/{item_id}/completed")
def update_calendar_item_completed(item_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    items = load_calendar_items()
    idx = next((i for i, row in enumerate(items) if row.get("id") == item_id), -1)
    if idx < 0:
        raise HTTPException(status_code=404, detail="Item not found.")
    item = items[idx]
    if item.get("kind") != "workout":
        raise HTTPException(status_code=400, detail="Only workout items support completed values.")

    def n(v: Any) -> float:
        try:
            return max(0.0, float(v or 0))
        except (TypeError, ValueError):
            return 0.0

    item["completed_duration_min"] = n(payload.get("completed_duration_min"))
    item["completed_distance_km"] = n(payload.get("completed_distance_km"))
    item["completed_distance_m"] = n(payload.get("completed_distance_m")) or (item["completed_distance_km"] * 1000)
    item["completed_elevation_m"] = n(payload.get("completed_elevation_m"))
    item["completed_tss"] = n(payload.get("completed_tss"))
    item["completed_if"] = n(payload.get("completed_if"))
    items[idx] = item
    save_calendar_items(items)
    return item


@app.delete("/calendar-items/{item_id}")
def delete_calendar_item(item_id: str) -> dict[str, bool]:
    items = load_calendar_items()
    target = next((row for row in items if row.get("id") == item_id), None)
    kept = [row for row in items if row.get("id") != item_id]
    if len(kept) == len(items):
        raise HTTPException(status_code=404, detail="Item not found.")
    save_calendar_items(kept)

    # Remove pair relationships involving this planned workout; if workout was paired,
    # also hide the linked completed activity.
    pairs = load_pairs()
    linked = [p for p in pairs if p.get("planned_id") == item_id]
    if target and target.get("kind") == "workout" and linked:
        overrides = load_activity_overrides()
        for link in linked:
            sid = str(link.get("strava_id", ""))
            if sid:
                current = overrides.get(sid, {})
                current["hidden"] = True
                overrides[sid] = current
        save_activity_overrides(overrides)
    pairs = [p for p in pairs if p.get("planned_id") != item_id]
    save_pairs(pairs)
    return {"ok": True}


@app.get("/pairs")
def get_pairs() -> list[dict[str, Any]]:
    return load_pairs()


@app.post("/pairs")
def create_pair(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    planned_id = str(payload.get("planned_id", "")).strip()
    strava_id = str(payload.get("strava_id", "")).strip()
    if not planned_id or not strava_id:
        raise HTTPException(status_code=400, detail="planned_id and strava_id are required.")

    planned_items = load_calendar_items()
    planned_item = next((row for row in planned_items if str(row.get("id")) == planned_id and row.get("kind") == "workout"), None)
    if not planned_item:
        raise HTTPException(status_code=404, detail="Planned workout not found.")

    def n(v: Any) -> float:
        try:
            return float(v or 0)
        except (TypeError, ValueError):
            return 0.0

    def sport_key(v: Any) -> str:
        t = str(v or "").strip().lower()
        if any(x in t for x in ("ride", "bike", "cycl")):
            return "ride"
        if any(x in t for x in ("run", "walk")):
            return "run"
        if "swim" in t:
            return "swim"
        if "row" in t:
            return "row"
        if any(x in t for x in ("strength", "weight")):
            return "strength"
        return "other"

    has_planned_content = (
        n(planned_item.get("duration_min")) > 0
        or n(planned_item.get("distance_km")) > 0
        or n(planned_item.get("planned_tss")) > 0
    )
    has_completed_on_planned = (
        n(planned_item.get("completed_duration_min")) > 0
        or n(planned_item.get("completed_distance_km")) > 0
        or n(planned_item.get("completed_tss")) > 0
        or n(planned_item.get("completed_if")) > 0
    )
    if not has_planned_content or has_completed_on_planned:
        raise HTTPException(status_code=400, detail="Pairing requires a planned-only workout.")

    completed_item = next((row for row in ui_activities() if str(row.get("id")) == strava_id), None)
    if not completed_item:
        raise HTTPException(status_code=404, detail="Completed workout not found.")
    planned_key = sport_key(planned_item.get("workout_type"))
    completed_key = sport_key(completed_item.get("type"))
    if planned_key != completed_key:
        raise HTTPException(status_code=400, detail="Pairing requires matching workout types.")
    pairs = load_pairs()
    pairs = [p for p in pairs if p.get("planned_id") != planned_id and p.get("strava_id") != strava_id]
    new_pair = {
        "id": str(uuid4()),
        "planned_id": planned_id,
        "strava_id": strava_id,
        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
    }
    pairs.append(new_pair)
    save_pairs(pairs)

    override_date = str(payload.get("override_date", "")).strip()
    override_title = str(payload.get("override_title", "")).strip()
    override_type = str(payload.get("override_type", "")).strip()
    if override_date or override_title or override_type:
        overrides = load_activity_overrides()
        current = overrides.get(strava_id, {})
        if override_title:
            current["title"] = override_title
        if override_type:
            current["type"] = override_type
        if override_date:
            try:
                _ = date.fromisoformat(override_date)
                current["date"] = override_date
            except ValueError:
                pass
        overrides[strava_id] = current
        save_activity_overrides(overrides)

    return new_pair


@app.delete("/pairs/{pair_id}")
def delete_pair(pair_id: str) -> dict[str, bool]:
    pairs = load_pairs()
    found = next((p for p in pairs if p.get("id") == pair_id), None)
    if not found:
        raise HTTPException(status_code=404, detail="Pair not found.")
    kept = [p for p in pairs if p.get("id") != pair_id]
    save_pairs(kept)

    strava_id = str(found.get("strava_id", "")).strip()
    planned_id = str(found.get("planned_id", "")).strip()
    planned_items = load_calendar_items()
    planned_item = next((row for row in planned_items if str(row.get("id")) == planned_id and row.get("kind") == "workout"), None)
    planned_date = str((planned_item or {}).get("date") or "").strip()
    if planned_item:
        reset_fields = (
            "completed_duration_min", "completed_distance_km", "completed_distance_m",
            "completed_elevation_m", "completed_tss", "completed_if", "completed_np",
            "completed_work_kj", "completed_calories", "completed_avg_speed",
            "completed_hr_min", "completed_hr_avg", "completed_hr_max",
            "completed_power_min", "completed_power_avg", "completed_power_max",
        )
        for field in reset_fields:
            planned_item[field] = 0
        for i, row in enumerate(planned_items):
            if str(row.get("id")) == planned_id:
                planned_items[i] = planned_item
                break
        save_calendar_items(planned_items)

    if strava_id:
        overrides = load_activity_overrides()
        current = overrides.get(strava_id, {})
        if planned_date:
            current["date"] = planned_date
        for key in ("date", "type", "title"):
            if key in current and key != "date":
                current.pop(key, None)
        if planned_date:
            current["date"] = planned_date
        if current:
            overrides[strava_id] = current
        elif strava_id in overrides:
            overrides.pop(strava_id, None)
        save_activity_overrides(overrides)

    return {"ok": True}


@app.get("/planned-workouts")
def get_planned_workouts() -> list[dict[str, Any]]:
    items = [i for i in load_calendar_items() if i.get("kind") == "workout"]
    return [
        {
            "id": i.get("id"),
            "date": i.get("date"),
            "start_time": i.get("start_time", ""),
            "workout_type": i.get("workout_type"),
            "title": i.get("title"),
            "planned_duration_min": i.get("duration_min", 0),
            "planned_distance_km": i.get("distance_km", 0),
            "planned_intensity": i.get("intensity", 6),
            "planned_if": i.get("planned_if", 0),
            "planned_tss": i.get("planned_tss", 0),
            "completed_duration_min": i.get("completed_duration_min", 0),
            "completed_distance_km": i.get("completed_distance_km", 0),
            "completed_tss": i.get("completed_tss", 0),
            "completed_if": i.get("completed_if", 0),
            "comments": i.get("comments", ""),
            "comments_feed": i.get("comments_feed", []),
            "feel": i.get("feel", 0),
            "rpe": i.get("rpe", 0),
            "description": i.get("description", ""),
            "created_at": i.get("created_at"),
        }
        for i in items
    ]

@app.post("/planned-workouts")
def create_planned_workout(payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    wrapped = {
        "kind": "workout",
        "date": payload.get("date"),
        "start_time": payload.get("start_time", ""),
        "workout_type": payload.get("workout_type", "Other"),
        "title": payload.get("title", "Untitled Workout"),
        "duration_min": payload.get("planned_duration_min", 0),
        "distance_km": payload.get("planned_distance_km", 0),
        "distance_m": payload.get("planned_distance_m", payload.get("planned_distance_km", 0) * 1000),
        "elevation_m": payload.get("planned_elevation_m", 0),
        "distance_unit": payload.get("distance_unit", "km"),
        "elevation_unit": payload.get("elevation_unit", "m"),
        "intensity": payload.get("planned_intensity", 6),
        "planned_if": payload.get("planned_if", 0),
        "planned_tss": payload.get("planned_tss", 0),
        "completed_duration_min": payload.get("completed_duration_min", 0),
        "completed_distance_km": payload.get("completed_distance_km", 0),
        "completed_distance_m": payload.get("completed_distance_m", payload.get("completed_distance_km", 0) * 1000),
        "completed_elevation_m": payload.get("completed_elevation_m", 0),
        "completed_tss": payload.get("completed_tss", 0),
        "completed_if": payload.get("completed_if", 0),
        "comments": payload.get("comments", ""),
        "comments_feed": payload.get("comments_feed", []),
        "feel": payload.get("feel", 0),
        "rpe": payload.get("rpe", 0),
        "description": payload.get("description", ""),
    }
    item = normalize_item(wrapped)
    items = load_calendar_items()
    items.append(item)
    save_calendar_items(items)
    return item
