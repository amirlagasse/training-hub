import json
import io
import os
import secrets
import threading
from datetime import date, datetime
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

TOKEN_FILE = Path("data/strava_tokens.json")
CALENDAR_FILE = Path("data/calendar_items.json")
PAIRS_FILE = Path("data/workout_pairs.json")
ACTIVITY_OVERRIDES_FILE = Path("data/activity_overrides.json")
IMPORTED_ACTIVITIES_FILE = Path("data/imported_activities.json")
FIT_PARSED_DIR = Path("data/fit_parsed")
SETTINGS_FILE = Path("data/settings.json")
PLANNED_FILE = Path("data/planned_workouts.json")
TEMPLATES_DIR = Path("app/templates")
STRAVA_TOKEN_URL = "https://www.strava.com/oauth/token"
STRAVA_ACTIVITIES_URL = "https://www.strava.com/api/v3/athlete/activities"
FILE_LOCK = threading.Lock()


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
    raw = read_json_file(ACTIVITY_OVERRIDES_FILE, {})
    if isinstance(raw, dict):
        return raw
    return {}


def save_activity_overrides(items: dict[str, dict[str, Any]]) -> None:
    write_json_file(ACTIVITY_OVERRIDES_FILE, items)


def load_imported_activities() -> list[dict[str, Any]]:
    raw = read_json_file(IMPORTED_ACTIVITIES_FILE, [])
    if isinstance(raw, list):
        return raw
    return []


def save_imported_activities(items: list[dict[str, Any]]) -> None:
    write_json_file(IMPORTED_ACTIVITIES_FILE, items)


def imported_activity_index(items: list[dict[str, Any]], activity_id: str) -> int:
    return next((i for i, row in enumerate(items) if str(row.get("id")) == activity_id), -1)


def apply_parsed_fit_to_activity(item: dict[str, Any], parsed: dict[str, Any], file_id: str, filename: str) -> dict[str, Any]:
    summary = parsed.get("summary", {})
    item["fit_id"] = file_id
    item["fit_filename"] = Path(filename).name
    item["distance"] = float(summary.get("distance_m") or item.get("distance") or 0)
    item["moving_time"] = float(summary.get("duration_s") or item.get("moving_time") or 0)
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
        "lthr": {"run": None, "ride": None},
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


def _iso(v: Any) -> str | None:
    if isinstance(v, datetime):
        return v.isoformat()
    if v is None:
        return None
    text = str(v).strip()
    return text or None


def _mean(values: list[float]) -> float | None:
    if not values:
        return None
    return sum(values) / len(values)


def _max(values: list[float]) -> float | None:
    if not values:
        return None
    return max(values)


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
    if settings:
        ftp_value = sanitize_ftp_value((settings.get("ftp") or {}).get(ftp_key))
    np_value = _normalized_power(points)
    if_value = None
    if ftp_value and np_value and np_value > 0:
        if_value = np_value / ftp_value
    tss_value = None
    if if_value and np_value and ftp_value and duration_s > 0:
        tss_value = (duration_s * np_value * if_value) / (ftp_value * 3600.0) * 100.0

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
            "if": if_value,
            "tss": tss_value,
            "normalized_power": np_value,
        },
        "series": points,
        "laps": laps,
    }


def save_fit_parsed(fit_id: str, data: dict[str, Any]) -> None:
    write_json_file(FIT_PARSED_DIR / f"{fit_id}.json", data)


def load_fit_parsed(fit_id: str) -> dict[str, Any]:
    path = FIT_PARSED_DIR / f"{fit_id}.json"
    data = read_json_file(path, {})
    if not data:
        raise HTTPException(status_code=404, detail="Parsed FIT data not found.")
    return data


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
        item["duration_min"] = max(0.0, duration)
        item["distance_km"] = max(0.0, distance)
        item["distance_m"] = max(0.0, distance_m)
        item["elevation_m"] = max(0.0, elevation_m)
        d_unit = str(payload.get("distance_unit", "km"))
        e_unit = str(payload.get("elevation_unit", "m"))
        item["distance_unit"] = d_unit if d_unit in {"km", "mi", "m"} else "km"
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
        if isinstance(raw_feed, list):
            item["comments_feed"] = [str(x).strip() for x in raw_feed if str(x).strip()]
        else:
            item["comments_feed"] = []
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

    if kind == "event":
        item["event_type"] = str(payload.get("event_type", "Race")).strip() or "Race"

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


@app.get("/connect")
def connect() -> RedirectResponse:
    client_id = os.getenv("STRAVA_CLIENT_ID")
    redirect_uri = os.getenv("STRAVA_REDIRECT_URI")
    if not client_id or not redirect_uri:
        raise HTTPException(status_code=500, detail="Missing STRAVA_CLIENT_ID or STRAVA_REDIRECT_URI.")
    state = secrets.token_urlsafe(24)

    auth_url = (
        "https://www.strava.com/oauth/authorize"
        f"?client_id={client_id}"
        "&response_type=code"
        f"&redirect_uri={redirect_uri}"
        f"&state={state}"
        "&approval_prompt=auto"
        "&scope=read,activity:read"
    )
    response = RedirectResponse(url=auth_url)
    response.set_cookie("strava_oauth_state", state, httponly=True, samesite="lax")
    return response


@app.get("/callback")
def callback(request: Request, code: str = Query(...), state: str = Query(...)) -> RedirectResponse:
    client_id = os.getenv("STRAVA_CLIENT_ID")
    client_secret = os.getenv("STRAVA_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise HTTPException(status_code=500, detail="Missing STRAVA_CLIENT_ID or STRAVA_CLIENT_SECRET.")
    expected_state = request.cookies.get("strava_oauth_state", "")
    if not expected_state or state != expected_state:
        raise HTTPException(status_code=400, detail="Invalid OAuth state.")

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
    response = RedirectResponse(url="/")
    response.delete_cookie("strava_oauth_state")
    return response


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

    out: list[dict[str, Any]] = []
    for row in merged:
        rid = str(row.get("id"))
        override = overrides.get(rid, {})
        updated = {**row}
        if "date" in override:
            old = str(updated.get("start_date_local", ""))
            time_part = old[10:] if len(old) > 10 else "T08:00:00"
            updated["start_date_local"] = f"{override['date']}{time_part}"
        if "title" in override:
            updated["name"] = str(override["title"])
        if "type" in override:
            updated["type"] = str(override["type"])
        for k in ["description", "comments", "comments_feed", "feel", "rpe", "tss_override", "if_value"]:
            if k in override:
                updated[k] = override[k]
        if override.get("hidden"):
            continue
        out.append(updated)
    return out


@app.delete("/activities/{activity_id}")
def delete_activity_local(activity_id: str) -> dict[str, bool]:
    imported = load_imported_activities()
    removed = next((a for a in imported if str(a.get("id")) == activity_id), None)
    filtered = [a for a in imported if str(a.get("id")) != activity_id]
    if len(filtered) != len(imported):
        save_imported_activities(filtered)
    if removed:
        fit_id = str(removed.get("fit_id") or "").strip()
        if fit_id:
            parsed_path = FIT_PARSED_DIR / f"{fit_id}.json"
            if parsed_path.exists():
                parsed_path.unlink()
            fit_path = Path("data/imports") / f"{fit_id}.fit"
            if fit_path.exists():
                fit_path.unlink()
    overrides = load_activity_overrides()
    current = overrides.get(activity_id, {})
    current["hidden"] = True
    overrides[activity_id] = current
    save_activity_overrides(overrides)

    # If this completed activity was paired, remove the pair.
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
    imports_dir = Path("data/imports")
    imports_dir.mkdir(parents=True, exist_ok=True)
    saved_path = imports_dir / f"{file_id}.fit"
    saved_path.write_bytes(content)
    settings = load_settings()
    try:
        parsed = parse_fit_bytes_to_json(content, settings=settings)
    except HTTPException:
        raise
    except Exception as err:
        raise HTTPException(status_code=400, detail=f"Failed to parse FIT: {err}") from err
    save_fit_parsed(file_id, parsed)

    safe_name = Path(filename).name
    name = Path(safe_name).stem.replace("_", " ").replace("-", " ").strip() or "Imported Workout"
    item = {
        "id": f"imported-{file_id}",
        "name": name.title(),
        "type": "Ride",
        "distance": 0,
        "moving_time": 0,
        "start_date_local": f"{date.today().isoformat()}T08:00:00",
        "description": "",
        "source": "fit",
        "fit_id": file_id,
    }
    item = apply_parsed_fit_to_activity(item, parsed, file_id, safe_name)
    imported = load_imported_activities()
    imported.append(item)
    save_imported_activities(imported)
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

    imported = load_imported_activities()
    idx = imported_activity_index(imported, activity_id)
    if idx < 0:
        raise HTTPException(status_code=404, detail="Activity not found.")

    file_id = str(uuid4())
    imports_dir = Path("data/imports")
    imports_dir.mkdir(parents=True, exist_ok=True)
    (imports_dir / f"{file_id}.fit").write_bytes(content)
    settings = load_settings()
    parsed = parse_fit_bytes_to_json(content, settings=settings)
    save_fit_parsed(file_id, parsed)

    imported[idx] = apply_parsed_fit_to_activity(imported[idx], parsed, file_id, filename)
    save_imported_activities(imported)
    return imported[idx]


@app.post("/activities/{activity_id}/fit/recalculate")
def recalculate_fit_for_activity(activity_id: str) -> dict[str, Any]:
    imported = load_imported_activities()
    idx = imported_activity_index(imported, activity_id)
    if idx < 0:
        raise HTTPException(status_code=404, detail="Activity not found.")
    fit_id = str(imported[idx].get("fit_id") or "").strip()
    if not fit_id:
        raise HTTPException(status_code=400, detail="No FIT attached.")
    fit_path = Path("data/imports") / f"{fit_id}.fit"
    if not fit_path.exists():
        raise HTTPException(status_code=404, detail="FIT file missing.")

    parsed = parse_fit_file_to_json(fit_path, settings=load_settings())
    save_fit_parsed(fit_id, parsed)
    filename = str(imported[idx].get("fit_filename") or f"{fit_id}.fit")
    imported[idx] = apply_parsed_fit_to_activity(imported[idx], parsed, fit_id, filename)
    save_imported_activities(imported)
    return imported[idx]


@app.delete("/activities/{activity_id}/fit")
def delete_fit_for_activity(activity_id: str) -> dict[str, Any]:
    imported = load_imported_activities()
    idx = imported_activity_index(imported, activity_id)
    if idx < 0:
        raise HTTPException(status_code=404, detail="Activity not found.")
    fit_id = str(imported[idx].get("fit_id") or "").strip()
    if fit_id:
        parsed_path = FIT_PARSED_DIR / f"{fit_id}.json"
        fit_path = Path("data/imports") / f"{fit_id}.fit"
        if parsed_path.exists():
            parsed_path.unlink()
        if fit_path.exists():
            fit_path.unlink()
    for key in [
        "fit_id",
        "fit_filename",
        "if_value",
        "tss_override",
        "avg_power",
        "avg_hr",
        "min_hr",
        "max_hr",
        "min_power",
        "max_power",
        "elev_gain_m",
    ]:
        imported[idx].pop(key, None)
    save_imported_activities(imported)
    return imported[idx]


@app.post("/activities/{activity_id}/fit/restore")
def restore_fit_for_activity(activity_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, Any]:
    imported = load_imported_activities()
    idx = imported_activity_index(imported, activity_id)
    if idx < 0:
        raise HTTPException(status_code=404, detail="Activity not found.")
    current_fit_id = str(imported[idx].get("fit_id") or "").strip()
    restore_fit_id = str(payload.get("fit_id") or "").strip()
    if current_fit_id and current_fit_id != restore_fit_id:
        current_fit_path = Path("data/imports") / f"{current_fit_id}.fit"
        current_parsed_path = FIT_PARSED_DIR / f"{current_fit_id}.json"
        if current_fit_path.exists():
            current_fit_path.unlink()
        if current_parsed_path.exists():
            current_parsed_path.unlink()
    for key in [
        "fit_id",
        "fit_filename",
        "distance",
        "moving_time",
        "avg_power",
        "avg_hr",
        "min_hr",
        "max_hr",
        "min_power",
        "max_power",
        "elev_gain_m",
        "if_value",
        "tss_override",
    ]:
        if key in payload:
            imported[idx][key] = payload.get(key)
    save_imported_activities(imported)
    return imported[idx]


@app.get("/activities/{activity_id}/fit/download")
def download_fit_for_activity(activity_id: str) -> FileResponse:
    imported = load_imported_activities()
    idx = imported_activity_index(imported, activity_id)
    if idx < 0:
        raise HTTPException(status_code=404, detail="Activity not found.")
    fit_id = str(imported[idx].get("fit_id") or "").strip()
    if not fit_id:
        raise HTTPException(status_code=404, detail="No FIT attached.")
    fit_path = Path("data/imports") / f"{fit_id}.fit"
    if not fit_path.exists():
        raise HTTPException(status_code=404, detail="FIT file missing.")
    filename = str(imported[idx].get("fit_filename") or f"{fit_id}.fit")
    return FileResponse(path=str(fit_path), filename=filename, media_type="application/octet-stream")


@app.put("/activities/{activity_id}/meta")
def update_activity_meta(activity_id: str, payload: dict[str, Any] = Body(...)) -> dict[str, bool]:
    overrides = load_activity_overrides()
    current = overrides.get(activity_id, {})
    if "description" in payload:
        current["description"] = str(payload.get("description", ""))
    if "comments" in payload:
        current["comments"] = str(payload.get("comments", ""))
    if "comments_feed" in payload:
        raw_feed = payload.get("comments_feed")
        if isinstance(raw_feed, list):
            current["comments_feed"] = [str(x).strip() for x in raw_feed if str(x).strip()]
    if "title" in payload:
        current["title"] = str(payload.get("title", "")).strip()
    if "type" in payload:
        current["type"] = str(payload.get("type", "")).strip()
    if "feel" in payload:
        try:
            current["feel"] = max(0, min(5, int(payload.get("feel") or 0)))
        except (TypeError, ValueError):
            current["feel"] = 0
    if "rpe" in payload:
        try:
            current["rpe"] = max(0, min(10, int(payload.get("rpe") or 0)))
        except (TypeError, ValueError):
            current["rpe"] = 0
    if "if_value" in payload:
        current["if_value"] = _as_float(payload.get("if_value"))
    if "tss_override" in payload:
        current["tss_override"] = _as_float(payload.get("tss_override"))
    overrides[activity_id] = current
    save_activity_overrides(overrides)
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
    if override_date:
        try:
            _ = date.fromisoformat(override_date)
            overrides = load_activity_overrides()
            current = overrides.get(strava_id, {})
            if override_title:
                current["title"] = override_title
            current["date"] = override_date
            overrides[strava_id] = current
            save_activity_overrides(overrides)
        except ValueError:
            pass

    return new_pair


@app.delete("/pairs/{pair_id}")
def delete_pair(pair_id: str) -> dict[str, bool]:
    pairs = load_pairs()
    kept = [p for p in pairs if p.get("id") != pair_id]
    if len(kept) == len(pairs):
        raise HTTPException(status_code=404, detail="Pair not found.")
    save_pairs(kept)
    return {"ok": True}


@app.get("/planned-workouts")
def get_planned_workouts() -> list[dict[str, Any]]:
    items = [i for i in load_calendar_items() if i.get("kind") == "workout"]
    return [
        {
            "id": i.get("id"),
            "date": i.get("date"),
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
