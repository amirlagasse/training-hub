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

TOKEN_FILE = Path("data/strava_tokens.json")
CALENDAR_FILE = Path("data/calendar_items.json")
PAIRS_FILE = Path("data/workout_pairs.json")
ACTIVITY_OVERRIDES_FILE = Path("data/activity_overrides.json")
IMPORTED_ACTIVITIES_FILE = Path("data/imported_activities.json")
FIT_PARSED_DIR = Path("data/fit_parsed")
SETTINGS_FILE = Path("data/settings.json")
PLANNED_FILE = Path("data/planned_workouts.json")
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
    item["tss_override"] = summary.get("tss")
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
                    "avg_cadence": _as_float(vals.get("avg_cadence")),
                    "max_cadence": _as_float(vals.get("max_cadence")),
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
    if_value = None
    if ftp_value and (_as_float(session_values.get("avg_power")) or _mean(power_values)):
        avg_p = _as_float(session_values.get("avg_power")) or _mean(power_values) or 0
        if avg_p > 0:
            if_value = avg_p / ftp_value
    tss_value = None
    if if_value and duration_s > 0:
        hours = duration_s / 3600.0
        tss_value = hours * if_value * if_value * 100.0

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
            "sport": sport,
            "sport_key": ftp_key,
            "ftp": ftp_value,
            "if": if_value,
            "tss": tss_value,
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
            planned_if = float(payload.get("planned_if", 0) or 0)
            planned_tss = float(payload.get("planned_tss", 0) or 0)
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
        item["planned_if"] = max(0.0, planned_if)
        item["planned_tss"] = max(0.0, planned_tss)
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
def page() -> str:
    return """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Training Freaks</title>
  <style>
    :root {
      --bg: #e8eef6;
      --panel: #ffffff;
      --line: #d7e0eb;
      --text: #172333;
      --muted: #6b7e93;
      --nav: #102947;
      --blue: #1e58d1;
      --pink: #ed4e95;
      --orange: #f06b23;
      --good: #148248;
      --planned: #b35d2a;
      --shadow: 0 12px 26px rgba(11, 25, 41, 0.08);
      --radius: 12px;
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      color: var(--text);
      /* TODO: add @font-face for TT Interphases Pro from /static/fonts when font files are available. */
      font-family: "TT Interphases Pro", system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
      background: linear-gradient(180deg, #f4f7fb 0%, var(--bg) 100%);
    }

    .top-nav {
      position: sticky;
      top: 0;
      z-index: 20;
      background: linear-gradient(180deg, #113154 0%, #0e2843 100%);
      border-bottom: 1px solid rgba(255,255,255,0.12);
      min-height: 45px;
      display: grid;
      grid-template-columns: 1fr auto 1fr;
      align-items: center;
      padding: 0 10px;
    }

    .brand {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      height: 30px;
    }

    .brand img {
      display: block;
      height: 24px;
      width: auto;
      object-fit: contain;
    }

    .brand .brand-icon {
      height: 28px;
    }

    .brand .brand-text {
      height: 16px;
    }

    .brand .brand-text.hidden {
      display: none;
    }

    .tabs {
      display: flex;
      gap: 8px;
      justify-content: center;
    }

    .tab {
      border: 0;
      border-radius: 8px;
      background: transparent;
      color: #d5e4f2;
      font-size: 13px;
      padding: 6px 12px;
      cursor: pointer;
    }

    .tab.active {
      background: rgba(255,255,255,0.16);
      color: #fff;
      font-weight: 700;
    }

    .nav-right {
      justify-self: end;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      color: #dbe8f4;
      font-size: 11px;
      font-weight: 700;
    }

    .nav-settings {
      border: 1px solid rgba(255, 255, 255, 0.25);
      background: rgba(255, 255, 255, 0.08);
      color: #eef4fb;
      border-radius: 8px;
      width: 24px;
      height: 24px;
      cursor: pointer;
    }

    .import-btn {
      border: 1px solid rgba(255, 255, 255, 0.28);
      background: rgba(255, 255, 255, 0.14);
      color: #eef4fb;
      border-radius: 8px;
      padding: 5px 10px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
    }

    .unit-btn {
      border: 1px solid rgba(255, 255, 255, 0.28);
      background: rgba(255, 255, 255, 0.14);
      color: #eef4fb;
      border-radius: 8px;
      padding: 5px 8px;
      font-size: 11px;
      font-weight: 700;
      cursor: pointer;
    }

    .main {
      padding: 14px;
    }

    .page-head {
      display: flex;
      align-items: center;
      justify-content: flex-start;
      margin-bottom: 10px;
      min-height: 32px;
    }
    .page-head.hidden { display: none; }

    h1 {
      margin: 0;
      font-size: 25px;
    }

    .view { display: none; }
    .view.active { display: block; }

    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 12px;
    }

    .panel-title {
      margin: 0 0 10px;
      color: #4f657d;
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.35px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
    }

    .plus-btn,
    .goal-btn {
      border: 1px solid #c7d5e6;
      background: #fff;
      color: #355578;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 700;
      padding: 4px 8px;
      cursor: pointer;
    }

    .home-grid {
      display: grid;
      grid-template-columns: 300px 1fr 380px;
      gap: 12px;
      align-items: start;
    }

    .stack {
      display: grid;
      gap: 10px;
    }

    .event-item,
    .goal-item {
      border: 1px solid #e0e8f3;
      background: #f8fbff;
      border-radius: 10px;
      padding: 9px;
      margin-bottom: 8px;
    }

    .event-item h4,
    .goal-item h4 {
      margin: 0 0 4px;
      font-size: 14px;
    }

    .event-item p,
    .goal-item p {
      margin: 0;
      color: #5a718a;
      font-size: 12px;
    }

    .list-item {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 8px;
      border-top: 1px solid #edf2f7;
      padding: 9px 0;
      width: 100%;
      background: transparent;
      border-left: 0;
      border-right: 0;
      border-bottom: 0;
      text-align: left;
      cursor: pointer;
    }

    .list-item:first-child { border-top: 0; padding-top: 2px; }

    .title { margin: 0 0 3px; font-size: 14px; font-weight: 700; }
    .meta { margin: 0; color: var(--muted); font-size: 12px; }

    .badge {
      align-self: center;
      border-radius: 999px;
      padding: 4px 7px;
      font-size: 10px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.25px;
    }

    .badge.done { background: #e9f8ef; color: var(--good); }
    .badge.planned { background: #fff1e9; color: var(--planned); }

    .metric-row {
      display: flex;
      gap: 8px;
      margin-bottom: 10px;
      flex-wrap: wrap;
    }

    .metric-chip {
      min-width: 88px;
      border-radius: 9px;
      color: #fff;
      text-align: center;
      padding: 6px 8px;
    }

    .metric-chip .num { display: block; font-size: 28px; font-weight: 800; line-height: 1; }
    .metric-chip .lbl { display: block; font-size: 10px; text-transform: uppercase; letter-spacing: 0.35px; }

    .chip-ctl { background: var(--blue); }
    .chip-atl { background: var(--pink); }
    .chip-tsb { background: var(--orange); }

    .explain {
      margin: 10px 0;
      padding: 10px;
      border: 1px solid #e0e8f2;
      border-radius: 10px;
      background: #f9fbfe;
      color: #5a7088;
      font-size: 12px;
      line-height: 1.4;
    }

    .spark-wrap { display: grid; gap: 8px; }

    .spark-box {
      border: 1px solid #e2ebf4;
      border-radius: 8px;
      background: #f8fbff;
      padding: 7px;
    }

    .spark-head {
      display: flex;
      justify-content: space-between;
      font-size: 11px;
      color: #57708a;
      margin-bottom: 6px;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }

    .sparkline {
      display: flex;
      gap: 2px;
      align-items: flex-end;
      height: 40px;
    }

    .sparkline span {
      display: block;
      width: 4px;
      flex: 1;
      border-radius: 2px 2px 0 0;
      min-height: 2px;
      background: #7d9fc8;
    }

    .calendar-wrap {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 10px;
    }

    .calendar-head {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }

    .cal-nav {
      display: flex;
      align-items: center;
      gap: 4px;
    }

    .btn {
      border: 0;
      border-radius: 10px;
      padding: 9px 13px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-block;
    }

    .btn.secondary { background: #fff; color: #284362; border: 1px solid var(--line); }
    .btn.primary { background: var(--blue); color: #fff; }

    .calendar-scroll {
      height: calc(100vh - 175px);
      overflow-y: auto;
      border: 1px solid #e2ebf4;
      border-radius: 10px;
      background: #f8fbff;
      padding: 8px;
      scroll-behavior: smooth;
    }

    .month {
      background: #fff;
      border: 1px solid #e3ebf5;
      border-radius: 10px;
      margin-bottom: 10px;
      padding: 9px;
    }

    .month.current-month {
      border-color: #b7ccee;
      box-shadow: 0 0 0 2px #e6efff inset;
    }

    .month-title {
      margin: 0 0 7px;
      font-size: 16px;
    }

    .dow,
    .week-row {
      display: grid;
      grid-template-columns: repeat(7, minmax(0, 1fr)) 240px;
      gap: 5px;
    }

    .dow {
      margin-bottom: 5px;
      color: #7087a0;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }

    .day {
      min-height: clamp(250px, 31vh, 420px);
      border: 1px solid #e7edf6;
      border-radius: 8px;
      background: #fdfefe;
      padding: 5px;
      display: flex;
      flex-direction: column;
      gap: 3px;
      position: relative;
      text-align: left;
    }

    .day.empty {
      border-color: transparent;
      background: transparent;
    }

    .day.today {
      border-color: #3d76df;
      box-shadow: 0 0 0 1px #dce9ff inset;
      background: #f3f8ff;
    }

    .quick-add {
      position: absolute;
      left: 6px;
      right: 6px;
      bottom: 6px;
      height: 28px;
      border: 1px solid #b9c9de;
      background: #f8fbff;
      border-radius: 6px;
      color: #3c5f87;
      font-size: 22px;
      line-height: 1;
      opacity: 0;
      pointer-events: none;
      cursor: pointer;
      transition: opacity 0.15s ease;
    }

    .day:hover .quick-add {
      opacity: 1;
      pointer-events: auto;
    }

    .d-num {
      font-size: 12px;
      font-weight: 700;
      color: #35516f;
    }

    .item {
      font-size: 10px;
      border-radius: 7px;
      padding: 2px 4px;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
      border: 1px solid transparent;
    }

    .item.done { background: #e9f7ed; color: #0f7a3f; }
    .item.workout { background: #fff2e8; color: #a75427; }
    .item.event { background: #e8efff; color: #1c4cb9; border-color: #cedcf8; }
    .item.goal { background: #edf8ef; color: #1f7e42; border-color: #d5efdb; }
    .item.note { background: #f6f0ff; color: #6134b6; border-color: #e7dafd; }
    .item.metrics { background: #eef8ff; color: #246d8f; border-color: #d6ecf8; }
    .item.availability { background: #f7f7f7; color: #596274; border-color: #eaecf0; }
    .item.paired-green { background: #edf9ef; border-color: #cdebd2; color: #1f7f3e; }
    .item.paired-orange { background: #fff2e8; border-color: #ffd9bf; color: #b4571f; }
    .item.paired-yellow { background: #fffde9; border-color: #f8edac; color: #8a7b20; }
    .item.paired-red { background: #ffecee; border-color: #f7c8cd; color: #b2313a; }
    .item.unplanned { background: #f2f3f5; border-color: #dfe2e6; color: #555f6f; }

    .work-card {
      width: 100%;
      border: 1px solid #d9e4ef;
      border-top: 4px solid #a1b8d4;
      border-radius: 6px;
      background: #f9fcff;
      text-align: left;
      padding: 5px 6px;
      color: #1d2a39;
      cursor: pointer;
      font-size: 10px;
      line-height: 1.25;
      font-family: inherit;
      position: relative;
    }

    .card-menu-btn {
      position: absolute;
      top: 2px;
      right: 2px;
      border: 0;
      background: transparent;
      color: #7087a0;
      cursor: pointer;
      padding: 2px 4px;
      font-size: 14px;
      line-height: 1;
      border-radius: 4px;
    }

    .card-menu-btn:hover { background: rgba(0, 0, 0, 0.06); }

    .work-card .wc-title {
      margin: 0 0 2px;
      font-size: 11px;
      font-weight: 700;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .work-card .wc-meta {
      margin: 0;
      color: #5f758e;
      font-size: 10px;
    }

    .work-card .wc-bottom {
      margin-top: 4px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      color: #506885;
      font-size: 10px;
    }

    .work-card.done { border-top-color: #4db347; background: #f2f8ef; }
    .work-card.workout { border-top-color: #c7d5e6; background: #ffffff; }
    .work-card.unplanned { border-top-color: #97a4b8; background: #f1f3f6; }
    .work-card.paired-green { border-top-color: #4db347; background: #f2f8ef; }
    .work-card.paired-orange { border-top-color: #f28d4c; background: #fff4ec; }
    .work-card.paired-yellow { border-top-color: #e0c53a; background: #fffde8; }
    .work-card.paired-red { border-top-color: #cc4a56; background: #fff1f2; }

    .delta-up { color: #b4571f; font-weight: 700; }
    .delta-down { color: #8a7b20; font-weight: 700; }
    .work-card.event { border-top-color: #3a70d8; background: #f1f5ff; }
    .work-card.goal { border-top-color: #3da86a; background: #f0faf4; }
    .work-card.note { border-top-color: #8b62cc; background: #f6f1ff; }
    .work-card.metrics { border-top-color: #3aa1be; background: #eefaff; }
    .work-card.availability { border-top-color: #8a95a6; background: #f6f7f9; }

    .week-summary {
      border: 1px solid #d9e3ef;
      border-radius: 8px;
      background: #f8fbff;
      padding: 7px;
      font-size: 11px;
      color: #415a74;
    }

    .ws-metrics {
      display: flex;
      gap: 6px;
      margin-bottom: 6px;
    }

    .ws-chip {
      flex: 1;
      border-radius: 6px;
      color: #fff;
      text-align: center;
      padding: 4px 2px;
      font-size: 10px;
      line-height: 1.2;
    }

    .ws-chip strong {
      display: block;
      font-size: 16px;
      line-height: 1;
    }

    .ws-ctl { background: var(--blue); }
    .ws-atl { background: var(--pink); }
    .ws-tsb { background: var(--orange); }

    .ws-row {
      display: flex;
      justify-content: space-between;
      border-top: 1px solid #e5edf6;
      padding-top: 5px;
      margin-top: 5px;
    }

    .grid { display: grid; grid-template-columns: repeat(12, minmax(0, 1fr)); gap: 10px; }

    .stat {
      font-size: 28px;
      font-weight: 800;
      margin: 4px 0 2px;
    }

    .sub { margin: 0; color: var(--muted); font-size: 12px; }

    .toggle {
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-top: 1px solid #edf2f7;
      padding: 9px 0;
      font-size: 14px;
    }

    .modal {
      position: fixed;
      inset: 0;
      background: rgba(11, 24, 39, 0.45);
      display: none;
      justify-content: center;
      align-items: center;
      padding: 16px;
      z-index: 50;
    }

    .modal.open { display: flex; }

    .settings-card {
      width: min(760px, 100%);
      max-height: 90vh;
      overflow: auto;
      border-radius: 14px;
      border: 1px solid #d4e1ef;
      background: #fff;
      box-shadow: 0 24px 44px rgba(12, 26, 42, 0.24);
      padding: 18px;
    }

    .modal-card {
      width: min(1320px, 100%);
      max-height: 94vh;
      overflow: auto;
      border-radius: 14px;
      border: 1px solid #d4e1ef;
      background: #fff;
      box-shadow: 0 24px 44px rgba(12, 26, 42, 0.24);
      padding: 20px;
    }

    .modal-top {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
      gap: 8px;
      flex-wrap: wrap;
    }

    .m-title { margin: 0; font-size: 62px; line-height: 1.05; color: #1f2b3d; }
    .m-title.small { font-size: 44px; }

    .icon-btn {
      border: 0;
      background: transparent;
      color: #495f7e;
      border-radius: 9px;
      padding: 4px;
      font-size: 32px;
      line-height: 1;
      cursor: pointer;
    }

    .section-label {
      margin: 18px 0 10px;
      font-size: 47px;
      color: #263348;
    }

    .type-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 8px;
      margin-top: 6px;
    }

    .type-btn {
      border: 0;
      background: transparent;
      border-radius: 10px;
      padding: 8px 10px;
      font-size: 22px;
      text-align: left;
      cursor: pointer;
      display: flex;
      flex-direction: row;
      align-items: center;
      gap: 8px;
      min-width: 0;
      overflow: hidden;
    }

    .type-btn span {
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .type-icon {
      width: 32px;
      height: 32px;
      text-align: center;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 8px;
      background: #eef4fb;
      overflow: hidden;
      flex: 0 0 32px;
    }

    .type-icon img {
      width: 22px;
      height: 22px;
      object-fit: contain;
      display: block;
    }

    .type-btn .type-icon {
      width: 48px;
      height: 48px;
      flex: 0 0 48px;
      background: transparent;
      display: grid;
      place-items: center;
      margin: 0;
    }

    .type-btn .type-icon img {
      width: 36px;
      height: 36px;
      object-fit: contain;
      object-position: center;
    }

    .type-btn:hover .type-icon { background: rgba(31, 91, 215, 0.08); }
    .type-btn:hover { background: rgba(31, 91, 215, 0.04); }

    .detail-shell {
      border: 1px solid #d7e1ef;
      border-radius: 12px;
      overflow: hidden;
      background: #fff;
    }

    .detail-top {
      display: grid;
      grid-template-columns: 1fr auto auto auto;
      gap: 10px;
      align-items: center;
      padding: 12px 14px;
      border-bottom: 1px solid #e3ebf4;
      background: #f8fbff;
    }

    .detail-date {
      color: #2753ce;
      font-size: 15px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.2px;
    }

    .mini-chip {
      border-radius: 8px;
      color: #fff;
      font-size: 14px;
      padding: 6px 10px;
      font-weight: 700;
    }
    .wv-row1 .mini-chip {
      font-size: 11px;
      padding: 4px 8px;
    }

    .mini-chip.ctl { background: var(--blue); }
    .mini-chip.atl { background: var(--pink); }
    .mini-chip.tsb { background: var(--orange); }

    .detail-body {
      display: grid;
      grid-template-columns: 1fr 340px;
      gap: 0;
      min-height: 460px;
    }

    .detail-left {
      padding: 14px;
      border-right: 1px solid #e3ebf4;
    }

    .detail-right {
      padding: 14px;
      background: #fbfdff;
    }

    .field {
      margin-bottom: 12px;
    }

    .field label {
      display: block;
      margin-bottom: 4px;
      color: #4e647d;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.32px;
      font-weight: 700;
    }

    .field input,
    .field textarea,
    .field select {
      width: 100%;
      border: 1px solid #cddaea;
      border-radius: 8px;
      padding: 10px;
      font-size: 14px;
      font-family: inherit;
      background: #fff;
    }

    .field textarea { min-height: 120px; resize: vertical; }

    .two-col {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }

    .comment-card {
      border: 1px solid #dde6f1;
      border-radius: 8px;
      padding: 16px;
      color: #8a9ab0;
      text-align: center;
      font-size: 14px;
      min-height: 180px;
      display: grid;
      place-items: center;
      background: #fff;
    }

    .modal-footer {
      border-top: 1px solid #e3ebf4;
      padding: 12px 14px;
      display: flex;
      justify-content: flex-end;
      gap: 8px;
      align-items: center;
      background: #fff;
    }

    .btn.ghost { background: transparent; color: #33528a; }

    .workout-view-modal {
      z-index: 60;
    }

    #actionModal .modal-card {
      width: min(920px, 100%);
      max-height: 76vh;
      overflow: hidden;
      padding: 12px;
    }

    #actionModal .m-title { font-size: 34px; }
    #actionModal .section-label { font-size: 24px; margin: 8px 0 6px; }
    #actionModal .type-btn { font-size: 20px; padding: 6px 8px; }

    .wv-shell {
      border: 1px solid #d7e1ef;
      border-radius: 12px;
      background: #fff;
      overflow: hidden;
    }

    .wv-top {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 10px;
      align-items: start;
      padding: 12px 14px;
      border-bottom: 1px solid #e3ebf4;
      background: #f8fbff;
      position: relative;
    }

    .wv-head-main {
      display: grid;
      gap: 6px;
    }

    .wv-row1 {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }

    .wv-row2 {
      display: grid;
      grid-template-columns: auto 1fr;
      gap: 8px;
      align-items: start;
    }

    .wv-date-line {
      color: #1f5bd7;
      font-size: 16px;
      font-weight: 700;
      letter-spacing: 0.2px;
      text-transform: uppercase;
      margin-right: 8px;
    }

    .wv-time-select {
      border: 1px solid #c9d7e9;
      border-radius: 8px;
      padding: 6px 8px;
      font-size: 14px;
      color: #2c4f79;
      background: #fff;
      font-family: inherit;
    }

    .wv-sport-wrap {
      display: inline-grid;
      align-content: start;
      justify-items: center;
      gap: 4px;
      position: relative;
      min-width: 72px;
    }

    .wv-sport-toggle {
      border: 1px solid #c9d7e9;
      border-radius: 6px;
      background: #fff;
      padding: 2px 5px;
      width: 24px;
      height: 24px;
      font-size: 12px;
      color: #32567d;
      font-family: inherit;
      cursor: pointer;
    }

    .wv-sport-wrap .type-icon {
      width: 64px;
      height: 64px;
      border-radius: 10px;
      flex: 0 0 64px;
    }

    .wv-sport-wrap .type-icon img {
      width: 44px;
      height: 44px;
    }

    .wv-title-input {
      border: 1px solid #c9d7e9;
      border-radius: 8px;
      padding: 8px 10px;
      font-size: 11px;
      font-weight: 700;
      color: #1f2f43;
      min-width: 260px;
      flex: 1;
      font-family: inherit;
    }

    .wv-sport-name {
      font-size: 11px;
      color: #345777;
      font-weight: 700;
      line-height: 1.1;
      text-align: center;
      max-width: 86px;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
    }

    .wv-title-stack {
      display: grid;
      gap: 6px;
      min-width: 0;
    }

    .wv-sport-select {
      border: 1px solid #c9d7e9;
      border-radius: 8px;
      padding: 7px 10px;
      font-size: 14px;
      color: #2c4f79;
      background: #fff;
      font-family: inherit;
    }

    .wv-meta-line {
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
      color: #5f758e;
      font-size: 12px;
    }

    .wv-meta-line strong {
      color: #1f2f43;
      font-size: 16px;
    }

    .wv-tabs {
      display: flex;
      flex-direction: column;
      gap: 6px;
      align-items: stretch;
    }

    .wv-tab {
      border: 1px solid #c9d7e9;
      background: #1f5bd7;
      border-radius: 8px;
      padding: 6px 10px;
      color: #fff;
      font-weight: 700;
      cursor: pointer;
    }

    .wv-tab.active { box-shadow: 0 0 0 2px rgba(31, 91, 215, 0.2) inset; }
    .wv-tab:hover { background: #2a66d2; }
    .wv-tab.outline {
      background: #fff;
      color: #1f5bd7;
      border-color: #8fb2e9;
    }
    .wv-tab.outline:hover { background: #edf4ff; color: #1b4fbf; }
    .wv-tab.solid:hover { background: #2a66d2; }

    .wv-tab.disabled {
      opacity: 0.45;
      cursor: not-allowed;
      pointer-events: none;
    }

    .wv-sport-menu {
      position: absolute;
      top: 70px;
      left: 0;
      background: #fff;
      border: 1px solid #d4e0ef;
      border-radius: 10px;
      box-shadow: 0 10px 22px rgba(18, 38, 62, 0.18);
      padding: 8px;
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 6px;
      width: 260px;
      z-index: 50;
    }

    .wv-sport-item {
      border: 1px solid transparent;
      background: #fff;
      border-radius: 8px;
      padding: 6px;
      display: grid;
      place-items: center;
      cursor: pointer;
    }
    .wv-sport-item:hover { border-color: #9bb8e2; background: #f3f8ff; }
    .wv-sport-item span {
      font-size: 10px;
      color: #36597f;
      margin-top: 2px;
      max-width: 56px;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
      text-align: center;
      display: block;
    }
    .wv-sport-item .type-icon {
      width: 48px;
      height: 48px;
      flex: 0 0 48px;
      display: grid;
      place-items: center;
    }
    .wv-sport-item .type-icon img {
      width: 28px;
      height: 28px;
      object-fit: contain;
      object-position: center;
    }

    .rpe-slider.rpe-unset {
      opacity: 0.6;
      accent-color: #c8d6ea;
    }

    .feel-row {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }

    .feel-btn {
      width: 34px;
      height: 34px;
      border-radius: 999px;
      border: 1px solid #c9d7e9;
      background: #fff;
      cursor: pointer;
      font-size: 18px;
      line-height: 1;
    }

    .feel-btn.active {
      border-color: #1f5bd7;
      box-shadow: 0 0 0 2px #dce9ff inset;
    }

    .feel-btn:disabled,
    .rpe-slider:disabled {
      opacity: 0.45;
      cursor: not-allowed;
    }

    .rpe-slider {
      width: 100%;
    }

    .wv-body {
      padding: 14px;
      max-height: 76vh;
      overflow: auto;
    }

    .workout-view-modal .modal-card {
      width: min(760px, 100%) !important;
      max-height: 95vh;
    }

    .wv-grid {
      display: grid;
      grid-template-columns: 1fr 330px;
      gap: 12px;
    }

    .wv-grid-vertical {
      display: grid;
      grid-template-columns: 1fr;
      gap: 10px;
    }

    .wv-card {
      border: 1px solid #dde7f3;
      border-radius: 10px;
      padding: 10px;
      background: #fbfdff;
    }

    .wv-kv {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 6px;
      margin-top: 8px;
    }

    .wv-kv div {
      border: 1px solid #e0e9f4;
      border-radius: 8px;
      background: #fff;
      padding: 8px;
      font-size: 12px;
      color: #4f657f;
    }

    .wv-kv strong {
      display: block;
      color: #1f2f43;
      font-size: 18px;
      margin-top: 2px;
    }

    .wv-file-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      border: 1px solid #dce6f2;
      border-radius: 8px;
      padding: 8px 10px;
      background: #fff;
      margin-top: 8px;
      gap: 8px;
      flex-wrap: wrap;
    }

    .wv-file-actions {
      display: flex;
      gap: 6px;
      flex-wrap: wrap;
    }

    .wv-files-popover {
      position: absolute;
      right: 58px;
      top: 44px;
      width: 360px;
      z-index: 40;
    }

    .comment-feed {
      border: 1px solid #dde6f1;
      border-radius: 8px;
      background: #fff;
      min-height: 90px;
      max-height: 180px;
      overflow: auto;
      padding: 8px;
      display: grid;
      gap: 6px;
    }

    .comment-feed .comment-item {
      border: 1px solid #e1e9f4;
      border-radius: 6px;
      padding: 6px 8px;
      font-size: 12px;
      color: #2f455f;
      background: #f9fcff;
      white-space: pre-wrap;
    }

    .comment-input-row {
      margin-top: 8px;
      display: flex;
      gap: 8px;
      align-items: center;
    }

    .comment-input-row input {
      flex: 1;
      border: 1px solid #cddaea;
      border-radius: 8px;
      padding: 8px 10px;
      font-family: inherit;
    }

    .pc-table {
      display: grid;
      grid-template-columns: 1fr 1fr 1fr;
      gap: 6px;
      margin-top: 8px;
      font-size: 12px;
    }

    .pc-table > div {
      border: 1px solid #dfe8f3;
      border-radius: 6px;
      background: #fff;
      padding: 6px;
      color: #445d79;
    }

    .pc-table .pc-head {
      background: #edf4fb;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .25px;
      font-size: 11px;
    }

    .chart-box {
      border: 1px solid #dbe6f2;
      border-radius: 10px;
      background: #f8fbff;
      padding: 10px;
    }

    .legend {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 8px;
      font-size: 11px;
      color: #536a84;
    }

    .legend span::before {
      content: "";
      display: inline-block;
      width: 10px;
      height: 10px;
      border-radius: 2px;
      margin-right: 4px;
      vertical-align: -1px;
      background: currentColor;
    }

    .l-hr { color: #f35353; }
    .l-pwr { color: #ff62f2; }
    .l-cad { color: #f39b1f; }
    .l-spd { color: #3fa144; }

    .lap-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10px;
      font-size: 12px;
    }

    .lap-table th,
    .lap-table td {
      border: 1px solid #dce6f2;
      padding: 6px;
      text-align: left;
    }

    .lap-table th {
      background: #eef4fb;
      color: #445d79;
    }

    .lap-table tr.selected {
      background: #e6f0ff;
    }

    .ctx-menu {
      position: fixed;
      z-index: 90;
      background: #fff;
      border: 1px solid #d2dfef;
      border-radius: 8px;
      box-shadow: 0 14px 22px rgba(12, 26, 42, 0.18);
      min-width: 160px;
      display: none;
      padding: 4px;
    }

    .ctx-menu button {
      width: 100%;
      text-align: left;
      border: 0;
      background: transparent;
      padding: 8px 9px;
      border-radius: 6px;
      cursor: pointer;
      color: #2b425d;
      font-weight: 600;
    }

    .ctx-menu button:hover {
      background: #edf4ff;
    }

    .tp-workout-layout {
      display: grid;
      grid-template-columns: 1.05fr 1fr;
      gap: 16px;
      align-items: start;
    }

    .tp-table-head {
      display: grid;
      grid-template-columns: 110px 1fr 1fr 80px;
      gap: 6px;
      margin-bottom: 6px;
      color: #28384d;
      font-size: 18px;
    }

    .tp-row {
      display: grid;
      grid-template-columns: 110px 1fr 1fr 80px;
      gap: 6px;
      align-items: center;
      margin-bottom: 6px;
    }

    .tp-row label {
      text-align: right;
      color: #24364c;
      font-size: 14px;
      font-weight: 500;
    }

    .tp-in {
      width: 100%;
      border: 1px solid #cfd9e7;
      border-radius: 6px;
      background: #fff;
      padding: 6px 8px;
      height: 34px;
      font-size: 14px;
      color: #2a3a51;
    }

    .tp-in.readonly {
      background: #eef2f7;
    }

    .tp-in.muted {
      background: #eef2f7;
    }

    .tp-unit {
      color: #24364c;
      font-size: 16px;
      text-align: left;
    }

    .tp-unit-select {
      width: 100%;
      border: 1px solid #cfd9e7;
      border-radius: 6px;
      background: #fff;
      height: 34px;
      padding: 5px 7px;
      font-size: 14px;
      font-family: inherit;
    }

    .tp-minmax-head {
      display: grid;
      grid-template-columns: 110px 1fr 1fr 1fr 80px;
      gap: 6px;
      margin: 14px 0 6px;
      color: #25364b;
      font-size: 13px;
      text-align: center;
    }

    .tp-minmax-row {
      display: grid;
      grid-template-columns: 110px 1fr 1fr 1fr 80px;
      gap: 6px;
      align-items: center;
      margin-bottom: 6px;
    }

    .tp-minmax-row label {
      text-align: right;
      font-size: 14px;
      color: #24364c;
    }

    .tp-right-block .field label {
      font-size: 14px;
      text-transform: none;
      letter-spacing: 0;
      font-weight: 500;
      color: #24364c;
    }

    .tp-right-block .field textarea,
    .tp-right-block .field input {
      border-radius: 6px;
      font-size: 14px;
      min-height: 34px;
    }

    .tp-equipment-title {
      font-size: 18px;
      color: #24364c;
      text-align: center;
      margin: 14px 0 8px;
    }

    .tp-eq-row {
      display: grid;
      grid-template-columns: 110px 1fr;
      gap: 8px;
      margin-bottom: 8px;
      align-items: center;
    }

    .tp-eq-row label {
      text-align: right;
      font-size: 14px;
      color: #24364c;
    }

    .tp-select {
      border: 1px solid #d4dce8;
      border-radius: 8px;
      background: #eef2f7;
      color: #55657d;
      padding: 8px 12px;
      font-size: 14px;
      height: 36px;
      width: 100%;
    }

    .hidden { display: none; }

    @media (max-width: 1450px) {
      .home-grid { grid-template-columns: 260px 1fr 330px; }
      .dow,
      .week-row { grid-template-columns: repeat(7, minmax(0, 1fr)) 220px; }
      .m-title { font-size: 46px; }
      .section-label { font-size: 34px; }
      .type-btn { font-size: 30px; }
    }

    @media (max-width: 1200px) {
      .home-grid { grid-template-columns: 1fr; }
      .dow,
      .week-row { grid-template-columns: repeat(7, minmax(0, 1fr)); }
      .dow .sum-head,
      .week-summary { display: none; }
      .detail-body { grid-template-columns: 1fr; }
      .detail-left { border-right: 0; border-bottom: 1px solid #e3ebf4; }
      .wv-grid { grid-template-columns: 1fr; }
      .tp-workout-layout { grid-template-columns: 1fr; }
    }

    @media (max-width: 820px) {
      .top-nav { grid-template-columns: 1fr; gap: 6px; padding: 8px; }
      .brand, .nav-right { justify-self: center; }
      .tabs { flex-wrap: wrap; }
      .type-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .m-title { font-size: 34px; }
      .section-label { font-size: 28px; }
      .type-btn { font-size: 24px; }
    }
  </style>
</head>
<body>
  <header class="top-nav">
    <div class="brand">
      <img id="brandIcon" class="brand-icon" src="/icons/logos/logo.png" alt="Training Freaks logo" />
      <img id="brandText" class="brand-text" src="/icons/logos/logotext.png" alt="Training Freaks" />
    </div>
    <nav class="tabs">
      <button class="tab active" data-view="home">Home</button>
      <button class="tab" data-view="calendar">Calendar</button>
      <button class="tab" data-view="dashboard">Dashboard</button>
    </nav>
    <div class="nav-right">
      <input id="uploadFitInput" type="file" accept=".fit" style="display:none;" />
      <span>Amir LaGasse</span>
      <button class="nav-settings" id="globalSettings" title="Settings">&#9881;</button>
    </div>
  </header>

  <main class="main">
    <div class="page-head">
      <h1 id="pageTitle">Home</h1>
    </div>

    <section id="view-home" class="view active">
      <div class="home-grid">
        <aside class="stack">
          <div class="panel">
            <h3 class="panel-title">
              <span>Events</span>
              <button class="plus-btn" id="addEventBtn">+</button>
            </h3>
            <div id="eventsList"></div>
          </div>

          <div class="panel">
            <h3 class="panel-title">
              <span>Goals</span>
              <button class="goal-btn" id="addGoalBtn">Add Goal</button>
            </h3>
            <div id="goalsList"></div>
          </div>
        </aside>

        <section class="stack">
          <div class="panel">
            <h3 class="panel-title"><span>Today - Completed Workouts</span></h3>
            <div id="todayDone"></div>
          </div>
          <div class="panel">
            <h3 class="panel-title"><span>Planned Workouts</span></h3>
            <div id="todayPlanned"></div>
          </div>
        </section>

        <aside class="panel">
          <h3 class="panel-title"><span>Performance Metrics</span></h3>
          <div class="metric-row">
            <div class="metric-chip chip-ctl">
              <span class="num" id="ctlVal">0</span>
              <span class="lbl">Fitness (CTL)</span>
            </div>
            <div class="metric-chip chip-atl">
              <span class="num" id="atlVal">0</span>
              <span class="lbl">Fatigue (ATL)</span>
            </div>
            <div class="metric-chip chip-tsb">
              <span class="num" id="tsbVal">0</span>
              <span class="lbl">Form (TSB)</span>
            </div>
          </div>

          <div class="explain">
            Training Stress Score (TSS) is estimated from duration and intensity (placeholder formula).<br/><br/>
            Fitness (CTL) is a 42-day rolling average of daily TSS.<br/>
            Fatigue (ATL) is a 7-day rolling average of daily TSS.<br/>
            Form (TSB) is yesterday's CTL minus yesterday's ATL.
          </div>

          <div class="spark-wrap">
            <div class="spark-box">
              <div class="spark-head"><span>Fitness Trend</span><strong id="ctlTrend">0</strong></div>
              <div class="sparkline" id="ctlSpark"></div>
            </div>
            <div class="spark-box">
              <div class="spark-head"><span>Fatigue Trend</span><strong id="atlTrend">0</strong></div>
              <div class="sparkline" id="atlSpark"></div>
            </div>
            <div class="spark-box">
              <div class="spark-head"><span>Form Trend</span><strong id="tsbTrend">0</strong></div>
              <div class="sparkline" id="tsbSpark"></div>
            </div>
          </div>
        </aside>
      </div>
    </section>

    <section id="view-calendar" class="view">
      <div class="calendar-wrap">
        <div class="calendar-head">
          <div class="cal-nav">
            <h3 id="calHeaderMonth" style="margin:0;font-size:18px;color:#31455f;">February 2026</h3>
            <button class="btn secondary" id="calTodayBtn" type="button" style="padding:5px 10px;font-size:12px;">Today</button>
          </div>
          <div style="display:flex;gap:8px;align-items:center;">
            <button class="import-btn" id="uploadFitBtn" title="Import FIT" aria-label="Import FIT">&#8682;</button>
          </div>
        </div>
        <div id="calendarScroll" class="calendar-scroll"></div>
      </div>
    </section>

    <section id="view-dashboard" class="view">
      <div class="grid" id="dashboardGrid">
        <div class="panel" data-widget="count" style="grid-column: span 3;">
          <h3 class="panel-title"><span>Completed Activities</span></h3>
          <p class="stat" id="statCount">0</p>
          <p class="sub">Strava imports loaded</p>
        </div>

        <div class="panel" data-widget="plannedCount" style="grid-column: span 3;">
          <h3 class="panel-title"><span>Planned Workouts</span></h3>
          <p class="stat" id="statPlanned">0</p>
          <p class="sub">Custom planned sessions</p>
        </div>

        <div class="panel" data-widget="distance" style="grid-column: span 3;">
          <h3 class="panel-title"><span>Total Distance</span></h3>
          <p class="stat" id="statDistance">0 km</p>
          <p class="sub">Across imported activities</p>
        </div>

        <div class="panel" data-widget="time" style="grid-column: span 3;">
          <h3 class="panel-title"><span>Total Moving Time</span></h3>
          <p class="stat" id="statTime">0 h</p>
          <p class="sub">Across imported activities</p>
        </div>

        <div class="panel" style="grid-column: span 12;">
          <h3 class="panel-title"><span>Customize Dashboard</span></h3>
          <label class="toggle"><span>Show completed activity count</span><input type="checkbox" data-toggle="count" checked /></label>
          <label class="toggle"><span>Show planned workout count</span><input type="checkbox" data-toggle="plannedCount" checked /></label>
          <label class="toggle"><span>Show distance</span><input type="checkbox" data-toggle="distance" checked /></label>
          <label class="toggle"><span>Show moving time</span><input type="checkbox" data-toggle="time" checked /></label>
        </div>
      </div>
    </section>

  </main>
  <div id="settingsModal" class="modal">
    <div class="settings-card">
      <div class="modal-top" style="margin-bottom:6px;">
        <h3 style="margin:0;">Settings</h3>
        <button class="icon-btn" id="closeSettings">&times;</button>
      </div>
      <div class="panel" style="box-shadow:none;">
        <h3 class="panel-title"><span>FTP Settings</span></h3>
        <div class="grid" style="grid-template-columns: repeat(2, minmax(0, 1fr));">
          <div class="field"><label>Bike FTP (W)</label><input id="ftpRide" type="number" min="50" max="600" placeholder="--" /></div>
          <div class="field"><label>Run FTP (optional)</label><input id="ftpRun" type="number" min="50" max="600" placeholder="--" /></div>
          <div class="field"><label>Row FTP (W)</label><input id="ftpRow" type="number" min="50" max="600" placeholder="--" /></div>
          <div class="field"><label>Swim FTP (optional)</label><input id="ftpSwim" type="number" min="50" max="600" placeholder="--" /></div>
          <div class="field"><label>Strength FTP (optional)</label><input id="ftpStrength" type="number" min="50" max="600" placeholder="--" /></div>
          <div class="field"><label>Other FTP</label><input id="ftpOther" type="number" min="50" max="600" placeholder="--" /></div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;margin-top:8px;">
          <button class="btn primary" id="saveSettingsBtn">Save</button>
          <span id="settingsSavedMsg" class="meta" style="display:none;color:#17733e;">Saved</span>
        </div>
      </div>
    </div>
  </div>

  <div id="actionModal" class="modal">
    <div class="modal-card">
      <div class="modal-top">
        <h2 class="m-title" id="actionDateTitle">Wednesday, February 11, 2026</h2>
        <button class="icon-btn" id="closeAction">&times;</button>
      </div>

      <h3 class="section-label">Add a Workout</h3>
      <div class="type-grid" id="workoutTypeGrid"></div>

      <h3 class="section-label">Add Other</h3>
      <div class="type-grid" id="otherTypeGrid"></div>
    </div>
  </div>

  <div id="contextMenu" class="ctx-menu"></div>

  <div id="detailModal" class="modal">
    <div class="modal-card" style="padding: 0;">
      <div class="detail-shell">
        <div class="detail-top">
          <div class="detail-date" id="detailDateLabel">WEDNESDAY FEBRUARY 11, 2026</div>
          <div id="detailMetricsChips" style="display:contents;">
            <div class="mini-chip ctl" id="miniCtl">Fitness 0</div>
            <div class="mini-chip atl" id="miniAtl">Fatigue 0</div>
            <div class="mini-chip tsb" id="miniTsb">Form 0</div>
          </div>
        </div>

        <div class="detail-body">
          <div class="detail-left">
            <div class="field">
              <label id="detailTitleLabel">Title</label>
              <input id="dTitle" placeholder="Untitled" />
            </div>

            <div class="field">
              <label>Date</label>
              <input id="dDate" type="date" />
            </div>

            <div id="workoutFields" class="hidden">
              <div class="tp-workout-layout">
                <div>
                  <div class="tp-table-head">
                    <div></div><div>Planned</div><div>Completed</div><div></div>
                  </div>
                  <div class="tp-row">
                    <label>Duration</label>
                    <input id="dDuration" class="tp-in" type="number" min="0" />
                    <input class="tp-in readonly" readonly />
                    <div class="tp-unit">h:m:s</div>
                  </div>
                  <div class="tp-row">
                    <label>Distance</label>
                    <input id="dDistance" class="tp-in" type="number" min="0" step="0.1" />
                    <input class="tp-in readonly" readonly />
                    <div class="tp-unit distance-unit-label">km</div>
                  </div>
                  <div class="tp-row">
                    <label>TSS</label>
                    <input class="tp-in readonly" readonly />
                    <input class="tp-in readonly" readonly />
                    <div class="tp-unit">TSS</div>
                  </div>
                  <div class="tp-row">
                    <label>IF</label>
                    <input class="tp-in readonly" readonly />
                    <input class="tp-in readonly" readonly />
                    <div class="tp-unit">IF</div>
                  </div>
                  <div class="tp-minmax-head">
                    <div></div><div>Min</div><div>Avg</div><div>Max</div><div></div>
                  </div>
                  <div class="tp-minmax-row">
                    <label>Heart Rate</label>
                    <input class="tp-in readonly" readonly />
                    <input class="tp-in readonly" readonly />
                    <input class="tp-in readonly" readonly />
                    <div class="tp-unit">bpm</div>
                  </div>
                  <div class="tp-minmax-row">
                    <label>Power</label>
                    <input class="tp-in readonly" readonly />
                    <input class="tp-in readonly" readonly />
                    <input class="tp-in readonly" readonly />
                    <div class="tp-unit">W</div>
                  </div>
                  <div class="tp-equipment-title">Equipment</div>
                  <div class="tp-eq-row">
                    <label>Bike</label>
                    <select class="tp-select"><option>Select Bike</option></select>
                  </div>
                  <div class="tp-eq-row">
                    <label>Shoes</label>
                    <select class="tp-select"><option>Select Shoe</option></select>
                  </div>
                </div>
                <div class="tp-right-block">
                  <div class="field">
                    <label>Description</label>
                    <textarea id="dDescription" style="min-height:72px;" placeholder="Add details"></textarea>
                  </div>
                  <div class="field">
                    <label>Post-activity Comments</label>
                    <input placeholder="Enter a new comment" />
                  </div>
                  <div class="field" style="margin-top:10px;">
                    <label>Workout Type</label>
                    <input id="dWorkoutType" />
                  </div>
                  <div class="field">
                    <label>Intensity (1-10)</label>
                    <input id="dIntensity" type="number" min="1" max="10" step="1" value="6" />
                  </div>
                </div>
              </div>
            </div>

            <div id="eventFields" class="hidden">
              <div class="field">
                <label>Event Type</label>
                <select id="dEventType">
                  <option>Race</option>
                  <option>Test</option>
                  <option>Camp</option>
                  <option>A Race</option>
                  <option>B Race</option>
                </select>
              </div>
            </div>

            <div id="availabilityFields" class="hidden">
              <div class="field">
                <label>Availability</label>
                <select id="dAvailability">
                  <option>Unavailable</option>
                  <option>Limited Availability</option>
                </select>
              </div>
            </div>

            <div id="nonWorkoutDescription" class="field">
              <label>Description</label>
              <textarea id="dDescriptionOther" placeholder="Add details"></textarea>
            </div>
          </div>

          <div class="detail-right">
            <h3 style="margin: 0 0 10px; color: #607690; font-size: 12px; letter-spacing: .35px; text-transform: uppercase;">Comments</h3>
            <div class="comment-card">Comments are available once the item has been saved.</div>
          </div>
        </div>

        <div class="modal-footer">
          <button class="btn ghost" id="deleteDetail">Delete</button>
          <button class="btn ghost" id="cancelDetail">Cancel</button>
          <button class="btn secondary" id="saveDetail">Save</button>
          <button class="btn primary" id="saveCloseDetail">Save &amp; Close</button>
        </div>
      </div>
    </div>
  </div>

  <div id="workoutViewModal" class="modal workout-view-modal">
    <div class="modal-card" style="padding: 0; width: min(900px, 100%);">
      <div class="wv-shell">
        <div class="wv-top">
          <div class="wv-head-main">
            <div class="wv-row1">
              <div id="wvDateLine" class="wv-date-line">WEDNESDAY, FEBRUARY 11, 2026</div>
              <select id="wvTimeSelect" class="wv-time-select"><option>8:07 am</option></select>
              <div class="mini-chip ctl" id="wvHeadCtl">Fitness 0</div>
              <div class="mini-chip atl" id="wvHeadAtl">Fatigue 0</div>
              <div class="mini-chip tsb" id="wvHeadTsb">Form 0</div>
            </div>
            <div class="wv-row2">
              <div class="wv-sport-wrap">
                <span id="wvSportIcon" class="type-icon"></span>
                <button type="button" id="wvSportToggle" class="wv-sport-toggle">▼</button>
                <div id="wvSportName" class="wv-sport-name">Workout</div>
                <div id="wvSportMenu" class="wv-sport-menu hidden"></div>
              </div>
              <div class="wv-title-stack">
                <input id="wvTitle" class="wv-title-input" value="Workout" />
                <div class="wv-meta-line">
                  <strong id="wvHeaderDuration">--:--:--</strong>
                  <strong id="wvHeaderDistance">-- km</strong>
                  <strong id="wvHeaderTss">-- TSS</strong>
                </div>
              </div>
            </div>
          </div>
          <div class="wv-tabs">
            <button class="wv-tab solid" id="wvFilesTabBtn">Files</button>
            <button class="wv-tab outline" id="wvAnalyzeBtn">Analyze</button>
            <button class="icon-btn" id="closeWorkoutView">&times;</button>
          </div>
          <div id="wvFilesPopover" class="wv-files-popover hidden">
            <div class="wv-card" style="padding:10px;">
              <h3 style="margin:0 0 8px;">Upload Device File</h3>
              <div style="display:flex;gap:8px;align-items:center;">
                <button class="btn secondary" id="wvBrowseFilesBtn">Browse Files</button>
              </div>
              <div id="wvFileList" style="margin-top:10px;"></div>
            </div>
          </div>
        </div>
        <div class="wv-body">
          <section id="wvSummary">
            <div class="wv-grid-vertical">
              <div class="wv-card">
                <h3 style="margin: 0 0 6px;">Workout Summary</h3>
                <p class="meta" id="wvSummaryText">-</p>
                <div class="tp-workout-layout">
                  <div>
                    <div class="tp-table-head">
                      <div></div><div>Planned</div><div>Completed</div><div></div>
                    </div>
                    <div class="tp-row">
                      <label>Duration</label>
                      <input id="pcDurPlan" class="tp-in" />
                      <input id="pcDurComp" class="tp-in" />
                      <div class="tp-unit">h:m:s</div>
                    </div>
                    <div class="tp-row">
                      <label>Distance</label>
                      <input id="pcDistPlan" class="tp-in" />
                      <input id="pcDistComp" class="tp-in" />
                      <div class="tp-unit">
                        <select id="pcDistanceUnit" class="tp-unit-select">
                          <option value="km">km</option>
                          <option value="mi">mi</option>
                          <option value="m">m</option>
                        </select>
                      </div>
                    </div>
                    <div class="tp-row">
                      <label>Elevation</label>
                      <input id="pcElevPlan" class="tp-in" />
                      <input id="pcElevComp" class="tp-in" />
                      <div class="tp-unit">
                        <select id="pcElevationUnit" class="tp-unit-select">
                          <option value="m">m</option>
                          <option value="ft">ft</option>
                        </select>
                      </div>
                    </div>
                    <div class="tp-row">
                      <label>TSS</label>
                      <input id="pcTssPlan" class="tp-in" />
                      <input id="pcTssComp" class="tp-in" />
                      <div class="tp-unit">TSS</div>
                    </div>
                    <div class="tp-row">
                      <label>IF</label>
                      <input id="pcIfPlan" class="tp-in" />
                      <input id="pcIfComp" class="tp-in" />
                      <div class="tp-unit">IF</div>
                    </div>
                    <div class="tp-minmax-head">
                      <div></div><div>Min</div><div>Avg</div><div>Max</div><div></div>
                    </div>
                    <div class="tp-minmax-row">
                      <label>Heart Rate</label>
                      <input id="wvHrMin" class="tp-in" />
                      <input id="wvHrAvg" class="tp-in" />
                      <input id="wvHrMax" class="tp-in" />
                      <div class="tp-unit">bpm</div>
                    </div>
                    <div class="tp-minmax-row">
                      <label>Power</label>
                      <input id="wvPowerMin" class="tp-in" />
                      <input id="wvPowerAvg" class="tp-in" />
                      <input id="wvPowerMax" class="tp-in" />
                      <div class="tp-unit">W</div>
                    </div>
                    <div class="tp-equipment-title">Equipment</div>
                    <div class="tp-eq-row">
                      <label>Bike</label>
                      <select class="tp-select"><option>Select Bike</option></select>
                    </div>
                    <div class="tp-eq-row">
                      <label>Shoes</label>
                      <select class="tp-select"><option>Select Shoe</option></select>
                    </div>
                  </div>
                  <div class="tp-right-block">
                    <div class="field">
                      <label>Description</label>
                      <textarea id="wvDescription" style="min-height:72px;"></textarea>
                    </div>
                    <div class="field">
                      <label>How did you feel?</label>
                      <div class="feel-row" id="feelRow">
                        <button class="feel-btn" type="button" data-feel="1" title="Very Weak">😫</button>
                        <button class="feel-btn" type="button" data-feel="2" title="Weak">🙁</button>
                        <button class="feel-btn" type="button" data-feel="3" title="Normal">😐</button>
                        <button class="feel-btn" type="button" data-feel="4" title="Strong">🙂</button>
                        <button class="feel-btn" type="button" data-feel="5" title="Very Strong">😁</button>
                      </div>
                    </div>
                    <div class="field">
                      <label>Rating of Perceived Exertion (RPE): <span id="wvRpeVal"></span></label>
                      <input id="wvRpe" class="rpe-slider rpe-unset" type="range" min="1" max="10" step="1" value="1" />
                    </div>
                    <div class="field">
                      <label>Post-activity comments</label>
                      <div id="wvCommentsFeed" class="comment-feed"></div>
                      <div class="comment-input-row">
                        <input id="wvCommentInput" placeholder="Enter a new comment" />
                        <button class="btn secondary" id="wvPostCommentBtn" type="button">Post</button>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </section>
          <section id="wvAnalyze" class="hidden">
            <div class="wv-grid">
              <div>
                <div class="chart-box">
                  <div class="legend">
                    <span class="l-hr">Heart Rate</span>
                    <span class="l-pwr">Watts</span>
                    <span class="l-cad">Cadence</span>
                    <span class="l-spd">Speed</span>
                  </div>
                  <svg id="wvChart" viewBox="0 0 1200 360" width="100%" height="320" role="img" aria-label="Workout analysis chart"></svg>
                  <svg id="wvViewfinder" viewBox="0 0 1200 90" width="100%" height="80" role="img" aria-label="Workout viewfinder"></svg>
                </div>
                <table class="lap-table" id="wvLapTable">
                  <thead>
                    <tr>
                      <th></th>
                      <th>Segment</th>
                      <th>Duration</th>
                      <th>Distance</th>
                      <th>Avg HR</th>
                      <th>Avg W</th>
                    </tr>
                  </thead>
                  <tbody></tbody>
                </table>
              </div>
              <div class="wv-card">
                <h3 style="margin: 0;">Selection</h3>
                <div class="wv-kv" id="wvSelectionKv"></div>
              </div>
            </div>
          </section>
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn ghost" id="deleteWorkoutView">Delete</button>
        <button class="btn ghost" id="cancelWorkoutView">Cancel</button>
        <button class="btn secondary" id="saveWorkoutView">Save</button>
        <button class="btn primary" id="saveCloseWorkoutView">Save &amp; Close</button>
      </div>
    </div>
  </div>

  <script>
    const WORKOUT_TYPES = [
      ['Run', 'run'], ['Bike', 'bike'], ['Swim', 'swim'], ['Brick', 'brick'],
      ['Crosstrain', 'pulse'], ['Day Off', 'rest'], ['Mtn Bike', 'mtb'], ['Strength', 'strength'],
      ['Custom', 'timer'], ['XC-Ski', 'ski'], ['Rowing', 'rowing'], ['Walk', 'walk'],
      ['Other', 'other'],
    ];

    const OTHER_TYPES = [
      ['Event', 'event', 'event'],
      ['Goals', 'goal', 'goal'],
      ['Note', 'note', 'note'],
      ['Metrics', 'metrics', 'metrics'],
      ['Availability', 'calendar', 'availability'],
    ];

    const ICON_ASSETS = {
      run: '/icons/workouts/run.png',
      bike: '/icons/workouts/bike.png',
      swim: '/icons/workouts/swim.png',
      brick: '/icons/workouts/brick.png',
      pulse: '/icons/workouts/crosstrain.png',
      rest: '/icons/workouts/day_off.png',
      mtb: '/icons/workouts/mountian_bike.png',
      strength: '/icons/workouts/strength.png',
      timer: '/icons/workouts/other_custom.png',
      ski: '/icons/workouts/XC_Ski.png',
      rowing: '/icons/workouts/row.png',
      walk: '/icons/workouts/walk.png',
      other: '/icons/workouts/other_custom.png',
      event: '/icons/workouts/event.png',
      goal: '/icons/workouts/goal.png',
      note: '/icons/workouts/note.png',
      metrics: '/icons/workouts/note.png',
      calendar: '/icons/workouts/note.png',
    };

    const DOW = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

    let activities = [];
    let calendarItems = [];
    let pairs = [];
    let selectedDate = todayKey();
    let selectedKind = 'workout';
    let selectedWorkoutType = 'Run';
    let editingItemId = null;
    let analyzeState = null;
    let initialMonthCentered = false;
    let appSettings = { units: { distance: 'km', elevation: 'm' }, ftp: {} };
    let distanceUnit = localStorage.getItem('distanceUnit') || 'km';
    let elevationUnit = localStorage.getItem('elevationUnit') || 'm';
    let currentFeel = 0;
    let fitUploadTargetActivityId = null;
    let fitUploadContext = 'global';
    let modalDraft = null;
    let calendarCursor = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    let calendarScrollBound = false;
    let calendarAnchorWeekStart = '';
    let calendarScrollTop = 0;

    function localDateKey(d) {
      const dt = new Date(d);
      const y = dt.getFullYear();
      const m = String(dt.getMonth() + 1).padStart(2, '0');
      const day = String(dt.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }

    function todayKey() {
      return localDateKey(new Date());
    }

    function parseDateKey(key) {
      return new Date(key + 'T00:00:00');
    }

    function dateKeyFromDate(d) {
      return localDateKey(d);
    }

    function mondayOfDate(dt) {
      const d = new Date(dt);
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
      return d;
    }

    function syncCalendarHeaderFromScroll() {
      const wrap = document.getElementById('calendarScroll');
      if (!wrap) return;
      const rows = Array.from(wrap.querySelectorAll('.week-row'));
      if (!rows.length) return;
      const top = wrap.getBoundingClientRect().top;
      let closest = null;
      let closestAbove = null;
      rows.forEach((row) => {
        const r = row.getBoundingClientRect();
        if (r.top <= top + 2) closestAbove = row;
        if (!closest && r.bottom >= top) closest = row;
      });
      const row = closestAbove || closest || rows[0];
      if (!row) return;
      document.getElementById('calHeaderMonth').textContent = row.dataset.weekLabel || '';
      const [yy, mm] = String(row.dataset.weekMonth || '').split('-');
      if (yy && mm) calendarCursor = new Date(Number(yy), Number(mm) - 1, 1);
      calendarAnchorWeekStart = String(row.dataset.weekStart || '');
      calendarScrollTop = wrap.scrollTop;
    }

    function toDisplayDistanceFromMeters(meters, unit = distanceUnit) {
      const m = Number(meters || 0);
      if (unit === 'm') return { value: m, unit: 'm' };
      if (unit === 'mi') return { value: m / 1609.344, unit: 'mi' };
      return { value: m / 1000, unit: 'km' };
    }

    function toDisplayDistanceFromKm(km, unit = distanceUnit) {
      return toDisplayDistanceFromMeters(Number(km || 0) * 1000, unit);
    }

    function fromDisplayDistanceToKm(val, unit = distanceUnit) {
      const n = Number(val || 0);
      if (!Number.isFinite(n)) return 0;
      if (unit === 'm') return n / 1000;
      if (unit === 'mi') return n * 1.609344;
      return n;
    }

    function fromDisplayDistanceToMeters(val, unit = distanceUnit) {
      return fromDisplayDistanceToKm(val, unit) * 1000;
    }

    function toDisplayElevationFromMeters(meters, unit = elevationUnit) {
      const m = Number(meters || 0);
      if (unit === 'ft') return { value: m * 3.28084, unit: 'ft' };
      return { value: m, unit: 'm' };
    }

    function fromDisplayElevationToMeters(val, unit = elevationUnit) {
      const n = Number(val || 0);
      if (!Number.isFinite(n)) return 0;
      if (unit === 'ft') return n / 3.28084;
      return n;
    }

    function fmtDistanceMeters(meters) {
      const d = toDisplayDistanceFromMeters(meters, distanceUnit);
      return `${d.value.toFixed(distanceUnit === 'm' ? 0 : 1)} ${d.unit}`;
    }

    function fmtDistanceMetersInUnit(meters, unit) {
      const useUnit = unit || distanceUnit;
      const d = toDisplayDistanceFromMeters(meters, useUnit);
      return `${d.value.toFixed(useUnit === 'm' ? 0 : 1)} ${d.unit}`;
    }

    function fmtDistanceKm(km) {
      const d = toDisplayDistanceFromKm(km, distanceUnit);
      return `${d.value.toFixed(distanceUnit === 'm' ? 0 : 1)} ${d.unit}`;
    }

    function fmtHours(seconds) {
      return ((seconds || 0) / 3600).toFixed(1) + ' h';
    }

    function fmtDateLabel(key) {
      return parseDateKey(key).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
    }

    function fmtDateUpper(key) {
      return parseDateKey(key).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }).toUpperCase();
    }

    function fmtElevation(meters) {
      const e = toDisplayElevationFromMeters(meters, elevationUnit);
      return `${Math.round(e.value)} ${e.unit}`;
    }

    function monthKey(year, month) {
      return `${year}-${String(month + 1).padStart(2, '0')}`;
    }

    function isCalendarActive() {
      const node = document.getElementById('view-calendar');
      return !!(node && node.classList.contains('active'));
    }

    function setView(name) {
      document.querySelectorAll('.tab').forEach(el => el.classList.toggle('active', el.dataset.view === name));
      document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
      document.getElementById('view-' + name).classList.add('active');
      const pageHead = document.querySelector('.page-head');
      if (pageHead) pageHead.classList.toggle('hidden', name === 'calendar');
      document.getElementById('pageTitle').textContent = name.charAt(0).toUpperCase() + name.slice(1);
      if (name === 'calendar') {
        if (!calendarAnchorWeekStart) {
          jumpToCurrentMonth();
          return;
        }
        renderCalendar({ forceAnchor: false, preserveScroll: true });
      }
    }

    function updateUnitButtons() {
      document.querySelectorAll('.distance-unit-label').forEach((el) => { el.textContent = distanceUnit; });
      document.querySelectorAll('.elevation-unit-label').forEach((el) => { el.textContent = elevationUnit; });
    }

    function intensityByType(type) {
      const map = {
        Run: 0.85, Bike: 0.82, Swim: 0.8, Brick: 0.9, Crosstrain: 0.7, 'Day Off': 0.2,
        'Mtn Bike': 0.86, Strength: 0.75, Custom: 0.72, 'XC-Ski': 0.88, Rowing: 0.84,
        Walk: 0.55, Other: 0.65, Ride: 0.82, Workout: 0.8,
      };
      return map[type] || 0.7;
    }

    function activitySportKey(activity) {
      const t = String(activity.type || activity.sport_key || '').toLowerCase();
      if (t.includes('ride') || t.includes('cycle') || t.includes('bike')) return 'ride';
      if (t.includes('run') || t.includes('walk')) return 'run';
      if (t.includes('swim')) return 'swim';
      if (t.includes('row')) return 'row';
      if (t.includes('strength')) return 'strength';
      return 'other';
    }

    function ftpForActivity(activity) {
      const key = activitySportKey(activity);
      const ftp = Number((appSettings.ftp || {})[key] || 0);
      return ftp > 0 ? ftp : null;
    }

    function estimateTss(durationMin, intensity) {
      const durH = Math.max(0, Number(durationMin || 0)) / 60;
      const ifac = Math.max(0.2, Number(intensity || 0.7));
      return Math.round(durH * ifac * ifac * 100);
    }

    function activityToTss(activity) {
      if (Number(activity.tss_override || 0) > 0) return Number(activity.tss_override);
      const ifv = activityIF(activity);
      const durationH = Number(activity.moving_time || 0) / 3600;
      if (ifv && durationH > 0) return durationH * ifv * ifv * 100;
      return estimateTss(Number(activity.moving_time || 0) / 60, intensityByType(activity.type || 'Other'));
    }

    function itemToTss(item) {
      if (item.kind !== 'workout') return 0;
      const plannedTss = Number(item.planned_tss || 0);
      if (plannedTss > 0) return plannedTss;
      const userIntensity = Number(item.intensity || 0);
      const intensity = userIntensity > 0 ? (0.4 + Math.min(10, userIntensity) / 10) : intensityByType(item.workout_type || 'Other');
      return estimateTss(Number(item.duration_min || 0), intensity);
    }

    function plannedIF(item) {
      const ifv = Number(item.planned_if || 0);
      if (ifv > 0) return ifv;
      const tss = Number(item.planned_tss || 0);
      const hours = Number(item.duration_min || 0) / 60;
      if (tss > 0 && hours > 0) return Math.sqrt(tss / (hours * 100));
      return null;
    }

    function completedIF(obj) {
      const ifv = Number(obj.if_value || obj.completed_if || 0);
      if (ifv > 0) return ifv;
      const tss = Number(obj.tss_override || obj.completed_tss || 0);
      const hours = Number(obj.moving_time || (Number(obj.completed_duration_min || 0) * 60) || 0) / 3600;
      if (tss > 0 && hours > 0) return Math.sqrt(tss / (hours * 100));
      const ftp = ftpForActivity(obj);
      const avgP = Number(obj.avg_power || 0);
      if (ftp && avgP > 0) return avgP / ftp;
      return null;
    }

    function activityIF(activity) {
      return completedIF(activity);
    }

    function pairForPlanned(plannedId) {
      return pairs.find(p => p.planned_id === plannedId) || null;
    }

    function pairForStrava(stravaId) {
      return pairs.find(p => p.strava_id === String(stravaId)) || null;
    }

    function plannedMetric(plannedItem, basis) {
      if (basis === 'distance') return Number(plannedItem.distance_km || 0);
      if (basis === 'tss') return itemToTss(plannedItem);
      return Number(plannedItem.duration_min || 0);
    }

    function completedFromPlanned(plannedItem) {
      const dur = Number(plannedItem.completed_duration_min || 0);
      const dist = Number(plannedItem.completed_distance_km || 0);
      const tss = Number(plannedItem.completed_tss || 0);
      const ifv = Number(plannedItem.completed_if || 0);
      if (dur <= 0 && dist <= 0 && tss <= 0 && ifv <= 0) return null;
      return {
        moving_time: dur * 60,
        distance: dist * 1000,
        tss_override: tss,
        if_value: ifv,
        type: plannedItem.workout_type || 'Workout',
      };
    }

    function completedMetric(completedItem, basis) {
      if (basis === 'distance') return Number(completedItem.distance || 0) / 1000;
      if (basis === 'tss') {
        const override = Number(completedItem.tss_override || 0);
        if (override > 0) return override;
        const ifv = Number(completedItem.if_value || 0);
        const h = Number(completedItem.moving_time || 0) / 3600;
        if (ifv > 0 && h > 0) return h * ifv * ifv * 100;
        return activityToTss(completedItem);
      }
      return Number(completedItem.moving_time || 0) / 60;
    }

    function complianceStatus(plannedItem, completedItem, dayKey) {
      const today = todayKey();
      if (!plannedItem && completedItem) return { cls: 'unplanned', arrow: '' };
      if (plannedItem && !completedItem) {
        if (dayKey < today) return { cls: 'paired-red', arrow: '' };
        return { cls: 'workout', arrow: '' };
      }
      if (!plannedItem || !completedItem) return { cls: 'workout', arrow: '' };
      const hasPlannedBasis = ['duration', 'distance', 'tss'].some((basis) => plannedMetric(plannedItem, basis) > 0);
      if (!hasPlannedBasis) return { cls: 'unplanned', arrow: '' };
      const bases = ['duration', 'distance', 'tss'];
      const pcts = [];
      for (const basis of bases) {
        const p = plannedMetric(plannedItem, basis);
        const c = completedMetric(completedItem, basis);
        if (p > 0) {
          pcts.push((c / p) * 100);
        }
      }
      if (!pcts.length) return { cls: 'unplanned', arrow: '' };
      const best = pcts.reduce((bestPct, pct) => Math.abs(pct - 100) < Math.abs(bestPct - 100) ? pct : bestPct, pcts[0]);
      if (best >= 80 && best <= 120) return { cls: 'paired-green', arrow: '' };
      if ((best >= 50 && best < 80) || (best > 120 && best <= 150)) {
        return { cls: 'paired-yellow', arrow: best > 120 ? 'up' : 'down' };
      }
      return { cls: 'paired-orange', arrow: best > 120 ? 'up' : 'down' };
    }

    function buildObservedDailyTssMap() {
      const map = {};
      const today = todayKey();

      activities.forEach(a => {
        const key = dateKeyFromDate(new Date(a.start_date_local));
        map[key] = (map[key] || 0) + activityToTss(a);
      });

      return map;
    }

    function buildMetricsToDate(endKey) {
      const observed = buildObservedDailyTssMap();
      const endDate = parseDateKey(endKey);
      const today = parseDateKey(todayKey());
      const start = new Date(endDate);
      start.setDate(endDate.getDate() - 119);
      const values = [];

      for (let i = 0; i < 120; i += 1) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        const key = dateKeyFromDate(d);
        values.push(d <= today ? Number(observed[key] || 0) : 0);
      }

      const ctlSeries = [];
      const atlSeries = [];
      const tsbSeries = [];
      let ctlPrev = 0;
      let atlPrev = 0;
      for (let i = 0; i < values.length; i += 1) {
        const tss = values[i];
        const tsb = ctlPrev - atlPrev;
        const ctl = ctlPrev + (tss - ctlPrev) / 42;
        const atl = atlPrev + (tss - atlPrev) / 7;
        tsbSeries.push(tsb);
        ctlSeries.push(ctl);
        atlSeries.push(atl);
        ctlPrev = ctl;
        atlPrev = atl;
      }

      return {
        ctl: Math.round(ctlSeries[ctlSeries.length - 1] || 0),
        atl: Math.round(atlSeries[atlSeries.length - 1] || 0),
        tsb: Math.round(tsbSeries[tsbSeries.length - 1] || 0),
        ctlSeries,
        atlSeries,
        tsbSeries,
      };
    }

    function renderSparkline(elId, series) {
      const el = document.getElementById(elId);
      el.innerHTML = '';
      if (!series.length) return;
      const recent = series.slice(-30);
      const maxVal = Math.max(...recent.map(v => Math.abs(v)), 1);
      recent.forEach(v => {
        const bar = document.createElement('span');
        bar.style.height = `${Math.max(3, Math.round((Math.abs(v) / maxVal) * 38))}px`;
        if (v < 0) bar.style.background = '#d5936f';
        el.appendChild(bar);
      });
    }

    function renderPerformanceMetrics() {
      const metrics = buildMetricsToDate(todayKey());
      document.getElementById('ctlVal').textContent = String(metrics.ctl);
      document.getElementById('atlVal').textContent = String(metrics.atl);
      document.getElementById('tsbVal').textContent = metrics.tsb > 0 ? `+${metrics.tsb}` : String(metrics.tsb);
      document.getElementById('ctlTrend').textContent = String(metrics.ctl);
      document.getElementById('atlTrend').textContent = String(metrics.atl);
      document.getElementById('tsbTrend').textContent = metrics.tsb > 0 ? `+${metrics.tsb}` : String(metrics.tsb);
      renderSparkline('ctlSpark', metrics.ctlSeries);
      renderSparkline('atlSpark', metrics.atlSeries);
      renderSparkline('tsbSpark', metrics.tsbSeries);
    }

    function iconSvg(name) {
      const src = ICON_ASSETS[name] || ICON_ASSETS.other;
      return `<span class="type-icon"><img src="${src}" alt="${name}"/></span>`;
    }

    function workoutIconKey(type) {
      const t = String(type || '').toLowerCase();
      if (t.includes('brick')) return 'brick';
      if (t.includes('cross')) return 'pulse';
      if (t.includes('day off') || t.includes('rest')) return 'rest';
      if (t.includes('mountain') || t.includes('mtn')) return 'mtb';
      if (t.includes('custom')) return 'timer';
      if (t.includes('ski')) return 'ski';
      if (t.includes('row')) return 'rowing';
      if (t.includes('run') || t.includes('walk')) return 'run';
      if (t.includes('swim')) return 'swim';
      if (t.includes('strength')) return 'strength';
      if (t.includes('ride') || t.includes('bike') || t.includes('cycling')) return 'bike';
      return 'other';
    }

    function cardIcon(type) {
      const key = workoutIconKey(type);
      const src = ICON_ASSETS[key] || ICON_ASSETS.other;
      return `<span style="display:inline-flex;align-items:center;margin-right:4px;"><img src="${src}" alt="${key}" style="width:18px;height:18px;object-fit:contain;" /></span>`;
    }

    function feelEmoji(v) {
      const map = { 1: '😫', 2: '🙁', 3: '😐', 4: '🙂', 5: '😁' };
      return map[Number(v)] || '';
    }

    function setFeelValue(v) {
      currentFeel = Number(v || 0);
      document.querySelectorAll('.feel-btn').forEach((btn) => {
        btn.classList.toggle('active', Number(btn.dataset.feel) === currentFeel);
      });
    }

    function commentsArrayFromEntity(entity) {
      if (!entity) return [];
      if (Array.isArray(entity.comments_feed)) {
        return entity.comments_feed.map((x) => String(x || '').trim()).filter(Boolean);
      }
      if (typeof entity.comments === 'string' && entity.comments.trim()) return [entity.comments.trim()];
      return [];
    }

    function renderCommentsFeed() {
      const wrap = document.getElementById('wvCommentsFeed');
      if (!modalDraft) {
        wrap.innerHTML = '<div class="meta">No comments yet.</div>';
        return;
      }
      const list = modalDraft.commentsFeed || [];
      if (!list.length) {
        wrap.innerHTML = '<div class="meta">No comments yet.</div>';
        return;
      }
      wrap.innerHTML = list.map((c) => `<div class="comment-item">${c.replace(/</g, '&lt;')}</div>`).join('');
    }

    function commentCount(entity) {
      return commentsArrayFromEntity(entity).length;
    }

    function formatStartClock(isoText) {
      if (!isoText) return '--:--';
      const dt = new Date(isoText);
      if (Number.isNaN(dt.getTime())) return '--:--';
      return dt.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).toLowerCase();
    }

    function parseDurationToMin(text) {
      const raw = String(text || '').trim();
      if (!raw) return 0;
      if (raw.includes(':')) {
        const parts = raw.split(':').map((x) => Number(x || 0));
        if (parts.length === 2) return parts[0] * 60 + parts[1];
        if (parts.length === 3) return parts[0] * 60 + parts[1] + (parts[2] / 60);
      }
      const n = Number(raw);
      return Number.isFinite(n) ? n : 0;
    }

    function formatDurationClock(mins) {
      const totalSec = Math.round(Math.max(0, Number(mins || 0)) * 60);
      const h = Math.floor(totalSec / 3600);
      const m = Math.floor((totalSec % 3600) / 60);
      const s = totalSec % 60;
      return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    }

    function toSportIcon(typeLabel) {
      const key = workoutIconKey(typeLabel);
      const src = ICON_ASSETS[key] || ICON_ASSETS.other;
      return `<span class="type-icon"><img src="${src}" alt="${key}" /></span>`;
    }

    function hasAnyCompletedMetric(obj, source) {
      if (source === 'strava') return true;
      return Number(obj.completed_duration_min || 0) > 0
        || Number(obj.completed_distance_m || 0) > 0
        || Number(obj.completed_distance_km || 0) > 0
        || Number(obj.completed_tss || 0) > 0
        || Number(obj.completed_if || 0) > 0;
    }

    function hasCompletedData(planned, completed) {
      if (completed) return true;
      if (!planned) return false;
      return Number(planned.completed_duration_min || 0) > 0
        || Number(planned.completed_distance_km || 0) > 0
        || Number(planned.completed_tss || 0) > 0;
    }

    function buildTypeGrids() {
      const workoutGrid = document.getElementById('workoutTypeGrid');
      workoutGrid.innerHTML = '';
      WORKOUT_TYPES.forEach(([name, icon]) => {
        const btn = document.createElement('button');
        btn.className = 'type-btn';
        btn.innerHTML = `${iconSvg(icon)}<span>${name}</span>`;
        btn.addEventListener('click', async () => {
          selectedKind = 'workout';
          selectedWorkoutType = name;
          const payload = {
            kind: 'workout',
            date: selectedDate || todayKey(),
            title: `Untitled ${name} Workout`,
            workout_type: name,
            duration_min: 0,
            distance_km: 0,
            intensity: 6,
          };
          const resp = await fetch('/calendar-items', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
          });
          if (!resp.ok) return;
          const created = await resp.json();
          closeActionModal();
          await loadData(false);
          openWorkoutModal({ source: 'planned', data: created });
        });
        workoutGrid.appendChild(btn);
      });

      const otherGrid = document.getElementById('otherTypeGrid');
      otherGrid.innerHTML = '';
      OTHER_TYPES.forEach(([name, icon, kind]) => {
        const btn = document.createElement('button');
        btn.className = 'type-btn';
        btn.innerHTML = `${iconSvg(icon)}<span>${name}</span>`;
        btn.addEventListener('click', () => {
          selectedKind = kind;
          selectedWorkoutType = 'Other';
          openDetailModal();
        });
        otherGrid.appendChild(btn);
      });
    }

    function openActionModal(dateKey, forcedKind) {
      selectedDate = dateKey || todayKey();
      document.getElementById('actionDateTitle').textContent = fmtDateLabel(selectedDate);
      document.getElementById('actionDateTitle').classList.toggle('small', window.innerWidth < 1200);
      document.getElementById('actionModal').classList.add('open');
      if (forcedKind === 'event') {
        selectedKind = 'event';
        openDetailModal();
      }
      if (forcedKind === 'goal') {
        selectedKind = 'goal';
        openDetailModal();
      }
    }

    function closeActionModal() {
      document.getElementById('actionModal').classList.remove('open');
    }

    function openDetailModal(existingItem) {
      closeActionModal();
      const metrics = buildMetricsToDate(selectedDate);
      document.getElementById('detailDateLabel').textContent = fmtDateUpper(selectedDate);
      document.getElementById('miniCtl').textContent = `Fitness ${metrics.ctl}`;
      document.getElementById('miniAtl').textContent = `Fatigue ${metrics.atl}`;
      document.getElementById('miniTsb').textContent = `Form ${metrics.tsb > 0 ? '+' + metrics.tsb : metrics.tsb}`;

      const titleMap = {
        workout: 'Workout Title',
        event: 'Event Name',
        goal: 'Goal',
        note: 'Note Title',
        metrics: 'Metrics Entry',
        availability: 'Availability Title',
      };

      document.getElementById('detailTitleLabel').textContent = titleMap[selectedKind] || 'Title';
      editingItemId = existingItem ? existingItem.id : null;
      document.getElementById('deleteDetail').style.visibility = editingItemId ? 'visible' : 'hidden';

      document.getElementById('dDate').value = existingItem ? existingItem.date : selectedDate;
      document.getElementById('dTitle').value = existingItem ? (existingItem.title || '') : '';
      const detailDesc = existingItem ? (existingItem.description || '') : '';
      document.getElementById('dDescription').value = detailDesc;
      document.getElementById('dDescriptionOther').value = detailDesc;
      document.getElementById('dWorkoutType').value = existingItem ? (existingItem.workout_type || selectedWorkoutType) : selectedWorkoutType;
      document.getElementById('dDuration').value = existingItem ? (existingItem.duration_min || '') : '';
      if (existingItem && Number(existingItem.distance_km || 0) > 0) {
        document.getElementById('dDistance').value = toDisplayDistanceFromKm(existingItem.distance_km).value.toFixed(distanceUnit === 'm' ? 0 : 1);
      } else {
        document.getElementById('dDistance').value = '';
      }
      document.getElementById('dIntensity').value = existingItem ? (existingItem.intensity || 6) : '6';
      document.getElementById('dEventType').value = existingItem ? (existingItem.event_type || 'Race') : 'Race';
      document.getElementById('dAvailability').value = existingItem ? (existingItem.availability || 'Unavailable') : 'Unavailable';

      document.getElementById('workoutFields').classList.toggle('hidden', selectedKind !== 'workout');
      document.getElementById('eventFields').classList.toggle('hidden', selectedKind !== 'event');
      document.getElementById('availabilityFields').classList.toggle('hidden', selectedKind !== 'availability');
      const isWorkout = selectedKind === 'workout';
      document.getElementById('detailMetricsChips').style.display = isWorkout ? 'contents' : 'none';
      document.querySelector('.detail-right').style.display = isWorkout ? 'block' : 'none';
      document.querySelector('.detail-body').style.gridTemplateColumns = isWorkout ? '1fr 340px' : '1fr';
      document.getElementById('nonWorkoutDescription').style.display = isWorkout ? 'none' : 'block';

      document.getElementById('detailModal').classList.add('open');
    }

    function closeDetailModal() {
      document.getElementById('detailModal').classList.remove('open');
    }

    async function saveDetail(closeAfter) {
      const payload = {
        kind: selectedKind,
        date: document.getElementById('dDate').value,
        title: document.getElementById('dTitle').value,
        description: selectedKind === 'workout'
          ? document.getElementById('dDescription').value
          : document.getElementById('dDescriptionOther').value,
      };

      if (selectedKind === 'workout') {
        payload.workout_type = document.getElementById('dWorkoutType').value || selectedWorkoutType;
        payload.duration_min = Number(document.getElementById('dDuration').value || 0);
        payload.distance_km = fromDisplayDistanceToKm(document.getElementById('dDistance').value);
        payload.intensity = Number(document.getElementById('dIntensity').value || 6);
      }

      if (selectedKind === 'event') {
        payload.event_type = document.getElementById('dEventType').value;
      }

      if (selectedKind === 'availability') {
        payload.availability = document.getElementById('dAvailability').value;
      }

      const url = editingItemId ? `/calendar-items/${editingItemId}` : '/calendar-items';
      const method = editingItemId ? 'PUT' : 'POST';
      const resp = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      if (!resp.ok) {
        const err = await resp.text();
        console.error('Could not save item:', err);
        return;
      }

      await loadData(false);
      if (closeAfter) {
        closeDetailModal();
      }
    }

    async function deleteCurrentDetail() {
      if (!editingItemId) return;
      const resp = await fetch(`/calendar-items/${editingItemId}`, { method: 'DELETE' });
      if (!resp.ok) return;
      closeDetailModal();
      await loadData(false);
    }

    function buildDayAggregateMap() {
      const map = {};
      const pairPlanned = new Set(pairs.map(p => String(p.planned_id)));

      activities.forEach(a => {
        const key = dateKeyFromDate(new Date(a.start_date_local));
        if (!map[key]) {
          map[key] = { done: [], items: [], durationMin: 0, tss: 0 };
        }
        map[key].done.push(a);
        map[key].durationMin += Number(a.moving_time || 0) / 60;
        map[key].tss += activityToTss(a);
      });

      calendarItems.forEach(item => {
        const key = item.date;
        if (!map[key]) {
          map[key] = { done: [], items: [], durationMin: 0, tss: 0 };
        }
        map[key].items.push(item);
        if (item.kind === 'workout' && !pairPlanned.has(String(item.id))) {
          const manual = completedFromPlanned(item);
          if (manual) {
            map[key].durationMin += Number(manual.moving_time || 0) / 60;
            map[key].tss += Number(manual.tss_override || 0) || activityToTss(manual);
          }
        }
      });

      return map;
    }

    function formatDurationMin(mins) {
      const total = Math.max(0, Math.round(mins));
      const h = Math.floor(total / 60);
      const m = total % 60;
      return `${h}:${String(m).padStart(2, '0')}`;
    }

    function getWeekMetrics(dateKeys, dayMap) {
      let durationMin = 0;
      let tss = 0;
      let weekEnd = null;

      dateKeys.forEach(key => {
        if (!key) return;
        weekEnd = key;
        const day = dayMap[key];
        if (!day) return;
        durationMin += day.durationMin;
        tss += day.tss;
      });

      const metrics = weekEnd ? buildMetricsToDate(weekEnd) : { ctl: 0, atl: 0, tsb: 0 };
      return {
        durationLabel: formatDurationMin(durationMin),
        tss: Math.round(tss),
        ctl: metrics.ctl,
        atl: metrics.atl,
        tsb: metrics.tsb,
      };
    }

    function closeContextMenu() {
      const menu = document.getElementById('contextMenu');
      menu.style.display = 'none';
      menu.innerHTML = '';
      menu.dataset.itemId = '';
    }

    function openContextMenu(x, y, options) {
      const menu = document.getElementById('contextMenu');
      menu.innerHTML = '';
      options.forEach(opt => {
        const btn = document.createElement('button');
        btn.textContent = opt.label;
        btn.addEventListener('click', async () => {
          closeContextMenu();
          await opt.onClick();
        });
        menu.appendChild(btn);
      });
      menu.style.left = `${x}px`;
      menu.style.top = `${y}px`;
      menu.style.display = 'block';
    }

    function showItemMenu(ev, payload) {
      ev.preventDefault();
      ev.stopPropagation();
      const opts = [];
      if (payload.source === 'planned') {
        opts.push({
          label: 'Edit',
          onClick: async () => {
            if (payload.data.kind === 'workout') {
              openWorkoutModal({ source: 'planned', data: payload.data, planned: payload.data });
            } else {
              selectedKind = payload.data.kind || 'workout';
              selectedDate = payload.data.date;
              selectedWorkoutType = payload.data.workout_type || 'Other';
              openDetailModal(payload.data);
            }
          },
        });
        opts.push({
          label: 'Copy',
          onClick: async () => {
            const copy = { ...payload.data };
            delete copy.id;
            delete copy.created_at;
            copy.title = `${copy.title || 'Copy'} (Copy)`;
            await fetch('/calendar-items', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(copy),
            });
            await loadData(false);
          },
        });
        opts.push({
          label: 'Delete',
          onClick: async () => {
            await fetch(`/calendar-items/${payload.data.id}`, { method: 'DELETE' });
            await loadData(false);
          },
        });
      }
      if (payload.source === 'strava') {
        opts.push({
          label: 'Delete',
          onClick: async () => {
            await fetch(`/activities/${payload.data.id}`, { method: 'DELETE' });
            await loadData(false);
          },
        });
      }

      const plannedId = payload.source === 'planned' ? payload.data.id : null;
      const stravaId = payload.source === 'strava' ? String(payload.data.id) : null;
      const currentPair = plannedId ? pairForPlanned(plannedId) : stravaId ? pairForStrava(stravaId) : null;
      if (currentPair) {
        opts.push({
          label: 'Unpair',
          onClick: async () => {
            await fetch(`/pairs/${currentPair.id}`, { method: 'DELETE' });
            await loadData(false);
          },
        });
      }

      if (!opts.length) return;
      openContextMenu(ev.clientX, ev.clientY, opts);
    }

    async function pairWorkouts(plannedId, stravaId) {
      if (!plannedId || !stravaId) return;
      const planned = calendarItems.find(i => String(i.id) === String(plannedId));
      const completed = activities.find(a => String(a.id) === String(stravaId));
      const typeLabel = (completed && completed.type) ? completed.type : (planned && planned.workout_type) ? planned.workout_type : 'Workout';
      const untitled = `Untitled ${typeLabel} Workout`;
      await fetch('/pairs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          planned_id: plannedId,
          strava_id: String(stravaId),
          override_date: planned ? planned.date : '',
          override_title: untitled,
        }),
      });
      await loadData(false);
    }

    function updateRpeLabel() {
      const slider = document.getElementById('wvRpe');
      const label = document.getElementById('wvRpeVal');
      if (!modalDraft || slider.disabled || !modalDraft.rpeTouched) {
        label.textContent = '';
        return;
      }
      label.textContent = String(slider.value || '--');
    }

    function setWorkoutMode(mode) {
      const analyzeBtn = document.getElementById('wvAnalyzeBtn');
      const showAnalyze = mode === 'analyze' && !analyzeBtn.classList.contains('disabled');
      document.getElementById('wvSummary').classList.toggle('hidden', showAnalyze);
      document.getElementById('wvAnalyze').classList.toggle('hidden', !showAnalyze);
      analyzeBtn.classList.toggle('active', showAnalyze);
      analyzeBtn.textContent = showAnalyze ? 'Summary' : 'Analyze';
      document.getElementById('wvFilesPopover').classList.add('hidden');
      document.getElementById('wvFilesTabBtn').classList.remove('active');
    }

    function syncUnitSelectValue(selectId, desired, allowed) {
      const el = document.getElementById(selectId);
      if (!el) return;
      if (allowed.includes(desired)) {
        el.value = desired;
      } else {
        el.value = allowed[0];
      }
    }

    function recalcIfTssRows() {
      const parseDurHours = (id) => parseDurationToMin(document.getElementById(id).value) / 60;
      const recalcPair = (durId, tssId, ifId) => {
        const hours = parseDurHours(durId);
        const tss = Number(document.getElementById(tssId).value || 0);
        const ifv = Number(document.getElementById(ifId).value || 0);
        if (hours <= 0) return;
        if (ifv > 0) {
          document.getElementById(tssId).value = (hours * ifv * ifv * 100).toFixed(1);
        } else if (tss > 0) {
          document.getElementById(ifId).value = Math.sqrt(tss / (hours * 100)).toFixed(2);
        }
      };
      recalcPair('pcDurPlan', 'pcTssPlan', 'pcIfPlan');
      recalcPair('pcDurComp', 'pcTssComp', 'pcIfComp');
    }

    function renderWorkoutFiles(payload) {
      const node = document.getElementById('wvFileList');
      const data = payload.data || {};
      const browseBtn = document.getElementById('wvBrowseFilesBtn');
      browseBtn.disabled = false;
      const rows = [];
      const hasExisting = !!data.fit_id;
      if (hasExisting && modalDraft && modalDraft.pendingDeleteFit) {
        rows.push(`
          <div class="wv-file-row">
            <div><strong>${data.fit_filename || `${data.fit_id}.fit`}</strong><div class="meta">Will be deleted on Save</div></div>
            <div class="wv-file-actions">
              <button class="btn secondary" id="wvUndoDeleteFitBtn">Undo</button>
            </div>
          </div>
        `);
      } else if (hasExisting) {
        const fileName = data.fit_filename || `${data.fit_id}.fit`;
        rows.push(`
          <div class="wv-file-row">
            <div><strong>${fileName}</strong><div class="meta">FIT attached</div></div>
            <div class="wv-file-actions">
              <button class="btn secondary" id="wvRecalcFitBtn">Recalculate</button>
              <button class="btn secondary" id="wvDeleteFitBtn">Delete</button>
              <a class="btn secondary" id="wvDownloadFitBtn" href="/activities/${encodeURIComponent(data.id)}/fit/download">Download</a>
            </div>
          </div>
        `);
      }
      if (!rows.length) {
        node.innerHTML = '<p class="meta">No files attached.</p>';
        return;
      }
      node.innerHTML = rows.join('');
      const undoBtn = document.getElementById('wvUndoDeleteFitBtn');
      if (undoBtn) {
        undoBtn.onclick = () => {
          if (!modalDraft) return;
          modalDraft.pendingDeleteFit = false;
          renderWorkoutFiles(payload);
        };
      }
      const recalcBtn = document.getElementById('wvRecalcFitBtn');
      if (recalcBtn) {
        recalcBtn.onclick = async () => {
          if (modalDraft && modalDraft.pendingDeleteFit) return;
          const resp = await fetch(`/activities/${data.id}/fit/recalculate`, { method: 'POST' });
          if (!resp.ok) return;
          const refreshed = await resp.json();
          window.currentWorkoutPayload = { ...payload, data: refreshed };
          renderWorkoutSummary(window.currentWorkoutPayload);
          await renderWorkoutAnalyze(window.currentWorkoutPayload);
          renderWorkoutFiles(window.currentWorkoutPayload);
          await loadData(false);
        };
      }
      const delBtn = document.getElementById('wvDeleteFitBtn');
      if (delBtn) {
        delBtn.onclick = () => {
          if (!modalDraft) return;
          modalDraft.pendingDeleteFit = true;
          renderWorkoutFiles(window.currentWorkoutPayload);
        };
      }
    }

    function openWorkoutModal(payload) {
      window.currentWorkoutPayload = payload;
      const data = payload.data || {};
      const parentPlanned = payload.planned || null;
      modalDraft = {
        pendingDeleteFit: false,
        uploadedNow: false,
        createdActivityId: null,
        createdPairId: null,
        originalFit: payload.source === 'strava' ? {
          fit_id: data.fit_id || null,
          fit_filename: data.fit_filename || null,
          distance: data.distance || 0,
          moving_time: data.moving_time || 0,
          avg_power: data.avg_power || null,
          avg_hr: data.avg_hr || null,
          min_hr: data.min_hr || null,
          max_hr: data.max_hr || null,
          min_power: data.min_power || null,
          max_power: data.max_power || null,
          elev_gain_m: data.elev_gain_m || null,
          if_value: data.if_value || null,
          tss_override: data.tss_override || null,
        } : null,
        commentsFeed: commentsArrayFromEntity(parentPlanned || data),
        sportType: parentPlanned
          ? (parentPlanned.workout_type || 'Other')
          : payload.source === 'strava' ? (data.type || 'Other') : (data.workout_type || 'Other'),
        rpeTouched: false,
      };
      const typeLabel = parentPlanned
        ? (parentPlanned.workout_type || 'Workout')
        : payload.source === 'strava' ? (data.type || 'Workout') : (data.workout_type || 'Workout');
      const dateLabel = parentPlanned
        ? `${parentPlanned.date} (Planned Day)`
        : payload.source === 'strava'
          ? new Date(data.start_date_local).toLocaleString()
          : `${data.date} (Planned)`;
      const dateKey = parentPlanned ? parentPlanned.date : payload.source === 'strava' ? dateKeyFromDate(new Date(data.start_date_local)) : data.date;
      const metrics = buildMetricsToDate(dateKey || todayKey());
      const dateText = parseDateKey(dateKey || todayKey()).toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }).toUpperCase();
      document.getElementById('wvDateLine').textContent = dateText;
      document.getElementById('wvHeadCtl').textContent = `Fitness ${metrics.ctl}`;
      document.getElementById('wvHeadAtl').textContent = `Fatigue ${metrics.atl}`;
      document.getElementById('wvHeadTsb').textContent = `Form ${metrics.tsb > 0 ? '+' + metrics.tsb : metrics.tsb}`;
      document.getElementById('wvTimeSelect').innerHTML = `<option>${payload.source === 'strava' ? formatStartClock(data.start_date_local) : '8:00 am'}</option>`;
      document.getElementById('wvTitle').value = parentPlanned ? (parentPlanned.title || 'Workout') : (data.title || data.name || 'Workout');
      const subNode = document.getElementById('wvSub');
      if (subNode) subNode.textContent = `${typeLabel} • ${dateLabel}`;
      modalDraft.sportType = WORKOUT_TYPES.some(([n]) => n === modalDraft.sportType) ? modalDraft.sportType : 'Other';
      document.getElementById('wvSportIcon').innerHTML = toSportIcon(modalDraft.sportType);
      document.getElementById('wvSportName').textContent = modalDraft.sportType;
      document.getElementById('wvCommentInput').value = '';

      const analyzeToggle = document.getElementById('wvAnalyzeBtn');
      const hasFile = !!(data.fit_id);
      analyzeToggle.classList.toggle('disabled', !hasFile);
      const sportMenu = document.getElementById('wvSportMenu');
      sportMenu.innerHTML = WORKOUT_TYPES.map(([name, icon]) => `
        <button class="wv-sport-item" type="button" data-sport="${name}">
          ${iconSvg(icon)}
          <span>${name}</span>
        </button>
      `).join('');
      sportMenu.querySelectorAll('.wv-sport-item').forEach((btn) => {
        btn.addEventListener('click', () => {
          if (!modalDraft) return;
          modalDraft.sportType = btn.dataset.sport || 'Other';
          document.getElementById('wvSportIcon').innerHTML = toSportIcon(modalDraft.sportType);
          document.getElementById('wvSportName').textContent = modalDraft.sportType;
          sportMenu.classList.add('hidden');
        });
      });
      if (hasFile) {
        renderWorkoutAnalyze(payload);
      } else {
        document.getElementById('wvAnalyze').classList.add('hidden');
        document.getElementById('wvSelectionKv').innerHTML = '<div>No FIT stream for this workout.</div>';
      }
      renderWorkoutSummary(payload);
      renderWorkoutFiles(payload);
      setWorkoutMode('summary');
      document.getElementById('wvFilesPopover').classList.add('hidden');
      document.getElementById('wvFilesTabBtn').classList.remove('active');
      document.getElementById('wvSportMenu').classList.add('hidden');
      document.getElementById('workoutViewModal').classList.add('open');
    }

    async function closeWorkoutModal(discard = true) {
      const payload = window.currentWorkoutPayload;
      if (discard && modalDraft && payload) {
        const data = payload.data || {};
        try {
          if (modalDraft.uploadedNow) {
            if (modalDraft.createdPairId) {
              await fetch(`/pairs/${modalDraft.createdPairId}`, { method: 'DELETE' });
            }
            if (modalDraft.createdActivityId) {
              await fetch(`/activities/${modalDraft.createdActivityId}`, { method: 'DELETE' });
            } else if (payload.source === 'strava') {
              const oldFit = modalDraft.originalFit || {};
              if (oldFit.fit_id) {
                await fetch(`/activities/${data.id}/fit/restore`, {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify(oldFit),
                });
              } else if (data.fit_id) {
                await fetch(`/activities/${data.id}/fit`, { method: 'DELETE' });
              }
            }
          }
        } catch (_err) {
          // best-effort rollback
        }
      }
      document.getElementById('workoutViewModal').classList.remove('open');
      modalDraft = null;
      window.currentWorkoutPayload = null;
      fitUploadContext = 'global';
      fitUploadTargetActivityId = null;
    }

    async function saveWorkoutView(closeAfter) {
      const payload = window.currentWorkoutPayload;
      if (!payload) return;
      const data = payload.data || {};
      const targetPlanned = payload.planned || (payload.source === 'planned' ? data : null);
      recalcIfTssRows();
      const description = document.getElementById('wvDescription').value;
      const commentsFeed = modalDraft ? modalDraft.commentsFeed.slice() : [];
      const comments = commentsFeed.length ? commentsFeed[commentsFeed.length - 1] : '';
      const sport = (modalDraft && modalDraft.sportType) || 'Other';
      const distanceUnitLocal = document.getElementById('pcDistanceUnit').value || 'km';
      const elevationUnitLocal = document.getElementById('pcElevationUnit').value || 'm';
      const plannedDuration = parseDurationToMin(document.getElementById('pcDurPlan').value);
      const plannedDistanceM = fromDisplayDistanceToMeters(document.getElementById('pcDistPlan').value, distanceUnitLocal);
      const plannedElevationM = fromDisplayElevationToMeters(document.getElementById('pcElevPlan').value, elevationUnitLocal);
      const plannedTss = Number(document.getElementById('pcTssPlan').value || 0);
      const plannedIf = Number(document.getElementById('pcIfPlan').value || 0);
      const completedDuration = parseDurationToMin(document.getElementById('pcDurComp').value);
      const completedDistanceM = fromDisplayDistanceToMeters(document.getElementById('pcDistComp').value, distanceUnitLocal);
      const completedElevationM = fromDisplayElevationToMeters(document.getElementById('pcElevComp').value, elevationUnitLocal);
      const completedTss = Number(document.getElementById('pcTssComp').value || 0);
      const completedIf = Number(document.getElementById('pcIfComp').value || 0);
      const hasCompleted = completedDuration > 0 || completedDistanceM > 0 || completedTss > 0 || payload.source === 'strava';
      const rpeVal = Number(document.getElementById('wvRpe').value || 0);
      const feel = hasCompleted ? currentFeel : 0;
      const rpeOut = (hasCompleted && modalDraft && modalDraft.rpeTouched) ? rpeVal : 0;

      let activityData = payload.source === 'strava' ? data : null;
      if (activityData && modalDraft && modalDraft.pendingDeleteFit && activityData.fit_id) {
        const delResp = await fetch(`/activities/${activityData.id}/fit`, { method: 'DELETE' });
        if (delResp.ok) activityData = await delResp.json();
      }

      if (targetPlanned) {
        await fetch(`/calendar-items/${targetPlanned.id}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ...targetPlanned,
            title: document.getElementById('wvTitle').value.trim() || targetPlanned.title || 'Untitled Workout',
            workout_type: sport,
            duration_min: plannedDuration,
            distance_km: plannedDistanceM / 1000,
            distance_m: plannedDistanceM,
            elevation_m: plannedElevationM,
            distance_unit: distanceUnitLocal,
            elevation_unit: elevationUnitLocal,
            planned_tss: plannedTss,
            planned_if: plannedIf,
            description,
            comments,
            comments_feed: commentsFeed,
            feel,
            rpe: rpeOut,
            completed_duration_min: completedDuration,
            completed_distance_km: completedDistanceM / 1000,
            completed_distance_m: completedDistanceM,
            completed_elevation_m: completedElevationM,
            completed_tss: completedTss,
            completed_if: completedIf,
          }),
        });
      } else if (payload.source === 'strava') {
        await fetch(`/activities/${data.id}/meta`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            description,
            comments,
            comments_feed: commentsFeed,
            feel,
            rpe: rpeOut,
            if_value: completedIf,
            tss_override: completedTss,
            title: document.getElementById('wvTitle').value.trim(),
            type: sport,
          }),
        });
      }
      await loadData(false);
      if (closeAfter) await closeWorkoutModal(false);
    }

    function num(v) {
      const n = Number(v);
      return Number.isFinite(n) ? n : null;
    }

    function timeToSec(iso, baseMs) {
      const t = new Date(iso).getTime();
      return Math.max(0, (t - baseMs) / 1000);
    }

    function hms(totalSec) {
      const s = Math.max(0, Math.round(totalSec));
      const h = Math.floor(s / 3600);
      const m = Math.floor((s % 3600) / 60);
      const sec = s % 60;
      return `${h}:${String(m).padStart(2, '0')}:${String(sec).padStart(2, '0')}`;
    }

    function fmtAxis(val, key) {
      if (key === 'speed') {
        if (distanceUnit === 'mi') return `${(val * 2.23694).toFixed(1)} mph`;
        if (distanceUnit === 'm') return `${val.toFixed(2)} m/s`;
        return `${(val * 3.6).toFixed(1)} km/h`;
      }
      if (key === 'distance') return fmtDistanceMeters(val);
      if (key === 'power') return `${Math.round(val)} W`;
      if (key === 'heart_rate') return `${Math.round(val)} bpm`;
      if (key === 'cadence') return `${Math.round(val)} rpm`;
      if (key === 'altitude') return fmtElevation(val);
      return String(Math.round(val));
    }

    function renderWorkoutSummary(payload) {
      const data = payload.data || {};
      const parentPlanned = payload.planned || null;
      const explicitCompleted = parentPlanned ? completedFromPlanned(parentPlanned) : completedFromPlanned(data);
      const completedDurationMin = payload.source === 'strava'
        ? Number(data.moving_time || 0) / 60
        : explicitCompleted ? Number(explicitCompleted.moving_time || 0) / 60 : 0;
      const completedDistanceM = payload.source === 'strava'
        ? Number(data.distance || 0)
        : explicitCompleted ? Number(explicitCompleted.distance || 0) : Number((parentPlanned || data).completed_distance_m || 0);
      const completedTss = payload.source === 'strava'
        ? activityToTss(data)
        : explicitCompleted ? Number(explicitCompleted.tss_override || 0) : 0;
      const typeLabel = parentPlanned
        ? (parentPlanned.workout_type || 'Workout')
        : payload.source === 'strava' ? (data.type || 'Workout') : (data.workout_type || 'Workout');
      const dateLabel = parentPlanned
        ? `${parentPlanned.date} (Planned Day)`
        : payload.source === 'strava'
          ? new Date(data.start_date_local).toLocaleString()
          : `${data.date} (Planned)`;
      const plannedObj = parentPlanned || data;
      let distanceUnitLocal = String(plannedObj.distance_unit || distanceUnit || 'km');
      let elevationUnitLocal = String(plannedObj.elevation_unit || elevationUnit || 'm');
      syncUnitSelectValue('pcDistanceUnit', distanceUnitLocal, ['km', 'mi', 'm']);
      syncUnitSelectValue('pcElevationUnit', elevationUnitLocal, ['m', 'ft']);

      document.getElementById('wvSummaryText').textContent = `${typeLabel} • ${dateLabel}`;
      document.getElementById('wvDescription').value = (parentPlanned && parentPlanned.description) || data.description || '';
      if (modalDraft) {
        if (!Array.isArray(modalDraft.commentsFeed) || !modalDraft.commentsFeed.length) {
          modalDraft.commentsFeed = commentsArrayFromEntity(parentPlanned || data);
        }
        renderCommentsFeed();
      }
      const savedRpe = Number((parentPlanned && parentPlanned.rpe) || data.rpe || 0);
      if (modalDraft) modalDraft.rpeTouched = savedRpe > 0;
      document.getElementById('wvRpe').value = String(savedRpe > 0 ? savedRpe : 1);
      document.getElementById('wvRpe').classList.toggle('rpe-unset', !(modalDraft && modalDraft.rpeTouched));
      setFeelValue((parentPlanned && parentPlanned.feel) || data.feel || 0);

      const plannedDuration = parentPlanned ? Number(parentPlanned.duration_min || 0) : Number(data.duration_min || 0);
      const plannedDistanceM = Number((plannedObj.distance_m || 0) || (Number(plannedObj.distance_km || 0) * 1000));
      const plannedElevationM = Number(plannedObj.elevation_m || 0);
      const completedElevationM = payload.source === 'strava'
        ? Number(data.elev_gain_m || 0)
        : Number(plannedObj.completed_elevation_m || 0);
      const plannedTss = Number(plannedObj.planned_tss || 0) || (parentPlanned ? itemToTss(parentPlanned) : itemToTss(data));
      const plannedIf = plannedIF(plannedObj);
      const completedIf = payload.source === 'strava' ? activityIF(data) : completedIF({
        completed_if: plannedObj.completed_if,
        completed_tss: completedTss,
        moving_time: completedDurationMin * 60,
        avg_power: data.avg_power,
        type: plannedObj.workout_type || data.type,
      });

      document.getElementById('pcDurPlan').value = plannedDuration ? formatDurationClock(plannedDuration) : '';
      document.getElementById('pcDurComp').value = completedDurationMin ? formatDurationClock(completedDurationMin) : '';
      const pd = toDisplayDistanceFromMeters(plannedDistanceM, distanceUnitLocal);
      const cd = toDisplayDistanceFromMeters(completedDistanceM, distanceUnitLocal);
      const pe = toDisplayElevationFromMeters(plannedElevationM, elevationUnitLocal);
      const ce = toDisplayElevationFromMeters(completedElevationM, elevationUnitLocal);
      document.getElementById('pcDistPlan').value = plannedDistanceM > 0 ? pd.value.toFixed(distanceUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcDistComp').value = completedDistanceM > 0 ? cd.value.toFixed(distanceUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcElevPlan').value = plannedElevationM > 0 ? pe.value.toFixed(elevationUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcElevComp').value = completedElevationM > 0 ? ce.value.toFixed(elevationUnitLocal === 'm' ? 0 : 1) : '';
      document.getElementById('pcTssPlan').value = plannedTss ? String(Math.round(plannedTss)) : '';
      document.getElementById('pcTssComp').value = completedTss ? String(Math.round(completedTss)) : '';
      document.getElementById('pcIfPlan').value = plannedIf ? Number(plannedIf).toFixed(2) : '';
      document.getElementById('pcIfComp').value = completedIf ? Number(completedIf).toFixed(2) : '';
      document.getElementById('wvHrMin').value = data.min_hr ? String(Math.round(data.min_hr)) : '';
      document.getElementById('wvHrAvg').value = data.avg_hr ? String(Math.round(data.avg_hr)) : '';
      document.getElementById('wvHrMax').value = data.max_hr ? String(Math.round(data.max_hr)) : '';
      document.getElementById('wvPowerMin').value = data.min_power ? String(Math.round(data.min_power)) : '';
      document.getElementById('wvPowerAvg').value = data.avg_power ? String(Math.round(data.avg_power)) : '';
      document.getElementById('wvPowerMax').value = data.max_power ? String(Math.round(data.max_power)) : '';
      if (data.fit_id) {
        fetch(`/fit/${data.fit_id}`).then(r => r.ok ? r.json() : null).then((fit) => {
          if (!fit) return;
          const s = fit.summary || {};
          if (s.min_hr) document.getElementById('wvHrMin').value = String(Math.round(s.min_hr));
          if (s.avg_hr) document.getElementById('wvHrAvg').value = String(Math.round(s.avg_hr));
          if (s.max_hr) document.getElementById('wvHrMax').value = String(Math.round(s.max_hr));
          if (s.min_power) document.getElementById('wvPowerMin').value = String(Math.round(s.min_power));
          if (s.avg_power) document.getElementById('wvPowerAvg').value = String(Math.round(s.avg_power));
          if (s.max_power) document.getElementById('wvPowerMax').value = String(Math.round(s.max_power));
        }).catch(() => {});
      }

      const hasCompleted = hasAnyCompletedMetric((parentPlanned || data), payload.source);
      document.querySelectorAll('.feel-btn').forEach((btn) => { btn.disabled = !hasCompleted; });
      document.getElementById('wvRpe').disabled = !hasCompleted;
      if (!hasCompleted && modalDraft) modalDraft.rpeTouched = false;
      document.getElementById('wvRpe').classList.toggle('rpe-unset', !hasCompleted || !(modalDraft && modalDraft.rpeTouched));
      updateRpeLabel();

      document.getElementById('wvHeaderDuration').textContent = completedDurationMin ? formatDurationClock(completedDurationMin) : '--:--:--';
      document.getElementById('wvHeaderDistance').textContent = completedDistanceM ? fmtDistanceMeters(completedDistanceM) : '--';
      document.getElementById('wvHeaderTss').textContent = completedTss ? `${Math.round(completedTss)} TSS` : '-- TSS';

      ['pcDurComp', 'pcDistComp', 'pcElevComp', 'pcTssComp', 'pcIfComp', 'pcDurPlan', 'pcDistPlan', 'pcElevPlan', 'pcTssPlan', 'pcIfPlan'].forEach((id) => {
        const el = document.getElementById(id);
        el.classList.toggle('muted', !el.value);
        el.oninput = () => el.classList.toggle('muted', !el.value);
      });

      const syncDistanceUnits = (nextUnit) => {
        const planMeters = fromDisplayDistanceToMeters(document.getElementById('pcDistPlan').value, distanceUnitLocal);
        const compMeters = fromDisplayDistanceToMeters(document.getElementById('pcDistComp').value, distanceUnitLocal);
        distanceUnitLocal = nextUnit;
        const nextPlan = toDisplayDistanceFromMeters(planMeters, nextUnit);
        const nextComp = toDisplayDistanceFromMeters(compMeters, nextUnit);
        document.getElementById('pcDistPlan').value = planMeters > 0 ? nextPlan.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
        document.getElementById('pcDistComp').value = compMeters > 0 ? nextComp.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
      };
      const syncElevationUnits = (nextUnit) => {
        const planMeters = fromDisplayElevationToMeters(document.getElementById('pcElevPlan').value, elevationUnitLocal);
        const compMeters = fromDisplayElevationToMeters(document.getElementById('pcElevComp').value, elevationUnitLocal);
        elevationUnitLocal = nextUnit;
        const nextPlan = toDisplayElevationFromMeters(planMeters, nextUnit);
        const nextComp = toDisplayElevationFromMeters(compMeters, nextUnit);
        document.getElementById('pcElevPlan').value = planMeters > 0 ? nextPlan.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
        document.getElementById('pcElevComp').value = compMeters > 0 ? nextComp.value.toFixed(nextUnit === 'm' ? 0 : 1) : '';
      };
      document.getElementById('pcDistanceUnit').onchange = (ev) => syncDistanceUnits(ev.target.value);
      document.getElementById('pcElevationUnit').onchange = (ev) => syncElevationUnits(ev.target.value);
      ['pcDurPlan', 'pcDurComp', 'pcTssPlan', 'pcTssComp', 'pcIfPlan', 'pcIfComp'].forEach((id) => {
        document.getElementById(id).onchange = recalcIfTssRows;
      });
    }

    async function renderWorkoutAnalyze(payload) {
      const data = payload.data || {};
      if (!data.fit_id) {
        document.getElementById('wvSelectionKv').innerHTML = '<div>No FIT stream for this workout.</div>';
        document.querySelector('#wvLapTable tbody').innerHTML = '';
        document.getElementById('wvChart').innerHTML = '';
        document.getElementById('wvViewfinder').innerHTML = '';
        return;
      }

      const resp = await fetch(`/fit/${data.fit_id}`);
      if (!resp.ok) {
        document.getElementById('wvSelectionKv').innerHTML = '<div>Could not load FIT data.</div>';
        return;
      }
      const fit = await resp.json();
      const series = Array.isArray(fit.series) ? fit.series : [];
      const laps = Array.isArray(fit.laps) ? fit.laps : [];
      const summary = fit.summary || {};
      if (!series.length) {
        document.getElementById('wvSelectionKv').innerHTML = '<div>No FIT points available.</div>';
        return;
      }

      const baseMs = new Date(series[0].timestamp).getTime();
      const pts = series.map((p) => ({
        t: timeToSec(p.timestamp, baseMs),
        heart_rate: num(p.heart_rate),
        speed: num(p.speed),
        distance: num(p.distance),
        cadence: num(p.cadence),
        power: num(p.power),
        altitude: num(p.altitude),
      }));
      const totalSec = Math.max(1, pts[pts.length - 1].t - pts[0].t);

      analyzeState = {
        pts,
        laps,
        totalSec,
        wStart: 0,
        wEnd: totalSec,
      };

      const chart = document.getElementById('wvChart');
      const finder = document.getElementById('wvViewfinder');
      const left = 54;
      const right = 170;
      const top = 14;
      const bottom = 28;
      const w = 1200;
      const h = 360;
      const cw = w - left - right;
      const ch = h - top - bottom;
      const lineMeta = [
        { key: 'heart_rate', color: '#f35353', label: 'HR' },
        { key: 'power', color: '#ff62f2', label: 'W' },
        { key: 'cadence', color: '#f39b1f', label: 'RPM' },
        { key: 'speed', color: '#3fa144', label: 'MPH' },
      ];

      function valPath(meta, inWindow) {
        const vals = inWindow.map(p => p[meta.key]).filter(v => v !== null);
        if (!vals.length) return { path: '', min: 0, max: 1, avg: null };
        const min = Math.min(...vals);
        const max = Math.max(...vals);
        const span = Math.max(0.001, max - min);
        const avg = vals.reduce((s, v) => s + v, 0) / vals.length;
        let started = false;
        let d = '';
        inWindow.forEach((p) => {
          const v = p[meta.key];
          if (v === null) return;
          const x = left + ((p.t - analyzeState.wStart) / Math.max(1, analyzeState.wEnd - analyzeState.wStart)) * cw;
          const y = top + (1 - ((v - min) / span)) * ch;
          d += `${started ? 'L' : 'M'}${x.toFixed(2)} ${y.toFixed(2)} `;
          started = true;
        });
        return { path: d.trim(), min, max, avg };
      }

      function inWindow() {
        return pts.filter(p => p.t >= analyzeState.wStart && p.t <= analyzeState.wEnd);
      }

      function renderSelectionStats() {
        const win = inWindow();
        const duration = Math.max(1, analyzeState.wEnd - analyzeState.wStart);
        const frac = duration / totalSec;
        const distance = Number(summary.distance_m || 0) * frac;
        const mean = (k) => {
          const vals = win.map(p => p[k]).filter(v => v !== null);
          return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null;
        };
        document.getElementById('wvSelectionKv').innerHTML = `
          <div>Duration<strong>${hms(duration)}</strong></div>
          <div>Distance<strong>${fmtDistanceMeters(distance)}</strong></div>
          <div>Avg HR<strong>${mean('heart_rate') ? Math.round(mean('heart_rate')) : '--'}</strong></div>
          <div>Avg Power<strong>${mean('power') ? `${Math.round(mean('power'))} W` : '--'}</strong></div>
          <div>Avg Cadence<strong>${mean('cadence') ? `${Math.round(mean('cadence'))} rpm` : '--'}</strong></div>
          <div>Avg Speed<strong>${mean('speed') ? fmtAxis(mean('speed'), 'speed') : '--'}</strong></div>
          <div>Elevation Gain<strong>${summary.elev_gain_m ? fmtElevation(summary.elev_gain_m) : '--'}</strong></div>
        `;
      }

      function renderMain() {
        const win = inWindow();
        const xTicks = 5;
        const paths = lineMeta.map(m => ({ ...m, ...valPath(m, win) }));
        let svg = `<rect x="0" y="0" width="${w}" height="${h}" fill="#f3f7fd" stroke="#d6e1ee"></rect>`;
        for (let i = 0; i <= xTicks; i += 1) {
          const x = left + (i / xTicks) * cw;
          svg += `<line x1="${x}" y1="${top}" x2="${x}" y2="${top + ch}" stroke="#e1eaf5"/>`;
          const sec = analyzeState.wStart + (i / xTicks) * (analyzeState.wEnd - analyzeState.wStart);
          svg += `<text x="${x}" y="${h - 8}" fill="#5b7290" font-size="11" text-anchor="middle">${hms(sec)}</text>`;
        }
        paths.forEach((p) => {
          if (p.path) svg += `<path d="${p.path}" stroke="${p.color}" stroke-width="2" fill="none"></path>`;
        });
        paths.forEach((p, idx) => {
          const y = top + 14 + idx * 24;
          svg += `<text x="${w - right + 6}" y="${y}" fill="${p.color}" font-size="11">${p.label} ${fmtAxis(p.max, p.key)} / ${fmtAxis(p.min, p.key)}</text>`;
        });
        chart.innerHTML = svg;
        renderSelectionStats();
      }

      function renderFinder() {
        const fw = 1200;
        const fh = 90;
        const px = (t) => (t / totalSec) * fw;
        const speedVals = pts.map(p => p.speed).filter(v => v !== null);
        const sMin = speedVals.length ? Math.min(...speedVals) : 0;
        const sMax = speedVals.length ? Math.max(...speedVals) : 1;
        const sSpan = Math.max(0.001, sMax - sMin);
        let d = '';
        let started = false;
        pts.forEach((p) => {
          if (p.speed === null) return;
          const x = px(p.t);
          const y = 6 + (1 - ((p.speed - sMin) / sSpan)) * (fh - 24);
          d += `${started ? 'L' : 'M'}${x.toFixed(2)} ${y.toFixed(2)} `;
          started = true;
        });
        const bx = px(analyzeState.wStart);
        const bw = Math.max(8, px(analyzeState.wEnd) - bx);
        finder.innerHTML = `
          <rect x="0" y="0" width="${fw}" height="${fh}" fill="#edf3fb" stroke="#d6e1ee"></rect>
          <path d="${d}" stroke="#3fa144" stroke-width="1.5" fill="none"></path>
          <rect id="wvBrush" x="${bx}" y="2" width="${bw}" height="${fh - 4}" fill="rgba(80,150,255,.22)" stroke="#2a66d2"></rect>
        `;
      }

      let dragging = false;
      let dragOffset = 0;
      finder.onmousedown = (ev) => {
        const rect = finder.getBoundingClientRect();
        const x = ((ev.clientX - rect.left) / rect.width) * 1200;
        const b = finder.querySelector('#wvBrush');
        const bx = Number(b.getAttribute('x'));
        const bw = Number(b.getAttribute('width'));
        if (x >= bx && x <= bx + bw) {
          dragging = true;
          dragOffset = x - bx;
        } else {
          const center = x / 1200;
          const span = (analyzeState.wEnd - analyzeState.wStart) / totalSec;
          let s = Math.max(0, center - span / 2);
          let e = Math.min(1, center + span / 2);
          if (e - s < span) s = Math.max(0, e - span);
          analyzeState.wStart = s * totalSec;
          analyzeState.wEnd = e * totalSec;
          renderFinder();
          renderMain();
        }
      };
      finder.onmousemove = (ev) => {
        if (!dragging) return;
        const rect = finder.getBoundingClientRect();
        const x = ((ev.clientX - rect.left) / rect.width) * 1200;
        const b = finder.querySelector('#wvBrush');
        const bw = Number(b.getAttribute('width'));
        let bx = x - dragOffset;
        bx = Math.max(0, Math.min(1200 - bw, bx));
        b.setAttribute('x', String(bx));
        analyzeState.wStart = (bx / 1200) * totalSec;
        analyzeState.wEnd = ((bx + bw) / 1200) * totalSec;
        renderMain();
      };
      window.onmouseup = () => { dragging = false; };

      const lapBody = document.querySelector('#wvLapTable tbody');
      lapBody.innerHTML = '';
      const lapRows = laps.length ? laps : [{
        name: 'Lap 1',
        start: series[0].timestamp,
        end: series[series.length - 1].timestamp,
        duration_s: totalSec,
      }];
      lapRows.forEach((lap, idx) => {
        const startSec = Math.max(0, timeToSec(lap.start || series[0].timestamp, baseMs));
        const endSec = Math.max(startSec, lap.end ? timeToSec(lap.end, baseMs) : (startSec + Number(lap.duration_s || 0)));
        const row = document.createElement('tr');
        row.innerHTML = `
          <td><input type="checkbox" /></td>
          <td>${lap.name || `Lap #${idx + 1}`}</td>
          <td>${hms(Number(lap.duration_s || (endSec - startSec)))}</td>
          <td>${lap.distance_m ? fmtDistanceMeters(lap.distance_m) : '--'}</td>
          <td>${lap.avg_hr ? Math.round(lap.avg_hr) : '--'}</td>
          <td>${lap.avg_power ? Math.round(lap.avg_power) : '--'}</td>
        `;
        row.querySelector('input').addEventListener('change', (ev) => {
          row.classList.toggle('selected', ev.target.checked);
          const checked = Array.from(lapBody.querySelectorAll('input')).map((el, i) => ({ checked: el.checked, i })).filter(x => x.checked);
          if (!checked.length) {
            analyzeState.wStart = 0;
            analyzeState.wEnd = totalSec;
          } else {
            const minI = Math.min(...checked.map(c => c.i));
            const maxI = Math.max(...checked.map(c => c.i));
            const minLap = lapRows[minI];
            const maxLap = lapRows[maxI];
            analyzeState.wStart = Math.max(0, timeToSec(minLap.start || series[0].timestamp, baseMs));
            analyzeState.wEnd = Math.max(analyzeState.wStart + 1, timeToSec(maxLap.end || series[series.length - 1].timestamp, baseMs));
          }
          renderFinder();
          renderMain();
        });
        lapBody.appendChild(row);
      });

      document.getElementById('wvHrMin').value = summary.min_hr ? String(Math.round(summary.min_hr)) : '';
      document.getElementById('wvHrAvg').value = summary.avg_hr ? String(Math.round(summary.avg_hr)) : '';
      document.getElementById('wvHrMax').value = summary.max_hr ? String(Math.round(summary.max_hr)) : '';
      document.getElementById('wvPowerMin').value = summary.min_power ? String(Math.round(summary.min_power)) : '';
      document.getElementById('wvPowerAvg').value = summary.avg_power ? String(Math.round(summary.avg_power)) : '';
      document.getElementById('wvPowerMax').value = summary.max_power ? String(Math.round(summary.max_power)) : '';
      renderFinder();
      renderMain();
    }

    function renderEvents() {
      const list = document.getElementById('eventsList');
      const events = calendarItems
        .filter(i => i.kind === 'event')
        .sort((a, b) => (a.date > b.date ? 1 : -1))
        .slice(0, 5);

      list.innerHTML = '';
      if (!events.length) {
        list.innerHTML = '<p class="meta">No events yet. Click + to add one.</p>';
        return;
      }

      events.forEach(e => {
        const node = document.createElement('div');
        node.className = 'event-item';
        node.innerHTML = `<h4>${e.title}</h4><p>${e.date} • ${e.event_type || 'Event'}</p>`;
        list.appendChild(node);
      });
    }

    function renderGoals() {
      const list = document.getElementById('goalsList');
      const goals = calendarItems
        .filter(i => i.kind === 'goal')
        .sort((a, b) => (a.date > b.date ? 1 : -1))
        .slice(0, 6);

      list.innerHTML = '';
      if (!goals.length) {
        list.innerHTML = '<p class="meta">No goals yet. Click Add Goal.</p>';
        return;
      }

      goals.forEach(g => {
        const node = document.createElement('div');
        node.className = 'goal-item';
        node.innerHTML = `<h4>${g.title}</h4><p>${g.date}</p>`;
        list.appendChild(node);
      });
    }

    function renderHome() {
      const today = todayKey();
      const doneToday = activities.filter(a => dateKeyFromDate(new Date(a.start_date_local)) === today);

      const plannedUpcoming = calendarItems
        .filter(i => i.kind === 'workout' && i.date >= today)
        .sort((a, b) => (a.date > b.date ? 1 : -1))
        .slice(0, 8);

      const doneNode = document.getElementById('todayDone');
      doneNode.innerHTML = '';
      if (!doneToday.length) {
        doneNode.innerHTML = '<p class="meta">No completed workouts for today yet.</p>';
      } else {
        doneToday.forEach(a => {
          const row = document.createElement('button');
          row.type = 'button';
          row.className = 'list-item';
          row.innerHTML = `
            <div>
              <p class="title">${a.name || 'Workout'}</p>
              <p class="meta">${a.type || 'Activity'} • ${fmtDistanceMeters(a.distance)} • ${fmtHours(a.moving_time)} • ${activityToTss(a)} TSS</p>
            </div>
            <span class="badge done">Done</span>
          `;
          row.addEventListener('click', () => openWorkoutModal({ source: 'strava', data: a }));
          row.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
          doneNode.appendChild(row);
        });
      }

      const plannedNode = document.getElementById('todayPlanned');
      plannedNode.innerHTML = '';
      if (!plannedUpcoming.length) {
        plannedNode.innerHTML = '<p class="meta">No planned workouts yet. Use + on any calendar day.</p>';
      } else {
        plannedUpcoming.forEach(p => {
          const pair = pairForPlanned(String(p.id));
          const pairedCompleted = pair ? activities.find(a => String(a.id) === String(pair.strava_id)) : null;
          const row = document.createElement('button');
          row.type = 'button';
          row.className = 'list-item';
          row.innerHTML = `
            <div>
              <p class="title">${p.title || p.workout_type}</p>
              <p class="meta">${p.date} • ${p.workout_type} • ${Number(p.duration_min || 0)} min • ${fmtDistanceKm(p.distance_km)} • ${itemToTss(p)} TSS</p>
            </div>
            <span class="badge planned">Planned</span>
          `;
          row.addEventListener('click', () => openWorkoutModal({ source: pairedCompleted ? 'strava' : 'planned', data: pairedCompleted || p, planned: p, pair }));
          row.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: p }));
          plannedNode.appendChild(row);
        });
      }

      renderEvents();
      renderGoals();
      renderPerformanceMetrics();
    }

    function renderCalendar(options = {}) {
      const { forceAnchor = false, preserveScroll = true, targetDateKey = '' } = options;
      const dayMap = buildDayAggregateMap();
      const plannedById = new Map(calendarItems.filter(i => i.kind === 'workout').map(i => [String(i.id), i]));
      const stravaById = new Map(activities.map(a => [String(a.id), a]));
      const pairByPlannedId = new Map(pairs.map(p => [String(p.planned_id), p]));
      const pairByStravaId = new Map(pairs.map(p => [String(p.strava_id), p]));
      const wrap = document.getElementById('calendarScroll');
      const prevScroll = preserveScroll ? wrap.scrollTop : 0;
      const priorRows = preserveScroll ? Array.from(wrap.querySelectorAll('.week-row')) : [];
      const wrapTop = wrap.getBoundingClientRect().top;
      let priorTopWeek = null;
      for (const row of priorRows) {
        if (row.getBoundingClientRect().bottom >= wrapTop + 2) {
          priorTopWeek = row;
          break;
        }
      }
      const preservedWeekStart = priorTopWeek ? String(priorTopWeek.dataset.weekStart || '') : '';
      wrap.innerHTML = '';
      const grid = document.createElement('section');
      grid.className = 'month';
      const dow = document.createElement('div');
      dow.className = 'dow';
      DOW.forEach((d) => {
        const el = document.createElement('span');
        el.textContent = d;
        dow.appendChild(el);
      });
      const sumHead = document.createElement('span');
      sumHead.className = 'sum-head';
      sumHead.textContent = 'Week Summary';
      dow.appendChild(sumHead);
      grid.appendChild(dow);

      const baseDate = targetDateKey ? parseDateKey(targetDateKey) : new Date(calendarCursor.getFullYear(), calendarCursor.getMonth(), 1);
      const startMonth = new Date(baseDate.getFullYear(), baseDate.getMonth() - 4, 1);
      const endMonth = new Date(baseDate.getFullYear(), baseDate.getMonth() + 9, 0);
      const startDate = new Date(startMonth);
      const startOffset = (startDate.getDay() + 6) % 7;
      startDate.setDate(startDate.getDate() - startOffset);
      const endDate = new Date(endMonth);
      const endOffset = (7 - ((endDate.getDay() + 6) % 7) - 1);
      endDate.setDate(endDate.getDate() + endOffset);

      const weekRows = [];
      for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 7)) {
        const row = document.createElement('div');
        row.className = 'week-row';
        const weekDateKeys = [];
        const weekStart = new Date(d);
        const weekMid = new Date(d);
        weekMid.setDate(weekMid.getDate() + 3);
        row.dataset.weekMonth = monthKey(weekMid.getFullYear(), weekMid.getMonth());
        row.dataset.weekLabel = weekMid.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
        row.dataset.weekStart = dateKeyFromDate(weekStart);

        for (let col = 0; col < 7; col += 1) {
          const dayDate = new Date(weekStart);
          dayDate.setDate(weekStart.getDate() + col);
          const key = dateKeyFromDate(dayDate);
          weekDateKeys.push(key);

          const cell = document.createElement('div');
          cell.className = 'day';
          cell.dataset.date = key;
          if (key === todayKey()) cell.classList.add('today');
          if (dayDate.getMonth() !== calendarCursor.getMonth()) cell.style.opacity = '0.92';

          const num = document.createElement('span');
          num.className = 'd-num';
          num.textContent = String(dayDate.getDate());
          cell.appendChild(num);

          const entries = dayMap[key] || { done: [], items: [] };
          const shownCompleted = new Set();
          const cardsToShow = [];
          entries.items.forEach((item) => {
            if (item.kind !== 'workout') {
              cardsToShow.push({ kind: 'other', item });
              return;
            }
            const pair = pairByPlannedId.get(String(item.id));
            const completed = pair ? stravaById.get(String(pair.strava_id)) : completedFromPlanned(item);
            if (completed) shownCompleted.add(String(completed.id));
            cardsToShow.push({ kind: 'planned', item, completed, pair, fromPair: !!pair });
          });
          entries.done.forEach((a) => {
            if (!shownCompleted.has(String(a.id))) {
              const pair = pairByStravaId.get(String(a.id));
              cardsToShow.push({ kind: 'completed', completed: a, pair });
            }
          });

          cardsToShow.slice(0, 6).forEach((entry) => {
            if (entry.kind === 'other') {
              const item = entry.item;
              const card = document.createElement('div');
              card.className = `work-card ${item.kind}`;
              card.innerHTML = `<button class="card-menu-btn" type="button">&#8942;</button><p class="wc-title">${item.title || item.kind.toUpperCase()}</p><p class="wc-meta">${item.kind.toUpperCase()} • ${item.date}</p>`;
              card.addEventListener('click', (ev) => {
                ev.stopPropagation();
                selectedKind = item.kind;
                selectedDate = item.date;
                selectedWorkoutType = item.workout_type || 'Other';
                openDetailModal(item);
              });
              card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              cell.appendChild(card);
              return;
            }
            if (entry.kind === 'planned') {
              const item = entry.item;
              const completed = entry.completed;
              const comp = complianceStatus(item, completed, key);
              const card = document.createElement('div');
              card.className = `work-card ${comp.cls}`;
              const cIf = completed ? completedIF(completed) : null;
              const distM = completed ? Number(completed.distance || 0) : Number(item.completed_distance_m || 0);
              const unitForCard = String(item.distance_unit || distanceUnit || 'km');
              const durMin = completed ? Number(completed.moving_time || 0) / 60 : Number(item.completed_duration_min || 0);
              const tssVal = completed ? Math.round(Number(completed.tss_override || 0) || activityToTss(completed)) : Math.round(Number(item.completed_tss || 0));
              const cTime = completed ? formatStartClock(completed.start_date_local) : '--:--';
              const arrow = comp.arrow === 'up' ? '<span class="delta-up">↑</span>' : comp.arrow === 'down' ? '<span class="delta-down">↓</span>' : '';
              const feedCount = commentCount(item) || commentCount(completed);
              const feedback = `${feelEmoji(item.feel)} ${Number(item.rpe || 0) > 0 ? `RPE ${item.rpe}` : ''}`.trim();
              card.innerHTML = `<button class="card-menu-btn" type="button">&#8942;</button><p class="wc-title">${cardIcon(item.workout_type)}${item.title || (item.workout_type || 'Workout')}</p><p class="wc-meta">${durMin ? formatDurationMin(durMin) : '--'} • ${distM ? fmtDistanceMetersInUnit(distM, unitForCard) : '--'} • ${tssVal || '--'} TSS ${cIf ? `• IF ${Number(cIf).toFixed(2)}` : ''}</p><p class="wc-meta">C ${cTime} ${arrow}</p><div class="wc-bottom"><span>${feedback || '&nbsp;'}</span><span>💬 ${feedCount}</span></div>`;
              card.draggable = true;
              card.dataset.kind = 'planned';
              card.dataset.plannedId = String(item.id);
              card.addEventListener('dragstart', (ev) => ev.dataTransfer.setData('text/plain', JSON.stringify({ source: 'planned', id: String(item.id) })));
              card.addEventListener('dragover', (ev) => ev.preventDefault());
              card.addEventListener('drop', async (ev) => {
                ev.preventDefault();
                const raw = ev.dataTransfer.getData('text/plain');
                if (!raw) return;
                const dragData = JSON.parse(raw);
                if (dragData.source === 'strava') await pairWorkouts(String(item.id), String(dragData.id));
              });
              card.addEventListener('click', (ev) => {
                ev.stopPropagation();
                const source = entry.fromPair && completed ? 'strava' : 'planned';
                openWorkoutModal({ source, data: source === 'strava' ? completed : item, planned: item, pair: entry.pair });
              });
              card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
              cell.appendChild(card);
              return;
            }
            const a = entry.completed;
            const pairedPlanned = entry.pair ? plannedById.get(String(entry.pair.planned_id)) : null;
            const compStat = pairedPlanned ? complianceStatus(pairedPlanned, a, key).cls : 'unplanned';
            const unitForCard = String((pairedPlanned && pairedPlanned.distance_unit) || a.distance_unit || distanceUnit || 'km');
            const card = document.createElement('div');
            card.className = `work-card ${compStat}`;
            const feedCount = commentCount(a);
            card.innerHTML = `<button class="card-menu-btn" type="button">&#8942;</button><p class="wc-title">${cardIcon(a.type)}${a.name || 'Completed Workout'}</p><p class="wc-meta">${formatDurationMin(Number(a.moving_time || 0) / 60)} • ${fmtDistanceMetersInUnit(Number(a.distance || 0), unitForCard)} • ${Math.round(activityToTss(a))} TSS</p><p class="wc-meta">C ${formatStartClock(a.start_date_local)} ${activityIF(a) ? `• IF ${Number(activityIF(a)).toFixed(2)}` : ''}</p><div class="wc-bottom"><span>${feelEmoji(a.feel)} ${Number(a.rpe || 0) > 0 ? `RPE ${a.rpe}` : ''}</span><span>💬 ${feedCount}</span></div>`;
            card.draggable = true;
            card.dataset.kind = 'strava';
            card.dataset.stravaId = String(a.id);
            card.addEventListener('dragstart', (ev) => ev.dataTransfer.setData('text/plain', JSON.stringify({ source: 'strava', id: String(a.id) })));
            card.addEventListener('dragover', (ev) => ev.preventDefault());
            card.addEventListener('drop', async (ev) => {
              ev.preventDefault();
              const raw = ev.dataTransfer.getData('text/plain');
              if (!raw) return;
              const dragData = JSON.parse(raw);
              if (dragData.source === 'planned') await pairWorkouts(String(dragData.id), String(a.id));
            });
            card.addEventListener('click', (ev) => {
              ev.stopPropagation();
              openWorkoutModal({ source: 'strava', data: a, pair: entry.pair });
            });
            card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
            card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
            cell.appendChild(card);
          });

          if (cardsToShow.length > 6) {
            const more = document.createElement('span');
            more.className = 'item';
            more.style.background = '#edf3fb';
            more.style.color = '#5c7898';
            more.textContent = `+${cardsToShow.length - 6} more`;
            cell.appendChild(more);
          }

          const addBar = document.createElement('button');
          addBar.type = 'button';
          addBar.className = 'quick-add';
          addBar.textContent = '+';
          addBar.title = 'Add item';
          addBar.addEventListener('click', (ev) => {
            ev.stopPropagation();
            openActionModal(key);
          });
          cell.appendChild(addBar);
          row.appendChild(cell);
        }

        const week = getWeekMetrics(weekDateKeys, dayMap);
        const weekCard = document.createElement('div');
        weekCard.className = 'week-summary';
        weekCard.innerHTML = `<div class="ws-metrics"><div class="ws-chip ws-ctl"><strong>${week.ctl}</strong>CTL</div><div class="ws-chip ws-atl"><strong>${week.atl}</strong>ATL</div><div class="ws-chip ws-tsb"><strong>${week.tsb > 0 ? '+' + week.tsb : week.tsb}</strong>TSB</div></div><div class="ws-row"><span>Total Duration</span><strong>${week.durationLabel}</strong></div><div class="ws-row"><span>Total TSS</span><strong>${week.tss}</strong></div>`;
        row.appendChild(weekCard);
        grid.appendChild(row);
        weekRows.push(row);
      }

      wrap.appendChild(grid);

      if (!calendarScrollBound) {
        wrap.addEventListener('scroll', syncCalendarHeaderFromScroll);
        calendarScrollBound = true;
      }
      const currentStartKey = dateKeyFromDate(new Date(baseDate.getFullYear(), baseDate.getMonth(), 1));
      let rowTarget = weekRows[0] || null;
      weekRows.forEach((row) => {
        if (String(row.dataset.weekStart || '') <= currentStartKey) rowTarget = row;
      });
      if (targetDateKey) {
        const targetCell = wrap.querySelector(`.day[data-date="${targetDateKey}"]`);
        const targetRow = targetCell ? targetCell.closest('.week-row') : null;
        if (targetRow) rowTarget = targetRow;
      }
      if (!forceAnchor) {
        const preferred = preservedWeekStart || calendarAnchorWeekStart;
        if (preferred) {
          const preferredRow = weekRows.find((row) => String(row.dataset.weekStart || '') === preferred);
          if (preferredRow) rowTarget = preferredRow;
        }
      }
      if (forceAnchor && rowTarget) {
        wrap.scrollTop = rowTarget.offsetTop;
        calendarAnchorWeekStart = String(rowTarget.dataset.weekStart || calendarAnchorWeekStart);
      } else if (preserveScroll && prevScroll > 0) {
        wrap.scrollTop = calendarScrollTop > 0 ? calendarScrollTop : prevScroll;
      } else if (rowTarget) {
        wrap.scrollTop = rowTarget.offsetTop;
      }
      syncCalendarHeaderFromScroll();
    }

    function jumpToCurrentMonth() {
      const now = new Date();
      calendarCursor = new Date(now.getFullYear(), now.getMonth(), 1);
      calendarAnchorWeekStart = dateKeyFromDate(mondayOfDate(now));
      calendarScrollTop = 0;
      renderCalendar({ forceAnchor: true, preserveScroll: false, targetDateKey: todayKey() });
    }

    function renderDashboard() {
      const totalDistance = activities.reduce((sum, a) => sum + Number(a.distance || 0), 0);
      const totalTime = activities.reduce((sum, a) => sum + Number(a.moving_time || 0), 0);
      document.getElementById('statCount').textContent = String(activities.length);
      document.getElementById('statPlanned').textContent = String(calendarItems.filter(i => i.kind === 'workout').length);
      document.getElementById('statDistance').textContent = fmtDistanceMeters(totalDistance);
      document.getElementById('statTime').textContent = fmtHours(totalTime);
    }

    function renderSettings() {
      const ftp = appSettings.ftp || {};
      const setVal = (id, key) => {
        const el = document.getElementById(id);
        if (!el) return;
        const v = ftp[key];
        el.value = v == null ? '' : String(v);
      };
      setVal('ftpRide', 'ride');
      setVal('ftpRun', 'run');
      setVal('ftpRow', 'row');
      setVal('ftpSwim', 'swim');
      setVal('ftpStrength', 'strength');
      setVal('ftpOther', 'other');
    }

    function applyWidgetPrefs() {
      const pref = JSON.parse(localStorage.getItem('dashboardWidgets') || '{}');
      ['count', 'plannedCount', 'distance', 'time'].forEach(key => {
        const visible = pref[key] !== false;
        const card = document.querySelector(`[data-widget="${key}"]`);
        const toggle = document.querySelector(`[data-toggle="${key}"]`);
        if (card) card.style.display = visible ? 'block' : 'none';
        if (toggle) toggle.checked = visible;
      });
    }

    function bindWidgetToggles() {
      document.querySelectorAll('[data-toggle]').forEach(input => {
        input.addEventListener('change', () => {
          const key = input.getAttribute('data-toggle');
          const pref = JSON.parse(localStorage.getItem('dashboardWidgets') || '{}');
          pref[key] = input.checked;
          localStorage.setItem('dashboardWidgets', JSON.stringify(pref));
          applyWidgetPrefs();
        });
      });
    }

    async function loadData(resetMonthPosition) {
      try {
        const [aResp, cResp, pResp, sResp] = await Promise.all([fetch('/ui/activities'), fetch('/calendar-items'), fetch('/pairs'), fetch('/settings')]);
        activities = aResp.ok ? await aResp.json() : [];
        calendarItems = cResp.ok ? await cResp.json() : [];
        pairs = pResp.ok ? await pResp.json() : [];
        appSettings = sResp.ok ? await sResp.json() : { units: { distance: 'km', elevation: 'm' }, ftp: {} };
        if (appSettings.units && appSettings.units.distance) {
          distanceUnit = appSettings.units.distance;
        }
        if (appSettings.units && appSettings.units.elevation) {
          elevationUnit = appSettings.units.elevation;
        }
      } catch (_err) {
        activities = [];
        calendarItems = [];
        pairs = [];
      }

      if (resetMonthPosition) initialMonthCentered = false;
      updateUnitButtons();
      renderHome();
      if (isCalendarActive()) {
        renderCalendar({ forceAnchor: resetMonthPosition, preserveScroll: !resetMonthPosition });
      }
      renderDashboard();
      renderSettings();
    }

    document.querySelectorAll('.tab').forEach(btn => {
      btn.addEventListener('click', () => setView(btn.dataset.view));
    });

    document.getElementById('saveSettingsBtn').addEventListener('click', async () => {
      const read = (id) => {
        const v = document.getElementById(id).value.trim();
        return v ? Number(v) : null;
      };
      const payload = {
        units: {
          distance: (appSettings.units && appSettings.units.distance) || distanceUnit || 'km',
          elevation: (appSettings.units && appSettings.units.elevation) || elevationUnit || 'm',
        },
        ftp: {
          ride: read('ftpRide'),
          run: read('ftpRun'),
          row: read('ftpRow'),
          swim: read('ftpSwim'),
          strength: read('ftpStrength'),
          other: read('ftpOther'),
        },
      };
      const resp = await fetch('/settings', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (!resp.ok) return;
      appSettings = await resp.json();
      const msg = document.getElementById('settingsSavedMsg');
      msg.style.display = 'inline';
      setTimeout(() => { msg.style.display = 'none'; }, 1400);
      await loadData(false);
    });
    document.getElementById('uploadFitBtn').addEventListener('click', () => {
      fitUploadContext = 'global';
      fitUploadTargetActivityId = null;
      document.getElementById('uploadFitInput').click();
    });
    document.getElementById('uploadFitInput').addEventListener('change', async (event) => {
      const input = event.target;
      const file = input.files && input.files[0];
      if (!file) return;
      input.value = '';
      if (fitUploadContext === 'modal') {
        const payload = window.currentWorkoutPayload;
        if (!payload || !modalDraft) return;
        const targetPlanned = payload.planned || (payload.source === 'planned' ? payload.data : null);
        if (payload.source === 'strava') {
          const resp = await fetch(`/activities/${encodeURIComponent(payload.data.id)}/fit/upload?filename=${encodeURIComponent(file.name)}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: file,
          });
          if (resp.ok) {
            const uploaded = await resp.json();
            modalDraft.uploadedNow = true;
            window.currentWorkoutPayload = { ...payload, data: uploaded };
            renderWorkoutSummary(window.currentWorkoutPayload);
            renderWorkoutFiles(window.currentWorkoutPayload);
            await renderWorkoutAnalyze(window.currentWorkoutPayload);
          }
        } else if (targetPlanned) {
          const upResp = await fetch(`/import-fit?filename=${encodeURIComponent(file.name)}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: file,
          });
          if (upResp.ok) {
            const uploaded = await upResp.json();
            const pairResp = await fetch('/pairs', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                planned_id: targetPlanned.id,
                strava_id: String(uploaded.id),
                override_date: targetPlanned.date,
                override_title: targetPlanned.title || 'Untitled Workout',
              }),
            });
            const pair = pairResp.ok ? await pairResp.json() : null;
            modalDraft.uploadedNow = true;
            modalDraft.createdActivityId = uploaded.id;
            modalDraft.createdPairId = pair ? pair.id : null;
            window.currentWorkoutPayload = { source: 'strava', data: uploaded, planned: targetPlanned, pair };
            renderWorkoutSummary(window.currentWorkoutPayload);
            renderWorkoutFiles(window.currentWorkoutPayload);
            await renderWorkoutAnalyze(window.currentWorkoutPayload);
          }
        }
        fitUploadTargetActivityId = null;
        fitUploadContext = 'global';
        return;
      }
      const endpoint = fitUploadTargetActivityId
        ? `/activities/${encodeURIComponent(fitUploadTargetActivityId)}/fit/upload?filename=${encodeURIComponent(file.name)}`
        : `/import-fit?filename=${encodeURIComponent(file.name)}`;
      const resp = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: file,
      });
      if (!resp.ok) {
        const err = await resp.text();
        alert(`Import failed: ${err}`);
        fitUploadTargetActivityId = null;
        fitUploadContext = 'global';
        return;
      }
      const uploaded = await resp.json();
      if (fitUploadTargetActivityId && window.currentWorkoutPayload && window.currentWorkoutPayload.data) {
        window.currentWorkoutPayload = { ...window.currentWorkoutPayload, data: uploaded };
        renderWorkoutFiles(window.currentWorkoutPayload);
        renderWorkoutSummary(window.currentWorkoutPayload);
        await renderWorkoutAnalyze(window.currentWorkoutPayload);
      }
      fitUploadTargetActivityId = null;
      fitUploadContext = 'global';
      await loadData(false);
    });
    document.getElementById('globalSettings').addEventListener('click', () => {
      document.getElementById('settingsModal').classList.add('open');
    });
    document.getElementById('closeSettings').addEventListener('click', () => {
      document.getElementById('settingsModal').classList.remove('open');
    });
    document.getElementById('settingsModal').addEventListener('click', (event) => {
      if (event.target.id === 'settingsModal') document.getElementById('settingsModal').classList.remove('open');
    });
    document.getElementById('addEventBtn').addEventListener('click', () => openActionModal(todayKey(), 'event'));
    document.getElementById('addGoalBtn').addEventListener('click', () => openActionModal(todayKey(), 'goal'));

    document.getElementById('closeAction').addEventListener('click', closeActionModal);
    document.getElementById('actionModal').addEventListener('click', (event) => {
      if (event.target.id === 'actionModal') closeActionModal();
    });

    document.getElementById('cancelDetail').addEventListener('click', closeDetailModal);
    document.getElementById('deleteDetail').addEventListener('click', deleteCurrentDetail);
    document.getElementById('saveDetail').addEventListener('click', () => saveDetail(false));
    document.getElementById('saveCloseDetail').addEventListener('click', () => saveDetail(true));
    document.getElementById('detailModal').addEventListener('click', (event) => {
      if (event.target.id === 'detailModal') closeDetailModal();
    });

    document.getElementById('closeWorkoutView').addEventListener('click', async () => { await closeWorkoutModal(true); await loadData(false); });
    document.getElementById('cancelWorkoutView').addEventListener('click', async () => { await closeWorkoutModal(true); await loadData(false); });
    document.getElementById('saveWorkoutView').addEventListener('click', () => saveWorkoutView(false));
    document.getElementById('saveCloseWorkoutView').addEventListener('click', () => saveWorkoutView(true));
    document.getElementById('deleteWorkoutView').addEventListener('click', async () => {
      const payload = window.currentWorkoutPayload;
      if (!payload) return;
      const data = payload.data || {};
      if (payload.planned || payload.source === 'planned') {
        const targetId = payload.planned ? payload.planned.id : data.id;
        await fetch(`/calendar-items/${targetId}`, { method: 'DELETE' });
      } else if (payload.source === 'strava') {
        await fetch(`/activities/${data.id}`, { method: 'DELETE' });
      }
      await closeWorkoutModal(false);
      await loadData(false);
    });
    document.getElementById('wvBrowseFilesBtn').addEventListener('click', () => {
      const payload = window.currentWorkoutPayload;
      if (!payload) return;
      fitUploadContext = 'modal';
      fitUploadTargetActivityId = payload.source === 'strava' ? payload.data.id : null;
      document.getElementById('uploadFitInput').click();
    });
    document.getElementById('wvFilesTabBtn').addEventListener('click', () => {
      const pop = document.getElementById('wvFilesPopover');
      const isHidden = pop.classList.contains('hidden');
      pop.classList.toggle('hidden', !isHidden);
      document.getElementById('wvFilesTabBtn').classList.toggle('active', isHidden);
    });
    document.getElementById('wvAnalyzeBtn').addEventListener('click', () => {
      const showingAnalyze = !document.getElementById('wvAnalyze').classList.contains('hidden');
      setWorkoutMode(showingAnalyze ? 'summary' : 'analyze');
    });
    document.getElementById('wvSportToggle').addEventListener('click', (ev) => {
      ev.stopPropagation();
      document.getElementById('wvSportMenu').classList.toggle('hidden');
    });
    document.querySelectorAll('.feel-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (btn.disabled) return;
        setFeelValue(btn.dataset.feel);
      });
    });
    document.getElementById('wvRpe').addEventListener('input', () => {
      if (modalDraft) modalDraft.rpeTouched = true;
      const rpe = document.getElementById('wvRpe');
      rpe.classList.remove('rpe-unset');
      updateRpeLabel();
    });
    document.getElementById('wvPostCommentBtn').addEventListener('click', () => {
      const input = document.getElementById('wvCommentInput');
      const txt = String(input.value || '').trim();
      if (!txt || !modalDraft) return;
      modalDraft.commentsFeed = modalDraft.commentsFeed || [];
      modalDraft.commentsFeed.push(txt);
      input.value = '';
      renderCommentsFeed();
    });
    document.getElementById('wvCommentInput').addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        document.getElementById('wvPostCommentBtn').click();
      }
    });
    document.getElementById('workoutViewModal').addEventListener('click', (event) => {
      if (event.target.id === 'workoutViewModal') {
        closeWorkoutModal(true).then(() => loadData(false));
      }
    });
    document.getElementById('contextMenu').addEventListener('click', (ev) => ev.stopPropagation());
    document.addEventListener('click', (ev) => {
      const sportMenu = document.getElementById('wvSportMenu');
      if (sportMenu && !sportMenu.classList.contains('hidden') && !ev.target.closest('#wvSportMenu') && !ev.target.closest('#wvSportToggle')) {
        sportMenu.classList.add('hidden');
      }
      if (!ev.target.closest('.card-menu-btn') && !ev.target.closest('#contextMenu')) {
        closeContextMenu();
      }
    });
    document.getElementById('calTodayBtn').addEventListener('click', jumpToCurrentMonth);

    buildTypeGrids();
    bindWidgetToggles();
    applyWidgetPrefs();
    updateUnitButtons();
    setView('home');
    loadData(true);
  </script>
</body>
</html>
    """


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
