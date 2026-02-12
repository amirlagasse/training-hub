import json
import os
import secrets
import threading
from datetime import date, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import requests
from dotenv import load_dotenv
from fastapi import Body, FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse

load_dotenv()

app = FastAPI()

TOKEN_FILE = Path("data/strava_tokens.json")
CALENDAR_FILE = Path("data/calendar_items.json")
PAIRS_FILE = Path("data/workout_pairs.json")
ACTIVITY_OVERRIDES_FILE = Path("data/activity_overrides.json")
IMPORTED_ACTIVITIES_FILE = Path("data/imported_activities.json")
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
            intensity = float(payload.get("intensity", 6) or 6)
        except (TypeError, ValueError) as err:
            raise HTTPException(status_code=400, detail="Workout values must be numeric.") from err

        item["workout_type"] = workout_type
        item["duration_min"] = max(0.0, duration)
        item["distance_km"] = max(0.0, distance)
        item["intensity"] = max(1.0, min(10.0, intensity))

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
  <title>Training Hub</title>
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
      font-family: "Segoe UI", Tahoma, sans-serif;
      background: linear-gradient(180deg, #f4f7fb 0%, var(--bg) 100%);
    }

    .top-nav {
      position: sticky;
      top: 0;
      z-index: 20;
      background: linear-gradient(180deg, #113154 0%, #0e2843 100%);
      border-bottom: 1px solid rgba(255,255,255,0.12);
      min-height: 56px;
      display: grid;
      grid-template-columns: 1fr auto 1fr;
      align-items: center;
      padding: 0 12px;
    }

    .brand {
      color: #d9e8f5;
      font-size: 16px;
      font-weight: 800;
      letter-spacing: 0.3px;
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
      padding: 8px 14px;
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
      font-size: 12px;
      font-weight: 700;
    }

    .nav-settings {
      border: 1px solid rgba(255, 255, 255, 0.25);
      background: rgba(255, 255, 255, 0.08);
      color: #eef4fb;
      border-radius: 8px;
      width: 28px;
      height: 28px;
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
      height: 74vh;
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
      min-height: 172px;
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
      gap: 12px;
      margin-top: 8px;
    }

    .type-btn {
      border: 1px solid #d3dfed;
      background: #fff;
      border-radius: 10px;
      padding: 14px 16px;
      font-size: 46px;
      text-align: left;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .type-btn .type-icon {
      width: 32px;
      height: 32px;
      text-align: center;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 8px;
      background: #eef4fb;
    }

    .type-btn .type-icon svg {
      width: 22px;
      height: 22px;
      stroke: #2a4b72;
      fill: none;
      stroke-width: 2;
      stroke-linecap: round;
      stroke-linejoin: round;
    }

    .type-btn:hover .type-icon {
      background: #dfeaff;
    }

    .type-btn:hover { background: #f7faff; border-color: #8dabdd; }

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
      align-items: center;
      padding: 12px 14px;
      border-bottom: 1px solid #e3ebf4;
      background: #f8fbff;
    }

    .wv-title {
      margin: 0;
      font-size: 24px;
      color: #1f2f43;
    }

    .wv-sub {
      margin: 2px 0 0;
      color: #5f758e;
      font-size: 12px;
    }

    .wv-tabs {
      display: flex;
      gap: 6px;
    }

    .wv-tab {
      border: 1px solid #c9d7e9;
      background: #fff;
      border-radius: 8px;
      padding: 6px 10px;
      color: #2c4f79;
      font-weight: 700;
      cursor: pointer;
    }

    .wv-tab.active {
      background: #1f5bd7;
      color: #fff;
      border-color: #1f5bd7;
    }

    .wv-tab.disabled {
      opacity: 0.45;
      cursor: not-allowed;
      pointer-events: none;
    }

    .wv-body {
      padding: 14px;
      max-height: 72vh;
      overflow: auto;
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

    .tp-unit {
      color: #24364c;
      font-size: 16px;
      text-align: left;
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
    <div class="brand">Training Hub</div>
    <nav class="tabs">
      <button class="tab active" data-view="home">Home</button>
      <button class="tab" data-view="calendar">Calendar</button>
      <button class="tab" data-view="dashboard">Dashboard</button>
    </nav>
    <div class="nav-right">
      <button class="import-btn" id="uploadFitBtn">Import FIT</button>
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
          <h3 style="margin: 0; color: #4f657d; text-transform: uppercase; letter-spacing: 0.35px; font-size: 13px;">Month-by-Month Calendar</h3>
          <button class="btn secondary" id="jumpToday">Jump to Current Month</button>
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
                    <div class="tp-unit">km</div>
                  </div>
                  <div class="tp-row">
                    <label>TSS</label>
                    <input class="tp-in readonly" readonly />
                    <input class="tp-in readonly" readonly />
                    <div class="tp-unit">TSS</div>
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
    <div class="modal-card" style="padding: 0; width: min(860px, 100%);">
      <div class="wv-shell">
        <div class="wv-top">
          <div>
            <h2 class="wv-title" id="wvTitle">Workout</h2>
            <p class="wv-sub" id="wvSub">Details</p>
          </div>
          <div class="wv-tabs">
            <button class="wv-tab active" data-wv-tab="summary">Summary</button>
            <button class="wv-tab" data-wv-tab="analyze">Analyze</button>
            <button class="wv-tab" id="wvUnpairBtn" style="display:none;">Unpair</button>
            <button class="wv-tab" id="wvDeleteBtn">Delete</button>
            <button class="icon-btn" id="closeWorkoutView">&times;</button>
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
                      <div class="tp-unit">km</div>
                    </div>
                    <div class="tp-row">
                      <label>TSS</label>
                      <input id="pcTssPlan" class="tp-in" />
                      <input id="pcTssComp" class="tp-in" />
                      <div class="tp-unit">TSS</div>
                    </div>
                    <div class="tp-minmax-head">
                      <div></div><div>Min</div><div>Avg</div><div>Max</div><div></div>
                    </div>
                    <div class="tp-minmax-row">
                      <label>Heart Rate</label>
                      <input class="tp-in" />
                      <input id="wvHrAvg" class="tp-in" />
                      <input class="tp-in" />
                      <div class="tp-unit">bpm</div>
                    </div>
                    <div class="tp-minmax-row">
                      <label>Power</label>
                      <input class="tp-in" />
                      <input id="wvPowerAvg" class="tp-in" />
                      <input class="tp-in" />
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
                      <div style="display:flex; gap:8px;">
                        <button class="wv-tab" data-feel="1">Very Weak</button>
                        <button class="wv-tab active" data-feel="3">Normal</button>
                        <button class="wv-tab" data-feel="5">Very Strong</button>
                      </div>
                    </div>
                    <div class="field">
                      <label>Rating of Perceived Exertion (RPE)</label>
                      <input id="wvRpe" type="range" min="1" max="10" value="6" />
                    </div>
                    <div class="field">
                      <label>Post-activity comments</label>
                      <textarea id="wvComments" style="min-height:120px;" placeholder="Enter comments"></textarea>
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
                </div>
                <table class="lap-table" id="wvLapTable">
                  <thead>
                    <tr>
                      <th></th>
                      <th>Segment</th>
                      <th>Duration</th>
                      <th>TSS</th>
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

    const ICONS = {
      run: '<svg viewBox="0 0 24 24"><path d="M13 4a2 2 0 1 0 0.01 0"/><path d="M8 12l4-2 2 2 3 1"/><path d="M7 18l3-4"/><path d="M12 13l-1 6"/><path d="M14 13l5 3"/></svg>',
      bike: '<svg viewBox="0 0 24 24"><circle cx="6" cy="17" r="3"/><circle cx="18" cy="17" r="3"/><path d="M7 17l4-6h4l2 6"/><path d="M10 9h3"/><path d="M14 7h2"/></svg>',
      swim: '<svg viewBox="0 0 24 24"><path d="M3 15c1.5 1 2.5 1 4 0s2.5-1 4 0 2.5 1 4 0 2.5-1 4 0"/><path d="M5 10l3-3 3 3"/><path d="M11 7l3 3"/></svg>',
      brick: '<svg viewBox="0 0 24 24"><rect x="3" y="8" width="18" height="8" rx="1"/><path d="M7 8v8M12 8v8M17 8v8"/></svg>',
      pulse: '<svg viewBox="0 0 24 24"><path d="M3 12h4l2-4 3 8 2-4h7"/></svg>',
      rest: '<svg viewBox="0 0 24 24"><rect x="3" y="9" width="18" height="7" rx="1"/><path d="M5 9V7M19 9V7M7 16v2M17 16v2"/></svg>',
      mtb: '<svg viewBox="0 0 24 24"><circle cx="6" cy="17" r="3"/><circle cx="18" cy="17" r="3"/><path d="M7 17l4-6 3 2 2 4"/><path d="M12 8l2-1"/></svg>',
      strength: '<svg viewBox="0 0 24 24"><path d="M4 10v4M8 9v6M16 9v6M20 10v4"/><path d="M8 12h8"/></svg>',
      timer: '<svg viewBox="0 0 24 24"><circle cx="12" cy="13" r="7"/><path d="M12 13V9M9 3h6"/></svg>',
      ski: '<svg viewBox="0 0 24 24"><path d="M5 19l6-14M13 19l6-14"/><path d="M3 21h8M13 21h8"/></svg>',
      rowing: '<svg viewBox="0 0 24 24"><path d="M3 18c2 2 16 2 18 0"/><path d="M9 8l4 4M13 12l3-5"/></svg>',
      walk: '<svg viewBox="0 0 24 24"><circle cx="13" cy="4" r="2"/><path d="M11 9l2-2 2 3"/><path d="M12 10l-2 4"/><path d="M14 12l4 2"/></svg>',
      other: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/><path d="M12 8v5M12 16h.01"/></svg>',
      event: '<svg viewBox="0 0 24 24"><path d="M8 4h8v4l-2 2 2 2v8H8v-8l2-2-2-2z"/></svg>',
      goal: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="7"/><circle cx="12" cy="12" r="3"/></svg>',
      note: '<svg viewBox="0 0 24 24"><path d="M6 3h9l3 3v15H6z"/><path d="M15 3v3h3"/></svg>',
      metrics: '<svg viewBox="0 0 24 24"><path d="M4 18h16"/><path d="M7 18v-6M12 18V8M17 18v-3"/></svg>',
      calendar: '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>',
    };
    const ICON_COLORS = {
      run: '#35a11a',
      bike: '#7a2ecf',
      swim: '#1697be',
      brick: '#9a3b32',
      pulse: '#b12b67',
      rest: '#60728f',
      mtb: '#7b5b1a',
      strength: '#50206d',
      timer: '#8d2fe2',
      ski: '#d66605',
      rowing: '#16a7c0',
      walk: '#3da220',
      other: '#8d2fe2',
      event: '#2151e0',
      goal: '#2151e0',
      note: '#5f6e89',
      metrics: '#5f6e89',
      calendar: '#5f6e89',
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

    function fmtDistanceMeters(meters) {
      return ((meters || 0) / 1000).toFixed(1) + ' km';
    }

    function fmtDistanceKm(km) {
      return Number(km || 0).toFixed(1) + ' km';
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

    function monthKey(year, month) {
      return `${year}-${String(month + 1).padStart(2, '0')}`;
    }

    function setView(name) {
      document.querySelectorAll('.tab').forEach(el => el.classList.toggle('active', el.dataset.view === name));
      document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
      document.getElementById('view-' + name).classList.add('active');
      document.getElementById('pageTitle').textContent = name.charAt(0).toUpperCase() + name.slice(1);
    }

    function intensityByType(type) {
      const map = {
        Run: 0.85, Bike: 0.82, Swim: 0.8, Brick: 0.9, Crosstrain: 0.7, 'Day Off': 0.2,
        'Mtn Bike': 0.86, Strength: 0.75, Custom: 0.72, 'XC-Ski': 0.88, Rowing: 0.84,
        Walk: 0.55, Other: 0.65, Ride: 0.82, Workout: 0.8,
      };
      return map[type] || 0.7;
    }

    function estimateTss(durationMin, intensity) {
      const durH = Math.max(0, Number(durationMin || 0)) / 60;
      const ifac = Math.max(0.2, Number(intensity || 0.7));
      return Math.round(durH * ifac * ifac * 100);
    }

    function activityToTss(activity) {
      return estimateTss(Number(activity.moving_time || 0) / 60, intensityByType(activity.type || 'Other'));
    }

    function itemToTss(item) {
      if (item.kind !== 'workout') return 0;
      const userIntensity = Number(item.intensity || 0);
      const intensity = userIntensity > 0 ? (0.4 + Math.min(10, userIntensity) / 10) : intensityByType(item.workout_type || 'Other');
      return estimateTss(Number(item.duration_min || 0), intensity);
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

    function completedMetric(completedItem, basis) {
      if (basis === 'distance') return Number(completedItem.distance || 0) / 1000;
      if (basis === 'tss') return activityToTss(completedItem);
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
      const bases = ['duration', 'distance', 'tss'];
      const pcts = [];
      for (const basis of bases) {
        const p = plannedMetric(plannedItem, basis);
        const c = completedMetric(completedItem, basis);
        if (p > 0) {
          pcts.push((c / p) * 100);
        }
      }
      if (!pcts.length) return { cls: 'paired-green', arrow: '' };
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
      const color = ICON_COLORS[name] || '#2a4b72';
      return `<span class="type-icon" style="color:${color}">${(ICONS[name] || ICONS.other).replace('<svg', '<svg style="stroke:currentColor"')}</span>`;
    }

    function buildTypeGrids() {
      const workoutGrid = document.getElementById('workoutTypeGrid');
      workoutGrid.innerHTML = '';
      WORKOUT_TYPES.forEach(([name, icon]) => {
        const btn = document.createElement('button');
        btn.className = 'type-btn';
        btn.innerHTML = `${iconSvg(icon)}<span>${name}</span>`;
        btn.addEventListener('click', () => {
          selectedKind = 'workout';
          selectedWorkoutType = name;
          openDetailModal();
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
      document.getElementById('dDistance').value = existingItem ? (existingItem.distance_km || '') : '';
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
        payload.distance_km = Number(document.getElementById('dDistance').value || 0);
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
            selectedKind = payload.data.kind || 'workout';
            selectedDate = payload.data.date;
            selectedWorkoutType = payload.data.workout_type || 'Other';
            openDetailModal(payload.data);
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

    function openWorkoutModal(payload) {
      window.currentWorkoutPayload = payload;
      const data = payload.data;
      const parentPlanned = payload.planned || null;
      const typeLabel = parentPlanned
        ? (parentPlanned.workout_type || 'Workout')
        : payload.source === 'strava' ? (data.type || 'Workout') : (data.workout_type || 'Workout');
      const dateLabel = parentPlanned
        ? `${parentPlanned.date} (Planned Day)`
        : payload.source === 'strava'
          ? new Date(data.start_date_local).toLocaleString()
          : `${data.date} (Planned)`;
      document.getElementById('wvTitle').textContent = parentPlanned ? (parentPlanned.title || 'Workout') : (data.title || data.name || 'Workout');
      document.getElementById('wvSub').textContent = `${typeLabel} • ${dateLabel}`;
      switchWorkoutTab('summary');
      renderWorkoutSummary(payload);
      const analyzeBtn = document.querySelector('.wv-tab[data-wv-tab="analyze"]');
      const hasFile = payload.source === 'strava' || !!payload.pair;
      analyzeBtn.classList.toggle('disabled', !hasFile);
      if (hasFile) {
        let analyzePayload = payload;
        if (payload.source !== 'strava' && payload.pair) {
          const paired = activities.find(a => String(a.id) === String(payload.pair.strava_id));
          if (paired) {
            analyzePayload = { ...payload, source: 'strava', data: paired };
          }
        }
        renderWorkoutAnalyze(analyzePayload);
      } else {
        document.getElementById('wvAnalyze').classList.add('hidden');
      }
      const unpairBtn = document.getElementById('wvUnpairBtn');
      const pair = payload.pair || (payload.planned ? pairForPlanned(payload.planned.id) : pairForStrava(String(data.id)));
      unpairBtn.style.display = pair ? 'inline-block' : 'none';
      unpairBtn.onclick = async () => {
        if (!pair) return;
        await fetch(`/pairs/${pair.id}`, { method: 'DELETE' });
        closeWorkoutModal();
        await loadData(false);
      };
      const deleteBtn = document.getElementById('wvDeleteBtn');
      deleteBtn.style.display = (payload.source === 'planned' || !!payload.planned || payload.source === 'strava') ? 'inline-block' : 'none';
      deleteBtn.onclick = async () => {
        if (payload.planned || payload.source === 'planned') {
          const targetId = payload.planned ? payload.planned.id : data.id;
          await fetch(`/calendar-items/${targetId}`, { method: 'DELETE' });
        } else if (payload.source === 'strava') {
          await fetch(`/activities/${data.id}`, { method: 'DELETE' });
        } else {
          return;
        }
        closeWorkoutModal();
        await loadData(false);
      };
      document.getElementById('workoutViewModal').classList.add('open');
    }

    function closeWorkoutModal() {
      document.getElementById('workoutViewModal').classList.remove('open');
    }

    function switchWorkoutTab(tab) {
      const analyzeBtn = document.querySelector('.wv-tab[data-wv-tab="analyze"]');
      if (tab === 'analyze' && analyzeBtn.classList.contains('disabled')) {
        return;
      }
      document.querySelectorAll('.wv-tab[data-wv-tab]').forEach(el => {
        el.classList.toggle('active', el.dataset.wvTab === tab);
      });
      document.getElementById('wvSummary').classList.toggle('hidden', tab !== 'summary');
      document.getElementById('wvAnalyze').classList.toggle('hidden', tab !== 'analyze');
    }

    function seededSeries(seed, count, min, max) {
      let x = seed % 2147483647;
      if (x <= 0) x += 2147483646;
      const out = [];
      for (let i = 0; i < count; i += 1) {
        x = (x * 48271) % 2147483647;
        const r = x / 2147483647;
        out.push(min + r * (max - min));
      }
      return out;
    }

    function pathFromSeries(series, width, height) {
      if (!series.length) return '';
      const min = Math.min(...series);
      const max = Math.max(...series);
      const span = Math.max(1, max - min);
      return series.map((v, i) => {
        const x = (i / (series.length - 1)) * width;
        const y = height - ((v - min) / span) * height;
        return `${i === 0 ? 'M' : 'L'}${x.toFixed(2)} ${y.toFixed(2)}`;
      }).join(' ');
    }

    function renderWorkoutSummary(payload) {
      const data = payload.data;
      const parentPlanned = payload.planned || null;
      const completedDurationMin = payload.source === 'strava'
        ? Number(data.moving_time || 0) / 60
        : Number(data.duration_min || 0);
      const completedDistanceKm = payload.source === 'strava'
        ? Number(data.distance || 0) / 1000
        : Number(data.distance_km || 0);
      const completedTss = payload.source === 'strava' ? activityToTss(data) : itemToTss(data);
      const typeLabel = parentPlanned
        ? (parentPlanned.workout_type || 'Workout')
        : payload.source === 'strava' ? (data.type || 'Workout') : (data.workout_type || 'Workout');
      const dateLabel = parentPlanned
        ? `${parentPlanned.date} (Planned Day)`
        : payload.source === 'strava'
          ? new Date(data.start_date_local).toLocaleString()
          : `${data.date} (Planned)`;
      document.getElementById('wvSummaryText').textContent = `${typeLabel} • ${dateLabel}`;
      document.getElementById('wvDescription').value = (parentPlanned && parentPlanned.description) || data.description || '';
      document.getElementById('wvComments').value = '';
      document.getElementById('wvRpe').value = '6';
      const plannedDuration = parentPlanned ? Number(parentPlanned.duration_min || 0) : Number(data.duration_min || 0);
      const plannedDistance = parentPlanned ? Number(parentPlanned.distance_km || 0) : Number(data.distance_km || 0);
      const plannedTss = parentPlanned ? itemToTss(parentPlanned) : itemToTss(data);
      document.getElementById('pcDurPlan').value = plannedDuration ? formatDurationMin(plannedDuration) : '--';
      document.getElementById('pcDurComp').value = completedDurationMin ? formatDurationMin(completedDurationMin) : '--';
      document.getElementById('pcDistPlan').value = plannedDistance ? `${plannedDistance.toFixed(1)}` : '--';
      document.getElementById('pcDistComp').value = completedDistanceKm ? `${completedDistanceKm.toFixed(1)}` : '--';
      document.getElementById('pcTssPlan').value = plannedTss ? String(plannedTss) : '--';
      document.getElementById('pcTssComp').value = completedTss ? String(completedTss) : '--';
      document.getElementById('wvHrAvg').value = String(Math.round(120 + completedTss * 0.5));
      document.getElementById('wvPowerAvg').value = String(Math.round(150 + completedTss * 1.8));
    }

    function renderWorkoutAnalyze(payload) {
      const data = payload.data;
      const tss = payload.source === 'strava' ? activityToTss(data) : itemToTss(data);
      const durationMin = payload.source === 'strava' ? Number(data.moving_time || 0) / 60 : Number(data.duration_min || 0);
      const totalMin = Math.max(60, Math.round(durationMin || 60));
      const points = totalMin * 6;
      const hr = [];
      const pwr = [];
      const cad = [];
      const spd = [];
      for (let i = 0; i < points; i += 1) {
        const t = i / 6;
        const block = Math.floor(t / 15);
        const phase = (t % 15) / 15;
        hr.push(148 + block * 2 + Math.sin(phase * 6.28) * 2);
        pwr.push(200 + block * 8 + Math.sin(phase * 6.28) * 12);
        cad.push(82 + block + Math.sin(phase * 6.28) * 1.5);
        spd.push(22 + block * 0.4 + Math.sin(phase * 6.28) * 0.6);
      }
      const svg = document.getElementById('wvChart');
      const width = 1120;
      const height = 280;
      svg.innerHTML = `
        <rect x="0" y="0" width="${width}" height="${height}" fill="#f3f7fd" stroke="#d6e1ee"></rect>
        <path d="${pathFromSeries(hr, width, height)}" stroke="#f35353" stroke-width="2" fill="none"></path>
        <path d="${pathFromSeries(pwr, width, height)}" stroke="#ff62f2" stroke-width="2" fill="none"></path>
        <path d="${pathFromSeries(cad, width, height)}" stroke="#f39b1f" stroke-width="2" fill="none"></path>
        <path d="${pathFromSeries(spd, width, height)}" stroke="#3fa144" stroke-width="2" fill="none"></path>
        <rect id="wvSelectionRect" x="0" y="0" width="0" height="${height}" fill="rgba(80,150,255,.22)" stroke="rgba(42,102,210,.5)" display="none"></rect>
      `;

      const lapBody = document.querySelector('#wvLapTable tbody');
      lapBody.innerHTML = '';
      const totalSec = Math.max(60, Math.round(durationMin * 60));
      const lapCount = 5;
      const lapSec = Math.round(totalSec / lapCount);
      analyzeState = {
        hr,
        pwr,
        cad,
        spd,
        points,
        totalSec,
        lapSec,
        selecting: false,
        startIndex: 0,
        endIndex: points - 1,
      };
      for (let i = 0; i < lapCount; i += 1) {
        const row = document.createElement('tr');
        row.innerHTML = `
          <td><input type="checkbox" /></td>
          <td>Lap #${i + 1}</td>
          <td>${formatDurationMin(lapSec / 60)}</td>
          <td>${Math.max(1, Math.round(tss / lapCount))}</td>
          <td>${Math.round(160 + (i - 2) * 3)}</td>
          <td>${Math.round(280 + (i - 2) * 6)}</td>
        `;
        row.querySelector('input').addEventListener('change', (ev) => {
          row.classList.toggle('selected', ev.target.checked);
          const checked = Array.from(lapBody.querySelectorAll('input')).map((el, idx) => ({ checked: el.checked, idx })).filter(x => x.checked);
          if (!checked.length) {
            updateAnalyzeSelection(0, analyzeState.points - 1);
            return;
          }
          const minLap = Math.min(...checked.map(c => c.idx));
          const maxLap = Math.max(...checked.map(c => c.idx));
          const start = Math.floor((minLap * analyzeState.points) / lapCount);
          const end = Math.floor(((maxLap + 1) * analyzeState.points) / lapCount) - 1;
          updateAnalyzeSelection(start, end);
        });
        lapBody.appendChild(row);
      }

      const distanceKm = payload.source === 'strava'
        ? Number(data.distance || 0) / 1000
        : Number(data.distance_km || 0);
      document.getElementById('wvSelectionKv').innerHTML = `
        <div>Duration<strong>${formatDurationMin(durationMin)}</strong></div>
        <div>Distance<strong>${distanceKm.toFixed(2)} km</strong></div>
        <div>TSS<strong>${tss}</strong></div>
        <div>Avg HR<strong>${Math.round(120 + tss * 0.5)}</strong></div>
        <div>Avg Power<strong>${Math.round(160 + tss * 1.6)} W</strong></div>
        <div>Cadence<strong>${Math.round(80 + tss * 0.09)} rpm</strong></div>
      `;

      const selRect = document.getElementById('wvSelectionRect');
      function xToIndex(clientX) {
        const rect = svg.getBoundingClientRect();
        const rel = Math.max(0, Math.min(1, (clientX - rect.left) / rect.width));
        return Math.round(rel * (analyzeState.points - 1));
      }
      svg.onmousedown = (ev) => {
        analyzeState.selecting = true;
        analyzeState.startIndex = xToIndex(ev.clientX);
        analyzeState.endIndex = analyzeState.startIndex;
        updateAnalyzeSelection(analyzeState.startIndex, analyzeState.endIndex);
      };
      window.onmouseup = () => {
        analyzeState.selecting = false;
      };
      svg.onmousemove = (ev) => {
        if (!analyzeState.selecting) return;
        analyzeState.endIndex = xToIndex(ev.clientX);
        updateAnalyzeSelection(analyzeState.startIndex, analyzeState.endIndex);
      };

      updateAnalyzeSelection(0, analyzeState.points - 1);

      function updateAnalyzeSelection(a, b) {
        const start = Math.max(0, Math.min(a, b));
        const end = Math.min(analyzeState.points - 1, Math.max(a, b));
        const x1 = (start / (analyzeState.points - 1)) * width;
        const x2 = (end / (analyzeState.points - 1)) * width;
        selRect.setAttribute('x', String(x1));
        selRect.setAttribute('width', String(Math.max(2, x2 - x1)));
        selRect.setAttribute('display', 'block');
        const segCount = Math.max(1, end - start + 1);
        const mean = (arr) => arr.slice(start, end + 1).reduce((s, v) => s + v, 0) / segCount;
        const selSec = Math.round((segCount / analyzeState.points) * analyzeState.totalSec);
        const totalDist = payload.source === 'strava'
          ? Number(data.distance || 0) / 1000
          : Number(data.distance_km || 0);
        const selDist = totalDist * (selSec / analyzeState.totalSec);
        const selTss = Math.round(tss * (selSec / analyzeState.totalSec));
        document.getElementById('wvSelectionKv').innerHTML = `
          <div>Duration<strong>${formatDurationMin(selSec / 60)}</strong></div>
          <div>Distance<strong>${selDist.toFixed(2)} km</strong></div>
          <div>TSS<strong>${selTss}</strong></div>
          <div>Avg HR<strong>${Math.round(mean(analyzeState.hr))}</strong></div>
          <div>Avg Power<strong>${Math.round(mean(analyzeState.pwr))} W</strong></div>
          <div>Cadence<strong>${Math.round(mean(analyzeState.cad))} rpm</strong></div>
        `;
      }
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

    function renderCalendar() {
      const dayMap = buildDayAggregateMap();
      const plannedById = new Map(calendarItems.filter(i => i.kind === 'workout').map(i => [String(i.id), i]));
      const stravaById = new Map(activities.map(a => [String(a.id), a]));
      const pairByPlannedId = new Map(pairs.map(p => [String(p.planned_id), p]));
      const pairByStravaId = new Map(pairs.map(p => [String(p.strava_id), p]));
      const now = new Date();
      const currentMonthKey = monthKey(now.getFullYear(), now.getMonth());
      const wrap = document.getElementById('calendarScroll');
      wrap.innerHTML = '';

      for (let offset = -6; offset <= 9; offset += 1) {
        const base = new Date(now.getFullYear(), now.getMonth() + offset, 1);
        const y = base.getFullYear();
        const m = base.getMonth();
        const daysInMonth = new Date(y, m + 1, 0).getDate();
        const firstDayJs = new Date(y, m, 1).getDay();
        const firstDayMon = (firstDayJs + 6) % 7;

        const month = document.createElement('section');
        month.className = 'month';
        month.dataset.month = monthKey(y, m);
        if (month.dataset.month === currentMonthKey) month.classList.add('current-month');

        const title = document.createElement('h4');
        title.className = 'month-title';
        title.textContent = base.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
        month.appendChild(title);

        const dow = document.createElement('div');
        dow.className = 'dow';
        DOW.forEach(d => {
          const el = document.createElement('span');
          el.textContent = d;
          dow.appendChild(el);
        });
        const sumHead = document.createElement('span');
        sumHead.className = 'sum-head';
        sumHead.textContent = 'Week Summary';
        dow.appendChild(sumHead);
        month.appendChild(dow);

        const totalSlots = Math.ceil((firstDayMon + daysInMonth) / 7) * 7;
        let cursorDay = 1;

        for (let slot = 0; slot < totalSlots; slot += 7) {
          const row = document.createElement('div');
          row.className = 'week-row';
          const weekDateKeys = [];

          for (let col = 0; col < 7; col += 1) {
            const globalIndex = slot + col;
            const dayNum = globalIndex >= firstDayMon && cursorDay <= daysInMonth ? cursorDay : null;

            if (dayNum) {
              const key = `${y}-${String(m + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
              weekDateKeys.push(key);

              const cell = document.createElement('div');
              cell.className = 'day';
              if (key === todayKey()) cell.classList.add('today');

              const num = document.createElement('span');
              num.className = 'd-num';
              num.textContent = String(dayNum);
              cell.appendChild(num);

              const entries = dayMap[key] || { done: [], items: [] };

              const shownCompleted = new Set();
              const cardsToShow = [];

              entries.items.forEach(item => {
                if (item.kind !== 'workout') {
                  cardsToShow.push({ kind: 'other', item });
                  return;
                }
                const pair = pairByPlannedId.get(String(item.id));
                const completed = pair ? stravaById.get(String(pair.strava_id)) : null;
                if (completed) shownCompleted.add(String(completed.id));
                cardsToShow.push({ kind: 'planned', item, completed, pair });
              });

              entries.done.forEach(a => {
                if (!shownCompleted.has(String(a.id))) {
                  const pair = pairByStravaId.get(String(a.id));
                  cardsToShow.push({ kind: 'completed', completed: a, pair });
                }
              });

              cardsToShow.slice(0, 3).forEach((entry) => {
                if (entry.kind === 'other') {
                  const item = entry.item;
                  const card = document.createElement('div');
                  card.className = `work-card ${item.kind}`;
                  card.innerHTML = `
                    <button class="card-menu-btn" type="button">&#8942;</button>
                    <p class="wc-title">${item.title || item.kind.toUpperCase()}</p>
                    <p class="wc-meta">${item.kind.toUpperCase()} • ${item.date}</p>
                  `;
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
                  const status = comp.cls;
                  const card = document.createElement('div');
                  card.className = `work-card ${status}`;
                  const plannedLine = `P ${Number(item.duration_min || 0)}m • ${itemToTss(item)} TSS`;
                  const completedLine = completed
                    ? `C ${formatDurationMin(Number(completed.moving_time || 0) / 60)} • ${activityToTss(completed)} TSS`
                    : 'C --';
                  const arrow = comp.arrow === 'up' ? '<span class="delta-up">↑</span>' : comp.arrow === 'down' ? '<span class="delta-down">↓</span>' : '';
                  card.innerHTML = `
                    <button class="card-menu-btn" type="button">&#8942;</button>
                    <p class="wc-title">${item.title || (item.workout_type || 'Workout')}</p>
                    <p class="wc-meta">${item.workout_type || 'Workout'} • ${plannedLine}</p>
                    <p class="wc-meta">${completedLine} ${arrow}</p>
                  `;
                  card.draggable = true;
                  card.dataset.kind = 'planned';
                  card.dataset.plannedId = String(item.id);
                  card.addEventListener('dragstart', (ev) => {
                    ev.dataTransfer.setData('text/plain', JSON.stringify({ source: 'planned', id: String(item.id) }));
                  });
                  card.addEventListener('dragover', (ev) => ev.preventDefault());
                  card.addEventListener('drop', async (ev) => {
                    ev.preventDefault();
                    const raw = ev.dataTransfer.getData('text/plain');
                    if (!raw) return;
                    const data = JSON.parse(raw);
                    if (data.source === 'strava') {
                      await pairWorkouts(String(item.id), String(data.id));
                    }
                  });
                  card.addEventListener('click', (ev) => {
                    ev.stopPropagation();
                    openWorkoutModal({ source: completed ? 'strava' : 'planned', data: completed || item, planned: item, pair: entry.pair });
                  });
                  card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
                  card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'planned', data: item }));
                  cell.appendChild(card);
                  return;
                }

                const a = entry.completed;
                const pairedPlanned = entry.pair ? plannedById.get(String(entry.pair.planned_id)) : null;
                const compStat = pairedPlanned ? complianceStatus(pairedPlanned, a, key).cls : 'unplanned';
                const card = document.createElement('div');
                card.className = `work-card ${compStat}`;
                card.innerHTML = `
                  <button class="card-menu-btn" type="button">&#8942;</button>
                  <p class="wc-title">${a.name || 'Completed Workout'}</p>
                  <p class="wc-meta">${a.type || 'Workout'} • ${formatDurationMin(Number(a.moving_time || 0) / 60)} • ${activityToTss(a)} TSS</p>
                `;
                card.draggable = true;
                card.dataset.kind = 'strava';
                card.dataset.stravaId = String(a.id);
                card.addEventListener('dragstart', (ev) => {
                  ev.dataTransfer.setData('text/plain', JSON.stringify({ source: 'strava', id: String(a.id) }));
                });
                card.addEventListener('dragover', (ev) => ev.preventDefault());
                card.addEventListener('drop', async (ev) => {
                  ev.preventDefault();
                  const raw = ev.dataTransfer.getData('text/plain');
                  if (!raw) return;
                  const data = JSON.parse(raw);
                  if (data.source === 'planned') {
                    await pairWorkouts(String(data.id), String(a.id));
                  }
                });
                card.addEventListener('click', (ev) => {
                  ev.stopPropagation();
                  openWorkoutModal({ source: 'strava', data: a, pair: entry.pair });
                });
                card.addEventListener('contextmenu', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
                card.querySelector('.card-menu-btn').addEventListener('click', (ev) => showItemMenu(ev, { source: 'strava', data: a }));
                cell.appendChild(card);
              });

              const allCount = cardsToShow.length;
              if (allCount > 3) {
                const more = document.createElement('span');
                more.className = 'item';
                more.style.background = '#edf3fb';
                more.style.color = '#5c7898';
                more.textContent = `+${allCount - 3} more`;
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
              cursorDay += 1;
            } else {
              weekDateKeys.push(null);
              const empty = document.createElement('div');
              empty.className = 'day empty';
              row.appendChild(empty);
            }
          }

          const week = getWeekMetrics(weekDateKeys, dayMap);
          const weekCard = document.createElement('div');
          weekCard.className = 'week-summary';
          weekCard.innerHTML = `
            <div class="ws-metrics">
              <div class="ws-chip ws-ctl"><strong>${week.ctl}</strong>CTL</div>
              <div class="ws-chip ws-atl"><strong>${week.atl}</strong>ATL</div>
              <div class="ws-chip ws-tsb"><strong>${week.tsb > 0 ? '+' + week.tsb : week.tsb}</strong>TSB</div>
            </div>
            <div class="ws-row"><span>Total Duration</span><strong>${week.durationLabel}</strong></div>
            <div class="ws-row"><span>Total TSS</span><strong>${week.tss}</strong></div>
          `;
          row.appendChild(weekCard);
          month.appendChild(row);
        }

        wrap.appendChild(month);
      }

      if (!initialMonthCentered) {
        jumpToCurrentMonth();
        initialMonthCentered = true;
      }
    }

    function jumpToCurrentMonth() {
      const wrap = document.getElementById('calendarScroll');
      const current = wrap.querySelector('.current-month');
      if (current) wrap.scrollTop = current.offsetTop - 7;
    }

    function renderDashboard() {
      const totalDistance = activities.reduce((sum, a) => sum + Number(a.distance || 0), 0);
      const totalTime = activities.reduce((sum, a) => sum + Number(a.moving_time || 0), 0);
      document.getElementById('statCount').textContent = String(activities.length);
      document.getElementById('statPlanned').textContent = String(calendarItems.filter(i => i.kind === 'workout').length);
      document.getElementById('statDistance').textContent = fmtDistanceMeters(totalDistance);
      document.getElementById('statTime').textContent = fmtHours(totalTime);
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
        const [aResp, cResp, pResp] = await Promise.all([fetch('/ui/activities'), fetch('/calendar-items'), fetch('/pairs')]);
        activities = aResp.ok ? await aResp.json() : [];
        calendarItems = cResp.ok ? await cResp.json() : [];
        pairs = pResp.ok ? await pResp.json() : [];
      } catch (_err) {
        activities = [];
        calendarItems = [];
        pairs = [];
      }

      if (resetMonthPosition) initialMonthCentered = false;
      renderHome();
      renderCalendar();
      renderDashboard();
    }

    document.querySelectorAll('.tab').forEach(btn => {
      btn.addEventListener('click', () => setView(btn.dataset.view));
    });

    document.getElementById('jumpToday').addEventListener('click', jumpToCurrentMonth);
    document.getElementById('uploadFitBtn').addEventListener('click', () => {
      document.getElementById('uploadFitInput').click();
    });
    document.getElementById('uploadFitInput').addEventListener('change', async (event) => {
      const input = event.target;
      const file = input.files && input.files[0];
      if (!file) return;
      const formData = new FormData();
      formData.append('file', file);
      const resp = await fetch('/import-fit', {
        method: 'POST',
        body: formData,
      });
      input.value = '';
      if (!resp.ok) {
        const err = await resp.text();
        alert(`Import failed: ${err}`);
        return;
      }
      await loadData(false);
    });
    document.getElementById('globalSettings').addEventListener('click', () => {
      openContextMenu(16, 64, [
        { label: 'Settings (coming next)', onClick: async () => {} },
      ]);
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

    document.getElementById('closeWorkoutView').addEventListener('click', closeWorkoutModal);
    document.getElementById('workoutViewModal').addEventListener('click', (event) => {
      if (event.target.id === 'workoutViewModal') closeWorkoutModal();
    });
    document.querySelectorAll('.wv-tab[data-wv-tab]').forEach(btn => {
      btn.addEventListener('click', () => switchWorkoutTab(btn.dataset.wvTab));
    });
    document.getElementById('contextMenu').addEventListener('click', (ev) => ev.stopPropagation());
    document.addEventListener('click', (ev) => {
      if (!ev.target.closest('.card-menu-btn') && !ev.target.closest('#contextMenu')) {
        closeContextMenu();
      }
    });

    buildTypeGrids();
    bindWidgetToggles();
    applyWidgetPrefs();
    setView('home');
    loadData(true);
  </script>
</body>
</html>
    """


@app.get("/health")
def health() -> dict[str, bool]:
    return {"ok": True}


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
        if override.get("hidden"):
            continue
        out.append(updated)
    return out


@app.delete("/activities/{activity_id}")
def delete_activity_local(activity_id: str) -> dict[str, bool]:
    imported = load_imported_activities()
    filtered = [a for a in imported if str(a.get("id")) != activity_id]
    if len(filtered) != len(imported):
        save_imported_activities(filtered)
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
async def import_fit(file: UploadFile = File(...)) -> dict[str, Any]:
    filename = file.filename or "workout.fit"
    ext = Path(filename).suffix.lower()
    if ext != ".fit":
        raise HTTPException(status_code=400, detail="Only .fit files are supported.")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Empty file.")

    file_id = str(uuid4())
    imports_dir = Path("data/imports")
    imports_dir.mkdir(parents=True, exist_ok=True)
    saved_path = imports_dir / f"{file_id}.fit"
    saved_path.write_bytes(content)

    name = Path(filename).stem.replace("_", " ").replace("-", " ").strip() or "Imported Workout"
    item = {
        "id": f"imported-{file_id}",
        "name": name.title(),
        "type": "Ride",
        "distance": 0,
        "moving_time": 0,
        "start_date_local": f"{date.today().isoformat()}T08:00:00",
        "description": f"Imported from {filename}",
        "source": "fit",
    }
    imported = load_imported_activities()
    imported.append(item)
    save_imported_activities(imported)
    return item


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
        "intensity": payload.get("planned_intensity", 6),
        "description": payload.get("description", ""),
    }
    item = normalize_item(wrapped)
    items = load_calendar_items()
    items.append(item)
    save_calendar_items(items)
    return item
