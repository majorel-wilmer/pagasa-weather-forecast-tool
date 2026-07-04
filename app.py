from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
import hashlib
import hmac
import json
import os
import re
import secrets
import time

import requests
from bs4 import BeautifulSoup
from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
SNAPSHOT = ROOT / "data" / "pagasa_selected_cities.html"
PAGASA_URL = "https://www.pagasa.dost.gov.ph/weather/weather-outlook-selected-philippine-cities"
PAGASA_WEEKLY_URL = "https://www.pagasa.dost.gov.ph/weather/weather-outlook-weekly"
LOCAL_OVERRIDES = ROOT / "data" / "overrides.json"
OVERRIDE_PREFIX = "pagasa-weather-overrides/"
FORECAST_WINDOW = "8:00 AM – 8:00 AM next day"
_override_cache = {"loaded_at": 0.0, "data": {}}

SITES = [
    ("LUZON", "Alabang", "Metro Manila"),
    ("", "Antipolo", "Metro Manila"),
    ("", "Baguio", "Baguio City"),
    ("", "Clark", "Sbma (Olongapo)"),
    ("", "Laoag", "Laoag City"),
    ("", "Metro Manila", "Metro Manila"),
    ("", "Molino", "Tagaytay City"),
    ("VISAYAS", "Bacolod", "Bacolod City"),
    ("", "Cebu", "Metro Cebu"),
    ("MINDANAO", "CDO", "Cagayan De Oro City"),
    ("", "Davao", "Metro Davao"),
    ("", "GenSan", "Metro Davao"),
]

app = FastAPI(title="PAGASA 5-Day Weather Tool")
app.mount("/static", StaticFiles(directory=STATIC), name="static")


class LoginPayload(BaseModel):
    password: str


class OverridePayload(BaseModel):
    site: str
    date: str
    red: bool


def automatic_severity(condition: str) -> str:
    text = condition.casefold()
    if any(word in text for word in ("torrential", "heavy", "intense")):
        return "orange"
    if "moderate" in text:
        return "yellow"
    if any(word in text for word in ("light rain", "rainfall", "rainshowers", "rain", "thunderstorm")):
        return "green"
    return "none"


def fetch_weekly_outlook() -> dict:
    try:
        response = requests.get(PAGASA_WEEKLY_URL, timeout=25, headers={"User-Agent": "Mozilla/5.0 PAGASA-Weather-Tool/1.0"})
        response.raise_for_status()
        text = " ".join(BeautifulSoup(response.text, "html.parser").get_text(" ", strip=True).replace("\xa0", " ").split())
        issued = re.search(r"Issued at:\s*(.+?)\s+Valid until:", text, re.I)
        valid = re.search(r"Valid until:\s*(.+?)\s+\d{1,2}(?:-\d{1,2})?\s+[A-Z]", text, re.I)
        return {
            "available": True,
            "issued": issued.group(1) if issued else "Issue time unavailable",
            "valid_until": valid.group(1) if valid else "",
            "summary": text,
            "source_url": PAGASA_WEEKLY_URL,
        }
    except requests.RequestException as exc:
        return {"available": False, "summary": "", "source_url": PAGASA_WEEKLY_URL, "error": str(exc)}


def weekly_context(site: str, date_text: str, condition: str, weekly: dict) -> dict:
    """Merge risks explicitly named in PAGASA's weekly narrative.

    This intentionally does not invent hourly probabilities. The weekly outlook
    sometimes supplies a qualitative time of day and hazards omitted by the
    selected-city table.
    """
    summary = weekly.get("summary", "").upper()
    timing = FORECAST_WINDOW
    alert = ""
    alert_level = "none"
    severity = automatic_severity(condition)
    try:
        date = datetime.strptime(date_text, "%A %B %d, %Y")
    except ValueError:
        date = None

    if "AFTERNOON OR EVENING" in summary and "THUNDERSTORM" in condition.upper():
        timing = "Afternoon to evening (PAGASA; exact hours unavailable)"

    if date and "BAVI" in summary and date.month == 7 and date.day in (8, 9):
        if site in {"Laoag", "Baguio"}:
            alert = "BAVI / INDAY: rains with gusty winds possible"
            alert_level = "cyclone"
        if site == "Bacolod" and "NEGROS ISLAND REGION" in summary and "AT TIMES HEAVY RAINS" in summary:
            severity = "orange"
            alert = "Enhanced Habagat: light to moderate, at times heavy rain"
            alert_level = "heavy-rain"
        if site == "GenSan" and "SOCCSKSARGEN" in summary and "AT TIMES HEAVY RAINS" in summary:
            severity = "orange"
            alert = "Enhanced Habagat: light to moderate, at times heavy rain"
            alert_level = "heavy-rain"

    if date and "BAVI" in summary and date.month == 7 and date.day == 10:
        if site in {"Laoag", "Baguio"}:
            alert = "BAVI / INDAY: rains with gusty winds possible"
            alert_level = "cyclone"
        elif site in {"Clark", "Molino", "Bacolod", "GenSan"}:
            alert = "Enhanced Habagat / monsoon rain risk"
            alert_level = "monsoon"

    return {"severity": severity, "forecast_window": timing, "weather_alert": alert, "alert_level": alert_level}


def override_key(site: str, date: str) -> str:
    return f"{site.strip().casefold()}|{date.strip().casefold()}"


def _local_overrides() -> dict:
    try:
        return json.loads(LOCAL_OVERRIDES.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def load_overrides(force: bool = False) -> dict:
    now = time.time()
    if not force and now - _override_cache["loaded_at"] < 10:
        return dict(_override_cache["data"])
    if not os.getenv("BLOB_READ_WRITE_TOKEN"):
        data = _local_overrides()
    else:
        try:
            from vercel.blob import list_objects

            result = list_objects(prefix=OVERRIDE_PREFIX, limit=100)
            latest = max(result.blobs, key=lambda item: item.uploaded_at, default=None)
            data = requests.get(latest.url, params={"v": int(now)}, timeout=10).json() if latest else {}
        except Exception:
            data = dict(_override_cache["data"])
    _override_cache.update({"loaded_at": now, "data": data})
    return dict(data)


def save_overrides(data: dict) -> None:
    if os.getenv("BLOB_READ_WRITE_TOKEN"):
        from vercel.blob import BlobClient

        filename = f"{OVERRIDE_PREFIX}{int(time.time() * 1000)}-{secrets.token_hex(4)}.json"
        BlobClient().put(
            filename,
            json.dumps(data, separators=(",", ":")).encode("utf-8"),
            access="public",
            content_type="application/json",
            cache_control_max_age=60,
        )
    else:
        LOCAL_OVERRIDES.write_text(json.dumps(data, indent=2), encoding="utf-8")
    _override_cache.update({"loaded_at": time.time(), "data": dict(data)})


def verify_password(password: str) -> bool:
    encoded = os.getenv("ADMIN_PASSWORD_HASH", "")
    try:
        iterations_text, salt_hex, expected_hex = encoded.split("$", 2)
        actual = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt_hex), int(iterations_text))
        return hmac.compare_digest(actual.hex(), expected_hex)
    except (TypeError, ValueError):
        return False


def session_token() -> str:
    expires = int(time.time()) + 8 * 60 * 60
    secret = os.getenv("ADMIN_SESSION_SECRET", "")
    signature = hmac.new(secret.encode(), str(expires).encode(), hashlib.sha256).hexdigest()
    return f"{expires}.{signature}"


def is_admin(request: Request) -> bool:
    token = request.cookies.get("pagasa_admin", "")
    secret = os.getenv("ADMIN_SESSION_SECRET", "")
    try:
        expires_text, signature = token.split(".", 1)
        expected = hmac.new(secret.encode(), expires_text.encode(), hashlib.sha256).hexdigest()
        return bool(secret) and int(expires_text) > int(time.time()) and hmac.compare_digest(signature, expected)
    except (TypeError, ValueError):
        return False


def fetch_html() -> tuple[str, str]:
    try:
        response = requests.get(PAGASA_URL, timeout=25, headers={"User-Agent": "Mozilla/5.0 PAGASA-Weather-Tool/1.0"})
        response.raise_for_status()
        return response.text, "live"
    except requests.RequestException:
        if SNAPSHOT.exists():
            return SNAPSHOT.read_text(encoding="utf-8"), "saved snapshot"
        raise HTTPException(503, "PAGASA is temporarily unavailable and no saved snapshot exists.")


def clean_text(node) -> str:
    return " ".join(node.get_text(" ", strip=True).replace("\xa0", " ").split())


def parse_pagasa(html: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    outlook = soup.select_one("#outlook-phil-cities")
    if not outlook:
        raise HTTPException(502, "PAGASA page format changed; forecast section was not found.")

    issue = soup.select_one(".validity")
    issued = clean_text(issue) if issue else "Issue time unavailable"
    cities = {}
    for panel in outlook.select(".panel.panel-default"):
        title = panel.select_one(".panel-title a")
        table = panel.select_one("table")
        if not title or not table:
            continue
        city = clean_text(title).replace("›", "").strip()
        headers = [clean_text(th) for th in table.select("thead.desktop-view-thead th")]
        desktop = table.select_one("tbody tr.desktop-view-tr")
        if not desktop:
            continue
        days = []
        for header, cell in zip(headers, desktop.select("td")):
            image = cell.select_one("img")
            condition = image.get("title", "Forecast unavailable") if image else "Forecast unavailable"
            low = clean_text(cell.select_one(".min")) if cell.select_one(".min") else "—"
            high = clean_text(cell.select_one(".max")) if cell.select_one(".max") else "—"
            rain_node = next((s for s in cell.select("span") if "Chance of rain" in clean_text(s)), None)
            rain_text = clean_text(rain_node) if rain_node else "Chance of rain: —"
            rain_match = re.search(r"(\d+)%", rain_text)
            days.append({
                "date": header,
                "condition": condition,
                "low": low,
                "high": high,
                "rain_chance": int(rain_match.group(1)) if rain_match else None,
                "icon": image.get("src") if image else None,
            })
        cities[city.casefold()] = {"name": city, "days": days}
    return {"issued": issued, "cities": cities}


def build_payload() -> dict:
    html, source_mode = fetch_html()
    parsed = parse_pagasa(html)
    weekly = fetch_weekly_outlook()
    rows = []
    overrides = load_overrides()
    for region, site, source_city in SITES:
        city_data = parsed["cities"].get(source_city.casefold())
        # Copy shared city forecasts so site-specific red overrides cannot overwrite one another.
        days = [dict(day) for day in city_data["days"]] if city_data else []
        for day in days:
            context = weekly_context(site, day["date"], day["condition"], weekly)
            auto = context["severity"]
            red = bool(overrides.get(override_key(site, day["date"])))
            day.update({
                "forecast_window": context["forecast_window"],
                "automatic_severity": auto,
                "severity": "red" if red else auto,
                "red_override": red,
                "weather_alert": context["weather_alert"],
                "alert_level": context["alert_level"],
            })
        rows.append({
            "region": region,
            "site": site,
            "source_city": source_city,
            "days": days,
            "available": bool(city_data),
        })
    return {
        "issued": parsed["issued"],
        "retrieved_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "source_mode": source_mode,
        "source_url": PAGASA_URL,
        "weekly_outlook": weekly,
        "rows": rows,
    }


def forecast_sentence(day: dict) -> str:
    rain = f"{day['rain_chance']}% chance of rain" if day["rain_chance"] is not None else "Rain chance unavailable"
    return f"{day['condition']}. Low {day['low']}, high {day['high']}; {rain}. Forecast window: {day['forecast_window']}."


def export_workbook(payload: dict) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "5-Day Forecast"
    ws.sheet_view.showGridLines = False
    dates = []
    for row in payload["rows"]:
        if row["days"]:
            dates = [day["date"] for day in row["days"][:5]]
            break
    while len(dates) < 5:
        dates.append(f"Day {len(dates) + 1}")

    ws.append([None, "Site", *dates])
    for row in payload["rows"]:
        values = [row["region"], row["site"]]
        values += [forecast_sentence(d) for d in row["days"][:5]]
        values += ["Forecast unavailable"] * (7 - len(values))
        ws.append(values)

    severity_fills = {
        "green": PatternFill("solid", fgColor="C6EFCE"),
        "yellow": PatternFill("solid", fgColor="FFF2CC"),
        "orange": PatternFill("solid", fgColor="F4B183"),
        "red": PatternFill("solid", fgColor="FF6B6B"),
        "none": PatternFill("solid", fgColor="F2F2F2"),
    }
    for row_index, row in enumerate(payload["rows"], start=2):
        for day_index, day in enumerate(row["days"][:5], start=3):
            ws.cell(row_index, day_index).fill = severity_fills[day["severity"]]

    navy, blue, white = "17365D", "D9EAF7", "FFFFFF"
    ws["B1"].fill = PatternFill("solid", fgColor=navy)
    for cell in ws[1][1:7]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(color=white, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="A6A6A6")
    for row in ws.iter_rows(min_row=2, max_row=13, min_col=1, max_col=7):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        row[0].fill = PatternFill("solid", fgColor=navy)
        row[0].font = Font(color=white, bold=True)
        row[1].fill = PatternFill("solid", fgColor=blue)
        row[1].font = Font(bold=True)
    ws.merge_cells("A2:A8")
    ws.merge_cells("A9:A10")
    ws.merge_cells("A11:A13")
    for cell in (ws["A2"], ws["A9"], ws["A11"]):
        cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
    widths = [13, 18, 38, 38, 38, 38, 38]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 34
    for i in range(2, 14):
        ws.row_dimensions[i].height = 84
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "B1:G13"

    meta = wb.create_sheet("Source")
    meta.append(["Field", "Value"])
    meta.append(["PAGASA URL", payload["source_url"]])
    meta.append(["Issued", payload["issued"]])
    meta.append(["Retrieved", payload["retrieved_at"]])
    meta.append(["Source mode", payload["source_mode"]])
    meta.column_dimensions["A"].width = 20
    meta.column_dimensions["B"].width = 95
    meta["A1"].font = meta["B1"].font = Font(bold=True, color=white)
    meta["A1"].fill = meta["B1"].fill = PatternFill("solid", fgColor=navy)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.get("/")
def home():
    return FileResponse(STATIC / "index.html")


@app.get("/api/forecast")
def forecast():
    return JSONResponse(build_payload())


@app.get("/api/admin/status")
def admin_status(request: Request):
    return {"authenticated": is_admin(request), "configured": bool(os.getenv("ADMIN_PASSWORD_HASH"))}


@app.post("/api/admin/login")
def admin_login(payload: LoginPayload, response: Response):
    if not os.getenv("ADMIN_PASSWORD_HASH"):
        raise HTTPException(503, "Admin access is not configured.")
    if not verify_password(payload.password):
        raise HTTPException(401, "Incorrect admin password.")
    response.set_cookie(
        "pagasa_admin",
        session_token(),
        max_age=8 * 60 * 60,
        httponly=True,
        secure=bool(os.getenv("VERCEL")),
        samesite="strict",
    )
    return {"authenticated": True}


@app.post("/api/admin/logout")
def admin_logout(response: Response):
    response.delete_cookie("pagasa_admin")
    return {"authenticated": False}


@app.put("/api/admin/override")
def update_override(payload: OverridePayload, request: Request):
    if not is_admin(request):
        raise HTTPException(401, "Admin login required.")
    valid_sites = {site.casefold() for _, site, _ in SITES}
    if payload.site.casefold() not in valid_sites or not payload.date.strip():
        raise HTTPException(400, "Invalid site or forecast date.")
    data = load_overrides(force=True)
    key = override_key(payload.site, payload.date)
    if payload.red:
        data[key] = {"red": True, "updated_at": datetime.now().astimezone().isoformat(timespec="seconds")}
    else:
        data.pop(key, None)
    save_overrides(data)
    return {"site": payload.site, "date": payload.date, "red_override": payload.red}


@app.get("/api/export")
def export():
    payload = build_payload()
    stream = export_workbook(payload)
    filename = f"PAGASA_5-Day_Forecast_{datetime.now():%Y%m%d}.xlsx"
    return StreamingResponse(stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="{filename}"'})
