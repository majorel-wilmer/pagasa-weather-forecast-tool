from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
import re

import requests
from bs4 import BeautifulSoup
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
SNAPSHOT = ROOT / "data" / "pagasa_selected_cities.html"
PAGASA_URL = "https://www.pagasa.dost.gov.ph/weather/weather-outlook-selected-philippine-cities"

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
    rows = []
    for region, site, source_city in SITES:
        city_data = parsed["cities"].get(source_city.casefold())
        rows.append({
            "region": region,
            "site": site,
            "source_city": source_city,
            "days": city_data["days"] if city_data else [],
            "available": bool(city_data),
        })
    return {
        "issued": parsed["issued"],
        "retrieved_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "source_mode": source_mode,
        "source_url": PAGASA_URL,
        "rows": rows,
    }


def forecast_sentence(day: dict) -> str:
    rain = f"{day['rain_chance']}% chance of rain" if day["rain_chance"] is not None else "Rain chance unavailable"
    return f"{day['condition']}. Low {day['low']}, high {day['high']}; {rain}."


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


@app.get("/api/export")
def export():
    payload = build_payload()
    stream = export_workbook(payload)
    filename = f"PAGASA_5-Day_Forecast_{datetime.now():%Y%m%d}.xlsx"
    return StreamingResponse(stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="{filename}"'})

