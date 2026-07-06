from __future__ import annotations

import asyncio
import os
import re
import shutil
import sqlite3
import urllib.request
from contextlib import asynccontextmanager
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any

from fastapi import BackgroundTasks, FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from playwright.async_api import async_playwright
try:
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps
    import pytesseract
except ImportError:
    Image = ImageEnhance = ImageFilter = ImageOps = pytesseract = None

ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "data" / "interruptions.db"
FB_PROFILE = ROOT / "data" / "facebook-profile"
FB_MARKER = ROOT / "data" / ".facebook_connected"
REFRESH_MINUTES = int(os.getenv("REFRESH_MINUTES", "60"))
POST_LIMIT = int(os.getenv("POST_LIMIT", "20"))

PROVIDERS = [
    ("BENECO", "Baguio / Benguet", "https://www.facebook.com/benguetelectric"),
    ("CEBECO I", "Cebu", "https://www.facebook.com/cebu1EC"),
    ("CEBECO II", "Cebu", "https://www.facebook.com/cebeco2.official"),
    ("CEBECO III", "Cebu", "https://www.facebook.com/CEBECOIIIToledo"),
    ("CEPALCO", "Cagayan de Oro", "https://www.facebook.com/cepalcoofficial"),
    ("Davao Light and Power Co.", "Davao", "https://www.facebook.com/DavaoLightOfficial"),
    ("INEC", "Ilocos Norte", "https://www.facebook.com/INECofficial"),
    ("MECO", "Mactan", "https://www.facebook.com/mecomactan"),
    ("MERALCO", "Metro Manila / service area", "https://www.facebook.com/meralco"),
    ("Negros Power", "Negros Occidental", "https://www.facebook.com/negrospowerph"),
    ("PELCO I", "Pampanga", "https://www.facebook.com/pelco1officialpage"),
    ("PELCO II", "Pampanga", "https://www.facebook.com/Pelco2"),
    ("PELCO III", "Pampanga", "https://www.facebook.com/pelco3official"),
    ("SOCOTECO I", "South Cotabato", "https://www.facebook.com/socoteco1.koronadal"),
    ("SOCOTECO II", "General Santos / Sarangani", "https://www.facebook.com/socoteco2.EC"),
    ("VECO", "Cebu", "https://www.facebook.com/visayanelectriccompany"),
]

KEYWORDS = re.compile(r"power interruption|service interruption|scheduled interruption|brownout|maintenance schedule|power advisory|cancelled|canceled|rescheduled", re.I)
CANCELLED = re.compile(r"\b(cancelled|canceled|cancellation|will no longer push through|called off)\b", re.I)
RESCHEDULED = re.compile(r"\b(rescheduled|moved to|new schedule)\b", re.I)
DATE_PATTERNS = [
    re.compile(r"(?:date\s*[:\-]?\s*)?((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2},?\s+20\d{2})", re.I),
    re.compile(r"(?:date\s*[:\-]?\s*)?(20\d{2}[-/]\d{1,2}[-/]\d{1,2})", re.I),
]
TIME_RE = re.compile(r"(?:time\s*[:\-]?\s*)?((?:\d{1,2}:\d{2}|\d{1,2})\s*(?:AM|PM)?)\s*(?:to|\-|–|until)\s*((?:\d{1,2}:\d{2}|\d{1,2})\s*(?:AM|PM)?)", re.I)


def db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with db() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS interruptions (
          id INTEGER PRIMARY KEY, provider TEXT NOT NULL, site TEXT NOT NULL,
          event_date TEXT, start_time TEXT, end_time TEXT, location TEXT,
          reason TEXT, status TEXT NOT NULL DEFAULT 'Scheduled', raw_text TEXT,
          source_url TEXT NOT NULL, image_url TEXT, source_post_id TEXT,
          confidence REAL NOT NULL DEFAULT 0, scraped_at TEXT NOT NULL,
          UNIQUE(provider, source_url, event_date, start_time, location)
        );
        CREATE TABLE IF NOT EXISTS runs (
          id INTEGER PRIMARY KEY, provider TEXT NOT NULL, started_at TEXT NOT NULL,
          finished_at TEXT, status TEXT NOT NULL, posts_seen INTEGER DEFAULT 0,
          records_found INTEGER DEFAULT 0, message TEXT
        );
        CREATE INDEX IF NOT EXISTS ix_interruptions_date ON interruptions(event_date);
        """)


def clean_text(value: str) -> str:
    return re.sub(r"[ \t]+", " ", re.sub(r"\r", "", value or "")).strip()


def parse_date(text: str) -> str | None:
    for pattern in DATE_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        raw = match.group(1).replace("/", "-")
        for fmt in ("%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(raw, fmt).date().isoformat()
            except ValueError:
                pass
    return None


def field(text: str, label: str, following: tuple[str, ...]) -> str:
    stop = "|".join(re.escape(x) for x in following)
    match = re.search(rf"{label}\s*[:\-]\s*(.+?)(?=\s+(?:{stop})\s*[:\-]|$)", text, re.I | re.S)
    return clean_text(match.group(1)) if match else ""


def parse_record(provider: str, site: str, text: str, source_url: str, image_url: str | None) -> dict[str, Any] | None:
    text = clean_text(text)
    if not KEYWORDS.search(text) and not (parse_date(text) and TIME_RE.search(text)):
        return None
    status = "Cancelled" if CANCELLED.search(text) else "Rescheduled" if RESCHEDULED.search(text) else "Scheduled"
    time_match = TIME_RE.search(text)
    event_date = parse_date(text)
    location = field(text, "(?:location|affected areas?|areas? affected)", ("reason", "purpose", "date", "time", "status"))
    reason = field(text, "(?:reason|purpose)", ("date", "time", "location", "affected area", "status"))
    confidence = sum([bool(event_date), bool(time_match), bool(location)]) / 3
    return {
        "provider": provider, "site": site, "event_date": event_date,
        "start_time": clean_text(time_match.group(1)) if time_match else None,
        "end_time": clean_text(time_match.group(2)) if time_match else None,
        "location": location or "Needs review", "reason": reason,
        "status": status, "raw_text": text, "source_url": source_url,
        "image_url": image_url, "confidence": confidence,
        "scraped_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }


def tesseract_path() -> str | None:
    if pytesseract is None:
        return None
    configured = os.getenv("TESSERACT_CMD")
    candidates = [configured, shutil.which("tesseract"), r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]
    return next((str(path) for path in candidates if path and Path(path).exists()), None)


def _image_ocr_sync(image_url: str) -> str:
    executable = tesseract_path()
    if not executable:
        return ""
    try:
        with urllib.request.urlopen(image_url, timeout=25) as response:
            raw = response.read()
        image = Image.open(BytesIO(raw)).convert("RGB")
        if image.width < 1800:
            ratio = 1800 / image.width
            image = image.resize((1800, int(image.height * ratio)), Image.Resampling.LANCZOS)
        image = ImageOps.grayscale(image)
        image = ImageOps.autocontrast(image)
        image = ImageEnhance.Contrast(image).enhance(1.6)
        image = image.filter(ImageFilter.SHARPEN)
        pytesseract.pytesseract.tesseract_cmd = executable
        return clean_text(pytesseract.image_to_string(image, lang="eng", config="--oem 3 --psm 6"))
    except Exception:
        return ""


async def image_ocr(image_url: str) -> str:
    return await asyncio.to_thread(_image_ocr_sync, image_url)


def facebook_profile() -> str | None:
    configured = os.getenv("FACEBOOK_PROFILE_DIR")
    if configured:
        return configured
    return str(FB_PROFILE) if FB_MARKER.exists() else None


connect_lock = asyncio.Lock()


async def connect_facebook() -> None:
    async with connect_lock:
        FB_PROFILE.mkdir(parents=True, exist_ok=True)
        async with async_playwright() as pw:
            context = await pw.chromium.launch_persistent_context(
                str(FB_PROFILE), headless=False, locale="en-PH", timezone_id="Asia/Manila"
            )
            page = context.pages[0] if context.pages else await context.new_page()
            await page.goto("https://www.facebook.com/", wait_until="domcontentloaded", timeout=60000)
            for _ in range(300):
                cookies = await context.cookies("https://www.facebook.com")
                if any(cookie.get("name") == "c_user" for cookie in cookies):
                    FB_MARKER.write_text(datetime.now().astimezone().isoformat(), encoding="utf-8")
                    break
                await page.wait_for_timeout(2000)
            await context.close()


async def scrape_provider(name: str, site: str, url: str) -> tuple[int, list[dict[str, Any]]]:
    profile = facebook_profile()
    async with async_playwright() as pw:
        if profile:
            context = await pw.chromium.launch_persistent_context(
                profile, headless=True, locale="en-PH", timezone_id="Asia/Manila",
                args=["--disable-blink-features=AutomationControlled"],
            )
            page = context.pages[0] if context.pages else await context.new_page()
        else:
            browser = await pw.chromium.launch(headless=True)
            context = await browser.new_context(locale="en-PH", timezone_id="Asia/Manila")
            page = await context.new_page()
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(3500)
        await page.evaluate("window.scrollBy(0, 900)")
        await page.wait_for_timeout(1500)
        posts = await page.locator("[role='article'], div[data-pagelet^='FeedUnit_']").all()
        if not posts:
            mobile_url = url.replace("www.facebook.com", "m.facebook.com")
            await page.goto(mobile_url, wait_until="domcontentloaded", timeout=60000)
            await page.wait_for_timeout(3000)
            posts = await page.locator("[role='article'], article").all()
        output: list[dict[str, Any]] = []
        for article in posts[:POST_LIMIT]:
            text = clean_text(await article.inner_text())
            links = await article.locator("a[href*='/posts/'], a[href*='story_fbid'], a[href*='/photos/']").all()
            source_url = await links[0].get_attribute("href") if links else url
            if source_url and source_url.startswith("/"):
                source_url = "https://www.facebook.com" + source_url
            images = await article.locator("img").all()
            image_url = None
            for image in images:
                candidate = await image.get_attribute("src")
                if candidate and ("scontent" in candidate or "fbcdn" in candidate):
                    image_url = candidate
                    break
            combined = text
            if image_url and (KEYWORDS.search(text) or len(text) < 300):
                combined += "\n" + await image_ocr(image_url)
            record = parse_record(name, site, combined, source_url or url, image_url)
            if record:
                output.append(record)
        await context.close()
        return len(posts), output


def save_records(records: list[dict[str, Any]]) -> int:
    count = 0
    with db() as conn:
        for item in records:
            if item["status"] == "Cancelled" and item["event_date"]:
                # Cancellation advisories are often separate Facebook posts. Preserve
                # the advisory and also mark the earlier schedule for that provider/date.
                conn.execute("""UPDATE interruptions SET status='Cancelled', scraped_at=?
                    WHERE provider=? AND event_date=? AND status!='Cancelled'
                    AND (?='Needs review' OR location LIKE '%' || ? || '%' OR ? LIKE '%' || location || '%')""",
                    (item["scraped_at"], item["provider"], item["event_date"], item["location"], item["location"], item["location"]))
            before = conn.total_changes
            conn.execute("""INSERT INTO interruptions
              (provider,site,event_date,start_time,end_time,location,reason,status,raw_text,source_url,image_url,confidence,scraped_at)
              VALUES (:provider,:site,:event_date,:start_time,:end_time,:location,:reason,:status,:raw_text,:source_url,:image_url,:confidence,:scraped_at)
              ON CONFLICT(provider,source_url,event_date,start_time,location) DO UPDATE SET
                status=excluded.status, reason=excluded.reason, raw_text=excluded.raw_text,
                image_url=excluded.image_url, confidence=excluded.confidence, scraped_at=excluded.scraped_at""", item)
            count += conn.total_changes - before
    return count


refresh_lock = asyncio.Lock()


async def refresh_all() -> None:
    if refresh_lock.locked() or connect_lock.locked():
        return
    async with refresh_lock:
        for name, site, url in PROVIDERS:
            started = datetime.now().astimezone().isoformat(timespec="seconds")
            with db() as conn:
                run_id = conn.execute("INSERT INTO runs(provider,started_at,status) VALUES (?,?,?)", (name, started, "Running")).lastrowid
            try:
                seen, records = await scrape_provider(name, site, url)
                saved = save_records(records)
                status, message = ("Complete", f"{saved} records updated") if seen else ("Blocked", "No public posts visible; configure FACEBOOK_PROFILE_DIR")
            except Exception as exc:
                seen, records, status, message = 0, [], "Failed", str(exc)[:500]
            with db() as conn:
                conn.execute("UPDATE runs SET finished_at=?,status=?,posts_seen=?,records_found=?,message=? WHERE id=?", (datetime.now().astimezone().isoformat(timespec="seconds"), status, seen, len(records), message, run_id))


async def scheduler() -> None:
    await asyncio.sleep(3)
    while True:
        await refresh_all()
        await asyncio.sleep(max(5, REFRESH_MINUTES) * 60)


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    task = None if os.getenv("DISABLE_SCHEDULER") == "1" else asyncio.create_task(scheduler())
    yield
    if task:
        task.cancel()


app = FastAPI(title="PowerWatch PH", lifespan=lifespan)


@app.get("/api/interruptions")
def interruptions(status: str | None = None, provider: str | None = None, start: str | None = None, end: str | None = None):
    clauses, args = [], []
    for column, value in (("status", status), ("provider", provider)):
        if value:
            clauses.append(f"{column}=?"); args.append(value)
    if start: clauses.append("event_date>=?"); args.append(start)
    if end: clauses.append("event_date<=?"); args.append(end)
    where = " WHERE " + " AND ".join(clauses) if clauses else ""
    with db() as conn:
        rows = conn.execute("SELECT * FROM interruptions" + where + " ORDER BY event_date,start_time,provider", args).fetchall()
    return [dict(row) for row in rows]


@app.get("/api/status")
def scrape_status():
    with db() as conn:
        runs = conn.execute("""SELECT r.* FROM runs r JOIN (SELECT provider,MAX(id) id FROM runs GROUP BY provider) x ON x.id=r.id ORDER BY provider""").fetchall()
        counts = conn.execute("SELECT status,COUNT(*) count FROM interruptions GROUP BY status").fetchall()
    return {"refreshing": refresh_lock.locked(), "connecting": connect_lock.locked(),
            "facebook_connected": bool(os.getenv("FACEBOOK_PROFILE_DIR") or FB_MARKER.exists()),
            "refresh_minutes": REFRESH_MINUTES,
            "ocr": {"available": bool(tesseract_path()), "engine": "Local Tesseract"},
            "runs": [dict(r) for r in runs], "counts": {r["status"]: r["count"] for r in counts}}


@app.post("/api/refresh")
async def trigger_refresh(background: BackgroundTasks):
    if refresh_lock.locked():
        return JSONResponse({"message": "Refresh already running"}, status_code=202)
    background.add_task(refresh_all)
    return JSONResponse({"message": "Refresh started"}, status_code=202)


@app.post("/api/facebook/connect")
async def start_facebook_connection(background: BackgroundTasks):
    if refresh_lock.locked():
        return JSONResponse({"message": "Wait for the current refresh to finish"}, status_code=409)
    if connect_lock.locked():
        return JSONResponse({"message": "Facebook sign-in is already open"}, status_code=202)
    background.add_task(connect_facebook)
    return JSONResponse({"message": "Facebook sign-in window opened. Log in there, then close it after the dashboard shows Connected."}, status_code=202)


@app.get("/api/export.xlsx")
def export_xlsx(start: str = Query(default_factory=lambda: date.today().isoformat()), days: int = Query(7, ge=1, le=366)):
    start_date = date.fromisoformat(start)
    end_date = start_date + timedelta(days=days - 1)
    with db() as conn:
        rows = [dict(r) for r in conn.execute("SELECT * FROM interruptions WHERE event_date BETWEEN ? AND ? ORDER BY provider,event_date", (start_date.isoformat(), end_date.isoformat()))]
    wb = Workbook(); summary = wb.active; summary.title = "Summary"
    headers = ["Power Provider", "Site"] + [(start_date + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(days)]
    summary.append(headers)
    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in rows: grouped.setdefault(row["provider"], []).append(row)
    for name, site, _ in PROVIDERS:
        cells = [name, site]
        for i in range(days):
            target = (start_date + timedelta(days=i)).isoformat()
            items = [r for r in grouped.get(name, []) if r["event_date"] == target]
            cells.append("\n\n".join(f"Status: {r['status']}\nLocation: {r['location']}\nTime: {r['start_time'] or ''} to {r['end_time'] or ''}\nReason: {r['reason'] or ''}" for r in items))
        summary.append(cells)
    fill = PatternFill("solid", fgColor="17324D")
    for cell in summary[1]: cell.fill = fill; cell.font = Font(color="FFFFFF", bold=True); cell.alignment = Alignment(horizontal="center")
    summary.freeze_panes = "C2"; summary.column_dimensions["A"].width = 28; summary.column_dimensions["B"].width = 24
    for col in range(3, 3 + days): summary.column_dimensions[summary.cell(1, col).column_letter].width = 42
    for row in summary.iter_rows(min_row=2):
        for cell in row: cell.alignment = Alignment(vertical="top", wrap_text=True)
    for name, _, url in PROVIDERS:
        sheet = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "", name)[:31]); sheet.append(["Power Provider", "Date", "Time", "Location", "Reason", "Status", "Source"])
        for r in grouped.get(name, []): sheet.append([name, r["event_date"], f"{r['start_time'] or ''} to {r['end_time'] or ''}", r["location"], r["reason"], r["status"], r["source_url"]])
        for cell in sheet[1]: cell.fill = fill; cell.font = Font(color="FFFFFF", bold=True)
        sheet.freeze_panes = "A2"; sheet.auto_filter.ref = sheet.dimensions
        for width, letter in zip((24, 13, 22, 65, 45, 14, 45), "ABCDEFG"): sheet.column_dimensions[letter].width = width
        for row in sheet.iter_rows():
            for cell in row: cell.alignment = Alignment(vertical="top", wrap_text=True)
    output = BytesIO(); wb.save(output); output.seek(0)
    filename = f"Scheduled_Power_Interruption_{start_date}_{end_date}.xlsx"
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="{filename}"'})


app.mount("/", StaticFiles(directory=ROOT / "static", html=True), name="static")
