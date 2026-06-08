"""
PSX stock scraper -> Google Sheets writer.

Usage:
    python psx_to_sheets.py KEL
    python psx_to_sheets.py OGDC

Each run always updates "PSX Stock Data" (latest snapshot).
"PSX Price History" is only appended when the run falls within
±5 minutes of a scheduled Task Scheduler trigger time.
"""

import sys
import os
import re
import json
import logging
from datetime import datetime, time as dtime

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials

# ── Config ────────────────────────────────────────────────────────────────────

CREDENTIALS_FILE   = "credentials.json"
SHEET_NAME         = "PSX Stock Data"
HISTORY_SHEET_NAME = "PSX Price History"
# SHEET_ID: env var takes priority (GitHub Actions); hardcoded value used locally
SHEET_ID           = os.environ.get("SHEET_ID", "1HI8lEQD9K9XjSVwfOplUN0Nc-Y1XMlE39fkwRpr-voc")
PSX_BASE_URL       = "https://dps.psx.com.pk/company/"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS = [
    "Symbol", "Company Name", "Current Price", "Change", "Change %",
    "Open", "High", "Low", "Volume", "LDCP",
    "1Y Change", "YTD Change", "Last Updated",
]

HISTORY_HEADERS = [
    "Symbol", "Current Price", "Open", "High", "Low", "Volume",
    "Change %", "Last Updated",
]

# Scheduled Task Scheduler trigger times (24-hour HH:MM).
# History is only appended when the script runs within ±5 min of these times.
#
# Monday–Thursday triggers: 09:33, 10:30, 11:30, 12:30, 13:30, 14:30, 15:30, 16:33
# Friday triggers          : 09:15, 10:15, 11:15, 12:00, 14:30, 15:30, 16:30
#   (PSX opens at 09:15 on Fridays; lunch break 12:00–14:00; no 13:30 slot)
TRIGGER_TIMES_WEEKDAY = [
    dtime(9,  33),
    dtime(10, 30),
    dtime(11, 30),
    dtime(12, 30),
    dtime(13, 30),
    dtime(14, 30),
    dtime(15, 30),
    dtime(16, 33),
]
TRIGGER_TIMES_FRIDAY = [
    dtime(9,  15),
    dtime(10, 15),
    dtime(11, 15),
    dtime(12,  0),
    dtime(14, 30),
    dtime(15, 30),
    dtime(16, 30),
]
TRIGGER_WINDOW_MINUTES = 5

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ── Scraper ───────────────────────────────────────────────────────────────────

def scrape_psx(ticker: str) -> dict:
    """Fetch and parse stock data for *ticker* from dps.psx.com.pk."""
    url = PSX_BASE_URL + ticker.upper()
    log.info("Fetching %s", url)

    resp = requests.get(url, headers=REQUEST_HEADERS, timeout=20)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    # ── Company name ──────────────────────────────────────────────────────────
    name_tag = soup.find(class_="quote__name")
    company_name = name_tag.get_text(strip=True) if name_tag else "N/A"

    # ── Current price ─────────────────────────────────────────────────────────
    price_tag = soup.find(class_="quote__close")
    raw_price = price_tag.get_text(strip=True) if price_tag else ""
    price_match = re.search(r"[\d,]+\.?\d*", raw_price)
    current_price = price_match.group() if price_match else "N/A"

    # ── Change & Change % ─────────────────────────────────────────────────────
    change_tag = soup.find(class_="quote__change")
    raw_change = change_tag.get_text(strip=True) if change_tag else ""
    # e.g. "-0.09(-1.11%)" or "+0.50(+0.62%)"
    change_match = re.match(r"([+-]?[\d.]+)\(([+-]?[\d.]+)%\)", raw_change)
    if change_match:
        change = change_match.group(1)
        change_pct = change_match.group(2)
    else:
        change = raw_change
        change_pct = "N/A"

    # ── Stats table (Open / High / Low / Volume / LDCP / 1Y / YTD) ───────────
    # The page has multiple stats sections (main + futures).
    # We use the FIRST occurrence of each label, which belongs to the main stock.
    stats: dict[str, str] = {}
    quote_stats = soup.find(class_="quote__stats")
    if quote_stats:
        for item in quote_stats.find_all(class_="stats_item"):
            label_tag = item.find(class_="stats_label")
            value_tag = item.find(class_="stats_value")
            if label_tag and value_tag:
                label = label_tag.get_text(strip=True)
                value = value_tag.get_text(strip=True)
                if label not in stats:
                    stats[label] = value

    def get_stat(*keys) -> str:
        for k in keys:
            if k in stats:
                return stats[k]
        return "N/A"

    def strip_pct(value: str) -> str:
        return value.rstrip("%") if value != "N/A" else "N/A"

    def strip_commas(value: str) -> str:
        return value.replace(",", "") if value != "N/A" else "N/A"

    return {
        "Symbol":        ticker.upper(),
        "Company Name":  company_name,
        "Current Price": current_price or "N/A",
        "Change":        change or "N/A",
        "Change %":      change_pct,
        "Open":          get_stat("Open"),
        "High":          get_stat("High"),
        "Low":           get_stat("Low"),
        "Volume":        strip_commas(get_stat("Volume")),
        "LDCP":          get_stat("LDCP"),
        "1Y Change":     strip_pct(get_stat("1-Year Change * ^", "1-Year Change")),
        "YTD Change":    strip_pct(get_stat("YTD Change * ^", "YTD Change")),
        "Last Updated":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


# ── Google Sheets helpers ─────────────────────────────────────────────────────

def open_spreadsheet() -> gspread.Spreadsheet:
    # GitHub Actions: credentials passed as env var JSON string
    gc_env = os.environ.get("GOOGLE_CREDENTIALS")
    if gc_env:
        info = json.loads(gc_env.lstrip(u'﻿'))  # strip UTF-8 BOM if present
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        log.info("Using GOOGLE_CREDENTIALS env var (GitHub Actions)")
    else:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        log.info("Using %s (local)", CREDENTIALS_FILE)

    client = gspread.authorize(creds)
    if SHEET_ID:
        return client.open_by_key(SHEET_ID)
    try:
        return client.open(SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        raise RuntimeError(
            f"Sheet '{SHEET_NAME}' not found. Create it in Google Drive, "
            "share it with the service account (Editor), and set SHEET_ID."
        )


def get_or_create_worksheet(spreadsheet: gspread.Spreadsheet,
                             tab_name: str,
                             headers: list) -> gspread.Worksheet:
    try:
        ws = spreadsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        log.info("Creating worksheet '%s'", tab_name)
        ws = spreadsheet.add_worksheet(title=tab_name, rows=10000, cols=len(headers))
    # Write headers if row 1 is empty or mismatched
    if ws.row_values(1) != headers:
        log.info("Writing headers to '%s'", tab_name)
        ws.update([headers], "A1", value_input_option="USER_ENTERED")
    return ws


# ── Mode 1: latest snapshot (upsert by symbol) ───────────────────────────────

def update_latest(data: dict, spreadsheet: gspread.Spreadsheet) -> None:
    ws = get_or_create_worksheet(spreadsheet, "PSX Stock Data", HEADERS)

    all_symbols = ws.col_values(1)
    ticker = data["Symbol"]
    try:
        row_idx = all_symbols.index(ticker) + 1
    except ValueError:
        row_idx = max(len(all_symbols) + 1, 2)

    row_values = [data[h] for h in HEADERS]
    ws.update([row_values], f"A{row_idx}", value_input_option="USER_ENTERED")
    log.info("[latest]  '%s' → row %d of '%s'", ticker, row_idx, "PSX Stock Data")


# ── Mode 2: history log (append only) ────────────────────────────────────────

def append_history(data: dict, spreadsheet: gspread.Spreadsheet) -> None:
    ws = get_or_create_worksheet(spreadsheet, HISTORY_SHEET_NAME, HISTORY_HEADERS)

    row_values = [data[h] for h in HISTORY_HEADERS]
    ws.append_row(row_values, value_input_option="USER_ENTERED")
    log.info("[history] '%s' appended to '%s'", data["Symbol"], HISTORY_SHEET_NAME)


# ── Trigger-time check ────────────────────────────────────────────────────────

def is_scheduled_run() -> bool:
    """Return True if this is a scheduled run (GitHub Actions or within trigger window)."""
    # GitHub Actions always runs on schedule — always append history
    if os.environ.get("GITHUB_ACTIONS") == "true":
        return True
    # Local: check if current time falls within a trigger window
    now = datetime.now()
    triggers = TRIGGER_TIMES_FRIDAY if now.weekday() == 4 else TRIGGER_TIMES_WEEKDAY
    now_minutes = now.hour * 60 + now.minute
    for t in triggers:
        if abs(now_minutes - (t.hour * 60 + t.minute)) <= TRIGGER_WINDOW_MINUTES:
            return True
    return False


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python psx_to_sheets.py <TICKER>")
        print("Example: python psx_to_sheets.py KEL")
        sys.exit(1)

    ticker = sys.argv[1].strip().upper()

    try:
        data = scrape_psx(ticker)
    except requests.HTTPError as exc:
        log.error("HTTP error fetching '%s': %s", ticker, exc)
        sys.exit(1)
    except Exception as exc:
        log.error("Scraping failed for '%s': %s", ticker, exc)
        sys.exit(1)

    log.info("Data: %s", data)

    scheduled = is_scheduled_run()
    if not scheduled:
        log.info("Manual run — 'PSX Price History' will NOT be updated.")

    try:
        spreadsheet = open_spreadsheet()
        update_latest(data, spreadsheet)
        if scheduled:
            append_history(data, spreadsheet)
    except FileNotFoundError:
        log.error("'%s' not found — place your service-account JSON in the same folder.", CREDENTIALS_FILE)
        sys.exit(1)
    except Exception as exc:
        log.error("Google Sheets error: %s", exc)
        sys.exit(1)

    if scheduled:
        print(f"[OK] {ticker} -> updated 'PSX Stock Data' + appended to 'PSX Price History'")
    else:
        print(f"[OK] {ticker} -> updated 'PSX Stock Data' (manual run, history skipped)")


if __name__ == "__main__":
    main()
