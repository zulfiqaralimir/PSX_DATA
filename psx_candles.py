"""
psx_candles.py  --  PSX intraday candlestick chart

Reads 'PSX Intraday OHLCV' from Google Sheets and produces an
interactive HTML candlestick chart with volume, gap highlights,
and key price levels.

Usage:
    python psx_candles.py KEL                 # today's date
    python psx_candles.py KEL 2026-06-09      # specific date

Output:
    charts/<TICKER>_<DATE>.html   (auto-opens in browser)
"""

# ── Auto-install missing packages ─────────────────────────────────────────────
import subprocess, sys

def _ensure(pkg, import_as=None):
    try:
        __import__(import_as or pkg)
    except ImportError:
        print(f"Installing {pkg} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure("plotly")
_ensure("gspread")
_ensure("google-auth",    "google.oauth2")
_ensure("pandas")
_ensure("requests")
_ensure("beautifulsoup4", "bs4")

# ── Imports ───────────────────────────────────────────────────────────────────
import os
import json
import logging
import tempfile
import time
import webbrowser
from datetime import date as dt_date, datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import gspread
import requests
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials

# ── Config ────────────────────────────────────────────────────────────────────

CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1HI8lEQD9K9XjSVwfOplUN0Nc-Y1XMlE39fkwRpr-voc"
SHEET_TAB        = "PSX Intraday OHLCV"
CHARTS_DIR       = Path("charts")
GAP_THRESHOLD    = 0.003      # minimum fractional gap size to highlight (0.3%)
PSX_BASE_URL     = "https://dps.psx.com.pk/company/"

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Sheet data is cached to a temp file so all 7 ticker invocations share one API read
CACHE_FILE        = Path(tempfile.gettempdir()) / "psx_sheet_cache.json"
CACHE_TTL_SECONDS = 600   # 10 minutes

ANNOUNCEMENTS = {
    "OBOY": {
        "2026-01-14": {
            "title": "Disclosure of Interest",
            "category": "Others",
            "summary": {
                "Disclosed":        "January 14, 2026",
                "Type":             "Insider Sell",
                "Executive":        "Mr. Inam Ullah (Company Secretary)",
                "Transaction Date": "January 13, 2026",
                "Shares Sold":      "4,500",
                "Rate":             "Rs. 13.31 per share",
                "Market":           "CDC Ready Market",
                "Regulation":       "Clause 5.6.4 of PSX Regulations",
                "Note":             "Transaction to be presented in "
                                    "next board meeting",
            },
        },
        "2025-12-29": {
            "title": "Material Information",
            "category": "Others",
            "summary": {
                "Disclosed":        "December 29, 2025",
                "Type":             "Insider Sell",
                "Executive":        "Mr. Khawaja Usman Arif",
                "Transaction Date": "December 18, 2025",
                "Shares Sold":      "207,000",
                "Rate":             "Rs. 11.91 per share",
                "Market":           "CDC Ready Market",
                "Regulation":       "Clause 5.6.4 of PSX Regulations",
                "Note":             "Transaction to be presented in "
                                    "next board meeting",
            },
        },
    },
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ── Sheet connection ──────────────────────────────────────────────────────────

def open_sheet() -> gspread.Worksheet:
    gc_env = os.environ.get("GOOGLE_CREDENTIALS")
    if gc_env:
        info = json.loads(gc_env.lstrip("﻿"))
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    ss = gspread.authorize(creds).open_by_key(SHEET_ID)
    return ss.worksheet(SHEET_TAB)


def _fetch_with_retry(ws: gspread.Worksheet, max_retries: int = 3) -> list:
    """Fetch all sheet records, retrying on 429 with exponential backoff."""
    for attempt in range(max_retries):
        try:
            return ws.get_all_records()
        except gspread.exceptions.APIError as exc:
            status = getattr(exc.response, "status_code", 0)
            if status == 429 and attempt < max_retries - 1:
                wait = 60 * (2 ** attempt)   # 60s, 120s
                log.warning("Rate limit (429) — waiting %ds (attempt %d/%d)", wait, attempt + 1, max_retries)
                time.sleep(wait)
            else:
                raise


# ── Data loading ──────────────────────────────────────────────────────────────

def load_all(ticker: str) -> pd.DataFrame:
    log.info("Loading data for %s from '%s'", ticker, SHEET_TAB)

    cache_valid = (
        CACHE_FILE.exists()
        and (time.time() - CACHE_FILE.stat().st_mtime) < CACHE_TTL_SECONDS
    )
    if cache_valid:
        log.info("Using cached sheet data (%s)", CACHE_FILE)
        raw = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    else:
        ws  = open_sheet()
        raw = _fetch_with_retry(ws)
        CACHE_FILE.write_text(json.dumps(raw), encoding="utf-8")
        log.info("Sheet data fetched and cached: %d rows", len(raw))

    df  = pd.DataFrame(raw)

    if df.empty:
        raise ValueError(f"Sheet '{SHEET_TAB}' is empty.")

    df = df[df["Symbol"].str.upper() == ticker.upper()].copy()
    if df.empty:
        raise ValueError(f"No data found for ticker '{ticker}'.")

    for col in ["Open", "High", "Low", "Close", "Volume"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Date"] = pd.to_datetime(df["Date"]).dt.date
    df["_sort_key"] = df["Interval"].str[:5]   # "09:33" from "09:33-10:30"
    df = df.sort_values(["Date", "_sort_key"]).drop(columns="_sort_key")
    return df


def get_day_df(all_df: pd.DataFrame, target_date: dt_date) -> pd.DataFrame:
    day = all_df[all_df["Date"] == target_date].reset_index(drop=True)
    return day


def get_ldcp(all_df: pd.DataFrame, target_date: dt_date) -> float | None:
    prev = all_df[all_df["Date"] < target_date]
    if prev.empty:
        return None
    last_date = prev["Date"].max()
    last_day  = prev[prev["Date"] == last_date]
    return float(last_day["Close"].iloc[-1])


# ── Gap detection ─────────────────────────────────────────────────────────────

def detect_gaps(df: pd.DataFrame) -> list:
    """
    True price gap: empty space between adjacent candles with zero overlap.
      Gap Up   — Low of next candle > High of previous candle (bullish)
      Gap Down — High of next candle < Low of previous candle (bearish)
    Overlapping candles are never flagged, regardless of Open/Close difference.
    """
    gaps = []
    for i in range(1, len(df)):
        prev_high = df["High"].iloc[i - 1]
        prev_low  = df["Low"].iloc[i - 1]
        curr_high = df["High"].iloc[i]
        curr_low  = df["Low"].iloc[i]

        if any(pd.isna(v) for v in [prev_high, prev_low, curr_high, curr_low]):
            continue
        if prev_high == 0 or prev_low == 0:
            continue

        if curr_low > prev_high:
            # Gap Up: shade between previous High and next Low
            size = curr_low - prev_high
            pct  = size / prev_high
            if pct < GAP_THRESHOLD:
                continue
            gaps.append({
                "interval_prev": df["Interval"].iloc[i - 1],
                "interval_curr": df["Interval"].iloc[i],
                "y0":  prev_high,
                "y1":  curr_low,
                "diff": round(size, 4),
                "pct":  round(pct * 100, 3),
                "up":   True,
                "idx":  i,
            })
        elif curr_high < prev_low:
            # Gap Down: shade between next High and previous Low
            size = prev_low - curr_high
            pct  = size / prev_low
            if pct < GAP_THRESHOLD:
                continue
            gaps.append({
                "interval_prev": df["Interval"].iloc[i - 1],
                "interval_curr": df["Interval"].iloc[i],
                "y0":  curr_high,
                "y1":  prev_low,
                "diff": round(-size, 4),
                "pct":  round(-pct * 100, 3),
                "up":   False,
                "idx":  i,
            })

    log.info("True price gaps detected: %d", len(gaps))
    return gaps


# ── News scraping ─────────────────────────────────────────────────────────────

def _parse_ann_date(date_str: str) -> datetime:
    try:
        return datetime.strptime(date_str, "%b %d, %Y")
    except Exception:
        return datetime.min


def scrape_news(ticker: str) -> list:
    """Fetch recent announcements for *ticker* from dps.psx.com.pk.

    Returns list of dicts: {date, title, url, category}
    """
    url = PSX_BASE_URL + ticker.upper()
    log.info("Fetching announcements for %s", ticker)
    try:
        resp = requests.get(url, headers=REQUEST_HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as exc:
        log.warning("News fetch failed for %s: %s", ticker, exc)
        return []

    panels = soup.find_all(class_="tabs__panel")
    # Tab list 1 maps to panel indices 4, 5, 6
    categories = [(4, "Financial Results"), (5, "Board Meetings"), (6, "Others")]

    items = []
    for panel_idx, category in categories:
        if panel_idx >= len(panels):
            continue
        tbl = panels[panel_idx].find(class_="tbl")
        if not tbl:
            continue
        for row in tbl.find_all("tr")[1:]:          # skip header row
            cells = row.find_all("td")
            if len(cells) < 2:
                continue
            date  = cells[0].get_text(strip=True)
            title = cells[1].get_text(strip=True)
            link  = ""
            if len(cells) >= 3:
                for a in cells[2].find_all("a", href=True):
                    href = a["href"]
                    if href and not href.startswith("javascript"):
                        link = ("https://dps.psx.com.pk" + href) if href.startswith("/") else href
                        break
            if title:
                items.append({"date": date, "title": title, "url": link, "category": category})

    items.sort(key=lambda x: _parse_ann_date(x["date"]), reverse=True)
    return items[:15]


def scrape_payouts(ticker: str) -> list:
    """Fetch payout history for *ticker* via POST to /company/payouts."""
    log.info("Fetching payouts for %s", ticker)
    try:
        resp = requests.post(
            "https://dps.psx.com.pk/company/payouts",
            data={"symbol": ticker.upper()},
            headers={
                **REQUEST_HEADERS,
                "X-Requested-With": "XMLHttpRequest",
                "Referer": f"https://dps.psx.com.pk/company/{ticker.upper()}",
            },
            timeout=20,
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as exc:
        log.warning("Payouts fetch failed for %s: %s", ticker, exc)
        return []

    items = []
    for row in soup.select("table tr")[1:]:   # skip header row
        cells = row.find_all("td")
        if len(cells) >= 4:
            items.append({
                "date":     cells[0].get_text(strip=True),
                "fin_year": cells[1].get_text(strip=True),
                "details":  cells[2].get_text(strip=True),
                "closure":  cells[3].get_text(strip=True),
            })
    return items


def _payouts_section_html(ticker: str, items: list) -> str:
    if not items:
        return (
            f'<div class="news-section">'
            f'<h2>Payouts — {ticker}</h2>'
            f'<p class="no-news">No payout history found.</p>'
            f'</div>'
        )

    rows = []
    for item in items:
        rows.append(
            f'<tr>'
            f'<td class="date-cell">{item["date"]}</td>'
            f'<td>{item["fin_year"]}</td>'
            f'<td><strong>{item["details"]}</strong></td>'
            f'<td class="date-cell">{item["closure"]}</td>'
            f'</tr>'
        )

    return (
        f'<div class="news-section">'
        f'<h2>Payouts — {ticker}</h2>'
        f'<table class="news-table">'
        f'<thead><tr>'
        f'<th>Date</th><th>Financial Year</th><th>Details</th><th>Book Closure</th>'
        f'</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody>'
        f'</table></div>'
    )


def _manual_announcements_html(ticker: str) -> str:
    entries = ANNOUNCEMENTS.get(ticker.upper(), {})
    if not entries:
        return ""

    cards = []
    for date_key, entry in sorted(entries.items(), reverse=True):
        title    = entry.get("title", "")
        category = entry.get("category", "")
        summary  = entry.get("summary", {})

        rows = "".join(
            f'<tr><td style="color:#555;font-size:12px;white-space:nowrap;'
            f'padding:5px 12px;border-bottom:1px solid #eef0f2;">{k}</td>'
            f'<td style="padding:5px 12px;font-size:13px;border-bottom:1px solid #eef0f2;">{v}</td></tr>'
            for k, v in summary.items()
        )

        cards.append(
            f'<div style="margin-bottom:16px;border:1px solid #dde1e7;border-radius:4px;overflow:hidden;">'
            f'<div style="background:#f4f6f8;padding:8px 12px;border-bottom:1px solid #dde1e7;'
            f'display:flex;align-items:center;gap:10px;">'
            f'<span style="font-size:12px;color:#666;">{date_key}</span>'
            f'<strong style="font-size:13px;">{title}</strong>'
            f'<span class="badge badge-other" style="margin-left:auto;">{category}</span>'
            f'</div>'
            f'<table style="width:100%;border-collapse:collapse;">{rows}</table>'
            f'</div>'
        )

    return (
        f'<div class="news-section">'
        f'<h2>Curated Announcements — {ticker}</h2>'
        f'{"".join(cards)}'
        f'</div>'
    )


def _news_section_html(ticker: str, items: list) -> str:
    if not items:
        return (
            f'<div class="news-section">'
            f'<h2>Announcements — {ticker}</h2>'
            f'<p class="no-news">No announcements found.</p>'
            f'</div>'
        )

    badge_cls = {
        "Financial Results": "badge-fin",
        "Board Meetings":    "badge-bm",
        "Others":            "badge-other",
    }

    rows = []
    for item in items:
        bcls  = badge_cls.get(item["category"], "badge-other")
        doc   = ""
        if item["url"]:
            label = "PDF" if item["url"].endswith(".pdf") else "View"
            doc = f'<a class="doc-link" href="{item["url"]}" target="_blank">{label}</a>'

        title_html = (
            f'<a href="{item["url"]}" target="_blank" '
            f'style="color:#1a1a1a;text-decoration:none;">{item["title"]}</a>'
            if item["url"] else item["title"]
        )
        rows.append(
            f'<tr>'
            f'<td class="date-cell">{item["date"]}</td>'
            f'<td>{title_html}</td>'
            f'<td><span class="badge {bcls}">{item["category"]}</span></td>'
            f'<td>{doc}</td>'
            f'</tr>'
        )

    return (
        f'<div class="news-section">'
        f'<h2>Announcements — {ticker}</h2>'
        f'<table class="news-table">'
        f'<thead><tr>'
        f'<th>Date</th><th>Title</th><th>Category</th><th>Document</th>'
        f'</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody>'
        f'</table></div>'
    )


# ── Chart ─────────────────────────────────────────────────────────────────────

def build_chart(
    df:          pd.DataFrame,
    ticker:      str,
    target_date: dt_date,
    ldcp:        float | None,
    gaps:        list,
    out_path:    Path,
    news_items:  list = None,
    payouts:     list = None,
) -> None:

    intervals   = df["Interval"].tolist()
    day_open    = float(df["Open"].iloc[0])
    date_label  = datetime.strptime(str(target_date), "%Y-%m-%d").strftime("%d %b %Y")
    n           = len(intervals)

    # ── Subplot layout: price (top 72%) + volume (bottom 28%) ────────────────
    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.04,
        row_heights=[0.72, 0.28],
    )

    # ── Candlestick ───────────────────────────────────────────────────────────
    fig.add_trace(
        go.Candlestick(
            x=intervals,
            open=df["Open"],
            high=df["High"],
            low=df["Low"],
            close=df["Close"],
            name=ticker,
            increasing=dict(
                line=dict(color="#009900", width=1.5),
                fillcolor="rgba(0,153,0,0.30)",
            ),
            decreasing=dict(
                line=dict(color="#cc0000", width=1.5),
                fillcolor="rgba(204,0,0,0.30)",
            ),
            whiskerwidth=0.4,
            hoverinfo="x+y",
        ),
        row=1, col=1,
    )

    # ── Volume bars ───────────────────────────────────────────────────────────
    vol_colors = [
        "rgba(0,153,0,0.55)" if c >= o else "rgba(204,0,0,0.55)"
        for o, c in zip(df["Open"], df["Close"])
    ]
    fig.add_trace(
        go.Bar(
            x=intervals,
            y=df["Volume"],
            name="Volume",
            marker_color=vol_colors,
            showlegend=False,
            hovertemplate="%{x}<br>Volume: %{y:,.0f}<extra></extra>",
        ),
        row=2, col=1,
    )

    # ── Gap shading and annotations ───────────────────────────────────────────
    shapes      = []
    annotations = []

    for g in gaps:
        fill_color   = "rgba(0,180,0,0.18)"  if g["up"] else "rgba(210,0,0,0.18)"
        border_color = "rgba(0,140,0,0.70)"  if g["up"] else "rgba(180,0,0,0.70)"
        text_color   = "#006600"             if g["up"] else "#990000"
        direction    = "Gap Up"              if g["up"] else "Gap Down"

        # Shaded rectangle spanning the empty space between the two candles
        shapes.append(dict(
            type="rect",
            xref="x", yref="y",
            x0=g["interval_prev"], x1=g["interval_curr"],
            y0=g["y0"],            y1=g["y1"],
            fillcolor=fill_color,
            line=dict(color=border_color, width=1, dash="dot"),
            layer="below",
        ))

        # Annotation: direction, PKR size, and % size
        gap_label = (
            f"{direction} {g['diff']:+.3f} PKR<br>"
            f"({g['pct']:+.2f}%)"
        )
        annotations.append(dict(
            x=g["interval_curr"],
            y=(g["y0"] + g["y1"]) / 2,
            xref="x", yref="y",
            text=gap_label,
            showarrow=True,
            arrowhead=2,
            arrowsize=0.8,
            arrowcolor=border_color,
            arrowwidth=1.2,
            ax=28, ay=0,
            font=dict(size=9, color=text_color),
            align="left",
            bgcolor="white",
            bordercolor=border_color,
            borderwidth=1,
            opacity=0.9,
        ))

    # ── Key price levels ──────────────────────────────────────────────────────
    price_lines = [
        dict(
            y=day_open, color="#1a6bbf", dash="dash",
            label=f"Day Open  {day_open:.2f}",
        ),
    ]
    if ldcp is not None:
        price_lines.append(dict(
            y=ldcp, color="#d4820a", dash="dot",
            label=f"LDCP  {ldcp:.2f}",
        ))

    for pl in price_lines:
        # Invisible scatter trace to show in legend
        fig.add_trace(
            go.Scatter(
                x=[intervals[0], intervals[-1]],
                y=[pl["y"], pl["y"]],
                mode="lines",
                name=pl["label"],
                line=dict(color=pl["color"], dash=pl["dash"], width=1.5),
                hovertemplate=f"{pl['label']}<extra></extra>",
            ),
            row=1, col=1,
        )

    # ── Layout ────────────────────────────────────────────────────────────────
    # Price range with 3% padding
    price_vals = pd.concat([df["High"], df["Low"]]).dropna()
    if ldcp:
        price_vals = pd.concat([price_vals, pd.Series([ldcp])])
    p_min = price_vals.min()
    p_max = price_vals.max()
    p_pad = (p_max - p_min) * 0.08
    y_min = p_min - p_pad
    y_max = p_max + p_pad

    fig.update_layout(
        title=dict(
            text=f"<b>{ticker}</b>  —  Intraday {date_label}",
            font=dict(size=17, family="Arial"),
        ),
        template="plotly_white",
        height=680,
        hovermode="x unified",
        xaxis_rangeslider_visible=False,
        shapes=shapes,
        annotations=annotations,
        legend=dict(
            orientation="h",
            yanchor="bottom", y=1.01,
            xanchor="left",   x=0,
            font=dict(size=11),
        ),
        margin=dict(l=65, r=55, t=70, b=45),
        font=dict(family="Arial", size=11),
        plot_bgcolor="#fafafa",
        paper_bgcolor="white",
    )

    fig.update_xaxes(
        type="category",
        tickangle=-30,
        showgrid=True,
        gridcolor="#e8e8e8",
        row=1, col=1,
    )
    fig.update_xaxes(
        type="category",
        tickangle=-30,
        showgrid=True,
        gridcolor="#e8e8e8",
        row=2, col=1,
    )
    fig.update_yaxes(
        title_text="Price (PKR)",
        tickformat=".2f",
        showgrid=True,
        gridcolor="#e8e8e8",
        range=[y_min, y_max],
        row=1, col=1,
    )
    fig.update_yaxes(
        title_text="Volume",
        tickformat=",.0f",
        showgrid=False,
        row=2, col=1,
    )

    # ── Save ──────────────────────────────────────────────────────────────────
    out_path.parent.mkdir(parents=True, exist_ok=True)

    chart_div     = fig.to_html(include_plotlyjs="cdn", full_html=False)
    curated_block = _manual_announcements_html(ticker)
    news_block    = _news_section_html(ticker, news_items or [])
    payouts_block = _payouts_section_html(ticker, payouts or [])

    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>{ticker} — Intraday {date_label}</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 0; background: #fff; color: #1a1a1a; }}
    .page {{ max-width: 1200px; margin: 0 auto; padding: 16px 20px 48px; }}
    .news-section {{ margin-top: 32px; }}
    .news-section h2 {{ font-size: 15px; font-weight: 600; color: #333; margin: 0 0 10px;
                        padding-bottom: 8px; border-bottom: 2px solid #e4e4e4; }}
    .news-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    .news-table th {{ text-align: left; padding: 8px 12px; background: #f4f6f8;
                      border-bottom: 2px solid #dde1e7; color: #555; font-size: 12px;
                      font-weight: 600; }}
    .news-table td {{ padding: 8px 12px; border-bottom: 1px solid #eef0f2;
                      vertical-align: middle; }}
    .news-table tr:hover td {{ background: #f7f9ff; }}
    .badge {{ display: inline-block; padding: 2px 9px; border-radius: 3px;
              font-size: 11px; font-weight: 600; color: #fff; white-space: nowrap; }}
    .badge-fin   {{ background: #1a6bbf; }}
    .badge-bm    {{ background: #d4820a; }}
    .badge-other {{ background: #6c757d; }}
    a.doc-link {{ color: #1a6bbf; text-decoration: none; font-size: 12px; }}
    a.doc-link:hover {{ text-decoration: underline; }}
    .date-cell {{ white-space: nowrap; color: #666; font-size: 12px; }}
    .no-news {{ color: #888; font-style: italic; font-size: 13px; }}
  </style>
</head>
<body>
  <div class="page">
    {chart_div}
    {curated_block}
    {payouts_block}
    {news_block}
  </div>
</body>
</html>"""

    out_path.write_text(html, encoding="utf-8")
    log.info("Chart saved: %s", out_path)


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python psx_candles.py <TICKER> [YYYY-MM-DD]")
        print("  e.g. python psx_candles.py KEL")
        print("  e.g. python psx_candles.py KEL 2026-06-10")
        sys.exit(1)

    ticker = sys.argv[1].strip().upper()

    if len(sys.argv) >= 3:
        try:
            target_date = datetime.strptime(sys.argv[2].strip(), "%Y-%m-%d").date()
        except ValueError:
            log.error("Invalid date '%s' — use YYYY-MM-DD format.", sys.argv[2])
            sys.exit(1)
    else:
        target_date = dt_date.today()

    log.info("Ticker: %s | Date: %s", ticker, target_date)

    # Load data
    try:
        all_df = load_all(ticker)
    except Exception as exc:
        log.error("%s", exc)
        sys.exit(1)

    day_df = get_day_df(all_df, target_date)
    if day_df.empty:
        available = sorted(all_df["Date"].unique())
        log.error(
            "No data for %s on %s.\nAvailable dates: %s",
            ticker, target_date,
            ", ".join(str(d) for d in available),
        )
        sys.exit(1)

    if len(day_df) < 2:
        log.warning("Only %d candle found — chart may look sparse.", len(day_df))

    ldcp    = get_ldcp(all_df, target_date)
    gaps    = detect_gaps(day_df)
    news    = scrape_news(ticker)
    payouts = scrape_payouts(ticker)

    log.info(
        "Candles: %d | Gaps: %d | LDCP: %s | Day Open: %.2f | News: %d | Payouts: %d",
        len(day_df), len(gaps),
        f"{ldcp:.2f}" if ldcp else "n/a",
        day_df["Open"].iloc[0],
        len(news),
        len(payouts),
    )

    out_path = CHARTS_DIR / f"{ticker}_{target_date}.html"
    build_chart(day_df, ticker, target_date, ldcp, gaps, out_path,
                news_items=news, payouts=payouts)

    # Also write a date-independent copy for GitHub Pages embedding
    import shutil
    latest_path = CHARTS_DIR / f"{ticker}_candles.html"
    shutil.copy2(out_path, latest_path)

    print(f"[OK] {ticker} {target_date}  —  {len(day_df)} candles, {len(gaps)} gap(s)")
    print(f"     Saved: {out_path.resolve()}")
    print(f"     Latest: {latest_path.resolve()}")

    if not os.environ.get("GITHUB_ACTIONS"):
        webbrowser.open(out_path.resolve().as_uri())


if __name__ == "__main__":
    main()
