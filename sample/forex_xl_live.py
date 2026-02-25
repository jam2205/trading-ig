"""
forex_xl_live.py
----------------
Streams live tick data for the 7 Forex majors (Spreadbet + CFD) from the IG
Lightstreamer feed and pushes updates into a live Excel workbook via xlwings.
Fetches FX news headlines from NewsAPI.org on a configurable interval.

Workbook layout
---------------
  Sheet "Tickers"
    Row 1        : Column headers
    Rows 2–8     : Live Spreadbet (TODAY.IP) ticks for 7 majors
    Row 9        : blank separator
    Rows 10–16   : Live CFD ticks for 7 majors

  Sheet "News"
    Row 1        : Column headers
    Rows 2+      : Latest FX headlines from NewsAPI.org

Requirements
------------
    pip install trading-ig xlwings requests python-dotenv

Setup
-----
  1. Copy .env.example → .env and fill in your IG Demo credentials:
         IG_SERVICE_USERNAME=your_ig_email
         IG_SERVICE_PASSWORD=your_ig_password
         IG_SERVICE_API_KEY=your_ig_api_key
         IG_SERVICE_ACC_TYPE=DEMO
         IG_SERVICE_ACC_NUMBER=your_account_number
         NEWSAPI_KEY=your_newsapi_key
     .env is gitignored so credentials are never committed.
  2. Get your IG Demo API key at:
         My IG → API (top-right menu) → Create API key
  3. Get a free NewsAPI key at https://newsapi.org/register
  3. Open Excel (Windows/macOS required for xlwings) and make sure
     IG_Forex_Live.xlsx is either open or accessible on disk.
  4. Run:
         python -m sample.forex_xl_live

Notes
-----
  - xlwings requires a locally installed Excel (Windows or macOS).
  - The script writes every TICK_REFRESH_SECS and saves the workbook each loop.
  - During market close tickers show the last known values.
  - Ctrl+C performs a clean shutdown (unsubscribes + disconnects).
"""

import logging
import os
import signal
import sys
import time
from math import isnan
from threading import Event, Thread

import requests

# Load .env before importing trading_ig config so env vars are available
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv optional; set env vars manually or use a shell export

from trading_ig import IGService, IGStreamService
from trading_ig.config import config
from trading_ig.streamer.manager import StreamingManager

from sample.forex_epics import ALL_EPICS, CFD_EPICS, EPIC_TO_META, FOREX_MAJORS, SB_EPICS

# ── User configuration ────────────────────────────────────────────────────────

# Get your free key at https://newsapi.org/register
NEWSAPI_KEY = os.environ.get("NEWSAPI_KEY", "YOUR_NEWSAPI_KEY_HERE")

# NewsAPI query — covers all 7 majors broadly; customise as needed
NEWS_QUERY = (
    "EURUSD OR GBPUSD OR USDJPY OR AUDUSD OR USDCAD OR USDCHF OR NZDUSD"
    " OR forex currency"
)

# How often to refresh from NewsAPI (seconds). Free tier: 100 req/day.
NEWS_REFRESH_SECS = 300   # every 5 minutes

# How often to write ticks to Excel (seconds)
TICK_REFRESH_SECS = 1

# Maximum news rows written to the News sheet
MAX_NEWS_ROWS = 25

# Workbook filename. Set to None to always create a new workbook.
WORKBOOK_NAME = "IG_Forex_Live.xlsx"

# ── Sheet structure ───────────────────────────────────────────────────────────

TICKER_SHEET = "Tickers"
NEWS_SHEET = "News"

TICKER_HEADERS = [
    "EPIC",
    "Type",
    "Pair",
    "Bid",
    "Offer",
    "Last Price",
    "Volume",
    "Incr Vol",
    "Day Open",
    "Net Chg",
    "% Chg",
    "Day High",
    "Day Low",
    "Updated",
]

NEWS_HEADERS = ["Currency", "Headline", "Source", "Published", "URL"]

# SB rows: 2–8, blank row 9, CFD rows: 10–16
_SB_START = 2
_CFD_START = _SB_START + len(FOREX_MAJORS) + 1   # +1 for blank separator

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

# ── Helpers ───────────────────────────────────────────────────────────────────


def _fmt(val, decimals: int = 5):
    """Return a rounded float or empty string for nan/None."""
    if val is None:
        return ""
    try:
        if isnan(float(val)):
            return ""
    except (TypeError, ValueError):
        return ""
    return round(float(val), decimals)


def _ticker_row(epic: str, ticker, meta: dict) -> list:
    """Build a single row list from a Ticker object for Excel."""
    ts = ""
    if getattr(ticker, "timestamp", None):
        ts = ticker.timestamp.strftime("%H:%M:%S")
    return [
        epic,
        meta.get("type", ""),
        meta.get("label", ""),
        _fmt(ticker.bid),
        _fmt(ticker.offer),
        _fmt(ticker.last_traded_price),
        ticker.last_traded_volume or "",
        ticker.incr_volume or "",
        _fmt(ticker.day_open_mid),
        _fmt(ticker.day_net_change_mid),
        _fmt(ticker.day_percent_change_mid, 4),
        _fmt(ticker.day_high),
        _fmt(ticker.day_low),
        ts,
    ]


def _placeholder_row(epic: str, meta: dict) -> list:
    """Row with epic/type/pair filled but tick fields blank (not yet received)."""
    return [epic, meta.get("type", ""), meta.get("label", "")] + [""] * 11


# ── Workbook setup ────────────────────────────────────────────────────────────


def _setup_workbook():
    """
    Open or create the Excel workbook.
    Tries: open books → disk file → new book.
    Returns an xlwings Book object.
    """
    try:
        import xlwings as xw
    except ImportError:
        raise ImportError(
            "xlwings is required: pip install xlwings\n"
            "Excel must be installed on Windows or macOS."
        )

    if WORKBOOK_NAME:
        # 1. Already open?
        for wb in xw.books:
            if wb.name in (WORKBOOK_NAME, os.path.basename(WORKBOOK_NAME)):
                logger.info(f"Using open workbook: {wb.name}")
                return wb
        # 2. Exists on disk?
        if os.path.exists(WORKBOOK_NAME):
            logger.info(f"Opening workbook from disk: {WORKBOOK_NAME}")
            return xw.Book(WORKBOOK_NAME)

    # 3. Create new
    logger.info("Creating new workbook")
    wb = xw.Book()
    if WORKBOOK_NAME:
        wb.save(WORKBOOK_NAME)
    return wb


def _ensure_sheets(wb) -> tuple:
    """Make sure Tickers and News sheets exist; return (ws_tick, ws_news)."""
    sheet_names = [s.name for s in wb.sheets]
    if TICKER_SHEET not in sheet_names:
        wb.sheets.add(TICKER_SHEET)
    if NEWS_SHEET not in sheet_names:
        wb.sheets.add(NEWS_SHEET)
    return wb.sheets[TICKER_SHEET], wb.sheets[NEWS_SHEET]


def _write_static_layout(ws_tick, ws_news):
    """
    Write headers and section labels (called once on startup).
    These don't change on each tick cycle.
    """
    # Tickers sheet
    ws_tick["A1"].value = TICKER_HEADERS
    ws_tick["A1"].font.bold = True

    # Section labels in column A (above each data block)
    ws_tick[f"A{_SB_START - 1}"].value = "SPREADBET (TODAY.IP)"   # row 1 if SB_START=2
    ws_tick[f"A{_CFD_START - 1}"].value = "CFD (CFD.IP)"

    # News sheet
    ws_news["A1"].value = NEWS_HEADERS
    ws_news["A1"].font.bold = True


# ── Excel write functions ─────────────────────────────────────────────────────


def write_tickers(ws_tick, sm: StreamingManager):
    """Write all 14 ticker rows to the Tickers sheet in one pass."""
    sb_data = []
    cfd_data = []

    for epic in SB_EPICS:
        meta = EPIC_TO_META.get(epic, {})
        ticker = sm.tickers.get(epic)
        row = _ticker_row(epic, ticker, meta) if ticker else _placeholder_row(epic, meta)
        sb_data.append(row)

    for epic in CFD_EPICS:
        meta = EPIC_TO_META.get(epic, {})
        ticker = sm.tickers.get(epic)
        row = _ticker_row(epic, ticker, meta) if ticker else _placeholder_row(epic, meta)
        cfd_data.append(row)

    # Write both blocks; xlwings accepts a 2-D list
    ws_tick[f"A{_SB_START}"].value = sb_data
    ws_tick[f"A{_CFD_START}"].value = cfd_data


def write_news(ws_news, headlines: list):
    """Write news headlines to the News sheet."""
    if not headlines:
        return
    rows = []
    for article in headlines[:MAX_NEWS_ROWS]:
        rows.append([
            article.get("pair", "FX"),
            article.get("title", ""),
            article.get("source", ""),
            article.get("published", ""),
            article.get("url", ""),
        ])
    ws_news[f"A2"].value = rows


# ── News poller thread ────────────────────────────────────────────────────────


class NewsPoller(Thread):
    """
    Background daemon thread that polls NewsAPI.org for FX news.
    Access .headlines for the latest list of article dicts.
    """

    def __init__(self, stop_event: Event):
        super().__init__(name="NewsPoller", daemon=True)
        self._stop = stop_event
        self.headlines: list = []
        # Fetch immediately on start
        self._refresh()

    def run(self):
        while not self._stop.wait(NEWS_REFRESH_SECS):
            self._refresh()

    def _refresh(self):
        if NEWSAPI_KEY == "YOUR_NEWSAPI_KEY_HERE":
            logger.warning(
                "NEWSAPI_KEY not set — news feed disabled. "
                "Set env var NEWSAPI_KEY or edit NEWSAPI_KEY in forex_xl_live.py"
            )
            return
        try:
            resp = requests.get(
                "https://newsapi.org/v2/everything",
                params={
                    "q": NEWS_QUERY,
                    "language": "en",
                    "sortBy": "publishedAt",
                    "pageSize": MAX_NEWS_ROWS,
                    "apiKey": NEWSAPI_KEY,
                },
                timeout=10,
            )
            resp.raise_for_status()
            articles = resp.json().get("articles", [])
            rows = []
            for a in articles:
                published = a.get("publishedAt", "")[:16].replace("T", " ")
                rows.append({
                    "pair": "FX",
                    "title": a.get("title", ""),
                    "source": a.get("source", {}).get("name", ""),
                    "published": published,
                    "url": a.get("url", ""),
                })
            self.headlines = rows
            logger.info(f"NewsPoller: fetched {len(rows)} articles")
        except requests.RequestException as exc:
            logger.warning(f"NewsPoller request error: {exc}")
        except Exception as exc:
            logger.warning(f"NewsPoller unexpected error: {exc}")


# ── Main ──────────────────────────────────────────────────────────────────────


def main():
    # ── IG connection ─────────────────────────────────────────────────────────
    print("Connecting to IG Markets...")
    ig_service = IGService(
        config.username,
        config.password,
        config.api_key,
        config.acc_type,
        acc_number=config.acc_number,
    )
    ig = IGStreamService(ig_service)
    ig.create_session(version="3")
    sm = StreamingManager(ig)

    print(f"Subscribing to {len(ALL_EPICS)} epics (7 majors × Spreadbet + CFD)...")
    for epic in ALL_EPICS:
        sm.start_tick_subscription(epic)
        print(f"  + {epic}")

    # ── Excel workbook ────────────────────────────────────────────────────────
    print(f"\nOpening Excel workbook: {WORKBOOK_NAME or '(new)'}")
    wb = _setup_workbook()
    ws_tick, ws_news = _ensure_sheets(wb)
    _write_static_layout(ws_tick, ws_news)

    # ── News poller ───────────────────────────────────────────────────────────
    stop_event = Event()
    news_poller = NewsPoller(stop_event)
    news_poller.start()
    if NEWSAPI_KEY == "YOUR_NEWSAPI_KEY_HERE":
        print(
            "\n[WARNING] NEWSAPI_KEY not configured — News sheet will be empty.\n"
            "  Set env var:  export NEWSAPI_KEY=your_key\n"
            "  Free key at:  https://newsapi.org/register\n"
        )

    # ── Graceful shutdown ─────────────────────────────────────────────────────
    def _shutdown(signum, frame):
        print("\nShutting down — unsubscribing and saving workbook...")
        stop_event.set()
        try:
            sm.stop_subscriptions()
        except Exception:
            pass
        try:
            wb.save()
        except Exception:
            pass
        sys.exit(0)

    signal.signal(signal.SIGINT, _shutdown)
    signal.signal(signal.SIGTERM, _shutdown)

    # ── Main loop ─────────────────────────────────────────────────────────────
    print(
        f"\nLive updating Excel every {TICK_REFRESH_SECS}s. "
        f"News refreshes every {NEWS_REFRESH_SECS // 60}min. "
        "Press Ctrl+C to stop.\n"
    )
    last_news_write = 0.0

    while True:
        try:
            write_tickers(ws_tick, sm)

            # Write news when the poller has new data
            if time.time() - last_news_write > NEWS_REFRESH_SECS:
                write_news(ws_news, news_poller.headlines)
                last_news_write = time.time()

            wb.save()

        except Exception as exc:
            logger.warning(f"Main loop error: {exc}")

        time.sleep(TICK_REFRESH_SECS)


if __name__ == "__main__":
    main()
