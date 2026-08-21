"""
Booked Profit Dashboard — Streamlit version (premium dark theme).

Runs entirely in the cloud (Streamlit Community Cloud, free): logs into
Angel One, holds a live WebSocket price feed in a background thread, and
renders the same open/closed positions view as the original local HTML
dashboard — but reachable from any device, with your PC turned off.

Local files reused unchanged from the original project:
    positions_builder.py   — FIFO matching (open + closed positions)
    token_resolver.py      — Angel One instrument-master token lookup

Secrets required (set in Streamlit Cloud's "Secrets" panel, never in code):
    APP_PASSWORD        — shared password gate for viewers
    ANGEL_API_KEY
    ANGEL_CLIENT_CODE
    ANGEL_PASSWORD
    ANGEL_TOTP_SECRET

The trade ledger Excel is uploaded through the sidebar each session (or
bundled in the repo as securities_today.xlsx as a fallback default) rather
than read from a local path, since Streamlit Cloud has no access to your PC.

LIVE UPDATES
    Instead of refreshing the whole page every few seconds (old
    st_autorefresh approach — flickers, resets scroll position, reruns
    everything including the sidebar), the KPI cards + tables now live
    inside an @st.fragment(run_every=...) block. Only that fragment reruns
    on its own clock, reading whatever ticks have landed in the background
    WebSocket thread since the last run — so the screen updates in near
    real time, tick by tick, without touching the rest of the app.
    Requires streamlit >= 1.33.
"""
import io
import os
import re
import threading
import time
from datetime import datetime
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import pyotp
import streamlit as st
from SmartApi import SmartConnect
from SmartApi.smartWebSocketV2 import SmartWebSocketV2

from positions_builder import load_trade_ledger, build_positions
from token_resolver import resolve_all

IST = ZoneInfo("Asia/Kolkata")

st.set_page_config(
    page_title="Booked Profit Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MASTER_CACHE_PATH = "instrument_master_cache.json"
MASTER_CACHE_MAX_AGE_HOURS = 20
SUBSCRIBE_MODE = 2
MAX_TOKENS_PER_SUBSCRIBE = 1000
TICK_REFRESH_SECONDS = 1  # how often the live fragment re-renders


# ── Premium theme ──────────────────────────────────────────────────────────
def inject_theme():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');

        :root {
            --bg: #070b14;
            --panel: #10161f;
            --panel-2: #131b27;
            --border: #1e2733;
            --accent: #34d5c8;
            --accent-2: #7c5cff;
            --pos: #2ee6a6;
            --neg: #ff5c7a;
            --muted: #7f8ba3;
            --text: #eaf0f7;
        }

        html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }

        .stApp {
            background:
                radial-gradient(circle at 15% 0%, rgba(124,92,255,0.10), transparent 40%),
                radial-gradient(circle at 85% 10%, rgba(52,213,200,0.08), transparent 35%),
                var(--bg);
            color: var(--text);
        }

        section[data-testid="stSidebar"] {
            background: var(--panel);
            border-right: 1px solid var(--border);
        }

        div.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1300px; }

        /* Header */
        .db-header {
            display: flex; align-items: center; justify-content: space-between;
            padding-bottom: 6px; margin-bottom: 22px;
            border-bottom: 1px solid var(--border);
        }
        .db-title { display: flex; align-items: center; gap: 12px; }
        .db-title h1 {
            font-size: 1.65rem; font-weight: 800; letter-spacing: -0.02em; margin: 0;
            background: linear-gradient(90deg, #ffffff, #b9c4d6);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }
        .db-icon {
            width: 40px; height: 40px; border-radius: 12px;
            background: linear-gradient(135deg, var(--accent), var(--accent-2));
            display: flex; align-items: center; justify-content: center;
            font-size: 1.15rem; box-shadow: 0 0 24px rgba(52,213,200,0.35);
        }

        .status-pill {
            display: inline-flex; align-items: center; gap: 8px;
            padding: 6px 14px; border-radius: 999px; font-size: 0.78rem; font-weight: 600;
            border: 1px solid var(--border); background: var(--panel-2); color: var(--muted);
        }
        .status-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--muted); }
        .status-live .status-dot { background: var(--pos); box-shadow: 0 0 8px var(--pos); animation: pulse 1.4s ease-in-out infinite; }
        .status-live { color: var(--pos); border-color: rgba(46,230,166,0.25); }
        .status-error .status-dot { background: var(--neg); }
        .status-error { color: var(--neg); border-color: rgba(255,92,122,0.25); }
        .status-warn .status-dot { background: #f5b942; }
        .status-warn { color: #f5b942; border-color: rgba(245,185,66,0.25); }
        @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.35; } }

        /* KPI cards */
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 26px; }
        .kpi-card {
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border); border-radius: 16px; padding: 18px 20px;
            position: relative; overflow: hidden;
        }
        .kpi-card::before {
            content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px;
            background: linear-gradient(90deg, var(--accent), var(--accent-2)); opacity: 0.85;
        }
        .kpi-label { font-size: 0.74rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; color: var(--muted); margin-bottom: 8px; }
        .kpi-value { font-family: 'JetBrains Mono', monospace; font-size: 1.55rem; font-weight: 700; letter-spacing: -0.01em; }
        .kpi-pos { color: var(--pos); }
        .kpi-neg { color: var(--neg); }
        .kpi-sub { font-size: 0.72rem; color: var(--muted); margin-top: 6px; font-family: 'JetBrains Mono', monospace; }

        /* Section headers */
        .section-label {
            display: flex; align-items: center; gap: 8px;
            font-size: 0.95rem; font-weight: 700; margin: 6px 0 10px 0; color: var(--text);
        }
        .section-label .badge {
            font-size: 0.68rem; font-weight: 600; padding: 2px 9px; border-radius: 999px;
            background: var(--panel-2); border: 1px solid var(--border); color: var(--muted);
        }

        /* Tabs */
        .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--border); }
        .stTabs [data-baseweb="tab"] {
            background: transparent; color: var(--muted); font-weight: 600; padding: 10px 18px;
        }
        .stTabs [aria-selected="true"] {
            color: var(--accent) !important; border-bottom: 2px solid var(--accent) !important;
        }

        /* Top status + live clock bar */
        .top-bar {
            display: flex; align-items: center; justify-content: space-between;
            background: var(--panel-2); border: 1px solid var(--border); border-radius: 12px;
            padding: 10px 18px; margin-bottom: 20px;
        }
        .top-clock {
            font-family: 'JetBrains Mono', monospace; font-size: 0.82rem; color: var(--text);
            display: flex; align-items: center; gap: 10px;
        }
        .top-clock .date-part { color: var(--muted); }
        .top-clock .tz-badge {
            font-size: 0.65rem; font-weight: 700; color: var(--accent);
            background: rgba(52,213,200,0.1); border: 1px solid rgba(52,213,200,0.25);
            padding: 1px 7px; border-radius: 999px;
        }

        /* Position cards */
        .pos-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 12px; margin-bottom: 8px; }
        .pos-card {
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border); border-radius: 14px; padding: 14px 16px;
            transition: border-color 0.2s ease;
        }
        .pos-card:hover { border-color: rgba(52,213,200,0.35); }
        .pos-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }
        .pos-symbol { font-weight: 700; font-size: 0.98rem; letter-spacing: -0.01em; }
        .pos-tags { display: flex; gap: 5px; margin-top: 4px; }
        .pos-chip {
            font-size: 0.62rem; font-weight: 700; padding: 1px 7px; border-radius: 999px;
            background: var(--panel); border: 1px solid var(--border); color: var(--muted);
            text-transform: uppercase; letter-spacing: 0.03em;
        }
        .pos-chip.long { color: var(--pos); border-color: rgba(46,230,166,0.3); }
        .pos-chip.short { color: var(--neg); border-color: rgba(255,92,122,0.3); }
        .pos-daychg {
            font-family: 'JetBrains Mono', monospace; font-size: 0.72rem; font-weight: 700;
            padding: 3px 9px; border-radius: 999px; white-space: nowrap;
        }
        .day-pos { background: rgba(46,230,166,0.12); color: var(--pos); }
        .day-neg { background: rgba(255,92,122,0.12); color: var(--neg); }
        .pos-rows { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 14px; margin-bottom: 10px; }
        .pos-row-label { font-size: 0.68rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
        .pos-row-value { font-family: 'JetBrains Mono', monospace; font-size: 0.82rem; font-weight: 600; }
        .pos-mtm-block {
            display: flex; justify-content: space-between; align-items: center;
            border-top: 1px solid var(--border); padding-top: 10px;
        }
        .pos-mtm-label { font-size: 0.68rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
        .pos-mtm-value { font-family: 'JetBrains Mono', monospace; font-size: 1.02rem; font-weight: 700; }
        .pos-tick { font-size: 0.66rem; color: var(--muted); font-family: 'JetBrains Mono', monospace; }

        /* Closed position cards */
        .closed-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(230px, 1fr)); gap: 12px; margin-bottom: 20px; }
        .closed-card {
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border); border-radius: 14px; padding: 13px 15px;
        }
        .closed-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px; }
        .closed-symbol { font-weight: 700; font-size: 0.92rem; }
        .closed-badge {
            font-size: 0.62rem; font-weight: 700; padding: 1px 7px; border-radius: 999px;
            background: var(--panel); border: 1px solid var(--border); color: var(--muted);
        }
        .closed-rows { display: flex; justify-content: space-between; font-size: 0.72rem; color: var(--muted); margin-bottom: 4px; font-family: 'JetBrains Mono', monospace; }
        .closed-pnl { font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 0.95rem; margin-top: 8px; }

        .chart-card {
            background: var(--panel-2); border: 1px solid var(--border); border-radius: 14px;
            padding: 16px 18px 6px 18px; margin-bottom: 18px;
        }

        /* Dataframe polish (still used where a raw table makes sense) */
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border); border-radius: 12px; overflow: hidden;
        }

        div[data-testid="stMetric"] { display: none; }  /* using custom KPI cards instead */
        </style>
        """,
        unsafe_allow_html=True,
    )


def status_pill(state: str, label: str) -> str:
    cls = {"live": "status-live", "error": "status-error", "warn": "status-warn"}.get(state, "")
    return f'<span class="status-pill {cls}"><span class="status-dot"></span>{label}</span>'


def flat(html: str) -> str:
    """Collapse a multi-line, indented HTML f-string to a single line.

    Streamlit's markdown renderer runs HTML through a CommonMark parser
    before allowing it through. An HTML block continues being treated as
    raw HTML only until a blank line; after that, any line indented 4+
    spaces (which our nested Python f-strings produce naturally) is read
    as an *indented code block* and shown as literal text instead of being
    rendered — this is what caused the closed-position cards to print
    "<div class=..." as visible text after the first card. Stripping all
    newlines/indentation removes any chance of a stray blank line or deep
    indentation confusing the parser, regardless of how deeply the
    generating Python code is nested.
    """
    return re.sub(r"\s*\n\s*", "", html.strip())


def kpi_card(label, value, positive=None, sub=None):
    cls = "kpi-pos" if positive is True else ("kpi-neg" if positive is False else "")
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return flat(f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value {cls}">{value}</div>
            {sub_html}
        </div>
    """)


# ── Password gate ─────────────────────────────────────────────────────────
def check_password():
    def _submit():
        st.session_state["_pw_ok"] = (
            st.session_state.get("_pw_input", "") == st.secrets.get("APP_PASSWORD", "")
        )

    if st.session_state.get("_pw_ok"):
        return True

    inject_theme()
    st.markdown(
        flat("""
        <div class="db-title" style="justify-content:center; margin: 60px 0 24px 0;">
            <div class="db-icon">📊</div>
            <h1>Booked Profit Dashboard</h1>
        </div>
        """),
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        st.text_input("Password", type="password", key="_pw_input", on_change=_submit)
        if "_pw_ok" in st.session_state and not st.session_state["_pw_ok"]:
            st.error("Incorrect password.")
    return False


# ── Background engine: login + FIFO build + token resolve + live WS feed ──
class LiveEngine:
    def __init__(self, ledger_rows):
        self.lock = threading.Lock()
        self.latest_prices = {}       # token -> tick dict
        self.token_to_symbol = {}     # token -> meta
        self.closed_positions = []
        self.booked_mtm_total = 0.0
        self.status = "starting"
        self.error = None
        self.last_tick_ts = None      # datetime of most recently received tick
        # Positions/token-resolution/login are all network calls and can take
        # several seconds. Do them in a background thread so get_engine()
        # (and therefore the page) returns to the browser immediately instead
        # of blocking behind Streamlit's "connecting" spinner. The UI polls
        # self.status via the live fragment until this flips to "live".
        threading.Thread(target=self._start, args=(ledger_rows,), daemon=True).start()

    def _start(self, ledger_rows):
        try:
            open_positions, closed_positions = build_positions(ledger_rows, verbose=False)
            self.closed_positions = closed_positions
            self.booked_mtm_total = round(sum(c["BookedPnL"] for c in closed_positions), 2)

            resolved, unresolved = resolve_all(open_positions, MASTER_CACHE_PATH, MASTER_CACHE_MAX_AGE_HOURS)
            if not resolved:
                self.status = "error"
                self.error = "Nothing resolved from the ledger — check symbol/exchange/expiry spelling."
                return

            for r in resolved:
                src = next((p for p in open_positions
                            if p["Symbol"] == r["symbol"] and p["Exchange"] == r["exchange"]), {})
                self.token_to_symbol[r["token"]] = {
                    "symbol": r["symbol"], "exchange": r["exchange"],
                    "qty": r.get("qty"), "avgPrice": r.get("avgPrice"),
                    "segment": src.get("Segment", "Other"),
                    "positionType": src.get("PositionType", "LONG"),
                }

            totp = pyotp.TOTP(st.secrets["ANGEL_TOTP_SECRET"]).now()
            sc = SmartConnect(api_key=st.secrets["ANGEL_API_KEY"])
            data = sc.generateSession(st.secrets["ANGEL_CLIENT_CODE"], st.secrets["ANGEL_PASSWORD"], totp)
            if not data.get("status"):
                self.status = "error"
                self.error = f"Angel One login failed: {data}"
                return
            jwt_token = data["data"]["jwtToken"]
            feed_token = data["data"]["feedToken"]

            self._start_ws(jwt_token, feed_token, resolved)
            self.status = "live"
        except Exception as e:
            self.status = "error"
            self.error = str(e)

    def _chunk(self, resolved):
        by_exch = {}
        for r in resolved:
            by_exch.setdefault(r["exchangeType"], []).append(r["token"])
        batches, current, current_count = [], [], 0
        for exch_type, tokens in by_exch.items():
            for i in range(0, len(tokens), 200):
                chunk = tokens[i:i + 200]
                if current_count + len(chunk) > MAX_TOKENS_PER_SUBSCRIBE:
                    batches.append(current)
                    current, current_count = [], 0
                current.append({"exchangeType": exch_type, "tokens": chunk})
                current_count += len(chunk)
        if current:
            batches.append(current)
        return batches

    def _start_ws(self, jwt_token, feed_token, resolved):
        sws = SmartWebSocketV2(
            auth_token=jwt_token, api_key=st.secrets["ANGEL_API_KEY"],
            client_code=st.secrets["ANGEL_CLIENT_CODE"], feed_token=feed_token,
        )
        batches = self._chunk(resolved)

        def on_open(wsapp):
            for i, token_list in enumerate(batches):
                sws.subscribe(f"feed_{i}", SUBSCRIBE_MODE, token_list)
                time.sleep(0.3)

        def on_data(wsapp, message):
            token = str(message.get("token"))
            meta = self.token_to_symbol.get(token, {})
            with self.lock:
                prev = self.latest_prices.get(token, {})
                ltp = message.get("last_traded_price", 0) / 100.0
                qty = meta.get("qty")
                avg_price = meta.get("avgPrice")
                now = datetime.now(IST)
                tick = {
                    "token": token, "symbol": meta.get("symbol", token),
                    "exchange": meta.get("exchange", ""), "segment": meta.get("segment", "Other"),
                    "positionType": meta.get("positionType", "LONG"), "ltp": ltp,
                    "prev_ltp": prev.get("ltp"),
                    "close": (message.get("closed_price", 0) / 100.0 if message.get("closed_price") else prev.get("close")),
                    "open": (message.get("open_price_of_the_day", 0) / 100.0 if message.get("open_price_of_the_day") else prev.get("open")),
                    "volume": message.get("volume_trade_for_the_day"), "qty": qty, "avgPrice": avg_price,
                    "mtm": round((ltp - avg_price) * qty, 2) if (qty is not None and avg_price is not None) else None,
                    "ts": now.isoformat(timespec="seconds"),
                }
                self.latest_prices[token] = tick
                self.last_tick_ts = now

        def on_error(wsapp, error):
            with self.lock:
                self.status = "error"
                self.error = f"WebSocket error: {error}"

        def on_close(wsapp, *a):
            with self.lock:
                self.status = "disconnected"

        sws.on_open = on_open
        sws.on_data = on_data
        sws.on_error = on_error
        sws.on_close = on_close
        threading.Thread(target=sws.connect, daemon=True).start()

    def snapshot(self):
        with self.lock:
            return list(self.latest_prices.values()), self.last_tick_ts


@st.cache_resource(show_spinner="Reading ledger...")
def get_engine(ledger_bytes: bytes):
    rows = load_trade_ledger(io.BytesIO(ledger_bytes))
    return LiveEngine(rows)


# ── Formatting helpers ─────────────────────────────────────────────────────
def fmt_money(x):
    if x is None:
        return "-"
    return f"(₹{abs(x):,.2f})" if x < 0 else f"₹{x:,.2f}"


def fmt_qty(qty):
    if qty is None:
        return "-"
    return f"{abs(round(qty)):,}"


def fmt_time(ts):
    """ts is an ISO datetime string (possibly tz-aware); show HH:MM:SS only."""
    if not ts:
        return "-"
    try:
        dt = datetime.fromisoformat(ts)
        return dt.strftime("%H:%M:%S")
    except ValueError:
        return ts


def alt_dark(chart):
    """Apply a shared dark, transparent-background theme to an Altair chart."""
    return (
        chart.configure_view(strokeWidth=0)
        .configure_axis(
            gridColor="#1e2733", domainColor="#1e2733", tickColor="#1e2733",
            labelColor="#7f8ba3", titleColor="#7f8ba3", labelFontSize=10.5, titleFontSize=11,
        )
        .configure_legend(labelColor="#7f8ba3", titleColor="#7f8ba3")
        .properties(background="transparent")
    )


def style_pnl_table(df, cols):
    """Return a pandas Styler that colors P&L-type columns green/red."""
    def _color(v):
        if pd.isna(v):
            return ""
        return "color: #2ee6a6; font-weight:600;" if v >= 0 else "color: #ff5c7a; font-weight:600;"

    fmt = {c: "₹{:,.2f}".format for c in cols if c in df.columns}
    styler = df.style.format(fmt)
    # pandas >=2.1 renamed Styler.applymap -> Styler.map (and removed
    # applymap entirely in pandas 3.x), so pick whichever exists at runtime.
    color_fn = styler.map if hasattr(styler, "map") else styler.applymap
    for c in cols:
        if c in df.columns:
            styler = color_fn(_color, subset=[c])
            color_fn = styler.map if hasattr(styler, "map") else styler.applymap
    return styler


# ── Live fragment: KPI cards + open/closed tables, refreshes on its own ───
@st.fragment(run_every=TICK_REFRESH_SECONDS)
def render_live(engine: "LiveEngine"):
    if engine.status == "error":
        st.markdown(status_pill("error", "Feed error"), unsafe_allow_html=True)
        st.error(f"Feed failed to start: {engine.error}")
        return
    elif engine.status == "starting":
        st.markdown(status_pill("warn", "Starting up..."), unsafe_allow_html=True)
        st.caption("Resolving instrument tokens and logging into Angel One — this can take a few seconds on a cold start.")
        return

    ticks, last_tick_ts = engine.snapshot()

    now_ist = datetime.now(IST)
    if engine.status == "disconnected":
        status_html = status_pill("error", "Disconnected — restart feed in sidebar")
    else:
        age = f"{(now_ist - last_tick_ts).seconds}s ago" if last_tick_ts else "waiting for first tick..."
        status_html = status_pill("live", f"Live • last tick {age}")

    st.markdown(
        flat(f"""
        <div class="top-bar">
            {status_html}
            <div class="top-clock">
                <span class="date-part">{now_ist.strftime("%A, %d %b %Y")}</span>
                <span>{now_ist.strftime("%H:%M:%S")}</span>
                <span class="tz-badge">IST</span>
            </div>
        </div>
        """),
        unsafe_allow_html=True,
    )

    equity = [t for t in ticks if t.get("segment") == "Equity"]
    fo = [t for t in ticks if t.get("segment") == "F&O"]

    def seg_totals(rows):
        # Capital deployed must use abs(qty): a short position has a
        # negative qty, and summing signed qty*avgPrice was silently
        # *subtracting* short positions' cost basis from the total instead
        # of adding it — that was the "wrong Investment Value" bug.
        buy_value = sum(abs(r["qty"] or 0) * (r["avgPrice"] or 0) for r in rows)
        mtm = sum(r["mtm"] for r in rows if r["mtm"] is not None)
        # Today's move only (ltp vs previous close), signed qty on purpose:
        # for a short, a falling price is a gain, and (ltp-close) is
        # negative while qty is negative too, so the product comes out
        # positive automatically.
        day_pnl = sum(
            (r["ltp"] - r["close"]) * r["qty"]
            for r in rows if r.get("close") not in (None, 0) and r.get("qty") is not None
        )
        return buy_value, mtm, day_pnl

    eq_buy, eq_mtm, eq_day = seg_totals(equity)
    fo_buy, fo_mtm, fo_day = seg_totals(fo)
    investment_value = eq_buy + fo_buy
    current_mtm = eq_mtm + fo_mtm
    total_mtm = engine.booked_mtm_total + current_mtm
    day_pnl_total = eq_day + fo_day
    day_pnl_pct = (day_pnl_total / investment_value * 100) if investment_value else 0.0

    st.markdown(
        flat(f"""
        <div class="kpi-grid">
            {kpi_card("Investment Value", fmt_money(investment_value), sub=f"{len(ticks)} instruments · cost basis")}
            {kpi_card("Today's P&amp;L", fmt_money(day_pnl_total), positive=day_pnl_total >= 0, sub=f"{'+' if day_pnl_pct >= 0 else ''}{day_pnl_pct:.2f}% vs prev close")}
            {kpi_card("Open MTM", fmt_money(current_mtm), positive=current_mtm >= 0, sub=f"Eq {fmt_money(eq_mtm)} · F&amp;O {fmt_money(fo_mtm)}")}
            {kpi_card("Booked MTM", fmt_money(engine.booked_mtm_total), positive=engine.booked_mtm_total >= 0, sub=f"{len(engine.closed_positions)} closed positions")}
            {kpi_card("Total MTM", fmt_money(total_mtm), positive=total_mtm >= 0, sub="Booked + Open, live")}
        </div>
        """),
        unsafe_allow_html=True,
    )

    tab_open, tab_closed = st.tabs(["📈 Open positions", "✅ Closed positions"])

    with tab_open:
        for label, rows in [("Equity", equity), ("F&O", fo)]:
            count_badge = f'<span class="badge">{len(rows)}</span>'
            st.markdown(f'<div class="section-label">{label} {count_badge}</div>', unsafe_allow_html=True)
            if not rows:
                st.caption("No open positions in this segment.")
                continue

            cards = []
            for r in sorted(rows, key=lambda x: (x.get("mtm") or 0)):
                pos_type = (r.get("positionType") or "LONG").upper()
                type_cls = "long" if pos_type == "LONG" else "short"
                close = r.get("close")
                day_pct = ((r["ltp"] - close) / close * 100) if close else None
                day_cls = "day-pos" if (day_pct or 0) >= 0 else "day-neg"
                day_txt = f"{'▲' if (day_pct or 0) >= 0 else '▼'} {abs(day_pct):.2f}%" if day_pct is not None else "–"
                mtm = r.get("mtm")
                mtm_cls = "kpi-pos" if (mtm or 0) >= 0 else "kpi-neg"
                cards.append(flat(f"""
                    <div class="pos-card">
                        <div class="pos-top">
                            <div>
                                <div class="pos-symbol">{r.get('symbol', '-')}</div>
                                <div class="pos-tags">
                                    <span class="pos-chip">{r.get('exchange', '-')}</span>
                                    <span class="pos-chip {type_cls}">{pos_type}</span>
                                </div>
                            </div>
                            <span class="pos-daychg {day_cls}">{day_txt}</span>
                        </div>
                        <div class="pos-rows">
                            <div><div class="pos-row-label">Qty</div><div class="pos-row-value">{fmt_qty(r.get('qty'))}</div></div>
                            <div><div class="pos-row-label">Avg Price</div><div class="pos-row-value">{fmt_money(r.get('avgPrice'))}</div></div>
                            <div><div class="pos-row-label">CMP</div><div class="pos-row-value">{fmt_money(r.get('ltp'))}</div></div>
                            <div><div class="pos-row-label">Prev Close</div><div class="pos-row-value">{fmt_money(close)}</div></div>
                        </div>
                        <div class="pos-mtm-block">
                            <div>
                                <div class="pos-mtm-label">MTM</div>
                                <div class="pos-mtm-value {mtm_cls}">{fmt_money(mtm)}</div>
                            </div>
                            <div class="pos-tick">🕐 {fmt_time(r.get('ts'))}</div>
                        </div>
                    </div>
                """))
            st.markdown(f'<div class="pos-grid">{"".join(cards)}</div>', unsafe_allow_html=True)

    with tab_closed:
        if not engine.closed_positions:
            st.caption("No closed positions in this ledger.")
        else:
            closed = engine.closed_positions
            wins = sum(1 for c in closed if c["BookedPnL"] >= 0)
            losses = len(closed) - wins
            win_rate = wins / len(closed) * 100 if closed else 0

            st.markdown(
                f'<div class="section-label">All closed positions <span class="badge">{len(closed)}</span> '
                f'<span class="badge">Win rate {win_rate:.0f}%</span></div>',
                unsafe_allow_html=True,
            )

            cards = []
            for c in sorted(closed, key=lambda x: x["BookedPnL"]):
                pnl = c["BookedPnL"]
                pnl_cls = "kpi-pos" if pnl >= 0 else "kpi-neg"
                cards.append(flat(f"""
                    <div class="closed-card">
                        <div class="closed-top">
                            <div class="closed-symbol">{c.get('Symbol', '-')}</div>
                            <span class="closed-badge">{c.get('Segment', '-')}</span>
                        </div>
                        <div class="closed-rows"><span>Exchange</span><span>{c.get('Exchange', '-')}</span></div>
                        <div class="closed-rows"><span>Qty</span><span>{fmt_qty(c.get('Qty'))}</span></div>
                        <div class="closed-rows"><span>Avg Buy</span><span>{fmt_money(c.get('AvgBuyPrice'))}</span></div>
                        <div class="closed-rows"><span>Avg Sell</span><span>{fmt_money(c.get('AvgSellPrice'))}</span></div>
                        <div class="closed-pnl {pnl_cls}">{fmt_money(pnl)}</div>
                    </div>
                """))
            st.markdown(f'<div class="closed-grid">{"".join(cards)}</div>', unsafe_allow_html=True)

            # ── Chart 1: cumulative booked profit over time (all segments
            # with a usable sell date; F&O sell dates aren't always present
            # in this ledger format, so that segment may be partial).
            legs = []
            for p in closed:
                for t in p.get("Trades", []):
                    if t.get("SellDate"):
                        legs.append({"SellDate": t["SellDate"], "Pnl": t["Pnl"], "Segment": p.get("Segment", "Other")})

            if legs:
                chart_df = pd.DataFrame(legs).sort_values("SellDate")
                chart_df["Cumulative"] = chart_df["Pnl"].cumsum()
                st.markdown('<div class="chart-card">', unsafe_allow_html=True)
                st.markdown('<div class="section-label">Cumulative booked profit over time</div>', unsafe_allow_html=True)
                area = alt.Chart(chart_df).mark_area(
                    line={"color": "#34d5c8", "strokeWidth": 2},
                    interpolate="monotone",
                    fillOpacity=0.15,
                    color=alt.Gradient(
                        gradient="linear",
                        stops=[alt.GradientStop(color="#34d5c8", offset=0), alt.GradientStop(color="transparent", offset=1)],
                        x1=1, x2=1, y1=1, y2=0,
                    ),
                ).encode(
                    x=alt.X("SellDate:T", title=None),
                    y=alt.Y("Cumulative:Q", title="Cumulative ₹"),
                    tooltip=[alt.Tooltip("SellDate:T", title="Date"), alt.Tooltip("Cumulative:Q", title="Cumulative", format=",.0f")],
                ).properties(height=240)
                st.altair_chart(alt_dark(area), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.caption("No sell-dated trade legs available yet to plot a cumulative profit trend.")

            # ── Chart 2: top gainers/losers by booked P&L ──────────────────
            top_df = pd.DataFrame(closed)[["Symbol", "BookedPnL"]].copy()
            top_df["AbsPnl"] = top_df["BookedPnL"].abs()
            top_df = top_df.sort_values("AbsPnl", ascending=False).head(12).drop(columns="AbsPnl")
            top_df["Direction"] = top_df["BookedPnL"].apply(lambda v: "Profit" if v >= 0 else "Loss")

            st.markdown('<div class="chart-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-label">Biggest movers — booked P&amp;L</div>', unsafe_allow_html=True)
            bar = alt.Chart(top_df).mark_bar(cornerRadiusEnd=4).encode(
                x=alt.X("BookedPnL:Q", title="Booked P&L (₹)"),
                y=alt.Y("Symbol:N", sort="-x", title=None),
                color=alt.Color(
                    "Direction:N",
                    scale=alt.Scale(domain=["Profit", "Loss"], range=["#2ee6a6", "#ff5c7a"]),
                    legend=None,
                ),
                tooltip=[alt.Tooltip("Symbol:N"), alt.Tooltip("BookedPnL:Q", title="Booked P&L", format=",.0f")],
            ).properties(height=max(220, 24 * len(top_df)))
            st.altair_chart(alt_dark(bar), use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # ── Chart 3: win rate donut ─────────────────────────────────────
            win_df = pd.DataFrame({"Outcome": ["Profitable", "Loss-making"], "Count": [wins, losses]})
            st.markdown('<div class="chart-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-label">Win / loss split</div>', unsafe_allow_html=True)
            donut = alt.Chart(win_df).mark_arc(innerRadius=65, cornerRadius=3).encode(
                theta=alt.Theta("Count:Q"),
                color=alt.Color(
                    "Outcome:N",
                    scale=alt.Scale(domain=["Profitable", "Loss-making"], range=["#2ee6a6", "#ff5c7a"]),
                    legend=alt.Legend(orient="right", title=None),
                ),
                tooltip=[alt.Tooltip("Outcome:N"), alt.Tooltip("Count:Q")],
            ).properties(height=220)
            st.altair_chart(alt_dark(donut), use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)


# ── UI ──────────────────────────────────────────────────────────────────
def main():
    if not check_password():
        return

    inject_theme()

    st.markdown(
        flat("""
        <div class="db-header">
            <div class="db-title">
                <div class="db-icon">📊</div>
                <h1>Booked Profit Dashboard</h1>
            </div>
        </div>
        """),
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.markdown("### Trade ledger")
        uploaded = st.file_uploader("Upload today's securities Excel", type=["xlsx"])
        default_path = "securities_today.xlsx"
        ledger_bytes = None
        if uploaded is not None:
            ledger_bytes = uploaded.getvalue()
        elif os.path.exists(default_path):
            with open(default_path, "rb") as f:
                ledger_bytes = f.read()
            st.caption("Using bundled `securities_today.xlsx` (no file uploaded this session).")
        st.divider()
        if st.button("🔄 Restart feed", use_container_width=True):
            st.cache_resource.clear()
            st.rerun()
        st.caption(f"Live tables refresh every {TICK_REFRESH_SECONDS}s, tick by tick.")

    if ledger_bytes is None:
        st.info("Upload your trade ledger Excel in the sidebar to start the live feed.")
        return

    engine = get_engine(ledger_bytes)
    render_live(engine)


if __name__ == "__main__":
    main()
