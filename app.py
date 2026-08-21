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
        @import url('https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap');

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

        html, body, [class*="css"]  { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }

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
            position: relative;
            container-type: inline-size; /* lets kpi-value size off THIS card's actual width */
        }
        .kpi-card::before {
            content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px;
            background: linear-gradient(90deg, var(--accent), var(--accent-2)); opacity: 0.85;
            border-radius: 16px 16px 0 0; /* rounds the bar itself now that the card no longer clips overflow */
        }
        .kpi-label { font-size: 0.74rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; color: var(--muted); margin-bottom: 8px; white-space: nowrap; }
        /* cqw = % of this card's own width, so the value shrinks exactly as much as its
           card needs regardless of how many KPI cards share the row. No overflow/ellipsis
           here on purpose — a clipped digit on a money figure is worse than a smaller font. */
        .kpi-value { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: clamp(0.72rem, 8.5cqw, 1.55rem); font-weight: 700; letter-spacing: -0.01em; white-space: nowrap; display: block; }
        .kpi-pos { color: var(--pos); }
        .kpi-neg { color: var(--neg); }
        .kpi-sub { font-size: 0.72rem; color: var(--muted); margin-top: 6px; font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }

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
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.82rem; color: var(--text);
            display: flex; align-items: center; gap: 10px;
        }
        .top-clock .date-part { color: var(--muted); }
        .top-clock .tz-badge {
            font-size: 0.65rem; font-weight: 700; color: var(--accent);
            background: rgba(52,213,200,0.1); border: 1px solid rgba(52,213,200,0.25);
            padding: 1px 7px; border-radius: 999px;
        }

        /* Segment-scoped one-line summary (shown inside each open/closed
           segment tab — reflects ONLY that segment, not the whole book). */
        .seg-summary {
            display: flex; flex-wrap: wrap; align-items: center; gap: 22px;
            background: var(--panel-2); border: 1px solid var(--border); border-radius: 12px;
            padding: 10px 18px; margin-bottom: 16px;
        }
        .seg-stat { display: flex; align-items: center; gap: 8px; }
        .seg-stat-label {
            font-size: 0.68rem; font-weight: 600; text-transform: uppercase;
            letter-spacing: 0.05em; color: var(--muted);
        }
        .seg-stat-value {
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif;
            font-size: 0.88rem; font-weight: 700;
        }

        /* Position cards */
        .pos-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 8px; }
        @media (max-width: 640px) { .pos-grid { grid-template-columns: 1fr; } }
        .pos-card {
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border); border-radius: 14px; padding: 14px 16px;
            transition: border-color 0.2s ease;
        }
        .pos-card:hover { border-color: rgba(52,213,200,0.35); }

        /* Highest daily gain card — animated RGB glow so it pops out of the grid.
           Recomputed every refresh, so this always follows whichever position
           currently has the top daily gain rather than sitting on one symbol. */
        .pos-card.top-gain {
            position: relative;
            border-color: transparent;
            background:
                linear-gradient(180deg, var(--panel-2), var(--panel)) padding-box,
                conic-gradient(from var(--rgb-angle, 0deg), #ff3cac, #784ba0, #2b86c5, #2ee6a6, #f5b942, #ff3cac) border-box;
            border: 2px solid transparent;
            animation: rgb-spin 4s linear infinite;
            box-shadow: 0 0 22px rgba(124,92,255,0.35), 0 0 42px rgba(52,213,200,0.18);
        }
        @property --rgb-angle {
            syntax: '<angle>'; inherits: false; initial-value: 0deg;
        }
        @keyframes rgb-spin {
            to { --rgb-angle: 360deg; }
        }
        .top-gain-badge {
            position: absolute; top: -10px; right: 14px;
            font-size: 0.6rem; font-weight: 800; letter-spacing: 0.04em;
            padding: 2px 9px; border-radius: 999px; text-transform: uppercase;
            background: linear-gradient(90deg, #ff3cac, #784ba0, #2b86c5, #2ee6a6);
            background-size: 300% 100%; animation: rgb-shift 3s linear infinite;
            color: #05070d; box-shadow: 0 0 10px rgba(124,92,255,0.5);
        }
        @keyframes rgb-shift {
            0% { background-position: 0% 50%; }
            100% { background-position: 300% 50%; }
        }
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
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.72rem; font-weight: 700;
            padding: 3px 9px; border-radius: 999px; white-space: nowrap;
        }
        .day-pos { background: rgba(46,230,166,0.12); color: var(--pos); }
        .day-neg { background: rgba(255,92,122,0.12); color: var(--neg); }
        .pos-rows { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 14px; margin-bottom: 10px; }
        .pos-row-label { font-size: 0.68rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
        .pos-row-value { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.82rem; font-weight: 600; }
        .pos-mtm-block {
            display: flex; justify-content: space-between; align-items: center;
            border-top: 1px solid var(--border); padding-top: 10px;
        }
        .pos-mtm-label { font-size: 0.68rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
        .pos-mtm-value { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 1.02rem; font-weight: 700; }
        .pos-tick { font-size: 0.66rem; color: var(--muted); font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }

        /* Closed position cards */
        .closed-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 20px; }
        @media (max-width: 640px) { .closed-grid { grid-template-columns: 1fr; } }
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
        .closed-rows { display: flex; justify-content: space-between; font-size: 0.72rem; color: var(--muted); margin-bottom: 4px; font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }
        .closed-pnl { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-weight: 700; font-size: 0.95rem; margin-top: 8px; }

        /* Position table — borderless rows, just a thin separator line
           between each position, replacing the boxed-card layout. */
        .pos-table-wrap { overflow-x: auto; margin-bottom: 20px; }
        .pos-table { width: 100%; min-width: 760px; border-collapse: collapse; }
        .pos-table-head, .pos-table-row {
            display: grid;
            align-items: center;
            column-gap: 14px;
        }
        .pos-table-head.cols-open, .pos-table-row.cols-open {
            grid-template-columns: 1.5fr 0.8fr 0.8fr 1fr 1fr 1fr 1.2fr 1fr;
        }
        .pos-table-head.cols-closed, .pos-table-row.cols-closed {
            grid-template-columns: 1.5fr 0.8fr 0.8fr 1fr 1fr 1fr 1.2fr;
        }
        .pos-table-head {
            padding: 6px 6px 10px 6px;
            border-bottom: 1px solid var(--border);
        }
        .pos-table-head > div {
            font-size: 0.66rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.05em; color: var(--muted);
        }
        .pos-table-row {
            padding: 13px 6px;
            border-bottom: 1px solid var(--border);
            transition: background 0.15s ease;
        }
        .pos-table-row:hover { background: rgba(255,255,255,0.025); }
        .pos-table-row:last-child { border-bottom: none; }
        .pt-symbol { display: flex; align-items: baseline; gap: 8px; }
        .pt-symbol-name { font-weight: 700; font-size: 0.9rem; letter-spacing: -0.01em; }
        .pt-tag {
            font-size: 0.58rem; font-weight: 700; padding: 1px 6px; border-radius: 999px;
            border: 1px solid var(--border); color: var(--muted); text-transform: uppercase;
        }
        .pt-tag.long { color: var(--pos); border-color: rgba(46,230,166,0.3); }
        .pt-tag.short { color: var(--neg); border-color: rgba(255,92,122,0.3); }
        .pt-cell { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.85rem; font-weight: 600; }
        .pt-cell.muted { color: var(--muted); font-weight: 500; font-size: 0.78rem; }
        .pt-cell.pos { color: var(--pos); }
        .pt-cell.neg { color: var(--neg); }
        .pt-arrow { font-size: 0.72rem; margin-right: 2px; }

        /* Leader row (top gain / least loss) — no box, just a soft glowing
           left accent bar + tinted background so it stands out among plain
           rows without reintroducing card borders. */
        .pos-table-row.top-gain-row {
            position: relative;
            background: linear-gradient(90deg, rgba(124,92,255,0.10), transparent 60%);
            border-left: 3px solid;
            border-image: linear-gradient(180deg, #ff3cac, #784ba0, #2b86c5, #2ee6a6) 1;
        }
        .pt-leader-badge {
            display: inline-flex; align-items: center; gap: 4px; margin-left: 8px;
            font-size: 0.58rem; font-weight: 800; letter-spacing: 0.04em;
            padding: 1px 8px; border-radius: 999px; text-transform: uppercase;
            background: linear-gradient(90deg, #ff3cac, #784ba0, #2b86c5, #2ee6a6);
            background-size: 300% 100%; animation: rgb-shift 3s linear infinite;
            color: #05070d;
        }

        /* Chart cards: st.markdown('<div class="chart-card">') + st.altair_chart(...) +
           st.markdown('</div>') used to be 3 SEPARATE calls — Streamlit renders each call
           as its own DOM node, so that div opened and closed empty and the real chart sat
           outside it, unstyled. Charts now render inside `with st.container(border=True):`,
           which is a genuine parent element, and we restyle Streamlit's own wrapper for it
           below so it matches the rest of the theme instead of Streamlit's default grey box. */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: var(--panel-2) !important; border: 1px solid var(--border) !important;
            border-radius: 14px !important; margin-bottom: 18px !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] > div { border-radius: 14px !important; }

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


def seg_summary_line(items):
    """items: list of (label, value_html, extra_css_class) tuples — a single
    horizontal strip of stats scoped to whichever segment tab it's rendered
    inside, so it never mixes numbers across segments."""
    stats = "".join(
        f'<div class="seg-stat"><span class="seg-stat-label">{lbl}</span>'
        f'<span class="seg-stat-value {cls}">{val}</span></div>'
        for lbl, val, cls in items
    )
    return flat(f'<div class="seg-summary">{stats}</div>')


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


def fmt_datetime(ts):
    """ts is an ISO datetime string (possibly tz-aware); show date + time."""
    if not ts:
        return "-"
    try:
        dt = datetime.fromisoformat(ts)
        return dt.strftime("%d %b, %H:%M:%S")
    except ValueError:
        return ts


def fmt_sell_date(d):
    """d may be a date/datetime object or an ISO-ish string; show a short date."""
    if not d:
        return "-"
    if isinstance(d, str):
        try:
            d = datetime.fromisoformat(d)
        except ValueError:
            return d
    try:
        return d.strftime("%d %b %Y")
    except AttributeError:
        return str(d)


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
    other = [t for t in ticks if t.get("segment") not in ("Equity", "F&O")]

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
    other_buy, other_mtm, other_day = seg_totals(other)
    seg_open_totals = {
        "Equity": (eq_buy, eq_mtm, eq_day),
        "F&O": (fo_buy, fo_mtm, fo_day),
        "Other": (other_buy, other_mtm, other_day),
    }
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
        open_segments = [("Equity", equity), ("F&O", fo)]
        if other:
            open_segments.append(("Other", other))
        # One sub-tab per segment so picking "F&O" shows only F&O, not every
        # segment stacked one after another.
        open_seg_tabs = st.tabs([f"{label} ({len(rows)})" for label, rows in open_segments])
        for (label, rows), seg_tab in zip(open_segments, open_seg_tabs):
          with seg_tab:
            count_badge = f'<span class="badge">{len(rows)}</span>'
            st.markdown(f'<div class="section-label">{label} {count_badge}</div>', unsafe_allow_html=True)

            # One-line summary scoped to THIS segment only — switching to the
            # F&O tab shows F&O's own investment/running-P&L/today's-P&L,
            # not the combined book total.
            seg_buy, seg_mtm, seg_day = seg_open_totals.get(label, (0.0, 0.0, 0.0))
            seg_mtm_pct = (seg_mtm / seg_buy * 100) if seg_buy else 0.0
            seg_day_pct = (seg_day / seg_buy * 100) if seg_buy else 0.0
            st.markdown(
                seg_summary_line([
                    (f"{label} Investment", fmt_money(seg_buy), ""),
                    (
                        "Running P&L",
                        f"{fmt_money(seg_mtm)} ({'+' if seg_mtm >= 0 else ''}{seg_mtm_pct:.2f}%)",
                        "kpi-pos" if seg_mtm >= 0 else "kpi-neg",
                    ),
                    (
                        "Today's P&L",
                        f"{fmt_money(seg_day)} ({'+' if seg_day >= 0 else ''}{seg_day_pct:.2f}%)",
                        "kpi-pos" if seg_day >= 0 else "kpi-neg",
                    ),
                ]),
                unsafe_allow_html=True,
            )

            if not rows:
                st.caption("No open positions in this segment.")
                continue

            def _daily_gain_pct(x):
                # Price-move % (ltp vs prev close), then flipped for shorts:
                # a short position LOSES when the price rises, so its
                # "daily gain" is the mirror image of the raw price move.
                close = x.get("close")
                if not close:
                    return None
                raw_pct = (x["ltp"] - close) / close * 100
                pos_type = (x.get("positionType") or "LONG").upper()
                return raw_pct if pos_type == "LONG" else -raw_pct

            # Highest daily gain first, always — never a fixed/pinned order.
            # Positions with no price yet (None) sort to the bottom.
            rows_html = []
            sorted_rows = sorted(
                rows,
                key=lambda x: (_daily_gain_pct(x) is None, -(_daily_gain_pct(x) or 0)),
            )
            for idx, r in enumerate(sorted_rows):
                pos_type = (r.get("positionType") or "LONG").upper()
                type_cls = "long" if pos_type == "LONG" else "short"
                close = r.get("close")
                day_pct = _daily_gain_pct(r)  # this position's actual gain/loss %, short-adjusted
                day_cls = "pt-cell pos" if (day_pct or 0) >= 0 else "pt-cell neg"
                day_txt = f"{'+' if (day_pct or 0) >= 0 else ''}{day_pct:.2f}%" if day_pct is not None else "–"
                # The single best-performing row in this segment always gets
                # the glow — if it's a genuine gain it reads "Top Gain"; if
                # every position in the segment is red today, the least-bad
                # one still gets marked so there's always a clear leader.
                is_top = idx == 0 and day_pct is not None
                row_cls = "pos-table-row cols-open top-gain-row" if is_top else "pos-table-row cols-open"
                if is_top and day_pct > 0:
                    leader_badge = f'<span class="pt-leader-badge">🔥 Top Gain</span>'
                elif is_top:
                    leader_badge = f'<span class="pt-leader-badge">🛡️ Least Loss</span>'
                else:
                    leader_badge = ""
                mtm = r.get("mtm")
                mtm_cls = "pt-cell pos" if (mtm or 0) >= 0 else "pt-cell neg"
                mtm_arrow = "▲" if (mtm or 0) >= 0 else "▼"
                rows_html.append(flat(f"""
                    <div class="{row_cls}">
                        <div class="pt-symbol">
                            <span class="pt-symbol-name">{r.get('symbol', '-')}</span>
                            <span class="pt-tag {type_cls}">{pos_type}</span>
                            {leader_badge}
                        </div>
                        <div class="pt-cell muted">{r.get('exchange', '-')}</div>
                        <div class="pt-cell">{fmt_qty(r.get('qty'))}</div>
                        <div class="pt-cell">{fmt_money(r.get('avgPrice'))}</div>
                        <div class="pt-cell">{fmt_money(r.get('ltp'))}</div>
                        <div class="{day_cls}">{day_txt}</div>
                        <div class="{mtm_cls}">{mtm_arrow} {fmt_money(mtm)}</div>
                        <div class="pt-cell muted">{fmt_datetime(r.get('ts'))}</div>
                    </div>
                """))
            table_html = flat(f"""
                <div class="pos-table-wrap">
                    <div class="pos-table">
                        <div class="pos-table-head cols-open">
                            <div>Symbol</div><div>Exchange</div><div>Qty</div>
                            <div>Avg Price</div><div>CMP</div><div>Day Chg %</div>
                            <div>MTM P&amp;L</div><div>Last Tick</div>
                        </div>
                        {"".join(rows_html)}
                    </div>
                </div>
            """)
            st.markdown(table_html, unsafe_allow_html=True)

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

            def _closed_sell_date(c):
                dates = [t.get("SellDate") for t in c.get("Trades", []) if t.get("SellDate")]
                return max(dates) if dates else None

            def closed_row_html(c, is_top=False):
                pnl = c["BookedPnL"]
                pnl_cls = "pt-cell pos" if pnl >= 0 else "pt-cell neg"
                pnl_arrow = "▲" if pnl >= 0 else "▼"
                row_cls = "pos-table-row cols-closed top-gain-row" if is_top else "pos-table-row cols-closed"
                if is_top and pnl > 0:
                    leader_badge = '<span class="pt-leader-badge">🔥 Top Gain</span>'
                elif is_top:
                    leader_badge = '<span class="pt-leader-badge">🛡️ Least Loss</span>'
                else:
                    leader_badge = ""
                sell_date = fmt_sell_date(_closed_sell_date(c))
                return flat(f"""
                    <div class="{row_cls}">
                        <div class="pt-symbol">
                            <span class="pt-symbol-name">{c.get('Symbol', '-')}</span>
                            {leader_badge}
                        </div>
                        <div class="pt-cell muted">{c.get('Exchange', '-')}</div>
                        <div class="pt-cell">{fmt_qty(c.get('Qty'))}</div>
                        <div class="pt-cell">{fmt_money(c.get('AvgBuyPrice'))}</div>
                        <div class="pt-cell">{fmt_money(c.get('AvgSellPrice'))}</div>
                        <div class="pt-cell muted">{sell_date}</div>
                        <div class="{pnl_cls}">{pnl_arrow} {fmt_money(pnl)}</div>
                    </div>
                """)

            closed_equity = [c for c in closed if c.get("Segment") == "Equity"]
            closed_fo = [c for c in closed if c.get("Segment") == "F&O"]
            closed_segments = [("Equity", closed_equity), ("F&O", closed_fo)]
            closed_other = [c for c in closed if c.get("Segment") not in ("Equity", "F&O")]
            if closed_other:
                closed_segments.append(("Other", closed_other))

            # One sub-tab per segment so picking "F&O" shows only F&O, not every
            # segment stacked one after another.
            closed_seg_tabs = st.tabs([f"{label} ({len(rows)})" for label, rows in closed_segments])
            for (label, rows), seg_tab in zip(closed_segments, closed_seg_tabs):
              with seg_tab:
                if not rows:
                    st.markdown(f'<div class="section-label">{label} <span class="badge">0</span></div>', unsafe_allow_html=True)
                    st.caption(f"No closed {label.lower()} positions in this ledger.")
                    continue
                seg_win_rate = sum(1 for c in rows if c["BookedPnL"] >= 0) / len(rows) * 100
                st.markdown(
                    f'<div class="section-label">{label} <span class="badge">{len(rows)}</span></div>',
                    unsafe_allow_html=True,
                )
                # One-line summary scoped to THIS segment's closed positions
                # only — cost basis deployed here vs. what was booked here.
                seg_invested = sum(abs(c.get("Qty") or 0) * (c.get("AvgBuyPrice") or 0) for c in rows)
                seg_booked = sum(c["BookedPnL"] for c in rows)
                seg_booked_pct = (seg_booked / seg_invested * 100) if seg_invested else 0.0
                st.markdown(
                    seg_summary_line([
                        (f"{label} Investment", fmt_money(seg_invested), ""),
                        (
                            "Booked P&L",
                            f"{fmt_money(seg_booked)} ({'+' if seg_booked >= 0 else ''}{seg_booked_pct:.2f}%)",
                            "kpi-pos" if seg_booked >= 0 else "kpi-neg",
                        ),
                        ("Win rate", f"{seg_win_rate:.0f}%", ""),
                    ]),
                    unsafe_allow_html=True,
                )
                # Whichever row has the single highest booked P&L in this
                # segment always gets the glow — a genuine gain reads "Top
                # Gain", otherwise the least-bad loss reads "Least Loss".
                sorted_rows = sorted(rows, key=lambda x: x["BookedPnL"], reverse=True)
                best_pnl = max((c["BookedPnL"] for c in rows), default=None)
                rows_html = [
                    closed_row_html(
                        c,
                        is_top=(best_pnl is not None and c["BookedPnL"] == best_pnl),
                    )
                    for c in sorted_rows
                ]
                table_html = flat(f"""
                    <div class="pos-table-wrap">
                        <div class="pos-table">
                            <div class="pos-table-head cols-closed">
                                <div>Symbol</div><div>Exchange</div><div>Qty</div>
                                <div>Avg Buy</div><div>Avg Sell</div><div>Sell Date</div>
                                <div>Booked P&amp;L</div>
                            </div>
                            {"".join(rows_html)}
                        </div>
                    </div>
                """)
                st.markdown(table_html, unsafe_allow_html=True)

            # ── Charts (combined across segments) — rendered below every
            # position card section so cards stay the primary focus. Chart 1
            # stays full-width since a date axis needs the room; charts 2
            # and 3 are compact enough to share a row.
            # Chart 1: cumulative booked profit over time (all segments with
            # a usable sell date; F&O sell dates aren't always present in
            # this ledger format, so that segment may be partial).
            legs = []
            for p in closed:
                for t in p.get("Trades", []):
                    if t.get("SellDate"):
                        legs.append({"SellDate": t["SellDate"], "Pnl": t["Pnl"], "Segment": p.get("Segment", "Other")})

            if legs:
                chart_df = pd.DataFrame(legs).sort_values("SellDate")
                chart_df["Cumulative"] = chart_df["Pnl"].cumsum()
                with st.container(border=True):
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
            else:
                st.caption("No sell-dated trade legs available yet to plot a cumulative profit trend.")

            # Charts 2 + 3 share one row.
            top_df = pd.DataFrame(closed)[["Symbol", "BookedPnL"]].copy()
            top_df["AbsPnl"] = top_df["BookedPnL"].abs()
            top_df = top_df.sort_values("AbsPnl", ascending=False).head(12).drop(columns="AbsPnl")
            top_df["Direction"] = top_df["BookedPnL"].apply(lambda v: "Profit" if v >= 0 else "Loss")
            win_df = pd.DataFrame({"Outcome": ["Profitable", "Loss-making"], "Count": [wins, losses]})

            col_movers, col_donut = st.columns(2)
            with col_movers:
                with st.container(border=True):
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

            with col_donut:
                with st.container(border=True):
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
