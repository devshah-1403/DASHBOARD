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
import threading
import time
from datetime import datetime

import pandas as pd
import pyotp
import streamlit as st
from SmartApi import SmartConnect
from SmartApi.smartWebSocketV2 import SmartWebSocketV2

from positions_builder import load_trade_ledger, build_positions
from token_resolver import resolve_all

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
        .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 26px; }
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

        /* Dataframe polish */
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border); border-radius: 12px; overflow: hidden;
        }

        div[data-testid="stMetric"] { display: none; }  /* using custom KPI cards instead */

        .ticker-strip {
            font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: var(--muted);
            margin-bottom: 4px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def status_pill(state: str, label: str) -> str:
    cls = {"live": "status-live", "error": "status-error", "warn": "status-warn"}.get(state, "")
    return f'<span class="status-pill {cls}"><span class="status-dot"></span>{label}</span>'


def kpi_card(label, value, positive=None, sub=None):
    cls = "kpi-pos" if positive is True else ("kpi-neg" if positive is False else "")
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value {cls}">{value}</div>
            {sub_html}
        </div>
    """


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
        """
        <div class="db-title" style="justify-content:center; margin: 60px 0 24px 0;">
            <div class="db-icon">📊</div>
            <h1>Booked Profit Dashboard</h1>
        </div>
        """,
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
        self._start(ledger_rows)

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
                now = datetime.now()
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


@st.cache_resource(show_spinner="Logging into Angel One and starting the live feed...")
def get_engine(ledger_bytes: bytes):
    rows = load_trade_ledger(io.BytesIO(ledger_bytes))
    return LiveEngine(rows)


# ── Formatting helpers ─────────────────────────────────────────────────────
def fmt_money(x):
    if x is None:
        return "-"
    return f"(₹{abs(x):,.2f})" if x < 0 else f"₹{x:,.2f}"


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
        return

    ticks, last_tick_ts = engine.snapshot()

    if engine.status == "disconnected":
        st.markdown(status_pill("error", "Disconnected — restart feed in sidebar"), unsafe_allow_html=True)
    else:
        age = f"{(datetime.now() - last_tick_ts).seconds}s ago" if last_tick_ts else "waiting for first tick..."
        st.markdown(status_pill("live", f"Live • last tick {age}"), unsafe_allow_html=True)

    equity = [t for t in ticks if t.get("segment") == "Equity"]
    fo = [t for t in ticks if t.get("segment") == "F&O"]

    def seg_totals(rows):
        buy_value = sum((r["qty"] or 0) * (r["avgPrice"] or 0) for r in rows)
        mtm = sum(r["mtm"] for r in rows if r["mtm"] is not None)
        return buy_value, mtm

    eq_buy, eq_mtm = seg_totals(equity)
    fo_buy, fo_mtm = seg_totals(fo)
    investment_value = eq_buy + fo_buy
    current_mtm = eq_mtm + fo_mtm
    total_mtm = engine.booked_mtm_total + current_mtm

    st.markdown(
        f"""
        <div class="kpi-grid">
            {kpi_card("Investment Value", fmt_money(investment_value), sub=f"{len(ticks)} instruments tracked")}
            {kpi_card("Booked MTM", fmt_money(engine.booked_mtm_total), positive=engine.booked_mtm_total >= 0, sub=f"{len(engine.closed_positions)} closed positions")}
            {kpi_card("Current (Open) MTM", fmt_money(current_mtm), positive=current_mtm >= 0, sub=f"Eq {fmt_money(eq_mtm)} · F&O {fmt_money(fo_mtm)}")}
            {kpi_card("Total MTM", fmt_money(total_mtm), positive=total_mtm >= 0, sub="Booked + Open, live")}
        </div>
        """,
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
            df = pd.DataFrame(rows)[["symbol", "exchange", "qty", "avgPrice", "ltp", "close", "mtm", "ts"]]
            df.columns = ["Symbol", "Exchange", "Qty", "Avg Price", "CMP", "Prev Close", "MTM", "Last Tick"]
            st.dataframe(
                style_pnl_table(df, ["MTM"]),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Avg Price": st.column_config.NumberColumn(format="₹%.2f"),
                    "CMP": st.column_config.NumberColumn(format="₹%.2f"),
                    "Prev Close": st.column_config.NumberColumn(format="₹%.2f"),
                },
            )

    with tab_closed:
        if not engine.closed_positions:
            st.caption("No closed positions in this ledger.")
        else:
            df = pd.DataFrame(engine.closed_positions)[
                ["Symbol", "Exchange", "Segment", "Qty", "AvgBuyPrice", "AvgSellPrice", "BookedPnL"]
            ]
            st.markdown(f'<div class="section-label">All closed positions <span class="badge">{len(df)}</span></div>', unsafe_allow_html=True)
            st.dataframe(
                style_pnl_table(df, ["BookedPnL"]),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "AvgBuyPrice": st.column_config.NumberColumn(format="₹%.2f"),
                    "AvgSellPrice": st.column_config.NumberColumn(format="₹%.2f"),
                },
            )

            # Cumulative booked profit chart — Equity only, since F&O sell
            # dates aren't reliably present in this ledger (see project notes).
            legs = []
            for p in engine.closed_positions:
                if p.get("Segment") != "Equity":
                    continue
                for t in p.get("Trades", []):
                    if t.get("SellDate"):
                        legs.append({"SellDate": t["SellDate"], "Pnl": t["Pnl"]})
            if legs:
                chart_df = pd.DataFrame(legs).sort_values("SellDate")
                chart_df["Cumulative"] = chart_df["Pnl"].cumsum()
                chart_df = chart_df.set_index("SellDate")
                st.markdown('<div class="section-label">Cumulative booked profit — Equity</div>', unsafe_allow_html=True)
                st.line_chart(chart_df["Cumulative"], color="#34d5c8")


# ── UI ──────────────────────────────────────────────────────────────────
def main():
    if not check_password():
        return

    inject_theme()

    st.markdown(
        """
        <div class="db-header">
            <div class="db-title">
                <div class="db-icon">📊</div>
                <h1>Booked Profit Dashboard</h1>
            </div>
        </div>
        """,
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
