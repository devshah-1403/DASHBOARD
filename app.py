"""
Booked Profit Dashboard — Streamlit version.

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
"""
import io
import os
import threading
import time
from datetime import datetime

import pandas as pd
import pyotp
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from SmartApi import SmartConnect
from SmartApi.smartWebSocketV2 import SmartWebSocketV2

from positions_builder import load_trade_ledger, build_positions
from token_resolver import resolve_all

st.set_page_config(page_title="Booked Profit Dashboard", layout="wide")

MASTER_CACHE_PATH = "instrument_master_cache.json"
MASTER_CACHE_MAX_AGE_HOURS = 20
SUBSCRIBE_MODE = 2
MAX_TOKENS_PER_SUBSCRIBE = 1000


# ── Password gate ─────────────────────────────────────────────────────────
def check_password():
    def _submit():
        st.session_state["_pw_ok"] = (
            st.session_state.get("_pw_input", "") == st.secrets.get("APP_PASSWORD", "")
        )

    if st.session_state.get("_pw_ok"):
        return True

    st.title("Booked Profit Dashboard")
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
                tick = {
                    "token": token, "symbol": meta.get("symbol", token),
                    "exchange": meta.get("exchange", ""), "segment": meta.get("segment", "Other"),
                    "positionType": meta.get("positionType", "LONG"), "ltp": ltp,
                    "close": (message.get("closed_price", 0) / 100.0 if message.get("closed_price") else prev.get("close")),
                    "open": (message.get("open_price_of_the_day", 0) / 100.0 if message.get("open_price_of_the_day") else prev.get("open")),
                    "volume": message.get("volume_trade_for_the_day"), "qty": qty, "avgPrice": avg_price,
                    "mtm": round((ltp - avg_price) * qty, 2) if (qty is not None and avg_price is not None) else None,
                    "ts": datetime.now().isoformat(timespec="seconds"),
                }
                self.latest_prices[token] = tick

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
            return list(self.latest_prices.values())


@st.cache_resource(show_spinner="Logging into Angel One and starting the live feed...")
def get_engine(ledger_bytes: bytes):
    rows = load_trade_ledger(io.BytesIO(ledger_bytes))
    return LiveEngine(rows)


# ── UI ──────────────────────────────────────────────────────────────────
def fmt_money(x):
    if x is None:
        return "-"
    return f"(₹{abs(x):,.2f})" if x < 0 else f"₹{x:,.2f}"


def main():
    if not check_password():
        return

    st.title("📊 Booked Profit Dashboard")
    st_autorefresh(interval=5000, key="live_refresh")

    with st.sidebar:
        st.subheader("Trade ledger")
        uploaded = st.file_uploader("Upload today's securities Excel", type=["xlsx"])
        default_path = "securities_today.xlsx"
        ledger_bytes = None
        if uploaded is not None:
            ledger_bytes = uploaded.getvalue()
        elif os.path.exists(default_path):
            with open(default_path, "rb") as f:
                ledger_bytes = f.read()
            st.caption("Using bundled securities_today.xlsx (no file uploaded this session).")
        if st.button("Restart feed (e.g. after uploading a new ledger)"):
            st.cache_resource.clear()
            st.rerun()

    if ledger_bytes is None:
        st.info("Upload your trade ledger Excel in the sidebar to start the live feed.")
        return

    engine = get_engine(ledger_bytes)

    if engine.status == "error":
        st.error(f"Feed failed to start: {engine.error}")
        return
    elif engine.status == "starting":
        st.warning("Starting up...")
        return
    elif engine.status == "disconnected":
        st.warning("Live feed disconnected — click 'Restart feed' in the sidebar.")

    ticks = engine.snapshot()
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

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Investment Value", fmt_money(investment_value))
    c2.metric("Booked MTM", fmt_money(engine.booked_mtm_total))
    c3.metric("Current (Open) MTM", fmt_money(current_mtm))
    c4.metric("Total MTM", fmt_money(total_mtm))

    tab_open, tab_closed = st.tabs(["Open positions", "Closed positions"])

    with tab_open:
        for label, rows in [("Equity", equity), ("F&O", fo)]:
            st.subheader(label)
            if not rows:
                st.caption("No open positions in this segment.")
                continue
            df = pd.DataFrame(rows)[["symbol", "exchange", "qty", "avgPrice", "ltp", "close", "mtm", "ts"]]
            df.columns = ["Symbol", "Exchange", "Qty", "Avg Price", "CMP", "Prev Close", "MTM", "Last Tick"]
            st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_closed:
        if not engine.closed_positions:
            st.caption("No closed positions in this ledger.")
        else:
            df = pd.DataFrame(engine.closed_positions)[
                ["Symbol", "Exchange", "Segment", "Qty", "AvgBuyPrice", "AvgSellPrice", "BookedPnL"]
            ]
            st.dataframe(df, use_container_width=True, hide_index=True)

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
                st.subheader("Cumulative booked profit — Equity")
                st.line_chart(chart_df["Cumulative"])


if __name__ == "__main__":
    main()
