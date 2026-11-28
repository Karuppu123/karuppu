import os
import time
import logging
import sqlite3
import threading
import openpyxl
import pandas as pd
from flask import Flask, render_template, jsonify, request
import yfinance as yf

logging.getLogger("yfinance").setLevel(logging.ERROR)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"))
DB_FILE = os.path.join(BASE_DIR, "watchlist.db")

CACHE = {"time": 0, "data": []}
FETCH_LOCK = threading.Lock()


def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS stocks (symbol TEXT PRIMARY KEY)")
    cur.execute("SELECT COUNT(*) FROM stocks")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO stocks VALUES (?)", [("RELIANCE.NS",)])
    conn.commit()
    conn.close()


def get_watchlist():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT symbol FROM stocks")
    symbols = [row[0] for row in cur.fetchall()]
    conn.close()
    return symbols


def _na_entry(sym):
    return {
        "symbol": sym.replace(".NS", "").replace(".BO", ""),
        "full_symbol": sym,
        "price": "N/A",
        "day_high": "N/A",
        "high_pct": 0,
        "day_low": "N/A",
        "low_pct": 0,
        "change": 0,
        "pct_change": 0,
    }


def _parse_history(sym, sub):
    sub = sub.dropna(subset=["Close"])
    if len(sub) == 0:
        return None

    latest = sub.iloc[-1]
    price = round(float(latest["Close"]), 2)
    day_high = round(float(latest["High"]), 2)
    day_low = round(float(latest["Low"]), 2)
    prev_close = (
        round(float(sub.iloc[-2]["Close"]), 2)
        if len(sub) >= 2
        else round(float(latest["Open"]), 2)
    )

    change = round(price - prev_close, 2)
    pct_change = round((change / prev_close) * 100, 2) if prev_close else 0.0
    high_pct = (
        round(((day_high - prev_close) / prev_close) * 100, 2)
        if prev_close
        else 0.0
    )
    low_pct = (
        round(((day_low - prev_close) / prev_close) * 100, 2)
        if prev_close
        else 0.0
    )

    return {
        "symbol": sym.replace(".NS", "").replace(".BO", ""),
        "full_symbol": sym,
        "price": price,
        "day_high": day_high,
        "high_pct": high_pct,
        "day_low": day_low,
        "low_pct": low_pct,
        "change": change,
        "pct_change": pct_change,
    }


def get_stock_data(symbols, force=False):
    if not symbols:
        return []

    with FETCH_LOCK:
        if not force and time.time() - CACHE["time"] < 60 and len(CACHE["data"]) == len(symbols):
            return CACHE["data"]

        data_map = {}
        chunk_size = 20

        for i in range(0, len(symbols), chunk_size):
            batch = symbols[i : i + chunk_size]
            try:
                df = yf.download(
                    tickers=batch,
                    period="5d",
                    interval="1d",
                    group_by="ticker",
                    threads=False,
                    progress=False,
                )
            except Exception as e:
                print(f"Batch download error: {e}", flush=True)
                df = None

            if df is not None and not df.empty:
                for sym in batch:
                    try:
                        if isinstance(df.columns, pd.MultiIndex):
                            if sym in df.columns.get_level_values(0):
                                sub = df[sym]
                            elif sym in df.columns.get_level_values(1):
                                sub = df.xs(sym, axis=1, level=1)
                            else:
                                continue
                        else:
                            sub = df

                        parsed = _parse_history(sym, sub)
                        if parsed:
                            data_map[sym] = parsed
                    except Exception as e:
                        print(f"Parse error for {sym}: {e}", flush=True)

            for sym in batch:
                if sym not in data_map:
                    try:
                        hist = yf.Ticker(sym).history(period="5d", interval="1d")
                        parsed = _parse_history(sym, hist)
                        if parsed:
                            data_map[sym] = parsed
                    except Exception as e:
                        print(f"Single ticker error for {sym}: {e}", flush=True)

        results = [data_map.get(sym, _na_entry(sym)) for sym in symbols]
        CACHE["time"] = time.time()
        CACHE["data"] = results
        return results


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/prices")
def api_prices():
    force = request.args.get("force") == "1"
    return jsonify(get_stock_data(get_watchlist(), force=force))


@app.route("/api/remove", methods=["POST"])
def remove_stocks():
    symbols = request.json.get("symbols", [])
    if symbols:
        conn = sqlite3.connect(DB_FILE)
        conn.executemany("DELETE FROM stocks WHERE symbol=?", [(s,) for s in symbols])
        conn.commit()
        conn.close()
        CACHE["time"] = 0
    return jsonify({"status": "ok", "removed": len(symbols)})


@app.route("/api/add", methods=["POST"])
def add_stock():
    symbol = request.json.get("symbol", "").strip().upper()
    if symbol and not (symbol.endswith(".NS") or symbol.endswith(".BO")):
        symbol += ".NS"
    if symbol:
        conn = sqlite3.connect(DB_FILE)
        conn.execute("INSERT OR IGNORE INTO stocks VALUES (?)", (symbol,))
        conn.commit()
        conn.close()
        CACHE["time"] = 0
    return jsonify({"status": "ok", "watchlist": get_watchlist()})


@app.route("/api/upload", methods=["POST"])
def upload_excel():
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "No file selected"}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return (
            jsonify({"status": "error", "message": "Please upload a .xlsx Excel file"}),
            400,
        )

    wb = openpyxl.load_workbook(file, data_only=True)
    sheet = wb.active

    conn = sqlite3.connect(DB_FILE)
    added_count = 0
    for row in sheet.iter_rows(values_only=True):
        if not row or not row[0]:
            continue
        sym = str(row[0]).strip().upper()
        if sym in ("SYMBOL", "SYMBOLS", "STOCK", "TICKER", "NAME"):
            continue
        if not (sym.endswith(".NS") or sym.endswith(".BO")):
            sym += ".NS"
        conn.execute("INSERT OR IGNORE INTO stocks VALUES (?)", (sym,))
        added_count += 1

    conn.commit()
    conn.close()
    CACHE["time"] = 0
    return jsonify({"status": "ok", "added": added_count})


@app.route("/cleanup")
def cleanup_invalid():
    data = get_stock_data(get_watchlist())
    invalid = [item["full_symbol"] for item in data if item["price"] == "N/A"]
    if invalid:
        conn = sqlite3.connect(DB_FILE)
        conn.executemany("DELETE FROM stocks WHERE symbol=?", [(s,) for s in invalid])
        conn.commit()
        conn.close()
        CACHE["time"] = 0
    return jsonify({"removed": invalid, "remaining": len(data) - len(invalid)})


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
    
