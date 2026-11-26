import os
import sqlite3
import openpyxl
from flask import Flask, render_template, jsonify, request
import yfinance as yf

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"))
DB_FILE = os.path.join(BASE_DIR, "watchlist.db")


def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS stocks (symbol TEXT PRIMARY KEY)")
    cur.execute("SELECT COUNT(*) FROM stocks")
    if cur.fetchone()[0] == 0:
        defaults = [("RELIANCE.NS",), ("TCS.NS",), ("HDFCBANK.NS",), ("INFY.NS",), ("ICICIBANK.NS",)]
        cur.executemany("INSERT INTO stocks VALUES (?)", defaults)
    conn.commit()
    conn.close()

def get_watchlist():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT symbol FROM stocks")
    symbols = [row[0] for row in cur.fetchall()]
    conn.close()
    return symbols


import time

CACHE = {"time": 0, "data": []}

def get_stock_data(symbols):
    if not symbols:
        return []
    if time.time() - CACHE["time"] < 30 and len(CACHE["data"]) == len(symbols):
        return CACHE["data"]

    try:
        df = yf.download(
            tickers=symbols,
            period="2d",
            interval="1d",
            group_by="ticker",
            threads=True,
            progress=False
        )
    except Exception:
        df = None

    results = []
    single = len(symbols) == 1

    for sym in symbols:
        try:
            sub = df if single else df[sym]
            sub = sub.dropna(subset=["Close"])
            if len(sub) == 0:
                raise ValueError("No data")

            latest = sub.iloc[-1]
            price = round(float(latest["Close"]), 2)
            day_high = round(float(latest["High"]), 2)
            day_low = round(float(latest["Low"]), 2)
            prev_close = round(float(sub.iloc[-2]["Close"]), 2) if len(sub) >= 2 else round(float(latest["Open"]), 2)

            change = round(price - prev_close, 2)
            pct_change = round((change / prev_close) * 100, 2) if prev_close else 0.0
            high_pct = round(((day_high - prev_close) / prev_close) * 100, 2) if prev_close else 0.0
            low_pct = round(((day_low - prev_close) / prev_close) * 100, 2) if prev_close else 0.0

            results.append({
                "symbol": sym.replace(".NS", "").replace(".BO", ""),
                "full_symbol": sym,
                "price": price,
                "day_high": day_high,
                "high_pct": high_pct,
                "day_low": day_low,
                "low_pct": low_pct,
                "change": change,
                "pct_change": pct_change
            })
        except Exception:
            results.append({
                "symbol": sym.replace(".NS", "").replace(".BO", ""),
                "full_symbol": sym,
                "price": "N/A",
                "day_high": "N/A",
                "high_pct": 0,
                "day_low": "N/A",
                "low_pct": 0,
                "change": 0,
                "pct_change": 0
            })

    CACHE["time"] = time.time()
    CACHE["data"] = results
    return results

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/prices")
def api_prices():
    return jsonify(get_stock_data(get_watchlist()))


@app.route("/api/add", methods=["POST"])
def add_stock():
    symbol = request.json.get("symbol", "").strip().upper()
    if symbol and not (symbol.endswith(".NS") or symbol.endswith(".BO")):
        symbol += ".NS"  # Default to NSE
    if symbol:
        conn = sqlite3.connect(DB_FILE)
        conn.execute("INSERT OR IGNORE INTO stocks VALUES (?)", (symbol,))
        conn.commit()
        conn.close()
    return jsonify({"status": "ok", "watchlist": get_watchlist()})

@app.route("/api/upload", methods=["POST"])
def upload_excel():
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "No file selected"}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"status": "error", "message": "Please upload a valid .xlsx Excel file"}), 400

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
    return jsonify({"status": "ok", "added": added_count})

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=False)


