"""
TikTok Dashboard — Flask backend
Accepts CSV uploads, returns chart-ready JSON.
"""

import io
import json
import os
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd

try:
    from flask import Flask, jsonify, render_template, request
except ModuleNotFoundError:
    Flask = None
    request = None

    def jsonify(payload):
        return payload

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input_data"
WORKSPACE_DIR = BASE_DIR.parent


class MiniApp:
    def route(self, *_args, **_kwargs):
        def decorator(func):
            return func
        return decorator

    def run(self, debug=False, host="127.0.0.1", port=5050):
        run_simple_server(host, port)


app = Flask(__name__) if Flask else MiniApp()
if Flask:
    app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024


# ---------------------------------------------------------------------------
# Cleaning helpers (same logic as merge_pipeline.py)
# ---------------------------------------------------------------------------

def clean_currency(s):
    return (
        s.astype(str)
        .str.replace(r"[\$,\s]", "", regex=True)
        .pipe(pd.to_numeric, errors="coerce")
    )

def clean_numeric(s):
    return (
        s.astype(str)
        .str.replace(r"[,\s]", "", regex=True)
        .pipe(pd.to_numeric, errors="coerce")
    )

def clean_pct(s):
    return (
        s.astype(str)
        .str.replace(r"[%\s]", "", regex=True)
        .pipe(pd.to_numeric, errors="coerce")
    )

def strip_cols(df):
    df.columns = df.columns.str.strip()
    return df

def pick(df, *candidates):
    """Return the first matching column name (case-insensitive)."""
    mapping = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in mapping:
            return mapping[cand.lower()]
    return None


def _source_name(source):
    if isinstance(source, (str, Path)):
        return str(source)
    return getattr(source, "filename", "")


def _source_bytes(source):
    if isinstance(source, (str, Path)):
        return source
    if hasattr(source, "file"):
        return source.file.read()
    return source.read()


def _as_readable(payload):
    if isinstance(payload, bytes):
        return io.BytesIO(payload)
    return payload


def _dedupe_columns(columns):
    seen = {}
    out = []
    for col in columns:
        name = str(col).strip() if pd.notna(col) else ""
        if not name:
            name = "Unnamed"
        count = seen.get(name, 0)
        seen[name] = count + 1
        out.append(name if count == 0 else f"{name}.{count}")
    return out


def _find_header_row(raw, keywords):
    if not keywords:
        return 0

    wanted = [k.lower() for k in keywords]
    best_idx = 0
    best_score = -1
    for idx, row in raw.head(40).iterrows():
        cells = [str(v).strip().lower() for v in row.tolist() if pd.notna(v) and str(v).strip()]
        score = sum(any(k == c or k in c for c in cells) for k in wanted)
        if score > best_score:
            best_idx = idx
            best_score = score
    return best_idx


def read_table(source, sheet_name=None, header_keywords=()):
    """Read CSV/XLSX exports, including TikTok sheets with metadata rows above headers."""
    name = _source_name(source).lower()
    payload = _source_bytes(source)

    if name.endswith((".xlsx", ".xls")):
        readable = _as_readable(payload)
        xl = pd.ExcelFile(readable)
        selected_sheet = sheet_name if sheet_name in xl.sheet_names else None
        if selected_sheet is None:
            for candidate in ("Product", "Trend", "Sheet1", "Sheet 1", xl.sheet_names[0]):
                if candidate in xl.sheet_names:
                    selected_sheet = candidate
                    break
        raw = pd.read_excel(_as_readable(payload), sheet_name=selected_sheet, header=None, dtype=str)
        header_idx = _find_header_row(raw, header_keywords)
        df = raw.iloc[header_idx + 1:].copy()
        df.columns = _dedupe_columns(raw.iloc[header_idx].tolist())
    else:
        df = pd.read_csv(_as_readable(payload), dtype=str)

    df = df.dropna(how="all")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed(\\.\\d+)?$")]
    return strip_cols(df)


def parse_dates(series):
    values = series.astype(str).str.strip()
    slash_dates = values.str.extract(r"^(\d{1,2})/(\d{1,2})/\d{4}")

    dayfirst = False
    if slash_dates.notna().any().any():
        first = pd.to_numeric(slash_dates[0], errors="coerce")
        second = pd.to_numeric(slash_dates[1], errors="coerce")
        if (first > 12).any():
            dayfirst = True
        elif (second > 12).any():
            dayfirst = False

    return pd.to_datetime(values, errors="coerce", dayfirst=dayfirst)


def add_period_columns(df, date_col):
    df[date_col] = parse_dates(df[date_col])
    df = df.dropna(subset=[date_col]).copy()
    df["_date"] = df[date_col].dt.normalize()
    df["_week"] = df[date_col].dt.to_period("W").apply(lambda p: str(p.start_time.date()))
    df["_month"] = df[date_col].dt.to_period("M").apply(lambda p: p.strftime("%b %Y"))
    return df


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_live_product(f):
    df = read_table(f, sheet_name="Product", header_keywords=("Product ID", "Product name", "Attributed GMV"))
    for col in df.columns:
        cl = col.lower()
        if "gmv" in cl or "revenue" in cl:
            df[col] = clean_currency(df[col])
        elif any(k in cl for k in ["sold", "order", "customer", "impression"]):
            df[col] = clean_numeric(df[col])
        elif "ctr" in cl:
            df[col] = clean_pct(df[col])
    return df


def load_shop_product(f):
    df = read_table(f, header_keywords=("Product ID", "Product Name", "GMV", "Orders"))
    for col in df.columns:
        cl = col.lower()
        if "gmv" in cl:
            df[col] = clean_currency(df[col])
        elif any(k in cl for k in ["sold", "order", "impression"]):
            df[col] = clean_numeric(df[col])
        elif "ctr" in cl:
            df[col] = clean_pct(df[col])
    return df


def load_live_perf(f):
    df = read_table(f, header_keywords=("Livestream", "Start time", "Attributed GMV", "Views"))
    if "Start time" in df.columns:
        df = add_period_columns(df, "Start time")
    for col in df.columns:
        cl = col.lower()
        if "gmv" in cl:
            df[col] = clean_currency(df[col])
        elif any(k in cl for k in ["order", "view", "follower"]):
            df[col] = clean_numeric(df[col])
    return df


def load_shop_daily(f):
    df = read_table(f, sheet_name="Sheet1", header_keywords=("Date", "GMV", "LIVE-attributed GMV", "Orders"))
    date_col = pick(df, "Date", "date")
    if date_col:
        df = add_period_columns(df, date_col)
    for col in df.columns:
        cl = col.lower()
        if "gmv" in cl:
            df[col] = clean_currency(df[col])
        elif any(k in cl for k in ["order", "sold", "impression"]):
            df[col] = clean_numeric(df[col])
    return df


def load_live_trend_as_session(path, session_date, session_name):
    df = read_table(path, header_keywords=("Time", "Attributed GMV", "Views", "New followers"))
    for col in df.columns:
        cl = col.lower()
        if "gmv" in cl:
            df[col] = clean_currency(df[col])
        elif any(k in cl for k in ["order", "view", "follower", "sold"]):
            df[col] = clean_numeric(df[col])

    def total(*candidates):
        col = pick(df, *candidates)
        return df[col].sum() if col else 0

    session = pd.DataFrame([{
        "Livestream": session_name,
        "Start time": session_date,
        "Duration": f"{len(df)} intervals",
        "Attributed GMV": total("Attributed GMV", "GMV"),
        "Attributed items sold": total("Attributed items sold", "Items sold"),
        "Attributed orders": total("Attributed orders", "Orders"),
        "Views": total("Views"),
        "New followers": total("New followers", "Followers"),
    }])
    return add_period_columns(session, "Start time")


# ---------------------------------------------------------------------------
# Chart builders
# ---------------------------------------------------------------------------

def _sort_labels(labels, period):
    """Sort period labels chronologically (not alphabetically)."""
    fmt = "%b %Y" if period == "_month" else "%Y-%m-%d"
    try:
        return sorted(labels, key=lambda l: pd.to_datetime(l, format=fmt))
    except Exception:
        return sorted(labels)


def chart_gmv_over_time(live_perf, shop_daily):
    """Bar chart: Live GMV vs Shop Total GMV grouped by week or month."""
    results = {}
    for period in ("_week", "_month"):
        rows = {}

        if shop_daily is not None and period in shop_daily.columns:
            gmv_col = pick(shop_daily, "GMV")
            live_attr_col = pick(shop_daily, "LIVE-attributed GMV", "LIVE GMV")
            if gmv_col:
                for p, grp in shop_daily.groupby(period):
                    rows.setdefault(p, {})["shop_total"] = round(grp[gmv_col].sum(), 2)
            if live_attr_col:
                for p, grp in shop_daily.groupby(period):
                    rows.setdefault(p, {})["shop_live_attr"] = round(grp[live_attr_col].sum(), 2)

        if live_perf is not None and period in live_perf.columns:
            gmv_col = pick(live_perf, "Attributed GMV", "GMV")
            if gmv_col:
                for p, grp in live_perf.groupby(period):
                    rows.setdefault(p, {})["live_gmv"] = round(grp[gmv_col].sum(), 2)

        labels = _sort_labels(list(rows.keys()), period)
        results[period.lstrip("_")] = {
            "labels": labels,
            "shop_total":    [rows[l].get("shop_total", 0) for l in labels],
            "shop_live_attr":[rows[l].get("shop_live_attr", 0) for l in labels],
            "live_gmv":      [rows[l].get("live_gmv", 0) for l in labels],
        }
    return results


def chart_live_pct(live_perf, shop_daily):
    """Line chart: Live GMV as % of shop total over time."""
    results = {}
    for period in ("_week", "_month"):
        rows = {}

        if shop_daily is not None and period in shop_daily.columns:
            gmv_col = pick(shop_daily, "GMV")
            if gmv_col:
                for p, grp in shop_daily.groupby(period):
                    rows.setdefault(p, {})["shop_total"] = grp[gmv_col].sum()

        if live_perf is not None and period in live_perf.columns:
            gmv_col = pick(live_perf, "Attributed GMV", "GMV")
            if gmv_col:
                for p, grp in live_perf.groupby(period):
                    rows.setdefault(p, {})["live_gmv"] = grp[gmv_col].sum()

        labels = _sort_labels(list(rows.keys()), period)
        pcts = []
        for l in labels:
            shop = rows[l].get("shop_total", 0)
            live = rows[l].get("live_gmv", 0)
            pcts.append(round(live / shop * 100, 2) if shop else 0)

        results[period.lstrip("_")] = {"labels": labels, "pct": pcts}
    return results


def chart_top_products(live_prod, shop_prod):
    """Horizontal bar: top products by Live GMV vs Shop Total GMV."""
    if live_prod is None and shop_prod is None:
        return {}

    rows = {}
    if live_prod is not None:
        pid_col  = pick(live_prod, "Product ID")
        name_col = pick(live_prod, "Product name", "Product Name")
        gmv_col  = pick(live_prod, "Attributed GMV", "GMV")
        if pid_col and gmv_col:
            for _, r in live_prod.iterrows():
                pid = str(r[pid_col]).strip()
                rows.setdefault(pid, {})["live_gmv"] = r[gmv_col]
                if name_col:
                    rows[pid]["name"] = r[name_col]

    if shop_prod is not None:
        pid_col  = pick(shop_prod, "Product ID")
        name_col = pick(shop_prod, "Product Name", "Product name")
        gmv_col  = pick(shop_prod, "GMV")
        if pid_col and gmv_col:
            for _, r in shop_prod.iterrows():
                pid = str(r[pid_col]).strip()
                rows.setdefault(pid, {})["shop_total"] = r[gmv_col]
                if name_col and "name" not in rows[pid]:
                    rows[pid]["name"] = r[name_col]

    # Sort by live GMV, take top 10
    sorted_rows = sorted(rows.items(), key=lambda x: x[1].get("live_gmv", 0), reverse=True)[:10]
    def short_label(text, limit=54):
        text = str(text)
        return text if len(text) <= limit else text[:limit - 1].rstrip() + "…"

    labels   = [short_label(v.get("name", pid)) for pid, v in sorted_rows]
    live_gmv = [round(v.get("live_gmv", 0), 2)  for _, v in sorted_rows]
    shop_gmv = [round(v.get("shop_total", 0), 2) for _, v in sorted_rows]
    return {"labels": labels, "live_gmv": live_gmv, "shop_gmv": shop_gmv}


def table_live_sessions(live_perf):
    """Table rows for each live session."""
    if live_perf is None:
        return []

    name_col     = pick(live_perf, "Livestream", "Stream name", "Name")
    date_col     = "Start time"
    dur_col      = pick(live_perf, "Duration")
    gmv_col      = pick(live_perf, "Attributed GMV", "GMV")
    orders_col   = pick(live_perf, "Attributed orders", "Orders")
    views_col    = pick(live_perf, "Views")
    followers_col= pick(live_perf, "New followers", "Followers")

    rows = []
    for _, r in live_perf.iterrows():
        def val(col):
            if col and col in r.index:
                v = r[col]
                if pd.isna(v):
                    return None
                if isinstance(v, float) and v == int(v):
                    return int(v)
                return v
            return None

        date_val = r.get(date_col) if date_col in r.index else None
        rows.append({
            "name":      val(name_col) or "—",
            "date":      str(date_val.date()) if pd.notna(date_val) else "—",
            "duration":  val(dur_col) or "—",
            "gmv":       val(gmv_col),
            "orders":    val(orders_col),
            "views":     val(views_col),
            "followers": val(followers_col),
        })

    rows.sort(key=lambda x: x["date"])
    return rows


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    return jsonify({"ok": True})


def build_dashboard_payload(live_prod, shop_prod, live_perf, shop_daily, files_received):
    payload = {
        "gmv_over_time":  chart_gmv_over_time(live_perf, shop_daily),
        "live_pct":       chart_live_pct(live_perf, shop_daily),
        "top_products":   chart_top_products(live_prod, shop_prod),
        "live_sessions":  table_live_sessions(live_perf),
        "files_received": files_received,
    }
    return to_jsonable(payload)


def to_jsonable(value):
    if isinstance(value, dict):
        return {k: to_jsonable(v) for k, v in value.items()}
    if isinstance(value, list):
        return [to_jsonable(v) for v in value]
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        return value.item()
    return value


def build_payload_from_uploads(files):
    live_prod = load_live_product(files["live_product"]) if "live_product" in files and getattr(files["live_product"], "filename", "") else None
    shop_prod = load_shop_product(files["shop_product"]) if "shop_product" in files and getattr(files["shop_product"], "filename", "") else None
    live_perf = load_live_perf(files["live_performance"]) if "live_performance" in files and getattr(files["live_performance"], "filename", "") else None
    shop_daily = load_shop_daily(files["shop_daily"]) if "shop_daily" in files and getattr(files["shop_daily"], "filename", "") else None

    files_received = [
        k for k in ["live_product", "shop_product", "live_performance", "shop_daily"]
        if k in files and getattr(files[k], "filename", "")
    ]
    return build_dashboard_payload(live_prod, shop_prod, live_perf, shop_daily, files_received)


def _existing(paths):
    return [p for p in paths if p.exists()]


def desktop_data_paths():
    return [
        WORKSPACE_DIR / "LIVEDATA/5:27 直播数据/0527live_luluthepiggyshop_7644709935498578718_product.xlsx",
        WORKSPACE_DIR / "SHOPDATA/product_list_20260523.xlsx",
        WORKSPACE_DIR / "LIVEDATA/March_Creator-Live-Performance_20260530033847.xlsx",
        WORKSPACE_DIR / "LIVEDATA/April_Creator-Live-Performance_20260530033910.xlsx",
        WORKSPACE_DIR / "LIVEDATA/5:27 直播数据/0527live_luluthepiggyshop_7644709935498578718_trend_stats.xlsx",
        WORKSPACE_DIR / "SHOPDATA/April.xlsx",
        WORKSPACE_DIR / "SHOPDATA/May 5:1-5:29.xlsx",
    ]


def load_desktop_dataset():
    live_product_paths = _existing([
        WORKSPACE_DIR / "LIVEDATA/5:27 直播数据/0527live_luluthepiggyshop_7644709935498578718_product.xlsx",
        WORKSPACE_DIR / "LIVEDATA/5:27 直播数据/LIVE Dashboard 0527.xlsx",
    ])
    shop_product_paths = _existing([
        WORKSPACE_DIR / "SHOPDATA/product_list_20260523.xlsx",
        WORKSPACE_DIR / "SHOPDATA/Product_Traffic_[total]_Key_Metrics_23_05_2026-29_05_2026.xlsx",
    ])
    live_perf_paths = _existing([
        WORKSPACE_DIR / "LIVEDATA/March_Creator-Live-Performance_20260530033847.xlsx",
        WORKSPACE_DIR / "LIVEDATA/April_Creator-Live-Performance_20260530033910.xlsx",
    ])
    live_trend_paths = _existing([
        WORKSPACE_DIR / "LIVEDATA/5:27 直播数据/0527live_luluthepiggyshop_7644709935498578718_trend_stats.xlsx",
    ])
    shop_daily_paths = _existing([
        WORKSPACE_DIR / "SHOPDATA/April.xlsx",
        WORKSPACE_DIR / "SHOPDATA/May 5:1-5:29.xlsx",
    ])

    live_prod = load_live_product(live_product_paths[0]) if live_product_paths else None
    shop_prod = load_shop_product(shop_product_paths[0]) if shop_product_paths else None

    live_frames = [load_live_perf(path) for path in live_perf_paths]
    live_frames.extend(
        load_live_trend_as_session(path, "2026-05-27 16:00", "LuLu 5/27 LIVE")
        for path in live_trend_paths
    )
    live_perf = pd.concat(live_frames, ignore_index=True) if live_frames else None

    shop_frames = [load_shop_daily(path) for path in shop_daily_paths]
    shop_daily = pd.concat(shop_frames, ignore_index=True) if shop_frames else None

    received = [
        str(p.relative_to(WORKSPACE_DIR))
        for p in live_product_paths[:1] + shop_product_paths[:1] + live_perf_paths + live_trend_paths + shop_daily_paths
    ]
    return live_prod, shop_prod, live_perf, shop_daily, received


@app.route("/app_config")
def app_config():
    input_paths = [
        INPUT_DIR / "live_product.csv",
        INPUT_DIR / "shop_product.csv",
        INPUT_DIR / "live_performance.csv",
        INPUT_DIR / "shop_daily.csv",
    ]
    return jsonify({
        "ok": True,
        "has_desktop_data": all(path.exists() for path in desktop_data_paths()),
        "has_input_data": all(path.exists() for path in input_paths),
        "is_public_deploy": os.environ.get("PUBLIC_DEPLOY", "").lower() in {"1", "true", "yes"},
    })


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files
    try:
        payload = build_payload_from_uploads(files)
        return jsonify({"ok": True, "data": payload})

    except Exception:
        return jsonify({"ok": False, "error": traceback.format_exc()}), 500


@app.route("/load_input_data", methods=["POST"])
def load_input_data():
    try:
        paths = {
            "live_product": INPUT_DIR / "live_product.csv",
            "shop_product": INPUT_DIR / "shop_product.csv",
            "live_performance": INPUT_DIR / "live_performance.csv",
            "shop_daily": INPUT_DIR / "shop_daily.csv",
        }
        missing = [name for name, path in paths.items() if not path.exists()]
        if missing:
            return jsonify({"ok": False, "error": f"Missing file(s) in input_data/: {', '.join(missing)}"}), 400

        live_prod = load_live_product(paths["live_product"])
        shop_prod = load_shop_product(paths["shop_product"])
        live_perf = load_live_perf(paths["live_performance"])
        shop_daily = load_shop_daily(paths["shop_daily"])
        payload = build_dashboard_payload(live_prod, shop_prod, live_perf, shop_daily, list(paths.keys()))
        return jsonify({"ok": True, "data": payload})

    except Exception:
        return jsonify({"ok": False, "error": traceback.format_exc()}), 500


@app.route("/load_desktop_data", methods=["POST"])
def load_desktop_data():
    try:
        live_prod, shop_prod, live_perf, shop_daily, files_received = load_desktop_dataset()
        if not files_received:
            return jsonify({"ok": False, "error": "No supported files found in SHOPDATA/ or LIVEDATA/."}), 400
        payload = build_dashboard_payload(live_prod, shop_prod, live_perf, shop_daily, files_received)
        return jsonify({"ok": True, "data": payload})

    except Exception:
        return jsonify({"ok": False, "error": traceback.format_exc()}), 500


def _json_bytes(payload, status=200):
    body = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
    return status, body


def _normalize_response(response):
    status = 200
    payload = response
    if isinstance(response, tuple):
        payload = response[0]
        if len(response) > 1:
            status = response[1]
    return _json_bytes(payload, status)


def run_simple_server(host, port):
    try:
        import cgi
    except ModuleNotFoundError:
        cgi = None

    class DashboardHandler(BaseHTTPRequestHandler):
        def _send(self, status, body, content_type):
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            path = urlparse(self.path).path
            if path == "/health":
                status, body = _json_bytes({"ok": True})
                self._send(status, body, "application/json; charset=utf-8")
                return
            if path == "/app_config":
                status, body = _normalize_response(app_config())
                self._send(status, body, "application/json; charset=utf-8")
                return
            if path != "/":
                self._send(404, b"Not found", "text/plain; charset=utf-8")
                return
            body = (BASE_DIR / "templates" / "index.html").read_bytes()
            self._send(200, body, "text/html; charset=utf-8")

        def do_POST(self):
            path = urlparse(self.path).path
            try:
                if path == "/load_desktop_data":
                    status, body = _normalize_response(load_desktop_data())
                elif path == "/load_input_data":
                    status, body = _normalize_response(load_input_data())
                elif path == "/upload":
                    if cgi is None:
                        status, body = _json_bytes({"ok": False, "error": "File upload parser is unavailable in this Python."}, 500)
                    else:
                        form = cgi.FieldStorage(
                            fp=self.rfile,
                            headers=self.headers,
                            environ={
                                "REQUEST_METHOD": "POST",
                                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                                "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                            },
                        )
                        files = {key: form[key] for key in form.keys() if getattr(form[key], "filename", "")}
                        payload = build_payload_from_uploads(files)
                        status, body = _json_bytes({"ok": True, "data": payload})
                else:
                    status, body = _json_bytes({"ok": False, "error": "Not found"}, 404)
            except Exception:
                status, body = _json_bytes({"ok": False, "error": traceback.format_exc()}, 500)
            self._send(status, body, "application/json; charset=utf-8")

        def log_message(self, fmt, *args):
            print(fmt % args)

    server = ThreadingHTTPServer((host, port), DashboardHandler)
    print(f"Dashboard running at http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5050"))
    host = os.environ.get("HOST", "127.0.0.1")
    app.run(debug=True, host=host, port=port)
