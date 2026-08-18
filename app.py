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


舞台