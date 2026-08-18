"""
Convert raw TikTok xlsx exports → 4 CSVs for the dashboard pipeline.

Source folders (relative to this script's location):
  ../LIVEDATA/   — live performance xlsx files
  ../SHOPDATA/   — shop overview xlsx files

Output (overwritten each run):
  input_data/live_performance.csv
  input_data/shop_daily.csv
  input_data/live_product.csv
  input_data/shop_product.csv
"""

import glob
import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent
LIVE = BASE.parent / "LIVEDATA"
SHOP = BASE.parent / "SHOPDATA"
OUT  = BASE / "input_data"
OUT.mkdir(exist_ok=True)


def read_sheet(path, header_row, sheet=0):
    return pd.read_excel(
        path, sheet_name=sheet, header=header_row - 1, dtype=str
    ).dropna(how="all")


def write_csv(df, name):
    path = OUT / name
    df.to_csv(path, index=False)
    print(f"  ✓ {name}  ({len(df)} rows, {len(df.columns)} cols)")


# ── 1. live_performance.csv ────────────────────────────────────────────────
# March + April Creator-Live-Performance xlsx, headers on row 3

def build_live_performance():
    print("\n[1] live_performance.csv")
    files = sorted(glob.glob(str(LIVE / "*Creator-Live-Performance*.xlsx")))
    if not files:
        print("  [SKIP] No Creator-Live-Performance files in LIVEDATA/")
        return

    frames = []
    for f in files:
        print(f"     {Path(f).name}")
        df = read_sheet(f, header_row=3)
        df = df.dropna(subset=[df.columns[0]])
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.rename(columns={"Viewers": "Views"})
    write_csv(combined, "live_performance.csv")


# ── 2. shop_daily.csv ──────────────────────────────────────────────────────
# April.xlsx + May*.xlsx — daily data section starts at row 9

def build_shop_daily():
    print("\n[2] shop_daily.csv")
    # Match shop overview files by name pattern; exclude Creator/Samples/affiliate files
    candidates = (
        glob.glob(str(SHOP / "April.xlsx")) +
        glob.glob(str(SHOP / "May*.xlsx"))
    )
    files = sorted(
        f for f in candidates
        if not any(x in Path(f).name for x in ["Creator", "Samples", "Affiliate"])
    )
    if not files:
        print("  [SKIP] No April.xlsx / May*.xlsx in SHOPDATA/")
        return

    frames = []
    for f in files:
        print(f"     {Path(f).name}")
        df = read_sheet(f, header_row=9)
        date_col = df.columns[0]
        # Drop blank rows and the repeated header row that sometimes appears
        df = df[df[date_col].notna() & (df[date_col].astype(str).str.strip() != "Date")]
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    keep = ["Date", "GMV", "Orders", "Items sold", "Product impressions",
            "LIVE-attributed GMV", "Customers"]
    combined = combined[[c for c in keep if c in combined.columns]]
    write_csv(combined, "shop_daily.csv")


# ── 3. live_product.csv ────────────────────────────────────────────────────
# 0527live_*_product.xlsx — headers on row 1

def build_live_product():
    print("\n[3] live_product.csv")
    files = sorted(glob.glob(str(LIVE / "**/*_product.xlsx"), recursive=True))
    if not files:
        print("  [SKIP] No *_product.xlsx in LIVEDATA/")
        return

    frames = []
    for f in files:
        print(f"     {Path(f).name}")
        df = read_sheet(f, header_row=1)
        df = df.dropna(subset=[df.columns[0]])
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.rename(columns={"Attributed SKU orders": "Attributed orders"})
    write_csv(combined, "live_product.csv")


# ── 4. shop_product.csv ────────────────────────────────────────────────────
# product_list_*.xlsx — headers on row 4

def build_shop_product():
    print("\n[4] shop_product.csv")
    files = sorted(glob.glob(str(SHOP / "product_list_*.xlsx")))
    if not files:
        print("  [SKIP] No product_list_*.xlsx in SHOPDATA/")
        return

    frames = []
    for f in files:
        print(f"     {Path(f).name}")
        df = read_sheet(f, header_row=4)
        df = df.dropna(subset=[df.columns[0]])
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.rename(columns={
        "Seller LIVE-attributed GMV": "Seller LIVE GMV",
        "Product card GMV":           "Seller product card GMV",
    })
    combined = combined.loc[:, ~combined.columns.duplicated()]
    write_csv(combined, "shop_product.csv")


# ── main ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=== TikTok XLSX → CSV Converter ===")
    build_live_performance()
    build_shop_daily()
    build_live_product()
    build_shop_product()
    print("\nDone. Files written to nextt_pipeline/input_data/\n")
