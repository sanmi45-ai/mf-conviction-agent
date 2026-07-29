#!/usr/bin/env python3
"""
mf_conviction_agent_amfi.py

AMFI-DIRECT version. Instead of relying on the third-party mfdata.in API
(which has shown repeated downtime), this pulls straight from AMFI's own
servers: the official, SEBI-mandated monthly portfolio disclosure file.

URL pattern (confirmed working):
    https://portal.amfiindia.com/spages/am{mon}{year}repo.xls
    e.g. https://portal.amfiindia.com/spages/amjun2026repo.xls

This is ONE big Excel file covering every scheme from every AMC for that
month. We don't yet know its exact internal column layout (sheet names,
header rows, etc.) — that's what --debug mode is for. Run this in debug
mode FIRST and share the printed output; then the FUND-matching / column
logic below can be corrected in one pass.

USAGE
    python mf_conviction_agent_amfi.py --debug     # inspect file structure first
    python mf_conviction_agent_amfi.py              # normal run (once calibrated)
    python mf_conviction_agent_amfi.py --top 30
"""

import argparse
import calendar
import datetime
import json
import os
import sys

import pandas as pd
import requests

SNAPSHOT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "snapshots")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
RAW_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "raw_downloads")

FUNDS = [
    {"label": "Bandhan Small Cap Fund", "match": ["bandhan", "small cap"]},
    {"label": "Nippon India Small Cap Fund", "match": ["nippon india", "small cap"]},
    {"label": "quant Small Cap Fund", "match": ["quant", "small cap"]},
]

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "Mozilla/5.0 (mf-conviction-agent personal research script)"})


def amfi_url_for(year, month):
    mon = calendar.month_abbr[month].lower()  # 'jun', 'jul', etc.
    return f"https://portal.amfiindia.com/spages/am{mon}{year}repo.xls"


def download_amfi_file(debug=False):
    """
    AMFI publishes month M's data by the 10th of month M+1. So depending on
    when this script runs, the most recently COMPLETE dataset is either
    this month's file (if it happens to already be published) or last
    month's. We try current month first, then fall back one month at a
    time (up to 2 months back) if the file doesn't exist yet.
    """
    today = datetime.date.today()
    candidates = []
    y, m = today.year, today.month
    for _ in range(3):
        candidates.append((y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1

    os.makedirs(RAW_DIR, exist_ok=True)
    for (year, month) in candidates:
        url = amfi_url_for(year, month)
        print(f"  Trying {url}")
        try:
            resp = SESSION.get(url, timeout=30)
            if resp.status_code == 200 and len(resp.content) > 10000:
                month_str = f"{year}-{month:02d}"
                raw_path = os.path.join(RAW_DIR, f"amfi_{month_str}.xls")
                with open(raw_path, "wb") as f:
                    f.write(resp.content)
                print(f"    -> downloaded ({len(resp.content)} bytes) as data for {month_str}")
                return raw_path, month_str
            else:
                print(f"    -> not available yet (status {resp.status_code}, "
                      f"{len(resp.content)} bytes)")
        except requests.exceptions.RequestException as e:
            print(f"    -> request failed: {e}")
    raise RuntimeError(
        "Could not download any of the last 3 months' AMFI files. "
        "Either AMFI's server is down, or the URL pattern needs updating "
        "— check https://www.amfiindia.com/research-information/other-data/monthly-portfolio-disclosures manually."
    )


def inspect_file(raw_path):
    """Debug helper: print sheet names, shape, and a sample of each sheet
    so we can figure out the real column layout."""
    xls = pd.ExcelFile(raw_path, engine="xlrd")
    print(f"\n  Sheet names ({len(xls.sheet_names)} total): {xls.sheet_names[:20]}"
          f"{' ...' if len(xls.sheet_names) > 20 else ''}")

    # Look at the first sheet (or first couple) in detail
    for sheet in xls.sheet_names[:2]:
        df = pd.read_excel(raw_path, sheet_name=sheet, engine="xlrd", header=None, nrows=15)
        print(f"\n  --- First 15 rows of sheet '{sheet}' (shape hint: {df.shape}) ---")
        with pd.option_context("display.max_columns", 10, "display.width", 160):
            print(df)


def find_fund_data(raw_path, fund, debug=False):
    """
    Placeholder matching logic — WILL need adjustment once we see the real
    file structure via --debug. Current assumption: either
      (a) one sheet per scheme, sheet name containing the fund name, or
      (b) one flat sheet with a "Scheme Name" column we can filter on.
    Tries (a) first, falls back to (b).
    """
    xls = pd.ExcelFile(raw_path, engine="xlrd")
    match_terms = [t.lower() for t in fund["match"]]

    # Attempt (a): sheet name matches
    for sheet in xls.sheet_names:
        name_lower = sheet.lower()
        if all(t in name_lower for t in match_terms):
            df = pd.read_excel(raw_path, sheet_name=sheet, engine="xlrd")
            if debug:
                print(f"  [debug] matched via sheet name: '{sheet}', columns: {list(df.columns)}")
            return df

    # Attempt (b): flat sheet(s) with a scheme-name-like column
    for sheet in xls.sheet_names:
        df = pd.read_excel(raw_path, sheet_name=sheet, engine="xlrd")
        scheme_col = None
        for col in df.columns:
            if "scheme" in str(col).lower():
                scheme_col = col
                break
        if scheme_col is None:
            continue
        mask = df[scheme_col].astype(str).str.lower().apply(
            lambda v: all(t in v for t in match_terms)
        )
        if mask.any():
            if debug:
                print(f"  [debug] matched via column '{scheme_col}' in sheet '{sheet}'")
            return df[mask]

    raise RuntimeError(
        f"Could not find data for '{fund['label']}' in the AMFI file using "
        f"either sheet-name or scheme-column matching. Run with --debug and "
        f"share the printed sheet/column structure so the matching logic "
        f"can be corrected."
    )


def extract_holdings(df, debug=False):
    """Given a fund's raw rows, extract [{isin, name, weight}, ...].
    Column-name guessing — will likely need tweaking after --debug run."""
    cols = {str(c).lower().strip(): c for c in df.columns}

    def find_col(*keywords):
        for k, orig in cols.items():
            if all(kw in k for kw in keywords):
                return orig
        return None

    isin_col = find_col("isin")
    name_col = find_col("instrument") or find_col("name") or find_col("company")
    weight_col = find_col("%") or find_col("nav") and find_col("percentage")

    if debug:
        print(f"  [debug] resolved columns -> isin: {isin_col}, name: {name_col}, weight: {weight_col}")

    if not isin_col or not name_col:
        raise RuntimeError(
            f"Couldn't confidently identify ISIN/name columns. Available "
            f"columns were: {list(df.columns)}. Run with --debug to inspect."
        )

    holdings = []
    for _, row in df.iterrows():
        isin = row.get(isin_col)
        name = row.get(name_col)
        weight = row.get(weight_col) if weight_col else None
        if pd.isna(isin) or pd.isna(name):
            continue
        try:
            weight = float(weight) if weight is not None and not pd.isna(weight) else 0.0
        except (TypeError, ValueError):
            weight = 0.0
        holdings.append({"isin": str(isin).strip(), "name": str(name).strip(), "weight": weight})
    return holdings


# ---------------------------------------------------------------------------
# Snapshot storage + ranking logic (same approach as the mfdata.in version)
# ---------------------------------------------------------------------------
def snapshot_path(label, month_str):
    safe_label = label.replace(" ", "_")
    return os.path.join(SNAPSHOT_DIR, f"{safe_label}_{month_str}.json")


def save_snapshot(label, month_str, holdings):
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    with open(snapshot_path(label, month_str), "w") as f:
        json.dump(holdings, f, indent=2)


def load_previous_snapshot(label, current_month_str):
    if not os.path.isdir(SNAPSHOT_DIR):
        return None
    safe_label = label.replace(" ", "_")
    prefix = f"{safe_label}_"
    candidates = []
    for fname in os.listdir(SNAPSHOT_DIR):
        if fname.startswith(prefix) and fname.endswith(".json"):
            month_part = fname[len(prefix):-len(".json")]
            if month_part < current_month_str:
                candidates.append(month_part)
    if not candidates:
        return None
    latest_prev = max(candidates)
    with open(snapshot_path(label, latest_prev)) as f:
        return {"month": latest_prev, "holdings": json.load(f)}


def build_conviction_list(fund_holdings, fund_prev_holdings, top_n):
    stocks = {}
    for label, holdings in fund_holdings.items():
        current_isins = {h["isin"] for h in holdings}
        prev = fund_prev_holdings.get(label)
        prev_isins = {h["isin"] for h in prev} if prev else set()

        for h in holdings:
            entry = stocks.setdefault(h["isin"], {"name": h["name"], "funds": {}, "entries": [], "exits": []})
            entry["funds"][label] = h["weight"]
            entry["name"] = h["name"]
            if prev is not None and h["isin"] not in prev_isins:
                entry["entries"].append(label)

        if prev is not None:
            for h in prev:
                if h["isin"] not in current_isins:
                    entry = stocks.setdefault(h["isin"], {"name": h["name"], "funds": {}, "entries": [], "exits": []})
                    entry["exits"].append(label)

    rows = []
    for isin, info in stocks.items():
        count = len(info["funds"])
        avg_weight = sum(info["funds"].values()) / count if count else 0.0
        rows.append({
            "isin": isin, "name": info["name"], "conviction_count": count,
            "avg_weight_pct": round(avg_weight, 2), "weights": info["funds"],
            "recent_entries": ", ".join(info["entries"]),
            "recent_exits": ", ".join(info["exits"]),
        })
    rows.sort(key=lambda r: (r["conviction_count"], r["avg_weight_pct"]), reverse=True)
    return rows[:top_n]


def write_csv(rows, all_labels, month_str):
    import csv
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f"conviction_list_{month_str}.csv")
    with open(out_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["rank", "stock", "isin", "conviction_count", "avg_weight_pct"]
                    + [f"weight_in::{l}" for l in all_labels] + ["recent_entries", "recent_exits"])
        for i, r in enumerate(rows, 1):
            w.writerow([i, r["name"], r["isin"], r["conviction_count"], r["avg_weight_pct"]]
                        + [r["weights"].get(l, "") for l in all_labels]
                        + [r["recent_entries"], r["recent_exits"]])
    return out_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--top", type=int, default=20)
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--inspect-only", action="store_true",
                         help="just download and print file structure, don't try to parse funds")
    args = parser.parse_args()

    print("Downloading AMFI monthly disclosure file...")
    raw_path, month_str = download_amfi_file(debug=args.debug)

    if args.debug or args.inspect_only:
        inspect_file(raw_path)
        if args.inspect_only:
            print("\n--inspect-only set, stopping here. Share the output above.")
            return

    existing_output = os.path.join(OUTPUT_DIR, f"conviction_list_{month_str}.csv")
    if os.path.exists(existing_output) and not args.debug:
        print(f"{existing_output} already exists — skipping.")
        return

    fund_holdings, fund_prev_holdings = {}, {}
    for fund in FUNDS:
        label = fund["label"]
        print(f"\nExtracting: {label}")
        df = find_fund_data(raw_path, fund, debug=args.debug)
        holdings = extract_holdings(df, debug=args.debug)
        print(f"  -> {len(holdings)} holdings extracted")
        fund_holdings[label] = holdings

        prev = load_previous_snapshot(label, month_str)
        fund_prev_holdings[label] = prev["holdings"] if prev else None
        save_snapshot(label, month_str, holdings)

    rows = build_conviction_list(fund_holdings, fund_prev_holdings, args.top)
    if not rows:
        print("No rows produced — check parsing logic with --debug.")
        sys.exit(1)

    print(f"\n=== Top {len(rows)} conviction stocks — {month_str} ===")
    for i, r in enumerate(rows, 1):
        flags = []
        if r["recent_entries"]:
            flags.append(f"NEW ({r['recent_entries']})")
        if r["recent_exits"]:
            flags.append(f"EXIT ({r['recent_exits']})")
        print(f"{i:2d}. {r['name']:<35} conviction={r['conviction_count']}/3 "
              f"avg_wt={r['avg_weight_pct']}%  {' | '.join(flags)}")

    out_path = write_csv(rows, [f["label"] for f in FUNDS], month_str)
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
