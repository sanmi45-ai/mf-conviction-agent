#!/usr/bin/env python3
"""
mf_conviction_agent_multisource.py

Pulls holdings for the three tracked funds from THREE DIFFERENT sources,
since each fund's data turned out to live in a different place:

  - Nippon India Small Cap  -> scraped directly from mf.nipponindiaim.com
                                (CONFIRMED working: plain HTML, real .xls links)
  - quant Small Cap         -> scraped directly from quantmutual.com
                                (BEST EFFORT: page structure not fully verified
                                yet -- run with --debug and share output if it
                                fails, same as we did for Nippon India)
  - Bandhan Small Cap       -> Bandhan's own site is a JS app we can't scrape.
                                Two options, both built in:
                                  (a) mfdata.in API, used ONLY for this fund
                                  (b) manual file drop: place the .xls you
                                      downloaded yourself into
                                      manual_downloads/bandhan_YYYY-MM.xls
                                      and the script will use it instead of
                                      hitting the network for Bandhan.

USAGE
    python mf_conviction_agent_multisource.py --debug
    python mf_conviction_agent_multisource.py
    python mf_conviction_agent_multisource.py --top 30
"""

import argparse
import datetime
import json
import os
import re
import sys

import pandas as pd
import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SNAPSHOT_DIR = os.path.join(BASE_DIR, "snapshots")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
RAW_DIR = os.path.join(BASE_DIR, "raw_downloads")
MANUAL_DIR = os.path.join(BASE_DIR, "manual_downloads")

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "Mozilla/5.0 (mf-conviction-agent personal research script)"})

MONTH_STR = datetime.date.today().strftime("%Y-%m")


def log(msg):
    print(msg)


# ---------------------------------------------------------------------------
# Shared xls parsing helpers
# ---------------------------------------------------------------------------
def find_col(df, *keyword_sets):
    """keyword_sets: list of tuples; each tuple's keywords must ALL appear
    (case-insensitive) in a column name for it to match. Returns first match
    across the tuples, in priority order."""
    cols = {str(c).lower().strip(): c for c in df.columns}
    for keywords in keyword_sets:
        for k, orig in cols.items():
            if all(kw in k for kw in keywords):
                return orig
    return None


def extract_holdings_from_df(df, debug=False, label=""):
    isin_col = find_col(df, ("isin",))
    name_col = find_col(df, ("instrument",), ("name of instrument",), ("company",), ("name",))
    weight_col = find_col(df, ("%", "nav"), ("percentage", "nav"), ("weightage",))

    if debug:
        log(f"    [debug:{label}] columns available: {list(df.columns)}")
        log(f"    [debug:{label}] resolved -> isin: {isin_col}, name: {name_col}, weight: {weight_col}")

    if not isin_col or not name_col:
        return None  # caller decides how to handle

    holdings = []
    for _, row in df.iterrows():
        isin = row.get(isin_col)
        name = row.get(name_col)
        weight = row.get(weight_col) if weight_col else None
        if pd.isna(isin) or pd.isna(name):
            continue
        isin_s = str(isin).strip()
        if not re.match(r"^IN[A-Z0-9]{10}$", isin_s):
            continue  # skip header/subtotal rows that aren't real ISINs
        try:
            weight_f = float(weight) if weight is not None and not pd.isna(weight) else 0.0
        except (TypeError, ValueError):
            weight_f = 0.0
        holdings.append({"isin": isin_s, "name": str(name).strip(), "weight": weight_f})
    return holdings


def find_scheme_sheet_or_rows(raw_path, match_terms, debug=False, label=""):
    """Try to locate a fund's holdings within a downloaded workbook, either
    via a matching sheet name or a scheme-name column inside a flat sheet."""
    engine = "xlrd" if raw_path.lower().endswith(".xls") else "openpyxl"
    xls = pd.ExcelFile(raw_path, engine=engine)
    terms = [t.lower() for t in match_terms]

    if debug:
        log(f"    [debug:{label}] sheets: {xls.sheet_names[:15]}"
            f"{' ...' if len(xls.sheet_names) > 15 else ''}")

    for sheet in xls.sheet_names:
        if all(t in sheet.lower() for t in terms):
            return pd.read_excel(raw_path, sheet_name=sheet, engine=engine)

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(raw_path, sheet_name=sheet, engine=engine)
        except Exception:
            continue
        scheme_col = find_col(df, ("scheme",))
        if scheme_col is None:
            continue
        mask = df[scheme_col].astype(str).str.lower().apply(lambda v: all(t in v for t in terms))
        if mask.any():
            return df[mask]

    return None


# ---------------------------------------------------------------------------
# SOURCE 1: Nippon India (confirmed scrapable)
# ---------------------------------------------------------------------------
def fetch_nippon_india(debug=False):
    log("  [Nippon India] fetching disclosures page...")
    url = "https://mf.nipponindiaim.com/investor-service/downloads/factsheet-portfolio-and-other-disclosures"
    resp = SESSION.get(url, timeout=30)
    resp.raise_for_status()
    html = resp.text

    # Find the first "Monthly portfolio for the month..." link (list is
    # newest-first). The href sits right after that link text.
    pattern = re.compile(
        r'Monthly portfolio[^<]*?(?:month[^<]*)?</a>.*?href="([^"]+\.xlsx?)"',
        re.IGNORECASE | re.DOTALL,
    )
    m = pattern.search(html)
    if not m:
        # Fallback: just grab hrefs near the text "Monthly portfolio"
        idx = html.lower().find("monthly portfolio")
        if idx == -1:
            raise RuntimeError("Could not find 'Monthly portfolio' text on Nippon India's page at all.")
        window = html[idx: idx + 1500]
        m2 = re.search(r'href="([^"]+\.xlsx?)"', window)
        if not m2:
            raise RuntimeError("Found 'Monthly portfolio' text but no .xls link nearby. "
                                "Site structure may have changed -- run --debug and inspect.")
        file_url = m2.group(1)
    else:
        file_url = m.group(1)

    if not file_url.startswith("http"):
        file_url = "https://mf.nipponindiaim.com" + file_url

    log(f"    -> found file: {file_url}")
    os.makedirs(RAW_DIR, exist_ok=True)
    raw_path = os.path.join(RAW_DIR, f"nippon_{MONTH_STR}.xls")
    fresp = SESSION.get(file_url, timeout=60)
    fresp.raise_for_status()
    with open(raw_path, "wb") as f:
        f.write(fresp.content)
    log(f"    -> downloaded ({len(fresp.content)} bytes)")

    df = find_scheme_sheet_or_rows(raw_path, ["small cap"], debug=debug, label="nippon")
    if df is None:
        raise RuntimeError("Downloaded Nippon India file but couldn't locate 'Small Cap' scheme rows. "
                            "Run --debug to inspect sheet/column names.")
    holdings = extract_holdings_from_df(df, debug=debug, label="nippon")
    if holdings is None:
        raise RuntimeError("Found Small Cap rows but couldn't identify ISIN/name columns. Run --debug.")
    return holdings


# ---------------------------------------------------------------------------
# SOURCE 2: quant (best-effort, page structure not fully verified)
# ---------------------------------------------------------------------------
def fetch_quant(debug=False):
    log("  [quant] fetching statutory-disclosures page...")
    url = "https://quantmutual.com/statutory-disclosures"
    resp = SESSION.get(url, timeout=30)
    resp.raise_for_status()
    html = resp.text

    # Look for any .xls/.xlsx link whose surrounding text mentions "small cap"
    # and "monthly" -- best effort, may need calibration.
    candidates = re.findall(r'href="([^"]+\.xlsx?)"[^>]*>([^<]*)</a>', html, re.IGNORECASE)
    if debug:
        log(f"    [debug:quant] found {len(candidates)} total xls/xlsx links on page")
        for href, text in candidates[:20]:
            log(f"      {text.strip()[:60]!r} -> {href}")

    scored = [
        (href, text) for href, text in candidates
        if "small" in text.lower() or "small" in href.lower()
    ]
    if not scored:
        # Fall back: any monthly-portfolio-looking link at all -- we'll need
        # to inspect it by hand this first time.
        scored = [
            (href, text) for href, text in candidates
            if "portfolio" in href.lower() or "monthly" in text.lower()
        ]
    if not scored:
        raise RuntimeError(
            "Could not find any Small Cap / monthly portfolio link on quant's page. "
            "Run --debug -- it prints every .xls/.xlsx link found so we can "
            "identify the right one and hardcode the correct match logic."
        )

    href, text = scored[0]
    file_url = href if href.startswith("http") else "https://quantmutual.com" + href
    log(f"    -> best-effort match: {text.strip()[:60]!r} -> {file_url}")

    os.makedirs(RAW_DIR, exist_ok=True)
    raw_path = os.path.join(RAW_DIR, f"quant_{MONTH_STR}.xls" if file_url.endswith(".xls")
                             else f"quant_{MONTH_STR}.xlsx")
    fresp = SESSION.get(file_url, timeout=60)
    fresp.raise_for_status()
    with open(raw_path, "wb") as f:
        f.write(fresp.content)
    log(f"    -> downloaded ({len(fresp.content)} bytes)")

    df = find_scheme_sheet_or_rows(raw_path, ["small cap"], debug=debug, label="quant")
    if df is None:
        raise RuntimeError("Downloaded a quant file but couldn't locate 'Small Cap' scheme rows "
                            "-- this file might not be the right one. Run --debug.")
    holdings = extract_holdings_from_df(df, debug=debug, label="quant")
    if holdings is None:
        raise RuntimeError("Found rows but couldn't identify ISIN/name columns. Run --debug.")
    return holdings


# ---------------------------------------------------------------------------
# SOURCE 3: Bandhan -- manual file drop OR mfdata.in fallback
# ---------------------------------------------------------------------------
def fetch_bandhan(debug=False):
    os.makedirs(MANUAL_DIR, exist_ok=True)
    manual_candidates = [
        f for f in os.listdir(MANUAL_DIR)
        if f.lower().startswith("bandhan") and (f.endswith(".xls") or f.endswith(".xlsx"))
    ]
    if manual_candidates:
        # Use the most recently modified manual file
        manual_candidates.sort(key=lambda f: os.path.getmtime(os.path.join(MANUAL_DIR, f)), reverse=True)
        raw_path = os.path.join(MANUAL_DIR, manual_candidates[0])
        log(f"  [Bandhan] using manually-provided file: {manual_candidates[0]}")
        df = find_scheme_sheet_or_rows(raw_path, ["small cap"], debug=debug, label="bandhan")
        if df is None:
            raise RuntimeError(f"Couldn't locate 'Small Cap' rows in {manual_candidates[0]}. Run --debug.")
        holdings = extract_holdings_from_df(df, debug=debug, label="bandhan")
        if holdings is None:
            raise RuntimeError("Found rows but couldn't identify ISIN/name columns. Run --debug.")
        return holdings

    log("  [Bandhan] no manual file found in manual_downloads/ -- trying mfdata.in as fallback...")
    try:
        search = SESSION.get("https://mfdata.in/api/v1/search",
                              params={"q": "Bandhan Small Cap Fund"}, timeout=20)
        search.raise_for_status()
        data = search.json()
        results = data.get("data") or data.get("results") or []
        if not results:
            raise RuntimeError("mfdata.in search returned no results for Bandhan Small Cap.")
        family_id = results[0].get("family_id") or results[0].get("familyId") or results[0].get("id")
        if not family_id:
            raise RuntimeError(f"Couldn't extract family_id from mfdata.in result: {results[0]}")

        hresp = SESSION.get(f"https://mfdata.in/api/v1/families/{family_id}/holdings", timeout=20)
        hresp.raise_for_status()
        hdata = hresp.json()
        raw_holdings = hdata.get("data") or hdata.get("holdings") or []
        if isinstance(raw_holdings, dict):
            raw_holdings = raw_holdings.get("equity_holdings") or raw_holdings.get("holdings") or []

        holdings = []
        for h in raw_holdings:
            isin = h.get("isin") or h.get("ISIN")
            name = h.get("name") or h.get("company") or h.get("instrument")
            weight = h.get("percentage_of_nav") or h.get("percent_nav") or h.get("weight") or 0
            if isin and name:
                holdings.append({"isin": str(isin).strip(), "name": str(name).strip(),
                                  "weight": float(weight) if weight else 0.0})
        if not holdings:
            raise RuntimeError("mfdata.in returned a response but no usable holdings rows.")
        return holdings
    except requests.exceptions.RequestException as e:
        raise RuntimeError(
            f"mfdata.in fallback also failed ({e}). Drop a manually-downloaded "
            f"Bandhan holdings file into manual_downloads/bandhan_{MONTH_STR}.xls "
            f"and re-run."
        )


FUND_SOURCES = {
    "Nippon India Small Cap Fund": fetch_nippon_india,
    "quant Small Cap Fund": fetch_quant,
    "Bandhan Small Cap Fund": fetch_bandhan,
}


# ---------------------------------------------------------------------------
# Snapshot storage + ranking (same as before)
# ---------------------------------------------------------------------------
def snapshot_path(label, month_str):
    safe = label.replace(" ", "_")
    return os.path.join(SNAPSHOT_DIR, f"{safe}_{month_str}.json")


def save_snapshot(label, month_str, holdings):
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    with open(snapshot_path(label, month_str), "w") as f:
        json.dump(holdings, f, indent=2)


def load_previous_snapshot(label, current_month_str):
    if not os.path.isdir(SNAPSHOT_DIR):
        return None
    safe = label.replace(" ", "_")
    prefix = f"{safe}_"
    candidates = [
        fname[len(prefix):-len(".json")]
        for fname in os.listdir(SNAPSHOT_DIR)
        if fname.startswith(prefix) and fname.endswith(".json")
        and fname[len(prefix):-len(".json")] < current_month_str
    ]
    if not candidates:
        return None
    latest = max(candidates)
    with open(snapshot_path(label, latest)) as f:
        return {"month": latest, "holdings": json.load(f)}


def build_conviction_list(fund_holdings, fund_prev_holdings, top_n):
    stocks = {}
    for label, holdings in fund_holdings.items():
        current_isins = {h["isin"] for h in holdings}
        prev = fund_prev_holdings.get(label)
        prev_isins = {h["isin"] for h in prev} if prev else set()
        for h in holdings:
            e = stocks.setdefault(h["isin"], {"name": h["name"], "funds": {}, "entries": [], "exits": []})
            e["funds"][label] = h["weight"]
            e["name"] = h["name"]
            if prev is not None and h["isin"] not in prev_isins:
                e["entries"].append(label)
        if prev is not None:
            for h in prev:
                if h["isin"] not in current_isins:
                    e = stocks.setdefault(h["isin"], {"name": h["name"], "funds": {}, "entries": [], "exits": []})
                    e["exits"].append(label)

    rows = []
    for isin, info in stocks.items():
        count = len(info["funds"])
        avg_w = sum(info["funds"].values()) / count if count else 0.0
        rows.append({"isin": isin, "name": info["name"], "conviction_count": count,
                      "avg_weight_pct": round(avg_w, 2), "weights": info["funds"],
                      "recent_entries": ", ".join(info["entries"]),
                      "recent_exits": ", ".join(info["exits"])})
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
    args = parser.parse_args()

    existing_output = os.path.join(OUTPUT_DIR, f"conviction_list_{MONTH_STR}.csv")
    if os.path.exists(existing_output) and not args.debug:
        log(f"{existing_output} already exists -- skipping (this month already succeeded).")
        return

    fund_holdings, fund_prev_holdings = {}, {}
    failures = []
    for label, fetch_fn in FUND_SOURCES.items():
        log(f"\n{label}:")
        try:
            holdings = fetch_fn(debug=args.debug)
            log(f"  -> {len(holdings)} holdings extracted")
            fund_holdings[label] = holdings
            prev = load_previous_snapshot(label, MONTH_STR)
            fund_prev_holdings[label] = prev["holdings"] if prev else None
            save_snapshot(label, MONTH_STR, holdings)
        except Exception as e:
            log(f"  ! FAILED: {e}")
            failures.append(label)

    if not fund_holdings:
        log("\nAll sources failed this run. Nothing to report.")
        sys.exit(1)

    rows = build_conviction_list(fund_holdings, fund_prev_holdings, args.top)
    log(f"\n=== Top {len(rows)} conviction stocks -- {MONTH_STR} "
        f"({len(fund_holdings)}/3 funds succeeded"
        f"{', missing: ' + ', '.join(failures) if failures else ''}) ===")
    for i, r in enumerate(rows, 1):
        flags = []
        if r["recent_entries"]:
            flags.append(f"NEW ({r['recent_entries']})")
        if r["recent_exits"]:
            flags.append(f"EXIT ({r['recent_exits']})")
        log(f"{i:2d}. {r['name']:<35} conviction={r['conviction_count']}/{len(fund_holdings)} "
            f"avg_wt={r['avg_weight_pct']}%  {' | '.join(flags)}")

    out_path = write_csv(rows, list(fund_holdings.keys()), MONTH_STR)
    log(f"\nSaved: {out_path}")
    if failures:
        log(f"NOTE: {', '.join(failures)} failed this run -- list above is based on "
            f"the funds that succeeded only.")


if __name__ == "__main__":
    main()
