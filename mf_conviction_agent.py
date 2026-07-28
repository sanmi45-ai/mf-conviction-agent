#!/usr/bin/env python3
"""
mf_conviction_agent.py

Tracks holdings across a chosen set of Indian mutual funds, and produces a
ranked "conviction list" of stocks based on:
  - how many of the tracked funds hold the stock (conviction count)
  - average weight (% of NAV) across the funds that hold it
  - whether the stock was recently ADDED or EXITED by any fund since the
    last time this script was run

Data source: https://mfdata.in (free, no-auth, community-run API that
aggregates AMFI's mandatory monthly portfolio disclosures).

USAGE
    python mf_conviction_agent.py                 # normal monthly run
    python mf_conviction_agent.py --debug          # print raw API responses
    python mf_conviction_agent.py --top 30         # change list size (default 20)

OUTPUT
    output/conviction_list_<YYYY-MM>.csv   <- the ranked list you asked for
    snapshots/<family_id>_<YYYY-MM>.json   <- raw holdings snapshot (kept so
                                               next month's run can diff
                                               against it for entries/exits)

SCHEDULING
    Run this once a month, a few days after the 10th (AMFI's disclosure
    deadline). See the bottom of this file / the README for cron and
    GitHub Actions examples.
"""

import argparse
import datetime
import json
import os
import sys
import time

import requests

BASE_URL = "https://mfdata.in/api/v1"
SNAPSHOT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "snapshots")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
FAMILY_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "family_id_cache.json")

# ---------------------------------------------------------------------------
# 1. THE FUNDS YOU CARE ABOUT
#    Edit this list to track different / additional funds.
#    "query" is what gets sent to mfdata.in's search endpoint to find the
#    scheme family. Keep it close to the official scheme name for a clean
#    match.
# ---------------------------------------------------------------------------
FUNDS = [
    {"label": "Bandhan Small Cap Fund", "query": "Bandhan Small Cap Fund"},
    {"label": "Nippon India Small Cap Fund", "query": "Nippon India Small Cap Fund"},
    {"label": "quant Small Cap Fund", "query": "quant Small Cap Fund"},
]

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "mf-conviction-agent/1.0 (personal research script)"})


def api_get(path, params=None, debug=False):
    url = f"{BASE_URL}{path}"
    for attempt in range(3):
        try:
            resp = SESSION.get(url, params=params, timeout=20)
            if debug:
                print(f"[debug] GET {resp.url} -> {resp.status_code}")
            resp.raise_for_status()
            data = resp.json()
            if debug:
                print(json.dumps(data, indent=2)[:2000])
            return data
        except requests.exceptions.RequestException as e:
            print(f"  ! request failed (attempt {attempt + 1}/3): {e}")
            time.sleep(2)
    raise RuntimeError(f"Could not reach {url} after 3 attempts")


# ---------------------------------------------------------------------------
# 2. RESOLVE FUND NAME -> family_id (cached locally so we don't re-search
#    every month)
# ---------------------------------------------------------------------------
def load_family_cache():
    if os.path.exists(FAMILY_CACHE_FILE):
        with open(FAMILY_CACHE_FILE) as f:
            return json.load(f)
    return {}


def save_family_cache(cache):
    with open(FAMILY_CACHE_FILE, "w") as f:
        json.dump(cache, f, indent=2)


def resolve_family_id(fund, cache, debug=False):
    if fund["query"] in cache:
        return cache[fund["query"]]

    print(f"  Searching for: {fund['query']}")
    data = api_get("/search", params={"q": fund["query"]}, debug=debug)
    results = data.get("data") or data.get("results") or []

    if not results:
        raise RuntimeError(
            f"No search results for '{fund['query']}'. Try adjusting the "
            f"query text in the FUNDS list, or run with --debug to inspect "
            f"the raw response."
        )

    # Try to find the best match: prefer a result whose name contains all
    # key words from the query (case-insensitive).
    query_words = [w.lower() for w in fund["query"].split()]
    best = None
    for r in results:
        name = (r.get("name") or r.get("scheme_name") or "").lower()
        if all(w in name for w in query_words):
            best = r
            break
    if best is None:
        best = results[0]  # fall back to top search hit

    family_id = best.get("family_id") or best.get("familyId") or best.get("id")
    resolved_name = best.get("name") or best.get("scheme_name")

    if family_id is None:
        raise RuntimeError(
            f"Found a result for '{fund['query']}' but couldn't extract a "
            f"family_id from it. Raw result: {json.dumps(best)}. "
            f"Run with --debug and adjust the key names in resolve_family_id()."
        )

    print(f"    -> matched '{resolved_name}' (family_id={family_id})")
    cache[fund["query"]] = {"family_id": family_id, "resolved_name": resolved_name}
    return cache[fund["query"]]


# ---------------------------------------------------------------------------
# 3. FETCH HOLDINGS FOR A FAMILY
# ---------------------------------------------------------------------------
def fetch_holdings(family_id, debug=False):
    data = api_get(f"/families/{family_id}/holdings", debug=debug)
    holdings = data.get("data") or data.get("holdings") or []
    if isinstance(holdings, dict):
        # some responses might nest the list under e.g. data["equity"]
        holdings = holdings.get("equity") or holdings.get("holdings") or []

    normalized = []
    for h in holdings:
        isin = h.get("isin") or h.get("ISIN")
        name = h.get("name") or h.get("company") or h.get("instrument") or h.get("instrument_name")
        weight = (
            h.get("percentage_of_nav")
            or h.get("percent_nav")
            or h.get("weight")
            or h.get("nav_percentage")
            or h.get("pct_nav")
        )
        if isin is None or name is None:
            continue
        try:
            weight = float(weight) if weight is not None else 0.0
        except (TypeError, ValueError):
            weight = 0.0
        normalized.append({"isin": isin, "name": name, "weight": weight})
    return normalized


# ---------------------------------------------------------------------------
# 4. SNAPSHOT STORAGE (for month-over-month entry/exit detection)
# ---------------------------------------------------------------------------
def snapshot_path(family_id, month_str):
    return os.path.join(SNAPSHOT_DIR, f"{family_id}_{month_str}.json")


def save_snapshot(family_id, month_str, holdings):
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    with open(snapshot_path(family_id, month_str), "w") as f:
        json.dump(holdings, f, indent=2)


def load_previous_snapshot(family_id, current_month_str):
    """Find the most recent snapshot for this family that is BEFORE the
    current month, so we can diff against it."""
    if not os.path.isdir(SNAPSHOT_DIR):
        return None
    candidates = []
    prefix = f"{family_id}_"
    for fname in os.listdir(SNAPSHOT_DIR):
        if fname.startswith(prefix) and fname.endswith(".json"):
            month_part = fname[len(prefix):-len(".json")]
            if month_part < current_month_str:
                candidates.append(month_part)
    if not candidates:
        return None
    latest_prev = max(candidates)
    with open(snapshot_path(family_id, latest_prev)) as f:
        return {"month": latest_prev, "holdings": json.load(f)}


# ---------------------------------------------------------------------------
# 5. MAIN ANALYSIS
# ---------------------------------------------------------------------------
def build_conviction_list(fund_holdings, fund_prev_holdings, top_n):
    """
    fund_holdings: {fund_label: [ {isin, name, weight}, ... ]}
    fund_prev_holdings: {fund_label: [ {isin, name, weight}, ... ] or None}
    """
    # stock_key -> {"name":..., "funds": {label: weight}, "entries": [...], "exits": [...]}
    stocks = {}

    for label, holdings in fund_holdings.items():
        current_isins = {h["isin"] for h in holdings}
        prev = fund_prev_holdings.get(label)
        prev_isins = {h["isin"] for h in prev} if prev else set()

        for h in holdings:
            key = h["isin"]
            entry = stocks.setdefault(key, {"name": h["name"], "funds": {}, "entries": [], "exits": []})
            entry["funds"][label] = h["weight"]
            entry["name"] = h["name"]  # keep freshest name spelling
            if prev is not None and key not in prev_isins:
                entry["entries"].append(label)

        if prev is not None:
            for h in prev:
                if h["isin"] not in current_isins:
                    key = h["isin"]
                    entry = stocks.setdefault(key, {"name": h["name"], "funds": {}, "entries": [], "exits": []})
                    entry["exits"].append(label)

    all_fund_labels = list(fund_holdings.keys())
    rows = []
    for isin, info in stocks.items():
        conviction_count = len(info["funds"])
        avg_weight = sum(info["funds"].values()) / conviction_count if conviction_count else 0.0
        rows.append({
            "isin": isin,
            "name": info["name"],
            "conviction_count": conviction_count,
            "avg_weight_pct": round(avg_weight, 2),
            "held_by": ", ".join(info["funds"].keys()),
            "weights": info["funds"],
            "recent_entries": ", ".join(info["entries"]),
            "recent_exits": ", ".join(info["exits"]),
        })

    # Rank: conviction count first, then average weight
    rows.sort(key=lambda r: (r["conviction_count"], r["avg_weight_pct"]), reverse=True)
    return rows[:top_n], all_fund_labels


def write_csv(rows, all_fund_labels, month_str):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f"conviction_list_{month_str}.csv")
    import csv
    with open(out_path, "w", newline="") as f:
        writer = csv.writer(f)
        header = ["rank", "stock", "isin", "conviction_count", "avg_weight_pct"]
        header += [f"weight_in::{lbl}" for lbl in all_fund_labels]
        header += ["recent_entries", "recent_exits"]
        writer.writerow(header)
        for i, r in enumerate(rows, 1):
            row = [i, r["name"], r["isin"], r["conviction_count"], r["avg_weight_pct"]]
            row += [r["weights"].get(lbl, "") for lbl in all_fund_labels]
            row += [r["recent_entries"], r["recent_exits"]]
            writer.writerow(row)
    return out_path


def print_summary(rows, month_str):
    print(f"\n=== Top {len(rows)} conviction stocks — {month_str} ===\n")
    for i, r in enumerate(rows, 1):
        flags = []
        if r["recent_entries"]:
            flags.append(f"NEW ENTRY ({r['recent_entries']})")
        if r["recent_exits"]:
            flags.append(f"EXIT ({r['recent_exits']})")
        flag_str = f"  [{' | '.join(flags)}]" if flags else ""
        print(f"{i:2d}. {r['name']:<35} conviction={r['conviction_count']}/3  "
              f"avg_wt={r['avg_weight_pct']}%{flag_str}")


def main():
    parser = argparse.ArgumentParser(description="Indian mutual fund conviction-list agent")
    parser.add_argument("--top", type=int, default=20, help="number of stocks in the final list")
    parser.add_argument("--debug", action="store_true", help="print raw API responses")
    args = parser.parse_args()

    month_str = datetime.date.today().strftime("%Y-%m")

    print(f"Run month: {month_str}")
    print("Resolving fund -> family_id mappings...")
    cache = load_family_cache()
    fund_family = {}
    for fund in FUNDS:
        try:
            fund_family[fund["label"]] = resolve_family_id(fund, cache, debug=args.debug)
        except RuntimeError as e:
            print(f"  ERROR resolving '{fund['label']}': {e}")
            sys.exit(1)
    save_family_cache(cache)

    print("\nFetching current holdings...")
    fund_holdings = {}
    fund_prev_holdings = {}
    for fund in FUNDS:
        label = fund["label"]
        family_id = fund_family[label]["family_id"]
        print(f"  {label} (family_id={family_id})")
        holdings = fetch_holdings(family_id, debug=args.debug)
        print(f"    -> {len(holdings)} holdings fetched")
        fund_holdings[label] = holdings

        prev = load_previous_snapshot(family_id, month_str)
        fund_prev_holdings[label] = prev["holdings"] if prev else None
        if prev:
            print(f"    -> comparing against previous snapshot from {prev['month']}")
        else:
            print("    -> no previous snapshot found (first run for this fund; "
                  "entries/exits will show up starting next month)")

        save_snapshot(family_id, month_str, holdings)

    print("\nBuilding conviction list...")
    rows, all_fund_labels = build_conviction_list(fund_holdings, fund_prev_holdings, args.top)

    if not rows:
        print("No holdings data was assembled — check the API responses with --debug.")
        sys.exit(1)

    print_summary(rows, month_str)
    out_path = write_csv(rows, all_fund_labels, month_str)
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
