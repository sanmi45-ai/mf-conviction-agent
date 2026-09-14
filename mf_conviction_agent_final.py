#!/usr/bin/env python3
"""
mf_conviction_agent_final.py

FINAL COMBINED SYSTEM.

Auto-scraped funds (direct from each AMC's own site via Playwright --
these show a clear "as on" date inside the file, unlike aggregator sites):
    - Nippon India Small Cap Fund   (confirmed working)
    - quant Small Cap Fund          (confirmed working after selector fix)
    - Invesco India Smallcap Fund   (best-effort, needs one calibration
      round -- see debug output if it fails)

Manual fund:
    - Bandhan Small Cap Fund -- its site actively blocks automated
      browsers, so you upload the .xls/.xlsx yourself each month into
      manual_downloads/. This script detects the file automatically.

MODES
    --mode=monthly   (default, run via cron on the 15th)
        Scrapes the 3 auto funds, saves snapshots, sends the main report
        email. If a not-yet-processed Bandhan file is sitting in
        manual_downloads/, includes it too. If not, sends a separate
        reminder email asking you to upload it.

    --mode=bandhan   (run automatically when a file is pushed to
                       manual_downloads/, via a separate workflow trigger)
        Loads this month's already-scraped auto-fund data (scraping fresh
        if not already done this month), parses the new Bandhan file, and
        sends an updated report with all 4 funds included.

REQUIRED GITHUB SECRETS:
    SMTP_USERNAME, SMTP_APP_PASSWORD, EMAIL_TO
"""

import argparse
import datetime
import json
import os
import re
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SNAPSHOT_DIR = os.path.join(BASE_DIR, "snapshots")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
RAW_DIR = os.path.join(BASE_DIR, "raw_downloads")
DEBUG_DIR = os.path.join(BASE_DIR, "debug_artifacts")
MANUAL_DIR = os.path.join(BASE_DIR, "manual_downloads")
MARKER_DIR = os.path.join(BASE_DIR, "bandhan_processed_markers")

MONTH_STR = datetime.date.today().strftime("%Y-%m")
BANDHAN_LABEL = "Bandhan Small Cap Fund"


def log(msg):
    print(msg)


def save_debug(page, name):
    os.makedirs(DEBUG_DIR, exist_ok=True)
    png_path = os.path.join(DEBUG_DIR, f"{name}.png")
    html_path = os.path.join(DEBUG_DIR, f"{name}.html")
    try:
        page.screenshot(path=png_path, full_page=True)
    except Exception as e:
        log(f"    (couldn't save screenshot: {e})")
    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(page.content())
    except Exception as e:
        log(f"    (couldn't save html: {e})")
    log(f"    -> saved debug artifacts: {png_path} , {html_path}")


def download_via_context(context, url, dest_path):
    resp = context.request.get(url)
    if resp.status != 200:
        raise RuntimeError(f"Download failed ({resp.status}) for {url}")
    with open(dest_path, "wb") as f:
        f.write(resp.body())
    return dest_path


# ---------------------------------------------------------------------------
# Excel parsing (content-based engine detection, ISIN-based extraction)
# ---------------------------------------------------------------------------
def detect_excel_engine(raw_path):
    with open(raw_path, "rb") as f:
        header = f.read(8)
    if header.startswith(b"PK\x03\x04"):
        return "openpyxl"
    if header.startswith(b"\xd0\xcf\x11\xe0"):
        return "xlrd"
    return "xlrd" if raw_path.lower().endswith(".xls") else "openpyxl"


def find_col(df, *keyword_sets):
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
        log(f"    [debug:{label}] columns: {list(df.columns)}")
        log(f"    [debug:{label}] resolved -> isin: {isin_col}, name: {name_col}, weight: {weight_col}")
    if not isin_col or not name_col:
        return None
    holdings = []
    for _, row in df.iterrows():
        isin, name = row.get(isin_col), row.get(name_col)
        weight = row.get(weight_col) if weight_col else None
        if pd.isna(isin) or pd.isna(name):
            continue
        isin_s = str(isin).strip()
        if not re.match(r"^IN[A-Z0-9]{10}$", isin_s):
            continue
        try:
            weight_f = float(weight) if weight is not None and not pd.isna(weight) else 0.0
        except (TypeError, ValueError):
            weight_f = 0.0
        holdings.append({"isin": isin_s, "name": str(name).strip(), "weight": weight_f})
    return holdings


def find_scheme_sheet_or_rows(raw_path, match_terms, debug=False, label=""):
    engine = detect_excel_engine(raw_path)
    if debug:
        log(f"    [debug:{label}] using engine '{engine}'")
    xls = pd.ExcelFile(raw_path, engine=engine)
    terms = [t.lower() for t in match_terms]
    if debug:
        log(f"    [debug:{label}] sheets: {xls.sheet_names[:15]}")
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
# AUTO-SCRAPED SOURCE 1: Nippon India (confirmed working)
# ---------------------------------------------------------------------------
def fetch_nippon_india(context, debug=False):
    log("  [Nippon India] loading page...")
    page = context.new_page()
    page.goto("https://mf.nipponindiaim.com/investor-service/downloads/factsheet-portfolio-and-other-disclosures",
               wait_until="networkidle", timeout=45000)
    links = page.eval_on_selector_all(
        "a[href$='.xls'], a[href$='.xlsx']",
        "els => els.map(e => ({href: e.href, text: e.innerText}))"
    )
    portfolio_links = [l for l in links if "portfolio" in l["text"].lower()] or \
                       [l for l in links if "portfolio" in l["href"].lower()]
    if not portfolio_links:
        save_debug(page, "nippon_fail")
        page.close()
        raise RuntimeError("No Monthly Portfolio link found for Nippon India.")
    file_url = portfolio_links[0]["href"]
    log(f"    -> {file_url}")
    os.makedirs(RAW_DIR, exist_ok=True)
    raw_path = os.path.join(RAW_DIR, f"nippon_{MONTH_STR}.xls")
    download_via_context(context, file_url, raw_path)
    page.close()
    df = find_scheme_sheet_or_rows(raw_path, ["small cap"], debug=debug, label="nippon")
    if df is None:
        raise RuntimeError("Couldn't find 'Small Cap' rows in Nippon India file.")
    holdings = extract_holdings_from_df(df, debug=debug, label="nippon")
    if holdings is None:
        raise RuntimeError("Couldn't identify ISIN/name columns for Nippon India.")
    return holdings


# ---------------------------------------------------------------------------
# AUTO-SCRAPED SOURCE 2: quant
# ---------------------------------------------------------------------------
def fetch_quant(context, debug=False):
    log("  [quant] loading page...")
    page = context.new_page()
    page.goto("https://quantmutual.com/statutory-disclosures", wait_until="networkidle", timeout=45000)
    try:
        page.click("[id='2026']", timeout=10000)
    except Exception as e:
        save_debug(page, "quant_fail_click")
        page.close()
        raise RuntimeError(f"Couldn't click the year tab: {e}")
    try:
        page.wait_for_function(
            "document.querySelector(\"div[id='MONTHLY PORTFOLIO - FUND - WISE']\")?.innerHTML.length > 50",
            timeout=15000
        )
    except Exception:
        pass
    links = page.eval_on_selector_all(
        "div[id='MONTHLY PORTFOLIO - FUND - WISE'] a",
        "els => els.map(e => ({href: e.href, text: e.innerText}))"
    )
    scored = [l for l in links if "small" in l["text"].lower() or "small" in l["href"].lower()] or links
    if not scored:
        save_debug(page, "quant_fail_empty")
        page.close()
        raise RuntimeError("No links found after clicking the year tab for quant.")
    file_url = scored[0]["href"]
    log(f"    -> {file_url}")
    os.makedirs(RAW_DIR, exist_ok=True)
    raw_path = os.path.join(RAW_DIR, f"quant_{MONTH_STR}{'.xls' if file_url.endswith('.xls') else '.xlsx'}")
    download_via_context(context, file_url, raw_path)
    page.close()
    df = find_scheme_sheet_or_rows(raw_path, ["small cap"], debug=debug, label="quant")
    if df is None:
        raise RuntimeError("Couldn't find 'Small Cap' rows in quant file.")
    holdings = extract_holdings_from_df(df, debug=debug, label="quant")
    if holdings is None:
        raise RuntimeError("Couldn't identify ISIN/name columns for quant.")
    return holdings


# ---------------------------------------------------------------------------
# AUTO-SCRAPED SOURCE 3: Invesco (best-effort, first calibration round)
# ---------------------------------------------------------------------------
def fetch_invesco(context, debug=False):
    log("  [Invesco] loading page...")
    page = context.new_page()
    page.goto("https://www.invescomutualfund.com/literature-forms/monthly-holdings",
               wait_until="networkidle", timeout=45000)
    page.wait_for_timeout(2000)

    # We don't yet know the exact filter UI. Try the generic approach first:
    # look for any link already on the page mentioning "small cap".
    links = page.eval_on_selector_all(
        "a[href$='.xls'], a[href$='.xlsx']",
        "els => els.map(e => ({href: e.href, text: e.innerText}))"
    )
    scored = [l for l in links if "small" in l["text"].lower() or "small" in l["href"].lower()]

    if not scored:
        # Dump every interactive control so we can see the real filter UI
        # structure and calibrate this in one more round if needed.
        if debug:
            selects = page.eval_on_selector_all(
                "select", "els => els.map(e => ({name: e.name || e.id, options: Array.from(e.options).map(o => o.text)}))"
            )
            inputs = page.eval_on_selector_all(
                "input", "els => els.map(e => ({type: e.type, placeholder: e.placeholder, name: e.name || e.id}))"
            )
            buttons = page.eval_on_selector_all(
                "button, [role='button']", "els => els.map(e => e.innerText.trim()).filter(t => t)"
            )
            log(f"    [debug:invesco] {len(links)} xls/xlsx links on page (none matched 'small cap')")
            log(f"    [debug:invesco] <select> elements: {selects}")
            log(f"    [debug:invesco] <input> elements: {inputs[:15]}")
            log(f"    [debug:invesco] buttons/clickable: {buttons[:20]}")
        save_debug(page, "invesco_fail_no_link")
        page.close()
        raise RuntimeError("No 'small cap' file link found on Invesco's page (may need a scheme filter "
                            "selection first). Debug artifacts + interactive element dump saved -- "
                            "share these so the filter interaction can be added.")

    file_url = scored[0]["href"]
    log(f"    -> {file_url}")
    os.makedirs(RAW_DIR, exist_ok=True)
    raw_path = os.path.join(RAW_DIR, f"invesco_{MONTH_STR}{'.xls' if file_url.endswith('.xls') else '.xlsx'}")
    download_via_context(context, file_url, raw_path)
    page.close()
    df = find_scheme_sheet_or_rows(raw_path, ["small cap"], debug=debug, label="invesco")
    if df is None:
        raise RuntimeError("Couldn't find 'Small Cap' rows in Invesco file.")
    holdings = extract_holdings_from_df(df, debug=debug, label="invesco")
    if holdings is None:
        raise RuntimeError("Couldn't identify ISIN/name columns for Invesco.")
    return holdings


AUTO_FUNDS = {
    "Nippon India Small Cap Fund": fetch_nippon_india,
    "quant Small Cap Fund": fetch_quant,
    "Invesco India Smallcap Fund": fetch_invesco,
}


# ---------------------------------------------------------------------------
# Bandhan: manual file detection
# ---------------------------------------------------------------------------
def find_manual_bandhan_file():
    os.makedirs(MANUAL_DIR, exist_ok=True)
    candidates = [
        f for f in os.listdir(MANUAL_DIR)
        if f.lower().startswith("bandhan") and (f.endswith(".xls") or f.endswith(".xlsx"))
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda f: os.path.getmtime(os.path.join(MANUAL_DIR, f)), reverse=True)
    return os.path.join(MANUAL_DIR, candidates[0])


def bandhan_already_processed_this_month():
    os.makedirs(MARKER_DIR, exist_ok=True)
    return os.path.exists(os.path.join(MARKER_DIR, f"{MONTH_STR}.done"))


def mark_bandhan_processed():
    os.makedirs(MARKER_DIR, exist_ok=True)
    with open(os.path.join(MARKER_DIR, f"{MONTH_STR}.done"), "w") as f:
        f.write(datetime.datetime.now().isoformat())


def fetch_bandhan_from_manual_file(debug=False):
    path = find_manual_bandhan_file()
    if path is None:
        return None
    log(f"  [Bandhan] using manually uploaded file: {os.path.basename(path)}")
    df = find_scheme_sheet_or_rows(path, ["small cap"], debug=debug, label="bandhan")
    if df is None:
        raise RuntimeError(f"Couldn't find 'Small Cap' rows in {path}.")
    holdings = extract_holdings_from_df(df, debug=debug, label="bandhan")
    if holdings is None:
        raise RuntimeError("Couldn't identify ISIN/name columns in the Bandhan file.")
    return holdings


# ---------------------------------------------------------------------------
# Snapshot storage
# ---------------------------------------------------------------------------
def _safe(label):
    return label.replace(" ", "_")


def save_snapshot(label, month_str, holdings):
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    with open(os.path.join(SNAPSHOT_DIR, f"{_safe(label)}_{month_str}.json"), "w") as f:
        json.dump(holdings, f, indent=2)


def load_snapshot(label, month_str):
    path = os.path.join(SNAPSHOT_DIR, f"{_safe(label)}_{month_str}.json")
    if not os.path.exists(path):
        return None
    with open(path) as f:
        return json.load(f)


def available_months(label, before_month_str):
    if not os.path.isdir(SNAPSHOT_DIR):
        return []
    prefix = f"{_safe(label)}_"
    months = [
        f[len(prefix):-len(".json")]
        for f in os.listdir(SNAPSHOT_DIR)
        if f.startswith(prefix) and f.endswith(".json")
        and f[len(prefix):-len(".json")] < before_month_str
    ]
    return sorted(months, reverse=True)


# ---------------------------------------------------------------------------
# Analyses (all keyed by ISIN)
# ---------------------------------------------------------------------------
def analyze_entries(fund_holdings, fund_prev):
    result = {}
    for label, holdings in fund_holdings.items():
        prev = fund_prev.get(label)
        if prev is None:
            result[label] = None
            continue
        prev_isins = {h["isin"] for h in prev}
        new = [h for h in holdings if h["isin"] not in prev_isins]
        result[label] = sorted(new, key=lambda h: -h["weight"])
    return result


def analyze_multi_fund_buys(fund_holdings):
    stock_funds = {}
    for label, holdings in fund_holdings.items():
        for h in holdings:
            e = stock_funds.setdefault(h["isin"], {"name": h["name"], "funds": {}})
            e["funds"][label] = h["weight"]
            e["name"] = h["name"]
    multi = [
        {"name": v["name"], "funds": v["funds"], "count": len(v["funds"])}
        for v in stock_funds.values() if len(v["funds"]) >= 2
    ]
    multi.sort(key=lambda x: (x["count"], sum(x["funds"].values()) / x["count"]), reverse=True)
    return multi


def analyze_stake_increases_3m(current, snapshot_3m):
    if snapshot_3m is None:
        return None
    old_weights = {h["isin"]: h["weight"] for h in snapshot_3m}
    increases = []
    for h in current:
        old_w = old_weights.get(h["isin"])
        if old_w is not None and h["weight"] > old_w:
            increases.append({"name": h["name"], "old_weight": old_w, "new_weight": h["weight"],
                               "change": round(h["weight"] - old_w, 2)})
    increases.sort(key=lambda x: -x["change"])
    return increases


def analyze_sells(current, prev, threshold_pct=10.0):
    if prev is None:
        return None, None
    current_map = {h["isin"]: h for h in current}
    prev_map = {h["isin"]: h for h in prev}
    complete = [h for k, h in prev_map.items() if k not in current_map]
    partial = []
    for k, old_h in prev_map.items():
        new_h = current_map.get(k)
        if new_h is None:
            continue
        if old_h["weight"] > 0 and (old_h["weight"] - new_h["weight"]) / old_h["weight"] * 100 >= threshold_pct:
            partial.append({"name": new_h["name"], "old_weight": old_h["weight"], "new_weight": new_h["weight"],
                             "pct_reduced": round((old_h["weight"] - new_h["weight"]) / old_h["weight"] * 100, 1)})
    partial.sort(key=lambda x: -x["pct_reduced"])
    return partial, complete


def build_conviction_list(fund_holdings, top_n=20):
    stocks = {}
    for label, holdings in fund_holdings.items():
        for h in holdings:
            e = stocks.setdefault(h["isin"], {"name": h["name"], "funds": {}})
            e["funds"][label] = h["weight"]
    rows = []
    for isin, info in stocks.items():
        count = len(info["funds"])
        avg_w = sum(info["funds"].values()) / count
        rows.append({"name": info["name"], "count": count, "avg_weight": round(avg_w, 2), "funds": info["funds"]})
    rows.sort(key=lambda r: (r["count"], r["avg_weight"]), reverse=True)
    return rows[:top_n]


# ---------------------------------------------------------------------------
# Report + email
# ---------------------------------------------------------------------------
def html_table(headers, rows):
    if not rows:
        return "<p><em>No data.</em></p>"
    h = "<tr>" + "".join(f"<th style='text-align:left;padding:4px 8px;border-bottom:1px solid #ccc'>{c}</th>" for c in headers) + "</tr>"
    body = ""
    for row in rows:
        body += "<tr>" + "".join(f"<td style='padding:4px 8px;border-bottom:1px solid #eee'>{c}</td>" for c in row) + "</tr>"
    return f"<table style='border-collapse:collapse;font-family:sans-serif;font-size:14px'>{h}{body}</table>"


def build_report_html(month_str, fund_holdings, entries, multi_buys, stake_increases,
                       sells, conviction_list, failures, bandhan_included):
    parts = [f"<h2>Mutual Fund Conviction Report — {month_str}</h2>"]
    parts.append(f"<p><b>Bandhan Small Cap:</b> "
                  f"{'included (manually uploaded file found)' if bandhan_included else 'NOT included yet -- see reminder below'}</p>")

    if failures:
        parts.append(f"<p style='color:#b00'><b>Note:</b> data could not be fetched for: {', '.join(failures)}.</p>")

    parts.append("<h3>1. Current holdings count</h3>")
    parts.append(html_table(["Fund", "# Holdings"], [[l, len(h)] for l, h in fund_holdings.items()]))

    parts.append("<h3>2. New stocks bought this month</h3>")
    for label, new in entries.items():
        parts.append(f"<p><b>{label}</b></p>")
        if new is None:
            parts.append("<p><em>No prior month data yet.</em></p>")
        else:
            parts.append(html_table(["Stock", "Weight %"], [[h["name"], h["weight"]] for h in new]) if new
                          else "<p><em>No new entries.</em></p>")

    parts.append("<h3>3. Stocks bought by multiple funds (conviction)</h3>")
    rows = [[m["name"], m["count"], ", ".join(f"{k} ({v}%)" for k, v in m["funds"].items())] for m in multi_buys[:20]]
    parts.append(html_table(["Stock", "# Funds", "Weight by fund"], rows))

    parts.append("<h3>4. Stake increases over the last ~3 months</h3>")
    for label, inc in stake_increases.items():
        parts.append(f"<p><b>{label}</b></p>")
        if inc is None:
            parts.append("<p><em>Not enough historical data yet.</em></p>")
        elif inc:
            parts.append(html_table(["Stock", "Weight then", "Weight now", "Change"],
                                     [[h["name"], h["old_weight"], h["new_weight"], f"+{h['change']}"] for h in inc[:15]]))
        else:
            parts.append("<p><em>No notable increases.</em></p>")

    parts.append("<h3>5. Partial and complete sells vs last month</h3>")
    for label, (partial, complete) in sells.items():
        parts.append(f"<p><b>{label}</b></p>")
        if partial is None:
            parts.append("<p><em>No prior month data yet.</em></p>")
            continue
        parts.append("<p>Partial sells:</p>")
        parts.append(html_table(["Stock", "Weight then", "Weight now", "% Reduced"],
                                 [[h["name"], h["old_weight"], h["new_weight"], h["pct_reduced"]] for h in partial]) if partial
                      else "<p><em>None.</em></p>")
        parts.append("<p>Complete sells:</p>")
        parts.append(html_table(["Stock", "Last known weight"], [[h["name"], h["weight"]] for h in complete]) if complete
                      else "<p><em>None.</em></p>")

    parts.append("<h3>Top 20 conviction list (overall)</h3>")
    rows = [[i, r["name"], r["count"], r["avg_weight"], ", ".join(r["funds"].keys())]
            for i, r in enumerate(conviction_list, 1)]
    parts.append(html_table(["Rank", "Stock", "# Funds", "Avg Weight %", "Held by"], rows))

    return "".join(parts)


def send_email(subject, html_body):
    smtp_user = os.environ.get("SMTP_USERNAME")
    smtp_pass = os.environ.get("SMTP_APP_PASSWORD")
    email_to = os.environ.get("EMAIL_TO")
    if not all([smtp_user, smtp_pass, email_to]):
        log("  ! Email not sent: SMTP secrets not set.")
        return
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = email_to
    msg.attach(MIMEText(html_body, "html"))
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, email_to, msg.as_string())
    log(f"  -> email sent to {email_to}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def scrape_auto_funds(debug=False):
    fund_holdings, failures = {}, []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ))
        for label, fetch_fn in AUTO_FUNDS.items():
            log(f"\n{label}:")
            try:
                holdings = fetch_fn(context, debug=debug)
                log(f"  -> {len(holdings)} holdings")
                fund_holdings[label] = holdings
                save_snapshot(label, MONTH_STR, holdings)
            except Exception as e:
                log(f"  ! FAILED: {e}")
                failures.append(label)
        browser.close()
    return fund_holdings, failures


def run_analysis_and_report(fund_holdings, failures, bandhan_included):
    fund_prev = {label: load_snapshot(label, available_months(label, MONTH_STR)[0])
                 if available_months(label, MONTH_STR) else None for label in fund_holdings}
    entries = analyze_entries(fund_holdings, fund_prev)

    stake_increases = {}
    for label, holdings in fund_holdings.items():
        months = available_months(label, MONTH_STR)
        snap_3m = load_snapshot(label, months[min(2, len(months) - 1)]) if months else None
        stake_increases[label] = analyze_stake_increases_3m(holdings, snap_3m)

    sells = {label: analyze_sells(holdings, fund_prev.get(label)) for label, holdings in fund_holdings.items()}
    multi_buys = analyze_multi_fund_buys(fund_holdings)
    conviction_list = build_conviction_list(fund_holdings, top_n=20)

    html_report = build_report_html(MONTH_STR, fund_holdings, entries, multi_buys,
                                     stake_increases, sells, conviction_list, failures, bandhan_included)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    suffix = "with_bandhan" if bandhan_included else "auto_only"
    report_path = os.path.join(OUTPUT_DIR, f"report_{MONTH_STR}_{suffix}.html")
    with open(report_path, "w") as f:
        f.write(html_report)
    log(f"\nSaved report: {report_path}")
    return html_report


def send_bandhan_reminder():
    html = (
        "<h2>Reminder: upload this month's Bandhan Small Cap portfolio</h2>"
        "<p>The auto-scraped report has been sent, but Bandhan's site can't be scraped automatically.</p>"
        "<p>Please download this month's file from:</p>"
        "<p><a href='https://bandhanmutual.com/mutual-funds/equity-funds/bandhan-small-cap-fund/direct#Downloads'>"
        "Bandhan Small Cap Fund — Downloads page</a> (tap 'Latest Portfolio')</p>"
        "<p>Then upload it to the <code>manual_downloads/</code> folder in the GitHub repo "
        "(filename should start with 'bandhan', e.g. <code>bandhan_2026-09.xlsx</code>).</p>"
        "<p>As soon as it's uploaded, an updated report including Bandhan will be sent automatically.</p>"
    )
    send_email(f"Reminder: upload Bandhan portfolio — {MONTH_STR}", html)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["monthly", "bandhan"], default="monthly")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--no-email", action="store_true")
    args = parser.parse_args()

    if args.mode == "monthly":
        fund_holdings, failures = scrape_auto_funds(debug=args.debug)
        if not fund_holdings:
            log("\nAll auto sources failed. No report generated.")
            sys.exit(1)

        bandhan_holdings = None
        if not bandhan_already_processed_this_month():
            try:
                bandhan_holdings = fetch_bandhan_from_manual_file(debug=args.debug)
            except Exception as e:
                log(f"  ! Bandhan file present but failed to parse: {e}")

        if bandhan_holdings:
            fund_holdings[BANDHAN_LABEL] = bandhan_holdings
            save_snapshot(BANDHAN_LABEL, MONTH_STR, bandhan_holdings)
            mark_bandhan_processed()

        html_report = run_analysis_and_report(fund_holdings, failures, bandhan_included=bandhan_holdings is not None)
        if not args.no_email:
            send_email(f"Mutual Fund Conviction Report — {MONTH_STR}", html_report)
            if bandhan_holdings is None:
                send_bandhan_reminder()
        else:
            log("  (--no-email set, skipping email)")

    elif args.mode == "bandhan":
        # Reuse this month's auto-fund snapshots if already scraped; else scrape fresh.
        fund_holdings, failures = {}, []
        for label in AUTO_FUNDS:
            snap = load_snapshot(label, MONTH_STR)
            if snap is not None:
                fund_holdings[label] = snap
        if not fund_holdings:
            log("No cached auto-fund data for this month yet -- scraping fresh.")
            fund_holdings, failures = scrape_auto_funds(debug=args.debug)

        bandhan_holdings = fetch_bandhan_from_manual_file(debug=args.debug)
        if bandhan_holdings is None:
            log("No Bandhan file found in manual_downloads/. Nothing to do.")
            sys.exit(0)

        fund_holdings[BANDHAN_LABEL] = bandhan_holdings
        save_snapshot(BANDHAN_LABEL, MONTH_STR, bandhan_holdings)
        mark_bandhan_processed()

        html_report = run_analysis_and_report(fund_holdings, failures, bandhan_included=True)
        if not args.no_email:
            send_email(f"Updated Mutual Fund Conviction Report (with Bandhan) — {MONTH_STR}", html_report)
        else:
            log("  (--no-email set, skipping email)")


if __name__ == "__main__":
    main()
