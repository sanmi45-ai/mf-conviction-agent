# Mutual Fund Conviction Agent — Final Version

Tracks 4 funds: Nippon India Small Cap, quant Small Cap, Invesco India
Smallcap (all auto-scraped from their own AMC sites), and Bandhan Small
Cap (manual upload, since its site blocks automated browsers).

## Setup from scratch

### 1. Create the repo
github.com/new → name it (e.g. `mf-conviction-agent`) → Private → Create.

### 2. Add the files
- `mf_conviction_agent_final.py`
- `requirements.txt`
- `.github/workflows/monthly.yml`
- `.github/workflows/bandhan_trigger.yml`

(Add file → Create new file → paste contents → Commit, for each one.)

### 3. Gmail App Password
1. myaccount.google.com/apppasswords (needs 2-Step Verification on)
2. Create one, copy the 16-character password

### 4. GitHub repo secrets
Settings → Secrets and variables → Actions → New repository secret:
| Name | Value |
|---|---|
| `SMTP_USERNAME` | your Gmail address |
| `SMTP_APP_PASSWORD` | the 16-character app password |
| `EMAIL_TO` | where to send reports |

### 5. Workflow permissions
Settings → Actions → General → Workflow permissions → "Read and write permissions" → Save.

### 6. First test run
Actions tab → "Monthly conviction report" → Run workflow.

## How the Bandhan flow works

1. Each month, on the scheduled run (or whenever you run it manually), the
   3 auto funds get scraped and a report is emailed.
2. If no Bandhan file has been processed yet this month, you'll also get a
   **reminder email** with a direct link to Bandhan's download page.
3. Download the file, then upload it to `manual_downloads/` in the repo
   (GitHub web UI: navigate into that folder — create it first with "Add
   file → Create new file" named `manual_downloads/.gitkeep` if it doesn't
   exist yet — then "Add file → Upload files"). Name it starting with
   `bandhan`, e.g. `bandhan_2026-09.xlsx`.
4. That upload automatically triggers the second workflow
   (`bandhan_trigger.yml`), which sends you an **updated report with all 4
   funds included** — no need to wait for next month or run anything manually.

## What's confirmed vs. best-effort

| Fund | Status |
|---|---|
| Nippon India | Confirmed working (tested successfully) |
| quant | Confirmed working after a selector fix |
| Invesco | **Best-effort — first run may need calibration.** If it fails, the debug output dumps every dropdown/input/button on the page so the exact filter interaction can be added in one more round. |
| Bandhan | Manual upload (site blocks automated browsers) |

## Schedule

15th of the month, 5:00 PM US **Eastern** Time (21:00 UTC). This drifts by
1 hour during winter (EST vs EDT) since GitHub cron is UTC-only — see the
comment in `monthly.yml` if you want to correct for that or use a
different US timezone.

## Notes

- All 4 funds are matched by **ISIN** (not name), which is more robust
  than name-matching since it doesn't break on minor spelling differences.
- 3-month trend analysis (stake increases) needs 3 months of history to
  fully populate — earlier runs will show "not enough data yet."
- Sector-level analysis was dropped in this version since none of the 3
  auto-scraped sources reliably expose a separate sector breakdown the way
  the stock-level data is exposed. It can be added back if you find a
  fund's own site includes one.
