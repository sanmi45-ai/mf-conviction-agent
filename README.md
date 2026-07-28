Mutual Fund Conviction Agent
Tracks holdings of Bandhan Small Cap, Nippon India Small Cap, and
quant Small Cap funds, and builds a ranked list of stocks where multiple
fund managers have conviction, flagging recent entries and exits.
Data source: mfdata.in — a free, no-auth API that
aggregates AMFI's SEBI-mandated monthly portfolio disclosures.
Setup
```bash
pip install -r requirements.txt
```
Run it
```bash
python mf_conviction_agent.py
```
First run: no previous snapshot exists yet, so "recent entries/exits" will
be empty — that's expected. From the second monthly run onward, the
script will diff against last month's saved snapshot and flag changes.
Useful flags:
```bash
python mf_conviction_agent.py --debug       # see raw API responses (for troubleshooting)
python mf_conviction_agent.py --top 30      # change list size from the default 20
```
What it produces
`output/conviction_list_<YYYY-MM>.csv` — the ranked stock list, with:
`conviction_count` — how many of the 3 funds hold the stock (0–3)
`avg_weight_pct` — average % of NAV across the funds holding it
per-fund weight columns
`recent_entries` / `recent_exits` — which fund(s) added or dropped the
stock since the last run
`snapshots/` — raw holdings saved every run, so future runs can detect
changes. Don't delete this folder.
`family_id_cache.json` — caches the fund-name → mfdata.in ID lookup so
the script doesn't re-search every time.
Adding or changing funds
Edit the `FUNDS` list near the top of `mf_conviction_agent.py`:
```python
FUNDS = [
    {"label": "Bandhan Small Cap Fund", "query": "Bandhan Small Cap Fund"},
    {"label": "Nippon India Small Cap Fund", "query": "Nippon India Small Cap Fund"},
    {"label": "quant Small Cap Fund", "query": "quant Small Cap Fund"},
]
```
`label` is just how the fund is displayed in output. `query` is what gets
sent to the search API — keep it close to the fund's official name.
Running it monthly, automatically
AMFI's disclosure deadline is the 10th of the following month, so schedule
the run for the 12th–15th to be safe.
Option A — cron (Linux/Mac, or WSL on Windows)
```bash
crontab -e
# Run at 9am on the 13th of every month:
0 9 13 * * cd /path/to/mf_agent && /usr/bin/python3 mf_conviction_agent.py >> run.log 2>&1
```
Option B — Windows Task Scheduler
Create a monthly trigger (day 13) that runs:
`python C:\path\to\mf_agent\mf_conviction_agent.py`
Option C — GitHub Actions (free, no machine needs to stay on)
Put this repo on GitHub and add `.github/workflows/monthly.yml`:
```yaml
name: Monthly conviction list
on:
  schedule:
    - cron: '0 9 13 * *'   # 9am UTC on the 13th
  workflow_dispatch: {}     # lets you trigger it manually too
jobs:
  run:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'
      - run: pip install -r requirements.txt
      - run: python mf_conviction_agent.py
      - run: |
          git config user.name "mf-agent-bot"
          git config user.email "bot@example.com"
          git add output snapshots family_id_cache.json
          git commit -m "Monthly conviction list $(date +%Y-%m)" || echo "nothing to commit"
          git push
```
This commits each month's CSV and snapshot back into the repo automatically
— free, and you get a running history for nothing extra.
Known limitations / things to sanity-check
`mfdata.in` is a small, community-run free API, not an official source.
It aggregates AMFI's official monthly disclosures, but treat AMFI's own
Excel files (amfiindia.com → Research & Information → Other Data →
Monthly Portfolio Disclosures) as ground truth if a number looks off.
The script matches funds by searching on name text. If a fund house
renames a scheme or the API's naming differs, `resolve_family_id()` might
pick the wrong result — check the printed "matched ..." line on first run
to confirm it found the right fund.
Small-cap funds can have 60–90+ holdings each; with only 3 funds tracked,
don't be surprised if very few stocks hit conviction_count = 3. You can
lower the bar by ranking on conviction_count = 2 as well, which the CSV
already supports (just filter/sort in Excel).
