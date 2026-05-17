# CLAUDE.md — bmnr-dashboard

Operational context for Claude Code sessions on this repo.

## CELH quarterly update

Trigger phrase: **"Do the Celsius Q[X] 20XX update. Follow the CELH manifest."**

Source of truth for the workflow is `celh/CELH_MANIFEST.txt` — read it first
every time. Summary of what to do:

1. **Fetch the data yourself.** Web access is available in this environment.
   Pull the quarter's results from the **primary source** — Celsius IR
   (ir.celsiusholdingsinc.com) or the BusinessWire press release — not
   secondary aggregators (they round/misstate). Only proceed once that
   quarter's results are actually released.
2. **Show the extracted numbers to the user for confirmation BEFORE editing
   scripts.** Financial line items are where a bad fetch silently corrupts
   output — never skip this check.
3. Update the hardcoded data in:
   - `celh/build_celsius_simple_v4_final.py` — the `ROWS` block (income
     statement: revenue + NA/Intl segments, COGS, gross profit/margin,
     SG&A / S&M / G&A, operating income, other income, pre-tax, tax, net
     income, net income to common, basic/diluted GAAP EPS, adjusted EPS,
     Adj. EBITDA + margin; plus year-ago comparison column and notes).
   - `celh/build_celh_projection_final.py` — the `DATA BLOCK` at top
     (`ENTRY_PRICE`, `BASE_REV`, `SHARES`, `Q1_REV`, `Q1_EBITDA_M`, and
     bull/base/bear assumptions if the user changes them).
4. Run both scripts; the self-test must print `SELF-TEST PASSED`.
5. The manifest's verification baselines (Revenue $782.6M, Bull SPL $217,
   etc.) are **Q1 2026** values. For a new quarter these SHOULD change — a
   mismatch is expected, not a failure. Update the baselines in
   `celh/CELH_MANIFEST.txt` and `celh/README.md` as part of the same commit.
6. Create a **new version file** (v5, v6, …) — never overwrite an existing
   version (manifest Rule 8).
7. Commit and push to the working branch.

## Drive policy

Per manifest Rule 4: **do NOT upload xlsx to Google Drive via API.** Deliver
the finished workbooks as downloadable files; the user uploads them to the
correct quarter folder manually. A Drive integration exists in this
environment but must not be used for xlsx upload unless the user explicitly
overrides this rule in-session.

## Secret handling

- Never hardcode tokens or API keys. Read them from environment variables
  with no committed fallback (see `update_dashboard.py` for the correct
  pattern; `baja_sync.py` now requires `GITHUB_TOKEN` from the env).
- Never embed a token in a git clone URL or in any committed file.
- Note: a previously committed GitHub PAT and a Finnhub key remain in git
  history and should be rotated by the user. The Finnhub key in
  `update_dashboard.py:15` was intentionally left in place per user request.

## Repo notes

- CELH scripts are pure `openpyxl` builders with local self-tests — no
  network, no credentials. Outputs go to `/mnt/user-data/outputs/`.
- `baja_sync.py` and `update_dashboard.py` push to GitHub / fetch market
  data and require their respective env vars to run after credential rotation.
