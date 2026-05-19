# BMNR Dashboard — Claude Instructions

At the start of every session, before doing anything else, read these files:

- `NEW_TICKER_SETUP_GUIDE.txt`
- `celh/CELH_MANIFEST.txt`
- `celh/run_celh.py`
- `celh/build_celh_income_v1.py`
- `celh/build_celh_projection_final.py`

## If the user gives you a stock ticker

Do all of the following without asking the user for any information:

1. Research the company — find the most recent earnings press release, pull
   all financials, segment breakdown, shares outstanding, and current stock
   price using web search.
2. Identify any one-time items, acquisitions, or anything that makes the
   headline numbers misleading.
3. Determine the correct valuation method (EV/EBITDA for consumer/beverage,
   P/E for semiconductor/tech).
4. Build all four deliverables following the CELH files as the gold standard:
   - `{ticker}/build_{ticker}_income_v1.py`
   - `{ticker}/build_{ticker}_projection_v1.py`
   - `{ticker}/run_{ticker}.py`
   - `{ticker}/{TICKER}_MANIFEST.txt`
5. Run `python3 {ticker}/run_{ticker}.py` — both self-tests must pass.
6. Present your proposed Y/Y colors and letter grade with reasoning,
   and wait for Tony's approval before committing.
7. Commit and push all files including generated xlsx outputs.
   Update the manifest per Rule 9 in the same commit.

## If the user says "update [TICKER]" or names an existing ticker

1. Read that ticker's manifest and all current scripts first.
2. Research the most recent earnings — do not ask Tony for financials.
3. Create new version files (increment v number — never overwrite).
4. Update DATA BLOCK only — never touch structure or formatting.
5. Run the ticker's run script — both self-tests must pass.
6. Present any proposed grade or color changes, wait for approval.
7. Update manifest per Rule 9, commit and push everything.

## The 9 Rules — no exceptions

1. NEVER rebuild from memory or scratch. Always pull from GitHub.
2. NEVER start any task without Tony's explicit approval.
3. NEVER save a script to Drive unless self-test printed SELF-TEST PASSED.
4. NEVER upload xlsx to Drive via API — present for download only.
5. NEVER change row structure, colors, or formatting without Tony's approval.
6. ALWAYS verify file sizes match manifest before using any script.
7. ALWAYS run self-test and show Tony the output before declaring done.
8. ALWAYS create a new version file for updates — never overwrite.
9. UPDATE THE MANIFEST immediately whenever anything changes.
