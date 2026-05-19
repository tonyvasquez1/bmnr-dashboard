# AMD Financial Analysis — Master Reference

Last updated: May 2026
Branch: `claude/setup-amd-financials-q7F3t`

---

## What This Is

A repeatable, self-testing Excel generation system for tracking the AMD investment thesis.
Every file is built from a Python script — no manual edits to spreadsheets, ever.
Scripts live on GitHub. Excel files are build artifacts (generated locally, never committed).

---

## How to Run Everything

### One-time setup
1. Install Python 3.12+ from python.org (check "Add to PATH" during install)
2. Clone the repo:
   ```
   git clone https://github.com/tonyvasquez1/bmnr-dashboard.git
   cd bmnr-dashboard
   ```

### Every time you want fresh files
```
python amd/run_amd.py
```

That's it. The runner auto-installs dependencies (`openpyxl`, `requests`) and generates all three files into the `amd/` folder.

---

## Output Files

| File | Script | Size | Description |
|------|--------|------|-------------|
| `AMD_Q1_2026_Income_Statement_v3.xlsx` | `build_amd_income_v3.py` | ~27 KB | Q1 2026 earnings summary |
| `AMD_5Year_Projection_v12.xlsx` | `build_amd_projection_engine_v7.py` | ~29 KB | 5-year Bull/Base/Bear scenario model |
| `AMD_Thesis_Tracker_v1.xlsx` | `build_amd_thesis_tracker_v1.py` | ~18 KB | Fundamentals-driven exit signal tool |

> Excel files are excluded from git (`.gitignore`). Always regenerate from scripts.

---

## Q1 2026 Income Statement — Key Figures

**Quarter: Q1 2026 | Grade: A− | Script: `build_amd_income_v3.py`**

| Line Item | Q1 2026 | Q4 2025 | Q1 2025 | Signal |
|-----------|---------|---------|---------|--------|
| Net Revenue | $10,253M | $10,270M | $7,438M | +37.9% Y/Y, flat Q/Q |
| Total Cost of Sales | $4,837M | $4,693M | $3,702M | Growing slower than revenue |
| Gross Profit | $5,416M | $5,577M | $3,736M | |
| Gross Margin | 52.8% | 54.0% | 50.2% | +260bps Y/Y |
| Operating Income | $1,476M | $1,752M | $806M | +83% Y/Y |
| GAAP Op Margin | 14.4% | 17.1% | 10.8% | Xilinx amortization drag |
| Net Income | $1,383M | $1,088M | $710M | +95% Y/Y |
| Diluted EPS | $0.84 | $0.69 | $0.44 | +90.9% Y/Y |
| Non-GAAP NI Margin | 22.1% | — | — | Base case target: 25% |

**Q2 2026 Guidance:** Revenue $11.2B midpoint | Non-GAAP EPS $0.96 midpoint | Non-GAAP op margin ~27%

### Income Statement Tabs

1. **Q1 2026 Income Statement** — full GAAP P&L with Y/Y and Q/Q comparisons, color-coded signals, and analyst notes on every line
2. **What's New** — all 5 confirmed deals, coherence check vs projection engine
3. **Data Sources** — cited sources for all figures

### Key Takeaways

- Data Center is the engine: revenue, margin expansion, and EPS growth all driven by MI300X/MI325X GPU ramp + EPYC server share gains
- GAAP margins depressed by ~$551M/quarter non-cash Xilinx amortization (declining over time — structural tailwind)
- Non-GAAP NI margin of 22.1% is below the Base case target of 25% — margin expansion still in progress
- EPS growing (+91% Y/Y) materially faster than revenue (+38% Y/Y) = operating leverage confirmed

---

## 5-Year Projection — Current Outputs (Q1 2026)

**Entry price: $455.00 | Diluted shares: 1.750B | Base revenue: $34.64B (FY2025 actual)**

### Scenario Assumptions

| | 2026 | 2027 | 2028 | 2029 | 2030 |
|---|---|---|---|---|---|
| **Bull rev growth** | +43% | +50% | +52% | +50% | +45% |
| **Bull NI margin** | 28% | 32% | 38% | 41% | 43% |
| **Bull PE range** | 45–55 | 48–58 | 50–60 | 50–60 | 48–58 |
| **Base rev growth** | +36% | +42% | +43% | +42% | +40% |
| **Base NI margin** | 25% | 26% | 26% | 28% | 29% |
| **Base PE range** | 35–40 | 36–41 | 37–42 | 38–43 | 40–45 |
| **Bear rev growth** | +32% | +18% | +12% | +8% | +10% |
| **Bear NI margin** | 22% | 21% | 22% | 24% | 26% |
| **Bear PE range** | 25–30 | 20–25 | 17–22 | 15–20 | 16–21 |

### Probability Weights

| Scenario | Weight | Rationale |
|----------|--------|-----------|
| Bull | 45% | All 5 confirmed GW-scale deals + strong Q1 execution |
| Base | 40% | ROCm adoption risk + competitive pressure caps upside |
| Bear | 15% | Execution failure or macro disruption |

### 2030 Key Outputs

| Scenario | Revenue | Net Income | EPS | SPL | SPH | CAGR (SPL) |
|----------|---------|------------|-----|-----|-----|------------|
| Bull | $245.7B | $105.6B | $60 | $2,897 | $3,501 | +44.8% |
| Base | $190.2B | $55.2B | $32 | $1,261 | $1,418 | +22.6% |
| Bear | $71.8B | $18.7B | $11 | $171 | $224 | -17.8% |
| **Expected Value** | — | — | — | **$1,834** | **$2,176** | **+32.1%** |

> SPL = Share Price Low (EPS × PE Low) | SPH = Share Price High (EPS × PE High)
> EV = (Bull × 45%) + (Base × 40%) + (Bear × 15%)

### Projection Tabs (locked — do not reorder)

1. **Inputs** — entry price, shares, scenario growth rates, PE ranges, probabilities
2. **Projection** — year-by-year outputs for all three scenarios
3. **Probability Weighted** — expected value calculation
4. **Assumptions** — all 5 confirmed deals, data center context, methodology notes
5. **Workflow** — version log, session history
6. **Data Sources** — cited sources for all inputs

### Locked Row Structure (Projection tab)

```
REVENUE
 REV GROWTH
NET INCOME
 NET INC. GROWTH
 NET INC. MARGINS
EPS (Non-GAAP)
 PE LOW EST
 PE HIGH EST
SHARE PRICE LOW
SHARE PRICE HIGH
 CAGR LOW
 CAGR HIGH
```

---

## Confirmed Deals (incorporated in v12)

| Deal | Size | Status | Score |
|------|------|--------|-------|
| OpenAI | 6GW — MI450 Series | H2 2026 deployment | 8/10 |
| Meta | 6GW — Custom MI450-based GPU | H2 2026 deployment | 8/10 |
| Oracle | ~5GW — MI300X cluster | Active | (in Assumptions) |
| HUMAIN | $10B / 500MW Sovereign AI | Multi-exaflop early 2026 | 7/10 |
| DoE / ORNL | Lux AI + Discovery | Active | 7/10 |

> Shares assume 1.650B actual + ~100M net warrant dilution from OpenAI/Meta contracts

---

## Thesis Tracker — Integrity Score

### What It Does

Scores the AMD investment thesis 0–100 each quarter. Below 50 for two consecutive quarters triggers an exit review. Below 65 = watch closely, no new buys.

**Current score: 76/100 — HOLD**

### Score Components

| Condition | Score | Weight | Notes |
|-----------|-------|--------|-------|
| OpenAI Delivery | 8/10 | 12% | MI450 deployment on track |
| New Deal Velocity | 7/10 | 10% | No new GW deals Q1 2026 |
| Meta Delivery | 8/10 | 10% | Custom GPU on track |
| ROCm Production Adoption | 6/10 | 18% | Highest risk — software moat |
| Competitive Position vs Nvidia | 7/10 | 8% | Holding, not gaining |
| EPYC CPU Share | auto | 8% | Live from SEC/quarterly data |
| HUMAIN + DoE Delivery | 7/10 | 8% | Early stage |
| GitHub Ecosystem Health | auto | 7% | Live from GitHub API |
| Guidance Tone | 8/10 | 7% | Bullish Q1 call |
| NI Margin Trajectory | auto | 7% | vs 25% base target |
| DC Revenue vs Base Case | auto | 5% | vs $5.5B quarterly target |

> ROCm weight (18%) is highest because software moat is the primary long-term risk to the bull case.

### Exit Signal Rules

| Range | Action |
|-------|--------|
| 65–100 | HOLD — thesis intact |
| 50–64 | WATCH — monitor closely, no new buys |
| Below 50 | REVIEW EXIT |
| Below 50 × 2 consecutive quarters | Exit review required |

### Live Data Fetching (runs locally, not in cloud)

| Source | Data Fetched |
|--------|-------------|
| GitHub API | ROCm/ROCm + ROCm/pytorch commit activity → auto GitHub health score |
| Yahoo Finance v8 | Live AMD stock price |
| SEC EDGAR | CIK 0000002488 → latest filing data |

Fetch failures fall back to manual DATA BLOCK values with "FETCH FAILED" labeling.

### Thesis Tracker Tabs

1. **Thesis Score** — overall score, component breakdown, exit signal
2. **Deal Execution** — per-deal status and delivery scores
3. **Software Moat** — ROCm adoption, competitive position vs Nvidia
4. **Financial Guardrails** — actual vs base case (DC revenue, NI margin, total revenue)
5. **Live Data** — fetched GitHub/Yahoo/SEC data with timestamps
6. **Data Sources** — all cited sources
7. **Workflow** — version log

---

## Rules — No Exceptions

1. **NEVER** rebuild from memory or scratch — always pull scripts from GitHub
2. **NEVER** start any task without explicit approval
3. **NEVER** change row structure, tab names, or colors without approval
4. **NEVER** save unless self-test prints `SELF-TEST PASSED`
5. **NEVER** upload xlsx to Drive via API — present for download only
6. **ALWAYS** verify script produces correct self-test output before using
7. **ALWAYS** create a new version file for any update — never overwrite
8. **ALWAYS** commit to GitHub after every change

---

## Script Naming Convention

```
build_{ticker}_{type}_v{n}.py
```

Examples: `build_amd_projection_engine_v7.py`, `build_amd_income_v3.py`, `build_amd_thesis_tracker_v1.py`

Quarterly updates increment the version number. Never overwrite a prior version.

---

## Version Log

| Version | Script | Output | Key Change |
|---------|--------|--------|------------|
| v1–v5 | projection_engine (legacy) | Projection_v1–v5 | Early iterations |
| v6 | projection_engine_v6 | Projection_v11 | Q1 2026 + OpenAI/Meta/Oracle deals |
| **v7** | **projection_engine_v7** | **Projection_v12** | HUMAIN + DoE added; Assumptions tab formatting fix (60pt row cap for iPad) |
| v1 | income_v1 | Income_Statement_v1 | Q1 2026 earnings |
| v2 | income_v2 | Income_Statement_v2 | Formatting update |
| **v3** | **income_v3** | **Income_Statement_v3** | Latest earnings |
| **v1** | **thesis_tracker_v1** | **Thesis_Tracker_v1** | New — exit signal tool with live data |

---

## Repository

- **Repo:** https://github.com/tonyvasquez1/bmnr-dashboard
- **Scripts path:** `/amd/`
- **Active branch:** `claude/setup-amd-financials-q7F3t`
- **Runner:** `amd/run_amd.py` — generates all files in one command
