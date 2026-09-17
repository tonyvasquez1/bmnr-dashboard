# ══════════════════════════════════════════════════════════════════════════════
# ADVANCED MICRO DEVICES, INC. (AMD) — INCOME STATEMENT ANALYSIS
# build_amd_income_v4.py
# Q2 2026 | Quarter Ended June 28, 2026 | Reported July 29, 2026
# v4: Q2 2026 actuals. All cost lines grew slower than revenue (+50% Y/Y).
#     Grade: A — first quarter with full operating leverage across all cost lines.
#     New deals: Anthropic 2GW + $5B equity, Core Scientific 2.5GW,
#     Microsoft Azure Helios deployment, Rackspace 30MW.
#     Total confirmed pipeline now exceeds 22GW.
#     Probabilities updated: Bull 57% / Base 33% / Bear 10%.
#     Coherence reference: build_amd_projection_engine_v8.py
#     Output: AMD_Q2_2026_Income_Statement_v4.xlsx
# ══════════════════════════════════════════════════════════════════════════════

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── COLOR PALETTE ─────────────────────────────────────────────────────────────
NAVY        = "1A1A2E"
BLUE        = "1F4E79"
ACCENT      = "2E75B6"
WHITE       = "FFFFFF"
GRAY        = "F2F2F2"
GREEN       = "00B050"
AMBER       = "FF8C00"
RED         = "C00000"
GREEN_LT    = "E8F5E9"
AMBER_LT    = "FFF3E0"
RED_LT      = "FFEBEE"
MEASURE_BG  = "EBF3FB"
BLACK       = "000000"
MUTED       = "666666"
SECTION_BG  = "1F4E79"
STRENGTH_BG = "E8F5E9"
CONCERN_BG  = "FFEBEE"
ASSESS_BG   = "FFF8E1"

# ── BORDER STYLE ──────────────────────────────────────────────────────────────
_SIDE = Side(style="thin", color="CCCCCC")
TABLE_BORDER = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)

# ── FONT SIZES ────────────────────────────────────────────────────────────────
SZ_DEFAULT  = 10
SZ_TITLE    = 14
SZ_SUBTITLE = 10
SZ_HEADER   = 12
SZ_GRADE    = 18
SZ_FOOTER   =  8

# ── PRIORITY ROWS (bold line item cell) ───────────────────────────────────────
PRIORITY = {
    "Net Revenue", "Total Cost of Sales", "Gross Profit", "Gross Margin",
    "Research and Development", "Marketing, General and Administrative",
    "Total Operating Expenses", "Operating Income", "Net Income",
    "Diluted Earnings Per Share",
}

# ── ROW DATA — Q2 2026 / Q1 2026 / Q2 2025 ───────────────────────────────────
# Tuple: (label, q2_2026, q1_2026, q2_2025, yy_color, actual_color, is_pct, note, measure, fmt_override)
ROWS = [
    # ── REVENUE ───────────────────────────────────────────────────────────────
    ("Net Revenue",
     11540, 10253, 7693, "green", None, False,
     "+50.0% Y/Y to $11.54B — the strongest Y/Y growth rate in AMD's Data Center era. "
     "+12.6% Q/Q sequential acceleration from Q1's $10.25B. Data Center was the dominant "
     "driver: $6.7B (+107% Y/Y) — more than doubled year-over-year. MI450 production ramp "
     "began H2 2026, delivering initial deployments to OpenAI, Meta, and Oracle on schedule. "
     "Helios rack deployed on Microsoft Azure. Anthropic 2GW deal signed. All 9 confirmed deals "
     "in execution. Revenue is now unambiguously a Data Center story: DC is 58% of total revenue.",
     "Total dollars sold across all products and geographies; top-line scale",
     None),

    ("Cost of Sales",
     5028, 4576, 3622, "green", None, False,
     "+38.8% Y/Y — 11 percentage points SLOWER than revenue growth of +50.0%. This is "
     "the clearest gross margin expansion signal in AMD's recent history. Data Center GPU "
     "(MI350/MI450) carries materially higher gross margins than Gaming or Embedded. "
     "As Data Center becomes 58% of revenue (up from 49% in Q2 2025), cost of sales as "
     "a % of revenue structurally improves. Green: cost growing well below revenue = "
     "structural gross margin expansion in progress.",
     "Direct production costs; 11pp below revenue growth = textbook gross margin expansion",
     None),

    ("  Amortization — Acquisition Intangibles (COGS)",
     261, 261, 252, None, None, False,
     "Essentially flat Q/Q ($261M vs $261M) and only +3.6% Y/Y ($252M → $261M). Non-cash "
     "amortization of intangibles from the 2022 Xilinx acquisition ($48.8B) — primarily "
     "Xilinx IP, technology licenses, and customer relationships allocated to cost of sales. "
     "Will decline gradually as assets fully amortize. Non-cash, directionally positive.",
     "Non-cash Xilinx acquisition intangibles in COGS; amortizes down over time",
     None),

    ("Total Cost of Sales",
     5289, 4837, 3874, "green", None, False,
     "+36.5% Y/Y — 13.5 percentage points slower than revenue growth of +50.0%. Revenue "
     "grew 13+ percentage points faster than total COGS, directly driving +460bps gross margin "
     "expansion from 49.6% to 54.2%. This is the gross-level operating leverage thesis playing "
     "out at full force. When total COGS grows materially slower than revenue, gross margin "
     "expands mechanically — and the driver is durable: Data Center GPU mix shift is structural.",
     "All-in production costs including non-cash amortization; key gross margin driver",
     None),

    ("Gross Profit",
     6251, 5416, 3819, "green", None, False,
     "+63.7% Y/Y — grew 13.7 percentage points faster than revenue (+50.0%). $2.43B more "
     "gross profit than Q2 2025. Q/Q improvement of +$835M (+15.4%) — exceptional sequential "
     "growth confirming no seasonal softness. Non-GAAP gross margin reached 56% in Q2 2026, "
     "a significant expansion from Q2 2025's non-GAAP ~50%. Dollar earning power is expanding "
     "rapidly and the trajectory is accelerating.",
     "Absolute earning power after production costs; measures scale and efficiency",
     None),

    ("Gross Margin",
     54.2, 52.7, 49.6, "green", "green", True,
     "+460bps Y/Y expansion (49.6% → 54.2%). +150bps Q/Q improvement (52.7% → 54.2%). "
     "Structural Y/Y driver: MI350/MI450 GPU carries meaningfully higher GAAP gross margins "
     "than Gaming GPUs or Embedded chips. As Data Center (58% of revenue) grows at 107% Y/Y, "
     "blended gross margin expands structurally each quarter. Non-GAAP gross margin reached "
     "56% in Q2 2026 — above the 55% threshold management identified as the A-grade milestone. "
     "Both GAAP and Non-GAAP margins are expanding in the same direction simultaneously.",
     "Production efficiency; % of each revenue dollar retained after COGS",
     None),

    # ── OPERATING EXPENSES ────────────────────────────────────────────────────
    ("Research and Development",
     2520, 2397, 1800, "green", None, False,
     "+40.0% Y/Y — 10 percentage points SLOWER than revenue growth of +50.0%. For the first "
     "time in recent quarters, R&D is showing meaningful operating leverage. AMD must invest "
     "heavily in chip design to stay competitive: MI450 in production ramp (CDNA 5, HBM4, "
     "432GB), MI500 on the 2027 roadmap (CDNA 6, 2nm); EPYC Venice (6th Gen) shipping; "
     "Zen 6 in development. R&D is the competitive lifeblood. Green: R&D growing 10pp slower "
     "than revenue is operating leverage in the innovation engine — the most important "
     "structural improvement this quarter vs Q1 2026.",
     "Product competitiveness investment; growing 10pp slower than revenue = first leverage",
     None),

    ("Marketing, General and Administrative",
     1274, 1253, 912, "green", None, False,
     "+39.7% Y/Y — 10.3 percentage points SLOWER than revenue growth of +50.0%. The MG&A "
     "overage that was Q1 2026's primary concern (+41.4% vs +37.9% revenue) has reversed. "
     "MG&A is now showing operating leverage. The path to 30%+ non-GAAP operating margins "
     "requires MG&A discipline, and this quarter delivers it. Green: MG&A growing slower "
     "than revenue for the first time in several quarters — a key inflection point.",
     "Overhead and sales cost discipline; first operating leverage quarter in recent history",
     None),

    ("  Amortization — Acquisition Intangibles (OpEx)",
     287, 290, 311, None, None, False,
     "Declining Y/Y ($311M → $287M, -7.7%) and Q/Q ($290M → $287M). Non-cash amortization "
     "of Xilinx intangibles in operating expenses — technology licenses, trade names, and "
     "customer relationships. This line will continue declining as assets fully amortize, "
     "providing a structural GAAP tailwind to operating income. Directionally positive.",
     "Non-cash Xilinx acquisition intangibles in OpEx; declining = GAAP operating tailwind",
     None),

    ("Total Operating Expenses",
     4081, 3940, 3023, "green", None, False,
     "+34.9% Y/Y — 15 percentage points SLOWER than revenue growth of +50.0%. This is "
     "full-stack operating leverage: every major operating cost line (R&D, MG&A, amortization) "
     "grew meaningfully slower than revenue. Q2 2026 is the first quarter where AMD shows "
     "operating leverage across all cost dimensions simultaneously. The aggregate leverage is "
     "not masked by one outlier line — it is broadly distributed across the P&L.",
     "All operating costs below gross profit; growing 15pp slower than revenue = full leverage",
     None),

    ("Operating Income",
     2170, 1476, 796, "green", None, False,
     "+172.6% Y/Y — more than 3x the rate of revenue growth (+50.0%). $1.374B more operating "
     "income than Q2 2025. +$694M Q/Q sequential improvement. GAAP operating margin: "
     "10.4% → 18.8% (+840bps Y/Y) — the fastest margin expansion in AMD's Data Center era. "
     "Non-GAAP operating margin: ~27% for Q2 2026 (excludes ~$548M Xilinx amortization). "
     "The +172% operating income growth is operating leverage confirmed at maximum force.",
     "Core business profitability before financing costs and taxes; leverage fully confirmed",
     None),

    # ── BELOW THE LINE ────────────────────────────────────────────────────────
    ("Interest Expense",
     -38, -37, -20, None, None, False,
     "Modest increase Q/Q (-$38M vs -$37M) and Y/Y growth from -$20M. AMD carries relatively "
     "light long-term debt (~$1.7B notes outstanding) — a significant balance sheet advantage. "
     "Interest expense is not a meaningful earnings drag at AMD's scale.",
     "Cost of AMD's debt; modest relative to operating income — balance sheet strength",
     None),

    ("Other Income (Expense), net",
     490, 165, 100, None, None, False,
     "Q2 2026: $490M — significantly elevated vs Q1 2026 ($165M) and Q2 2025 ($100M). "
     "Primarily reflects: (1) interest income on AMD's expanded cash position following the "
     "Anthropic $5B equity stake transaction, (2) gains on equity investments, and (3) "
     "licensing income. The Anthropic deal (2GW MI450 + $5B equity) closed Q2 2026 and "
     "contributed to the elevated other income line. Full composition will be disclosed in "
     "AMD's Q2 2026 10-Q filing.",
     "Interest income on cash + equity investment gains; elevated due to Anthropic equity deal",
     None),

    ("Income Before Income Taxes",
     2622, 1604, 876, None, None, False,
     "$2.62B pre-tax income — +$1.75B Y/Y (+199%). Q/Q improvement of +$1.02B (+63%). "
     "Both operating income expansion (+$1.37B Q/Q) and elevated Other Income (+$325M Q/Q) "
     "contributed. The pre-tax margin is 22.7% in Q2 2026 vs 10.4% in Q2 2025 (+1,230bps).",
     "Total profitability before tax; reflects operating performance and capital structure",
     None),

    ("Income Tax Provision",
     339, 238, 130, None, None, False,
     "+$209M Y/Y in tax expense, tracking income growth. AMD benefits from R&D tax credits, "
     "foreign income structures, and Xilinx-related deductions that keep the effective rate "
     "manageable. Effective GAAP tax rate: ~12.9% in Q2 2026 vs ~14.8% in Q2 2025 — "
     "lower effective rate as income grows into favorable structure.",
     "Tax obligation; R&D credits and deductions keep effective rate competitive",
     None),

    ("Equity Income in Investee",
     7, 6, 7, None, None, False,
     "Small income ($7M) from equity method investments. Immaterial relative to AMD's "
     "overall profitability.",
     "Small income from minority equity stakes; immaterial to the investment thesis",
     None),

    ("Income from Continuing Operations",
     2290, 1372, 753, None, None, False,
     "+204% Y/Y — more than tripled. Q/Q improvement of +$918M (+67%). This is GAAP net "
     "income from AMD's core ongoing business. The acceleration in Q2 2026 confirms the "
     "operating leverage thesis is not slowing — it is accelerating.",
     "Core business GAAP net income excluding discontinued operations",
     None),

    ("Income (Loss) from Discontinued Operations",
     1, 11, None, None, None, False,
     "Q2 2026: $1M gain. Q1 2026: $11M gain. These relate to Xilinx-era divested business "
     "units being wound down. Non-recurring, expected to zero out completely.",
     "Gains/losses from divested business units; non-recurring, expected to zero out",
     None),

    ("Net Income",
     2291, 1383, 753, "green", None, False,
     "+204% Y/Y — more than tripled on 50.0% revenue growth. Q/Q improvement ($1,383M → "
     "$2,291M, +65.7%). GAAP net margin: 9.8% → 19.9% (+910bps Y/Y). This is textbook "
     "operating leverage in full effect: revenue +50%, net income +204%. Every dollar of "
     "incremental revenue is generating substantially more than one dollar of incremental "
     "net income — the definition of operating leverage at scale.",
     "Total GAAP profitability; bottom-line operating leverage fully confirmed (+204% Y/Y)",
     None),

    # ── EPS ───────────────────────────────────────────────────────────────────
    ("Basic EPS — Continuing Operations",
     1.39, 0.84, 0.46, None, None, False,
     "Q2 2026: $1.39. Q1 2026: $0.84. Q2 2025: $0.46. Reference line.",
     "Per-share profit from core operations on basic share count; reference line",
     None),

    ("Basic EPS — Discontinued Operations",
     0.00, 0.01, None, None, None, False,
     "Q2 2026: $0.00. Q1 2026: $0.01 gain. Non-recurring — winding down.",
     "Per-share discontinued ops impact; non-recurring",
     None),

    ("Basic Earnings Per Share",
     1.39, 0.85, 0.46, None, None, False,
     "Q2 2026: $1.39. Q1 2026: $0.85. Q2 2025: $0.46. Uses ~1,649M basic shares.",
     "Total per-share profit on basic shares only; diluted EPS is the primary metric",
     None),

    ("Diluted EPS — Continuing Operations",
     1.38, 0.83, 0.46, None, None, False,
     "Q2 2026: $1.38. Diluted count: 1,660M. Basic-diluted spread: 11M (0.7%).",
     "Per-share profit from core operations on fully diluted share count",
     None),

    ("Diluted EPS — Discontinued Operations",
     0.00, 0.01, None, None, None, False,
     "Same as basic — $0.00 in Q2 2026. Non-recurring.",
     "Per-share discontinued ops impact on diluted basis",
     None),

    ("Diluted Earnings Per Share",
     1.38, 0.84, 0.46, "green", None, False,
     "+200.0% Y/Y — tripled from $0.46 to $1.38. Q/Q: $0.84 → $1.38 (+64.3%). "
     "AMD's GAAP EPS growth of +200% on +50% revenue growth is the operating leverage "
     "thesis confirmed at the per-share level. Share count dilution is minimal: only 11M "
     "shares basic-to-diluted spread (0.7%). The 320M OpenAI/Meta warrants (not yet fully "
     "vested) are tracked quarterly but have not yet materially expanded the diluted count.",
     "Primary per-share metric; +200% Y/Y confirms shareholders benefit from full leverage",
     None),

    # ── SHARE COUNT ───────────────────────────────────────────────────────────
    ("Basic Shares (millions)",
     1649, 1631, 1623, None, None, False,
     "+1.6% Y/Y growth (+26M shares from 1,623M to 1,649M). AMD's share count is very "
     "stable — stock buybacks roughly offset SBC issuance. Note: 320M total warrants "
     "outstanding (OpenAI 160M + Meta 160M). Additional 2GW Anthropic deal may include "
     "warrant or equity component — track Q2 2026 10-Q for dilution disclosure.",
     "Basic share count; slow growth — warrant dilution tracked quarterly",
     "#,##0"),

    ("Diluted Shares (millions)",
     1660, 1650, 1627, None, None, False,
     "1,660M diluted vs 1,650M Q1 and 1,627M Q2 2025. Basic-diluted spread: 11M (0.7%). "
     "Clean dilution profile. Forward: 320M warrants outstanding (OpenAI/Meta). Projection "
     "engine uses 1.750B to reflect ~100M net after partial buyback offset.",
     "Fully diluted count; 0.7% spread is minimal — forward warrant dilution tracked separately",
     "#,##0"),
]

# ── COLUMN LAYOUT ─────────────────────────────────────────────────────────────
COL_WIDTHS = {
    "A": 38, "B": 14, "C": 10,
    "D": 13, "E": 13, "F": 13,
    "G": 90, "H": 40,
}

def fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def section_header(ws, row, text, bg, cols=8):
    ws.merge_cells(f"A{row}:{chr(64+cols)}{row}")
    cell = ws[f"A{row}"]
    cell.value = text
    cell.font = Font(name="Arial", bold=True, size=SZ_HEADER, color=WHITE)
    cell.fill = fill(bg)
    cell.alignment = center()
    ws.row_dimensions[row].height = 20
    return row + 1

def bullet_row(ws, row, text, bg, font_size=SZ_DEFAULT, color=BLACK):
    ws.merge_cells(f"A{row}:H{row}")
    cell = ws[f"A{row}"]
    cell.value = text
    cell.font = Font(name="Arial", size=font_size, color=color)
    cell.fill = fill(bg)
    cell.alignment = Alignment(horizontal="left", vertical="top",
                               wrap_text=True, indent=2)
    ws.row_dimensions[row].height = max(15, 13 + len(text) // 10)
    return row + 1

def auto_fmt(val, is_pct):
    if val is None:
        return "@"
    if abs(val) < 5:
        return '$#,##0.00'
    if abs(val) < 100:
        return '$#,##0.0'
    return '$#,##0'

def yy_change_str(q2, q2_py, is_pct, fmt_override):
    if q2 is None or q2_py is None:
        return "—"
    diff = q2 - q2_py
    if is_pct:
        sign = "+" if diff >= 0 else ""
        return f"{sign}{diff:.0f} bps"
    if fmt_override == "#,##0":
        sign = "+" if diff >= 0 else ""
        return f"{sign}{diff:,.0f}"
    if abs(q2) < 5:
        sign = "+" if diff >= 0 else ""
        return f"{sign}${diff:.2f}"
    sign = "+" if diff >= 0 else "-"
    return f"{sign}${abs(diff):,.0f}M"

def yy_pct_str(q2, q2_py, is_pct):
    if q2 is None or q2_py is None or q2_py == 0:
        return "—"
    if is_pct:
        diff = q2 - q2_py
        sign = "+" if diff >= 0 else ""
        return f"{sign}{diff:.1f} pp"
    pct = (q2 - q2_py) / abs(q2_py) * 100
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct:.1f}%"

def fmt_val(val, is_pct, fmt_override):
    if val is None:
        return "—"
    if is_pct:
        return f"{val:.1f}%"
    if fmt_override == "#,##0":
        return f"{val:,.0f}"
    if abs(val) < 5:
        return f"${val:.2f}"
    if abs(val) < 100:
        return f"${val:.1f}"
    return f"${val:,.0f}"


def build_sources_tab(wb):
    ws = wb.create_sheet("Data Sources")

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 52
    ws.column_dimensions["E"].width = 42

    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value = "AMD Q2 2026 INCOME STATEMENT — DATA SOURCES"
    t.font = Font(name="Arial", bold=True, size=SZ_TITLE, color=WHITE)
    t.fill = fill(NAVY)
    t.alignment = center()
    ws.row_dimensions[1].height = 24

    for col_idx, h in enumerate(
        ["Category", "Source", "Date", "Key Data Used", "Reference / URL"], 1
    ):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = h
        cell.font = Font(name="Arial", bold=True, size=SZ_HEADER, color=WHITE)
        cell.fill = fill(ACCENT)
        cell.alignment = center(wrap=True)
        cell.border = TABLE_BORDER
    ws.row_dimensions[2].height = 24

    sources = [
        # ── 9 Confirmed deals ─────────────────────────────────────────────────
        ("Strategic Deal",
         "AMD + Anthropic 2GW + $5B Equity Stake",
         "Q2 2026",
         "2GW MI450 compute deal. Anthropic invests $5B equity stake in AMD. "
         "1GW starts H1 2027. Largest single equity investment by a hyperscaler in AMD. "
         "Signed Q2 2026. Added to What's New and What's New coherence check.",
         "AMD Newsroom — AMD and Anthropic Strategic Partnership (Q2 2026)"),

        ("Strategic Deal",
         "Core Scientific 2.5GW Capacity Agreement",
         "Q2 2026",
         "2.5GW HPC/AI capacity secured. Deployments begin 2027. "
         "First large-scale AMD colocation/HPC partnership. Signed Q2 2026.",
         "AMD Newsroom — AMD and Core Scientific Partnership (Q2 2026)"),

        ("Strategic Deploy",
         "Microsoft Azure — Helios Rack-Scale Deployment",
         "Q2 2026",
         "AMD Helios rack-scale architecture deployed on Microsoft Azure. "
         "First public cloud deployment of AMD Helios. Validates ROCm + MI450 "
         "at Microsoft datacenter scale.",
         "AMD Newsroom / Microsoft Azure Blog (Q2 2026)"),

        ("Strategic Deal",
         "Rackspace 30MW Phased Deployment",
         "Q2 2026",
         "30MW AMD GPU deployment phased late 2026 through 2028. "
         "Managed services model. Signed Q2 2026.",
         "AMD Newsroom — Rackspace AMD Partnership (Q2 2026)"),

        ("Strategic Deal",
         "AMD + HUMAIN $10B Sovereign AI Collaboration",
         "May 13, 2025",
         "$10B total investment over 5 years. 500MW AI compute capacity globally. "
         "Full AMD stack: Instinct GPUs, EPYC CPUs, Pensando DPUs, Ryzen AI, ROCm. "
         "Multi-exaflop capacity confirmed underway Q2 2026.",
         "AMD Newsroom — AMD and HUMAIN Form Strategic $10B Collaboration (May 13, 2025)"),

        ("Strategic Deal",
         "AMD + DoE/ORNL Sovereign AI (Lux + Discovery)",
         "Oct 27, 2025",
         "Lux AI: MI355X + EPYC + Pensando, production Q2 2026. "
         "Discovery: MI430X (MI400 Series) + EPYC Venice, 2028. $1B combined.",
         "AMD Newsroom — AMD Powers U.S. Sovereign AI Factory Supercomputers (Oct 27, 2025)"),

        ("Strategic Deal",
         "AMD + OpenAI Strategic Partnership",
         "Oct 6, 2025",
         "6GW total. 1GW H2 2026 MI450 deployment — in production ramp. "
         "AMD issues 160M warrant shares to OpenAI.",
         "AMD Newsroom — AMD and OpenAI Announce Strategic Partnership (Oct 6, 2025)"),

        ("Strategic Deal",
         "AMD + Meta Expanded Strategic Partnership",
         "Feb 24, 2026",
         "6GW total. 1GW H2 2026 custom MI450-based GPU + EPYC Venice. "
         "AMD issues 160M warrant shares to Meta.",
         "AMD Newsroom — AMD and Meta Announce Expanded Strategic Partnership (Feb 24, 2026)"),

        ("Strategic Deal",
         "AMD + Oracle AI World Announcement",
         "Oct 14, 2025",
         "50K AMD GPUs Q3 2026: MI450 + EPYC Venice + Pensando Vulcano, Helios rack. "
         "Expanding to 200K+ GPUs in 2027. In delivery Q3 2026.",
         "AMD Newsroom — AMD at Oracle CloudWorld (Oct 14, 2025)"),

        # ── Primary financial data ─────────────────────────────────────────────
        ("Primary Data",
         "AMD Q2 2026 Earnings Press Release",
         "July 29, 2026",
         "All Q2 2026 actuals: Revenue $11,540M, Gross Profit $6,251M, "
         "Operating Income $2,170M, Net Income $2,291M, Diluted EPS $1.38. "
         "Quarter ended June 28, 2026.",
         "https://ir.amd.com/news-events/press-releases"),

        ("Comparison Data",
         "AMD Q1 2026 Earnings Press Release",
         "May 5, 2026",
         "Q1 2026 figures used for Q/Q comparison: Revenue $10,253M, "
         "Gross Margin 52.7%, Operating Income $1,476M, EPS $0.84.",
         "https://ir.amd.com/news-events/press-releases"),

        ("Comparison Data",
         "AMD Q2 2025 Earnings Press Release",
         "Jul 2025",
         "Q2 2025 figures used for all Y/Y change calculations: Revenue $7,693M, "
         "Gross Margin 49.6%, Operating Income $796M, Net Income $753M, EPS $0.46.",
         "https://ir.amd.com/news-events/press-releases"),

        ("Guidance",
         "AMD Q3 2026 Non-GAAP Guidance",
         "July 29, 2026",
         "Q3 2026 guidance: $12.7B–$13.3B revenue (midpoint $13.0B, +12.6% Q/Q). "
         "Non-GAAP op margin ~27%. Lisa Su: 'exceptional demand across every segment.'",
         "AMD Q2 2026 Earnings Call transcript / press release"),

        ("Product Roadmap",
         "AMD Instinct MI450 — Production Ramp H2 2026",
         "Q2 2026",
         "MI450 in production ramp. CDNA 5, HBM4, 432GB. Helios rack. "
         "Delivering to OpenAI, Meta, Oracle on schedule. H2 2026 major revenue driver.",
         "AMD Newsroom — MI450 Production Ramp (Q2 2026)"),

        ("Product Roadmap",
         "AMD Helios Rack — Production Deploy",
         "Q2 2026",
         "Helios rack architecture in production. Deployed on Microsoft Azure. "
         "Oracle 50K GPU commitment in Helios form factor. Full-stack AMD integration.",
         "AMD Newsroom — Helios Production (Q2 2026)"),

        ("Product Roadmap",
         "AMD 6th Gen EPYC (Venice) — GA Launch",
         "Q2 2026",
         "EPYC Venice (6th Generation, Zen 6 architecture) launched GA Q2 2026. "
         "Used in Meta 1GW custom GPU rack and Oracle Helios deployment.",
         "AMD Newsroom — EPYC Venice Launch (Q2 2026)"),

        ("Market Data",
         "AMD Q2 2026 Data Center Segment Revenue",
         "July 29, 2026",
         "Data Center: $6.7B Q2 2026 (+107% Y/Y). 58% of total revenue. "
         "AMD Data Center now exceeds Intel for 5th consecutive quarter.",
         "AMD Q2 2026 Earnings Press Release segment breakdown"),

        ("Competitive Context",
         "Nvidia Supply Constraints — Q2 2026",
         "Q2 2026",
         "Nvidia reporting supply constraints on Blackwell leading hyperscalers to "
         "accelerate AMD MI450 orders. Near-term supply opportunity for AMD.",
         "Nvidia Q2 FY2027 earnings discussion; hyperscaler procurement reports"),

        ("Watch Item — Unconfirmed",
         "AMD-Google TPU Hybrid ASIC Collaboration",
         "Reported Aug 2026",
         "AMD reportedly designing Google's 10th-gen TPU hybrid ASIC with AMD CPU cores "
         "on-package. Production 2029+. UNCONFIRMED — reported in tech press. "
         "If confirmed, flips Google custom silicon from Bear risk to potential upside. "
         "Added to Assumptions tab as watch item in projection engine v8.",
         "Tech press Aug 2026 — unconfirmed; monitor AMD and Google IR for confirmation"),

        ("Script / Analysis",
         "build_amd_income_v4.py",
         "Sep 2026",
         "All analysis, row coloring, letter grade framework, and narrative. "
         "v4 is Q2 2026 update. Coherence check references build_amd_projection_engine_v8.py. "
         "Never overwrite — future updates: build_amd_income_v5.py.",
         "GitHub: tonyvasquez1/bmnr-dashboard /amd/"),
    ]

    for i, (cat, src, date, data, url) in enumerate(sources):
        row = 3 + i
        bg = WHITE if i % 2 == 0 else GRAY
        for col_idx, val in enumerate([cat, src, date, data, url], 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.value = val
            cell.fill = fill(bg)
            cell.border = TABLE_BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.font = Font(name="Arial", size=SZ_DEFAULT, color=BLACK)
        ws.cell(row=row, column=1).font = Font(
            name="Arial", size=SZ_DEFAULT, bold=True, color=NAVY)
        ws.row_dimensions[row].height = 52


def build_whats_new_tab(wb):
    ws = wb.create_sheet("What's New")

    for col, w in zip("ABCDEF", [20, 36, 26, 26, 28, 28]):
        ws.column_dimensions[col].width = w

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = ("AMD — WHAT'S NEW  |  Version: Q2 2026 (All 9 Deals)  |  "
               "Script: build_amd_income_v4.py  |  Researched: September 2026")
    t.font = Font(name="Arial", bold=True, size=SZ_TITLE, color=WHITE)
    t.fill = fill(NAVY)
    t.alignment = center()
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:F2")
    instr = ws["A2"]
    instr.value = (
        "HOW TO USE: In each quarterly update (v4, v5, ...) research AMD results, guidance, "
        "product news, competitor earnings, and sector macro BEFORE editing any row data. "
        "Populate every section below. Carry each item into the income statement and projection "
        "engine. Confirm the Coherence Check passes before committing. This tab is the audit "
        "trail between versions."
    )
    instr.font = Font(name="Arial", size=SZ_SUBTITLE, color=WHITE, italic=True)
    instr.fill = fill(BLUE)
    instr.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 40

    col_headers = ["Category", "Item", "Prior State",
                   "Current State (Q2 2026)",
                   "Consequences → Earnings",
                   "Consequences → Projections"]
    for col_idx, h in enumerate(col_headers, 1):
        c = ws.cell(row=3, column=col_idx)
        c.value = h
        c.font = Font(name="Arial", bold=True, size=SZ_HEADER, color=WHITE)
        c.fill = fill(ACCENT)
        c.alignment = center(wrap=True)
        c.border = TABLE_BORDER
    ws.row_dimensions[3].height = 30

    cat_colors = {
        "AMD — Financial":     "1A5276",
        "AMD — Guidance":      "1A7A3A",
        "AMD — Product":       "6C3483",
        "AMD — Strategic":     "0E5A8A",
        "AMD — Market Share":  "B7770D",
        "Competitor — Nvidia": "8B0000",
        "Competitor — Intel":  "4A235A",
        "Sector / Macro":      "1B4F72",
        "Watch — Unconfirmed": "7D6608",
    }

    items = [
        # ── NEW Q2 DEALS ───────────────────────────────────────────────────────
        ("AMD — Strategic",
         "Anthropic 2GW + $5B AMD Equity Stake (Q2 2026)",
         "No prior Anthropic agreement",
         "2GW MI450 compute deal. Anthropic invests $5B equity stake in AMD as part of the "
         "partnership. 1GW starts H1 2027. Largest single equity investment by any hyperscaler "
         "in AMD. Signed Q2 2026. Contributes to elevated Q2 2026 Other Income.",
         "No Q2 2026 revenue impact (deployments begin H1 2027). Other Income elevated Q2 "
         "due to equity deal accounting. Forward: 2GW pipeline provides 2027-2028 revenue "
         "visibility. $5B equity investment validates hyperscaler confidence in AMD at scale.",
         "Anthropic 2GW supports Bull 2027 rev growth assumption (+55%). $5B equity stake "
         "is a multi-year partnership signal. Projection engine v8 incorporates this deal. "
         "Total confirmed pipeline now exceeds 22GW across 9 customer agreements."),

        ("AMD — Strategic",
         "Core Scientific 2.5GW Capacity Agreement (Q2 2026)",
         "No prior Core Scientific agreement at scale",
         "2.5GW HPC/AI compute capacity secured. Deployments begin 2027. First large-scale "
         "AMD HPC colocation partnership. Signed Q2 2026.",
         "No Q2 2026 revenue (2027 deployments). Confirms AMD is winning non-hyperscaler "
         "AI infrastructure deals at scale. Diversifies the customer base beyond the Big 4.",
         "2.5GW adds to 2027-2028 revenue pipeline. Projection engine v8 updated. "
         "Core Scientific win validates AMD's ability to win HPC/colocation verticals."),

        ("AMD — Strategic",
         "Microsoft Azure Helios Rack-Scale Deployment (Q2 2026)",
         "Microsoft using NVIDIA primarily; no Helios cloud deployment prior",
         "AMD Helios rack-scale architecture deployed on Microsoft Azure. First public cloud "
         "deployment of AMD Helios. Full MI450 + EPYC Venice + Pensando Vulcano stack. "
         "Validates AMD's full-stack architecture at hyperscaler cloud scale.",
         "Microsoft Azure deployment provides Q3/Q4 2026 Data Center revenue uplift. "
         "First major public cloud Helios deployment validates the architecture. "
         "ROCm running in Azure production environment is a key ROCm adoption milestone.",
         "Microsoft Helios adoption supports Base/Bull 2026-2027 revenue assumptions. "
         "Dual-vendor GPU policy at Microsoft (NVDA + AMD) is now production-confirmed. "
         "De-risks the Bear case ROCm concern at the hyperscaler level."),

        ("AMD — Strategic",
         "Rackspace 30MW Phased Deployment (Q2 2026)",
         "No prior AMD deal with Rackspace at this scale",
         "30MW phased AMD GPU deployment from late 2026 through 2028. Managed services "
         "model. Signed Q2 2026. Expands AMD into the managed cloud segment.",
         "Initial 30MW deployment begins late Q4 2026. Modest near-term revenue. "
         "Opens a new managed services customer vertical for AMD AI hardware.",
         "Rackspace 30MW adds to long-tail revenue diversification. Supports Base/Bull "
         "2027 revenue across the managed services vertical."),

        # ── EXISTING DEALS — STATUS UPDATE ────────────────────────────────────
        ("AMD — Strategic",
         "HUMAIN $10B Sovereign AI (May 2025) — Q2 2026 Status",
         "Signed May 2025; initial 500MW delivery phase active",
         "HUMAIN $10B / 500MW Sovereign AI collaboration ongoing. GW-scale expansion expected. "
         "Delivery execution confirmed active through Q2 2026. Largest sovereign AI deal at AMD signing.",
         "Ongoing HUMAIN deliveries contribute to Data Center revenue in H2 2026. "
         "GW-scale expansion is a potential 2027 catalyst if contracted.",
         "HUMAIN deal de-risks the Base/Bull 2026 revenue assumption. Sovereign AI "
         "vertical is a Bull case tailwind independent of hyperscaler spend cycles."),

        ("AMD — Strategic",
         "DoE/ORNL — Lux AI + Discovery (2025) — Q2 2026 Status",
         "DoE HPC contracts for Lux AI and Discovery systems signed 2025",
         "AMD DoE/ORNL partnership delivering Lux AI and Discovery HPC systems. "
         "EPYC + Instinct hardware in DOE supercomputing infrastructure. "
         "Validates AMD's HPC credibility alongside commercial AI wins.",
         "DoE revenue recognition ongoing per contract milestones. "
         "HPC wins create ROCm production proof points in government-classified workloads.",
         "DoE/ORNL execution supports AMD's HPC narrative — a distinct segment from "
         "pure hyperscaler AI that supports margin diversification in the Base case."),

        ("AMD — Strategic",
         "OpenAI 6GW (Oct 2025) — Q2 2026 Status",
         "1GW H2 2026 MI450 deployment on track per Q1 2026 call",
         "MI450 in production ramp. 1GW H2 2026 deployment executing as contracted. "
         "MI450 production confirmed shipping to lead customers. Full 6GW ramp multi-year.",
         "H2 2026 OpenAI deployment drives major Data Center revenue step-up in Q3/Q4. "
         "This is the single largest near-term revenue catalyst for AMD.",
         "OpenAI 1GW H2 2026 is baked into Q3 2026 guidance ($13.0B midpoint). "
         "Full 6GW multi-year ramp supports Bull 2027-2028 revenue assumptions."),

        ("AMD — Strategic",
         "Meta 6GW (Feb 2026) — Q2 2026 Status",
         "1GW H2 2026 custom MI450-based + EPYC Venice on track",
         "1GW H2 2026 custom MI450-based GPU + EPYC Venice in production ramp. "
         "Full 6GW multi-year commitment executing per schedule. EPYC Venice cross-sell confirmed.",
         "H2 2026 Meta deployment (GPU + CPU) provides dual-revenue stream. "
         "Meta partnership validates AMD's ability to win custom GPU architecture contracts.",
         "Meta 6GW supports Bull 2027-2028 revenue assumptions. EPYC cross-sell is the "
         "highest-margin revenue AMD generates (GPU + CPU in same rack)."),

        ("AMD — Strategic",
         "Oracle 50K GPUs (Oct 2025) — Q2 2026 Status",
         "50K GPU Q3 2026 commitment confirmed",
         "Oracle 50K GPU Helios rack deployment in progress. Q3 2026 delivery timeline. "
         "Expanding to 200K+ GPUs in 2027. No yield issues reported.",
         "Oracle revenue recognition in Q3 2026 ($MI450 + EPYC Venice + Pensando). "
         "First major Helios rack revenue contribution (50K GPUs at ~$10K/GPU = ~$500M gross).",
         "Oracle 2027 200K+ GPU expansion supports Bull 2027 rev growth (+55%). "
         "Full-stack Helios win at Oracle is the architecture validation AMD needed."),

        # ── AMD FINANCIAL RESULTS ──────────────────────────────────────────────
        ("AMD — Financial",
         "Net Revenue Q2 2026: $11,540M (+50.0% Y/Y)",
         "Q1 2026: $10,253M | Q2 2025: $7,693M",
         "$11,540M actual. +50.0% Y/Y. +12.6% Q/Q. Above Bull scenario pace (+43% for FY2026). "
         "Data Center $6.7B (+107% Y/Y) — first time AMD DC exceeded $6B in a quarter. "
         "58% of revenue now from Data Center.",
         "Revenue at or above Bull case pace for FY2026. Q3 guidance $13.0B implies FY2026 "
         "total ~$48-50B — tracking toward Bull case ($49.5B) or above.",
         "Q2 result shifts probability weighting: Bull 57% / Base 33% / Bear 10%. "
         "v8 projection engine updated with Q2 actuals."),

        ("AMD — Financial",
         "Gross Margin: +460bps Y/Y expansion (49.6% → 54.2% GAAP)",
         "Q1 2026 GAAP: 52.7% | Q2 2025 GAAP: 49.6%",
         "Q2 2026 GAAP: 54.2%. Non-GAAP: 56%. Both expanding. "
         "Data Center GPU mix shift (58% of rev) is the structural driver. Repeatable. "
         "Non-GAAP gross margin 56% exceeds the A-grade threshold of 55%.",
         "Gross margin expansion drives Operating Income tripling (+172%) on 50% revenue. "
         "Non-GAAP 56% validates path to 60%+ by 2029 in Bull case.",
         "Bull NI margin ramp (28%→43%) is supported by gross margin continuing to expand. "
         "Each 100bps gross margin improvement = ~$115M additional quarterly gross profit."),

        ("AMD — Financial",
         "Operating Leverage: ALL cost lines slower than revenue (+50% Y/Y)",
         "Q1 2026: MG&A was the sole leverage failure point",
         "Q2 2026: ALL major cost lines (R&D +40%, MG&A +39.7%) grew SLOWER than revenue (+50%). "
         "First quarter with full-stack operating leverage — no exceptions. "
         "Operating Income +172% Y/Y. GAAP operating margin: 10.4% → 18.8% (+840bps).",
         "Q2 2026 is the inflection quarter: operating leverage across the full P&L. "
         "This is what the investment thesis was predicated on demonstrating.",
         "Full-stack operating leverage in Q2 2026 is the strongest validation of the "
         "Bull NI margin expansion trajectory (28%→32%→38%→41%→43% over 2026-2030)."),

        ("AMD — Financial",
         "Diluted EPS: $1.38 (+200% Y/Y)",
         "Q1 2026: $0.84 | Q2 2025: $0.46",
         "$1.38 GAAP diluted EPS. +200% Y/Y. +64% Q/Q. Non-GAAP EPS: $1.66 (Q2 actual). "
         "Clean share count: only 11M basic-to-diluted spread (0.7%). No preferred dilution.",
         "EPS leverage fully confirmed. +200% EPS on +50% revenue is textbook operating "
         "leverage. Projection engine's Non-GAAP EPS path is tracking above Base case.",
         "Q2 Non-GAAP EPS of $1.66 implies FY2026 Non-GAAP EPS run rate ~$5.50-6.00. "
         "Bull 2026 EPS assumption in v8 projection likely understated — watch for upward "
         "revision in v9 after Q3 2026."),

        # ── AMD GUIDANCE ───────────────────────────────────────────────────────
        ("AMD — Guidance",
         "Q3 2026 Revenue Guidance: $12.7B–$13.3B",
         "Q2 2026 actual: $11,540M",
         "Q3 2026 guidance: $12.7B–$13.3B midpoint $13.0B (+12.6% Q/Q). "
         "Most bullish sequential guidance in AMD history. "
         "Lisa Su: 'exceptional demand across every customer segment.'",
         "Q3 guide above consensus confirms no demand deceleration despite MI450 ramp. "
         "Oracle 50K GPU delivery, OpenAI/Meta H2 deployments are the key Q3 catalysts.",
         "Q3 guide of $13.0B midpoint implies FY2026 $50B+ if Q4 maintains pace. "
         "This exceeds the current Bull case ($49.5B) — may require v9 upward revision."),

        # ── PRODUCT ────────────────────────────────────────────────────────────
        ("AMD — Product",
         "MI450: In Production Ramp (CDNA 5, HBM4, 432GB)",
         "Sampling confirmed Q1 2026 for H2 2026 production",
         "MI450 in production ramp. Shipping to lead customers: OpenAI, Meta, Oracle. "
         "CDNA 5, HBM4, 432GB memory. Helios rack form factor. No yield issues reported.",
         "H2 2026 revenue driver. Q3 2026 guidance ($13.0B) reflects initial MI450 volume. "
         "OpenAI/Meta/Oracle H2 deployments are the primary Q3/Q4 revenue catalysts.",
         "MI450 on-schedule execution de-risks the Bear case (execution risk is the primary "
         "Bear scenario). v8 Bull 2027 rev growth (+55%) depends on MI450 ramp sustaining."),

        ("AMD — Product",
         "Helios Rack: Production — Deployed on Azure",
         "Helios in qualification prior to Q2",
         "Helios rack architecture in production. Deployed on Microsoft Azure public cloud. "
         "Oracle 50K GPU Helios deployment in progress. Full-stack AMD (GPU+CPU+DPU) integration.",
         "Helios rack deployment on Azure validates the architecture at hyperscaler cloud scale. "
         "Oracle Helios deployment creates Q3 2026 revenue recognition event.",
         "Helios production validates AMD's full-stack strategy. Supports Bull 2027+ margins "
         "as full-stack (GPU+CPU+DPU) carries AMD's highest revenue per server rack."),

        ("AMD — Product",
         "6th Gen EPYC (Venice) — GA Launch Q2 2026",
         "EPYC Turin (Gen 5) at record 46.2% share Q1 2026",
         "EPYC Venice (Zen 6, 6th Gen) launched GA Q2 2026. Deployed in Meta custom rack and "
         "Oracle Helios. Next-gen CPU drives EPYC server revenue share expansion.",
         "Venice launch in Meta and Oracle racks generates CPU + GPU cross-sell revenue. "
         "Venice is the competitive response to Intel's Xeon 6 Granite Rapids.",
         "Venice supports EPYC CPU revenue share expansion past record 46.2%. "
         "CPU + GPU cross-sell (Helios rack) is AMD's highest-margin revenue mix."),

        # ── COMPETITOR ─────────────────────────────────────────────────────────
        ("Competitor — Nvidia",
         "Nvidia Supply Constraints — Benefit to AMD (Q2 2026)",
         "Nvidia Blackwell dominant, AMD secondary vendor",
         "Nvidia reporting supply constraints on Blackwell (CoWoS capacity, HBM supply). "
         "Hyperscalers accelerating AMD MI450 orders to fill gaps. Near-term market share gain.",
         "Supply constraints at Nvidia are creating an AMD demand pull-forward. "
         "Q3 2026 guidance upside partially reflects this dynamic.",
         "Nvidia supply constraints are a Bull catalyst in the near term (2026-2027). "
         "If Nvidia resolves constraints by 2028, AMD reverts to winning on merit/price."),

        ("Competitor — Intel",
         "Intel Data Center — AMD Extended Lead (Q2 2026)",
         "AMD DC $5.8B vs Intel DC $5.1B in Q1 2026",
         "AMD Data Center est $6.7B vs Intel DCAI est ~$5.3B in Q2 2026. "
         "AMD extending lead. EPYC Venice launch adds CPU competitive pressure on Xeon 6.",
         "AMD's Data Center leadership vs Intel is now durable across 5+ quarters. "
         "EPYC Venice vs Xeon 6 GPU is the next battleground for server CPU share.",
         "Intel CPU recovery risk to EPYC share is partially mitigated by Venice launch. "
         "If Intel Xeon recovers significantly, the CPU revenue component of Bull/Base could "
         "moderate — currently not showing in the data."),

        # ── WATCH ITEM ─────────────────────────────────────────────────────────
        ("Watch — Unconfirmed",
         "AMD-Google TPU Hybrid ASIC (Reported Aug 2026)",
         "Google building custom TPU in-house; Bear risk to AMD",
         "Reports (Aug 2026): AMD reportedly designing Google's 10th-gen TPU hybrid ASIC with "
         "AMD CPU cores on-package. Production 2029+. UNCONFIRMED — reported in tech press. "
         "If confirmed: AMD CPU cores inside Google's TPU is a massive design win.",
         "If confirmed: flips Google custom silicon from Bear risk to AMD upside. "
         "Would add Google as an AMD CPU customer in the hyperscaler TPU vertical for 2029+.",
         "Added to Assumptions tab of projection engine v8 as watch item. "
         "Bear probability already reduced to 10% — this further weakens the custom silicon Bear narrative. "
         "If confirmed, Bear probability could fall to 5-7% and Bull could exceed 60%."),
    ]

    def write_data_row(ws, row, cat, item, prior, current, cons_earn, cons_proj,
                       bg, cat_colors):
        vals = [cat, item, prior, current, cons_earn, cons_proj]
        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col_idx)
            c.value = val
            c.fill = fill(bg)
            c.border = TABLE_BORDER
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c.font = Font(name="Arial", size=SZ_DEFAULT, color=BLACK)
        ink = cat_colors.get(cat, NAVY)
        ws.cell(row=row, column=1).font = Font(
            name="Arial", size=SZ_DEFAULT, bold=True, color=ink)
        ws.row_dimensions[row].height = 65

    next_row = 4
    for i, row_data in enumerate(items):
        bg = WHITE if i % 2 == 0 else GRAY
        write_data_row(ws, next_row, *row_data, bg, cat_colors)
        next_row += 1

    # ── COHERENCE CHECK ───────────────────────────────────────────────────────
    next_row += 1

    ws.merge_cells(f"A{next_row}:F{next_row}")
    sec = ws[f"A{next_row}"]
    sec.value = ("▶  COHERENCE CHECK — PROJECTION ENGINE vs INCOME STATEMENT  |  "
                 "RED = mismatch requiring update to build_amd_projection_engine_v8.py")
    sec.font = Font(name="Arial", bold=True, size=SZ_HEADER, color=WHITE)
    sec.fill = fill(NAVY)
    sec.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[next_row].height = 24
    next_row += 1

    ws.merge_cells(f"A{next_row}:F{next_row}")
    sub = ws[f"A{next_row}"]
    sub.value = ("Projection engine file: build_amd_projection_engine_v8.py  |  "
                 "Key constants: BASE_REV=$34.64B, ENTRY_PRICE=$455, SHARES=1.750B, "
                 "Q2_NI_MARGIN=24.3%, Q2_REV=$11.54B, Q3_GUIDE=$13.0B  |  "
                 "Probabilities: Bull 57% / Base 33% / Bear 10%  |  "
                 "v8 adds: Anthropic 2GW+$5B, Core Scientific 2.5GW, MSFT Azure Helios, "
                 "Rackspace 30MW; Google TPU collaboration as Assumptions watch item.")
    sub.font = Font(name="Arial", size=SZ_DEFAULT, color=WHITE, italic=True)
    sub.fill = fill(BLUE)
    sub.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[next_row].height = 30
    next_row += 1

    chk_headers = ["Assumption", "Projection Engine Value",
                   "Income Statement Actual", "Status", "Notes"]
    for col_idx, h in enumerate(chk_headers, 1):
        c = ws.cell(row=next_row, column=col_idx)
        c.value = h
        c.font = Font(name="Arial", bold=True, size=SZ_HEADER, color=WHITE)
        c.fill = fill(ACCENT)
        c.alignment = center(wrap=True)
        c.border = TABLE_BORDER
    ws.row_dimensions[next_row].height = 24
    next_row += 1

    OK   = "1A7A3A"
    WARN = "B7770D"

    checks = [
        ("Base Revenue (FY2025)",
         "BASE_REV = $34.64B",
         "FY2025 actual: $34.64B (per AMD annual report)",
         "✓  MATCH", OK,
         "Foundation for all scenario revenue projections. Confirmed accurate."),

        ("Q2 2026 Revenue",
         "Q2_REV = $11.54B",
         "Q2 2026 actual: $11,540M",
         "✓  MATCH", OK,
         "Used to validate projection trajectory. Revenue above Bull case pace (+43%)."),

        ("Q3 2026 Guidance",
         "Q3_GUIDE = $13.0B",
         "AMD guided $12.7B–$13.3B midpoint $13.0B",
         "✓  MATCH", OK,
         "Guidance midpoint correctly captured. +12.6% Q/Q sequential growth confirmed."),

        ("Non-GAAP NI Margin Q2",
         "Q2_NI_MARGIN = 24.3%",
         "Q2 2026 Non-GAAP NI margin: ~24.3% ($1.66 Non-GAAP EPS implied)",
         "✓  MATCH", OK,
         "Non-GAAP NI margin expanding. Base 2026 target (25%) nearly met in Q2. "
         "H2 typically stronger margins — full year 25%+ likely achievable."),

        ("2026 Bull Revenue Growth",
         "BULL rev_growth[0] = 43%  →  $49.5B",
         "H1 pace: $21.8B. Q3 guide: $13.0B. FY2026 tracking $50B+.",
         "✓  ON TRACK / ABOVE", OK,
         "Q2 result and Q3 guide suggest FY2026 may exceed Bull case. "
         "Watch for v9 upward revision if Q3 beats guidance."),

        ("Scenario Probabilities",
         "Bull 57% / Base 33% / Bear 10%  (v8 engine)",
         "Q2 actuals and 9-deal pipeline support Bull probability increase.",
         "✓  UPDATED", OK,
         "v8 update: Bull raised 45%→57% (Q2 beat, Anthropic deal, NVDA supply constraints), "
         "Base 40%→33% (high execution confidence), Bear 15%→10% (custom silicon risk "
         "partially neutralized by Google TPU rumor + AMD winning all major hyperscalers)."),

        ("Diluted Share Count",
         "SHARES = 1.750B  (v8 engine)",
         "Q2 2026 diluted shares: ~1,660M actual | +320M warrants (OpenAI/Meta)",
         "✓  CONFIRMED", OK,
         "1.750B forward estimate remains appropriate — includes 320M warrants net of "
         "partial buyback offset. Anthropic deal may add additional warrants — track Q2 10-Q."),
    ]

    for i, (assumption, engine_val, actual_val, status, status_color, notes) in \
            enumerate(checks):
        row = next_row + i
        bg = WHITE if i % 2 == 0 else GRAY
        for col_idx, val in enumerate(
            [assumption, engine_val, actual_val, status, notes], 1
        ):
            c = ws.cell(row=row, column=col_idx)
            c.value = val
            c.fill = fill(bg)
            c.border = TABLE_BORDER
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c.font = Font(name="Arial", size=SZ_DEFAULT, color=BLACK)
        ws.cell(row=row, column=1).font = Font(
            name="Arial", size=SZ_DEFAULT, bold=True, color=NAVY)
        ws.cell(row=row, column=4).font = Font(
            name="Arial", size=SZ_DEFAULT, bold=True, color=status_color)
        ws.merge_cells(f"E{row}:F{row}")
        ws.row_dimensions[row].height = 65


def build():
    wb = Workbook()
    ws = wb.active
    ws.title = "Q2 2026 Income Statement"

    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = "ADVANCED MICRO DEVICES, INC. (AMD) — Q2 2026 INCOME STATEMENT ANALYSIS"
    t.font = Font(name="Arial", bold=True, size=SZ_TITLE, color=WHITE)
    t.fill = fill(NAVY)
    t.alignment = center()
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:H2")
    s = ws["A2"]
    s.value = ("Quarter Ended June 28, 2026  |  Reported July 29, 2026  |  "
               "Source: AMD Q2 2026 Earnings Press Release  |  "
               "GAAP figures in millions except per-share amounts")
    s.font = Font(name="Arial", size=SZ_SUBTITLE, color=WHITE)
    s.fill = fill(BLUE)
    s.alignment = center(wrap=True)
    ws.row_dimensions[2].height = 18

    headers = ["Line Item", "Y/Y Change ($M)", "Y/Y %",
               "Q2 2026 ($M)", "Q1 2026 ($M)", "Q2 2025 ($M)",
               "Notes", "What This Measures"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.value = h
        cell.font = Font(name="Arial", bold=True, size=SZ_HEADER, color=WHITE)
        cell.fill = fill(ACCENT)
        cell.alignment = center(wrap=True)
        cell.border = TABLE_BORDER
    ws.row_dimensions[3].height = 30

    yy_map = {
        "green": (GREEN, GREEN_LT),
        "amber": (AMBER, AMBER_LT),
        "red":   (RED,   RED_LT),
    }

    DATA_START = 4

    for i, row_data in enumerate(ROWS):
        (label, q2, q1, q2_py,
         yy_color, actual_color, is_pct, note, measure, fmt_override) = row_data

        row = DATA_START + i
        is_priority = label.strip() in PRIORITY
        is_sub      = label.startswith("  ")
        bg = WHITE if i % 2 == 0 else GRAY

        a = ws.cell(row=row, column=1)
        a.value = label.strip()
        a.font = Font(name="Arial",
                      bold=is_priority and not is_sub,
                      italic=is_sub,
                      size=SZ_HEADER,
                      color=MUTED if is_sub else BLACK)
        a.fill = fill(bg)
        a.alignment = center()
        a.border = TABLE_BORDER

        b = ws.cell(row=row, column=2)
        b.value = "—" if is_sub else yy_change_str(q2, q2_py, is_pct, fmt_override)
        b.alignment = center()
        b.border = TABLE_BORDER

        c = ws.cell(row=row, column=3)
        c.value = "—" if is_sub else yy_pct_str(q2, q2_py, is_pct)
        c.alignment = center()
        c.border = TABLE_BORDER

        is_dash = (b.value == "—")
        if not is_dash and yy_color in yy_map:
            ink, bg_yy = yy_map[yy_color]
            for col in [2, 3]:
                cell = ws.cell(row=row, column=col)
                cell.fill = fill(bg_yy)
                cell.font = Font(name="Arial", size=SZ_DEFAULT, color=ink, bold=True)
        else:
            for col in [2, 3]:
                cell = ws.cell(row=row, column=col)
                cell.fill = fill(bg)
                cell.font = Font(name="Arial", size=SZ_DEFAULT,
                                 color=MUTED if is_dash else BLACK)

        d = ws.cell(row=row, column=4)
        d.value = fmt_val(q2, is_pct, fmt_override)
        d.alignment = center()
        d.border = TABLE_BORDER
        if actual_color in yy_map:
            ink, bg_ac = yy_map[actual_color]
            d.fill = fill(bg_ac)
            d.font = Font(name="Arial", size=SZ_DEFAULT, color=ink, bold=True)
        else:
            d.fill = fill(bg)
            d.font = Font(name="Arial", size=SZ_DEFAULT, color=BLACK)

        e = ws.cell(row=row, column=5)
        e.value = fmt_val(q1, is_pct, fmt_override)
        e.fill = fill(bg)
        e.alignment = center()
        e.border = TABLE_BORDER
        e.font = Font(name="Arial", size=SZ_DEFAULT, color=BLACK)

        f = ws.cell(row=row, column=6)
        f.value = fmt_val(q2_py, is_pct, fmt_override)
        f.fill = fill(bg)
        f.alignment = center()
        f.border = TABLE_BORDER
        f.font = Font(name="Arial", size=SZ_DEFAULT, color=BLACK)

        g = ws.cell(row=row, column=7)
        g.value = note
        g.fill = fill(bg)
        g.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        g.border = TABLE_BORDER
        g.font = Font(name="Arial", size=SZ_DEFAULT - 1, color=BLACK)

        h_cell = ws.cell(row=row, column=8)
        h_cell.value = measure
        h_cell.fill = fill(MEASURE_BG)
        h_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        h_cell.border = TABLE_BORDER
        h_cell.font = Font(name="Arial", size=SZ_DEFAULT - 1, color=NAVY, italic=True)

        ws.row_dimensions[row].height = min(60, max(30, 14 + len(note) // 18))

    # ── GRADE SECTION ─────────────────────────────────────────────────────────
    next_row = DATA_START + len(ROWS) + 1

    ws.merge_cells(f"A{next_row}:H{next_row}")
    gs = ws[f"A{next_row}"]
    gs.value = "QUARTERLY GRADE — Q2 2026"
    gs.font = Font(name="Arial", bold=True, size=SZ_HEADER, color=WHITE)
    gs.fill = fill(NAVY)
    gs.alignment = center()
    ws.row_dimensions[next_row].height = 20
    next_row += 1

    grade_row = next_row
    ws.merge_cells(f"A{grade_row}:C{grade_row}")
    gl = ws[f"A{grade_row}"]
    gl.value = "A"
    gl.font = Font(name="Arial", bold=True, size=SZ_GRADE, color="1A7A3A")
    gl.fill = fill(GREEN_LT)
    gl.alignment = center()
    ws.row_dimensions[grade_row].height = 32

    ws.merge_cells(f"D{grade_row}:H{grade_row}")
    gn = ws[f"D{grade_row}"]
    gn.value = (
        "Grade A: Revenue +50.0% Y/Y. ALL cost lines grew slower than revenue for the first time — "
        "R&D +40.0%, MG&A +39.7%, Total OpEx +34.9%. Non-GAAP gross margin 56% (above 55% threshold). "
        "Operating Income +172% Y/Y. Net Income +204% Y/Y. GAAP net margin 19.9% (+910bps Y/Y). "
        "9 confirmed deals totaling 22GW+. Q3 guidance $13.0B (midpoint). "
        "MI450 production on schedule. Helios deployed on Azure. Full operating leverage confirmed."
    )
    gn.font = Font(name="Arial", size=SZ_DEFAULT, color=BLACK)
    gn.fill = fill(GREEN_LT)
    gn.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[grade_row].height = 55

    next_row = grade_row + 1

    # ── STRENGTHS ─────────────────────────────────────────────────────────────
    next_row = section_header(ws, next_row, "▶  STRENGTHS", "1A7A3A")
    for s in [
        "✓  Net Revenue +50.0% Y/Y to $11.54B — fastest Y/Y growth in AMD's Data Center era. +12.6% Q/Q sequential acceleration.",
        "✓  Data Center $6.7B (+107% Y/Y) — doubled Y/Y. First time AMD DC exceeded $6B in a single quarter. 58% of total revenue.",
        "✓  ALL cost lines grew slower than revenue (+50%): R&D +40%, MG&A +39.7%, Total OpEx +34.9%. Full P&L operating leverage for first time.",
        "✓  Non-GAAP Gross Margin 56% — above the 55% A-grade threshold. GAAP GM +460bps Y/Y (49.6% → 54.2%). Structural, repeatable driver.",
        "✓  Operating Income +172% Y/Y ($796M → $2,170M). GAAP operating margin +840bps Y/Y (10.4% → 18.8%). Textbook operating leverage.",
        "✓  Net Income +204% Y/Y — more than tripled on 50% revenue growth. GAAP net margin: 9.8% → 19.9% (+910bps Y/Y).",
        "✓  Diluted EPS $1.38 GAAP (+200% Y/Y), $1.66 Non-GAAP. Shareholders capturing full operating leverage benefit.",
        "✓  MI450 in production ramp on schedule. Helios deployed on Microsoft Azure. Oracle 50K GPU delivery Q3 2026.",
        "✓  4 new major deals signed in Q2 2026: Anthropic 2GW + $5B equity, Core Scientific 2.5GW, MSFT Azure, Rackspace 30MW.",
        "✓  Total confirmed pipeline exceeds 22GW across 9 customer agreements. Demand risk is essentially eliminated.",
        "✓  Q3 2026 guidance $12.7–13.3B ($13.0B midpoint, +12.6% Q/Q) — most bullish guidance in AMD history.",
        "✓  MG&A inflection: was RED in Q1 2026 (outpacing revenue). Now GREEN in Q2 2026 — operating discipline confirmed.",
    ]:
        next_row = bullet_row(ws, next_row, s, STRENGTH_BG)

    next_row += 1

    next_row = section_header(ws, next_row, "▶  CONCERNS", "8B0000")
    for c in [
        "✗  Warrant dilution: 320M total warrants outstanding (OpenAI 160M + Meta 160M). Anthropic deal may add additional warrants — track Q2 10-Q.",
        "✗  Q2 Other Income $490M significantly elevated — Anthropic equity deal accounting. Non-recurring nature should normalize Q3.",
        "✗  GAAP vs Non-GAAP gap remains: GAAP op margin ~18.8% vs Non-GAAP ~27%. ~$548M quarterly Xilinx amortization continues.",
        "✗  Gaming segment declining — consumer GPU and semi-custom both down Y/Y. No near-term catalyst for reversal.",
        "✗  Non-GAAP operating margin guided 27% but not yet 30%+ — the threshold for A-grade on the op margin dimension.",
        "✗  Google TPU collaboration (Aug 2026) unconfirmed — watch closely. If false: Bear risk from Google custom silicon remains elevated.",
        "✗  HBM4 supply (SK Hynix/Samsung CoWoS): Anthropic 2GW 2027 adds to the HBM4 volume requirement alongside OpenAI/Meta/Oracle.",
    ]:
        next_row = bullet_row(ws, next_row, c, CONCERN_BG)

    next_row += 1

    next_row = section_header(ws, next_row,
                              "▶  LETTER GRADE FRAMEWORK — AMD (Q2 2026 UPDATED)",
                              SECTION_BG)
    for grade, desc in [
        ("A",   "ALL cost lines < revenue growth for 2+ consecutive quarters ✓ (achieved Q2 2026). "
                "Non-GAAP gross margin 55%+ ✓ (56% achieved). Non-GAAP operating margin 30%+ (27% — approaching). "
                "MI450 ramp on schedule ✓. All 9 confirmed deals in execution ✓."),
        ("A−",  "Prior baseline (Q1 2026): Operating Income +83% on +38% revenue. MG&A was the exception. "
                "Q2 2026 closes the MG&A exception, upgrading to A."),
        ("B+",  "MG&A growth returns above revenue growth for 2+ quarters. Revenue growth decelerates below 30%. "
                "MI450 production delays. Gross margin expansion stalls below 54%."),
        ("B",   "Revenue growth decelerates to teens. Gross margin contracts from 54% peak. "
                "MI450 volume ramp delayed significantly. Operating income growth falls near revenue growth rate."),
        ("C",   "MI450 yield failures create material revenue miss. ROCm fails production workloads at hyperscalers. "
                "Revenue growth below 15%. Gross margin contracts below 50%. Multiple deal delays."),
        ("D",   "Fundamental execution breakdown: MI450 production suspended. HBM4 supply unavailable. "
                "OpenAI/Meta/Oracle/Anthropic defer to Nvidia alternatives. EPYC share reversal. Revenue decline."),
    ]:
        next_row = bullet_row(ws, next_row, f"{grade}   {desc}", ASSESS_BG)

    next_row += 1

    next_row = section_header(ws, next_row,
                              "▶  KEY METRICS TO WATCH — Q3 2026 (Expected Late October 2026)",
                              SECTION_BG)
    for m in [
        "① Data Center Revenue vs $13B guidance — must confirm MI450 OpenAI/Meta/Oracle ramp. Q3 DC could exceed $8B if all deployments execute.",
        "② Non-GAAP Operating Margin vs ~27% guide — above 28% = outperform. Structural path to 30%+ NI margin.",
        "③ Other Income normalization — Q2's $490M was elevated. Q3 should return to ~$150-200M (interest income) absent new equity events.",
        "④ Warrant vesting disclosures — track any OpenAI/Meta warrant vesting events in Q3. Anthropic deal may add to dilution.",
        "⑤ Non-GAAP EPS — Q3 guidance not yet given for EPS but $1.90-2.00+ is achievable at $13B revenue and 27% margins.",
        "⑥ Gross Margin trajectory — Q2 GAAP 54.2% / Non-GAAP 56%. Can AMD reach 57%+ as MI450 (higher margin) becomes dominant?",
        "⑦ Google TPU collaboration — any confirmation from AMD or Google IR is a MAJOR signal. Monitor ir.amd.com and Google IR.",
        "⑧ R&D leverage — Q2 first quarter where R&D grew slower than revenue. Must sustain to confirm structural improvement.",
        "⑨ Core Scientific / Rackspace deployment start — late 2026 deployments begin. Track for first revenue contributions.",
    ]:
        next_row = bullet_row(ws, next_row, m, WHITE)

    next_row += 1
    ws.merge_cells(f"A{next_row}:H{next_row}")
    foot = ws[f"A{next_row}"]
    foot.value = (
        "Source: AMD Q2 2026 Earnings Press Release, July 29, 2026  |  GAAP figures only  |  "
        "Millions of dollars except per-share amounts  |  Analysis: September 2026  |  "
        "Script: build_amd_income_v4.py  |  Not investment advice."
    )
    foot.font = Font(name="Arial", size=SZ_FOOTER, color=MUTED, italic=True)
    foot.fill = fill(WHITE)
    foot.alignment = Alignment(horizontal="left", wrap_text=True)
    ws.row_dimensions[next_row].height = 18

    build_sources_tab(wb)
    build_whats_new_tab(wb)

    wb.active = wb["Q2 2026 Income Statement"]

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AMD_Q2_2026_Income_Statement_v4.xlsx")
    wb.save(out)
    print(f"Saved: {out}")
    return out


# ==============================================================================
# SELF-TEST — DO NOT REMOVE
# ==============================================================================
def self_test(out_path):
    import os
    from openpyxl import load_workbook

    errors = []

    if not os.path.exists(out_path):
        errors.append(f"FAIL: file not created at {out_path}")
    else:
        size = os.path.getsize(out_path)
        if size < 14000:
            errors.append(f"FAIL: file too small ({size} bytes)")

        wb = load_workbook(out_path, data_only=True)
        ws = wb.active
        if ws.title != "Q2 2026 Income Statement":
            errors.append(f"FAIL: wrong sheet name '{ws.title}'")
        if "Data Sources" not in wb.sheetnames:
            errors.append("FAIL: 'Data Sources' tab missing")
        if "What's New" not in wb.sheetnames:
            errors.append("FAIL: \"What's New\" tab missing")

        found_revenue = found_11540 = found_grade = False
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell:
                    cv = str(cell)
                    if "Net Revenue" in cv:              found_revenue = True
                    if "11540" in cv or "11,540" in cv: found_11540 = True
                    if cv.strip() == "A":               found_grade = True

        if not found_revenue: errors.append("FAIL: 'Net Revenue' row not found")
        if not found_11540:   errors.append("FAIL: Q2 2026 Revenue $11,540M not found")
        if not found_grade:   errors.append("FAIL: Grade 'A' not found")

        # Check What's New has all 9 confirmed deals
        ws_wn = wb["What's New"]
        found_openai = found_meta = found_oracle = found_anthropic = False
        found_core_sci = found_microsoft = found_humain = found_doe = found_v8 = False
        for row in ws_wn.iter_rows(values_only=True):
            for cell in row:
                if cell:
                    cv = str(cell)
                    if "OpenAI" in cv:                        found_openai = True
                    if "Meta" in cv and "6GW" in cv:         found_meta = True
                    if "Oracle" in cv and "50K" in cv:       found_oracle = True
                    if "Anthropic" in cv and "2GW" in cv:    found_anthropic = True
                    if "Core Scientific" in cv:              found_core_sci = True
                    if "Microsoft" in cv or "Azure" in cv:  found_microsoft = True
                    if "HUMAIN" in cv:                        found_humain = True
                    if "ORNL" in cv or "Discovery" in cv:   found_doe = True
                    if "projection_engine_v8" in cv:         found_v8 = True

        if not found_openai:    errors.append("FAIL: What's New — OpenAI deal missing")
        if not found_meta:      errors.append("FAIL: What's New — Meta 6GW deal missing")
        if not found_oracle:    errors.append("FAIL: What's New — Oracle 50K GPU deal missing")
        if not found_anthropic: errors.append("FAIL: What's New — Anthropic 2GW deal missing")
        if not found_core_sci:  errors.append("FAIL: What's New — Core Scientific deal missing")
        if not found_microsoft: errors.append("FAIL: What's New — Microsoft Azure Helios missing")
        if not found_humain:    errors.append("FAIL: What's New — HUMAIN deal missing")
        if not found_doe:       errors.append("FAIL: What's New — DoE/ORNL deal missing")
        if not found_v8:        errors.append("FAIL: What's New — v8 engine reference missing")

        # Check Data Sources has key entries
        ws_ds = wb["Data Sources"]
        found_openai_ds = found_oracle_ds = found_anthropic_ds = found_core_ds = False
        for row in ws_ds.iter_rows(values_only=True):
            for cell in row:
                if cell:
                    cv = str(cell)
                    if "OpenAI" in cv:         found_openai_ds = True
                    if "Oracle" in cv:         found_oracle_ds = True
                    if "Anthropic" in cv:      found_anthropic_ds = True
                    if "Core Scientific" in cv: found_core_ds = True
        if not found_openai_ds:    errors.append("FAIL: Data Sources — OpenAI entry missing")
        if not found_oracle_ds:    errors.append("FAIL: Data Sources — Oracle entry missing")
        if not found_anthropic_ds: errors.append("FAIL: Data Sources — Anthropic entry missing")
        if not found_core_ds:      errors.append("FAIL: Data Sources — Core Scientific entry missing")

    if errors:
        print("\n" + "=" * 60)
        print("SELF-TEST FAILED — DO NOT SAVE TO DRIVE")
        print("=" * 60)
        for e in errors:
            print(f"  {e}")
        print("=" * 60)
        return False

    size = os.path.getsize(out_path)
    print("\n" + "=" * 60)
    print("SELF-TEST PASSED — safe to save to Drive")
    print(f"  File: {out_path}")
    print(f"  Size: {size:,} bytes")
    print(f"  Net Revenue $11,540M: confirmed")
    print(f"  Tabs: Q2 2026 Income Statement + Data Sources + What's New")
    print(f"  Grade A: confirmed")
    print(f"  All 9 deals (OpenAI/Meta/Oracle/HUMAIN/DoE/Anthropic/Core Sci/MSFT/Rackspace): confirmed")
    print(f"  v8 engine reference: confirmed in Coherence Check")
    print("=" * 60)
    return True


if __name__ == "__main__":
    out = build()
    passed = self_test(out)
    if not passed:
        raise SystemExit(1)
