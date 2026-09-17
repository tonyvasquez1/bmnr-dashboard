"""
AMD Monitoring Agent
--------------------
Utility script for Claude-driven daily monitoring of AMD financials, deals, and
competitive landscape.

Claude Code runs this script for state management, Excel generation, and
formatted output. Claude handles web searching (WebSearch tool) and MCP calls
(Gmail, Drive).

Usage:
    python amd_agent.py report              # 8-quarter projection + current state
    python amd_agent.py detect-quarter      # Check if new quarter available
    python amd_agent.py carry-scores        # Output current scores as JSON
    python amd_agent.py add-finding <json>  # Add a monitoring finding to log
    python amd_agent.py generate-log-xlsx   # Create AMD_Monitor_Log.xlsx
    python amd_agent.py email-body          # Full formatted email body (stdout)
    python amd_agent.py send-email          # Build + send email via SMTP (needs GMAIL_APP_PASSWORD)
    python amd_agent.py update-quarter <json_file>  # Update last_run.json with new quarter
    python amd_agent.py self-test           # Verify all modes run cleanly

Requires: openpyxl  (auto-installed if missing)
"""

import json
import os
import sys
import subprocess
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
LAST_RUN_PATH    = os.path.join(DATA_DIR, "last_run.json")
MONITOR_LOG_PATH = os.path.join(DATA_DIR, "monitor_log.json")
EVENT_WATCH_PATH = os.path.join(DATA_DIR, "event_watch.json")


# ── auto-install ───────────────────────────────────────────────────────────────
def _ensure(pkg):
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure("openpyxl")
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════════════
# DATA BLOCK — updated each quarter by Claude; carried forward from last_run.json
# ══════════════════════════════════════════════════════════════════════════════

VERSION     = "v1.0 — Q2 2026 Baseline"
AS_OF       = "Q2 2026"
LAST_FILING = "2026-07-29"

# Current quarter financials
Q2_REV       = 11.540   # B
Q2_DC_REV    = 6.700    # B
Q2_DC_GROWTH = 1.07     # Y/Y
Q2_NI_MARGIN = 0.243
Q2_EPS_GAAP  = 1.38
Q2_EPS_NONGAAP = 1.66
Q3_GUIDE_MID = 13.000   # B

# Projection parameters (from build_amd_projection_engine_v8.py)
ENTRY_PRICE = 455.00
BASE_REV    = 34.64     # FY 2025 actual revenue $B
SHARES      = 1.750     # B diluted shares

PROB_BULL   = 0.57
PROB_BASE   = 0.33
PROB_BEAR   = 0.10

BULL = dict(
    rev_growth = [0.43, 0.55, 0.55, 0.50, 0.45],
    ni_margin  = [0.28, 0.32, 0.38, 0.41, 0.43],
    pe_low     = [45, 48, 50, 50, 48],
    pe_high    = [55, 58, 60, 60, 58],
)
BASE = dict(
    rev_growth = [0.36, 0.45, 0.43, 0.42, 0.40],
    ni_margin  = [0.25, 0.26, 0.26, 0.28, 0.29],
    pe_low     = [35, 36, 37, 38, 40],
    pe_high    = [40, 41, 42, 43, 45],
)
BEAR = dict(
    rev_growth = [0.32, 0.18, 0.12, 0.08, 0.10],
    ni_margin  = [0.22, 0.21, 0.22, 0.24, 0.26],
    pe_low     = [25, 20, 17, 15, 16],
    pe_high    = [30, 25, 22, 20, 21],
)

# Qualitative scores (Q2 2026)
SCORES = {
    "SCORE_NEW_DEAL_VELOCITY": 10,
    "SCORE_ROCM_PRODUCTION":   7,
    "SCORE_COMPETITIVE":       8,
    "SCORE_GUIDANCE_TONE":     9,
    "SCORE_OPENAI_DELIVERY":   8,
    "SCORE_META_DELIVERY":     8,
    "SCORE_HUMAIN_DELIVERY":   7,
}

# Confirmed deals pipeline
CONFIRMED_DEALS = [
    ("HUMAIN",           "500 MW",  "Delivery ongoing — GW scale expansion expected"),
    ("OpenAI",           "6 GW",    "MI450 series; active deployment"),
    ("Meta",             "6 GW",    "Custom MI450-based GPU; multi-year program"),
    ("Oracle",           "50K GPU", "MI450 + EPYC Venice; Q2 2026 active"),
    ("DoE / ORNL",       "HPC",     "Lux AI + Discovery systems"),
    ("Anthropic",        "2 GW",    "MI450; $5B AMD equity stake; 1 GW starts H1 2027"),
    ("Core Scientific",  "2.5 GW",  "Capacity secured; deployments begin 2027"),
    ("Microsoft/Azure",  "Helios",  "Rack-scale system deployed on Azure"),
    ("Rackspace",        "30 MW",   "Phased deployment late 2026–2028"),
]

WATCH_ITEMS = [
    ("Google TPU hybrid ASIC", "UNCONFIRMED — Aug 2026 reports AMD designing 10th-gen TPU. "
     "Bear risk (custom silicon) may become upside. Confirm before moving to confirmed deals."),
]

# Monitoring sources
DAILY_SOURCES = [
    ("AMD IR press releases",    "https://ir.amd.com/press-releases"),
    ("AMD newsroom",             "https://www.amd.com/en/newsroom"),
    ("SEC EDGAR 8-K filings",    "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000002488&type=8-K&dateb=&owner=include&count=5"),
]
WEEKLY_SOURCES = [
    ("AMD IR events calendar",   "https://ir.amd.com/events"),
    ("Nvidia press releases",    "https://nvidianews.nvidia.com/"),
    ("Intel newsroom",           "https://www.intel.com/content/www/us/en/newsroom/home.html"),
]
QUARTERLY_SOURCES = [
    ("Mercury Research server CPU share", "Mercury Research"),
    ("TSMC manufacturing announcements",  "https://www.tsmc.com/english/news/pressReleases"),
]

DRIVE_AMD_FINANCIALS_FOLDER = "1i1GdOQGreQuxv2s6xoqILyY_XIqKuX5a"
DRIVE_Q2_2026_FOLDER        = "1D_GOhT8kgZBTlVPA_7VXvO0-lR5ZyTbB"
DRIVE_Q1_2026_FOLDER        = "1Df4g1sgPkwy_7BwBLjNTsYNDPh28RiWc"

GMAIL_TO = "tonyvasquez1@gmail.com"


# ══════════════════════════════════════════════════════════════════════════════
# QUARTERLY PROJECTION — 8-quarter table
# ══════════════════════════════════════════════════════════════════════════════

def _annual_to_quarterly(annual_rate):
    """Convert annual growth rate to approximate quarterly growth rate."""
    return (1 + annual_rate) ** 0.25 - 1


def build_8q_projection():
    """
    Build 8-quarter revenue projection for Bull / Base / Bear scenarios.
    Starting point: Q2 2026 actuals. Q3 2026 uses guided midpoint.
    Returns list of dicts: {quarter, bull_rev, base_rev, bear_rev}
    """
    # Quarter labels
    quarters = [
        "Q3 2026", "Q4 2026",
        "Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027",
        "Q1 2028", "Q2 2028",
    ]

    # Annual growth rates indexed by fiscal year offset (FY2026=0, FY2027=1, FY2028=2)
    # FY2026 covers Q3+Q4 2026, FY2027 covers Q1-Q4 2027, FY2028 starts Q1 2028
    fy_map = {
        "Q3 2026": 0, "Q4 2026": 0,
        "Q1 2027": 1, "Q2 2027": 1, "Q3 2027": 1, "Q4 2027": 1,
        "Q1 2028": 2, "Q2 2028": 2,
    }

    # Seasonal Q/Q patterns (roughly: Q3 up, Q4 up, Q1 down, Q2 recovery)
    # Applied on top of annual trend; simplified to keep numbers coherent
    seasonal = {
        "Q3 2026": ( 0.127,  0.127,  0.120),   # Q3 guide anchors this: $13.0B / $11.54B − 1
        "Q4 2026": ( 0.098,  0.080,  0.050),
        "Q1 2027": (-0.040, -0.050, -0.060),    # typical Q1 seasonal dip
        "Q2 2027": ( 0.130,  0.100,  0.060),
        "Q3 2027": ( 0.120,  0.095,  0.040),
        "Q4 2027": ( 0.105,  0.085,  0.030),
        "Q1 2028": (-0.035, -0.045, -0.055),
        "Q2 2028": ( 0.125,  0.098,  0.045),
    }

    rows = []
    bull_prev = Q2_REV
    base_prev = Q2_REV
    bear_prev = Q2_REV

    for q in quarters:
        bs, bs2, bs3 = seasonal[q]
        bull_rev = round(bull_prev * (1 + bs),  2)
        base_rev = round(base_prev * (1 + bs2), 2)
        bear_rev = round(bear_prev * (1 + bs3), 2)

        # Q3 2026: anchor to guide
        if q == "Q3 2026":
            bull_rev = Q3_GUIDE_MID * 1.02   # slightly above guide in bull
            base_rev = Q3_GUIDE_MID
            bear_rev = Q3_GUIDE_MID * 0.96   # slightly below in bear

        bull_prev = bull_rev
        base_prev = base_rev
        bear_prev = bear_rev

        rows.append(dict(
            quarter  = q,
            bull_rev = round(bull_rev, 2),
            base_rev = round(base_rev, 2),
            bear_rev = round(bear_rev, 2),
        ))

    return rows


def build_5y_terminal():
    """Build 5-year terminal values for email summary."""
    rows = []
    for i, yr in enumerate(range(2026, 2031)):
        rev_b = BASE_REV
        for g in BULL["rev_growth"][:i+1]:
            rev_b *= (1 + g)
        eps_b  = rev_b * BULL["ni_margin"][i] / SHARES
        spl_b  = eps_b * BULL["pe_low"][i]
        sph_b  = eps_b * BULL["pe_high"][i]

        rev_ba = BASE_REV
        for g in BASE["rev_growth"][:i+1]:
            rev_ba *= (1 + g)
        eps_ba  = rev_ba * BASE["ni_margin"][i] / SHARES
        spl_ba  = eps_ba * BASE["pe_low"][i]
        sph_ba  = eps_ba * BASE["pe_high"][i]

        rev_br = BASE_REV
        for g in BEAR["rev_growth"][:i+1]:
            rev_br *= (1 + g)
        eps_br  = rev_br * BEAR["ni_margin"][i] / SHARES
        spl_br  = eps_br * BEAR["pe_low"][i]
        sph_br  = eps_br * BEAR["pe_high"][i]

        ev_spl = round(PROB_BULL * spl_b + PROB_BASE * spl_ba + PROB_BEAR * spl_br, 0)
        ev_sph = round(PROB_BULL * sph_b + PROB_BASE * sph_ba + PROB_BEAR * sph_br, 0)

        rows.append(dict(
            year     = yr,
            bull_rev = round(rev_b,  1), bull_spl = round(spl_b,  0), bull_sph = round(sph_b,  0),
            base_rev = round(rev_ba, 1), base_spl = round(spl_ba, 0), base_sph = round(sph_ba, 0),
            bear_rev = round(rev_br, 1), bear_spl = round(spl_br, 0), bear_sph = round(sph_br, 0),
            ev_spl   = ev_spl, ev_sph = ev_sph,
        ))
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# STATE MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

def _ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def load_last_run():
    _ensure_data_dir()
    if not os.path.exists(LAST_RUN_PATH):
        return _default_last_run()
    with open(LAST_RUN_PATH) as f:
        return json.load(f)


def _default_last_run():
    return {
        "last_quarter":      AS_OF,
        "last_filing_date":  LAST_FILING,
        "last_run_date":     None,
        "scores":            SCORES,
        "probabilities":     {"bull": PROB_BULL, "base": PROB_BASE, "bear": PROB_BEAR},
        "version":           VERSION,
    }


def save_last_run(data):
    _ensure_data_dir()
    with open(LAST_RUN_PATH, "w") as f:
        json.dump(data, f, indent=2)


def load_monitor_log():
    _ensure_data_dir()
    if not os.path.exists(MONITOR_LOG_PATH):
        return []
    with open(MONITOR_LOG_PATH) as f:
        return json.load(f)


def save_monitor_log(log):
    _ensure_data_dir()
    with open(MONITOR_LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)


def load_event_watch():
    _ensure_data_dir()
    if not os.path.exists(EVENT_WATCH_PATH):
        return []
    with open(EVENT_WATCH_PATH) as f:
        return json.load(f)


def save_event_watch(events):
    _ensure_data_dir()
    with open(EVENT_WATCH_PATH, "w") as f:
        json.dump(events, f, indent=2)


# ══════════════════════════════════════════════════════════════════════════════
# DETECT QUARTER
# ══════════════════════════════════════════════════════════════════════════════

def detect_quarter():
    """
    Check whether a new quarter is available since last run.
    Claude must search SEC EDGAR (via WebSearch) for the latest 10-Q/8-K filing.
    This function checks local state only.

    Returns dict with:
        last_quarter: what we've already processed
        last_filing_date: date of last processed filing
        days_since_run: days since last run
        action: "check_for_new_quarter" | "no_action_needed" | "new_quarter_available"
    """
    lr = load_last_run()
    today = datetime.date.today().isoformat()
    last_run = lr.get("last_run_date")

    if last_run is None:
        days_since = 999
    else:
        try:
            d = datetime.date.fromisoformat(last_run)
            days_since = (datetime.date.today() - d).days
        except Exception:
            days_since = 999

    return {
        "last_quarter":       lr.get("last_quarter", AS_OF),
        "last_filing_date":   lr.get("last_filing_date", LAST_FILING),
        "last_run_date":      last_run,
        "days_since_run":     days_since,
        "action":             "check_for_new_quarter" if days_since >= 1 else "no_action_needed",
        "instruction":        (
            "Search SEC EDGAR for AMD latest 10-Q or 8-K earnings filing. "
            "Compare filing date to last_filing_date. If newer filing exists, "
            "a new quarter is available. Use WebSearch: "
            "\"AMD 10-Q OR 8-K site:sec.gov 2026\" or check "
            "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000002488&type=10-Q&dateb=&owner=include&count=5"
        ),
    }


# ══════════════════════════════════════════════════════════════════════════════
# MONITORING LOG MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

def add_finding(finding_json_str):
    """
    Add a monitoring finding to monitor_log.json.

    Expected JSON keys:
        date: ISO date string
        source: source URL or name
        title: headline
        classification: HIGH | MEDIUM | LOW
        disposition: action taken or "logged only"
        impact: what it means for the model
        confirmed: true | false (false = rumor/unconfirmed)
        files_affected: list of file names or []
    """
    finding = json.loads(finding_json_str)
    required = {"date", "source", "title", "classification", "disposition", "impact"}
    missing = required - set(finding.keys())
    if missing:
        print(f"ERROR: Missing required keys: {missing}", file=sys.stderr)
        sys.exit(1)

    # Defaults
    finding.setdefault("confirmed", True)
    finding.setdefault("files_affected", [])

    log = load_monitor_log()
    log.append(finding)
    save_monitor_log(log)

    lr = load_last_run()
    lr["last_run_date"] = datetime.date.today().isoformat()
    save_last_run(lr)

    print(f"Added finding: [{finding['classification']}] {finding['title']}")
    print(f"Total log entries: {len(log)}")


def update_quarter(json_file):
    """
    Update last_run.json with new quarter data.

    Expected JSON keys in file:
        quarter: e.g. "Q3 2026"
        filing_date: e.g. "2026-10-28"
        scores: dict of score keys/values (optional — carries forward if missing)
        probabilities: {bull, base, bear} (optional — carries forward)
    """
    with open(json_file) as f:
        data = json.load(f)

    lr = load_last_run()

    lr["last_quarter"]     = data.get("quarter", lr["last_quarter"])
    lr["last_filing_date"] = data.get("filing_date", lr["last_filing_date"])
    lr["last_run_date"]    = datetime.date.today().isoformat()

    if "scores" in data:
        lr["scores"].update(data["scores"])
    if "probabilities" in data:
        lr["probabilities"].update(data["probabilities"])

    save_last_run(lr)
    print(f"Updated last_run.json → quarter={lr['last_quarter']}, filing_date={lr['last_filing_date']}")


# ══════════════════════════════════════════════════════════════════════════════
# FORMATTED REPORT OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

def _fmt_b(val):
    return f"${val:.2f}B"


def _fmt_price(val):
    return f"${val:,.0f}"


def build_report_text():
    """Return the plain-text monitoring report body (used for email and terminal)."""
    lr      = load_last_run()
    today   = datetime.date.today().strftime("%B %d, %Y")
    quarter = lr.get("last_quarter", AS_OF)
    probs   = lr.get("probabilities", {"bull": PROB_BULL, "base": PROB_BASE, "bear": PROB_BEAR})
    scores  = lr.get("scores", SCORES)

    lines = []
    lines.append("═" * 72)
    lines.append(f"  AMD MONITOR — {today.upper()}")
    lines.append("═" * 72)
    lines.append("")
    lines.append(f"  Model as of: {quarter}  |  Entry price: ${ENTRY_PRICE:,.2f}  |  Shares: {SHARES:.3f}B")
    lines.append("")

    # ── Probabilities ─────────────────────────────────────────────────────────
    lines.append("─" * 72)
    lines.append("  SCENARIO PROBABILITIES")
    lines.append("─" * 72)
    pb = probs["bull"]
    pba = probs["base"]
    pbr = probs["bear"]
    lines.append(f"  Bull:  {pb*100:.0f}%   Base:  {pba*100:.0f}%   Bear:  {pbr*100:.0f}%")
    lines.append("")

    # ── Q2 Actuals ────────────────────────────────────────────────────────────
    lines.append("─" * 72)
    lines.append(f"  Q2 2026 ACTUALS  (Most Recent Quarter)")
    lines.append("─" * 72)
    lines.append(f"  Total Revenue:     {_fmt_b(Q2_REV)}   (+50% Y/Y)")
    lines.append(f"  Data Center Rev:   {_fmt_b(Q2_DC_REV)}   (+{Q2_DC_GROWTH*100:.0f}% Y/Y, {Q2_DC_REV/Q2_REV*100:.0f}% of total)")
    lines.append(f"  Non-GAAP GM:       56%")
    lines.append(f"  GAAP EPS:          ${Q2_EPS_GAAP:.2f}   Non-GAAP EPS: ${Q2_EPS_NONGAAP:.2f}")
    lines.append(f"  Q3 2026 Guide:     ${Q3_GUIDE_MID:.1f}B midpoint ($12.7B–$13.3B)")
    lines.append("")

    # ── Qualitative Scores ────────────────────────────────────────────────────
    lines.append("─" * 72)
    lines.append("  THESIS TRACKER SCORES  (carried from last_run.json)")
    lines.append("─" * 72)
    label_map = {
        "SCORE_NEW_DEAL_VELOCITY": "Deal Velocity",
        "SCORE_ROCM_PRODUCTION":   "ROCm Production",
        "SCORE_COMPETITIVE":       "Competitive Position",
        "SCORE_GUIDANCE_TONE":     "Guidance Tone",
        "SCORE_OPENAI_DELIVERY":   "OpenAI Delivery",
        "SCORE_META_DELIVERY":     "Meta Delivery",
        "SCORE_HUMAIN_DELIVERY":   "HUMAIN Delivery",
    }
    for key, label in label_map.items():
        val = scores.get(key, "?")
        lines.append(f"  {label:<26} {val}/10")
    lines.append("")

    # ── Confirmed Deals ───────────────────────────────────────────────────────
    lines.append("─" * 72)
    lines.append("  CONFIRMED DEALS  (9 total, 22+ GW pipeline)")
    lines.append("─" * 72)
    for name, scale, status in CONFIRMED_DEALS:
        lines.append(f"  • {name:<20} {scale:<12}  {status}")
    lines.append("")

    # ── Watch Items ───────────────────────────────────────────────────────────
    lines.append("─" * 72)
    lines.append("  WATCH ITEMS  (unconfirmed / developing)")
    lines.append("─" * 72)
    for name, note in WATCH_ITEMS:
        lines.append(f"  ⚠  {name}")
        lines.append(f"     {note}")
    lines.append("")

    # ── 8-Quarter Revenue Projection Table ────────────────────────────────────
    lines.append("═" * 72)
    lines.append("  8-QUARTER REVENUE PROJECTION  (from Q3 2026, all figures $B)")
    lines.append("═" * 72)
    lines.append(f"  {'Quarter':<12} {'Bull':>10} {'Base':>10} {'Bear':>10}  Notes")
    lines.append("  " + "─" * 68)
    proj = build_8q_projection()
    for row in proj:
        notes = ""
        if row["quarter"] == "Q3 2026":
            notes = "← guided $12.7–$13.3B"
        lines.append(
            f"  {row['quarter']:<12} "
            f"{'$'+str(row['bull_rev'])+'B':>10} "
            f"{'$'+str(row['base_rev'])+'B':>10} "
            f"{'$'+str(row['bear_rev'])+'B':>10}  {notes}"
        )
    lines.append("")
    lines.append(f"  Probabilities: Bull {probs['bull']*100:.0f}% / Base {probs['base']*100:.0f}% / Bear {probs['bear']*100:.0f}%")
    lines.append(f"  Entry: ${ENTRY_PRICE:,.2f}  |  Diluted shares: {SHARES:.3f}B")
    lines.append("")

    # ── 5-Year Share Price Targets ────────────────────────────────────────────
    lines.append("═" * 72)
    lines.append("  5-YEAR SHARE PRICE TARGETS  (EPS × P/E range)")
    lines.append("═" * 72)
    lines.append(f"  {'Year':<8} {'Bull SPL':>10} {'Bull SPH':>10} {'Base SPL':>10} {'Base SPH':>10} {'Bear SPL':>8} {'EV SPL':>8}")
    lines.append("  " + "─" * 68)
    five_y = build_5y_terminal()
    for row in five_y:
        lines.append(
            f"  {row['year']:<8} "
            f"{_fmt_price(row['bull_spl']):>10} "
            f"{_fmt_price(row['bull_sph']):>10} "
            f"{_fmt_price(row['base_spl']):>10} "
            f"{_fmt_price(row['base_sph']):>10} "
            f"{_fmt_price(row['bear_spl']):>8} "
            f"{_fmt_price(row['ev_spl']):>8}"
        )
    lines.append("")
    lines.append("  SPL = Share Price Low (EPS × PE Low)  |  SPH = Share Price High")
    lines.append("  EV SPL = Probability-weighted Expected Value at low P/E")
    lines.append("")

    # ── Footer ────────────────────────────────────────────────────────────────
    lines.append("─" * 72)
    lines.append("  Sources tracked:")
    for label, url in DAILY_SOURCES:
        lines.append(f"  [Daily]   {label}")
    for label, url in WEEKLY_SOURCES:
        lines.append(f"  [Weekly]  {label}")
    lines.append("")
    lines.append("  AMD_Monitor_Log.xlsx in Drive — AMD Financials / Q2 2026")
    lines.append("  AMD_Thesis_Tracker_v1.xlsx   in Drive — AMD Financials / Q2 2026")
    lines.append("─" * 72)
    lines.append("")

    return "\n".join(lines)


def send_email_smtp(subject, html_body, to_addr=GMAIL_TO):
    """
    Send HTML email via Gmail SMTP using App Password.
    Requires env var GMAIL_APP_PASSWORD (16-char password from myaccount.google.com/apppasswords).
    Optional env var GMAIL_FROM (defaults to same as GMAIL_TO).
    """
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    from_addr = os.environ.get("GMAIL_FROM", GMAIL_TO)
    password  = os.environ.get("GMAIL_APP_PASSWORD")

    if not password:
        raise RuntimeError(
            "GMAIL_APP_PASSWORD not set. "
            "Go to myaccount.google.com/apppasswords, create an App Password (requires 2-step verification), "
            "then set GMAIL_APP_PASSWORD=<16-char-password> in your environment or Claude Code env vars."
        )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = from_addr
    msg["To"]      = to_addr
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(from_addr, password)
        smtp.sendmail(from_addr, [to_addr], msg.as_string())

    print(f"Email sent → {to_addr}")
    print(f"Subject:    {subject}")


def build_email_subject(high_count=0, medium_count=0, low_count=0):
    today = datetime.date.today().strftime("%Y-%m-%d")
    return f"AMD Monitor — {today} | {high_count} HIGH · {medium_count} MEDIUM · {low_count} LOW"


def build_email_body(findings_today=None, drive_log_url="", drive_doc_url=""):
    """
    Build the full HTML email body.
    findings_today: list of finding dicts from today's monitoring run.
    """
    if findings_today is None:
        findings_today = []

    today = datetime.date.today().strftime("%B %d, %Y")
    lr    = load_last_run()
    probs = lr.get("probabilities", {"bull": PROB_BULL, "base": PROB_BASE, "bear": PROB_BEAR})
    scores = lr.get("scores", SCORES)

    highs   = [f for f in findings_today if f.get("classification") == "HIGH"]
    mediums = [f for f in findings_today if f.get("classification") == "MEDIUM"]
    lows    = [f for f in findings_today if f.get("classification") == "LOW"]

    total_items = len(findings_today)
    subject = build_email_subject(len(highs), len(mediums), len(lows))

    proj = build_8q_projection()
    five_y = build_5y_terminal()

    # ── helper to format one finding block ───────────────────────────────────
    def _finding_block(f):
        confirmed_tag = "" if f.get("confirmed", True) else " <b style='color:#c0392b'>[UNCONFIRMED]</b>"
        source = f.get("source", "")
        if source.startswith("http"):
            source_html = f'<a href="{source}">{source}</a>'
        else:
            source_html = source

        return (
            f"<tr><td style='padding:10px 0; border-bottom:1px solid #eee;'>"
            f"<b>{f['title']}</b>{confirmed_tag}<br>"
            f"<small style='color:#666'>{source_html} · {f.get('date','')}</small><br>"
            f"<span style='color:#444'>{f['impact']}</span><br>"
            f"<small><i>Disposition: {f.get('disposition','logged only')}</i></small>"
            f"</td></tr>"
        )

    def _section(label, color, items):
        if not items:
            return f"<h3 style='color:{color}'>{label}</h3><p style='color:#888'>None today.</p>"
        rows_html = "".join(_finding_block(f) for f in items)
        return (
            f"<h3 style='color:{color}'>{label}</h3>"
            f"<table style='width:100%;border-collapse:collapse'>{rows_html}</table>"
        )

    # ── 8-quarter revenue table ───────────────────────────────────────────────
    q_rows = ""
    for row in proj:
        note = "← Guided $12.7–$13.3B" if row["quarter"] == "Q3 2026" else ""
        q_rows += (
            f"<tr>"
            f"<td style='padding:4px 8px'><b>{row['quarter']}</b></td>"
            f"<td style='padding:4px 8px;text-align:right;color:#27ae60'>${row['bull_rev']:.2f}B</td>"
            f"<td style='padding:4px 8px;text-align:right;color:#2980b9'>${row['base_rev']:.2f}B</td>"
            f"<td style='padding:4px 8px;text-align:right;color:#c0392b'>${row['bear_rev']:.2f}B</td>"
            f"<td style='padding:4px 8px;color:#888;font-size:12px'>{note}</td>"
            f"</tr>"
        )

    # ── 5-year price table ────────────────────────────────────────────────────
    fy_rows = ""
    for row in five_y:
        fy_rows += (
            f"<tr>"
            f"<td style='padding:4px 8px'><b>{row['year']}</b></td>"
            f"<td style='padding:4px 8px;text-align:right;color:#27ae60'>${row['bull_spl']:,.0f}–${row['bull_sph']:,.0f}</td>"
            f"<td style='padding:4px 8px;text-align:right;color:#2980b9'>${row['base_spl']:,.0f}–${row['base_sph']:,.0f}</td>"
            f"<td style='padding:4px 8px;text-align:right;color:#c0392b'>${row['bear_spl']:,.0f}–${row['bear_sph']:,.0f}</td>"
            f"<td style='padding:4px 8px;text-align:right;color:#8e44ad'><b>${row['ev_spl']:,.0f}</b></td>"
            f"</tr>"
        )

    drive_log_link  = f'<a href="{drive_log_url}">AMD_Monitor_Log.xlsx</a>'  if drive_log_url  else "AMD_Monitor_Log.xlsx"
    drive_doc_link  = f'<a href="{drive_doc_url}">Drive summary doc</a>'     if drive_doc_url  else "Drive summary doc"

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;color:#222;line-height:1.5">

<div style="background:#1a1a2e;color:#fff;padding:20px 24px;border-radius:6px 6px 0 0">
  <h2 style="margin:0;font-size:20px">AMD MONITOR — {today.upper()}</h2>
  <p style="margin:4px 0 0;opacity:0.8;font-size:14px">
    {len(highs)} HIGH &nbsp;·&nbsp; {len(mediums)} MEDIUM &nbsp;·&nbsp; {len(lows)} LOW
    &nbsp;·&nbsp; {total_items} items reviewed today
  </p>
</div>

<div style="background:#f4f6f9;padding:16px 24px;border:1px solid #dce1e9;font-size:14px">
  <b>Model Status:</b>
  &nbsp; Bull <b style="color:#27ae60">{probs['bull']*100:.0f}%</b>
  &nbsp;·&nbsp; Base <b style="color:#2980b9">{probs['base']*100:.0f}%</b>
  &nbsp;·&nbsp; Bear <b style="color:#c0392b">{probs['bear']*100:.0f}%</b>
  &nbsp;&nbsp;|&nbsp;&nbsp;
  Q2 Rev <b>${Q2_REV:.2f}B</b> (+50% Y/Y)
  &nbsp;·&nbsp; DC <b>${Q2_DC_REV:.1f}B</b> (+{Q2_DC_GROWTH*100:.0f}%)
  &nbsp;·&nbsp; Entry <b>${ENTRY_PRICE:,.2f}</b>
</div>

<div style="padding:0 24px">

{_section('🔴 HIGH IMPACT', '#c0392b', highs)}
{_section('🟡 MEDIUM IMPACT', '#e67e22', mediums)}
{_section('⚪ LOW IMPACT / LOGGED', '#666', lows)}

{"<p style='color:#888;font-style:italic'>No items to report today — all sources checked, nothing new found.</p>" if total_items == 0 else ""}

<hr style="border:none;border-top:2px solid #1a1a2e;margin:32px 0 24px">

<h2 style="color:#1a1a2e">8-QUARTER REVENUE PROJECTION</h2>
<p style="font-size:13px;color:#666">
  All figures in $B. Starting from Q2 2026 actuals (${Q2_REV:.2f}B).
  Probabilities: Bull {probs['bull']*100:.0f}% / Base {probs['base']*100:.0f}% / Bear {probs['bear']*100:.0f}%.
</p>

<table style="width:100%;border-collapse:collapse;font-size:14px">
  <thead>
    <tr style="background:#1a1a2e;color:#fff">
      <th style="padding:8px;text-align:left">Quarter</th>
      <th style="padding:8px;text-align:right;color:#7ecf9a">Bull</th>
      <th style="padding:8px;text-align:right;color:#85b9e0">Base</th>
      <th style="padding:8px;text-align:right;color:#e88080">Bear</th>
      <th style="padding:8px;text-align:left"></th>
    </tr>
  </thead>
  <tbody>
    {q_rows}
  </tbody>
</table>

<hr style="border:none;border-top:1px solid #dce1e9;margin:24px 0">

<h2 style="color:#1a1a2e">5-YEAR SHARE PRICE TARGETS</h2>
<p style="font-size:13px;color:#666">
  EPS × P/E ranges. EV = probability-weighted at low P/E.
  Entry price ${ENTRY_PRICE:,.2f}. Diluted shares {SHARES:.3f}B.
</p>

<table style="width:100%;border-collapse:collapse;font-size:14px">
  <thead>
    <tr style="background:#1a1a2e;color:#fff">
      <th style="padding:8px;text-align:left">Year</th>
      <th style="padding:8px;text-align:right;color:#7ecf9a">Bull Range</th>
      <th style="padding:8px;text-align:right;color:#85b9e0">Base Range</th>
      <th style="padding:8px;text-align:right;color:#e88080">Bear Range</th>
      <th style="padding:8px;text-align:right;color:#c39bd3">EV (low P/E)</th>
    </tr>
  </thead>
  <tbody>
    {fy_rows}
  </tbody>
</table>

<p style="font-size:12px;color:#888;margin-top:8px">
  SPL = EPS × PE Low &nbsp;·&nbsp; SPH = EPS × PE High &nbsp;·&nbsp;
  EV SPL = Bull×{probs['bull']*100:.0f}% + Base×{probs['base']*100:.0f}% + Bear×{probs['bear']*100:.0f}% at low P/E
</p>

<hr style="border:none;border-top:1px solid #dce1e9;margin:24px 0">

<h3 style="color:#1a1a2e">Confirmed Deals Pipeline (9 deals, 22+ GW)</h3>
<table style="width:100%;border-collapse:collapse;font-size:13px">
  <tr style="background:#f4f6f9"><th style="padding:6px 8px;text-align:left">Partner</th><th style="padding:6px 8px;text-align:left">Scale</th><th style="padding:6px 8px;text-align:left">Status</th></tr>
{"".join(f"<tr><td style='padding:4px 8px'>{n}</td><td style='padding:4px 8px'>{s}</td><td style='padding:4px 8px;color:#444'>{st}</td></tr>" for n, s, st in CONFIRMED_DEALS)}
</table>

{"".join(f"<div style='background:#fff8e1;border-left:4px solid #f39c12;padding:12px 16px;margin:16px 0'><b>⚠ Watch: {name}</b><br><span style='color:#444'>{note}</span></div>" for name, note in WATCH_ITEMS)}

</div>

<div style="background:#f4f6f9;padding:16px 24px;border:1px solid #dce1e9;border-radius:0 0 6px 6px;font-size:13px;color:#666;margin-top:16px">
  <b>Resources:</b>
  {drive_log_link} &nbsp;·&nbsp; {drive_doc_link}
  <br>
  <b>Sources checked today:</b>
  {", ".join(l for l, _ in DAILY_SOURCES)}
</div>

</body>
</html>"""

    return subject, html


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL — AMD_Monitor_Log.xlsx
# ══════════════════════════════════════════════════════════════════════════════

def _cell_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _hdr_style(ws, row, headers, widths, fill_hex="1a1a2e"):
    fill = _cell_fill(fill_hex)
    font = Font(bold=True, color="FFFFFF", size=10)
    for col_idx, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=row, column=col_idx, value=h)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def generate_log_xlsx():
    """Generate AMD_Monitor_Log.xlsx from monitor_log.json."""
    log = load_monitor_log()
    output_path = os.path.join(HERE, "AMD_Monitor_Log.xlsx")

    wb = openpyxl.Workbook()

    # ── Tab 1: Full Log ───────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Monitor Log"
    ws.sheet_view.showGridLines = True

    headers = ["Date", "Source", "Title", "Classification", "Disposition", "Impact Notes", "Files Affected", "Confirmed"]
    widths  = [12,     30,       50,      16,               22,            50,              25,               12]
    _hdr_style(ws, 1, headers, widths)
    ws.row_dimensions[1].height = 20

    class_colors = {
        "HIGH":   "FDECEA",
        "MEDIUM": "FFF8E1",
        "LOW":    "F1F8E9",
    }

    # Sort descending by date
    sorted_log = sorted(log, key=lambda x: x.get("date", ""), reverse=True)

    for r_idx, entry in enumerate(sorted_log, start=2):
        cls = entry.get("classification", "LOW")
        fill = _cell_fill(class_colors.get(cls, "FFFFFF"))

        values = [
            entry.get("date", ""),
            entry.get("source", ""),
            entry.get("title", ""),
            cls,
            entry.get("disposition", ""),
            entry.get("impact", ""),
            ", ".join(entry.get("files_affected", [])),
            "Yes" if entry.get("confirmed", True) else "No — Rumor/Unconfirmed",
        ]
        for c_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cls == "HIGH" and c_idx == 4:
                cell.font = Font(bold=True, color="C0392B")
            elif cls == "MEDIUM" and c_idx == 4:
                cell.font = Font(bold=True, color="E67E22")

        ws.row_dimensions[r_idx].height = 45

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H1"

    # ── Tab 2: Summary by Classification ─────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 55

    ws2["A1"] = "AMD Monitor Log — Summary"
    ws2["A1"].font = Font(bold=True, size=14, color="1A1A2E")

    from collections import Counter
    cls_counts = Counter(e.get("classification", "LOW") for e in log)

    ws2["A3"] = "Classification"
    ws2["B3"] = "Count"
    ws2["C3"] = "Latest"
    for cell in [ws2["A3"], ws2["B3"], ws2["C3"]]:
        cell.fill = _cell_fill("1A1A2E")
        cell.font = Font(bold=True, color="FFFFFF")

    row = 4
    for cls in ["HIGH", "MEDIUM", "LOW"]:
        latest = next((e.get("title","") for e in sorted_log if e.get("classification") == cls), "—")
        ws2.cell(row=row, column=1, value=cls)
        ws2.cell(row=row, column=2, value=cls_counts.get(cls, 0))
        ws2.cell(row=row, column=3, value=latest)
        row += 1

    ws2.cell(row=row+1, column=1, value=f"Total entries: {len(log)}")
    ws2.cell(row=row+2, column=1, value=f"Generated: {datetime.date.today().isoformat()}")

    wb.save(output_path)
    print(f"Generated: {output_path}  ({os.path.getsize(output_path)//1024} KB)")
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# SELF-TEST
# ══════════════════════════════════════════════════════════════════════════════

def self_test():
    failures = []

    # 1. State management round-trip
    _ensure_data_dir()
    test_lr = _default_last_run()
    test_lr["last_run_date"] = "2026-09-17"
    save_last_run(test_lr)
    loaded = load_last_run()
    if loaded.get("last_quarter") != AS_OF:
        failures.append(f"last_run.json round-trip failed: expected {AS_OF}, got {loaded.get('last_quarter')}")

    # 2. 8-quarter projection has 8 rows
    proj = build_8q_projection()
    if len(proj) != 8:
        failures.append(f"8Q projection has {len(proj)} rows, expected 8")
    if proj[0]["quarter"] != "Q3 2026":
        failures.append(f"First quarter is {proj[0]['quarter']}, expected Q3 2026")

    # 3. Q3 2026 anchored to guide
    q3 = proj[0]
    if not (12.0 <= q3["base_rev"] <= 14.0):
        failures.append(f"Q3 base revenue {q3['base_rev']} out of expected range 12.0–14.0")

    # 4. 5-year terminal: 5 rows
    five_y = build_5y_terminal()
    if len(five_y) != 5:
        failures.append(f"5Y terminal has {len(five_y)} rows, expected 5")

    # 5. Bull 2030 SPL in reasonable range
    bull_2030_spl = five_y[-1]["bull_spl"]
    if not (2500 <= bull_2030_spl <= 3500):
        failures.append(f"Bull 2030 SPL {bull_2030_spl} out of range 2500–3500")

    # 6. Report text contains key strings
    report = build_report_text()
    for s in ["Q3 2026", "$11.54B", "Bull", "CONFIRMED DEALS", "8-QUARTER"]:
        if s not in report:
            failures.append(f"Report text missing: '{s}'")

    # 7. Email body builds without error
    subj, html = build_email_body([], "", "")
    if "AMD MONITOR" not in html:
        failures.append("Email HTML missing 'AMD MONITOR'")
    if "8-QUARTER REVENUE PROJECTION" not in html:
        failures.append("Email HTML missing 8-quarter projection section")

    # 8. Add-finding round-trip
    test_finding = json.dumps({
        "date": "2026-09-17",
        "source": "https://ir.amd.com/test",
        "title": "Self-test finding",
        "classification": "LOW",
        "disposition": "logged only",
        "impact": "Self-test. No action.",
        "confirmed": True,
        "files_affected": [],
    })
    add_finding(test_finding)
    log = load_monitor_log()
    if not any(e.get("title") == "Self-test finding" for e in log):
        failures.append("add_finding round-trip failed")

    # 9. Excel generation
    xlsx_path = generate_log_xlsx()
    if not os.path.exists(xlsx_path):
        failures.append("generate_log_xlsx did not create file")
    if os.path.getsize(xlsx_path) < 2000:
        failures.append(f"AMD_Monitor_Log.xlsx too small: {os.path.getsize(xlsx_path)} bytes")

    # 10. detect_quarter returns expected structure
    dq = detect_quarter()
    for k in ["last_quarter", "last_filing_date", "action", "instruction"]:
        if k not in dq:
            failures.append(f"detect_quarter missing key: {k}")

    # ── result ────────────────────────────────────────────────────────────────
    if failures:
        print("\n".join(f"  FAIL: {f}" for f in failures))
        print("\nSELF-TEST FAILED")
        sys.exit(1)
    else:
        print("SELF-TEST PASSED — all 10 checks passed")


# ══════════════════════════════════════════════════════════════════════════════
# MONITORING SCHEDULE (documentation — Claude reads this as its run-book)
# ══════════════════════════════════════════════════════════════════════════════

RUNBOOK = """
AMD MONITORING AGENT — CLAUDE RUN-BOOK
=======================================

This run-book tells Claude what to check at each cadence.
Claude uses the WebSearch tool for all web lookups.
Email is sent via Python SMTP: python amd_agent.py send-email
  Requires GMAIL_APP_PASSWORD env var (myaccount.google.com/apppasswords).
Claude uses mcp__Google_Drive__ for Drive uploads (when MCP available).

DAILY CHECKS (every run):
  1. AMD IR press releases: search "AMD press release site:ir.amd.com"
  2. AMD newsroom: search "AMD newsroom announcement site:amd.com"
  3. SEC 8-K filings: search "AMD 8-K SEC EDGAR filing 2026"
  4. Check event_watch.json for any events that are "today" or "yesterday"
     — For events marked "pending" with date == today: check AMD IR and newsroom
     — For events marked "checked_day_of" with date == yesterday: check for
       transcripts and follow-up coverage (conference transcripts often post 24h late)

WEEKLY CHECKS (Monday or first run of the week):
  5. AMD IR events calendar: search "AMD investor conference event 2026 site:ir.amd.com"
     — Log any NEW events to event_watch.json with status "pending"
  6. Nvidia announcements: search "Nvidia announcement news September 2026"
  7. Intel announcements: search "Intel data center news September 2026"
  8. Partner confirmations: search "OpenAI AMD OR Meta AMD OR Anthropic AMD deployment 2026"

QUARTERLY CHECKS (when new earnings are detected):
  9. TSMC: search "TSMC manufacturing announcement 2026"
  10. Mercury Research: search "Mercury Research server CPU market share 2026"
  11. Hyperscaler capex: search "Google Microsoft AWS Meta capex AI spending 2026"

CLASSIFICATION RULES:
  HIGH:   New GW-scale deal, earnings release, guidance change, acquisition,
          AMD TAM change (like CFO $3T statement), major competitive shift.
          → Triggers immediate email with "ACTION REQUIRED" flag.
  MEDIUM: Product launch, partnership, ROCm milestone, conference statement with
          new data, competitor weakness, rumored deal >100MW.
          → Updates thesis scores, included in daily email.
  LOW:    Exec hire, governance, conference attendance with no new data,
          rumored deals <100MW, general market commentary.
          → Logged only, included in daily email.

CONFERENCE EVENT TRACKING (special rule):
  AMD or competitor conference detected → log to event_watch.json:
    {event: "Name", date: "YYYY-MM-DD", status: "pending", source: "ir.amd.com/events"}
  On event date: check AMD IR and newsroom → update status "checked_day_of"
  Day after: check for transcript/follow-up → update status "checked_day_after"
  Example: Citi TMT Sep 8 2026 — CFO raised TAM to $3T, appeared only in
           day-after transcript coverage, never as standalone press release.

AFTER EACH RUN:
  1. Run: python amd_agent.py add-finding '<json>'  (for each finding)
  2. Run: python amd_agent.py generate-log-xlsx     (creates AMD_Monitor_Log.xlsx)
  3. Upload AMD_Monitor_Log.xlsx to Drive Q2 2026 folder via mcp__Google_Drive__update_file (if MCP available)
  4. Run: python amd_agent.py send-email            (builds email from today's findings + sends via SMTP)
     — Requires GMAIL_APP_PASSWORD env var. Generate at myaccount.google.com/apppasswords.
  5. Update Drive summary doc via mcp__Google_Drive__update_file (if MCP available)

GRACEFUL FALLBACK:
  If WebSearch returns no results or errors:
  - Log the failure: {"classification":"LOW","title":"Search failed: <source>","impact":"STALE DATA — check manually"}
  - Continue with remaining sources
  - Never skip email or Excel generation

DRIVE FOLDER MAPPING:
  AMD Financials master: 1i1GdOQGreQuxv2s6xoqILyY_XIqKuX5a
  Q2 2026:               1D_GOhT8kgZBTlVPA_7VXvO0-lR5ZyTbB
  Q1 2026:               1Df4g1sgPkwy_7BwBLjNTsYNDPh28RiWc
  New quarter: Claude creates a subfolder under master using mcp__Google_Drive__create_file
               with mimeType="application/vnd.google-apps.folder"

QUARTERLY REBUILD TRIGGER:
  1. Run: python amd_agent.py detect-quarter
  2. If action == "check_for_new_quarter": search SEC EDGAR for new AMD 10-Q
  3. If new filing found: flag as HIGH, recommend full rebuild in email
  4. Full rebuild = Claude creates new version scripts (v5, v9, etc.) with new DATA BLOCK
"""


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    cmd = sys.argv[1].lower()

    if cmd == "report":
        print(build_report_text())

    elif cmd == "detect-quarter":
        result = detect_quarter()
        print(json.dumps(result, indent=2))

    elif cmd == "carry-scores":
        lr = load_last_run()
        print(json.dumps(lr.get("scores", SCORES), indent=2))

    elif cmd == "add-finding":
        if len(sys.argv) < 3:
            print("Usage: python amd_agent.py add-finding '<json>'", file=sys.stderr)
            sys.exit(1)
        add_finding(sys.argv[2])

    elif cmd == "generate-log-xlsx":
        path = generate_log_xlsx()
        print(f"Output: {path}")

    elif cmd == "email-body":
        # Read findings from stdin if available, otherwise use empty list
        findings = []
        subj, html = build_email_body(findings)
        print(f"SUBJECT: {subj}")
        print("─" * 60)
        print(html)

    elif cmd == "send-email":
        # Build email from today's findings in monitor_log.json, then send via SMTP
        today_str = datetime.date.today().isoformat()
        log = load_monitor_log()
        findings_today = [e for e in log if e.get("date", "").startswith(today_str)]
        subj, html = build_email_body(findings_today)
        send_email_smtp(subj, html)

    elif cmd == "update-quarter":
        if len(sys.argv) < 3:
            print("Usage: python amd_agent.py update-quarter <json_file>", file=sys.stderr)
            sys.exit(1)
        update_quarter(sys.argv[2])

    elif cmd == "runbook":
        print(RUNBOOK)

    elif cmd == "self-test":
        self_test()

    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
