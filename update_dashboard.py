#!/usr/bin/env python3
"""
BMNR Dashboard Auto-Updater
Runs daily via GitHub Actions.
Checks PRNewswire for new press releases and SEC EDGAR for new 8-K filings.
Only updates dashboard if new data is detected.
"""

import urllib.request, urllib.error, json, base64, sys, re, os
from datetime import datetime, timezone

GITHUB_TOKEN = os.environ.get('GITHUB_TOKEN', '')
GITHUB_USER  = 'tonyvasquez1'
GITHUB_REPO  = 'bmnr-dashboard'
FINNHUB_KEY  = 'd7foe0hr01qqb8rh1gr0d7foe0hr01qqb8rh1grg'

# ── HELPERS ───────────────────────────────────────────────────────────
def ok(m):   print(f"  ✓ {m}")
def info(m): print(f"  · {m}")
def warn(m): print(f"  ⚠ {m}")
def fail(m): print(f"  ✗ {m}"); sys.exit(1)

def fetch_url(url, timeout=15):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'BMNR-Dashboard-Updater/1.0'})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read().decode('utf-8', errors='ignore')
    except Exception as e:
        warn(f"Failed to fetch {url[:60]}: {e}")
        return None

def gh(method, path, body=None):
    url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/{path}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
        "User-Agent": "BMNR-Deploy"
    }
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, headers=headers, data=data, method=method)
    try:
        with urllib.request.urlopen(req) as r:
            return r.status, json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try: msg = json.loads(e.read().decode()).get("message","")
        except: msg = str(e)
        return e.code, msg

# ── STEP 1: Check SEC EDGAR for latest BMNR 8-K ─────────────────────
def get_latest_sec_filing():
    info("Checking SEC EDGAR for latest BMNR 8-K...")
    url = "https://data.sec.gov/submissions/CIK0001829311.json"
    raw = fetch_url(url)
    if not raw:
        return None
    try:
        data = json.loads(raw)
        filings = data.get('filings', {}).get('recent', {})
        forms = filings.get('form', [])
        dates = filings.get('filingDate', [])
        accNums = filings.get('accessionNumber', [])
        for i, form in enumerate(forms):
            if form == '8-K':
                acc = accNums[i].replace('-','')
                filing_url = f"https://www.sec.gov/Archives/edgar/data/1829311/{acc}/"
                ok(f"Latest 8-K: {dates[i]} — {filing_url}")
                return {'date': dates[i], 'url': filing_url, 'acc': accNums[i]}
    except Exception as e:
        warn(f"SEC parse error: {e}")
    return None

# ── STEP 2: Check PRNewswire for latest BMNR press release ───────────
def get_latest_prnewswire():
    info("Checking PRNewswire for latest BMNR press release...")
    # Search SEC EDGAR full-text for most recent 8-K exhibit (press release)
    url = "https://efts.sec.gov/LATEST/search-index?q=%22Bitmine%22&dateRange=custom&startdt=2026-01-01&forms=8-K&hits.hits._source.period_of_report=&hits.hits.total.value=1"
    # Use the EDGAR full text search instead
    search_url = "https://efts.sec.gov/LATEST/search-index?q=%224%2C976%2C485%22+%22Bitmine%22&forms=8-K"
    
    # Primary: fetch PRNewswire RSS
    rss_url = "https://www.prnewswire.com/rss/news-releases-list.rss?category=BL&keywords=Bitmine+BMNR"
    raw = fetch_url(rss_url)
    if raw:
        # Find most recent BMNR release
        matches = re.findall(r'<title>(.*?Bitmine.*?)</title>', raw, re.IGNORECASE)
        if matches:
            ok(f"PRNewswire latest: {matches[0][:80]}")
            return matches[0]
    
    # Fallback: check SEC EDGAR exhibits directly
    info("PRNewswire RSS unavailable — using SEC EDGAR exhibits")
    return None

# ── STEP 3: Fetch and parse the actual press release content ─────────
def parse_press_release(filing_info):
    if not filing_info:
        return None
    
    info(f"Fetching filing index: {filing_info['url']}")
    index_raw = fetch_url(filing_info['url'])
    if not index_raw:
        return None
    
    # Find exhibit 99.1 (press release)
    ex_match = re.search(r'ex99[-_]?1[^"]*\.htm', index_raw, re.IGNORECASE)
    if not ex_match:
        warn("Could not find exhibit 99.1 in filing")
        return None
    
    ex_url = f"https://www.sec.gov/Archives/edgar/data/1829311/{filing_info['acc'].replace('-','')}/{ex_match.group(0)}"
    info(f"Fetching press release: {ex_url}")
    pr_raw = fetch_url(ex_url)
    if not pr_raw:
        return None
    
    # Parse key figures
    data = {}
    
    # ETH held
    m = re.search(r'([\d,]+)\s*ETH at \$([\d,]+(?:\.\d+)?)\s*per ETH', pr_raw)
    if m:
        data['eth_held'] = int(m.group(1).replace(',',''))
        ok(f"ETH held: {data['eth_held']:,}")
    
    # BTC held
    m = re.search(r'(\d+)\s*Bitcoin \(BTC\)', pr_raw)
    if m:
        data['btc_held'] = int(m.group(1))
        ok(f"BTC held: {data['btc_held']}")
    
    # Cash
    m = re.search(r'total cash of \$([\d.]+)\s*(billion|million)', pr_raw, re.IGNORECASE)
    if m:
        val = float(m.group(1))
        mult = 1e9 if 'billion' in m.group(2).lower() else 1e6
        data['cash'] = val * mult
        ok(f"Cash: ${val}{m.group(2)[0].upper()}")
    
    # ORBS
    m = re.search(r'\$([\d]+)\s*million stake in Eightco', pr_raw, re.IGNORECASE)
    if m:
        data['orbs'] = float(m.group(1)) * 1e6
        ok(f"ORBS: ${m.group(1)}M")
    
    # ETH staked
    m = re.search(r'([\d,]+)\s*staked ETH', pr_raw)
    if m:
        data['eth_staked'] = int(m.group(1).replace(',',''))
        ok(f"ETH staked: {data['eth_staked']:,}")
    
    # Staking revenue
    m = re.search(r'\$([\d]+)\s*million.*?annualized|Annualized.*?\$([\d]+)\s*million', pr_raw, re.IGNORECASE)
    if m:
        val = m.group(1) or m.group(2)
        data['staking_rev'] = float(val) * 1e6
        ok(f"Staking revenue: ${val}M/yr")
    
    # Yield %
    m = re.search(r'([\d.]+)%\s*7-day BMNR yield', pr_raw)
    if m:
        data['staking_yield'] = float(m.group(1))
        ok(f"Staking yield: {data['staking_yield']}%")
    
    # CESR
    m = re.search(r'CESR.*?is\s*([\d.]+)%', pr_raw)
    if m:
        data['cesr'] = float(m.group(1))
        ok(f"CESR: {data['cesr']}%")
    
    # Projected full staking
    m = re.search(r'\$([\d]+)\s*million annually.*?2\.\d+%|projected.*?\$([\d]+)\s*million annually', pr_raw, re.IGNORECASE)
    if m:
        val = m.group(1) or m.group(2)
        data['staking_proj'] = float(val) * 1e6
        ok(f"Staking projected: ${val}M/yr")
    
    data['filing_date'] = filing_info['date']
    return data

# ── STEP 4: Read current dashboard and check if update needed ────────
def get_current_dashboard():
    status, data = gh("GET", "contents/index.html")
    if status != 200:
        fail(f"Could not fetch current dashboard: {status}")
    content = base64.b64decode(data['content'].replace('\n','')).decode('utf-8')
    sha = data['sha']
    
    # Extract current ETH_HELD value
    m = re.search(r'const ETH_HELD=(\d+)', content)
    current_eth = int(m.group(1)) if m else 0
    info(f"Current dashboard ETH_HELD: {current_eth:,}")
    return content, sha, current_eth

# ── STEP 5: Update dashboard constants ───────────────────────────────
def update_dashboard(content, new_data):
    if not new_data:
        return content, False
    
    changed = False
    
    if 'eth_held' in new_data:
        old = re.search(r'const ETH_HELD=(\d+)', content)
        if old and int(old.group(1)) != new_data['eth_held']:
            content = re.sub(r'const ETH_HELD=\d+', f"const ETH_HELD={new_data['eth_held']}", content)
            changed = True
            ok(f"Updated ETH_HELD: {new_data['eth_held']:,}")
    
    if 'btc_held' in new_data:
        content = re.sub(r'BTC_HELD=\d+', f"BTC_HELD={new_data['btc_held']}", content)
        changed = True
    
    if 'cash' in new_data:
        cash_val = int(new_data['cash'])
        content = re.sub(r'CASH=[\d.e+]+', f"CASH={cash_val}", content)
        changed = True
        ok(f"Updated CASH: ${cash_val/1e9:.2f}B")
    
    if 'orbs' in new_data:
        orbs_val = int(new_data['orbs'])
        content = re.sub(r'ORBS=[\d.e+]+', f"ORBS={orbs_val}", content)
        changed = True
    
    if 'eth_staked' in new_data:
        content = re.sub(r'ETH_STAKED=\d+', f"ETH_STAKED={new_data['eth_staked']}", content)
        changed = True
        ok(f"Updated ETH_STAKED: {new_data['eth_staked']:,}")
    
    if 'staking_rev' in new_data:
        rev = int(new_data['staking_rev'])
        content = re.sub(r'STAKING_REV=[\d.e+]+', f"STAKING_REV={rev}", content)
        changed = True
        ok(f"Updated STAKING_REV: ${rev/1e6:.0f}M")
    
    if 'staking_yield' in new_data:
        content = re.sub(r'STAKING_YIELD=[\d.]+', f"STAKING_YIELD={new_data['staking_yield']}", content)
        changed = True
    
    if 'cesr' in new_data:
        content = re.sub(r'CESR=[\d.]+', f"CESR={new_data['cesr']}", content)
        changed = True
    
    if 'staking_proj' in new_data:
        proj = int(new_data['staking_proj'])
        content = re.sub(r'STAKING_PROJ=[\d.e+]+', f"STAKING_PROJ={proj}", content)
        changed = True
        ok(f"Updated STAKING_PROJ: ${proj/1e6:.0f}M")
    
    # Update footer date
    if 'filing_date' in new_data:
        content = re.sub(
            r'Balance sheet: 8-K \w+ \d+, \d+',
            f"Balance sheet: 8-K {new_data['filing_date']}",
            content
        )
    
    return content, changed

# ── STEP 6: Push updated dashboard ───────────────────────────────────
def push_dashboard(content, sha, filing_date):
    b64 = base64.b64encode(content.encode()).decode()
    body = {
        "message": f"Auto-update: BMNR 8-K {filing_date}",
        "content": b64,
        "sha": sha,
        "branch": "main"
    }
    status, data = gh("PUT", "contents/index.html", body)
    if status in (200, 201):
        ok("Dashboard pushed to GitHub Pages")
    else:
        fail(f"Push failed {status}: {data}")

# ── MAIN ─────────────────────────────────────────────────────────────
def main():
    print()
    print("=" * 56)
    print("  BMNR Dashboard Auto-Updater")
    print(f"  {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
    print("=" * 56)
    print()

    # Get current dashboard state
    current_content, current_sha, current_eth = get_current_dashboard()

    # Check SEC EDGAR for latest 8-K
    filing = get_latest_sec_filing()
    if not filing:
        warn("Could not get SEC filing info — skipping update")
        return

    # Check PRNewswire for press release
    get_latest_prnewswire()

    # Parse the press release
    print()
    info("Parsing press release figures...")
    new_data = parse_press_release(filing)

    if not new_data:
        warn("Could not parse press release — skipping update")
        return

    # Check if ETH held changed (main signal that new data exists)
    new_eth = new_data.get('eth_held', current_eth)
    if new_eth == current_eth:
        ok(f"No new data detected (ETH_HELD still {current_eth:,}) — dashboard up to date")
        return

    ok(f"New data detected: ETH_HELD {current_eth:,} → {new_eth:,}")
    print()

    # Update dashboard
    updated_content, changed = update_dashboard(current_content, new_data)
    if not changed:
        ok("No changes needed")
        return

    # Push
    print()
    info("Pushing updated dashboard...")
    push_dashboard(updated_content, current_sha, filing['date'])

    print()
    print("=" * 56)
    print("  Update complete!")
    print("  https://tonyvasquez1.github.io/bmnr-dashboard/")
    print("=" * 56)
    print()

if __name__ == "__main__":
    main()
