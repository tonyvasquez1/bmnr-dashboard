const https = require('https');
const FINNHUB_KEY = 'd7foe0hr01qqb8rh1gr0d7foe0hr01qqb8rh1grg';

function get(url, timeout=10000) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, {
      headers: { 'User-Agent': 'BMNR-Dashboard/1.0', 'Accept': 'application/json' }
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode !== 200) { reject(new Error(`HTTP ${res.statusCode}`)); return; }
        try { resolve(JSON.parse(data)); } catch(e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.setTimeout(timeout, () => { req.destroy(); reject(new Error('Timeout')); });
  });
}

// ── 1. SHARES OUTSTANDING ─────────────────────────────────────────────
// Fast: Finnhub basic financials only — no SEC document parsing
const ANCHOR_SHARES = 537628819;
const ANCHOR_DATE   = '2026-04-13';

async function fetchShares() {
  let finnhubShares = null, finnhubSource = null;

  // Finnhub: shares outstanding from latest filing
  try {
    const d = await get(`https://finnhub.io/api/v1/stock/metric?symbol=BMNR&metric=all&token=${FINNHUB_KEY}`, 8000);
    const s = d?.metric?.sharesOutstanding; // in millions
    if (s && s > 400) { finnhubShares = Math.round(s * 1e6); finnhubSource = 'Finnhub'; }
  } catch(e) {}

  // SEC EDGAR: count 424B5 filings since anchor (index only — fast)
  let atmCount = 0, latestAtmDate = null;
  try {
    const d = await get('https://data.sec.gov/submissions/CIK0001829311.json', 8000);
    const forms = d?.filings?.recent?.form || [];
    const dates = d?.filings?.recent?.filingDate || [];
    for (let i = 0; i < forms.length; i++) {
      if (dates[i] < ANCHOR_DATE) break;
      if (forms[i] === '424B5') {
        atmCount++;
        if (!latestAtmDate) latestAtmDate = dates[i];
      }
    }
  } catch(e) {}

  // Best estimate: Finnhub if available, else anchor + conservative ATM estimate
  const ATM_PER_FILING_EST = 12e6; // ~12M shares per 424B5 filing (conservative)
  const atmEstimate = atmCount * ATM_PER_FILING_EST;
  const best = finnhubShares || (ANCHOR_SHARES + atmEstimate);

  return {
    best,
    finnhub: finnhubShares,
    anchor: ANCHOR_SHARES,
    anchorDate: ANCHOR_DATE,
    atmFilings: atmCount,
    latestAtmDate,
    atmEstimate: Math.round(atmEstimate),
    weeklyEst: ATM_PER_FILING_EST,
    source: finnhubSource || `SEC anchor + ${atmCount} ATM filings est.`
  };
}

// ── 2. CLARITY ACT POLYMARKET ─────────────────────────────────────────
async function fetchClarity() {
  // Polymarket gamma API
  try {
    const d = await get('https://gamma-api.polymarket.com/markets?slug=clarity-act-signed-into-law-in-2026', 8000);
    if (d?.[0]) {
      const prices = JSON.parse(d[0].outcomePrices || '[]');
      const yes = parseFloat(prices[0]);
      if (yes > 0) return { odds: Math.round(yes * 100), raw: yes, source: 'Polymarket' };
    }
  } catch(e) {}

  // Fallback: Polymarket CLOB
  try {
    const slug = 'will-the-digital-asset-market-clarity-act-be-signed-into-law-in-2026';
    const d = await get(`https://gamma-api.polymarket.com/markets?slug=${slug}`, 8000);
    if (d?.[0]) {
      const prices = JSON.parse(d[0].outcomePrices || '[]');
      const yes = parseFloat(prices[0]);
      if (yes > 0) return { odds: Math.round(yes * 100), raw: yes, source: 'Polymarket' };
    }
  } catch(e) {}

  return { odds: 62, source: 'cached (May 3)', cached: true };
}

// ── 3. ETH SUPPLY ─────────────────────────────────────────────────────
async function fetchEthSupply() {
  try {
    const d = await get('https://api.coingecko.com/api/v3/coins/ethereum?localization=false&tickers=false&market_data=true&community_data=false&developer_data=false', 8000);
    const supply = d?.market_data?.circulating_supply;
    if (supply > 100e6) return { supply: Math.round(supply), source: 'CoinGecko' };
  } catch(e) {}
  return { supply: 120700000, source: 'fallback' };
}

// ── 4. RWA ON-CHAIN (DeFiLlama) ───────────────────────────────────────
// DeFiLlama tracks RWA category TVL — free, no auth, reliable
async function fetchRWA() {
  let exclStablecoins = null, inclStablecoins = null, ethShare = null;

  // Try DeFiLlama categories for RWA
  try {
    const d = await get('https://api.llama.fi/overview/protocols?category=RWA', 6000);
    // Sum all RWA protocols TVL
    if (Array.isArray(d)) {
      const total = d.reduce((sum, p) => sum + (p.tvl || 0), 0);
      if (total > 1e9) exclStablecoins = total;
    }
  } catch(e) {}

  // Try DeFiLlama global stablecoin data
  try {
    const d = await get('https://stablecoins.llama.fi/stablecoins?includePrices=true', 6000);
    if (d?.peggedAssets) {
      const total = d.peggedAssets.reduce((sum, s) => sum + (s.circulating?.peggedUSD || 0), 0);
      if (total > 100e9) inclStablecoins = total;
    }
  } catch(e) {}

  // Try DeFiLlama chains for ETH RWA share
  try {
    const d = await get('https://api.llama.fi/v2/chains', 6000);
    const eth = d?.find(c => c.name === 'Ethereum');
    if (eth?.tvl && exclStablecoins) {
      // ETH share is approximate based on known ~62% split
    }
  } catch(e) {}

  // Fallback: use known recent values from rwa.xyz
  const fallbackExcl = 30.91e9;  // $30.91B as of May 3
  const fallbackIncl = 330e9;    // ~$330B incl stablecoins

  return {
    exclStablecoins: exclStablecoins || fallbackExcl,
    inclStablecoins: inclStablecoins || fallbackIncl,
    ethShare: 0.62,             // ETH ~62% of RWA (rwa.xyz data)
    ethRWA: (exclStablecoins || fallbackExcl) * 0.62,
    change30d: '+7.5%',
    source: exclStablecoins ? 'DeFiLlama' : 'rwa.xyz cached (May 3)',
    cached: !exclStablecoins,
    baseTarget: 40e9,           // $40-50B = Base scenario confirmed
    bullTarget: 80e9,           // $80B+ = Bull scenario
    baseProgress: ((exclStablecoins || fallbackExcl) / 45e9 * 100).toFixed(0),
    gateProgress: ((exclStablecoins || fallbackExcl) / 100e9 * 100).toFixed(0),
    blackrockBuidl: 1.9e9
  };
}

// ── MAIN ──────────────────────────────────────────────────────────────
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0');
  res.setHeader('Content-Type', 'application/json');

  const [sharesR, clarityR, ethSupplyR, rwaR] = await Promise.allSettled([
    fetchShares(),
    fetchClarity(),
    fetchEthSupply(),
    fetchRWA()
  ]);

  res.status(200).json({
    timestamp: new Date().toISOString(),
    shares:    sharesR.status    === 'fulfilled' ? sharesR.value    : { error: sharesR.reason?.message,    anchor: 537628819 },
    clarity:   clarityR.status   === 'fulfilled' ? clarityR.value   : { error: clarityR.reason?.message,   odds: 62 },
    ethSupply: ethSupplyR.status === 'fulfilled' ? ethSupplyR.value : { error: ethSupplyR.reason?.message, supply: 120700000 },
    rwa:       rwaR.status       === 'fulfilled' ? rwaR.value       : { error: rwaR.reason?.message,       exclStablecoins: 30.91e9, inclStablecoins: 330e9 }
  });
};
