const https = require('https');

const FINNHUB_KEY = 'd7foe0hr01qqb8rh1gr0d7foe0hr01qqb8rh1grg';

function get(url, timeout = 10000) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, {
      headers: { 'User-Agent': 'BMNR-Dashboard/1.0', 'Accept': 'application/json' }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode !== 200) { reject(new Error(`HTTP ${res.statusCode}`)); return; }
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error(`JSON parse error`)); }
      });
    });
    req.on('error', reject);
    req.setTimeout(timeout, () => { req.destroy(); reject(new Error('Timeout')); });
  });
}

// ── 1. SHARES OUTSTANDING ─────────────────────────────────────────────
// Primary: Finnhub basic financials (shares outstanding)
// Fallback: SEC EDGAR 10-Q filing
// Supplement: Count 424B5 ATM issuances since 10-Q anchor
async function fetchShares() {
  let sharesFromFinancials = null;
  let source = null;

  // Try Finnhub shares outstanding
  try {
    const d = await get(`https://finnhub.io/api/v1/stock/metric?symbol=BMNR&metric=all&token=${FINNHUB_KEY}`);
    const shares = d?.metric?.sharesOutstanding;
    if (shares && shares > 400) { // in millions
      sharesFromFinancials = Math.round(shares * 1e6);
      source = 'Finnhub';
    }
  } catch (e) {}

  // Try SEC EDGAR submissions for latest 10-Q share count
  let secShares = null, secDate = null;
  try {
    const d = await get('https://data.sec.gov/submissions/CIK0001829311.json');
    const filings = d?.filings?.recent;
    const forms = filings?.form || [];
    const dates = filings?.filingDate || [];
    const docs  = filings?.primaryDocument || [];
    const accNums = filings?.accessionNumber || [];

    // Find most recent 10-Q
    for (let i = 0; i < forms.length; i++) {
      if (forms[i] === '10-Q') {
        secDate = dates[i];
        // Try to get shares from the cover page
        const acc = accNums[i].replace(/-/g, '');
        try {
          const idx = await get(`https://data.sec.gov/Archives/edgar/data/1829311/${acc}/index.json`);
          const files = idx?.directory?.item || [];
          const htm = files.find(f => f.name?.match(/\.htm$/i) && !f.name?.includes('R'));
          if (htm) {
            // Get the filing and search for shares outstanding
            const fUrl = `https://data.sec.gov/Archives/edgar/data/1829311/${acc}/${htm.name}`;
            const raw = await new Promise((res, rej) => {
              const r = https.get(fUrl, { headers: { 'User-Agent': 'BMNR-Dashboard/1.0' } }, resp => {
                let d = '';
                resp.on('data', c => d += c);
                resp.on('end', () => res(d));
              });
              r.on('error', rej);
              r.setTimeout(12000, () => { r.destroy(); rej(new Error('timeout')); });
            });
            // Parse shares from cover page
            const m = raw.match(/(\d{2,3},\d{3},\d{3})\s*(?:shares|common shares)\s*(?:of common stock\s*)?(?:issued and\s*)?outstanding/i);
            if (m) {
              secShares = parseInt(m[1].replace(/,/g, ''));
            }
          }
        } catch(e) {}
        break;
      }
    }
  } catch (e) {}

  // Count 424B5 ATM issuances since the 10-Q anchor date
  let atmShares = 0, atmFilings = [], atmError = null;
  const ANCHOR_DATE = '2026-04-13'; // Apr 13 10-Q anchor
  const ANCHOR_SHARES = 537628819;
  try {
    const d = await get('https://data.sec.gov/submissions/CIK0001829311.json');
    const filings = d?.filings?.recent;
    const forms = filings?.form || [];
    const dates = filings?.filingDate || [];
    const accNums = filings?.accessionNumber || [];

    for (let i = 0; i < forms.length; i++) {
      if ((forms[i] === '424B5' || forms[i] === 'S-3ASR') && dates[i] > ANCHOR_DATE) {
        atmFilings.push({ date: dates[i], acc: accNums[i] });
      }
      if (dates[i] < ANCHOR_DATE) break; // stop looking at older filings
    }

    // For each 424B5, try to extract shares issued
    for (const filing of atmFilings.slice(0, 10)) {
      try {
        const acc = filing.acc.replace(/-/g, '');
        const idx = await get(`https://data.sec.gov/Archives/edgar/data/1829311/${acc}/index.json`);
        const files = idx?.directory?.item || [];
        const htm = files.find(f => f.name?.match(/424b5.*\.htm$/i) || f.name?.match(/prospectus.*\.htm$/i));
        if (htm) {
          const fUrl = `https://data.sec.gov/Archives/edgar/data/1829311/${acc}/${htm.name}`;
          const raw = await new Promise((res, rej) => {
            const r = https.get(fUrl, { headers: { 'User-Agent': 'BMNR-Dashboard/1.0' } }, resp => {
              let d = '';
              resp.on('data', c => { d += c; if (d.length > 200000) { r.destroy(); res(d); } });
              resp.on('end', () => res(d));
            });
            r.on('error', rej);
            r.setTimeout(12000, () => { r.destroy(); rej(new Error('timeout')); });
          });
          // Look for shares being offered/sold
          const m = raw.match(/(?:offering|sold|selling)\s+(?:up to\s+)?(\d{1,3},\d{3},\d{3})\s+shares/i);
          if (m) {
            const n = parseInt(m[1].replace(/,/g, ''));
            if (n > 0 && n < 200000000) {
              atmShares += n;
              filing.sharesIssued = n;
            }
          }
        }
      } catch(e) {}
    }
  } catch (e) {
    atmError = e.message;
  }

  // Best estimate
  const bestShares = sharesFromFinancials || secShares || (ANCHOR_SHARES + atmShares);
  const estimatedCurrent = ANCHOR_SHARES + atmShares;

  return {
    finnhub: sharesFromFinancials,
    secAnchor: secShares || ANCHOR_SHARES,
    secAnchorDate: secDate || ANCHOR_DATE,
    atmFilingsSinceAnchor: atmFilings.length,
    atmSharesIssued: atmShares,
    estimatedCurrent,
    best: bestShares,
    source: sharesFromFinancials ? 'Finnhub' : (secShares ? 'SEC 10-Q' : 'SEC anchor + ATM'),
    weeklyRun: atmFilings.length > 0 ? Math.round(atmShares / Math.max(1, atmFilings.length)) : null,
    error: atmError
  };
}

// ── 2. CLARITY ACT POLYMARKET ODDS ───────────────────────────────────
async function fetchClarity() {
  try {
    // Polymarket public API — CLARITY Act market
    const MARKET_ID = '0x5f65177b394277fd294cd75650044e32ba009a0';
    const d = await get(`https://clob.polymarket.com/markets/${MARKET_ID}`);
    if (d?.tokens) {
      const yes = d.tokens.find(t => t.outcome === 'Yes');
      if (yes?.price) {
        return {
          odds: Math.round(yes.price * 100),
          raw: yes.price,
          source: 'Polymarket CLOB',
          marketId: MARKET_ID,
          updated: new Date().toISOString()
        };
      }
    }
  } catch (e) {}

  // Fallback: Polymarket gamma API
  try {
    const d = await get('https://gamma-api.polymarket.com/markets?slug=clarity-act-signed-into-law-in-2026');
    if (d?.[0]?.outcomePrices) {
      const prices = JSON.parse(d[0].outcomePrices);
      const yes = parseFloat(prices[0]);
      if (yes > 0) {
        return {
          odds: Math.round(yes * 100),
          raw: yes,
          source: 'Polymarket Gamma',
          updated: new Date().toISOString()
        };
      }
    }
  } catch (e) {}

  return { odds: null, source: 'unavailable', error: 'All Polymarket endpoints failed' };
}

// ── 3. ETH SUPPLY (real-time) ─────────────────────────────────────────
async function fetchEthSupply() {
  try {
    const d = await get('https://api.coingecko.com/api/v3/coins/ethereum?localization=false&tickers=false&market_data=true&community_data=false&developer_data=false');
    const supply = d?.market_data?.circulating_supply;
    if (supply > 100e6) return { supply: Math.round(supply), source: 'CoinGecko' };
  } catch (e) {}
  return { supply: 120700000, source: 'fallback' }; // static fallback
}

// ── 4. RWA ON-CHAIN (Ethereum) ────────────────────────────────────────
async function fetchRWA() {
  try {
    // RWA.xyz public API
    const d = await get('https://api.rwa.xyz/v1/metrics/tvl?chain=ethereum');
    if (d?.tvl) return { tvl: d.tvl, source: 'rwa.xyz', updated: new Date().toISOString() };
  } catch (e) {}

  // Fallback: DeFiLlama for tokenized assets category
  try {
    const d = await get('https://api.llama.fi/protocol/ondo-finance');
    if (d?.tvl) {
      // This is just Ondo — use as a proxy indicator
      return { tvl: null, source: 'unavailable', note: 'rwa.xyz API unavailable' };
    }
  } catch (e) {}

  return { tvl: null, source: 'unavailable' };
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
    shares:   sharesR.status   === 'fulfilled' ? sharesR.value   : { error: sharesR.reason?.message },
    clarity:  clarityR.status  === 'fulfilled' ? clarityR.value  : { error: clarityR.reason?.message },
    ethSupply: ethSupplyR.status === 'fulfilled' ? ethSupplyR.value : { error: ethSupplyR.reason?.message },
    rwa:      rwaR.status      === 'fulfilled' ? rwaR.value      : { error: rwaR.reason?.message }
  });
};
