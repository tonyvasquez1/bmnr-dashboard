const https = require('https');

const FINNHUB_KEY = 'd7foe0hr01qqb8rh1gr0d7foe0hr01qqb8rh1grg';

function get(url) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, {
      headers: { 'User-Agent': 'BMNR-Dashboard/1.0', 'Accept': 'application/json' }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode !== 200) { reject(new Error(`HTTP ${res.statusCode}`)); return; }
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.setTimeout(8000, () => { req.destroy(); reject(new Error('Timeout')); });
  });
}

async function fetchFinnhub(symbol) {
  const url = `https://finnhub.io/api/v1/quote?symbol=${encodeURIComponent(symbol)}&token=${FINNHUB_KEY}`;
  const d = await get(url);
  const price = d.c > 0 ? d.c : d.pc; // use prev close after hours
  if (!price || price <= 0) throw new Error(`No valid price for ${symbol}`);
  return { price, changePct: d.dp || 0, change: d.d || 0, prevClose: d.pc || 0, isMarketOpen: d.c > 0 };
}

async function fetchETH() {
  try {
    const d = await get('https://api.coingecko.com/api/v3/simple/price?ids=ethereum&vs_currencies=usd&include_24hr_change=true');
    const price = d?.ethereum?.usd;
    const changePct = d?.ethereum?.usd_24h_change;
    if (price > 0) { const prev = price/(1+(changePct||0)/100); return { price, changePct:changePct||0, change:price-prev, src:'CoinGecko' }; }
  } catch(e) {}
  return { ...(await fetchFinnhub('BINANCE:ETHUSDT')), src:'Finnhub' };
}

async function fetchBTC() {
  try { return { ...(await fetchFinnhub('BINANCE:BTCUSDT')), src:'Finnhub' }; } catch(e) {}
  const d = await get('https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=usd&include_24hr_change=true');
  const price = d?.bitcoin?.usd;
  if (!price) throw new Error('BTC failed');
  return { price, changePct: d.bitcoin.usd_24h_change||0, src:'CoinGecko' };
}

async function fetchBMNR() {
  return { ...(await fetchFinnhub('BMNR')), src:'Finnhub/NYSE' };
}

async function fetchGold() {
  try {
    const d = await get('https://gold-api.com/price/XAU');
    if (d?.price > 500) return { price:d.price, changePct:d.chp||0, change:d.ch||0, src:'gold-api.com' };
  } catch(e) {}
  try {
    const d = await get('https://api.coingecko.com/api/v3/simple/price?ids=pax-gold&vs_currencies=usd&include_24hr_change=true');
    const price = d?.['pax-gold']?.usd;
    if (price > 500) return { price, changePct:d['pax-gold'].usd_24h_change||0, src:'PAXG' };
  } catch(e) {}
  const d = await get('https://api.frankfurter.app/latest?from=XAU&to=USD');
  const price = d?.rates?.USD;
  if (price > 500) return { price, changePct:0, src:'ECB' };
  throw new Error('All gold sources failed');
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Content-Type', 'application/json');

  const [ethR, btcR, bmnrR, goldR] = await Promise.allSettled([
    fetchETH(), fetchBTC(), fetchBMNR(), fetchGold()
  ]);

  const errors = {};
  if (ethR.status  === 'rejected') errors.eth  = ethR.reason?.message;
  if (btcR.status  === 'rejected') errors.btc  = btcR.reason?.message;
  if (bmnrR.status === 'rejected') errors.bmnr = bmnrR.reason?.message;
  if (goldR.status === 'rejected') errors.gold = goldR.reason?.message;

  res.status(200).json({
    timestamp: new Date().toISOString(),
    eth:   ethR.status  === 'fulfilled' ? ethR.value  : null,
    btc:   btcR.status  === 'fulfilled' ? btcR.value  : null,
    bmnr:  bmnrR.status === 'fulfilled' ? bmnrR.value : null,
    gold:  goldR.status === 'fulfilled' ? goldR.value : null,
    errors: Object.keys(errors).length > 0 ? errors : undefined
  });
};
