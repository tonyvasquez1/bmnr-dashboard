const https = require('https');

const FINNHUB_KEY = 'd7foe0hr01qqb8rh1gr0d7foe0hr01qqb8rh1grg';

function get(url) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: { 'User-Agent': 'BMNR-Dashboard/1.0' } }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(e); }
      });
    }).on('error', reject);
  });
}

async function fetchFinnhub(symbol) {
  const url = `https://finnhub.io/api/v1/quote?symbol=${symbol}&token=${FINNHUB_KEY}`;
  const d = await get(url);
  if (!d?.c || d.c <= 0) throw new Error(`Bad Finnhub response for ${symbol}`);
  return { price: d.c, changePct: d.dp || 0, change: d.d || 0, prevClose: d.pc || 0 };
}

async function fetchETH() {
  const d = await get('https://api.coingecko.com/api/v3/simple/price?ids=ethereum&vs_currencies=usd&include_24hr_change=true');
  const price = d?.ethereum?.usd;
  const changePct = d?.ethereum?.usd_24h_change;
  if (!price || price <= 0) throw new Error('Bad ETH response');
  const prev = price / (1 + (changePct || 0) / 100);
  return { price, changePct: changePct || 0, change: price - prev, prevClose: prev };
}

async function fetchGold() {
  try {
    const d = await get('https://gold-api.com/price/XAU');
    if (d?.price > 500) return { price: d.price, changePct: d.chp || 0, change: d.ch || 0, src: 'gold-api.com' };
  } catch (e) {}
  const d = await get('https://api.coingecko.com/api/v3/simple/price?ids=pax-gold&vs_currencies=usd&include_24hr_change=true');
  const price = d?.['pax-gold']?.usd;
  const changePct = d?.['pax-gold']?.usd_24h_change;
  if (!price || price <= 0) throw new Error('Bad Gold response');
  return { price, changePct: changePct || 0, src: 'PAXG' };
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.setHeader('Content-Type', 'application/json');

  const [ethR, btcR, bmnrR, goldR] = await Promise.allSettled([
    fetchETH(),
    fetchFinnhub('BINANCE:BTCUSDT'),
    fetchFinnhub('BMNR'),
    fetchGold()
  ]);

  const result = {
    timestamp: new Date().toISOString(),
    eth:  ethR.status  === 'fulfilled' ? ethR.value  : null,
    btc:  btcR.status  === 'fulfilled' ? btcR.value  : null,
    bmnr: bmnrR.status === 'fulfilled' ? bmnrR.value : null,
    gold: goldR.status === 'fulfilled' ? goldR.value : null,
  };

  res.status(200).json(result);
};
