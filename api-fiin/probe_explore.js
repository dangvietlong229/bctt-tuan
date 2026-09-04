'use strict';

/**
 * probe_explore.js — POLITE sampling probe of unexplored controllers.
 *
 * One representative read-only GET per controller, circuit-breaker ON,
 * jittered pacing. GET-only (never mutations). If a sample 200s the whole
 * controller is likely guest-free; if it 401s we mark it gated and move on.
 */

const fs = require('fs');
const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');

const H = {
  accept: 'application/json, text/plain, */*', 'accept-language': 'vi',
  authorization: 'Bearer', origin: 'https://fiintrade.vn', referer: 'https://fiintrade.vn/',
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/149.0.0.0',
};
const C = 'language=vi&Code=ACB&OrganCode=ACB&ComGroupCode=VNINDEX&TimeRange=OneDay&Frequency=Daily&Page=1&PageSize=20&Type=Stock';

// One representative GET per unexplored controller.
const samples = [
  ['Calendar',        `https://market.fiintrade.vn/Calendar/GetCorporateEarning?${C}`],
  ['SectorIndepth',   `https://market.fiintrade.vn/SectorIndepth/GetSectorPerformance?${C}`],
  ['News',            `https://news.fiintrade.vn/News/GetMostRecent?${C}`],
  ['Strategy',        `https://strategy.fiintrade.vn/Strategy/GetLeaders?${C}`],
  ['TAStrategy',      `https://strategy.fiintrade.vn/TAStrategy/GetBreakout?${C}`],
  ['Rankings',        `https://strategy.fiintrade.vn/Rankings/GetRanking?${C}`],
  ['Ownership',       `https://fundamental.fiintrade.vn/Ownership/GetOwnership?${C}`],
  ['BUSD',            `https://market.fiintrade.vn/BUSD/GetBUSD?${C}`],
  ['MarketInDepth',   `https://market.fiintrade.vn/MarketInDepth/GetProspect?${C}`],
  ['Snapshot',        `https://fundamental.fiintrade.vn/Snapshot/GetSnapshot?${C}`],
  ['Valuation',       `https://tools.fiintrade.vn/Valuation/GetValuation?${C}`],
  ['TopMover',        `https://market.fiintrade.vn/TopMover/GetTopLosers?${C}`],
];

function classify(s) {
  if (s === 200) return 'FREE ✓';
  if (s === 401 || s === 403) return 'gated 🔒';
  if (s === 400) return 'needs-params (exists)';
  if (s === 404) return 'wrong-host';
  return `other-${s}`;
}

async function main() {
  // Breaker ON (default 3) — polite. Slow pacing.
  const client = new FiinClient({ minDelayMs: 700, jitterMs: 600 });
  const tokens = new TokenProvider();
  await tokens.start(); await tokens.ready();
  console.log(`Sampling ${samples.length} controllers with the circuit breaker enabled.\n`);

  const out = [];
  for (const [ctrl, url] of samples) {
    let status, size = 0;
    try {
      const res = await client.fetch(url, { headers: { ...H, u0: tokens.token } });
      status = res.status;
      size = (await res.text()).length;
    } catch (e) {
      if (e.message.includes('Circuit breaker')) { console.log('\n🔒 Breaker tripped (3 consecutive gated) — stopping politely.'); break; }
      status = 'ERR';
    }
    out.push({ ctrl, status, size });
    console.log(`${String(status).padEnd(4)} ${classify(status).padEnd(22)} ${ctrl}${size?`  (${size}b)`:''}`);
  }
  tokens.stop();
  fs.writeFileSync(path.join(__dirname, 'explore_probe.json'), JSON.stringify(out, null, 2));
  console.log('\nSaved explore_probe.json');
}

main().catch((e) => { console.error(e.message); process.exit(1); });
