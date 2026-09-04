'use strict';

/**
 * probe_catalog.js — classify every GET endpoint in catalog.json against the
 * fiintrade.vn guest token. GET-only: never fires POST mutations.
 *
 * Sends a kitchen-sink set of common params so "missing param" 400s don't hide
 * a reachable route. Classifies:
 *   200 free | 401/403 gated | 400 needs-params(exists) | 404 wrong | other
 */

const fs = require('fs');
const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');

const rows = require('./catalog.json').filter((r) => r.verb === 'GET');

const COMMON = {
  language: 'vi', Code: 'ACB', OrganCode: 'ACB', ComGroupCode: 'VNINDEX',
  TimeRange: 'OneDay', Frequency: 'Daily', Page: '1', Type: 'Stock',
  Exchange: 'All', Criteria: 'FrBuyVal', Option: 'NetBuyVol',
};
const qs = Object.entries(COMMON).map(([k, v]) => `${k}=${v}`).join('&');

const H = {
  accept: 'application/json, text/plain, */*', 'accept-language': 'vi',
  authorization: 'Bearer', origin: 'https://fiintrade.vn', referer: 'https://fiintrade.vn/',
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/149.0.0.0',
};

function classify(s) {
  if (s === 200) return 'free';
  if (s === 401 || s === 403) return 'gated';
  if (s === 400) return 'needs-params';
  if (s === 404) return 'not-here';
  return `other-${s}`;
}

async function main() {
  const client = new FiinClient({ maxConsecutiveBlocks: 1e9, minDelayMs: 500, jitterMs: 500 });
  const tokens = new TokenProvider();
  await tokens.start(); await tokens.ready();
  console.log(`Probing ${rows.length} GET endpoints with a live token.\n`);

  const results = [];
  const tally = {};
  for (const r of rows) {
    const url = `${r.url}?${qs}`;
    let status;
    try {
      const res = await client.fetch(url, { headers: { ...H, u0: tokens.token } });
      status = res.status;
      await res.text();
    } catch (e) { status = 'ERR'; }
    const cls = classify(status);
    tally[cls] = (tally[cls] || 0) + 1;
    results.push({ method: r.method, base: r.base, status, class: cls });
    const mark = cls === 'free' ? '✓' : cls === 'gated' ? '🔒' : '·';
    console.log(`${mark} ${String(status).padEnd(4)} ${cls.padEnd(12)} ${r.base.replace('https://','').replace(/\/$/,'')}/${r.method}`);
  }

  tokens.stop();
  fs.writeFileSync(path.join(__dirname, 'catalog_probe.json'), JSON.stringify(results, null, 2));
  console.log('\n=== Summary ===');
  for (const [k, v] of Object.entries(tally).sort((a,b)=>b[1]-a[1])) console.log(`  ${String(v).padStart(3)}  ${k}`);
  console.log('\nSaved catalog_probe.json');
}

main().catch((e) => { console.error(e.message); process.exit(1); });
