'use strict';

/**
 * pull_ticker.js <CODE> [years]
 *
 * Pulls the full "Thống kê giá" (Time & Sales / price-statistics) panel for a
 * single symbol through the rate-limited FiinClient:
 *
 *   1. TimeAndSales/GetTimeAndSales          (tick table, auto-paginated)
 *   2. TimeAndSales/GetTimeAndSalesBuSdChart (buy/sell up-down chart)
 *   3. PriceData/GetLatestPrice              (quote header)
 *   4. PriceData/GetVWAP                     (volume-weighted avg price)
 *   5. PriceDepth/GetPriceDepth              (bid/ask depth)
 *   6. TradingView/GetStockChartData         (OHLC history, year-chunked)
 *   7. TradingView/GetStockEvents            (corporate events)
 *
 * Type (Stock/Index/Derivative) and OrganCode are auto-resolved from the
 * master data in responses/. Only the ticker is required.
 *
 * Usage:  node pull_ticker.js TCB
 *         node pull_ticker.js VNINDEX 3
 */

const fs = require('fs');
const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const { writeTextAtomic } = require('./lib/io');
const { normalizeTicker, parseBoundedInteger } = require('./lib/validation');
const { fetchWithLiveToken } = require('./lib/api');

const T = 'https://technical.fiintrade.vn';
const RESP = path.join(__dirname, 'responses');

const baseHeaders = {
  accept: 'application/json, text/plain, */*',
  'accept-language': 'vi,en-US;q=0.9,en;q=0.8',
  authorization: 'Bearer',
  origin: 'https://fiintrade.vn',
  referer: 'https://fiintrade.vn/',
  'user-agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
};

// Known index group codes -> Type=Index
const INDEX_CODES = new Set([
  'VNINDEX', 'HNXINDEX', 'HNXIndex', 'UPCOMINDEX', 'UpcomIndex',
  'VN30', 'VN100', 'VNX50', 'HNX30', 'VNALL', 'VNMID', 'VNSML',
]);

function loadOrganizations() {
  try {
    const d = require(path.join(RESP, 'GetListOrganization_vi.json'));
    const arr = d.items || d.data || (Array.isArray(d) ? d : []);
    return new Map(arr.map((o) => [String(o.ticker).toUpperCase(), o]));
  } catch {
    return new Map();
  }
}

/** Resolve the instrument Type and OrganCode for a code. */
function resolve(code) {
  const up = code.toUpperCase();
  if (INDEX_CODES.has(up)) return { type: 'Index', organCode: up };
  const orgs = loadOrganizations();
  if (orgs.has(up)) {
    const organization = orgs.get(up);
    return { type: 'Stock', organCode: organization.organCode || organization.code || up };
  }
  // Unknown -> default to Stock; the API will 400/404 if wrong.
  return { type: 'Stock', organCode: up };
}

function isoYearsAgo(years) {
  const d = new Date();
  d.setFullYear(d.getFullYear() - years);
  return d.toISOString();
}

async function main() {
  if (!process.argv[2]) {
    console.error('Usage: node pull_ticker.js <CODE> [years]');
    process.exit(1);
  }
  const code = normalizeTicker(process.argv[2]);
  const years = parseBoundedInteger(process.argv[3], 2, { min: 1, max: 10, label: 'years' });

  const { type, organCode } = resolve(code);
  console.log(`Symbol ${code}  ->  Type=${type}, OrganCode=${organCode}, history=${years}y`);

  const client = new FiinClient();
  const tokens = new TokenProvider();
  await tokens.start();
  await tokens.ready();
  console.log('Live token acquired.\n');

  const outDir = path.join(RESP, code);
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const save = (name, text) => {
    let body = text;
    try { body = JSON.stringify(JSON.parse(text), null, 2); } catch {}
    writeTextAtomic(path.join(outDir, `${name}.json`), `${body}\n`, (candidate) => JSON.parse(candidate));
  };

  let failures = 0;

  async function get(name, url) {
    try {
      const res = await fetchWithLiveToken(client, tokens, url, {}, { baseHeaders });
      const text = await res.text();
      const ok = res.status === 200;
      console.log(`${ok ? '✓' : '✗'} [${res.status}] ${name} ${ok ? `(${text.length}b)` : text.slice(0, 40)}`);
      if (ok) save(name, text);
      if (!ok) failures++;
      return ok ? text : null;
    } catch (e) {
      if (e.message.includes('Circuit breaker')) { console.error(`\n${e.message}`); throw e; }
      console.log(`✗ [ERR] ${name} ${e.message}`);
      failures++;
      return null;
    }
  }

  try {
    // Simple single-shot endpoints
    await get('GetLatestPrice', `${T}/PriceData/GetLatestPrice?Code=${code}&language=vi`);
    await get('GetVWAP', `${T}/PriceData/GetVWAP?Code=${code}&language=vi`);
    await get('GetPriceDepth', `${T}/PriceDepth/GetPriceDepth?Code=${code}&language=vi`);
    await get('GetTimeAndSalesBuSdChart', `${T}/TimeAndSales/GetTimeAndSalesBuSdChart?Code=${code}&offset=0&page=1&language=vi`);
    await get('GetStockEvents', `${T}/TradingView/GetStockEvents?OrganCode=${encodeURIComponent(organCode)}&From=${isoYearsAgo(years)}&To=${new Date().toISOString()}&language=vi`);

    // OHLC history, chunked one year at a time (the panel's own anti-limit pattern)
    const chart = [];
    for (let y = 0; y < years; y++) {
      const to = isoYearsAgo(y);
      const from = isoYearsAgo(y + 1);
      const txt = await get(
        `GetStockChartData_y${y}`,
        `${T}/TradingView/GetStockChartData?Code=${code}&DerivativeCode=&Frequency=Daily&From=${from}&To=${to}&Type=${type}&language=vi`
      );
      if (txt) chart.push({ from, to, txt });
    }

    // Time & Sales tick table, walk pages until empty (cap to stay polite)
    const MAX_PAGES = Number(process.env.FIIN_TAS_PAGES) || 5;
    for (let page = 1; page <= MAX_PAGES; page++) {
      const txt = await get(`GetTimeAndSales_p${page}`, `${T}/TimeAndSales/GetTimeAndSales?Code=${code}&offset=0&page=${page}&language=vi`);
      if (!txt) break;
      let items;
      try { const j = JSON.parse(txt); items = j.items || j.data || []; } catch { items = []; }
      if (!items || items.length === 0) { console.log(`  (page ${page} empty — stopping tick pagination)`); break; }
    }

    if (failures > 0) throw new Error(`${failures} ticker requests failed; incomplete run was not accepted`);
    console.log(`\nDone. Saved to ${path.relative(__dirname, outDir)}/`);
  } finally {
    tokens.stop();
  }
}

main().catch(() => process.exit(1));
