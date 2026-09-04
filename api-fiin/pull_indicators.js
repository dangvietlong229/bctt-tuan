'use strict';

/**
 * pull_indicators.js <CODE> [indicator...]
 *
 * Pulls fundamental-ratio time-series for a ticker via the FREE
 * technical/TradingView/GetIndicators endpoint — the guest-accessible route
 * to ROE/margins/EPS/PE/etc. that sidesteps the gated FinancialAnalysis.
 *
 * Each indicator returns items[] of {tradingDate, value}. Results are merged
 * into one wide table (date x indicator) and saved to responses/<CODE>/.
 *
 * Usage:
 *   node pull_indicators.js TCB
 *   node pull_indicators.js TCB ROE PE PB
 */

const fs = require('fs');
const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const { normalizeTicker } = require('./lib/validation');
const { writeJsonAtomic } = require('./lib/io');
const { fetchWithLiveToken } = require('./lib/api');

const TV = 'https://technical.fiintrade.vn/TradingView/GetIndicators';

// Verified guest-free indicator codes (from the bundle's chart-indicator config).
// NOTE: the symbol param is the misspelled "OganCode" — that is what the
// backend expects; the correctly-spelled "OrganCode" returns blanked zeros.
const DEFAULT_INDICATORS = [
  // valuation (daily series)
  'PE', 'PB', 'PS', 'DividendYield',
  // profitability (report-date steps)
  'ROE', 'ROA', 'NetProfitMargin', 'GrossProfitMargin', 'EBITMargin',
  'NetProfit', 'GrossProfit', 'NetIncome', 'NetRevenue', 'OI',
  // balance sheet aggregates (report-date steps)
  'TotalAsset', 'Equity', 'TotalLiabilities', 'CurrentAsset',
  'CurrentLiabilities', 'TotalLoans', 'TotalDeposit',
];

const baseHeaders = {
  accept: 'application/json, text/plain, */*',
  'accept-language': 'vi,en-US;q=0.9,en;q=0.8',
  authorization: 'Bearer',
  origin: 'https://fiintrade.vn',
  referer: 'https://fiintrade.vn/',
  'user-agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
};

async function main() {
  if (!process.argv[2]) {
    console.error('Usage: node pull_indicators.js <CODE> [indicator...]');
    process.exit(1);
  }
  const code = normalizeTicker(process.argv[2]);
  const indicators = process.argv.slice(3).length
    ? process.argv.slice(3).map((value) => {
        if (!/^[A-Za-z][A-Za-z0-9]{1,31}$/.test(value)) throw new Error(`Invalid indicator: ${value}`);
        return value;
      })
    : DEFAULT_INDICATORS;

  const client = new FiinClient();
  const tokens = new TokenProvider();
  await tokens.start();
  await tokens.ready();
  console.log(`Indicators for ${code} using a live token.\n`);

  const outDir = path.join(__dirname, 'responses', code);
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const From = '1999-12-31T17:00:00.000Z';
  const To = new Date().toISOString();
  const merged = {}; // date -> { indicator: value }
  let failures = 0;

  for (const ind of indicators) {
    // "OganCode" misspelling is intentional — the backend expects it.
    const url =
      `${TV}?OganCode=${code}&Code=${ind}&Frequency=Daily` +
      `&From=${From}&To=${To}&Type=Stock&language=vi`;
    try {
      const res = await fetchWithLiveToken(client, tokens, url, {}, { baseHeaders });
      const text = await res.text();
      if (!res.ok) { console.log(`✗ [${res.status}] ${ind}`); failures++; continue; }
      const items = (JSON.parse(text).items) || [];
      // keep only points where value changes (these series step at report dates)
      let last;
      let kept = 0;
      for (const it of items) {
        if (it.value === last) continue;
        last = it.value;
        (merged[it.tradingDate] ||= {})[ind] = it.value;
        kept++;
      }
      console.log(`✓ [200] ${ind.padEnd(16)} ${items.length} pts (${kept} change-points)`);
    } catch (e) {
      if (e.message.includes('Circuit breaker')) { console.error(`\n${e.message}`); break; }
      console.log(`✗ [ERR] ${ind} ${e.message}`);
      failures++;
    }
  }

  tokens.stop();

  const dates = Object.keys(merged).sort();
  const table = dates.map((d) => ({ tradingDate: d, ...merged[d] }));
  if (failures > 0 || table.length === 0) {
    throw new Error(`Indicator pull incomplete: ${failures} failures, ${table.length} rows`);
  }
  writeJsonAtomic(path.join(outDir, 'Indicators.json'), table);
  console.log(`\nSaved ${table.length} dated rows -> responses/${code}/Indicators.json`);
}

main().catch(() => process.exit(1));
