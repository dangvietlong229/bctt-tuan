'use strict';

/**
 * Probe endpoints discovered from the SSI iBoard iframe dump
 * (fiin-*.ssi.com.vn) against their fiintrade.vn equivalents using our
 * guest u0 token. Tells us which SSI-only endpoints are also reachable
 * on the public fiintrade gateway without SSI's x-fiin-* auth.
 */

const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');

const client = new FiinClient();
const tokens = new TokenProvider();

// SSI host  ->  fiintrade.vn host
// fiin-fundamental has no obvious fiintrade twin; we try technical + core.
const baseHeaders = {
  accept: 'application/json, text/plain, */*',
  'accept-language': 'vi,en-US;q=0.9,en;q=0.8',
  authorization: 'Bearer',
  origin: 'https://fiintrade.vn',
  referer: 'https://fiintrade.vn/',
  'user-agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
};

// Endpoints seen ONLY in the SSI dump (not in retrieve_endpoints.js).
const probes = [
  { name: 'TopMover_GetTopGainers', url: 'https://market.fiintrade.vn/TopMover/GetTopGainers?language=vi&ComGroupCode=All' },
  { name: 'TopMover_GetTopForeignTrading', url: 'https://market.fiintrade.vn/TopMover/GetTopForeignTrading?language=vi&ComGroupCode=All&Option=NetBuyVol' },
  { name: 'TopMover_GetTopBreakout', url: 'https://market.fiintrade.vn/TopMover/GetTopBreakout?language=vi&ComGroupCode=All&TimeRange=OneWeek&Rate=OnePointFive' },
  { name: 'HeatMap_GetHeatMap', url: 'https://market.fiintrade.vn/HeatMap/GetHeatMap?language=vi&Exchange=All&Criteria=FrBuyVal' },
  { name: 'MarketInDepth_GetProspectV2', url: 'https://market.fiintrade.vn/MarketInDepth/GetProspectV2?language=vi&ComGroupCode=VNINDEX' },
  { name: 'MoneyFlow_GetContribution', url: 'https://market.fiintrade.vn/MoneyFlow/GetContribution?language=vi&ComGroupCode=VNINDEX&Type=FreeFloat' },
  { name: 'MoneyFlow_GetForeign', url: 'https://market.fiintrade.vn/MoneyFlow/GetForeign?language=vi&ComGroupCode=VNINDEX' },
  { name: 'Screener_GetScreenerParameters', url: 'https://tools.fiintrade.vn/Screener/GetScreenerParameters?language=vi' },
  { name: 'Alert_GetNotificationList', url: 'https://tools.fiintrade.vn/Alert/GetNotificationList?language=vi&OrganCode=&Page=1&PageSize=50&AlertCode=' },
  { name: 'UserSetting_getTopScreeners', url: 'https://core.fiintrade.vn/UserSetting/getTopScreeners?language=vi' },
  { name: 'UserSetting_GetWorkspace', url: 'https://core.fiintrade.vn/UserSetting/GetWorkspace?language=vi&name=MarketOverview' },
  // fiin-fundamental.ssi.com.vn -> try a couple of fiintrade hosts
  { name: 'Fundamental_GetBalanceSheet_technical', url: 'https://technical.fiintrade.vn/FinancialStatement/GetBalanceSheet?language=vi&OrganCode=ACB' },
  { name: 'Fundamental_GetBalanceSheet_core', url: 'https://core.fiintrade.vn/FinancialStatement/GetBalanceSheet?language=vi&OrganCode=ACB' },
  { name: 'Fundamental_GetCompanyScore_technical', url: 'https://technical.fiintrade.vn/Snapshot/GetCompanyScore?language=vi&OrganCode=ACB' },
  // POST screener
  {
    name: 'Screener_GetScreenerItems',
    url: 'https://tools.fiintrade.vn/Screener/GetScreenerItems',
    method: 'POST',
    body: JSON.stringify({ comGroupCode: 'All', icbCode: 'All', parameters: [], page: 1, pageSize: 30, OrderBy: 'StockScreenerItem.Ticker', Direction: 'ASC' }),
  },
];

async function main() {
  console.log('Capturing live u0 token...');
  await tokens.start();
  await tokens.ready();
  console.log('Live token acquired.\n');

  const results = [];
  for (const p of probes) {
    const headers = { ...baseHeaders, u0: tokens.token };
    if (p.method === 'POST') headers['content-type'] = 'application/json';
    const init = { method: p.method || 'GET', headers };
    if (p.body) init.body = p.body;

    let status, note = '';
    try {
      const res = await client.fetch(p.url, init);
      status = res.status;
      const text = await res.text();
      note = res.ok ? `${text.length} bytes` : text.slice(0, 80).replace(/\s+/g, ' ');
    } catch (e) {
      if (e.message.includes('Circuit breaker')) { console.log('\nCircuit breaker tripped — stopping.'); break; }
      status = 'ERR';
      note = e.message;
    }
    const flag = status === 200 ? '✓' : '✗';
    console.log(`${flag} [${status}] ${p.name}  ${note}`);
    results.push({ name: p.name, status, note });
  }

  tokens.stop();
  const ok = results.filter((r) => r.status === 200).length;
  console.log(`\nReachable on fiintrade.vn with guest u0: ${ok}/${results.length}`);
}

main();
