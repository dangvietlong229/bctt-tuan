'use strict';

const fs = require('fs');
const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');

const client = new FiinClient();
const tokens = new TokenProvider();

const baseHeaders = {
  accept: 'application/json, text/plain, */*',
  'accept-language': 'vi,en-US;q=0.9,en;q=0.8',
  authorization: 'Bearer',
  origin: 'https://fiintrade.vn',
  referer: 'https://fiintrade.vn/',
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
};

async function testEndpoint(name, url, headers) {
  try {
    const res = await client.fetch(url, { headers });
    if (res.ok) {
      const data = await res.json();
      console.log(`\n=================== [${name}] ===================`);
      console.log(`URL: ${url}`);
      console.log(`Type: ${Array.isArray(data) ? 'Array' : 'Object'}`);
      const items = data.items || (Array.isArray(data) ? data : null);
      if (items) {
        console.log(`Count: ${items.length}`);
        if (items.length > 0) {
          console.log(`Sample Item:`, JSON.stringify(items[0], null, 2).slice(0, 800));
        }
      } else {
        console.log(`Keys:`, Object.keys(data));
        console.log(`Content Sample:`, JSON.stringify(data, null, 2).slice(0, 800));
      }
    } else {
      console.log(`\n[${name}] Failed: ${res.status}`);
    }
  } catch (e) {
    console.log(`\n[${name}] Error: ${e.message}`);
  }
}

async function main() {
  await tokens.start();
  await tokens.ready();
  const u0Headers = { ...baseHeaders, u0: tokens.token };

  // 1. Check Index Series (Market Liquidity)
  await testEndpoint(
    "GetIndexSeries_VNINDEX",
    "https://market.fiintrade.vn/MarketInDepth/GetIndexSeries?ComGroupCode=VNINDEX&TimeRange=OneDay&id=1&language=vi",
    u0Headers
  );

  // 2. Check Liquidity Series
  await testEndpoint(
    "GetLiquiditySeries",
    "https://market.fiintrade.vn/MarketInDepth/GetLiquiditySeries?ComGroupCode=VNINDEX&language=vi",
    u0Headers
  );

  // 3. Check Foreign Trading (GetForeign for TCB and All)
  await testEndpoint(
    "GetForeign_TCB",
    "https://market.fiintrade.vn/MoneyFlow/GetForeign?ComGroupCode=TCB&language=vi",
    u0Headers
  );

  // 4. Check Proprietary Trading (GetProprietaryV2 for VNINDEX)
  await testEndpoint(
    "GetProprietaryV2_VNINDEX",
    "https://market.fiintrade.vn/MoneyFlow/GetProprietaryV2?ComGroupCode=VNINDEX&language=vi",
    u0Headers
  );

  // 5. Check Sector Performance Index via TradingView Chart (e.g. VNREAL, VNFIN)
  await testEndpoint(
    "TradingView_VNREAL",
    "https://technical.fiintrade.vn/TradingView/GetStockChartData?Code=VNREAL&DerivativeCode=&Frequency=Daily&From=2026-06-15T00:00:00.000Z&To=2026-06-24T00:00:00.000Z&Type=Index&language=vi",
    u0Headers
  );

  tokens.stop();
}

main().catch(console.error);
