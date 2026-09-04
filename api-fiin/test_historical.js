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

async function testUrl(name, url, headers) {
  try {
    const res = await client.fetch(url, { headers });
    if (res.ok) {
      const data = await res.json();
      const count = data.items ? data.items.length : (Array.isArray(data) ? data.length : 0);
      console.log(`[HistoricalTest] ${name}: count=${count}`);
      if (count > 0) {
        const item = data.items ? data.items[0] : (Array.isArray(data) ? data[0] : data);
        console.log(`Sample:`, JSON.stringify(item).slice(0, 300));
      }
    } else {
      console.log(`[HistoricalTest] ${name} failed: ${res.status}`);
    }
  } catch (e) {
    console.log(`[HistoricalTest] ${name} error: ${e.message}`);
  }
}

async function main() {
  await tokens.start();
  await tokens.ready();
  const u0Headers = { ...baseHeaders, u0: tokens.token };

  const endpoints = [
    // 1. GetIndexSeries with larger TimeRange
    { name: "GetIndexSeries_ThreeMonths", url: "https://market.fiintrade.vn/MarketInDepth/GetIndexSeries?ComGroupCode=VNINDEX&TimeRange=ThreeMonths&id=1&language=vi" },
    { name: "GetIndexSeries_OneYear", url: "https://market.fiintrade.vn/MarketInDepth/GetIndexSeries?ComGroupCode=VNINDEX&TimeRange=OneYear&id=1&language=vi" },
    
    // 2. GetForeign with TimeRange
    { name: "GetForeign_ThreeMonths", url: "https://market.fiintrade.vn/MoneyFlow/GetForeign?ComGroupCode=TCB&TimeRange=ThreeMonths&language=vi" },
    { name: "GetForeign_OneYear", url: "https://market.fiintrade.vn/MoneyFlow/GetForeign?ComGroupCode=TCB&TimeRange=OneYear&language=vi" },
    
    // 3. TradingView Chart for Index with matchValue
    { name: "TradingView_VNINDEX_Daily", url: "https://technical.fiintrade.vn/TradingView/GetStockChartData?Code=VNINDEX&DerivativeCode=&Frequency=Daily&From=2026-06-01T00:00:00.000Z&To=2026-06-24T00:00:00.000Z&Type=Index&language=vi" },
  ];

  for (const ep of endpoints) {
    await testUrl(ep.name, ep.url, u0Headers);
  }

  tokens.stop();
}

main().catch(console.error);
