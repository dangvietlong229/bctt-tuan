'use strict';

const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const { fetchWithLiveToken, readJsonResponse, requireItems } = require('./lib/api');
const { writeJsonAtomic } = require('./lib/io');

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

async function main() {
  console.log("Starting bridge_pull.js...");
  console.log("Opening realtime hub to capture a live u0 token...");
  await tokens.start();
  await tokens.ready();
  console.log('Successfully obtained a live token.\n');

  const sampleData = {};
  const getJson = async (label, url) => {
    const response = await fetchWithLiveToken(client, tokens, url, {}, { baseHeaders });
    return readJsonResponse(response, label);
  };

  // 1. Fetch Market Liquidity (Index Series with matchValue - last 5 days)
  console.log("Fetching VNINDEX Index Series (with matchValue)...");
  {
    const url = 'https://market.fiintrade.vn/MarketInDepth/GetIndexSeries?ComGroupCode=VNINDEX&TimeRange=ThreeMonths&id=1&language=vi';
    sampleData.vnindexLiquidity = requireItems(await getJson('VNINDEX liquidity', url), 'VNINDEX liquidity').slice(-5);
    console.log(`✓ Fetched VNINDEX historical series: ${sampleData.vnindexLiquidity.length} items`);
  }

  // 2. Fetch Foreign Investor Statistics for TCB (including YTD)
  console.log("Fetching Foreign Investor Statistics for TCB...");
  {
    const url = 'https://market.fiintrade.vn/MoneyFlow/GetStatisticInvestor?ComGroupCode=TCB&language=vi';
    sampleData.foreignInvestorTCB = requireItems(await getJson('TCB foreign statistics', url), 'TCB foreign statistics');
    console.log('✓ Fetched TCB foreign investor stats');
  }

  // 3. Fetch Proprietary Trading (Tự doanh) for VNINDEX (all stocks)
  console.log("Fetching Proprietary Trading for VNINDEX (entire market)...");
  {
    const url = 'https://market.fiintrade.vn/MoneyFlow/GetProprietaryV2?ComGroupCode=VNINDEX&language=vi';
    sampleData.proprietaryMarket = requireItems(await getJson('Proprietary market', url), 'Proprietary market');
    console.log('✓ Fetched market proprietary trading data');
  }

  // 4. Fetch Sector Index Chart (e.g. VNREAL and VNFIN)
  console.log("Fetching Sector Index (VNREAL)...");
  {
    const to = new Date();
    const from = new Date(to);
    from.setMonth(from.getMonth() - 3);
    const url = `https://technical.fiintrade.vn/TradingView/GetStockChartData?Code=VNREAL&DerivativeCode=&Frequency=Daily&From=${encodeURIComponent(from.toISOString())}&To=${encodeURIComponent(to.toISOString())}&Type=Index&language=vi`;
    sampleData.sectorVnReal = requireItems(await getJson('VNREAL sector index', url), 'VNREAL sector index').slice(-5);
    console.log(`✓ Fetched VNREAL sector index: ${sampleData.sectorVnReal.length} items`);
  }

  // Save sample data
  const outPath = path.join(__dirname, 'responses', 'bridge_sample.json');
  writeJsonAtomic(outPath, sampleData);
  console.log(`\nSaved sample data to: ${outPath}`);
  console.log("Done!");
}

main().catch(err => {
  console.error("Fatal error:", err);
  tokens.stop();
  process.exit(1);
}).finally(() => tokens.stop());
