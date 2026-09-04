'use strict';

const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const { fetchWithLiveToken, readJsonResponse, requireItems } = require('./lib/api');
const { writeJsonAtomic } = require('./lib/io');
const { normalizeTicker, parseAsOf, validatePullData } = require('./lib/validation');

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
  // Parse tickers from CLI arguments: --tickers=ACB,TCB,...
  let tickers = [];
  const tickersArg = process.argv.find(arg => arg.startsWith('--tickers='));
  if (tickersArg) {
    tickers = [...new Set(tickersArg.split('=')[1].split(',').filter(Boolean).map(t => normalizeTicker(t)))];
  }
  if (tickers.length === 0) throw new Error('No valid --tickers were supplied');
  const asOfArg = process.argv.find(arg => arg.startsWith('--as-of='));
  const asOf = parseAsOf(asOfArg ? asOfArg.split('=')[1] : undefined);

  console.log(`Starting pull_all_inputs.js. Tickers to query: ${tickers.length}`);
  console.log("Opening realtime hub to capture a live u0 token...");
  await tokens.start();
  await tokens.ready();
  console.log('Successfully obtained a live token.\n');

  const rawData = {
    generatedAt: new Date().toISOString(),
    asOf: asOf.toISOString().slice(0, 10),
    sectors: {},
    indices: {},
    tickers: {}
  };
  const getJson = async (label, url) => {
    const response = await fetchWithLiveToken(client, tokens, url, {}, { baseHeaders });
    return readJsonResponse(response, label);
  };

  // 1. Fetch Sector Performance via HeatMap (Today, OneWeek, OneMonth)
  console.log("Fetching Sector performance from HeatMap...");
  const ranges = ['OneDay', 'OneWeek', 'OneMonth'];
  for (const range of ranges) {
    const url = `https://market.fiintrade.vn/HeatMap/GetHeatMap?ComGroupCode=VNINDEX&Type=Sector&TimeRange=${range}&language=vi`;
    rawData.sectors[range] = requireItems(await getJson(`Sector ${range}`, url), `Sector ${range}`);
    console.log(`✓ Fetched Sector HeatMap for ${range}`);
  }

  // 2. Fetch Index Series (Daily History for VNINDEX and VN30)
  console.log("Fetching VNINDEX and VN30 historical series...");
  const indexCodes = ['VNINDEX', 'VN30'];
  for (const index of indexCodes) {
    const url = `https://market.fiintrade.vn/MarketInDepth/GetIndexSeries?ComGroupCode=${index}&TimeRange=OneYear&id=1&language=vi`;
    rawData.indices[index] = requireItems(await getJson(`${index} index series`, url), `${index} index series`);
    console.log(`✓ Fetched ${index} index series`);
  }

  // 3. Fetch Index Vốn Hóa (Market Cap) from LatestIndices
  console.log("Fetching latest indices (market cap)...");
  {
    const url = 'https://market.fiintrade.vn/MarketInDepth/GetLatestIndices?pageSize=99999&status=1&language=vi';
    rawData.latestIndices = requireItems(await getJson('Latest indices', url), 'Latest indices');
    console.log('✓ Fetched latest indices');
  }

  // 4. Fetch detail for specific tickers (PriceDepth, History, and Indicators)
  if (tickers.length > 0) {
    const organizations = requireItems(
      await getJson(
        'Organizations',
        'https://core.fiintrade.vn/Master/GetListOrganization?language=vi'
      ),
      'Organizations'
    );
    const organizationCodes = new Map(
      organizations.map((item) => [String(item.ticker || '').toUpperCase(), item.organCode || item.code])
    );
    console.log(`Fetching detail for ${tickers.length} tickers...`);
    const todayStr = asOf.toISOString();
    const oneYearAgo = new Date(asOf);
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    const oneYearAgoStr = oneYearAgo.toISOString();

    for (const ticker of tickers) {
      console.log(`-> Fetching details for ticker: ${ticker}`);
      rawData.tickers[ticker] = {};
      const organCode = organizationCodes.get(ticker);
      if (!organCode) throw new Error(`No OrganCode mapping found for ${ticker}`);

      // A. Price Depth (for current price, bid/ask, foreign room, foreign transactions, foreignerPercentage)
      {
        const url = `https://technical.fiintrade.vn/PriceDepth/GetPriceDepth?Code=${encodeURIComponent(ticker)}&language=vi`;
        rawData.tickers[ticker].priceDepth = await getJson(`${ticker}.priceDepth`, url);
      }

      // B. Stock History (for 1w, 1m, YTD, 52w high/low calculation)
      {
        const url = `https://technical.fiintrade.vn/TradingView/GetStockChartData?Code=${encodeURIComponent(ticker)}&DerivativeCode=&Frequency=Daily&From=${encodeURIComponent(oneYearAgoStr)}&To=${encodeURIComponent(todayStr)}&Type=Stock&language=vi`;
        rawData.tickers[ticker].history = await getJson(`${ticker}.history`, url);
      }

      // C. PE Indicator
      {
        const url = `https://technical.fiintrade.vn/TradingView/GetIndicators?OganCode=${encodeURIComponent(organCode)}&Code=PE&Frequency=Daily&From=${encodeURIComponent(oneYearAgoStr)}&To=${encodeURIComponent(todayStr)}&Type=Stock&language=vi`;
        rawData.tickers[ticker].pe = await getJson(`${ticker}.pe`, url);
      }

      // D. PB Indicator
      {
        const url = `https://technical.fiintrade.vn/TradingView/GetIndicators?OganCode=${encodeURIComponent(organCode)}&Code=PB&Frequency=Daily&From=${encodeURIComponent(oneYearAgoStr)}&To=${encodeURIComponent(todayStr)}&Type=Stock&language=vi`;
        rawData.tickers[ticker].pb = await getJson(`${ticker}.pb`, url);
      }
    }
  }

  validatePullData(rawData, tickers, { asOf, maxAgeDays: 7 });
  const outPath = path.join(__dirname, 'responses', 'api_raw_inputs.json');
  writeJsonAtomic(outPath, rawData);
  console.log(`\nAggregated raw data saved to: ${outPath}`);
  console.log("Completed successfully.");
}

main().catch(err => {
  console.error("Fatal error during fetch:", err);
  tokens.stop();
  process.exit(1);
}).finally(() => tokens.stop());
