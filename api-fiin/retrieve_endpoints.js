const fs = require('fs');
const path = require('path');
const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const { writeJsonAtomic } = require('./lib/io');
const { fetchWithLiveToken } = require('./lib/api');

const client = new FiinClient();
const tokens = new TokenProvider();

const OUTPUT_DIR = path.join(__dirname, 'responses');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR);
}

// Exactly match the browser headers that returned 200 OK with the latest u0 token
const headers = {
  "accept": "application/json, text/plain, */*",
  "accept-language": "vi,en-US;q=0.9,en;q=0.8,zh-CN;q=0.7,zh;q=0.6",
  "authorization": "Bearer",
  "sec-ch-ua": "\"Google Chrome\";v=\"149\", \"Chromium\";v=\"149\", \"Not)A;Brand\";v=\"24\"",
  "sec-ch-ua-mobile": "?0",
  "sec-ch-ua-platform": "\"Windows\"",
  "sec-fetch-dest": "empty",
  "sec-fetch-mode": "cors",
  "sec-fetch-site": "same-site",
  "origin": "https://fiintrade.vn",
  "referer": "https://fiintrade.vn/",
  "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36"
};

const negotiateHeaders = {
  "accept": "*/*",
  "accept-language": "vi,en-US;q=0.9,en;q=0.8,zh-CN;q=0.7,zh;q=0.6",
  "content-type": "text/plain;charset=UTF-8",
  "sec-ch-ua": "\"Google Chrome\";v=\"149\", \"Chromium\";v=\"149\", \"Not)A;Brand\";v=\"24\"",
  "sec-ch-ua-mobile": "?0",
  "sec-ch-ua-platform": "\"Windows\"",
  "sec-fetch-dest": "empty",
  "sec-fetch-mode": "cors",
  "sec-fetch-site": "same-site",
  "x-requested-with": "XMLHttpRequest",
  "origin": "https://fiintrade.vn",
  "referer": "https://fiintrade.vn/",
  "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36"
};

const chartTo = new Date();
const chartFrom = new Date(chartTo);
chartFrom.setFullYear(chartFrom.getFullYear() - 1);

const endpoints = [
  // SignalR Negotiation Endpoints (POST, credentials omitted)
  { name: "RealtimeHub_negotiate", url: "https://realtime.fiintrade.vn/RealtimeHub/negotiate", method: "POST", headers: negotiateHeaders, body: "" },
  { name: "MarketHub_negotiate", url: "https://realtime.fiintrade.vn/MarketHub/negotiate", method: "POST", headers: negotiateHeaders, body: "" },
  { name: "NotificationHub_negotiate", url: "https://realtime.fiintrade.vn/NotificationHub/negotiate", method: "POST", headers: negotiateHeaders, body: "" },
  
  // REST API Endpoints
  { name: "GetTimeOffset", url: "https://core.fiintrade.vn/Master/GetTimeOffset?clientTime=" + new Date().toISOString() },
  { name: "GetHotNews", url: "https://core.fiintrade.vn/UtilFeature/GetHotNews?language=vi" },
  { name: "GetChangeLogsV3", url: "https://core.fiintrade.vn/UserSetting/GetChangeLogsV3?lastVersion=&language=vi" },
  { name: "GetGuestSettings", url: "https://core.fiintrade.vn/Master/GetGuestSettings?language=vi" },
  { name: "GetListOrganization", url: "https://core.fiintrade.vn/Master/GetListOrganization?language=vi" },
  { name: "GetAllCoveredWarrants", url: "https://core.fiintrade.vn/Master/GetAllCoveredWarrants?language=vi" },
  { name: "GetAllCompanyGroup", url: "https://core.fiintrade.vn/Master/GetAllCompanyGroup?language=vi" },
  { name: "GetAllIcbIndustry", url: "https://core.fiintrade.vn/Master/GetAllIcbIndustry?language=vi" },
  { name: "GetAllDerivatives", url: "https://core.fiintrade.vn/Master/GetAllDerivatives?language=vi" },
  { name: "GetAllDerivativesForChart", url: "https://core.fiintrade.vn/Master/GetAllDerivativesForChart?language=vi" },
  { name: "GetAllChartEconomy", url: "https://core.fiintrade.vn/Master/GetAllChartEconomy?language=vi" },
  { name: "GetChartLayouts", url: "https://core.fiintrade.vn/UserSetting/GetChartLayouts?language=vi" },
  { name: "GetLatestIndices", url: "https://market.fiintrade.vn/MarketInDepth/GetLatestIndices?pageSize=99999&status=1&language=vi" },
  { name: "GetIndexSeries", url: "https://market.fiintrade.vn/MarketInDepth/GetIndexSeries?ComGroupCode=VNINDEX&TimeRange=OneDay&id=1&language=vi" },
  { name: "GetAllSystemAlerts", url: "https://tools.fiintrade.vn/Alert/GetAllSystemAlerts?language=vi" },
  { name: "GetPersonalSubsribedAlerts", url: "https://tools.fiintrade.vn/PersonalAlert/GetPersonalSubsribedAlerts?language=vi" },
  { name: "GetPersonalAlertTypes", url: "https://tools.fiintrade.vn/PersonalAlert/GetPersonalAlertTypes?language=vi" },
  { name: "GetLatestPrice", url: "https://technical.fiintrade.vn/PriceData/GetLatestPrice?Code=AAA&language=vi" },
  { name: "GetVWAP", url: "https://technical.fiintrade.vn/PriceData/GetVWAP?Code=AAA&language=vi" },
  { name: "GetTimeAndSales", url: "https://technical.fiintrade.vn/TimeAndSales/GetTimeAndSales?Code=AAA&offset=0&page=1&language=vi" },
  { name: "GetTimeAndSalesBuSdChart", url: "https://technical.fiintrade.vn/TimeAndSales/GetTimeAndSalesBuSdChart?Code=AAA&offset=0&page=1&language=vi" },
  { name: "GetPriceDepth", url: "https://technical.fiintrade.vn/PriceDepth/GetPriceDepth?Code=AAA&language=vi" },
  { name: "GetStockChartData", url: `https://technical.fiintrade.vn/TradingView/GetStockChartData?Code=VNINDEX&DerivativeCode=&Frequency=Daily&From=${encodeURIComponent(chartFrom.toISOString())}&To=${encodeURIComponent(chartTo.toISOString())}&Type=Index&language=vi` }
];

async function run() {
  console.log(`Starting API retrieval with updated browser User-Agent and headers for ${endpoints.length} endpoints...`);
  let successCount = 0;
  let errorCount = 0;

  for (const ep of endpoints) {
    const method = ep.method || "GET";
    console.log(`\nFetching ${ep.name} [${method}] from: ${ep.url}`);
    let response;
    let dataText = "";

    try {
      // Negotiate endpoints use their own header set; REST GETs get a live u0.
      const fetchOptions = {
        method: method,
        headers: ep.headers || {}
      };
      if (ep.body !== undefined) {
        fetchOptions.body = ep.body;
      }

      // Paced, retrying, circuit-broken fetch (see lib/http.js).
      response = ep.headers
        ? await client.fetch(ep.url, fetchOptions)
        : await fetchWithLiveToken(client, tokens, ep.url, fetchOptions, { baseHeaders: headers });

      console.log(`Response Status: ${response.status} ${response.statusText}`);
      dataText = await response.text();

      if (!response.ok) {
        throw new Error(`${ep.name} failed with HTTP ${response.status}: ${dataText.slice(0, 120)}`);
      }
      if (!dataText.trim()) throw new Error(`${ep.name} returned an empty body`);

      const fileSuffix = ep.url.includes("language=vi") ? "_vi" : "";
      const filename = path.join(OUTPUT_DIR, `${ep.name}${fileSuffix}.json`);
      
      const parsed = JSON.parse(dataText);
      writeJsonAtomic(filename, parsed);
      console.log(`Saved output to ${filename}`);
      successCount++;
    } catch (error) {
      // Circuit breaker tripped -> stop the whole run instead of grinding on.
      if (error.message && error.message.includes('Circuit breaker open')) {
        console.error(`\n${error.message}`);
        break;
      }
      console.error(`Fetch error for ${ep.name}:`, error.message);
      errorCount++;
    }
    // Pacing (min-delay + jitter + token bucket) is enforced inside client.fetch().
  }

  console.log(`\nExecution finished. Success: ${successCount}, Errors/Skipped: ${errorCount}`);
  if (errorCount > 0) throw new Error(`${errorCount} endpoints failed; existing response files were preserved`);
}

async function main() {
  console.log("Opening realtime hub to capture a live u0 token...");
  await tokens.start();
  await tokens.ready();
  console.log('Live token acquired.');
  try {
    await run();
  } finally {
    tokens.stop();
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
