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
      const firstItem = data.items ? data.items[0] : (Array.isArray(data) ? data[0] : null);
      if (firstItem) {
        console.log(`[HeatMapTest] ${name}: name=${firstItem.name}, rate=${firstItem.rate}`);
      } else {
        console.log(`[HeatMapTest] ${name} returned empty items`);
      }
    } else {
      console.log(`[HeatMapTest] ${name} failed: ${res.status}`);
    }
  } catch (e) {
    console.log(`[HeatMapTest] ${name} error: ${e.message}`);
  }
}

async function main() {
  await tokens.start();
  await tokens.ready();
  const u0Headers = { ...baseHeaders, u0: tokens.token };

  await testUrl("Today", "https://market.fiintrade.vn/HeatMap/GetHeatMap?ComGroupCode=VNINDEX&Type=Sector&TimeRange=Today&language=vi", u0Headers);
  await testUrl("OneWeek", "https://market.fiintrade.vn/HeatMap/GetHeatMap?ComGroupCode=VNINDEX&Type=Sector&TimeRange=OneWeek&language=vi", u0Headers);
  await testUrl("OneMonth", "https://market.fiintrade.vn/HeatMap/GetHeatMap?ComGroupCode=VNINDEX&Type=Sector&TimeRange=OneMonth&language=vi", u0Headers);

  tokens.stop();
}

main().catch(console.error);
