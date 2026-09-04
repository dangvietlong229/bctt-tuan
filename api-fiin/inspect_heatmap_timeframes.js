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

async function test(name, url, headers) {
  try {
    const res = await client.fetch(url, { headers });
    if (res.ok) {
      const data = await res.json();
      console.log(`\n=== ${name} ===`);
      console.log(JSON.stringify(data, null, 2).slice(0, 500));
    }
  } catch (e) {
    console.error(e);
  }
}

async function main() {
  await tokens.start();
  await tokens.ready();
  const u0Headers = { ...baseHeaders, u0: tokens.token };

  await test("OneWeek", "https://market.fiintrade.vn/HeatMap/GetHeatMap?ComGroupCode=VNINDEX&Type=Sector&TimeRange=OneWeek&language=vi", u0Headers);

  tokens.stop();
}

main().catch(console.error);
