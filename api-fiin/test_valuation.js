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

async function main() {
  await tokens.start();
  await tokens.ready();
  const u0Headers = { ...baseHeaders, u0: tokens.token };

  const url = "https://market.fiintrade.vn/MarketInDepth/GetValuationSeriesV2?ComGroupCode=VNINDEX&TimeRange=OneYear&language=vi";
  try {
    const res = await client.fetch(url, { headers: u0Headers });
    if (res.ok) {
      const data = await res.json();
      const count = data.items ? data.items.length : (Array.isArray(data) ? data.length : 0);
      console.log(`[ValuationTest] count=${count}`);
      if (count > 0) {
        const item = data.items ? data.items[0] : (Array.isArray(data) ? data[0] : data);
        console.log(`Sample:`, JSON.stringify(item, null, 2));
      }
    }
  } catch (e) {
    console.error(e);
  }

  tokens.stop();
}

main().catch(console.error);
