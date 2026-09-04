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
      console.log(`[InvestorTest] ${name}: count=${count}, keys=${Object.keys(data)}`);
      if (count > 0) {
        const item = data.items ? data.items[0] : (Array.isArray(data) ? data[0] : data);
        console.log(`Sample:`, JSON.stringify(item, null, 2).slice(0, 800));
      } else {
        console.log(`Payload Sample:`, JSON.stringify(data, null, 2).slice(0, 800));
      }
    } else {
      console.log(`[InvestorTest] ${name} failed: ${res.status}`);
    }
  } catch (e) {
    console.log(`[InvestorTest] ${name} error: ${e.message}`);
  }
}

async function main() {
  await tokens.start();
  await tokens.ready();
  const u0Headers = { ...baseHeaders, u0: tokens.token };

  const endpoints = [
    { name: "GetStatisticInvestor_TCB", url: "https://market.fiintrade.vn/MoneyFlow/GetStatisticInvestor?ComGroupCode=TCB&language=vi" },
    { name: "GetStatisticInvestorChart_TCB", url: "https://market.fiintrade.vn/MoneyFlow/GetStatisticInvestorChart?ComGroupCode=TCB&TimeRange=OneYear&language=vi" },
    { name: "GetStatisticInvestorChart_VNINDEX", url: "https://market.fiintrade.vn/MoneyFlow/GetStatisticInvestorChart?ComGroupCode=VNINDEX&TimeRange=OneYear&language=vi" },
  ];

  for (const ep of endpoints) {
    await testUrl(ep.name, ep.url, u0Headers);
  }

  tokens.stop();
}

main().catch(console.error);
