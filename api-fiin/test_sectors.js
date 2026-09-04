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
      console.log(`[SectorTest] ${name}: count=${count}, keys=${Object.keys(data)}`);
      if (count > 0) {
        console.log(`Sample item from ${name}:`, JSON.stringify(data.items ? data.items[0] : (Array.isArray(data) ? data[0] : data)).slice(0, 300));
      }
    } else {
      console.log(`[SectorTest] ${name} failed: ${res.status}`);
    }
  } catch (e) {
    console.log(`[SectorTest] ${name} error: ${e.message}`);
  }
}

async function main() {
  await tokens.start();
  await tokens.ready();
  const u0Headers = { ...baseHeaders, u0: tokens.token };

  const endpoints = [
    { name: "GetLatestSector", url: "https://market.fiintrade.vn/SectorIndepth/GetLatestSector?language=vi" },
    { name: "GetSectorProportionSummary", url: "https://market.fiintrade.vn/SectorIndepth/GetSectorProportionSummary?language=vi&ComGroupCode=VNINDEX" },
    { name: "GetSectorRatio", url: "https://market.fiintrade.vn/SectorIndepth/GetSectorRatio?language=vi&ComGroupCode=VNINDEX" },
    { name: "GetSectorSeries", url: "https://market.fiintrade.vn/SectorIndepth/GetSectorSeries?language=vi&ComGroupCode=VNINDEX&TimeRange=OneDay" },
    { name: "GetSectorPerformance_Bare", url: "https://market.fiintrade.vn/SectorIndepth/GetSectorPerformance" },
    { name: "GetSectorFMI", url: "https://market.fiintrade.vn/SectorIndepth/GetSectorFMI?language=vi" },
  ];

  for (const ep of endpoints) {
    await testUrl(ep.name, ep.url, u0Headers);
  }

  tokens.stop();
}

main().catch(console.error);
