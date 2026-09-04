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

  const url = "https://fundamental.fiintrade.vn/Snapshot/GetSnapshot?OrganCode=AAA&language=vi";
  try {
    const res = await client.fetch(url, { headers: u0Headers });
    if (res.ok) {
      const data = await res.json();
      console.log(JSON.stringify(data, null, 2));
    }
  } catch (e) {
    console.error(e);
  }

  tokens.stop();
}

main().catch(console.error);
