const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const client = new FiinClient({ maxConsecutiveBlocks: 99 });
const tokens = new TokenProvider();
const H = { accept:'application/json, text/plain, */*','accept-language':'vi', authorization:'Bearer', origin:'https://fiintrade.vn', referer:'https://fiintrade.vn/', 'user-agent':'Mozilla/5.0 Chrome/149.0.0.0' };
const T = 'https://technical.fiintrade.vn';
const M = 'https://market.fiintrade.vn';

// "Thống kê giá" = Price Data panel + its tabs
const urls = [
  ['PriceData/GetPriceData (Daily)',      `${T}/PriceData/GetPriceData?Code=ACB&OrganCode=ACB&Frequency=Daily&Page=1&PageSize=20&language=vi`],
  ['PriceData/GetPriceData (From/To)',    `${T}/PriceData/GetPriceData?Code=ACB&Frequency=Daily&From=2025-01-01T00:00:00.000Z&To=2026-06-12T00:00:00.000Z&language=vi`],
  ['TimeAndSales/GetTimeAndSalesStatChart',`${T}/TimeAndSales/GetTimeAndSalesStatChart?Code=ACB&offset=0&page=1&language=vi`],
  ['MoneyFlow GetStatisticInvestor',      `${M}/MoneyFlow/GetStatisticInvestor?Code=ACB&ComGroupCode=ACB&language=vi`],
  ['MoneyFlow GetForeign (by ticker)',    `${M}/MoneyFlow/GetForeign?ComGroupCode=ACB&language=vi`],
  ['MoneyFlow GetProprietaryV2',          `${M}/MoneyFlow/GetProprietaryV2?ComGroupCode=ACB&language=vi`],
  ['PriceData GetPutThrough',             `${T}/PriceData/GetPutThrough?Code=ACB&language=vi`],
  ['PriceData GetClosing',                `${T}/PriceData/GetClosing?Code=ACB&language=vi`],
];
(async()=>{
  await tokens.start(); await tokens.ready();
  console.log('Live token acquired.\n');
  for (const [n,u] of urls){
    const res = await client.fetch(u,{headers:{...H,u0:tokens.token}});
    const t = await res.text();
    let head = res.ok ? t.slice(0,70).replace(/\s+/g,' ') : t.slice(0,40);
    console.log(`${res.status===200?'✓':'✗'} [${res.status}] ${n}\n      ${head}`);
  }
  tokens.stop();
})();
