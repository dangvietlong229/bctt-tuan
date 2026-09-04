const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const client = new FiinClient({ maxConsecutiveBlocks: 99 }); // probe-all, don't abort on 401s
const tokens = new TokenProvider();
const H = { accept:'application/json, text/plain, */*','accept-language':'vi', authorization:'Bearer', origin:'https://fiintrade.vn', referer:'https://fiintrade.vn/', 'user-agent':'Mozilla/5.0 Chrome/149.0.0.0' };

const F = 'https://fundamental.fiintrade.vn';
const urls = [
  // EarningCorner family (the one you captured)
  ['EarningCorner/ProfitDisclosureOverview', `${F}/EarningCorner/ProfitDisclosureOverview?yearReport=2022&lengthReport=1&language=vi`],
  ['EarningCorner/ProfitGrowthOverview',      `${F}/EarningCorner/ProfitGrowthOverview?yearReport=2022&lengthReport=1&language=vi`],
  ['EarningCorner/ProfitGrowthDetail',        `${F}/EarningCorner/ProfitGrowthDetail?yearReport=2022&lengthReport=1&language=vi`],
  ['EarningCorner/ProfitGrowthChartOverview', `${F}/EarningCorner/ProfitGrowthChartOverview?yearReport=2022&lengthReport=1&language=vi`],
  // financial statements (controller guess from SSI dump)
  ['FinancialStatement/GetBalanceSheet',      `${F}/FinancialStatement/GetBalanceSheet?language=vi&OrganCode=ACB`],
  ['FinancialStatement/GetIncomeStatement',   `${F}/FinancialStatement/GetIncomeStatement?language=vi&OrganCode=ACB`],
  ['Snapshot/GetCompanyScore',                `${F}/Snapshot/GetCompanyScore?language=vi&OrganCode=ACB`],
];
(async()=>{
  await tokens.start(); await tokens.ready();
  console.log('Live token acquired.\n');
  for (const [n,u] of urls){
    // refresh u0 each call from the live socket so we never use a stale token
    const res = await client.fetch(u,{headers:{...H,u0:tokens.token}});
    const t = await res.text();
    console.log(`${res.status===200?'✓':'✗'} [${res.status}] ${n}  ${res.ok? t.length+'b' : t.slice(0,40)}`);
  }
  tokens.stop();
})();
