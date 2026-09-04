const { FiinClient } = require('./lib/http');
const { TokenProvider } = require('./lib/token');
const client = new FiinClient({ maxConsecutiveBlocks: 99, minDelayMs: 250, jitterMs: 300 });
const tokens = new TokenProvider();
const H = { accept:'application/json, text/plain, */*','accept-language':'vi',authorization:'Bearer',origin:'https://fiintrade.vn',referer:'https://fiintrade.vn/','user-agent':'M/5 Chrome/149' };
const M='https://market.fiintrade.vn', F='https://fundamental.fiintrade.vn', T='https://technical.fiintrade.vn';
const tl = Array.from({length:21},(_,i)=>`Timeline=${2006+i}_5`).join('&');
const u=[
  // corrected controllers (were 404 before because wrong controller)
  ['MarketInDepth/GetLiquiditySeries',  `${M}/MarketInDepth/GetLiquiditySeries?language=vi&ComGroupCode=ACB&TimeRange=OneDay`],
  ['MarketInDepth/GetValuationSeriesV2',`${M}/MarketInDepth/GetValuationSeriesV2?language=vi&Code=ACB&TimeRange=SixMonths&FromDate=&ToDate=`],
  ['MarketInDepth/GetMarketAnomaly',    `${M}/MarketInDepth/GetMarketAnomaly?language=vi&Code=ACB&TimeRange=FiveYears`],
  ['WatchList/GetTickerSeries',         `${M}/WatchList/GetTickerSeries?language=vi&OrganCode=ACB&TimeRange=OneDay`],
  // financial ratio with the real Timeline=YEAR_5 param shape
  ['FinancialAnalysis/GetFinancialRatioV2', `${F}/FinancialAnalysis/GetFinancialRatioV2?language=vi&Type=Company&OrganCode=ACB&${tl}`],
  // financial statements (were 401 on guest) - retest
  ['FinancialStatement/GetBalanceSheet',`${F}/FinancialStatement/GetBalanceSheet?language=vi&OrganCode=ACB`],
];
(async()=>{await tokens.start();await tokens.ready();console.log('Live token acquired.\n');
for(const [n,url] of u){const r=await client.fetch(url,{headers:{...H,u0:tokens.token}});const t=await r.text();console.log((r.status===200?'✓':'✗')+' ['+r.status+'] '+n.padEnd(40)+' '+(r.ok?t.length+'b':t.slice(0,45)));}
tokens.stop();})();
