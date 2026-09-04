# fiintrade.vn Guest-Token Coverage Map

Controller-level reachability with the anonymous `u0` token (no login).
Derived from `catalog.json` (216 endpoints across 33 controllers / 7 hosts)
+ live sampling probes.

## ✅ Guest-FREE controllers (reachable, no account)

| Host | Controller | Sample | Notes |
|------|-----------|--------|-------|
| core | Master | GetListOrganization | reference/master data |
| core | UserSetting (read) | GetWorkspace | guest workspaces only |
| core | UtilFeature | GetHotNews | |
| technical | PriceData | GetLatestPrice / GetVWAP | (GetPriceData table = gated) |
| technical | PriceDepth | GetPriceDepth | |
| technical | TimeAndSales | GetTimeAndSales | tick data |
| technical | TradingView | GetStockChartData, GetStockEvents, GetIndicators | OHLC, events, fundamentals (see ⭐ note) |
| market | MarketInDepth | GetProspect, **GetValuationSeriesV2** (the valuation panel), GetLiquiditySeries, GetMarketAnomaly | 12–43 KB |
| market | MoneyFlow | GetContribution, GetForeign, GetProprietaryV2, GetStatisticInvestor | money-flow trend |
| market | TopMover | GetTopGainers, GetTopLosers | (ForeignTrading/Breakout = gated) |
| market | HeatMap | GetHeatMap | |
| market | WatchList | GetTickerSeries | 138 KB |
| market | SectorIndepth | GetSectorPerformance | reachable |
| market | BUSD | GetBUSD | reachable |
| fundamental | EarningCorner | ProfitDisclosure/ProfitGrowth* | profit data |
| fundamental | Snapshot | GetCompanyScore, GetSnapshot | 8.8 KB |
| fundamental | Ownership | GetOwnership | 26 KB — shareholders/BoD |
| news | News | GetMostRecent | reachable |

## ⭐ Free path to fundamental data (TradingView/GetIndicators)

`technical/TradingView/GetIndicators?OganCode=<T>&Code=<IND>&Frequency=Daily&From=..&To=..&Type=Stock`
is **guest-free** and returns `{tradingDate,value}` series — and it serves not
just ratios but **balance-sheet aggregates and P&L line items**, partially
routing around the gated `FinancialStatement`/`FinancialAnalysis`.

CRITICAL: the symbol param is the misspelled **`OganCode`** (that is what the
backend expects). The correctly-spelled `OrganCode` returns 200 with all values
blanked to 0 — a silent soft-gate that looks like data but isn't.

Verified working codes (TCB): PE, PB, PS, DividendYield, ROE, ROA,
NetProfitMargin, GrossProfitMargin, NetProfit, GrossProfit, NetIncome,
NetRevenue, TotalAsset, Equity, TotalLiabilities, CurrentAsset,
CurrentLiabilities, TotalLoans, TotalDeposit. (EBITMargin/OI sparse for banks.)

Pull with: `node pull_indicators.js <CODE>` -> merged table in responses/<CODE>/Indicators.json.
Still does NOT give the full itemized statement (every line) — but the major
aggregates are recoverable for free.

## 🔒 GATED controllers (401 — need login/subscription)

| Host | Controller | Why |
|------|-----------|-----|
| fundamental | FinancialStatement | GetBalanceSheet/IncomeStatement/CashFlow |
| fundamental | FinancialAnalysis | GetFinancialRatioV2 |
| market | Calendar | corporate-events calendar |
| strategy | Strategy / TAStrategy / Rankings | **entire strategy host gated** |
| tools | Screener | GetScreenerParameters/Items |
| tools | Valuation | GetValuation — **separate** standalone tool, NOT the panel (panel = MarketInDepth/GetValuationSeriesV2, which is FREE) |
| technical | **TechnicalAnalysisSignals** (config name "Deceptive") | order-flow: GetAggressive/Pressing/Cancelled/Closing/CEFLAbnormality/PriceVolumeAnalysis — all 401 |
| (all hosts) | Download/* and Export/* | xlsx exports — login-only |
| core | UserSetting (personal) | watchlists/screeners of a user |
| tools | PersonalAlert | personal alerts |

## Request-body schemas (from captures)

`tools/Screener/GetScreenerItems` (POST, gated):
```json
{"comGroupCode":"All|HNXIndex|...","icbCode":"All|1000|...",
 "parameters":[{"name":"...","code":"AverageVolume2Week","type":"Range",
   "selectedValue":[min,max],"valueRange":[min,max],"unit":"ThousandUnit"}],
 "page":1,"pageSize":30,"OrderBy":"StockScreenerItem.Ticker","Direction":"ASC"}
```
Filter `code`/`unit` vocabulary comes from `Screener/GetScreenerParameters` (gated).

`fundamental/FinancialAnalysis/GetFinancialRatioV2` (gated) takes `Type=Company`
(the stock) OR `Type=Icb` (industry peer benchmark), with repeated
`Timeline=YEAR_5` params (2006_5 … 2026_5).

## Summary
- **216** endpoints mapped to host+controller (`catalog.json`) — incl. the
  late-added `TradingView` (free) and `TechnicalAnalysisSignals` (gated) controllers
  that the extractor first missed (bare config keys `Charting`/`Deceptive`).
- **~18 controllers guest-FREE**, spanning price, money-flow, sector, ownership,
  snapshot, news, top-movers, valuation-series, market-depth, and TradingView
  (chart/events/indicators).
- Note: a `200` is not proof of data — `GetIndicators` returns a zeroed series
  unless the symbol param is the backend's misspelled `OganCode`.
- **GATED:** financial statements, the entire `strategy.*` host, screeners,
  calendar, valuation, all Download/Export, and all personal/mutation endpoints.
- Genuinely out of guest reach without a fiintrade.vn account: financial
  statements, strategy/ranking screens, screener, calendar, file exports.
