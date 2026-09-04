# Fiintrade API Scraper & Realtime Handshake Client

Maps the Fiintrade (`*.fiintrade.vn`) API surface and provides a rate-limited
Node.js client that authenticates as an anonymous guest by replicating the
SignalR WebSocket handshake, then pulls data from the guest-accessible endpoints.

The full endpoint map lives in [`catalog.json`](catalog.json) (216 endpoints
across 33 controllers / 7 hosts) and the free-vs-gated classification in
[`COVERAGE.md`](COVERAGE.md).

---

## Authentication (`u0` header)

REST requests carry a header named `"u" + userId` — `u0` for guests.
- Its value is a rotating token (`window._kc`) the server pushes over a SignalR
  WebSocket, changing roughly every **30 seconds**.
- An expired/missing token gives `403`; an endpoint your account tier can't see
  gives `401`.
- `lib/token.js` opens the realtime hub, captures that token, answers keep-alive
  pings, and refreshes it continuously so long runs never use a stale token.

Anything beyond guest data (financial statements, screener, strategy host, file
exports, personal/watchlist endpoints) requires a logged-in fiintrade.vn account
(`Authorization: Bearer <jwt>` + `u<userId>`) — see `COVERAGE.md`.

---

## Directory structure

**Core library**
- `lib/http.js` — `FiinClient`: rate-limited, retrying, circuit-broken fetch.
- `lib/token.js` — `TokenProvider`: live `u0` token via the SignalR hub.

**Pullers**
- `retrieve_endpoints.js` — batch-pull the seed endpoint list (live token).
- `pull_ticker.js <CODE> [years]` — full "Thống kê giá" set for one symbol
  (price, VWAP, depth, time&sales, chart history, events). Auto-resolves
  `Type`/`OrganCode` from master data in `responses/`.
- `pull_indicators.js <CODE> [indicator...]` — fundamental ratios & balance-sheet
  aggregates as time-series via `TradingView/GetIndicators` (see note below).
- `signalr_pull.js` — original one-shot handshake demo.

**Mapping / analysis**
- `build_catalog.js` — parses the bundles into `catalog.json` (host→controller→method).
- `verify_hosts.js` — confirms each controller's single authoritative host.
- `probe_*.js` — polite sampling probes used to classify free vs gated.
- `2.9e4f9559.chunk.js` / `main.6abe6580.chunk.js` — client bundles (reverse-eng source).

**Output**
- `catalog.json` — complete endpoint map.
- `COVERAGE.md` — guest free/gated map per controller.
- `responses/` — saved JSON payloads (and per-ticker subfolders).

---

## Usage

Requires **Node.js v21+** (uses native global `fetch` and `WebSocket`; no npm deps).

```bash
node retrieve_endpoints.js          # batch-pull the seed endpoints
node pull_ticker.js TCB 2           # full price/T&S/chart set for TCB, 2y history
node pull_indicators.js TCB         # ROE/PE/PB/TotalAsset/... time-series for TCB
node build_catalog.js               # rebuild catalog.json from the bundles
node verify_hosts.js                # confirm controller→host bindings
```

---

## Gotchas worth knowing

- **`TradingView/GetIndicators` needs the misspelled `OganCode` param.** The
  backend literally expects `OganCode` (not `OrganCode`); the correctly-spelled
  version returns `200` with every value **blanked to 0**. This endpoint is the
  free path to ROE/PE/PB/margins and balance-sheet aggregates (TotalAsset,
  Equity, Liabilities, Loans, Deposits) — partially routing around the gated
  `FinancialStatement`/`FinancialAnalysis`.
- **A `200` is not proof of data.** Some endpoints soft-gate by returning a valid
  shape with zeroed/empty values.
- **Gating is per-controller (sometimes per-method), not per-host.** The same host
  serves free and gated controllers side by side (e.g. on `fundamental`,
  `EarningCorner` is free but `FinancialStatement` is gated; on `technical`,
  `PriceData/GetLatestPrice` is free but `PriceData/GetPriceData` is gated).
- **For individual tickers** `GetStockChartData` needs `Type=Stock` (not `Index`),
  and history is best pulled in 1-year windows.

---

## Rate-limit & anti-ban layer (`lib/http.js`)

All REST traffic goes through `FiinClient`, built to stay under limits and avoid
IP/token bans:

- **Token-bucket pacing** caps sustained throughput (default burst 5, refill 1/s).
- **Randomized jitter** between requests (`800ms + rand(0..1200ms)`) so traffic
  lacks the tell-tale "exactly 1000ms" bot signature.
- **`429` handling** obeys the server's `Retry-After` header exactly.
- **Exponential backoff + jitter** for transient `5xx` / network errors.
- **Circuit breaker** aborts after 3 consecutive `401/403/429` blocks — the single
  most important defense, since hammering a soft-blocked endpoint is what escalates
  to a hard ban.

### Tuning (env vars — no code edits needed)

| Var | Default | Meaning |
|---|---|---|
| `FIIN_MIN_DELAY_MS` | 800 | Minimum gap between requests |
| `FIIN_JITTER_MS` | 1200 | Random extra gap added on top |
| `FIIN_BURST` | 5 | Token-bucket capacity |
| `FIIN_REFILL_PER_SEC` | 1 | Tokens refilled per second |
| `FIIN_MAX_RETRIES` | 4 | Retries for transient failures |
| `FIIN_MAX_BLOCKS` | 3 | Consecutive blocks before circuit opens |
| `FIIN_TIMEOUT_MS` | 15000 | Per-request hard timeout |

Go slower (safer) with e.g. `FIIN_MIN_DELAY_MS=2000 FIIN_BURST=2 node retrieve_endpoints.js`.
