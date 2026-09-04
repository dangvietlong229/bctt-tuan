'use strict';

const SAFE_TICKER = /^[A-Z0-9]{2,12}$/;

function normalizeTicker(value, label = 'ticker') {
  const ticker = String(value || '').trim().toUpperCase();
  if (!SAFE_TICKER.test(ticker)) {
    throw new Error(`Invalid ${label}: expected 2-12 ASCII letters or digits`);
  }
  return ticker;
}

function parseBoundedInteger(value, fallback, { min = 1, max = 100, label = 'value' } = {}) {
  if (value === undefined || value === null || value === '') return fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < min || parsed > max) {
    throw new Error(`${label} must be an integer from ${min} to ${max}`);
  }
  return parsed;
}

function parseAsOf(value) {
  if (!value) return new Date();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error('--as-of must use YYYY-MM-DD');
  }
  const date = new Date(`${value}T23:59:59.999Z`);
  if (!Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== value) {
    throw new Error(`Invalid --as-of date: ${value}`);
  }
  return date;
}

function latestDate(items) {
  let latest = null;
  for (const item of items || []) {
    const raw = item?.tradingDate;
    if (!raw) continue;
    const date = new Date(raw);
    if (!Number.isFinite(date.getTime())) continue;
    if (!latest || date > latest) latest = date;
  }
  return latest;
}

function assertFresh(items, label, asOf, maxAgeDays) {
  const latest = latestDate(items);
  if (!latest) throw new Error(`${label} has no valid tradingDate`);
  const ageDays = (asOf.getTime() - latest.getTime()) / 86400000;
  if (ageDays > maxAgeDays) {
    throw new Error(
      `${label} is stale: latest ${latest.toISOString().slice(0, 10)}, ` +
        `as-of ${asOf.toISOString().slice(0, 10)}`
    );
  }
  if (ageDays < -1) {
    throw new Error(`${label} contains a trading date after the requested as-of date`);
  }
}

function payloadItems(payload) {
  return Array.isArray(payload) ? payload : payload?.items;
}

function validatePullData(rawData, tickers, options = {}) {
  const asOf = options.asOf || new Date();
  const maxAgeDays = options.maxAgeDays ?? 7;
  const ranges = ['OneDay', 'OneWeek', 'OneMonth'];
  const signatures = [];

  for (const range of ranges) {
    const rows = rawData?.sectors?.[range];
    if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(rows[0]?.sectors)) {
      throw new Error(`Sector ${range} data is incomplete`);
    }
    if (rows[0].sectors.length === 0) throw new Error(`Sector ${range} data is empty`);
    signatures.push(JSON.stringify(rows));
  }
  if (new Set(signatures).size === 1) {
    throw new Error('Sector OneDay/OneWeek/OneMonth payloads are identical');
  }

  for (const index of ['VNINDEX', 'VN30']) {
    const rows = rawData?.indices?.[index];
    if (!Array.isArray(rows) || rows.length === 0) {
      throw new Error(`${index} index series is incomplete`);
    }
    assertFresh(rows, index, asOf, maxAgeDays);
  }

  if (!Array.isArray(rawData.latestIndices) || rawData.latestIndices.length === 0) {
    throw new Error('Latest-indices data is incomplete');
  }

  const requiredParts = ['priceDepth', 'history', 'pe', 'pb'];
  for (const ticker of tickers) {
    const record = rawData?.tickers?.[ticker];
    if (!record) throw new Error(`${ticker} data is missing`);
    for (const part of requiredParts) {
      const items = payloadItems(record[part]);
      if (!Array.isArray(items) || items.length === 0) {
        throw new Error(`${ticker}.${part} is incomplete`);
      }
    }
    assertFresh(payloadItems(record.history), `${ticker}.history`, asOf, maxAgeDays);
  }
  return true;
}

module.exports = {
  SAFE_TICKER,
  normalizeTicker,
  parseBoundedInteger,
  parseAsOf,
  latestDate,
  assertFresh,
  validatePullData,
};
