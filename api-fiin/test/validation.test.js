'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { normalizeTicker, parseBoundedInteger, validatePullData } = require('../lib/validation');
const { writeJsonAtomic } = require('../lib/io');
const { TokenProvider } = require('../lib/token');

test('ticker and integer inputs are bounded', () => {
  assert.equal(normalizeTicker(' tcb '), 'TCB');
  assert.throws(() => normalizeTicker('../TCB'));
  assert.equal(parseBoundedInteger('3', 2, { min: 1, max: 10 }), 3);
  assert.throws(() => parseBoundedInteger('100', 2, { min: 1, max: 10 }));
});

test('validation rejects partial or identical sector payloads', () => {
  assert.throws(() => validatePullData({ sectors: {} }, [], { asOf: new Date('2026-08-21') }));
  const same = [{ sectors: [{ name: 'Bank' }] }];
  assert.throws(() => validatePullData({ sectors: { OneDay: same, OneWeek: same, OneMonth: same } }, []));
});

test('atomic writer leaves valid JSON', () => {
  const tempRoot = path.join(__dirname, '.tmp');
  fs.mkdirSync(tempRoot, { recursive: true });
  const directory = fs.mkdtempSync(path.join(tempRoot, 'bctt-api-'));
  const output = path.join(directory, 'result.json');
  try {
    writeJsonAtomic(output, { ok: true });
    assert.deepEqual(JSON.parse(fs.readFileSync(output, 'utf8')), { ok: true });
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
    try { fs.rmdirSync(tempRoot); } catch {}
  }
});

test('token refresh waits for a different token', async () => {
  const provider = new TokenProvider();
  provider._set('old-token-value-that-is-long-enough==');
  const changed = provider.waitForChange(provider.token, 1000);
  provider._set('new-token-value-that-is-long-enough==');
  assert.equal(await changed, 'new-token-value-that-is-long-enough==');
  provider.stop();
});
