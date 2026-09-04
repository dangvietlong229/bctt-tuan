'use strict';

/**
 * Hardened HTTP layer for the Fiintrade client.
 *
 * Goal: stay under rate limits and avoid IP/token bans by behaving like a
 * polite, well-mannered client instead of a tight loop.
 *
 * Strategy:
 *   1. Token-bucket pacing       -> caps sustained request rate.
 *   2. Randomized jitter         -> breaks the predictable "exactly 1000ms"
 *                                   signature that makes a bot trivial to flag.
 *   3. Respect Retry-After / 429 -> back off exactly as long as the server asks.
 *   4. Exponential backoff+jitter-> for transient 5xx / network errors.
 *   5. Circuit breaker           -> after N consecutive hard blocks, stop
 *                                   hammering (the single fastest way to a ban).
 */

// ---------------------------------------------------------------------------
// Tunables (override via env so you never have to edit code to slow down).
// ---------------------------------------------------------------------------
const CONFIG = {
  // Sustained pace: minimum gap between requests, plus random extra jitter.
  minDelayMs: numEnv('FIIN_MIN_DELAY_MS', 800),
  jitterMs: numEnv('FIIN_JITTER_MS', 1200), // actual gap = minDelay + rand(0..jitter)

  // Token bucket: at most `burst` requests, refilled at `refillPerSec`.
  burst: numEnv('FIIN_BURST', 5),
  refillPerSec: numEnv('FIIN_REFILL_PER_SEC', 1),

  // Retry policy for transient failures (5xx, network, 429 without Retry-After).
  maxRetries: numEnv('FIIN_MAX_RETRIES', 4),
  baseBackoffMs: numEnv('FIIN_BASE_BACKOFF_MS', 1000),
  maxBackoffMs: numEnv('FIIN_MAX_BACKOFF_MS', 30000),

  // Circuit breaker: consecutive 401/403/429 before we abort the whole run.
  maxConsecutiveBlocks: numEnv('FIIN_MAX_BLOCKS', 3),

  // Per-request hard timeout.
  requestTimeoutMs: numEnv('FIIN_TIMEOUT_MS', 15000),
};

function numEnv(name, def) {
  const v = Number(process.env[name]);
  return Number.isFinite(v) && v > 0 ? v : def;
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const rand = (max) => Math.floor(Math.random() * max);

// ---------------------------------------------------------------------------
// Token bucket — limits sustained throughput regardless of caller loop speed.
// ---------------------------------------------------------------------------
class TokenBucket {
  constructor(capacity, refillPerSec) {
    this.capacity = capacity;
    this.tokens = capacity;
    this.refillPerSec = refillPerSec;
    this.last = Date.now();
  }
  async take() {
    for (;;) {
      const now = Date.now();
      this.tokens = Math.min(
        this.capacity,
        this.tokens + ((now - this.last) / 1000) * this.refillPerSec
      );
      this.last = now;
      if (this.tokens >= 1) {
        this.tokens -= 1;
        return;
      }
      const needed = (1 - this.tokens) / this.refillPerSec;
      await sleep(Math.ceil(needed * 1000));
    }
  }
}

// ---------------------------------------------------------------------------
// Circuit breaker — once tripped, every call rejects immediately so a single
// bad token / soft ban doesn't turn into thousands of blocked requests.
// ---------------------------------------------------------------------------
class CircuitBreaker {
  constructor(threshold) {
    this.threshold = threshold;
    this.consecutive = 0;
    this.tripped = false;
  }
  recordSuccess() {
    this.consecutive = 0;
  }
  recordBlock() {
    this.consecutive += 1;
    if (this.consecutive >= this.threshold) this.tripped = true;
  }
  check() {
    if (this.tripped) {
      throw new Error(
        `Circuit breaker open after ${this.consecutive} consecutive blocks. ` +
          `Aborting to avoid an IP/token ban. Refresh credentials and retry later.`
      );
    }
  }
}

class FiinClient {
  constructor(config = {}) {
    this.cfg = { ...CONFIG, ...config };
    this.bucket = new TokenBucket(this.cfg.burst, this.cfg.refillPerSec);
    this.breaker = new CircuitBreaker(this.cfg.maxConsecutiveBlocks);
    this._lastRequestAt = 0;
  }

  /** Honor minDelay + jitter between consecutive requests (human-like pacing). */
  async _pace() {
    const gap = this.cfg.minDelayMs + rand(this.cfg.jitterMs);
    const wait = this._lastRequestAt + gap - Date.now();
    if (wait > 0) await sleep(wait);
    this._lastRequestAt = Date.now();
  }

  /** Parse Retry-After (seconds or HTTP-date) into ms. */
  static _retryAfterMs(res) {
    const h = res.headers.get('retry-after');
    if (!h) return null;
    const secs = Number(h);
    if (Number.isFinite(secs)) return secs * 1000;
    const date = Date.parse(h);
    return Number.isFinite(date) ? Math.max(0, date - Date.now()) : null;
  }

  _backoff(attempt) {
    const exp = Math.min(
      this.cfg.maxBackoffMs,
      this.cfg.baseBackoffMs * 2 ** attempt
    );
    return exp / 2 + rand(exp / 2); // full-ish jitter
  }

  /**
   * Rate-limited, retrying fetch.
   * Returns the final Response (caller reads body). Throws only on circuit
   * breaker / exhausted retries / network error after all attempts.
   *
   * @param {{onBlock?:(status:number)=>void}} [opts]
   */
  async fetch(url, init = {}, opts = {}) {
    this.breaker.check();
    await this.bucket.take();
    await this._pace();

    let attempt = 0;
    for (;;) {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), this.cfg.requestTimeoutMs);
      let res;
      try {
        const resolvedInit = typeof init === 'function' ? await init({ attempt }) : init;
        res = await fetch(url, { ...resolvedInit, signal: ctrl.signal });
      } catch (err) {
        clearTimeout(timer);
        if (attempt >= this.cfg.maxRetries) throw err;
        const wait = this._backoff(attempt++);
        console.warn(`  network error (${err.name}); retry in ${wait}ms`);
        await sleep(wait);
        continue;
      }
      clearTimeout(timer);

      // Rate limited -> obey Retry-After, else exponential backoff.
      if (res.status === 429) {
        this.breaker.recordBlock();
        opts.onBlock?.(429);
        const ra = FiinClient._retryAfterMs(res) ?? this._backoff(attempt);
        if (attempt >= this.cfg.maxRetries) return res;
        try { await res.body?.cancel(); } catch {}
        console.warn(`  429 rate-limited; backing off ${Math.round(ra)}ms`);
        attempt++;
        await sleep(ra);
        this.breaker.check();
        continue;
      }

      // Auth/forbidden -> token likely expired or soft ban. Count it, let the
      // caller decide whether to refresh; do NOT blindly retry the same token.
      if (res.status === 401 || res.status === 403) {
        this.breaker.recordBlock();
        opts.onBlock?.(res.status);
        return res;
      }

      // Transient server errors -> retry with backoff.
      if (res.status >= 500 && attempt < this.cfg.maxRetries) {
        try { await res.body?.cancel(); } catch {}
        const wait = this._backoff(attempt++);
        console.warn(`  ${res.status} server error; retry in ${Math.round(wait)}ms`);
        await sleep(wait);
        continue;
      }

      if (res.ok) this.breaker.recordSuccess();
      return res;
    }
  }
}

module.exports = { FiinClient, sleep, CONFIG };
