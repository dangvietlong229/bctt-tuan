'use strict';

/**
 * Live `u0` token provider.
 *
 * The Fiintrade server rotates the guest token (`window._kc`) roughly every
 * 30s and pushes it over a SignalR WebSocket. Hardcoded tokens go stale fast
 * and produce 403s. This module opens the realtime hub, captures the current
 * token, and keeps the socket alive so callers always read a fresh value.
 */

const HUB_URL = 'https://realtime.fiintrade.vn/RealtimeHub';
const RECORD_SEPARATOR = String.fromCharCode(0x1e);

const browserHeaders = {
  accept: '*/*',
  'accept-language': 'vi,en-US;q=0.9,en;q=0.8,zh-CN;q=0.7,zh;q=0.6',
  origin: 'https://fiintrade.vn',
  referer: 'https://fiintrade.vn/',
  'user-agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
};

function looksLikeToken(v) {
  return typeof v === 'string' && v.length > 20 && v.endsWith('==');
}

class TokenProvider {
  constructor() {
    this.token = null;
    this.ws = null;
    this._waiters = [];
    this._changeWaiters = [];
  }

  /** Resolve only after the server publishes a token different from previous. */
  waitForChange(previous, timeoutMs = 15000) {
    if (this.token && this.token !== previous) return Promise.resolve(this.token);
    return new Promise((resolve, reject) => {
      const waiter = { previous, resolve, reject, timer: null };
      waiter.timer = setTimeout(() => {
        this._changeWaiters = this._changeWaiters.filter((item) => item !== waiter);
        reject(new Error('Timed out waiting for a refreshed token'));
      }, timeoutMs);
      this._changeWaiters.push(waiter);
    });
  }

  /** Resolve once a token is available (or immediately if we already have one). */
  ready(timeoutMs = 15000) {
    if (this.token) return Promise.resolve(this.token);
    return new Promise((resolve, reject) => {
      const t = setTimeout(
        () => reject(new Error('Timed out waiting for live token')),
        timeoutMs
      );
      this._waiters.push((tok) => {
        clearTimeout(t);
        resolve(tok);
      });
    });
  }

  _set(tok) {
    this.token = tok;
    while (this._waiters.length) this._waiters.shift()(tok);
    const remaining = [];
    for (const waiter of this._changeWaiters) {
      if (tok !== waiter.previous) {
        clearTimeout(waiter.timer);
        waiter.resolve(tok);
      } else {
        remaining.push(waiter);
      }
    }
    this._changeWaiters = remaining;
  }

  async start() {
    const negotiateRes = await fetch(`${HUB_URL}/negotiate`, {
      method: 'POST',
      headers: {
        ...browserHeaders,
        'content-type': 'text/plain;charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest',
      },
      body: '',
    });
    if (!negotiateRes.ok) {
      throw new Error(`Negotiate failed: ${negotiateRes.status}`);
    }
    const { connectionId } = await negotiateRes.json();
    const wsUrl = `wss://realtime.fiintrade.vn/RealtimeHub?id=${connectionId}`;
    const ws = new WebSocket(wsUrl, { headers: browserHeaders });
    this.ws = ws;

    ws.onopen = () => {
      ws.send(JSON.stringify({ protocol: 'json', version: 1 }) + RECORD_SEPARATOR);
    };

    ws.onmessage = (event) => {
      for (const raw of event.data.split(RECORD_SEPARATOR)) {
        if (!raw.trim()) continue;
        let msg;
        try {
          msg = JSON.parse(raw);
        } catch {
          continue;
        }
        if (msg.type === 6) {
          ws.send(JSON.stringify({ type: 6 }) + RECORD_SEPARATOR); // keep-alive
          continue;
        }
        if (Array.isArray(msg.arguments) && looksLikeToken(msg.arguments[0])) {
          this._set(msg.arguments[0]); // refresh on every push
        }
      }
    };

    ws.onerror = () => {};
    return this;
  }

  stop() {
    try {
      this.ws?.close();
    } catch {
      /* noop */
    }
    for (const waiter of this._changeWaiters.splice(0)) {
      clearTimeout(waiter.timer);
      waiter.reject(new Error('Token provider stopped while waiting for a refreshed token'));
    }
  }
}

module.exports = { TokenProvider };
