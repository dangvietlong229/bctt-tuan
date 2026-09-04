'use strict';

/**
 * build_catalog.js — host-aware endpoint extractor for fiintrade.vn.
 *
 * Service classes in the bundle look like:
 *    ...{key:"getForeign",value:function(e,t){return this.get("GetForeign",e,t)}}])
 *       ,t}(dn.a))(pn.a.Market.MoneyFlow.ServiceUrl)
 * i.e. a class body full of this.get("Method") calls, terminated by an
 * instantiation that injects a config reference ( ....MoneyFlow.ServiceUrl ).
 *
 * Config references resolve to base-URL literals defined elsewhere:
 *    MoneyFlow:{ServiceUrl:"".concat("https://market.fiintrade.vn","/MoneyFlow/")
 *    PriceDataUrl:"".concat("https://technical.fiintrade.vn","/PriceData/")
 *
 * Strategy:
 *   1. Build keyMap: config-property-name -> base URL.
 *   2. Split the source at each instantiation tail; the this.get/post calls in
 *      each segment belong to that segment's trailing config reference.
 *   3. Emit catalog.json: { verb, url, base, method } per endpoint.
 */

const fs = require('fs');
const path = require('path');

const FILES = ['2.9e4f9559.chunk.js', 'main.6abe6580.chunk.js'].map((f) =>
  path.join(__dirname, f)
);

// Config base-URL definitions ------------------------------------------------
// (a)  Name:{ServiceUrl:"".concat("https://host","/Ctrl/")
const DEF_SERVICE = /([A-Za-z0-9]+):\{ServiceUrl:"".concat\("(https:\/\/[a-z]+\.fiintrade\.vn)","(\/[A-Za-z0-9]+\/)"\)/g;
// (b)  XxxUrl:"".concat("https://host","/Ctrl/")
const DEF_NAMEDURL = /([A-Za-z0-9]+Url):"".concat\("(https:\/\/[a-z]+\.fiintrade\.vn)","(\/[A-Za-z0-9]+\/)"\)/g;

// Instantiation tail: ...}]),X}(BASE))(CONFREF)
const TAIL = /\}\]\),[a-zA-Z_$][\w$]*\}\([\w$.]+\)\)\(([\w$.]+)\)/g;

const CALL = /this\.(get|post|put|delete)\("([A-Za-z0-9_]+)"/g;

function buildKeyMap(src, keyMap) {
  let m;
  DEF_SERVICE.lastIndex = 0;
  while ((m = DEF_SERVICE.exec(src)) !== null) {
    // referenced as  ....<Name>.ServiceUrl  -> key "<Name>.ServiceUrl"
    keyMap[`${m[1]}.ServiceUrl`] = `${m[2]}${m[3]}`;
  }
  DEF_NAMEDURL.lastIndex = 0;
  while ((m = DEF_NAMEDURL.exec(src)) !== null) {
    keyMap[m[1]] = `${m[2]}${m[3]}`; // referenced as ....<XxxUrl>
  }
}

/** Resolve an instantiation config reference to a base URL. */
function resolveRef(ref, keyMap) {
  const parts = ref.split('.');
  const last = parts[parts.length - 1];
  if (last === 'ServiceUrl') {
    const name = parts[parts.length - 2];
    return keyMap[`${name}.ServiceUrl`];
  }
  if (/Url$/.test(last)) return keyMap[last];
  return undefined;
}

function main() {
  const keyMap = {};
  const sources = [];
  for (const f of FILES) {
    if (!fs.existsSync(f)) continue;
    const src = fs.readFileSync(f, 'utf8');
    sources.push(src);
    buildKeyMap(src, keyMap);
  }

  const catalog = {}; // base -> { method: verb }
  let unresolved = 0;

  for (const src of sources) {
    let m;
    let segStart = 0;
    TAIL.lastIndex = 0;
    while ((m = TAIL.exec(src)) !== null) {
      const ref = m[1];
      const segEnd = m.index + m[0].length;
      const base = resolveRef(ref, keyMap);
      const segment = src.slice(segStart, m.index);
      segStart = segEnd;
      if (!base) {
        // still advance; count only segments that had endpoint-ish calls
        if (/this\.(get|post)\("(Get|Download|Create|Update|Delete)/.test(segment)) unresolved++;
        continue;
      }
      if (!catalog[base]) catalog[base] = {};
      let c;
      CALL.lastIndex = 0;
      while ((c = CALL.exec(segment)) !== null) {
        catalog[base][c[2]] = c[1].toUpperCase();
      }
    }
  }

  const rows = [];
  for (const [base, methods] of Object.entries(catalog)) {
    for (const [method, verb] of Object.entries(methods)) {
      rows.push({ verb, url: `${base}${method}`, base, method });
    }
  }
  rows.sort((a, b) => a.url.localeCompare(b.url));
  fs.writeFileSync(path.join(__dirname, 'catalog.json'), JSON.stringify(rows, null, 2));

  const byCtrl = {};
  for (const r of rows) byCtrl[r.base.replace('https://', '').replace(/\/$/, '')] = (byCtrl[r.base.replace('https://', '').replace(/\/$/, '')] || 0) + 1;
  console.log(`Catalog: ${rows.length} endpoints across ${Object.keys(catalog).length} controllers (${unresolved} unresolved segments)\n`);
  for (const [ctrl, n] of Object.entries(byCtrl).sort()) console.log(`  ${String(n).padStart(3)}  ${ctrl}`);
  console.log('\nSaved catalog.json');
}

main();
