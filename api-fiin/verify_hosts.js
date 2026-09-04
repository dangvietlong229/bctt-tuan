'use strict';
const fs = require('fs');
const src =
  fs.readFileSync('main.6abe6580.chunk.js', 'utf8') +
  fs.readFileSync('2.9e4f9559.chunk.js', 'utf8');

// gated controllers -> confirm exactly one authoritative host binding
const ctrls = [
  'FinancialStatement', 'FinancialAnalysis', 'Calendar', 'Strategy',
  'TAStrategy', 'Rankings', 'Screener', 'Valuation',
  'TechnicalAnalysisSignals', 'PersonalAlert', 'PriceData', 'Tracking',
];

for (const ctrl of ctrls) {
  const re = new RegExp(
    'concat\\("(https://[a-z]+\\.fiintrade\\.vn)","(/' + ctrl + '/)"\\)',
    'g'
  );
  const hosts = new Set();
  let m;
  while ((m = re.exec(src))) hosts.add(m[1]);
  const list = [...hosts];
  const tag = list.length === 1 ? 'UNIQUE' : list.length === 0 ? 'NONE' : 'MULTI';
  console.log(`${ctrl.padEnd(24)} [${tag}] ${list.join(', ') || '(no binding)'}`);
}
