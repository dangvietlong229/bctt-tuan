const fs = require('fs');

const files = [
  'c:/Users/admin/Documents/api-fiin/2.9e4f9559.chunk.js',
  'c:/Users/admin/Documents/api-fiin/main.6abe6580.chunk.js'
];

const methods = new Set();
const fullUrls = new Set();

// Scan for .get("...") and .post("...")
const getPostRegex = /\.(get|post|put|delete)\(\s*['"`]([a-zA-Z0-9_\-\/]+)['"`]/g;

// Scan for absolute URLs as well
const urlRegex = /https?:\/\/[a-zA-Z0-9.-]+\.fiintrade\.vn\/[a-zA-Z0-9.\/_-]*/g;

for (const file of files) {
  if (!fs.existsSync(file)) continue;
  const content = fs.readFileSync(file, 'utf8');

  // Extract absolute URLs
  const urls = content.match(urlRegex) || [];
  urls.forEach(u => fullUrls.add(u));

  // Extract HTTP method calls
  let match;
  while ((match = getPostRegex.exec(content)) !== null) {
    methods.add(`${match[1].toUpperCase()}: ${match[2]}`);
  }
}

console.log('HTTP calls found:');
console.log(JSON.stringify(Array.from(methods).sort(), null, 2));

console.log('Absolute URLs found:');
console.log(JSON.stringify(Array.from(fullUrls).sort(), null, 2));
