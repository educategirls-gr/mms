// Syntax-check every inline <script> in a page:  node scripts/tests/jscheck.js docs/dashboard.html
const fs = require('fs');
const vm = require('vm');
const file = process.argv[2];
const html = fs.readFileSync(file, 'utf8');
const re = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
let m, n = 0, bad = 0;
while ((m = re.exec(html))) {
  if (/\bsrc\s*=/.test(m[1])) continue;
  n++;
  const before = html.slice(0, m.index).split('\n').length;
  try { new vm.Script(m[2], { filename: file + ' block' + n }); }
  catch (e) { bad++; console.log('FAIL block ' + n + ' (starts near line ' + before + '): ' + e.message); }
}
console.log(file + ': ' + n + ' inline script block(s), ' + (bad ? bad + ' with syntax errors' : 'all parse clean'));
process.exit(bad ? 1 : 0);
