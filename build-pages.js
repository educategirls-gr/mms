// Generates districtreports.html + teamperformance.html from docs/report.html
// (single source of truth). Run after ANY edit to report.html:
//   node build-pages.js
const fs = require('fs');
const SRC = 'docs/report.html';
const src = fs.readFileSync(SRC, 'utf8');
const pages = {
  'docs/districtreports.html': 'reports',
  'docs/teamperformance.html': 'perf'
};
Object.keys(pages).forEach(function(out) {
  const view = pages[out];
  const content = src.replace(
    /var PORTAL_VIEW = '[^']*'; \/\* BUILD:VIEW \*\//,
    "var PORTAL_VIEW = '" + view + "'; /* BUILD:VIEW */"
  );
  if (content.indexOf("var PORTAL_VIEW = '" + view + "'") === -1) {
    console.error('❌ PORTAL_VIEW marker not found — aborting for ' + out);
    process.exit(1);
  }
  fs.writeFileSync(out, content);
  console.log('✅ Generated ' + out + '  (view=' + view + ')');
});
