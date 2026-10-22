// The post list is written twice: in getDropdownData (Code.gs) and as POSTS in
// docs/dashboard.html, so the dropdown fills without waiting for the server.
// Same for the zone districts. This fails when one copy is edited and the
// other is not.
//   node scripts/tests/lists.test.js
const fs = require('fs');
const gs = fs.readFileSync('Code.gs', 'utf8');
const page = fs.readFileSync('docs/dashboard.html', 'utf8');

function arrayAfter(src, marker) {
  const a = src.indexOf(marker);
  if (a < 0) throw new Error('marker not found: ' + marker);
  const open = src.indexOf('[', a), close = src.indexOf('];', open);
  return Function('return ' + src.slice(open, close + 1))();
}
function objectAfter(src, marker) {
  const a = src.indexOf(marker);
  if (a < 0) throw new Error('marker not found: ' + marker);
  const open = src.indexOf('{', a), close = src.indexOf('};', open);
  return Function('return ' + src.slice(open, close + 1))();
}

const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };

const serverPosts = arrayAfter(gs, 'var stakeholders = [');
const pagePosts = arrayAfter(page, 'var POSTS = [');
ok('post list on the page matches getDropdownData in Code.gs',
   JSON.stringify(serverPosts) === JSON.stringify(pagePosts),
   'only in Code.gs: ' + serverPosts.filter(p => !pagePosts.includes(p)).join(', ') +
   ' | only on the page: ' + pagePosts.filter(p => !serverPosts.includes(p)).join(', '));
ok('post list ends with Other (the page opens the Other box on it)', pagePosts[pagePosts.length - 1] === 'Other');

const purposes = arrayAfter(page, 'var PURPOSES_DEFAULT = [');
ok('default purposes do not repeat Other (the page adds "Other (Specify)" itself)',
   !purposes.some(p => p.trim().toLowerCase() === 'other'));

const zonesGs = objectAfter(gs, 'var ZONE_DISTRICTS = {');
const zonesPage = objectAfter(page, 'var ZONE_DISTRICTS = {');
ok('zone districts on the page match Code.gs', JSON.stringify(zonesGs) === JSON.stringify(zonesPage));
ok('state-level extra districts match',
   JSON.stringify(arrayAfter(gs, 'var STATE_EXTRA_DISTRICTS')) === JSON.stringify(arrayAfter(page, 'var STATE_EXTRA_DISTRICTS')));

console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
