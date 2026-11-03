// Commitments to check (3 Nov 2026): the REAL updateCommitment out of Code.gs
// on a fake Conducted sheet, and the REAL page helpers out of dashboard.html.
//   node scripts/tests/commitment-followup.test.js
const fs = require('fs');
const gs = fs.readFileSync('Code.gs', 'utf8');
const page = fs.readFileSync('docs/dashboard.html', 'utf8');
function fnSrc(src, name) {
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n}', a) + 2);
}
function pageFn(name) {
  const a = page.indexOf('  function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return page.slice(a, page.indexOf('\n  }', a) + 4);
}
const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };

// ── a fake Conducted sheet: rows of 34 columns (AG/AH exist, AI/AJ not yet) ──
const OFFICER = 'ravi.kumar@educategirls.ngo';
function makeSheet() {
  const rows = [new Array(34).fill('h')];
  rows[0][32] = 'Official Will Do'; rows[0][33] = 'Commitment Due';
  function r(id, email, what, due) { const x = new Array(34).fill(''); x[0] = id; x[1] = 'SITAPUR'; x[4] = email; x[32] = what; x[33] = due; return x; }
  rows.push(r('MTG-1', OFFICER, 'Share the village-wise list', '2026-11-02'));
  rows.push(r('MTG-2', 'someone.else@educategirls.ngo', 'Issue a letter', '2026-11-01'));
  rows.push(r('MTG-3', OFFICER, '', ''));
  const sh = { rows, maxCols: 34, log: [], fmt: {},
    getLastRow: () => rows.length,
    getMaxColumns: () => sh.maxCols,
    insertColumnsAfter: (a, n) => { sh.log.push('insert ' + a + '+' + n); sh.maxCols += n; rows.forEach(x => { while (x.length < sh.maxCols) x.push(''); }); },
    getRange: (row, col, nr, nc) => {
      nr = nr || 1; nc = nc || 1;
      if (col + nc - 1 > sh.maxCols) throw new Error('out of bounds: column ' + (col + nc - 1) + ' of ' + sh.maxCols);
      const rg = {
        getValues: () => rows.slice(row - 1, row - 1 + nr).map(x => x.slice(col - 1, col - 1 + nc).map(v => v === undefined ? '' : v)),
        getValue: () => { const v = rows[row - 1][col - 1]; return v === undefined ? '' : v; },
        setValue: (v) => { sh.log.push('set ' + row + ',' + col); rows[row - 1][col - 1] = v; return rg; },
        setNumberFormat: (f) => { sh.fmt[row + ',' + col] = f; return rg; }
      };
      return rg;
    } };
  return sh;
}
let sheet, invalidated;
const E = {
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }), flush: () => {} },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  Session: { getScriptTimeZone: () => 'Asia/Kolkata' },
  Utilities: { formatDate: (d, tz, f) => '2026-11-03' },   // "today" for the test
  SPREADSHEET_ID: 'x', CONDUCTED_SHEET: 'Conducted Meetings', SHEET_BUSY_MSG: 'busy',
  sheetBreakerTripped_: () => false,
  invalidateUser: (email, district) => { invalidated.push(email + '|' + district); }
};
const consts = gs.match(/var COL_CON_WILLDO = \d+;[\s\S]*?var COL_CON_DUE\s*= \d+;/)[0] + '\n' +
               gs.match(/var COL_CON_CSTATUS = \d+;[\s\S]*?var COMMIT_ANSWERS\s*= [^\n]+/)[0];
const body = [consts, fnSrc(gs, 'withScriptLock_'), fnSrc(gs, 'dayKey_'), fnSrc(gs, 'fmtDateVal'), fnSrc(gs, 'updateCommitment')].join('\n');
const names = Object.keys(E);
const S = new Function(...names, body + '\nreturn { updateCommitment, COL_CON_CSTATUS, COL_CON_CLOG };')(...names.map(k => E[k]));
const me = { email: 'Ravi.Kumar@educategirls.ngo' };
const reset = () => { sheet = makeSheet(); invalidated = []; };

ok('AI and AJ, right after AH (34)', S.COL_CON_CSTATUS === 35 && S.COL_CON_CLOG === 36);
reset();
let res = S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Done' });
ok('Done saved for the officer who filed it (email case ignored)', res.success && !res.already && sheet.rows[1][34] === 'Done', JSON.stringify(res));
ok('the two columns are added at the far right, headers written', sheet.log[0] === 'insert 34+2' && sheet.rows[0][34] === 'Commitment Status' && sheet.rows[0][35] === 'Commitment Updates', sheet.log.join());
ok('history line written', sheet.rows[1][35] === '3 Nov 2026: Done', sheet.rows[1][35]);
ok('due date left alone for Done', sheet.rows[1][33] === '2026-11-02');
ok('caches cleared for that officer and district', invalidated.join() === 'ravi.kumar@educategirls.ngo|SITAPUR', invalidated.join());
const writes = sheet.log.length;
res = S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Done' });
ok('the same answer twice is written once', res.success && res.already && sheet.log.length === writes, JSON.stringify(res));

reset();
res = S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Not yet', newDue: '2026-11-15' });
ok('Not yet moves the date and keeps the old one in the history', res.success && res.due === '2026-11-15' && sheet.rows[1][33] === '2026-11-15' &&
   sheet.rows[1][35] === '3 Nov 2026: Not yet, new date 15 Nov 2026 (was 2 Nov 2026)' && sheet.fmt['2,34'] === '@', sheet.rows[1][35]);
res = S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Not yet', newDue: '2026-11-20' });
ok('moved a second time: both lines kept', sheet.rows[1][35].split('\n').length === 2 && /\(was 15 Nov 2026\)$/.test(sheet.rows[1][35]), sheet.rows[1][35]);
res = S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Done' });
ok('then Done: three lines, status Done', sheet.rows[1][34] === 'Done' && sheet.rows[1][35].split('\n').length === 3);

reset();
ok('Not yet without a date: refused, nothing written', S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Not yet' }).message === 'Please pick the new date.' && sheet.log.length === 0);
ok('Not yet with a past date: refused', S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Not yet', newDue: '2026-11-01' }).message === 'The new date cannot be in the past.' && sheet.log.length === 0);
ok('a new date of today is fine', S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Not yet', newDue: '2026-11-03' }).success === true);
reset();
ok('an unknown answer: refused', S.updateCommitment(me, { meetingId: 'MTG-1', status: 'Maybe' }).success === false && sheet.log.length === 0);
ok('someone else\'s meeting: refused, nothing written', /Only the officer/.test(S.updateCommitment(me, { meetingId: 'MTG-2', status: 'Done' }).message) && sheet.log.length === 0);
ok('a meeting with no commitment: refused', /no commitment/.test(S.updateCommitment(me, { meetingId: 'MTG-3', status: 'Done' }).message) && sheet.log.length === 0);
ok('an unknown meeting: refused', S.updateCommitment(me, { meetingId: 'MTG-9', status: 'Done' }).message === 'Meeting not found.');

// wired in: POST only, and BODY_MISSING when the body was lost
ok('POST only', /POST_ONLY = \{[^}]*updateCommitment:1/.test(gs));
ok('routed, with BODY_MISSING for a lost body', /action === 'updateCommitment'\)\s+result = body\.meetingId \? updateCommitment\(session, body\)\s+: \{ success:false, message:'BODY_MISSING' \}/.test(gs));
ok('getAllMyMeetings carries the commitment', /commitWhat:\s+\(cd\[j\]\[COL_CON_WILLDO-1\]/.test(gs) && /commitDue:\s+dayKey_\(cd\[j\]\[COL_CON_DUE-1\]\)/.test(gs) && /commitStatus: \(cd\[j\]\[COL_CON_CSTATUS-1\]/.test(gs));

// ── page: which ones the card shows ──
const P = new Function('dayStr', '_MON_NAMES',
  pageFn('ymdLabel') + pageFn('daysFrom') + pageFn('commitmentsDue') + '\nreturn { ymdLabel, daysFrom, commitmentsDue };')(
  () => '2026-11-03', ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']);
const C = (id, due, status, what) => ({ meetingId: id, status: 'Conducted', commitWhat: what === undefined ? 'Task ' + id : what, commitDue: due, commitStatus: status || '' });
const list = P.commitmentsDue([
  C('late', '2026-11-01'), C('today', '2026-11-03'), C('later', '2026-11-10'), C('done', '2026-10-30', 'Done'),
  C('dropped', '2026-10-30', 'Not happening'), C('moved', '2026-11-02', 'Not yet'), C('nodate', ''), C('notask', '2026-11-01', '', ''),
  { meetingId: 'plan', status: 'Planned', commitWhat: 'x', commitDue: '2026-11-01' }]);
ok('shows only open ones whose date has come, oldest first', list.map(m => m.meetingId).join() === 'late,moved,today', list.map(m => m.meetingId).join());
ok('days late counted right', P.daysFrom('2026-11-01', '2026-11-03') === 2 && P.daysFrom('2026-11-03', '2026-11-03') === 0);
ok('date shown like the rest of the app', P.ymdLabel('2026-11-02') === '2 Nov 2026');
ok('empty list when nothing is due', P.commitmentsDue([C('later', '2026-12-01')]).length === 0 && P.commitmentsDue(null).length === 0);
ok('can be switched off in one place', /var COMMIT_CARD_ON = true;/.test(page) && /var list = COMMIT_CARD_ON \? commitmentsDue\(all\) : \[\];/.test(page));
ok('names and the typed task are never translated', /<div data-notr style="font-weight:700/.test(page) && /'<span data-notr>' \+ esc\(m\.adhikariName\)/.test(page));
['Commitments to check', 'The date has come for these. Did the official do what they said?', 'Met on', 'Due', 'day late', 'days late', 'Due today',
 'Not yet', 'Not happening', 'New date', 'Save new date', 'Please pick the new date.', 'The new date cannot be in the past.',
 'Mark this as not happening? It will leave this list.', 'Marked as done. Thank you.', 'Saved. It will come back on the new date.',
 'Marked as not happening.', 'Could not save. Please try again.', 'Could not save. Please try again when you have signal.',
 'Only the officer who recorded this meeting can answer for it.', 'This meeting has no commitment recorded.', 'Please choose what happened.']
  .forEach(s => ok('Hindi for: ' + s.slice(0, 48), page.indexOf("'" + s + "':") > 0));
ok('no em dash anywhere', gs.indexOf(String.fromCharCode(8212)) < 0 && page.indexOf(String.fromCharCode(8212)) < 0);
console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
