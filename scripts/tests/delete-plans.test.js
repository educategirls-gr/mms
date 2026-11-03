// Delete Meeting for everyone (own meeting, still a plan) and the editor
// helpers that remove plans more than 90 days old: the REAL functions out of
// Code.gs on a fake Plan Meetings sheet and a fake "Deleted Meetings" tab.
//   node scripts/tests/delete-plans.test.js
const fs = require('fs');
const gs = fs.readFileSync('Code.gs', 'utf8');
const page = fs.readFileSync('docs/dashboard.html', 'utf8');
function fnSrc(name) {
  const a = gs.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return gs.slice(a, gs.indexOf('\n}', a) + 2);
}
const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };

const TODAY = new Date(2026, 10, 3);          // 3 Nov 2026
const RAVI = 'ravi@educategirls.ngo', SITA = 'sita@educategirls.ngo';
function plan(id, email, date, status, ev) {
  const r = new Array(24).fill('');
  r[0] = id; r[1] = 'SITAPUR'; r[2] = email === RAVI ? 'Ravi' : 'Sita'; r[4] = email; r[5] = date;
  r[9] = 'Ramesh Kumar'; r[10] = 'BSA'; r[12] = '=agenda that looks like a formula'; r[13] = status; r[22] = ev || '';
  return r;
}
function makeSheet(rows, maxCols) {
  const sh = { rows, maxCols: maxCols || 24, log: [],
    getLastRow: () => rows.length,
    getMaxColumns: () => sh.maxCols,
    insertColumnsAfter: (a, n) => { sh.maxCols += n; sh.log.push('insertCols ' + n); },
    setFrozenRows: () => {},
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    deleteRow: (n) => { sh.log.push('delete ' + rows[n - 1][0]); rows.splice(n - 1, 1); },
    getRange: (row, col, nr, nc) => {
      nr = nr || 1; nc = nc || 1;
      if (col + nc - 1 > sh.maxCols) throw new Error('out of bounds ' + (col + nc - 1) + ' > ' + sh.maxCols);
      const rg = {
        getValue: () => (rows[row - 1] || [])[col - 1] || '',
        getValues: () => { const out = []; for (let i = 0; i < nr; i++) out.push(((rows[row - 1 + i]) || []).slice(col - 1, col - 1 + nc)); return out; },
        setValues: (v) => { v.forEach((x, i) => { rows[row - 1 + i] = rows[row - 1 + i] || []; x.forEach((c, j) => { rows[row - 1 + i][col - 1 + j] = c; }); }); sh.log.push('setValues ' + v.length); return rg; },
        setFontWeight: () => rg
      };
      return rg;
    } };
  return sh;
}
let plans, deleted, invalidated, calDeleted, order;
function reset() {
  const header = ['Meeting ID','District','Employee','Post','Email','Date','Time','Dur','Type','Stakeholder','Post','Purpose','Agenda','Status'];
  while (header.length < 24) header.push('h' + header.length);
  plans = makeSheet([header,
    plan('OLD-1', RAVI, new Date(2026, 5, 10), 'Planned', 'ev-1'),   // 10 Jun: old
    plan('OLD-2', RAVI, '2026-07-01', 'Postponed'),                   // text date: old
    plan('OLD-3', SITA, '15 Jul 2026', 'Follow-up'),                  // "d Mon yyyy": old
    plan('KEEP-C', RAVI, new Date(2026, 4, 1), 'Conducted'),           // old but conducted
    plan('KEEP-X', RAVI, new Date(2026, 4, 1), 'Cancelled'),           // old but cancelled
    plan('KEEP-NEW', RAVI, new Date(2026, 9, 1), 'Planned'),           // 1 Oct: only 33 days
    plan('KEEP-BAD', SITA, 'sometime soon', 'Planned'),               // date unreadable
    plan('EDGE', SITA, new Date(2026, 7, 5), 'Planned')                // 5 Aug: exactly 90 days
  ]);
  deleted = null; invalidated = []; calDeleted = []; order = [];
}
const E = {
  SpreadsheetApp: { openById: () => ({
      getSheetByName: (n) => n === 'Plan Meetings' ? plans : (n === 'Deleted Meetings' ? deleted : null),
      insertSheet: (n) => { order.push('archive'); deleted = makeSheet([], 26); return deleted; } }), flush: () => {} },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  Session: { getScriptTimeZone: () => 'Asia/Kolkata' },
  Utilities: { formatDate: (d, tz, f) => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') },
  CalendarApp: { getDefaultCalendar: () => ({ getEventById: (id) => ({ deleteEvent: () => calDeleted.push(id) }) }) },
  Logger: { log: (s) => { E._log = s; } },
  SPREADSHEET_ID: 'x', MEETINGS_SHEET: 'Plan Meetings', COL_CAL_EVENT: 23,
  invalidateUser: (e, d) => invalidated.push(e + '|' + d), cDel: () => {}
};
const RealDate = Date;
global.Date = class extends RealDate { constructor(...a) { if (a.length) super(...a); else super(TODAY.getTime()); } static now() { return TODAY.getTime(); } };
const consts = gs.match(/var DELETED_SHEET = [^\n]+/)[0] + '\n' + gs.match(/var PLAN_STILL_A_PLAN = [^\n]+/)[0] + '\n' +
               gs.match(/var PLAN_OLD_DAYS = [^\n]+/)[0] + '\n' + gs.match(/var PLAN_OLD_STATUSES = [^\n]+/)[0];
const body = [consts, fnSrc('withScriptLock_'), fnSrc('fmtDateVal'), fnSrc('archiveDeleted_'), fnSrc('planDay_'),
              fnSrc('planOld_'), fnSrc('PLAN_old_1_preview'), fnSrc('PLAN_old_2_DELETE'), fnSrc('deleteMeeting')].join('\n');
const names = Object.keys(E).filter(k => k !== '_log');
const S = new Function(...names, body + '\nreturn { PLAN_old_1_preview, PLAN_old_2_DELETE, deleteMeeting, planDay_ };')(...names.map(k => E[k]));
const ids = () => plans.rows.slice(1).map(r => r[0]).join(',');

// ── preview ──
reset();
let r = S.PLAN_old_1_preview();
ok('preview finds exactly the three old open plans', r.preview && r.count === 3 && r.rows.map(x => x.id).join() === 'OLD-1,OLD-2,OLD-3', JSON.stringify(r.rows.map(x => x.id)));
ok('preview removes nothing and copies nothing', plans.rows.length === 9 && deleted === null && plans.log.length === 0);
ok('preview log says how to go ahead', /Nothing was removed/.test(E._log) && /PLAN_old_2_DELETE/.test(E._log) && /could not be read, left alone: 1/.test(E._log), E._log.split('\n')[1]);
ok('the 90-day line itself is kept (5 Aug, exactly 90 days)', r.rows.every(x => x.id !== 'EDGE'));
ok('dates read in all three shapes', S.planDay_(new Date(2026, 5, 10), 'x') === '2026-06-10' && S.planDay_('2026-07-01') === '2026-07-01' &&
   S.planDay_('15 Jul 2026') === '2026-07-15' && S.planDay_('01-07-2026') === '2026-07-01' && S.planDay_('sometime soon') === '');

// ── delete ──
reset();
r = S.PLAN_old_2_DELETE();
ok('removes the three, keeps conducted, cancelled, recent and unreadable', r.removed === 3 && ids() === 'KEEP-C,KEEP-X,KEEP-NEW,KEEP-BAD,EDGE', ids());
ok('copied to "Deleted Meetings" BEFORE anything was removed', order[0] === 'archive' && deleted.rows.length === 4, JSON.stringify(order));
ok('the copy has when, by whom, why, then the whole row', deleted.rows[0].slice(0, 4).join('|') === 'Deleted At|Deleted By|Reason|Meeting ID' &&
   deleted.rows[1][1] === 'PLAN_old_2_DELETE' && /90 days/.test(deleted.rows[1][2]) && deleted.rows[1][3] === 'OLD-1');
ok('the tab was widened to fit 3 + 24 columns', deleted.log.indexOf('insertCols 1') >= 0, deleted.log.join());
ok('text that looks like a formula stays text in the copy', deleted.rows[1][3 + 12] === "'=agenda that looks like a formula");
ok('caches cleared for each officer affected', invalidated.sort().join() === 'ravi@educategirls.ngo|SITAPUR,sita@educategirls.ngo|SITAPUR', invalidated.join());
ok('no calendar invites touched for old plans', calDeleted.length === 0);
r = S.PLAN_old_2_DELETE();
ok('running it again removes nothing more', r.removed === 0 && ids() === 'KEEP-C,KEEP-X,KEEP-NEW,KEEP-BAD,EDGE');

// a row moved under it: found again by its id, never the wrong row
reset();
const realGetRange = plans.getRange;
let first = true;
plans.getRange = (row, col, nr, nc) => {
  // the first id check: a new plan lands at the top just before it, so every
  // row number read earlier is now one out
  if (first && col === 1 && !nr) { first = false; plans.rows.splice(1, 0, plan('NEW-ROW', SITA, new Date(2026, 10, 3), 'Planned')); }
  return realGetRange(row, col, nr, nc);
};
r = S.PLAN_old_2_DELETE();
ok('if a row is added mid-way, only the intended ids go', r.removed === 3 && /NEW-ROW/.test(ids()) && !/OLD-/.test(ids()), ids());

// ── Delete Meeting from Manage Meetings ──
reset();
ok('someone else\'s meeting: refused, nothing written', S.deleteMeeting('OLD-3', RAVI).message === 'You can delete only your own meetings.' && plans.log.length === 0 && deleted === null);
ok('a postponed meeting: refused', S.deleteMeeting('OLD-2', RAVI).message === 'Only a meeting that is still planned can be deleted.' && plans.log.length === 0);
ok('a conducted meeting: refused', S.deleteMeeting('KEEP-C', RAVI).success === false && plans.log.length === 0);
r = S.deleteMeeting('OLD-1', 'Ravi@educategirls.ngo');
ok('own planned meeting: deleted (email case ignored)', r.success && !/OLD-1/.test(ids()), JSON.stringify(r));
ok('copied first, marked "Deleted from Manage Meetings", by the officer', deleted.rows[1][2] === 'Deleted from Manage Meetings' && deleted.rows[1][1] === RAVI);
ok('its calendar invite is removed', calDeleted.join() === 'ev-1');
ok('caches cleared, so the list does not show it again', invalidated.join() === 'ravi@educategirls.ngo|SITAPUR');
ok('a follow-up that is still a plan can be deleted too', S.deleteMeeting('OLD-3', SITA).success === true);
ok('an unknown meeting: not found', S.deleteMeeting('NOPE', RAVI).message === 'Meeting not found.');

// ── page ──
ok('the Field-only hiding is gone', page.indexOf("if (_delRow) _delRow.style.display = 'none';\n    }") < 0);
ok('Delete shows only for a meeting that is still a plan', /_delRow\.style\.display = \(!m\.status \|\| m\.status === 'Planned' \|\| m\.status === 'Follow-up'\) \? '' : 'none';/.test(page));
ok('Hindi for the two new refusals', page.indexOf("'You can delete only your own meetings.':") > 0 && page.indexOf("'Only a meeting that is still planned can be deleted.':") > 0);
ok('no em dash anywhere', gs.indexOf(String.fromCharCode(8212)) < 0 && page.indexOf(String.fromCharCode(8212)) < 0);
global.Date = RealDate;
console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
