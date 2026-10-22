// The REAL COMMIT_dryRun and extractCommitments_ out of Code.gs, on a stub
// Conducted sheet and a stub AI. Checks that it picks three-question notes,
// judges the AI against the outcome the officer chose, and writes nothing.
//   node scripts/tests/commit-dry-run.test.js
const fs = require('fs');
const src = fs.readFileSync('Code.gs', 'utf8');
function fnSrc(name) {
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n}', a) + 2);
}
function lineSrc(name) {                       // one-line helpers such as _parseJson_
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n', a));
}

function row(id, officer, note, outcome, date) {
  const r = new Array(32).fill('');
  r[0] = id; r[1] = 'SITAPUR'; r[2] = officer; r[9] = 'Ramesh Kumar'; r[10] = 'BEO'; r[11] = 'Enrollment';
  r[13] = date || '2026-10-01'; r[15] = note; r[31] = outcome || '';
  return r;
}
const NEW = (what, said, next) => 'Discussed: ' + what + '\nOfficial said: ' + said + (next ? '\nNext step: ' + next : '');
const SHEET = [new Array(32).fill('h'),
  // old single box: never picked, however long
  row('OLD1', 'Ravi', 'A very long old style note. BEO agreed to share the data by Friday and promised a letter to all schools in the block.', 'Commitment'),
  row('C1', 'Ravi', NEW('Out of school girls in the block', 'BEO agreed to share the village wise list by Friday', 'Collect the list'), 'Commitment', '2026-10-02'),
  row('C2', 'Sita', NEW('KGBV admissions', 'Warden will send the vacant seat list next week'), 'Commitment', '2026-10-03'),
  row('C3', 'Sita', NEW('Shiksha Rath route', 'BDO will issue a letter to the pradhans'), 'Commitment', '2026-10-04'),
  row('X1', 'Ravi', NEW('Introduction to the new BSA', 'BSA welcomed the work of the team'), 'Courtesy', '2026-10-05'),
  row('X2', 'Sita', NEW('Data review', 'Will see later'), 'Nothing', '2026-10-06'),
  // other outcomes are neither side of the test
  row('I1', 'Ravi', NEW('Enrollment data', 'Shared the block data in the meeting'), 'Information', '2026-10-07'),
];

// The stub AI: a real-looking answer for each note, one invented quote, and one
// meeting where the AI does not answer at all.
const AI = {
  'village wise list': '```json\n{"commitments":[{"what":"Share village wise list","by":"BEO","due":"Friday","evidence":"BEO agreed to share the village wise list by Friday"}]}\n```',
  'vacant seat list': '{"commitments":[{"what":"Send vacant seat list","by":"Warden","due":"next week","evidence":"The warden promised to send it within two days"}]}',
  'letter to the pradhans': '',                       // the AI did not answer
  'welcomed the work': '{"commitments":[]}',
  'Will see later': '{"commitments":[{"what":"Review data","by":"Official","due":"","evidence":"Will see later"}]}'
};
const asked = [], logs = [], writes = [];
const sheet = { getDataRange: () => ({ getValues: () => SHEET }),
                getRange: () => { writes.push('getRange'); return { setValue: () => writes.push('setValue'), setValues: () => writes.push('setValues') }; },
                appendRow: () => writes.push('appendRow') };
const E = {
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }) },
  SPREADSHEET_ID: 'x', CONDUCTED_SHEET: 'Conducted Meetings', COL_CON_OUTCOME: 32,
  Logger: { log: (s) => logs.push(s) }, Utilities: { sleep: () => {} },
  MailApp: { sendEmail: () => writes.push('mail') },
  fmtDateVal: (d) => String(d),
  callLLM: (prompt) => { asked.push(prompt); const k = Object.keys(AI).find(k => prompt.indexOf(k) >= 0); return k ? AI[k] : '{"commitments":[]}'; }
};
const names = Object.keys(E);
const body = [lineSrc('_parseJson_'), fnSrc('extractCommitments_'), fnSrc('COMMIT_dryRun')].join('\n');
const COMMIT_dryRun = new Function(...names, body + '\nreturn COMMIT_dryRun;')(...names.map(k => E[k]));

const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };
const r = COMMIT_dryRun();
const t = r.summary;

ok('only three-question notes, only Commitment / Courtesy / Nothing: 5 meetings sent to the AI', asked.length === 5, asked.length);
ok('the old single-box note is never sent', !asked.some(p => p.indexOf('A very long old style note') >= 0));
ok('the Information meeting is not part of the test', !asked.some(p => p.indexOf('Shared the block data') >= 0));
ok('newest Commitment meeting comes first', r.results[0].date === '2026-10-04' && r.results[0].officerChose === 'Commitment', JSON.stringify(r.results.map(x => x.date)));
ok('AI that did not answer is counted apart, not as "no commitment"', t.aiFailed === 1 && r.results[0].aiFailed === true, JSON.stringify(t));
ok('Commitment side: AI found one in 2 of 2 answered', t.marked === 2 && t.markedFound === 2, JSON.stringify(t));
ok('Courtesy / Nothing side: AI found one in 1 of 2 (the "Will see later" one)', t.control === 2 && t.controlFound === 1, JSON.stringify(t));
ok('an invented quote is caught: 2 of 3 quotes are in the note', t.quotes === 3 && t.quotesOk === 2, JSON.stringify(t));
ok('JSON inside ```json fences is still read', r.results.some(x => x.commitments.length && x.commitments[0].what === 'Share village wise list'));
ok('limit works: 2 Commitment and 1 other', (COMMIT_dryRun(2), asked.length === 5 + 3), asked.length);
ok('writes nothing to the sheet and sends no email', writes.length === 0, writes.join());

const log = logs[0];
ok('log opens with the plain summary', /^COMMITMENT DRY RUN/.test(log) && /Officer chose Commitment: AI found one in 2 of 2/.test(log) && /should be 0/.test(log));
ok('log names the failed AI call and how to check it', /AI did not answer for 1 meeting/.test(log) && /LLM_probe\(\)/.test(log));
ok('log shows only what the official said, not the whole note', /Official said: BEO agreed to share the village wise list by Friday\n/.test(log) && log.indexOf('Official said: Out of school') < 0);
ok('log marks the invented quote', /QUOTE NOT IN NOTE/.test(log) && /quote checked/.test(log));
ok('no em dash', log.indexOf(String.fromCharCode(8212)) < 0 && src.indexOf(String.fromCharCode(8212)) < 0);
console.log('\n--- what the execution log will show ---\n' + log);
console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
