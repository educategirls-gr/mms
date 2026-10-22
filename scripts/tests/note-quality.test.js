// The REAL NOTE_quality out of Code.gs on a stub Conducted sheet.
//   node scripts/tests/note-quality.test.js
const fs = require('fs');
const src = fs.readFileSync('Code.gs', 'utf8');
function fnSrc(name) {
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n}', a) + 2);
}
function row(id, officer, note, outcome) {
  const r = new Array(32).fill('');
  r[0] = id; r[2] = officer; r[15] = note; r[12] = 'Agenda text for the meeting'; r[31] = outcome || '';
  return r;
}
const NEW = (what, said, next) => 'Discussed: ' + what + '\nOfficial said: ' + said + (next ? '\nNext step: ' + next : '');
const SHEET = [new Array(32).fill('h'),
  row('M1', 'Ravi', 'ok'),
  row('M2', 'Ravi', 'BSA se mile, data ki baat hui'),
  row('M3', 'Ravi', 'Meeting held'),
  row('M4', 'Ravi', NEW('Enrollment progress of the block was reviewed in detail', 'BSA agreed to share block wise data by Friday', 'Collect data'), 'Commitment'),
  row('M5', 'Ravi', NEW('KGBV vacant seats were discussed with the warden', 'Warden will send the list of vacant seats next week'), 'Information'),
  row('M6', 'Ravi', NEW('Shiksha Rath route for the block was discussed', 'Permission given for the route on Monday'), 'Permission'),
  row('M7', 'Sita', 'Discussion on enrollment and retention of girls in the district schools.'),
  row('M8', 'Sita', 'Discussion on enrollment and retention of girls in the district schools.'),
  row('M9', 'Sita', NEW('Short', 'Ok')),
];

const logs = [];
const E = { SpreadsheetApp: { openById: () => ({ getSheetByName: () => ({ getDataRange: () => ({ getValues: () => SHEET }) }) }) },
            SPREADSHEET_ID: 'x', CONDUCTED_SHEET: 'Conducted Meetings', COL_CON_OUTCOME: 32, Logger: { log: (s) => logs.push(s) } };
const names = Object.keys(E);
const NOTE_quality = new Function(...names, fnSrc('NOTE_quality') + '\nreturn NOTE_quality;')(...names.map(k => E[k]));
const r = NOTE_quality();

const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };
ok('splits the two forms by the labels: 5 old, 4 new', r.oldForm.meetings === 5 && r.threeQuestionForm.meetings === 4, JSON.stringify([r.oldForm, r.threeQuestionForm]));
ok('labels are not counted as the officer\'s words', r.threeQuestionForm.medianChars < NEW('Enrollment progress of the block was reviewed in detail', 'BSA agreed to share block wise data by Friday', 'Collect data').length);
ok('copied note counted in the old form', r.oldForm.repeated === 1 && r.threeQuestionForm.repeated === 0);
ok('outcome recorded: 3 of 4 new, 0 of 5 old', r.threeQuestionForm.outcomeRecorded === '75%' && r.oldForm.outcomeRecorded === '0%');
ok('under 30: "ok", "Meeting held" and the 29-character Hinglish note (3 of 5)', r.oldForm.under30 === '60%', r.oldForm.under30);
ok('same-officer comparison only for 3+ notes each side (Ravi yes, Sita no)', r.sameOfficers.length === 1 && r.sameOfficers[0].officer === 'Ravi' && r.sameOfficers[0].after > r.sameOfficers[0].before);
ok('old overall fields still returned', r.conductedMeetings === 9 && r.repeatedNotes === 1 && r.shortestWriters.length > 0);
ok('the log is a readable table', /Old box\s+Three questions/.test(logs[0]) && /Ravi\s+\d+ -> \d+/.test(logs[0]));
ok('no em dash', logs[0].indexOf(String.fromCharCode(8212)) < 0 && src.indexOf(String.fromCharCode(8212)) < 0);
console.log('\n--- what the execution log will show ---\n' + logs[0]);
console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
