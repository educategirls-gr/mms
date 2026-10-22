// The REAL NOTE_quality out of Code.gs on a stub Conducted sheet.
//   node scripts/tests/note-quality.test.js
const fs = require('fs');
const src = fs.readFileSync('Code.gs', 'utf8');
function fnSrc(name) {
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n}', a) + 2);
}
function row(id, officer, note, outcome, email) {
  const r = new Array(32).fill('');
  r[0] = id; r[2] = officer; r[4] = email || ''; r[15] = note; r[12] = 'Agenda text for the meeting'; r[31] = outcome || '';
  return r;
}
const NEW = (what, said, next) => 'Discussed: ' + what + '\nOfficial said: ' + said + (next ? '\nNext step: ' + next : '');
const OM1 = 'om.prakash@educategirls.ngo', OM2 = 'om.prakash2@educategirls.ngo';
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
  // Two different people with one name. The first wrote 2 old notes and 2 new,
  // the second 1 old and 1 new: apart, neither has 3 on each side. Counted by
  // name they would add up to 3 and 3 and show as one improving officer.
  row('M10', 'OM PRAKASH', 'Met the BEO, talked about enrollment', '', OM1),
  row('M11', 'OM PRAKASH', 'Visited KGBV with the team', '', OM1),
  row('M12', 'OM PRAKASH', NEW('Girls out of school in two villages were discussed', 'BEO will call the ARPs on Monday'), 'Commitment', OM1),
  row('M13', 'OM PRAKASH', NEW('Shiksha Rath dates in the block were discussed', 'Dates will be confirmed after the review meeting'), 'Awaiting', OM1),
  row('M14', 'Om Prakash', 'BSA meeting', '', OM2.toUpperCase()),
  row('M15', 'Om Prakash', NEW('Courtesy visit to the new BDO of the block', 'BDO welcomed the work of the team'), 'Courtesy', OM2),
];

const logs = [];
const E = { SpreadsheetApp: { openById: () => ({ getSheetByName: () => ({ getDataRange: () => ({ getValues: () => SHEET }) }) }) },
            SPREADSHEET_ID: 'x', CONDUCTED_SHEET: 'Conducted Meetings', COL_CON_OUTCOME: 32, Logger: { log: (s) => logs.push(s) } };
const names = Object.keys(E);
const NOTE_quality = new Function(...names, fnSrc('NOTE_quality') + '\nreturn NOTE_quality;')(...names.map(k => E[k]));
const r = NOTE_quality();

const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };
ok('splits the two forms by the labels: 8 old, 7 new', r.oldForm.meetings === 8 && r.threeQuestionForm.meetings === 7, JSON.stringify([r.oldForm, r.threeQuestionForm]));
ok('labels are not counted as the officer\'s words', r.threeQuestionForm.medianChars < NEW('Enrollment progress of the block was reviewed in detail', 'BSA agreed to share block wise data by Friday', 'Collect data').length);
ok('copied note counted in the old form', r.oldForm.repeated === 1 && r.threeQuestionForm.repeated === 0);
ok('outcome recorded: 6 of 7 new, 0 of 8 old', r.threeQuestionForm.outcomeRecorded === '86%' && r.oldForm.outcomeRecorded === '0%', r.threeQuestionForm.outcomeRecorded);
ok('under 30: "ok", "Meeting held", the 29-character Hinglish note, "Visited KGBV..." and "BSA meeting" (5 of 8)', r.oldForm.under30 === '63%', r.oldForm.under30);
ok('same-officer comparison only for 3+ notes each side (Ravi yes, Sita no)', r.sameOfficers.length === 1 && r.sameOfficers[0].officer === 'Ravi' && r.sameOfficers[0].after > r.sameOfficers[0].before, JSON.stringify(r.sameOfficers));
ok('two people named OM PRAKASH are counted apart (by email), so neither reaches 3+3', !r.sameOfficers.some(x => /om prakash/i.test(x.officer)));
const oms = r.shortestWriters.concat(r.longestWriters).filter(x => /om prakash/i.test(x.officer));
ok('a shared name is shown with the email beside it', oms.length >= 2 && oms.every(x => x.officer.indexOf('<' + OM1 + '>') > 0 || x.officer.indexOf('<' + OM2 + '>') > 0), JSON.stringify(oms));
ok('email case does not split one person', !r.shortestWriters.concat(r.longestWriters).some(x => x.officer.indexOf(OM2.toUpperCase()) >= 0) &&
   r.shortestWriters.concat(r.longestWriters).filter(x => x.officer.indexOf(OM2) >= 0).every(x => x.meetings === 2));
ok('a name nobody shares is shown plain, no email', r.shortestWriters.concat(r.longestWriters).some(x => x.officer === 'Ravi'));
ok('old overall fields still returned', r.conductedMeetings === 15 && r.repeatedNotes === 1 && r.shortestWriters.length > 0);
ok('repeated notes listed by name', r.repeatedNotesByOfficer.Sita === 1, JSON.stringify(r.repeatedNotesByOfficer));
ok('the log is a readable table', /Old box\s+Three questions/.test(logs[0]) && /Ravi\s+\d+ -> \d+/.test(logs[0]));
ok('no em dash', logs[0].indexOf(String.fromCharCode(8212)) < 0 && src.indexOf(String.fromCharCode(8212)) < 0);
console.log('\n--- what the execution log will show ---\n' + logs[0]);
console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
