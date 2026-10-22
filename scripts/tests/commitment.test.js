// The commitment boxes on the conduct form (22 Oct 2026): the REAL server
// helpers out of Code.gs on a fake sheet, and the REAL page checks out of
// docs/dashboard.html on fake inputs.
//   node scripts/tests/commitment.test.js
const fs = require('fs');
const gs = fs.readFileSync('Code.gs', 'utf8');
const page = fs.readFileSync('docs/dashboard.html', 'utf8');
function fnSrc(src, name) {
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n}', a) + 2);
}
const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };

// ── server ─────────────────────────────────────────────────────────────
function fakeSheet(maxCols, header) {
  const cells = {}, fmt = {}, log = [];
  Object.keys(header || {}).forEach(c => { cells['1,' + c] = header[c]; });
  const sh = {
    cells, fmt, log, maxCols,
    getMaxColumns: () => sh.maxCols,
    insertColumnsAfter: (after, n) => { log.push('insert ' + after + '+' + n); sh.maxCols += n; },
    getRange: (r, c) => {
      if (c > sh.maxCols) throw new Error('Range out of bounds: column ' + c + ' of ' + sh.maxCols);
      const rg = {
        getValue: () => cells[r + ',' + c] || '',
        setValue: (v) => { cells[r + ',' + c] = v; return rg; },
        setNumberFormat: (f) => { fmt[r + ',' + c] = f; return rg; }
      };
      return rg;
    }
  };
  return sh;
}
const consts = gs.match(/var COL_CON_WILLDO = \d+;[\s\S]*?var COL_CON_DUE\s*= \d+;/)[0];
const S = new Function(consts + '\n' + fnSrc(gs, 'cleanCommitment_') + '\n' + fnSrc(gs, 'saveCommitment_') +
                       '\nreturn { cleanCommitment_, saveCommitment_, COL_CON_WILLDO, COL_CON_DUE };')();
ok('AG and AH, right after Outcome (AF = 32)', S.COL_CON_WILLDO === 33 && S.COL_CON_DUE === 34);

let sh = fakeSheet(32, { 32: 'Outcome' });
ok('saved for Commitment', S.saveCommitment_(sh, 50, { outcome: 'Commitment', commitWhat: '  Share the  village-wise list ', commitDue: '2026-11-05' }) === true);
ok('the sheet only had 32 columns: two added at the far right, nothing moved', sh.log.join() === 'insert 32+2' && sh.maxCols === 34, sh.log.join());
ok('headers written once', sh.cells['1,33'] === 'Official Will Do' && sh.cells['1,34'] === 'Commitment Due');
ok('task tidied and saved in AG', sh.cells['50,33'] === 'Share the village-wise list', sh.cells['50,33']);
ok('date saved in AH as text, so Sheets cannot turn it into another day', sh.cells['50,34'] === '2026-11-05' && sh.fmt['50,34'] === '@');
S.saveCommitment_(sh, 51, { outcome: 'Commitment', commitWhat: 'Issue a letter to the BEOs', commitDue: '2026-11-10' });
ok('second save: no more columns added, headers left alone', sh.log.length === 1 && sh.cells['51,33'] === 'Issue a letter to the BEOs');

sh = fakeSheet(40, { 33: 'Official Will Do', 34: 'Commitment Due' });
S.saveCommitment_(sh, 9, { outcome: 'Commitment', commitWhat: '- share the list with the team', commitDue: '' });
ok('a sheet that is already wide enough is not touched', sh.log.length === 0 && sh.maxCols === 40);
ok('text starting with "-" stays text, not a formula', sh.cells['9,33'] === "'- share the list with the team", sh.cells['9,33']);
ok('no date sent: AH left empty', !('9,34' in sh.cells));
S.saveCommitment_(sh, 10, { outcome: 'Commitment', commitWhat: '=IMPORTRANGE("x")', commitDue: '5 Nov' });
ok('"=" cannot become a formula', sh.cells['10,33'] === "'=IMPORTRANGE(\"x\")");
ok('a date in any other shape is dropped, never guessed', !('10,34' in sh.cells));

sh = fakeSheet(32, {});
ok('another outcome: nothing written, even with text left in the box', S.saveCommitment_(sh, 5, { outcome: 'Courtesy', commitWhat: 'left over', commitDue: '2026-11-05' }) === false && sh.log.length === 0 && Object.keys(sh.cells).length === 0);
ok('Commitment with an empty task (an older page): nothing written', S.saveCommitment_(sh, 5, { outcome: 'Commitment' }) === false && sh.log.length === 0);
ok('very long text cut at 300', S.cleanCommitment_({ outcome: 'Commitment', commitWhat: 'x'.repeat(500) }).what.length === 300);

// Wired in after the outcome is written, in a try of its own, so a failure
// here can never fail the conduct or lose the outcome.
const cm = fnSrc(gs, 'conductMeeting');
const iOutcome = cm.indexOf('cSheet.getRange(clr, COL_CON_OUTCOME).setValue(oc)');
const iCommit = cm.indexOf('try { saveCommitment_(cSheet, clr, payload); } catch (cmErr) {}');
ok('conductMeeting calls it after the outcome, inside its own try', iOutcome > 0 && iCommit > iOutcome);
const mom = fnSrc(gs, 'createMoMDoc');
ok('MoM shows the commitment through the same check', /var _cm = cleanCommitment_\(d\);/.test(mom) && /Commitment by the Official/.test(mom));

// ── page ───────────────────────────────────────────────────────────────
const vague = page.match(/var _VAGUE_COMMIT = [^\n]+/)[0];
const els = {};
const document = { getElementById: (id) => els[id] || null };
// page functions are indented two spaces and close with "\n  }"
function pageFn(name) {
  const a = page.indexOf('  function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return page.slice(a, page.indexOf('\n  }', a) + 4);
}
const P = new Function('document', vague + '\n' + pageFn('readCommitment') + '\n' + pageFn('syncCommitBox') +
                       '\nreturn { readCommitment, syncCommitBox };')(document);
function set(outcome, what, due, date) {
  els['c-outcome'] = { value: outcome };
  els['c-commit-what'] = { value: what };
  els['c-commit-due'] = { value: due };
  els['c-commit-box'] = { style: { display: 'none' } };
  els['c-date'] = { value: date || '2026-10-22' };
  return P.readCommitment(outcome, date || '2026-10-22');
}
ok('other outcomes need nothing', set('Courtesy', '', '').ok === true && set('Nothing', '', '').what === '');
ok('Commitment with no task: asked for it', /what exactly the official will do/.test(set('Commitment', '', '2026-11-01').msg));
['Agreed to support', 'Full support.', 'will cooperate', 'BEO agreed to support', 'BEO sir agreed to support',
 'Agreed to support with all activities', 'he will help', 'Will provide full cooperation']
  .forEach(v => { const r = set('Commitment', v, '2026-11-01'); ok('vague "' + v + '" is stopped', r.ok === false && /on its own/.test(r.msg), JSON.stringify(r)); });
['support', 'Sahyog'].forEach(v => ok('"' + v + '" is too short to be a task', /what exactly/.test(set('Commitment', v, '2026-11-01').msg)));
['Share the village-wise list of out-of-school girls', 'Will support the enrolment drive on 5 Nov in all schools',
 'BSA will issue a letter to all BEOs', 'Visit KGBV Bangarmau next week', 'Data dega 15 tak']
  .forEach(v => ok('real task passes: "' + v + '"', set('Commitment', v, '2026-11-01').ok === true));
ok('no date: asked for it', /by when/.test(set('Commitment', 'Issue a letter to the BEOs', '').msg));
ok('date before the meeting: refused', /cannot be before the meeting/.test(set('Commitment', 'Issue a letter to the BEOs', '2026-10-20', '2026-10-22').msg));
ok('same day as the meeting is fine', set('Commitment', 'Issue a letter to the BEOs', '2026-10-22', '2026-10-22').ok === true);
const good = set('Commitment', '  Issue   a letter to the BEOs ', '2026-11-01');
ok('task sent tidied, with its date', good.what === 'Issue a letter to the BEOs' && good.due === '2026-11-01');

set('Commitment', '', ''); P.syncCommitBox();
ok('box shows for Commitment, its date cannot go before the meeting', els['c-commit-box'].style.display === '' && els['c-commit-due'].min === '2026-10-22');
els['c-outcome'].value = 'Information'; P.syncCommitBox();
ok('box hides for any other outcome', els['c-commit-box'].style.display === 'none');

ok('payload sends the two only for Commitment', /commitWhat: outcome === 'Commitment' \? readCommitment\(outcome, date\)\.what : ''/.test(page));
ok('kept in the draft like the other boxes', /_DRAFT_FIELDS = \[[^\]]*'c-commit-what','c-commit-due'/.test(page));
ok('cleared when the window opens, shown again for a restored draft', /restoreConductDraft\(m\.meetingId\);\s*syncCommitBox\(\);/.test(page) && /_cw\.value = ''/.test(page));

// every new English string has its Hindi line
['Commitment, the official agreed to do a specific task', 'What exactly will the official do?',
 'Share the village-wise list of out-of-school girls', 'By when?',
 'If no date was given, pick the day you will check back with them.',
 'Only a specific task counts, such as sharing a list, issuing a letter or visiting a school. "Agreed to support" on its own is not a commitment, so choose another outcome for that.',
 'Please write what exactly the official will do.',
 '"Support" on its own is not a task anyone can check. Please write the specific thing the official will do, or choose another outcome.',
 'Please choose by when the official will do it.', 'The date for the commitment cannot be before the meeting.']
  .forEach(s => ok('Hindi for: ' + s.slice(0, 50), page.indexOf("'" + s + "':") > 0));
ok('the old option text and its Hindi line are gone', page.indexOf('Commitment obtained, the official agreed to do something') < 0);
ok('no em dash anywhere', gs.indexOf(String.fromCharCode(8212)) < 0 && page.indexOf(String.fromCharCode(8212)) < 0);
console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
