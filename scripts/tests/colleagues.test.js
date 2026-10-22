// The REAL colleague lookup and the two colleague emails out of Code.gs, on
// the kinds of text officers actually type into the colleague box.
//   node scripts/tests/colleagues.test.js
const fs = require('fs');
const src = fs.readFileSync('Code.gs', 'utf8');
function fnSrc(name) {
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n}', a) + 2);
}
const FNS = ['_emailEsc', 'colleagueNames_', 'colleagueKey_', 'colleaguesFound_',
             'sendColleagueNotification', 'sendColleagueNotificationTo_',
             'sendMOMNotification', 'sendMOMNotificationTo_'];

// A slice of the real employee master (the sign-in copy is keyed by email).
const MASTER = {};
[['Nigar Yasin', 'District Operational Lead'], ['Vikrant kumar', 'District Program Officer'],
 ['Sashi Prakash', 'District Program Training Officer'], ['Buddh Vilas', 'District Impact Specialist'],
 ['Shyam Narayan Nath', 'District Program Officer'], ['AJAY KUMAR', 'Block Program Officer'],
 ['Sanjay Kumar', 'Block Program Officer'], ['Anand Kumar Shukla', 'Block Program Officer'],
 ['Balwant Singh', 'District Operational Lead']].forEach(([name, des]) => {
  const email = name.toLowerCase().replace(/\s+/g, '.') + '@educategirls.ngo';
  MASTER[email] = { name, designation: des, email };
});

function load(master) {
  const mail = [];
  const E = {
    STATE_NAME: 'Uttar Pradesh', SITE_HOST: 'dataimpact.in',
    MailApp: { sendEmail: (o) => mail.push(o) },
    empMirrorRead_: () => master, empReadFromSheet_: () => null
  };
  const names = Object.keys(E);
  const api = new Function(...names, FNS.map(fnSrc).join('\n') + '\nreturn {' + FNS.join(',') + '};')(...names.map(k => E[k]));
  api.mail = mail;
  return api;
}

const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };
const api = load(MASTER);
const found = (raw, me) => api.colleaguesFound_(raw, me).map(c => c.name);

ok('splits on commas', JSON.stringify(api.colleagueNames_('Nigar Yasin, Vikrant Kumar, Shashi Prakash')) === JSON.stringify(['Nigar Yasin', 'Vikrant Kumar', 'Shashi Prakash']));
ok('splits on & and "and"', JSON.stringify(api.colleagueNames_('AJAY KUMAR & SANJAY KUMAR and Balwant Singh')) === JSON.stringify(['AJAY KUMAR', 'SANJAY KUMAR', 'Balwant Singh']));
ok('"and" inside a name does not split it', JSON.stringify(api.colleagueNames_('Anand Kumar Shukla, Chandan')) === JSON.stringify(['Anand Kumar Shukla', 'Chandan']));
ok('empty box: nobody', api.colleagueNames_('').length === 0 && api.colleagueNames_(null).length === 0);

ok('three typed, two in the master (Shashi is Sashi in the sheet)',
   JSON.stringify(found('Nigar Yasin, Vikrant Kumar, Shashi Prakash')) === JSON.stringify(['Nigar Yasin', 'Vikrant kumar']), found('Nigar Yasin, Vikrant Kumar, Shashi Prakash'));
ok('"ji" and case do not matter', JSON.stringify(found('Shyam ji and Buddh vilas ji')) === JSON.stringify(['Buddh Vilas']), found('Shyam ji and Buddh vilas ji'));
ok('& between two block officers finds both', found('AJAY KUMAR & SANJAY KUMAR').length === 2);
ok('the same person twice gets one email', found('Balwant Singh, balwant singh ji').length === 1);
ok('the officer filing it is never mailed', found('Balwant Singh, Buddh Vilas', 'balwant.singh@educategirls.ngo').join() === 'Buddh Vilas');
ok('field coordinators outside the master: nobody, no error', found('Premlata ji, Team Member, FC').length === 0);

const meeting = { email: 'officer@educategirls.ngo', colleagueName: 'Nigar Yasin, Vikrant Kumar, Shashi Prakash', colleaguePost: 'DOL, DPO, DPTO',
                  adhikariName: 'Ramesh Kumar', adhikariPost: 'BSA', district: 'SITAPUR', meetingDate: '2026-10-25', meetingTime: '11:00 AM',
                  purpose: 'Enrollment', agenda: 'Agenda', employeeName: 'Officer', meetingId: 'MTG-1', keyPoints: 'Discussed: x', conductDate: '2026-10-25' };
api.sendColleagueNotification(meeting, 'MTG-1');
ok('invitation: one email each to the two found', api.mail.length === 2 &&
   api.mail[0].to === 'nigar.yasin@educategirls.ngo' && api.mail[1].to === 'vikrant.kumar@educategirls.ngo', api.mail.map(m => m.to).join());
ok('each invitation greets its own reader', /Dear <strong>Nigar Yasin<\/strong>/.test(api.mail[0].htmlBody) && /Dear <strong>Vikrant kumar<\/strong>/.test(api.mail[1].htmlBody));
ok('no invitation greets the whole list', api.mail.every(m => m.htmlBody.indexOf('Dear <strong>Nigar Yasin, ') < 0));
api.mail.length = 0;
api.sendMOMNotification(meeting, 'https://docs.google.com/x', 'https://drive.google.com/y', '');
ok('MoM: one email each to the two found', api.mail.length === 2 && /Dear <strong>Vikrant kumar<\/strong>/.test(api.mail[1].htmlBody));
api.mail.length = 0;
api.sendColleagueNotification(Object.assign({}, meeting, { colleagueName: 'Balwant Singh' }), 'MTG-2');
ok('a single colleague works as before', api.mail.length === 1 && api.mail[0].to === 'balwant.singh@educategirls.ngo');
api.mail.length = 0;
api.sendColleagueNotification(Object.assign({}, meeting, { colleagueName: '' }), 'MTG-3');
ok('no colleague: no email', api.mail.length === 0);

// One email failing must not stop the others.
const flaky = load(MASTER); let n = 0;
flaky.mail.push = function (o) { if (n++ === 0) throw new Error('quota'); return Array.prototype.push.call(this, o); };
flaky.sendColleagueNotification(meeting, 'MTG-4');
ok('if the first email fails, the second still goes', flaky.mail.length === 1);

ok('no em dash in Code.gs', src.indexOf('—') < 0);
console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
