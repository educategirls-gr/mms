// The REAL email builders out of Code.gs, run with stub services. Checks that
// every email names the state and the site from STATE_NAME / SITE_HOST, and
// that changing those two lines is enough for another state.
//   node scripts/tests/state-name.test.js
const fs = require('fs');
const src = fs.readFileSync('Code.gs', 'utf8');
function fnSrc(name) {
  const a = src.indexOf('function ' + name + '(');
  if (a < 0) throw new Error('missing ' + name);
  return src.slice(a, src.indexOf('\n}', a) + 2);
}
function varVal(name) {
  const m = new RegExp('var ' + name + '\\s*=\\s*\'([^\']*)\'').exec(src);
  if (!m) throw new Error('missing var ' + name);
  return m[1];
}

const FNS = ['_emailEsc', 'sendOTP', 'sendColleagueNotification', 'sendMOMNotification',
             'buildEscalationEmail_', 'buildWeeklyReminder_'];

function load(state, host) {
  const mail = [];
  const E = {
    STATE_NAME: state, SITE_HOST: host, ALLOWED_DOMAIN: 'educategirls.ngo', OTP_EXPIRY_SEC: 600,
    MailApp: { sendEmail: (o) => mail.push(o) },
    CacheService: { getScriptCache: () => ({ put: () => {}, get: () => null }) },
    getEmployeeByEmail: (e) => (e === 'known@educategirls.ngo' ? { name: 'Known Person' } : null),
    getEmployeeByName: () => ({ email: 'colleague@educategirls.ngo', name: 'Colleague' }),
    Logger: { log: () => {} }, Session: { getScriptTimeZone: () => 'Asia/Kolkata' },
    Utilities: { formatDate: (d) => String(d) }
  };
  const names = Object.keys(E);
  const api = new Function(...names, FNS.map(fnSrc).join('\n') + '\nreturn {' + FNS.join(',') + '};')(...names.map(k => E[k]));
  api.mail = mail;
  return api;
}

const fails = [];
const ok = (label, cond, extra) => { console.log((cond ? 'ok   ' : 'FAIL ') + label + (cond ? '' : '   ' + (extra || ''))); if (!cond) fails.push(label); };

const STATE = varVal('STATE_NAME'), HOST = varVal('SITE_HOST');
ok('Code.gs sets STATE_NAME and SITE_HOST', STATE === 'Uttar Pradesh' && HOST === 'dataimpact.in', STATE + ' / ' + HOST);

const meeting = { colleagueName: 'Colleague', adhikariName: 'Ramesh Kumar', adhikariPost: 'BSA', district: 'UNNAO',
                  meetingDate: '2026-10-05', meetingTime: '11:00 AM', purpose: 'Enrollment', agenda: 'Agenda',
                  employeeName: 'Officer', meetingId: 'MTG-1', keyPoints: 'Discussed: x\nOfficial said: y',
                  conductDate: '2026-10-05', duration: '1 hr', meetingType: 'One-on-One' };

function checkAll(state, host, tag) {
  const api = load(state, host);

  // Sign-in code: only the words changed; who may sign in did not.
  const r = api.sendOTP('Known@educategirls.ngo ');
  const otp = api.mail[0] || {};
  ok(tag + 'OTP still sent and still succeeds', r.success === true && api.mail.length === 1, JSON.stringify(r));
  ok(tag + 'OTP sender names the state', otp.name === 'EG-MMS ' + state, otp.name);
  ok(tag + 'OTP subject names the state', otp.subject === 'EG Meeting Management System (' + state + ') - Login OTP', otp.subject);
  ok(tag + 'OTP text names state and site', otp.body.indexOf(state + ' (' + host + ') is: ') > 0 && /is: \d{6}\n/.test(otp.body), otp.body.split('\n')[2]);
  ok(tag + 'OTP still refused for another domain', api.sendOTP('x@gmail.com').success === false);
  ok(tag + 'OTP still refused for someone not in the sheet', api.sendOTP('nobody@educategirls.ngo').success === false && api.mail.length === 1);

  api.sendColleagueNotification(meeting, 'MTG-1');
  api.sendMOMNotification(meeting, 'https://docs.google.com/x', 'https://drive.google.com/y', '');
  [['colleague invitation', api.mail[1]], ['MoM email', api.mail[2]]].forEach(([what, m]) => {
    m = m || { htmlBody: '' };
    ok(tag + what + ': sender names the state', m.name === 'EG-MMS ' + state, m.name);
    ok(tag + what + ': header and sign-off name the state',
       m.htmlBody.indexOf('Government Relations, ' + state + '</p>') > 0 && m.htmlBody.indexOf('Government Relations Team, ' + state) > 0);
    ok(tag + what + ': logo from our own site', m.htmlBody.indexOf('src="https://' + host + '/eg-logo.png"') > 0 && m.htmlBody.indexOf('educategirls.ngo/wp-content') < 0);
    ok(tag + what + ': no "undefined" in it', m.htmlBody.indexOf('undefined') < 0);
  });

  const esc = api.buildEscalationEmail_({ priority: 'High', category: 'Resource needed', district: 'UNNAO', conductDate: '5 Oct',
                                          officerName: 'Officer', stakeholder: 'BSA', purpose: 'x', flag: 'Blocked', nextAction: 'y', keyPoints: 'z' });
  ok(tag + 'escalation footer names the site', esc.indexOf(host + '</div>') > 0);

  const wk = api.buildWeeklyReminder_({ name: 'Anupam Gupta', meetings: [] }, '5 Oct to 11 Oct');
  ok(tag + 'weekly reminder footer names the site', wk.indexOf('https://' + host + '</div>') > 0);
}

checkAll(STATE, HOST, '');
// The point of the constants: another state is two lines, and nothing of UP leaks.
checkAll('Rajasthan', 'raj.dataimpact.in', '[RJ] ');

// Nothing outside the constants still spells the state or the site out.
const lines = src.split('\n');
const spelled = lines.filter(l => /Uttar Pradesh|dataimpact\.in/.test(l) && !/^\s*(\/\/|var STATE_NAME|var SITE_HOST)/.test(l));
ok('no other line of Code.gs spells out "Uttar Pradesh" or "dataimpact.in"', spelled.length === 0, spelled.map(l => l.trim().slice(0, 80)).join(' | '));
ok('no em dash in Code.gs', src.indexOf('\u2014') < 0);

console.log('\n' + (fails.length ? fails.length + ' FAILED' : 'all checks pass'));
process.exit(fails.length ? 1 : 0);
