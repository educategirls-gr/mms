// Builds docs/_test.html: the real dashboard with a stub in front that answers
// every server call with canned data, and a fake State session that is only
// ever read, never written. The real stored session on this origin is not
// read, replaced or removed. docs/_*.html is git-ignored; delete it after use.
//   node scripts/tests/make-harness.js     then open http://localhost:8765/_test.html
const fs = require('fs');
const src = fs.readFileSync('docs/dashboard.html', 'utf8');

const STUB = `<script>
(function(){
  var FAKE = { name:'Test Lead', email:'test.lead@example.org', role:'State', designation:'State Lead',
               district:'UNNAO', districts:['UNNAO'], token:'fake' };
  var gi = Storage.prototype.getItem, ri = Storage.prototype.removeItem, cl = Storage.prototype.clear;
  Storage.prototype.getItem = function(k){ return k === 'eg_grm_session' ? JSON.stringify(FAKE) : gi.call(this, k); };
  Storage.prototype.removeItem = function(k){ if (/^eg_grm_/.test(k)) return; return ri.call(this, k); };
  Storage.prototype.clear = function(){ if (this === window.sessionStorage) return; return cl.call(this); };
  function d(offset) {
    var M = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var x = new Date(); x.setDate(x.getDate() + offset);
    return x.getDate() + ' ' + M[x.getMonth()] + ' ' + x.getFullYear();
  }
  var PLANNED = [
    { meetingId:'MTG-001', date:d(-3), meetingTime:'11:00 AM', adhikariName:'Ramesh Kumar', adhikariPost:'BSA', district:'UNNAO', purpose:'Enrollment', status:'Planned' },
    { meetingId:'MTG-002', date:d(1),  meetingTime:'3:00 PM',  adhikariName:'K.K Singh',    adhikariPost:'BEO', district:'UNNAO', purpose:'Training',   status:'Postponed' }
  ];
  var R = {
    getPlanDistricts: ['UNNAO','SITAPUR','LUCKNOW'],
    getDropdownData: { stakeholders:['BSA','BEO','DM','Other'], purposes:['Enrollment','Training'], blocksByDistrict:{}, metBefore:[] },
    getMyMeetings: PLANNED, getAllMyMeetings: [], getStateAllMeetings: []
  };
  window.fetch = function(url) {
    var a = (/[?&]action=([^&]+)/.exec(url) || [])[1] || '';
    var body = R.hasOwnProperty(a) ? R[a] : { success:true };
    return new Promise(function(res){ setTimeout(function(){ res(new Response(JSON.stringify(body))); }, 60); });
  };
})();
</script>`;

const out = src.replace('<head>', '<head>\n' + STUB);
if (out === src) throw new Error('no <head> found');
fs.writeFileSync('docs/_test.html', out);
console.log('docs/_test.html written');
