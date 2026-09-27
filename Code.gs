// ============================================================
//  EG GR Reporting System - Main Backend
//  Google Apps Script | Bound to Employee_DB Spreadsheet
// ============================================================

// Switched on the night of 17 Sep 2026. The original document
// (1a7068K07gE40PLkxIs39A6OJvalCK7IgDJTZB5NQH40) stopped opening from Apps Script
// with 'Service Spreadsheets timed out' while a brand new sheet opened in a
// second, so a copy was made from the browser and the script pointed at it.
// The original is untouched and kept as an archive; nothing writes to it now.
var SPREADSHEET_ID   = '1a7068K07gE40PLkxIs39A6OJvalCK7IgDJTZB5NQH40';
var EMPLOYEE_SHEET   = 'Employee_DB';
var MEETINGS_SHEET   = 'Plan Meetings';
var CONDUCTED_SHEET  = 'Conducted Meetings';
var POSTPONED_SHEET  = 'Postponed Meetings';
var CANCELLED_SHEET  = 'Cancelled Meetings';
var DRIVE_ROOT_ID    = '1S_Y79rGOxkaRq5bD_ZEk5nWL1AcajMcd'; // EG-GR-Meetings Drive folder (gr@educategirls.ngo)
var OTP_EXPIRY_SEC   = 600;
var ALLOWED_DOMAIN   = 'educategirls.ngo';

// ============================================================
//  CACHE HELPERS  (GAS CacheService - script-level, 6 hr max)
// ============================================================
var C_TTL_EMP    = 1800;  // 30 min - employee data (rarely changes)
var C_TTL_LIVE   = 600;   // 10 min - meeting lists, stats, reports. Safe to
                          // hold this long because invalidateUser() below clears
                          // every affected key on save, conduct, postpone,
                          // cancel, delete and Govt MoM upload.
var C_TTL_DROP   = 900;   // 15 min - dropdown / colleague lists

function cGet(key) {
  try {
    var v = CacheService.getScriptCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch(e) { return null; }
}
function cPut(key, data, ttl) {
  try {
    var s = JSON.stringify(data);
    if (s.length < 95000) CacheService.getScriptCache().put(key, s, ttl || C_TTL_LIVE);
  } catch(e) {}
}
function cDel() {
  var keys = Array.prototype.slice.call(arguments);
  try { CacheService.getScriptCache().removeAll(keys); } catch(e) {}
}
// ── Not letting one stuck document take the day down ────────────────────
// On the night of 17 Sep 2026 the sheet stopped answering, and every request
// that touched it sat there for the full six minutes Apps Script allows before
// giving up. Sixty of those is the entire daily runtime allowance, so the
// failure fed itself: the more people tried, the less was left for anyone.
//
// The first failure now trips a breaker for five minutes. While it is tripped
// nothing even attempts the document: requests come back in milliseconds
// saying so. The flag is simply a cache entry with a five minute life, so it
// clears itself, and the first request after that tries for real. If the
// document is still stuck that one request pays the six minutes and trips it
// again, which is one hang every five minutes instead of all of them.
//
// Signing in and the open portal are untouched by this, because neither reads
// the document: one reads Script Properties, the other a published file.
var SHEET_TRIP_KEY  = 'SHEET_TRIPPED';
var SHEET_TRIP_SECS = 300;
var SHEET_MISS_KEY  = 'SHEET_MISSES';
var SHEET_MISS_BEFORE_TRIP = 2;
var SHEET_MISS_WINDOW      = 120;
var SHEET_BUSY_MSG  = 'The meetings sheet is not responding at the moment. Nothing was lost. Please try again in a few minutes.';

function sheetBreakerTripped_() {
  try { return CacheService.getScriptCache().get(SHEET_TRIP_KEY) === '1'; }
  catch (e) { return false; }
}
// Two failures, not one. Sheets throws the occasional one-off that clears
// itself, and tripping on that would stop everybody's meetings for five minutes
// over nothing: a cure worse than the illness. Two in a row inside two minutes
// is a pattern rather than a blip. It costs one more slow request before the
// protection starts, which against six hours of lost runtime is nothing.
function sheetBreakerTrip_() {
  try {
    var cache = CacheService.getScriptCache();
    var misses = parseInt(cache.get(SHEET_MISS_KEY) || '0', 10) + 1;
    if (misses >= SHEET_MISS_BEFORE_TRIP) {
      cache.put(SHEET_TRIP_KEY, '1', SHEET_TRIP_SECS);
      cache.remove(SHEET_MISS_KEY);
    } else {
      cache.put(SHEET_MISS_KEY, String(misses), SHEET_MISS_WINDOW);
    }
  } catch (e) {}
}
function sheetBreakerClear_() {
  // A good read clears the count as well, so two failures an hour apart never
  // add up to a trip.
  try { CacheService.getScriptCache().removeAll([SHEET_TRIP_KEY, SHEET_MISS_KEY]); } catch (e) {}
}

// A request that has already spent a long time must not start another read that
// could cost six more minutes. getReportData opens five sheets; if the first one
// hangs there is no sense attempting the other four.
var REQ_START_MS  = new Date().getTime();
var REQ_BUDGET_MS = 90000;
var IS_WEB_REQUEST = false;   // set by apiResponse; the timed jobs leave it false
function requestBudgetSpent_() {
  return IS_WEB_REQUEST && (new Date().getTime() - REQ_START_MS) > REQ_BUDGET_MS;
}

// Thrown rather than returned, so a caller can never mistake "could not read"
// for "there is nothing there" and show an empty screen as if it were the truth.
function sheetBusy_() { throw new Error(SHEET_BUSY_MSG); }

// Editor helpers, for when someone wants to look or to let people back in early.
function BREAKER_status() {
  var on = sheetBreakerTripped_();
  var misses = 0;
  try { misses = parseInt(CacheService.getScriptCache().get(SHEET_MISS_KEY) || '0', 10); } catch (e) {}
  Logger.log((on ? 'Tripped. The sheet is being left alone; it clears itself within five minutes.'
                 : 'Clear. The sheet is being read normally.') +
             String.fromCharCode(10) + 'Recent failures counted: ' + misses + ' of ' + SHEET_MISS_BEFORE_TRIP);
  return on ? 'tripped' : 'clear';
}
function BREAKER_reset() { sheetBreakerClear_(); Logger.log('Breaker cleared. The next request will try the sheet.'); return 'clear'; }

// ── One read per sheet, shared by everyone ──────────────────────────────
// The caches held each person's finished answer, so fifty people opening My
// Meetings meant the same three sheets were read fifty times over inside ten
// minutes, with the hourly jobs and the half hourly snapshot on top of that.
// The document went under exactly that weight on the night of 17 Sep 2026.
// These hold the sheet itself instead, so fifty people cost one read.
//
// A cache value is capped near 100KB, so a sheet is split across numbered keys
// and fetched back in a single getAll. Chunks are deliberately small in
// characters, because a Devanagari note is three bytes per character and the
// cap is on bytes.
//
// Dates survive the trip: getValues hands back real Date objects, JSON turns
// them into ISO text, and the reviver turns that text back. The pattern is
// strict enough that a note, a name or a meeting id can never be caught by it.
var SHEET_ROWS_TTL   = 600;
var SHEET_ROWS_CHUNK = 30000;
var ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/;

function sheetRowsKey_(name) { return 'rows_' + name.replace(/[^A-Za-z0-9]/g, ''); }

function sheetRowsCached_(name) {
  try {
    var base = sheetRowsKey_(name), cache = CacheService.getScriptCache();
    var n = parseInt(cache.get(base + '_n') || '0', 10);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(base + '_' + i);
    var got = cache.getAll(keys), s = '';
    for (var j = 0; j < n; j++) {
      var part = got[base + '_' + j];
      if (part == null) return null;          // one piece gone, treat the lot as gone
      s += part;
    }
    return JSON.parse(s, function(k, v) {
      return (typeof v === 'string' && ISO_DATE_RE.test(v)) ? new Date(v) : v;
    });
  } catch (e) { return null; }
}

function sheetRowsStore_(name, rows) {
  try {
    var base = sheetRowsKey_(name), cache = CacheService.getScriptCache();
    var s = JSON.stringify(rows), parts = {}, n = 0;
    for (var i = 0; i < s.length; i += SHEET_ROWS_CHUNK) {
      parts[base + '_' + n] = s.slice(i, i + SHEET_ROWS_CHUNK);
      n++;
    }
    parts[base + '_n'] = String(n);
    cache.putAll(parts, SHEET_ROWS_TTL);
  } catch (e) { /* too big, or no cache: the sheet is simply read again next time */ }
}

// Everything that used sheet.getDataRange().getValues() on a read path calls
// this instead. Same array of arrays, same order, same types.
function sheetRows_(name) {
  var hit = sheetRowsCached_(name);
  if (hit) return hit;                      // the copy is free, breaker or not
  if (sheetBreakerTripped_()) sheetBusy_();
  if (requestBudgetSpent_())  sheetBusy_();
  try {
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
    if (!sh) return null;                   // a missing tab is not a stuck sheet
    var rows = sh.getDataRange().getValues();
    sheetRowsStore_(name, rows);
    sheetBreakerClear_();                   // it answered, so let everyone back in
    return rows;
  } catch (e) {
    sheetBreakerTrip_();
    sheetBusy_();
  }
}

// Called after every write, so nobody is served a copy that predates it.
function sheetRowsDropAll_() {
  try {
    var cache = CacheService.getScriptCache(), keys = [];
    [MEETINGS_SHEET, CONDUCTED_SHEET, POSTPONED_SHEET, CANCELLED_SHEET, EMPLOYEE_SHEET].forEach(function(name) {
      var base = sheetRowsKey_(name);
      var n = parseInt(cache.get(base + '_n') || '0', 10);
      keys.push(base + '_n');
      for (var i = 0; i < n; i++) keys.push(base + '_' + i);
    });
    if (keys.length) cache.removeAll(keys);
  } catch (e) {}
}

// Editor helper: is the shared copy doing its job?
function ROWS_status() {
  var cache = CacheService.getScriptCache(), out = [];
  [MEETINGS_SHEET, CONDUCTED_SHEET, POSTPONED_SHEET, CANCELLED_SHEET, EMPLOYEE_SHEET].forEach(function(name) {
    var n = parseInt(cache.get(sheetRowsKey_(name) + '_n') || '0', 10);
    out.push(name + ': ' + (n ? ('cached in ' + n + ' piece(s)') : 'not cached, next read will fill it'));
  });
  Logger.log(out.join(String.fromCharCode(10)));
  return out.join(' | ');
}

function invalidateUser(email, district) {
  var keys = ['emp_' + email,
       'stats_' + email + '_0', 'stats_' + email + '_1',
       'rep_' + email,
       'mymt_' + email, 'allmymt_' + email,
       'mymtg_' + email, 'planmtg_' + email,
       'stateMtg_all', 'docUrlMap', 'meetingZoneMap', 'reportData'];
  // The district and zone lists are shared, not one person's. Clearing only the
  // writer's own keys meant a meeting they had just filed stayed invisible to
  // their district and zone leads until those caches expired by themselves.
  // The employee is read before the keys go, so that lookup is still served warm.
  try {
    var ds = [];
    var d0 = (district || '').toString().trim();
    if (d0) ds.push(d0);
    var me = getEmployeeByEmail(email);
    if (me) {
      if (me.district) ds.push(me.district.toString().trim());
      (me.districts || []).forEach(function(x) { if (x) ds.push(x.toString().trim()); });
    }
    var seen = {};
    ds.forEach(function(x) {
      var k = x.toLowerCase();
      if (!k || seen[k]) return;
      seen[k] = 1;
      keys.push('distMtg_' + k);            // matches getDistrictAllMeetings
      var z = districtToZone_(x);
      if (z) keys.push('zoneMtg_' + z);     // matches getZoneAllMeetings
    });
  } catch (e) { /* the user's own keys are cleared either way */ }
  cDel.apply(null, keys);
  sheetRowsDropAll_();   // the shared sheet copies too, or a write would stay invisible
}

// ------------------------------------------------------------
//  DEMO / PRESENTATION HELPER - Run from GAS editor
//  Sheet mein role change karne ke baad ye run karo
//  Turant cache clear hoga - logout/login ke baad naya role dikhega
// ------------------------------------------------------------
function clearCacheForDemo() {
  var email = 'alok.mohan@educategirls.ngo'; // ← apna email yahan rakho
  invalidateUser(email);
  CacheService.getScriptCache().remove('distMtg_sitapur');
  CacheService.getScriptCache().remove('stateMtg_all');
  Logger.log('✅ Cache cleared for: ' + email + ' - ab logout karke login karo');
}

// ------------------------------------------------------------
//  DEMO ROLE SWITCHERS - Run ONE of these from the GAS editor,
//  then Logout + Login in the app. Role badal jayega + cache
//  clear ho jayega automatically. Sheet manually edit nahi karni.
//  (Email niche DEMO_EMAIL mein set hai)
// ------------------------------------------------------------
var DEMO_EMAIL = 'alok.mohan@educategirls.ngo';

function demoSetField()    { setDemoRole_('Field');    }
function demoSetDistrict() { setDemoRole_('District'); }
function demoSetState()    { setDemoRole_('State');    }

function setDemoRole_(role) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  var data  = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if ((data[i][4] || '').toString().trim().toLowerCase() === DEMO_EMAIL.toLowerCase()) {
      sheet.getRange(i + 1, 6).setValue(role); // col F = Role
      found = true;
      break;
    }
  }
  if (!found) { Logger.log('❌ Email not found: ' + DEMO_EMAIL); return; }

  // Clear caches so the new role + meetings show immediately
  invalidateUser(DEMO_EMAIL.toLowerCase());
  var c = CacheService.getScriptCache();
  ['distMtg_sitapur','stateMtg_all','allEmp'].forEach(function(k){ c.remove(k); });

  Logger.log('✅ Role set to "' + role + '" for ' + DEMO_EMAIL +
             '\n👉 Ab app mein LOGOUT karke LOGIN karo.');
}

// ------------------------------------------------------------
//  GRANT ADDITIONAL DISTRICT CHARGE - Run ONCE from GAS editor
//  Gives a user charge of one or more extra districts (beyond their
//  primary). Writes to the "Additional Districts" column (col H) in
//  Employee_DB and clears their cache so it applies on next login.
//  To reuse for someone else, just edit the two lines below and re-run.
// ------------------------------------------------------------
function setupDualCharge() {
  // Each entry: [identifier, extra district(s)] - identifier can be an EMAIL
  // (contains '@') or the exact employee NAME. extra MUST match the exact
  // district name used in Employee_DB / meeting sheets (e.g. "LAKHIMPUR KHERI",
  // not "LAKHIMPUR"; "FARRUKHABAD"). Only the listed rows are touched - other
  // people's existing charge in col H is left untouched.
  var GRANTS = [
    ['Manvendra Mishra', 'FARRUKHABAD']   // primary HARDOI stays; adds FARRUKHABAD charge
  ];

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) { return '❌ Employee sheet not found: ' + EMPLOYEE_SHEET; }

  var COL_H = 8;   // H = Additional Districts (A=1 … G=7 Zone, H=8)
  if (!sheet.getRange(1, COL_H).getValue()) {
    sheet.getRange(1, COL_H).setValue('Additional Districts');
  }

  var data  = sheet.getDataRange().getValues();
  var cache = CacheService.getScriptCache();
  var out = [];
  GRANTS.forEach(function(g) {
    var id      = (g[0] || '').toString().trim().toLowerCase();
    var byEmail = id.indexOf('@') !== -1;
    var extra   = g[1] || '';
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][4] || '').toString().trim().toLowerCase();
      var rowName  = (data[i][2] || '').toString().trim().toLowerCase();
      if ((byEmail && rowEmail === id) || (!byEmail && rowName === id)) {
        sheet.getRange(i + 1, COL_H).setValue(extra);
        try { cache.remove('emp_' + rowEmail); } catch(e) {}
        var msg = '✅ ' + (data[i][2] || '') + ' (' + (data[i][0] || '') + ' + ' + extra + ')  [cache cleared]';
        Logger.log(msg); out.push(msg);
        return;
      }
    }
    var nf = '❌ Not found in Employee sheet: ' + g[0];
    Logger.log(nf); out.push(nf);
  });
  Logger.log('🔄 Done. Ask these users to LOGOUT and LOGIN again.');
  return out.join('\n') + '\n🔄 Done - user should LOGOUT and LOGIN again.';
}

// ------------------------------------------------------------
//  CHANGE PRIMARY DISTRICT - Run ONCE from GAS editor
//  Corrects a user's home district (col A) in Employee_DB + clears cache.
//  Only the listed rows are touched. District MUST be the exact spelling
//  used across the system (e.g. "BUDAUN", "LAKHIMPUR KHERI").
// ------------------------------------------------------------
function setPrimaryDistrict() {
  // [identifier (email preferred, or exact name), new primary district]
  var CHANGES = [
    ['rahul.kumar3@educategirls.ngo', 'BUDAUN']   // was HARDOI
  ];

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return '❌ Employee sheet not found: ' + EMPLOYEE_SHEET;

  var data  = sheet.getDataRange().getValues();
  var cache = CacheService.getScriptCache();
  var out = [];
  CHANGES.forEach(function(c) {
    var id      = (c[0] || '').toString().trim().toLowerCase();
    var byEmail = id.indexOf('@') !== -1;
    var dist    = c[1] || '';
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][4] || '').toString().trim().toLowerCase();
      var rowName  = (data[i][2] || '').toString().trim().toLowerCase();
      if ((byEmail && rowEmail === id) || (!byEmail && rowName === id)) {
        var old = (data[i][0] || '').toString();
        sheet.getRange(i + 1, 1).setValue(dist);   // col A = District
        try { cache.remove('emp_' + rowEmail); } catch(e) {}
        var msg = '✅ ' + (data[i][2] || '') + ': ' + old + ' → ' + dist + '  [cache cleared]';
        Logger.log(msg); out.push(msg);
        return;
      }
    }
    var nf = '❌ Not found in Employee sheet: ' + c[0];
    Logger.log(nf); out.push(nf);
  });
  return out.join('\n') + '\n🔄 Done - user should LOGOUT and LOGIN again.';
}

// ------------------------------------------------------------
//  NORMALIZE DISTRICTS - Run ONCE from GAS editor
//  Trims + UPPERCASEs the district column in Employee_DB and all
//  4 meeting sheets so spelling/casing is consistent everywhere.
//  Only touches the district column; nothing else is modified.
// ------------------------------------------------------------
function normalizeAllDistricts() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var report = {};
  normalizeDistrictColumn_(ss, EMPLOYEE_SHEET,  0, report); // Employee_DB → col A
  normalizeDistrictColumn_(ss, MEETINGS_SHEET,  1, report); // Plan       → col B
  normalizeDistrictColumn_(ss, CONDUCTED_SHEET, 1, report);
  normalizeDistrictColumn_(ss, POSTPONED_SHEET, 1, report);
  normalizeDistrictColumn_(ss, CANCELLED_SHEET, 1, report);

  // Clear all caches so the cleaned data shows immediately
  try { CacheService.getScriptCache().remove('stateMtg_all'); } catch(e){}
  try { CacheService.getScriptCache().remove('allEmp'); } catch(e){}
  try {
    var ks = ['distMtg_sitapur','distMtg_prayagraj','distMtg_shahjahanpur',
              'distMtg_hardoi','distMtg_gonda','distMtg_fatehpur','distMtg_bahraich'];
    ks.forEach(function(k){ CacheService.getScriptCache().remove(k); });
  } catch(e){}

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function normalizeDistrictColumn_(ss, sheetName, colIdx, report) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) { report[sheetName] = 'SHEET NOT FOUND'; return; }
  var lastRow = sh.getLastRow();
  if (lastRow < 2) { report[sheetName] = 'no data rows'; return; }
  var rng  = sh.getRange(2, colIdx + 1, lastRow - 1, 1);
  var vals = rng.getValues();
  var changed = 0;
  for (var i = 0; i < vals.length; i++) {
    var orig = (vals[i][0] || '').toString();
    var norm = orig.trim().toUpperCase();
    if (norm !== orig) { vals[i][0] = norm; changed++; }
  }
  rng.setValues(vals);
  report[sheetName] = changed + ' of ' + vals.length + ' rows normalized';
}

// ------------------------------------------------------------
//  DIAGNOSTIC - district column audit across all meeting sheets
//  Returns a summary of district values so we can see why a
//  district filter (e.g. SITAPUR) shows fewer meetings than expected.
// ------------------------------------------------------------
// Canonicalize role names so typos/variants map to the 4 system roles.
// "Zonal" / "Zonal Lead" → "Zone"; case-corrects State/District/Field/Zone.
function normalizeRole_(raw) {
  var r = (raw || 'Field').toString().trim();
  var lc = r.toLowerCase();
  if (lc.indexOf('zone') === 0 || lc.indexOf('zonal') === 0) return 'Zone';
  if (lc === 'state')    return 'State';
  if (lc === 'district') return 'District';
  if (lc === 'field')    return 'Field';
  return r;
}

function diagnoseZoneTeam() {
  var EMAIL = 'alok.mohan@educategirls.ngo'; // ← apna email
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  var data = sheet.getDataRange().getValues();
  // Cols: District(0) Block(1) Name(2) Desig(3) Email(4) Role(5) Zone(6)
  var me = null, roleCount = {}, zoneCount = {};
  for (var i = 1; i < data.length; i++) {
    var em = (data[i][4]||'').toString().trim().toLowerCase();
    var role = (data[i][5]||'(blank)').toString().trim();
    var zone = (data[i][6]||'(blank)').toString().trim();
    roleCount[role] = (roleCount[role]||0)+1;
    zoneCount[zone] = (zoneCount[zone]||0)+1;
    if (em === EMAIL.toLowerCase()) {
      me = { district:data[i][0], block:data[i][1], name:data[i][2], desig:data[i][3], email:data[i][4], role:data[i][5], zone:data[i][6] };
    }
  }
  Logger.log('MY ROW: ' + JSON.stringify(me));
  Logger.log('ROLE counts: ' + JSON.stringify(roleCount));
  Logger.log('ZONE counts: ' + JSON.stringify(zoneCount));
  if (me) {
    CacheService.getScriptCache().remove('zoneEmp_' + (me.zone||'').toString().trim().toUpperCase());
    var res = getZoneTeamEmployees(me.zone, me.email);
    Logger.log('getZoneTeamEmployees("' + me.zone + '") → ' + res.length + ' people');
    Logger.log('Names: ' + res.map(function(r){return r.name + ' [' + r.district + '/' + (r._email)+']';}).join('  |  '));
  }
  return me;
}

function diagnoseDistricts() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = [MEETINGS_SHEET, CONDUCTED_SHEET, POSTPONED_SHEET, CANCELLED_SHEET];
  var out = {};
  sheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { out[name] = 'SHEET NOT FOUND'; return; }
    var data = sh.getDataRange().getValues();
    var counts = {};
    for (var i = 1; i < data.length; i++) {
      var d = (data[i][1] || '(blank)').toString().trim();
      counts[d] = (counts[d] || 0) + 1;
    }
    out[name] = counts;
  });
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// ------------------------------------------------------------
//  AUTHORIZE ALL SERVICES - Run this once from GAS editor
//  to grant all required permissions (Spreadsheet, Drive, Mail)
// ------------------------------------------------------------
function authorizeAll() {
  try {
    // 1. Spreadsheet access
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('✅ Spreadsheet: ' + ss.getName());

    // 2. Drive access
    var folder = getRootMeetingsFolder();
    Logger.log('✅ Drive folder: ' + folder.getName() + ' (' + folder.getId() + ')');

    // 3. Mail access
    var quota = MailApp.getRemainingDailyQuota();
    Logger.log('✅ Mail quota remaining: ' + quota);

    // 4. Session / user
    Logger.log('✅ Running as: ' + Session.getEffectiveUser().getEmail());

    Logger.log('🎉 All services authorized successfully!');
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
  }
}

// ------------------------------------------------------------
//  ENTRY POINT
// ------------------------------------------------------------
// ------------------------------------------------------------
//  API HANDLER - called from GitHub Pages frontend via fetch()
// ------------------------------------------------------------
function doPost(e) {
  return apiResponse(e, 'POST');
}

// Admin allowlist - only these emails can call destructive/import actions
var ADMIN_EMAILS = ['gr@educategirls.ngo', 'alok.mohan@educategirls.ngo'];
function isAdmin_(email) {
  email = (email || '').toString().trim().toLowerCase();
  for (var i = 0; i < ADMIN_EMAILS.length; i++) {
    if (ADMIN_EMAILS[i].toLowerCase() === email) return true;
  }
  return false;
}

// ── Maintenance switch ──────────────────────────────────────────────────
// When SpreadsheetApp cannot open the document, every request that needs it
// holds an execution for its full six minutes. Other people's open tabs keep
// sending those requests, so the script never gets a free slot and even work
// that needs no sheet at all cannot run. This turns the sheet-backed actions
// away at the door in a few milliseconds, which lets the queue drain.
//
// Signing in is deliberately still allowed: it reads the copy in Script
// Properties, not the document, so it costs nothing and nobody is locked out.
//
//   MAINT_on()   turn it on      MAINT_off()   turn it off
var MAINT_KEY = 'MAINTENANCE_MODE_2';   // renamed so the old 'on' from the outage cannot linger
var MAINT_ALLOWED = { sendOTP:1, verifyOTP:1, getPlanDistricts:1, getDropdownData:1 };

function maintOn_() {
  try { return PropertiesService.getScriptProperties().getProperty(MAINT_KEY) === 'on'; }
  catch (e) { return false; }
}
function MAINT_on() {
  PropertiesService.getScriptProperties().setProperty(MAINT_KEY, 'on');
  Logger.log('Maintenance ON. Sign-in still works; anything that reads the sheet is turned away at once.');
  return 'on';
}
function MAINT_off() {
  PropertiesService.getScriptProperties().deleteProperty(MAINT_KEY);
  Logger.log('Maintenance OFF. Everything is served normally again.');
  return 'off';
}

function apiResponse(e, method) {
  IS_WEB_REQUEST = true;   // somebody is waiting at a screen for this one
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var result;
  try {
    var body = {}, bodyBroken = false;
    // The front door delivers a POST without its body often enough to be the
    // commonest complaint about saving: the script then sees no date, no name,
    // no purpose and rightly refuses. Nothing is written when that happens, so
    // the honest answer is to say the body never arrived and let the browser
    // send it again, rather than making someone press Save a second time and
    // wonder what they did wrong.
    if (method === 'POST' && (!e.postData || !e.postData.contents)) bodyBroken = true;
    if (method === 'POST' && e.postData && e.postData.contents) {
      // A truncated or malformed body used to fall through as {} and the write
      // went ahead anyway, stamping the session fields onto an otherwise empty
      // row. Refuse instead: a request that did not arrive must not become a
      // half meeting in the sheet.
      try { body = JSON.parse(e.postData.contents); } catch(pe) { body = {}; bodyBroken = true; }
    }
    var token  = (e && e.parameter && e.parameter.token) ? e.parameter.token : '';
    // getDashboardStats / getDistrictReport are public - power the open
    // State Analytics Portal (report.html), which needs no login.
    var PUBLIC = { sendOTP: 1, verifyOTP: 1, loginPassword: 1, getDashboardStats: 1, getDistrictReport: 1, getReportData: 1, getEmployeeMaster: 1 };
    var ADMIN  = { bulkUpdateEmployeeDB: 1, importFromSource: 1, peekSourceSheet: 1 };
    // Writes and heavy requests that the pages send only as a POST. Sign-in is
    // deliberately not in this list: it is left exactly as it is.
    var POST_ONLY = { saveMeeting:1, conductMeeting:1, postponeMeeting:1, cancelMeeting:1, uploadGovtMom:1,
                      askMeetings:1, transcribeVoice:1, sendMeetingFeedback:1 };

    // Turned away at the door while the document is unreachable, so the queue
    // can empty. Signing in is on the allowed list and keeps working.
    if (maintOn_() && !MAINT_ALLOWED[action]) {
      result = { success:false, message:'The system is being repaired right now. Please try again in a little while. Nothing you have saved is affected.' };
    } else if (bodyBroken) {
      result = { success:false, message:'BODY_MISSING' };
    } else if (method !== 'POST' && POST_ONLY[action]) {
      // The pages only ever send these as a POST. Google's front door
      // sometimes answers a POST with a redirect the browser follows as a GET:
      // the action and token survive in the URL, the body does not. Run as-is
      // it wrote half rows (a conduct with no meeting id) or answered as if
      // nothing had been asked. BODY_MISSING makes the page send it again.
      result = { success:false, message:'BODY_MISSING' };
    } else if (PUBLIC[action]) {
      // ── No auth required ──────────────────────────────────────
      if      (action === 'sendOTP')           result = sendOTP(e.parameter.email || '');
      else if (action === 'verifyOTP')         result = verifyOTP(e.parameter.email || '', e.parameter.otp || '');
      else if (action === 'loginPassword')      result = loginPassword(e.parameter.email || '', (body.password || ''));   // body only: a password must never reach a log through the URL
      else if (action === 'getDashboardStats') result = getDashboardStats(e.parameter.email || '', e.parameter.all === '1');
      else if (action === 'getDistrictReport') result = getDistrictReport(e.parameter.district || '');
      else if (action === 'getReportData')     result = getReportData();
      else if (action === 'getEmployeeMaster') result = getEmployeeMaster();
    } else {
      // ── Auth required: identity comes from the session token, ──
      //    NOT from client-supplied params (prevents spoofing)
      var session = getSession(token);
      if (!session) {
        result = { success: false, message: 'AUTH_REQUIRED' };
      } else if (ADMIN[action] && !isAdmin_(session.email)) {
        result = { success: false, message: 'ADMIN_ONLY' };
      } else {
        // Sliding expiry: refresh 1-hour TTL on every authenticated call
        try { CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(session), 3600); } catch(se) {}
        var role = (session.role || '').toString();

        if      (action === 'setPassword')          result = setPassword(session.email, (body.password || ''));
        else if (action === 'getDropdownData')      result = getDropdownData(session.email);
        else if (action === 'getMyMeetings')        result = getMyMeetings(session.email);
        else if (action === 'getAllMyMeetings')     result = getAllMyMeetings(session.email);
        else if (action === 'getMonthlyReport')     result = getMonthlyReport(session, e.parameter.month || '');
        else if (action === 'getMeetingPrep')       result = getMeetingPrep(session, e.parameter.meetingId || '');
        else if (action === 'getDistrictEmployees') result = getDistrictEmployees(resolveActiveDistrict_(session, e.parameter.district), session.email);
        else if (action === 'getAllEmployees')      result = getAllEmployees(session.email);
        else if (action === 'getZoneTeamEmployees') result = getZoneTeamEmployees(session.zone, session.email);
        else if (action === 'getPlanDistricts')     result = getPlanDistricts(session.role, session.zone, (session.districts && session.districts.length) ? session.districts : [session.district]);
        else if (action === 'getDistrictAllMeetings') {
          // Active district: own/charge districts for District role; any for State
          result = getDistrictAllMeetings(resolveActiveDistrict_(session, e.parameter.district));
        }
        else if (action === 'getStateAllMeetings')  result = (role === 'State')
                                                       ? getStateAllMeetings()
                                                       : { success: false, message: 'FORBIDDEN' };
        else if (action === 'getZoneAllMeetings') {
          // Zone role locked to own zone; State may query any zone
          var zn = (role === 'State') ? (e.parameter.zone || session.zone) : session.zone;
          result = (role === 'Zone' || role === 'State')
                     ? getZoneAllMeetings(zn)
                     : { success: false, message: 'FORBIDDEN' };
        }
        // These three only ever come as a POST, and the page never sends them
        // with the main field empty. Google's front door sometimes turns a
        // POST into a GET on a redirect: the action and token survive in the
        // URL, the body does not, and bodyBroken only looks at POSTs. So an
        // empty question, recording or feedback means the request lost its
        // body on the way. BODY_MISSING makes the page send it again, where
        // answering "please type a question" to someone who just did would not.
        // (saveMeeting's "the meeting details did not arrive" is the same thing.)
        else if (action === 'askMeetings')                 result = body.question ? askMeetings(session, body.question, body.history || [])
                                                                                  : { success:false, message:'BODY_MISSING' };
        else if (action === 'transcribeVoice')             result = body.audio ? transcribeVoice(session, body.audio, body.mime || '')
                                                                               : { success:false, message:'BODY_MISSING' };
        else if (action === 'sendMeetingFeedback')         result = (body.meetingId && body.text) ? sendMeetingFeedback(session, body.meetingId, body.text)
                                                                                               : { success:false, message:'BODY_MISSING' };
        else if (action === 'getDashboardStats')    result = getDashboardStats(session.email, e.parameter.all === '1', resolveActiveDistrict_(session, e.parameter.district));
        else if (action === 'getDistrictReport') {
          result = getDistrictReport(resolveActiveDistrict_(session, e.parameter.district));
        }
        else if (action === 'getAllReports')        result = getAllReports(session.email);
        else if (action === 'deleteMeeting')        result = deleteMeeting(e.parameter.meetingId || '', session.email);
        else if (action === 'saveMeeting') {
          // Stamp identity from session - fixes attribution + blank district.
          // District = the active/charge district requested by the client, validated against
          // the user's authorized list (so a dual-charge lead files under the right district).
          body.email = session.email; body.employeeName = session.name;
          body.district = resolveActiveDistrict_(session, body.district); body.designation = session.designation; body.block = session.block;
          result = saveMeeting(body);
        }
        else if (action === 'conductMeeting')   { body.email = session.email; result = conductMeeting(body); }
        else if (action === 'uploadGovtMom')    { result = uploadGovtMom(body, session); }
        else if (action === 'postponeMeeting')  { body.email = session.email; result = postponeMeeting(body); }
        else if (action === 'cancelMeeting')    { body.email = session.email; result = cancelMeeting(body); }
        else if (action === 'updateMeetingStatus') result = updateMeetingStatus(body.meetingId || '', body);
        else if (action === 'clearMyCache')     result = clearMyCache(session.email);
        else if (action === 'bulkUpdateEmployeeDB') result = bulkUpdateEmployeeDB(body.rows || []);
        else if (action === 'peekSourceSheet')  result = peekSourceSheet(e.parameter.sourceId || '', e.parameter.sheetIndex || '0');
        else if (action === 'importFromSource') result = importFromSource(e.parameter.sourceId || '', e.parameter.sheetIndex || '0');
        else                                    result = { success: false, message: 'Unknown action: ' + action };
      }
    }
  } catch(err) {
    result = { success: false, message: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
//  TEST FUNCTION - Run this once from GAS editor to authorize MailApp
// ------------------------------------------------------------
function authorizeMailApp() {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'EG MMS - MailApp Authorization Successful',
    body: 'MailApp is now authorized. Colleague notifications will work.'
  });
  Logger.log('MailApp authorized successfully.');
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action) return apiResponse(e, 'GET');

  var page  = (e && e.parameter && e.parameter.page)  ? e.parameter.page  : 'login';
  var token = (e && e.parameter && e.parameter.token) ? e.parameter.token : '';

  var execUrl = ScriptApp.getService().getUrl();

  if (page === 'dashboard') {
    var userData = token ? getSession(token) : null;

    if (!userData) {
      var loginTmpl = HtmlService.createTemplateFromFile('Index');
      loginTmpl.execUrl = execUrl;
      return loginTmpl.evaluate()
        .setTitle('EG Meeting Management System')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Refresh session so active users stay logged in (reset 1-hour TTL)
    CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(userData), 3600);

    var tmpl = HtmlService.createTemplateFromFile('MeetingForm');
    tmpl.execUrl     = execUrl;
    tmpl.sessionJson = JSON.stringify({
      token:       token,
      email:       userData.email,
      name:        userData.name,
      district:    userData.district,
      districts:   userData.districts || [userData.district],
      block:       userData.block,
      designation: userData.designation,
      role:        userData.role
    });

    return tmpl.evaluate()
      .setTitle('EG Meeting Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var loginTmpl = HtmlService.createTemplateFromFile('Index');
  loginTmpl.execUrl = execUrl;
  return loginTmpl.evaluate()
    .setTitle('EG Meeting Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include helper for Stylesheet.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ------------------------------------------------------------
//  RUN THIS ONCE FROM EDITOR TO AUTHORIZE PERMISSIONS
// ------------------------------------------------------------
function authorizeApp() {
  SpreadsheetApp.openById(SPREADSHEET_ID).getName();
  getRootMeetingsFolder().getName();
  MailApp.getRemainingDailyQuota();
  Logger.log('Authorization successful.');
}

// ------------------------------------------------------------
//  RUN THIS ONCE to create Drive folder under gr account
//  and get the new DRIVE_ROOT_ID to paste in Code.gs
// ------------------------------------------------------------
function setupDriveFolder() {
  var folderName = 'EG-GR-Meetings';
  var root = DriveApp.getRootFolder();
  var it = root.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : root.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('=== COPY THIS FOLDER ID ===');
  Logger.log('DRIVE_ROOT_ID = "' + folder.getId() + '"');
  Logger.log('Folder URL: ' + folder.getUrl());
}

// ------------------------------------------------------------
//  OTP - SEND
// ------------------------------------------------------------
// Editor helper: why is no OTP arriving? Three separate things can stop it and
// they look identical from the login page, so this reports all three at once.
// Sends nothing. Run LOGIN_debug('alok.mohan@educategirls.ngo').
// Editor helper: is Sheets refusing everything, or only our document?
// Tries a brand new throwaway spreadsheet FIRST, because if our document hangs
// it eats the whole six minutes and nothing after it would ever run. The
// throwaway is deleted again straight away.
// Editor helper: how big is each tab, without reading a single cell.
// getMaxRows/getLastRow are metadata, so this answers in a second even on a
// document that getDataRange().getValues() cannot get through. "empty tail" is
// the part that holds nothing and still gets dragged along on every read.
// ── Reading the document through the Sheets REST API ────────────────────
// SpreadsheetApp.openById builds the whole document in memory before it will
// answer, and on the night of 17 Sep 2026 it stopped being able to do that for
// this document at all: a one line script with no other code hung on it just as
// the app did, while a brand new sheet opened in a second. The REST API is a
// different road to the same data and asks only for the range named, so it can
// come back when the other cannot.
function sheetsApiGet_(range) {
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + SPREADSHEET_ID +
            '/values/' + encodeURIComponent(range) + '?majorDimension=ROWS';
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    // Hiding this cost an hour once. Whatever the API objects to, say it.
    Logger.log('Sheets API refused ' + range + ': HTTP ' + res.getResponseCode() +
               String.fromCharCode(10) + res.getContentText().slice(0, 400));
    return null;
  }
  var j = JSON.parse(res.getContentText());
  return j.values || [];
}

// The API's own error body, because "refused" told us nothing: a wrong tab name
// and a missing permission look identical from outside and need opposite fixes.
function sheetsApiRaw_(path) {
  var res = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets/' + SPREADSHEET_ID + path, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  return { code: res.getResponseCode(), body: res.getContentText() };
}

// ── Reading and writing the document without SpreadsheetApp ──────────────
// Switched on with SHEETS_API_on() and off again with SHEETS_API_off(), so the
// normal path comes straight back the moment openById works again. Nothing here
// changes what is stored or how; it is the same rows through a different door.
var SHEETS_API_KEY = 'USE_SHEETS_API';

function sheetsApiOn_() {
  try { return PropertiesService.getScriptProperties().getProperty(SHEETS_API_KEY) === 'on'; }
  catch (e) { return false; }
}
function SHEETS_API_on()  { PropertiesService.getScriptProperties().setProperty(SHEETS_API_KEY, 'on');
                            Logger.log('Reads and writes now go through the Sheets API.'); return 'on'; }
function SHEETS_API_off() { PropertiesService.getScriptProperties().deleteProperty(SHEETS_API_KEY);
                            Logger.log('Back to the normal SpreadsheetApp path.'); return 'off'; }

// Every row of a tab, the way getDataRange().getValues() would have given them.
// Rows the API returns are ragged, so they are padded out to a fixed width and
// the callers that index by column keep working unchanged.
function apiValues_(tabName, width) {
  var rows = sheetsApiGet_("'" + tabName + "'");
  if (rows === null) return null;
  var w = width || 40, out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] || [];
    while (r.length < w) r.push('');
    out.push(r);
  }
  return out;
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// appendRow's replacement. The API appends after the last filled row of the
// tab, which is what appendRow does, so ordering and the row number match.
function apiAppend_(tabName, row) {
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + SPREADSHEET_ID +
            '/values/' + encodeURIComponent("'" + tabName + "'!A1") +
            ':append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ values: [row] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return null;
  // "'Plan Meetings'!A413:X413" -> 413, so callers that then write one more cell
  // into that row know where it landed.
  var m = /![A-Z]+(\d+)/.exec((JSON.parse(res.getContentText()).updates || {}).updatedRange || '');
  return m ? parseInt(m[1], 10) : 0;
}

// Writing one cell, for the follow-up columns the save path fills in after.
function apiSetCell_(tabName, row, col, value) {
  var a1 = "'" + tabName + "'!" + colLetter_(col) + row;
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + SPREADSHEET_ID +
            '/values/' + encodeURIComponent(a1) + '?valueInputOption=USER_ENTERED';
  var res = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ values: [[value]] }),
    muteHttpExceptions: true
  });
  return res.getResponseCode() === 200;
}

// Editor helper: prove the three operations work before anything depends on them.
function SHEETS_API_selftest() {
  var out = [];
  var t = new Date().getTime();
  var rows = apiValues_(MEETINGS_SHEET, 24);
  out.push('read Plan Meetings : ' + (rows === null ? 'refused' : rows.length + ' rows') +
           ' in ' + (new Date().getTime() - t) + ' ms');
  t = new Date().getTime();
  var emp = apiValues_(EMPLOYEE_SHEET, 8);
  out.push('read Employee_DB   : ' + (emp === null ? 'refused' : emp.length + ' rows') +
           ' in ' + (new Date().getTime() - t) + ' ms');
  out.push('');
  out.push('Nothing was written. If both read, the API path is ready.');
  Logger.log(out.join(String.fromCharCode(10)));
  return out.join(' || ');
}

// Editor helper: ask the document what its tabs are actually called. This is a
// metadata call, so it comes back even while the document will not open.
function SHEET_apiTabs() {
  var t = new Date().getTime();
  var r = sheetsApiRaw_('?fields=properties.title,sheets.properties(title,gridProperties)');
  var out = ['HTTP ' + r.code + ' in ' + (new Date().getTime() - t) + ' ms'];
  if (r.code !== 200) {
    out.push(r.body.slice(0, 600));
  } else {
    var j = JSON.parse(r.body);
    out.push('Document: ' + (j.properties && j.properties.title));
    out.push('');
    (j.sheets || []).forEach(function(s) {
      var g = s.properties.gridProperties || {};
      out.push('  "' + s.properties.title + '"   ' + (g.rowCount || '?') + ' rows x ' + (g.columnCount || '?') + ' cols');
    });
  }
  Logger.log(out.join(String.fromCharCode(10)));
  return r.code;
}

// Editor helper: does the other road work? Reads five rows and reports.
function SHEET_apiTest() {
  var out = [];
  ['Employee', 'Plan Meetings'].forEach(function(name) {
    var t = new Date().getTime();
    try {
      var rows = sheetsApiGet_("'" + name + "'!A1:F5");
      out.push(name + ': ' + (rows === null ? 'refused' : rows.length + ' rows') +
               ' in ' + (new Date().getTime() - t) + ' ms');
      if (rows && rows.length) out.push('    first row: ' + rows[0].join(' | '));
    } catch (e) {
      out.push(name + ': FAILED after ' + (new Date().getTime() - t) + ' ms - ' + e.message);
    }
  });
  out.push('');
  out.push('Tab names must match exactly. If one says refused, tell me the real tab name.');
  Logger.log(out.join(String.fromCharCode(10)));
  return out.join(' || ');
}

// Fills the sign-in copy for EVERYONE, through the REST API, without
// SpreadsheetApp ever being asked to open the document. This is what makes
// other people able to log in while the document itself is unreachable.
function EMP_refreshViaApi(tabName) {
  tabName = tabName || EMPLOYEE_SHEET;   // the Run button passes nothing, so this default matters
  var rows = sheetsApiGet_("'" + tabName + "'");
  if (rows === null) {
    Logger.log('The API refused, or the tab name is wrong. Tab tried: ' + tabName);
    return 0;
  }
  var map = {}, n = 0;
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i] || [];
    var em = (r[4] || '').toString().trim().toLowerCase();
    if (!em) continue;
    var primaryDist = (r[0] || '').toString().trim();
    var districts = [primaryDist];
    (r[7] || '').toString().split(/[,;]/).forEach(function(x) {
      var d = x.toString().trim();
      if (d && districts.map(function(z){ return z.toLowerCase(); }).indexOf(d.toLowerCase()) === -1) districts.push(d);
    });
    map[em] = {
      district:    primaryDist,
      districts:   districts,
      block:       (r[1] || '').toString().trim(),
      name:        (r[2] || '').toString().trim(),
      designation: (r[3] || '').toString().trim(),
      email:       (r[4] || '').toString().trim(),
      role:        normalizeRole_(r[5]),
      zone:        (r[6] || '').toString().trim()
    };
    n++;
  }
  if (!n) { Logger.log('Read the tab but found no email addresses in column E. Tab: ' + tabName); return 0; }
  var chunks = empMirrorWrite_(map);
  Logger.log('Sign-in copy filled from the API: ' + n + ' people, ' + chunks + ' chunk(s).' +
             String.fromCharCode(10) + 'Everyone can sign in now.');
  return n;
}

function SHEET_sizes() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var out = [], worst = 0;
  ss.getSheets().forEach(function(sh) {
    var maxR = sh.getMaxRows(), lastR = sh.getLastRow();
    var maxC = sh.getMaxColumns(), lastC = sh.getLastColumn();
    var tail = maxR - lastR;
    if (tail > worst) worst = tail;
    out.push(pad_(sh.getName(), 26) +
             ' rows ' + pad_(lastR + ' of ' + maxR, 16) +
             ' cols ' + pad_(lastC + ' of ' + maxC, 10) +
             (tail > 1000 ? '   <-- ' + tail + ' empty rows to delete' : ''));
  });
  out.push('');
  out.push(worst > 1000 ? 'Delete the empty rows on the tabs marked above.'
                        : 'Every tab looks trim. The size is not the problem any more.');
  Logger.log(out.join(String.fromCharCode(10)));
  return worst;
}
function pad_(s, n) { s = String(s); while (s.length < n) s += ' '; return s; }

// Editor helper: removes the empty tail from every tab, using only row counts,
// never a data read. Preview first, then pass 'DELETE'. It never touches a row
// at or above getLastRow(), so nothing that holds anything can be removed.
function SHEET_trimTail(confirm) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = [], total = 0;
  ss.getSheets().forEach(function(sh) {
    var maxR = sh.getMaxRows(), lastR = sh.getLastRow();
    var keep = Math.max(lastR, 1) + 50;        // a little room to append into
    if (maxR <= keep) { log.push(sh.getName() + ': already trim (' + maxR + ' rows)'); return; }
    var from = keep + 1, count = maxR - keep;
    total += count;
    if (confirm === 'DELETE') {
      sh.deleteRows(from, count);
      log.push(sh.getName() + ': deleted rows ' + from + '-' + maxR + '  (' + count + ' empty rows)');
    } else {
      log.push(sh.getName() + ': would delete rows ' + from + '-' + maxR + '  (' + count + ' empty rows)');
    }
  });
  log.push('');
  log.push(confirm === 'DELETE' ? ('Removed ' + total + ' empty rows in total.')
                                : ('Preview only. Run SHEET_trimTail("DELETE") to remove ' + total + ' empty rows.'));
  Logger.log(log.join(String.fromCharCode(10)));
  return total;
}

function SHEET_probe() {
  var out = [];

  var t1 = new Date().getTime();
  try {
    var tmp = SpreadsheetApp.create('EG probe ' + t1);
    tmp.getSheets()[0].getRange(1, 1).setValue('ok');
    var id = tmp.getId();
    out.push('A BRAND NEW SHEET : worked in ' + (new Date().getTime() - t1) + ' ms');
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e0) {}
  } catch (e1) {
    out.push('A BRAND NEW SHEET : FAILED after ' + (new Date().getTime() - t1) + ' ms - ' + e1.message);
  }

  var t2 = new Date().getTime();
  try {
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EMPLOYEE_SHEET);
    out.push('OUR DOCUMENT      : opened, ' + (sh ? sh.getLastRow() + ' rows' : 'sheet missing') +
             ', in ' + (new Date().getTime() - t2) + ' ms');
  } catch (e2) {
    out.push('OUR DOCUMENT      : FAILED after ' + (new Date().getTime() - t2) + ' ms - ' + e2.message);
  }

  Logger.log(out.join(String.fromCharCode(10)));
  return out.join(' || ');
}

function LOGIN_debug(email) {
  var out = [];
  email = (email || 'alok.mohan@educategirls.ngo').toString().trim().toLowerCase();

  // 1. Mail quota. Once this hits zero MailApp throws and no code can go out,
  //    however healthy everything else is. It resets on a rolling 24 hours.
  var quota = 'could not read';
  try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { quota = 'ERROR: ' + e.message; }
  out.push('Mail left today: ' + quota);

  // 2. Is this person still in the employee sheet? sendOTP refuses first if not.
  var emp = null;
  try { emp = getEmployeeByEmail(email); } catch (e2) { out.push('employee lookup ERROR: ' + e2.message); }
  out.push('Employee found: ' + (emp ? (emp.name + '  role ' + emp.role + '  district ' + emp.district) : 'NO - sendOTP would refuse'));

  // 3. What is running on a timer. A job that has started failing every hour is
  //    the usual reason a script runs out of room to answer anything.
  try {
    var ts = ScriptApp.getProjectTriggers();
    out.push('Triggers installed: ' + ts.length);
    ts.forEach(function(t) { out.push('     ' + t.getHandlerFunction()); });
  } catch (e3) { out.push('triggers ERROR: ' + e3.message); }

  // 4. Is a code already sitting in the cache for this person, unused?
  try {
    var held = CacheService.getScriptCache().get('OTP_' + email);
    out.push('Code waiting in cache: ' + (held ? 'yes, and it is still valid' : 'no'));
  } catch (e4) { out.push('cache ERROR: ' + e4.message); }

  Logger.log(out.join(String.fromCharCode(10)));
  return out.join(' | ');
}

// ── Password sign-in ────────────────────────────────────────────────────
// A code by email needs two slow round trips through a front door that has been
// taking half a minute, and every fresh code kills the one before it, so a
// person who presses the button again while waiting is left holding a dead
// code. A password needs one trip and nothing expires. The code stays, for the
// first sign-in and for anyone who has forgotten theirs.
//
// The password itself is never stored. Each person gets their own random salt,
// and only the salt and the SHA-256 of salt+password are kept. Nobody reading
// the stored values, this script's author included, can work back to it.
var PW_MIN_LEN   = 8;
var PW_MAX_TRIES = 5;          // per email, per window
var PW_WINDOW_S  = 900;        // 15 minutes
var PW_OBVIOUS = {
  'password':1, 'password1':1, '12345678':1, '123456789':1, '1234567890':1,
  'qwertyui':1, 'educategirls':1, 'educate123':1, 'abcd1234':1, 'admin123':1
};

function pwHash_(salt, password) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}
function pwKey_(email) { return 'PW_' + email.trim().toLowerCase(); }

function pwStored_(email) {
  try { return PropertiesService.getScriptProperties().getProperty(pwKey_(email)); }
  catch (e) { return null; }
}

// Called by someone who has already proved who they are with a code, so the
// session token is the authority here, not anything the browser says.
function setPassword(email, password) {
  password = (password || '').toString();
  if (password.length < PW_MIN_LEN) {
    return { success: false, message: 'Please choose at least ' + PW_MIN_LEN + ' characters.' };
  }
  if (PW_OBVIOUS[password.toLowerCase()]) {
    return { success: false, message: 'That one is too easy to guess. Please choose another.' };
  }
  if (password.toLowerCase() === email.split('@')[0].toLowerCase()) {
    return { success: false, message: 'Please do not use your own name as the password.' };
  }
  try {
    var salt = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty(pwKey_(email), salt + '$' + pwHash_(salt, password));
    CacheService.getScriptCache().remove('PWTRY_' + email.trim().toLowerCase());
    return { success: true, message: 'Password saved. You can sign in with it from now on.' };
  } catch (e) {
    return { success: false, message: 'Could not save the password. Please try again.' };
  }
}

// Public, so it is rate limited: anyone at all can reach this URL, and without
// a limit a machine could work through passwords at its leisure.
function loginPassword(email, password) {
  email = (email || '').toString().trim().toLowerCase();
  password = (password || '').toString();
  if (!email || !password) return { success: false, message: 'Please enter your email and password.' };

  var cache = CacheService.getScriptCache();
  var tkey  = 'PWTRY_' + email;
  var tries = parseInt(cache.get(tkey) || '0', 10);
  if (tries >= PW_MAX_TRIES) {
    return { success: false, message: 'Too many attempts. Please wait fifteen minutes, or sign in with a code instead.' };
  }

  var stored = pwStored_(email);
  var emp    = getEmployeeByEmail(email);
  // The same answer whether the email is unknown, has no password yet, or the
  // password is wrong, so this cannot be used to find out who is registered.
  function refuse() {
    cache.put(tkey, String(tries + 1), PW_WINDOW_S);
    return { success: false, message: 'Email or password is incorrect. If this is your first time, ask for a code instead.' };
  }
  if (!stored || !emp) return refuse();

  var parts = stored.split('$');
  if (parts.length !== 2 || pwHash_(parts[0], password) !== parts[1]) return refuse();

  cache.remove(tkey);
  return sessionFor_(emp);
}

// The session half of verifyOTP, so a password sign-in and a code sign-in hand
// the browser exactly the same thing and nothing downstream can tell them apart.
function sessionFor_(emp) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify({
    email:       emp.email,
    name:        emp.name,
    district:    emp.district,
    districts:   emp.districts || [emp.district],
    block:       emp.block,
    designation: emp.designation,
    role:        emp.role,
    zone:        emp.zone || '',
    loginTime:   new Date().toISOString()
  }), 3600);
  return {
    success:     true,
    token:       token,
    role:        emp.role,
    name:        emp.name,
    district:    emp.district,
    districts:   emp.districts || [emp.district],
    block:       emp.block,
    designation: emp.designation,
    zone:        emp.zone || '',
    email:       emp.email
  };
}

// Editor helper: who has set one, and clearing one for someone who is stuck.
function PW_status() {
  var all = PropertiesService.getScriptProperties().getProperties(), n = 0, who = [];
  for (var k in all) { if (k.indexOf('PW_') === 0) { n++; who.push(k.slice(3)); } }
  Logger.log(n + ' people have a password set' + String.fromCharCode(10) + who.sort().join(String.fromCharCode(10)));
  return n;
}
function PW_clear(email) {
  if (!email) { Logger.log('Pass an email.'); return 0; }
  PropertiesService.getScriptProperties().deleteProperty(pwKey_(email));
  Logger.log('Password cleared for ' + email + '. They can set a new one after a code.');
  return 1;
}

function sendOTP(email) {
  email = email.trim().toLowerCase();

  // Only allow office domain
  var domain = email.split('@')[1] || '';
  if (domain !== ALLOWED_DOMAIN) {
    return { success: false, message: 'Only @' + ALLOWED_DOMAIN + ' email addresses are allowed.' };
  }

  var employee = getEmployeeByEmail(email);
  if (!employee) {
    return { success: false, message: 'Your email is not registered in the system. Please contact Admin.' };
  }

  var otp = Math.floor(100000 + Math.random() * 900000).toString();
  CacheService.getScriptCache().put('OTP_' + email, otp, OTP_EXPIRY_SEC);

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'EG Meeting Management System - Login OTP',
      body: 'Dear ' + employee.name + ',\n\n' +
            'Your OTP for EG Meeting Management System is: ' + otp + '\n\n' +
            'This OTP is valid for 10 minutes. Do not share it with anyone.\n\n' +
            'Educate Girls Meeting Management System'
    });
    return { success: true, message: 'OTP sent to: ' + email, name: employee.name };
  } catch (err) {
    return { success: false, message: 'Failed to send OTP: ' + err.message };
  }
}

// ------------------------------------------------------------
//  OTP - VERIFY
// ------------------------------------------------------------
function verifyOTP(email, otp) {
  email = email.trim().toLowerCase();
  otp   = otp.trim();

  var cache     = CacheService.getScriptCache();
  var storedOTP = cache.get('OTP_' + email);

  // The same code, offered again within a few minutes, gets the same answer.
  // The front door is slow enough that a verify can reach the script, spend the
  // code and hand back a session that never arrives, leaving someone holding a
  // code the server has already used and being told it is invalid. This is not
  // a second login, it is the same one answered again.
  var againKey = 'OTPOK_' + email + '_' + otp;
  var again    = cache.get(againKey);
  if (again) { try { return JSON.parse(again); } catch (ae) {} }

  if (!storedOTP) {
    return { success: false, message: 'OTP has expired. Please request a new OTP.' };
  }
  if (storedOTP !== otp) {
    return { success: false, message: 'Invalid OTP. Please try again.' };
  }

  cache.remove('OTP_' + email);

  var emp = getEmployeeByEmail(email);
  if (!emp) {
    return { success: false, message: 'Employee record not found.' };
  }

  // Create session token
  var token = Utilities.getUuid();
  var session = JSON.stringify({
    email:       emp.email,
    name:        emp.name,
    district:    emp.district,
    districts:   emp.districts || [emp.district],   // all districts under this user's charge
    block:       emp.block,
    designation: emp.designation,
    role:        emp.role,
    zone:        emp.zone || '',
    loginTime:   new Date().toISOString()
  });
  cache.put('SESSION_' + token, session, 3600); // 1 hour

  var out = {
    success:     true,
    token:       token,
    role:        emp.role,
    name:        emp.name,
    district:    emp.district,
    districts:   emp.districts || [emp.district],   // client uses this to show the district switcher
    block:       emp.block,
    designation: emp.designation,
    zone:        emp.zone || '',
    email:       emp.email
  };
  // Held briefly, so a reply lost on the way back can be asked for again.
  try { cache.put(againKey, JSON.stringify(out), 300); } catch (pe) {}
  return out;
}

// ------------------------------------------------------------
//  SESSION - GET
// ------------------------------------------------------------
function getSession(token) {
  var data = CacheService.getScriptCache().get('SESSION_' + token);
  if (!data) return null;
  return JSON.parse(data);
}

// ------------------------------------------------------------
//  ACTIVE DISTRICT RESOLVER
//  Given a session and a client-requested district, return the
//  district to actually operate on. Prevents a user from querying
//  a district they have no charge over.
//   • State role  → any district (whole-state visibility, unchanged)
//   • Other roles → the requested district only if it is in the
//     user's authorized list (primary + additional charge);
//     otherwise falls back to their primary district.
// ------------------------------------------------------------
function resolveActiveDistrict_(session, requested) {
  var req  = (requested || '').toString().trim();
  // State-level-only districts (e.g. LUCKNOW): ANY role may file a meeting here.
  if (req && STATE_EXTRA_DISTRICTS.some(function(d){ return d.toLowerCase() === req.toLowerCase(); })) return req;
  var role = (session && session.role || '').toString();
  if (role === 'State') return req || (session && session.district) || '';
  if (role === 'Zone') {
    // Zone leads may file a meeting under any district within their zone
    if (req && districtToZone_(req) === findZoneKey_(session && session.zone)) return req;
    return (session && session.district) || '';
  }

  var allowed = (session && session.districts && session.districts.length)
                  ? session.districts
                  : [session && session.district];
  if (req) {
    for (var i = 0; i < allowed.length; i++) {
      if ((allowed[i] || '').toString().trim().toLowerCase() === req.toLowerCase()) {
        return allowed[i];   // authorized → honour the request
      }
    }
  }
  return (session && session.district) || '';   // default to primary
}

// ------------------------------------------------------------
//  PLAN DISTRICTS - districts a State/Zone user may file a meeting
//  under. State → all districts; Zone → districts in their zone.
// ------------------------------------------------------------
function getPlanDistricts(role, zone, ownDistricts) {
  role = (role || '').toString().trim();
  var base;
  if (role === 'Zone') {
    // The admin districts in this lead's zone (from the fixed mapping)
    var zkey = findZoneKey_(zone);
    base = zkey ? ZONE_DISTRICTS[zkey].slice() : [];
  } else if (role === 'State') {
    // All admin districts across every zone
    base = [];
    for (var z in ZONE_DISTRICTS) base = base.concat(ZONE_DISTRICTS[z]);
  } else {
    // District / Field → their own (primary + any charge) district(s)
    base = (ownDistricts || []).slice();
  }
  // Everyone also gets the state-level-only districts (e.g. LUCKNOW)
  base = base.concat(STATE_EXTRA_DISTRICTS);
  // De-dup (case-insensitive) + sort
  var seen = {}, out = [];
  base.forEach(function(d){
    var k = (d || '').toString().trim();
    if (k && !seen[k.toUpperCase()]) { seen[k.toUpperCase()] = 1; out.push(k); }
  });
  return out.sort();
}

// ------------------------------------------------------------
//  ACCESS CHECK - re-verify employee still active in sheet
//  Returns emp object if active, null if removed/not found
// ------------------------------------------------------------
function checkAccess(email) {
  return getEmployeeByEmail(email.trim().toLowerCase());
}

// ------------------------------------------------------------
//  DROPDOWN DATA - Stakeholder Type (hardcoded) + Meeting Purpose (sheet)
// ------------------------------------------------------------
// ── Meeting purposes, with a copy that does not depend on Sheets ─────────
var PURPOSES_KEY = 'PURPOSES_MIRROR';

function purposesRead_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('EG_PURPOSES');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  var copy = null;
  try { copy = JSON.parse(PropertiesService.getScriptProperties().getProperty(PURPOSES_KEY) || 'null'); } catch (e2) {}
  if (copy && copy.length) {
    try { cache.put('EG_PURPOSES', JSON.stringify(copy), 600); } catch (e3) {}
    return copy;
  }

  var list = purposesReadFromSheet_();
  if (!list) return [];                    // Sheets will not answer and there is no copy yet
  purposesWrite_(list);
  return list;
}

function purposesReadFromSheet_() {
  try {
    var ws = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Meeting Purpose');
    if (!ws) return null;
    var d = ws.getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++) { if (d[i][0]) out.push(d[i][0].toString().trim()); }
    return out;
  } catch (e) { return null; }
}

function purposesWrite_(list) {
  try {
    PropertiesService.getScriptProperties().setProperty(PURPOSES_KEY, JSON.stringify(list));
    CacheService.getScriptCache().put('EG_PURPOSES', JSON.stringify(list), 600);
    return list.length;
  } catch (e) { return 0; }
}

// Editor helper: refresh the purpose copy from the Sheet. calendarJob does this
// hourly; run it yourself after editing the Meeting Purpose sheet.
function PURPOSE_refresh() {
  var list = purposesReadFromSheet_();
  if (!list) { Logger.log('Could not read the Meeting Purpose sheet - Sheets is not answering. Try again in a minute.'); return 0; }
  purposesWrite_(list);
  Logger.log('Purpose copy refreshed: ' + list.length + ' purposes.' + String.fromCharCode(10) + list.join(', '));
  return list.length;
}

// Editor helper for the case this was written in: Sheets is refusing the
// document, so the copy cannot be taken from it. This puts the purposes people
// have actually been choosing all year into the copy, so the form works tonight.
// The hourly refresh replaces it with the real sheet the moment Sheets answers.
function PURPOSE_seedFallback() {
  var list = ['Enrollment','Learning','Introductory Meeting','MPR Submission','Review Meeting',
              'School Liasioning','Courtesy Meeting','Retention','DTF','Feedback letter',
              'Invitation','Task Force Meeting'];
  purposesWrite_(list);
  Logger.log('Seeded ' + list.length + ' purposes as a stand-in:' + String.fromCharCode(10) + list.join(', ') +
             String.fromCharCode(10) + 'The real sheet replaces this automatically once Sheets answers.');
  return list.length;
}

function getDropdownData(email) {
  // Re-verify access on every page load
  if (email && !getEmployeeByEmail(email.trim().toLowerCase())) {
    return { error: 'ACCESS_REVOKED' };
  }

  var stakeholders = [
    'ACS', 'DGSE', 'UIC', 'SPD', 'ASPD', 'JD',
    'Development Partner Cell',
    'BSA', 'DC- Gender', 'DC-Training', 'DC- MIS',
    'DC- Community', 'DC-IED',
    'District Collector', 'CDO', 'DIET Principal',
    'ABSA', 'ARP', 'Head Teacher', 'Teacher', 'Other'
  ];

  // Cache, then the copy in Script Properties, and only then the Sheet. When
  // Sheets stopped answering on 17 Sep 2026 this read hung for six minutes and
  // took the whole Plan Meeting form down with it, including the posts, which
  // are not even stored in a sheet.
  var purposes = purposesRead_();

  // Blocks per district, so the plan form can offer the block a block-level
  // official actually sits in. Built from the sign-in copy rather than by
  // reopening the employee sheet.
  // The copy already carries each person's district and block, and an empty
  // block list only means the optional block dropdown has nothing to offer.
  var blocksByDistrict = {};
  try {
    var _mir = empMirrorRead_() || {};
    var _emps = [];
    for (var _k in _mir) _emps.push(_mir[_k]);
    _emps.forEach(function(e) {
      var d = (e.district || '').toString().trim();
      var b = (e.block || '').toString().trim();
      if (!d || !b) return;
      var list = blocksByDistrict[d] || (blocksByDistrict[d] = []);
      if (list.indexOf(b) === -1) list.push(b);
    });
    Object.keys(blocksByDistrict).forEach(function(d){ blocksByDistrict[d].sort(); });
  } catch (be) { blocksByDistrict = {}; }

  // When this office, and this person, were last met and about what. It rides
  // along with the form rather than being fetched when a post is picked: the
  // front door costs 3 to 17 seconds even for a request that runs no code, and
  // this has to appear while someone is still filling the form in.
  // One record per district + post + person, holding only the most recent
  // meeting, which keeps the whole state to a few hundred lines. Wrapped,
  // because a plan form that cannot be filled is far worse than one with no
  // history line, which is the lesson of 17 Sep.
  var metBefore = [];
  try {
    var _cd = sheetRows_(CONDUCTED_SHEET) || [];
    var _seen = {};
    for (var _i = 1; _i < _cd.length; _i++) {
      if (!_cd[_i][0]) continue;
      var _nm = (_cd[_i][9]  || '').toString().trim();
      var _pt = (_cd[_i][10] || '').toString().trim();
      if (!_nm && !_pt) continue;
      var _dt = _cd[_i][13];
      var _ts = (_dt instanceof Date) ? _dt.getTime() : Date.parse(_dt);
      // The whole line is "last met on <date>", so a row without one has
      // nothing to contribute and would render an empty date.
      if (!_ts || isNaN(_ts)) continue;
      var _key = normDist_((_cd[_i][1] || '').toString()) + '|' +
                 _prepNorm_(_pt) + '|' + _prepNorm_(_nm);
      var _prev = _seen[_key];
      if (_prev && _prev.ts >= _ts) continue;
      var _rec = { dist: (_cd[_i][1] || '').toString(), post: _pt, name: _nm,
                   block: (_cd[_i][COL_CON_SKBLOCK - 1] || '').toString(),
                   purpose: (_cd[_i][11] || '').toString(),
                   date: fmtDateVal(_dt), ts: _ts };
      if (_prev) { for (var _f in _rec) _prev[_f] = _rec[_f]; }
      else { _seen[_key] = _rec; metBefore.push(_rec); }
    }
    metBefore.sort(function(a, b){ return b.ts - a.ts; });
    if (metBefore.length > 600) metBefore.length = 600;
  } catch (mbe) { metBefore = []; }

  return { stakeholders: stakeholders, purposes: purposes,
           blocksByDistrict: blocksByDistrict, metBefore: metBefore };
}

// ------------------------------------------------------------
//  FEEDBACK ON A CONDUCTED MEETING
// ------------------------------------------------------------
// A State or Zone lead reading a conducted meeting can send the officer who
// filed it whatever they have to say about it: a question, a correction, a
// request, or that it was good work. It goes out as one email and the reply
// comes back to them, so the exchange lives in both mailboxes.
//
// NOTHING IS STORED. That is the decision, not an omission: the email already
// sits in both people's inboxes, searchable, so a copy in a sheet would be a
// second worse record to keep in step with the first. It also means no Comments
// tab, no unread counts, and no new screen to build or maintain.
//
// The officer's address is looked up here from the meeting id rather than sent
// up by the browser, so the page never has to hold anyone's email.
function sendMeetingFeedback(session, meetingId, text) {
  meetingId = (meetingId || '').toString().trim();
  text      = (text || '').toString().trim();
  if (!meetingId) return { success:false, message:'No meeting id' };
  if (text.length < 5) return { success:false, message:'Please write your feedback first.' };
  if (text.length > 4000) text = text.substring(0, 4000);

  var role = (session.role || '').toString();
  if (role !== 'State' && role !== 'Zone') return { success:false, message:'FORBIDDEN' };

  var cd = sheetRows_(CONDUCTED_SHEET) || [];
  var m = null;
  for (var i = 1; i < cd.length; i++) {
    if ((cd[i][0] || '').toString().trim() !== meetingId) continue;
    m = { district:(cd[i][1]||'').toString(),  officer:(cd[i][2]||'').toString(),
          email:(cd[i][4]||'').toString().trim(),
          name:(cd[i][9]||'').toString(),      post:(cd[i][10]||'').toString(),
          purpose:(cd[i][11]||'').toString(),  date:fmtDateVal(cd[i][13]),
          notes:(cd[i][15]||'').toString(),    mom:(cd[i][17]||'').toString() };
    break;
  }
  if (!m) return { success:false, message:'That meeting was not found.' };
  if (!m.email) return { success:false, message:'No email is recorded against that meeting, so nobody can be written to.' };

  // A zone lead may only write about their own zone. State may write about any.
  // Both sides go through the canonical key first. The Zone column is typed by
  // hand, so "UP Zone 1" and "up zone-1" are the same zone as far as every
  // other screen is concerned, and comparing the raw string would lock out a
  // lead whose Zone Meetings tab works perfectly well.
  if (role === 'Zone') {
    var mine = findZoneKey_((session.zone || '').toString());
    if (!mine || districtToZone_(m.district) !== mine) return { success:false, message:'FORBIDDEN' };
  }

  var from = (session.name || session.email || 'A colleague').toString();
  var subj = 'Feedback on your meeting with ' + (m.name || m.post || 'an official') +
             ' (' + m.date + ')';

  var body =
    '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#111827;line-height:1.6;">' +
    '<p>' + _emailEsc(from) + ' has sent you feedback on a meeting you recorded.</p>' +
    '<div style="border-left:3px solid #7B1010;background:#FAFAFA;padding:10px 14px;margin:14px 0;white-space:pre-wrap;">' +
      _emailEsc(text) +
    '</div>' +
    '<p style="margin:16px 0 6px;font-weight:600;">The meeting</p>' +
    '<table style="font-size:13px;color:#374151;border-collapse:collapse;">' +
      '<tr><td style="padding:2px 12px 2px 0;color:#6B7280;">Date</td><td>' + _emailEsc(m.date) + '</td></tr>' +
      '<tr><td style="padding:2px 12px 2px 0;color:#6B7280;">District</td><td>' + _emailEsc(m.district) + '</td></tr>' +
      '<tr><td style="padding:2px 12px 2px 0;color:#6B7280;">Met</td><td>' + _emailEsc(m.name) +
        (m.post ? ', ' + _emailEsc(m.post) : '') + '</td></tr>' +
      '<tr><td style="padding:2px 12px 2px 0;color:#6B7280;">Purpose</td><td>' + _emailEsc(m.purpose) + '</td></tr>' +
      '<tr><td style="padding:2px 12px 2px 0;color:#6B7280;">Meeting ID</td><td>' + _emailEsc(meetingId) + '</td></tr>' +
    '</table>' +
    (m.notes ? '<p style="margin:16px 0 6px;font-weight:600;">What you wrote</p>' +
               '<div style="font-size:13px;color:#374151;white-space:pre-wrap;background:#F9FAFB;padding:10px 14px;border-radius:6px;">' +
               _emailEsc(m.notes) + '</div>' : '') +
    (m.mom ? '<p style="margin:14px 0 0;"><a href="' + _emailEsc(m.mom) + '">Open the MoM</a></p>' : '') +
    '<p style="margin-top:18px;color:#6B7280;font-size:13px;">Just reply to this email to respond. ' +
      _emailEsc(from) + ' is on it and will get your reply.</p>' +
    '</div>';

  try {
    MailApp.sendEmail({
      to: m.email,
      cc: session.email,          // so Reply All reaches the asker
      replyTo: session.email,     // and so does a plain Reply
      subject: subj,
      htmlBody: body,
      name: 'EG-MMS'
    });
  } catch (e) {
    return { success:false, message:'The email could not be sent: ' + e.message };
  }
  return { success:true, sentTo:m.officer || m.email };
}

// ------------------------------------------------------------
//  MEETINGS - SAVE
// ------------------------------------------------------------
function saveMeeting(data) {
  try {
    // Re-verify employee is still active before saving
    if (!getEmployeeByEmail((data.email || '').trim().toLowerCase())) {
      return { success: false, message: 'ACCESS_REVOKED' };
    }

    // Second guard, in case a request ever reaches here without its form data.
    // The identity fields are stamped from the session, so without this a lost
    // body writes a row carrying only who and when, and no meeting at all.
    var req = [];
    if (!(data.meetingDate  || '').toString().trim()) req.push('date');
    if (!(data.adhikariName || '').toString().trim()) req.push('stakeholder name');
    if (!(data.adhikariPost || '').toString().trim()) req.push('stakeholder post');
    if (!(data.purpose      || '').toString().trim()) req.push('purpose');
    if (req.length) {
      // The page will not send a plan without these, so their absence means the
      // body was lost on the way. BODY_MISSING has the page send it again on
      // its own instead of asking the officer to press Save a second time.
      // Safe to repeat: nothing was written, and findRecentPlan_ below stops a
      // plan that did land from being filed twice.
      return { success: false, message: 'BODY_MISSING', missing: req };
    }

    if (sheetBreakerTripped_()) return { success: false, message: SHEET_BUSY_MSG };
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MEETINGS_SHEET);
    if (!sheet) return { success: false, message: 'Meetings sheet not found.' };

    // ── The same plan twice ───────────────────────────────────
    // Apps Script's front door is slow and sometimes answers with an error page
    // even though the script has already written the row, so an officer who
    // never sees a confirmation presses Save again. Instead of letting that
    // become a second meeting, hand back the one already in the sheet.
    var dupId = findRecentPlan_(sheet, data);
    if (dupId) return { success: true, meetingId: dupId, duplicate: true };

    var now   = new Date();
    var mtgId = 'MTG-' + now.getFullYear() +
                ('0'+(now.getMonth()+1)).slice(-2) +
                ('0'+now.getDate()).slice(-2) + '-' +
                ('0'+now.getHours()).slice(-2) + ('0'+now.getMinutes()).slice(-2) + ('0'+now.getSeconds()).slice(-2);

    // ── Upload meeting documents to Drive (optional) ──────────
    var docFolderUrl = '';
    try {
      if (data.documents && data.documents.length > 0) {
        var droot = getRootMeetingsFolder();
        var ddist = getOrCreateFolder(droot, data.district || 'General');
        var dmtg  = getOrCreateFolder(ddist, mtgId);
        var ddoc  = getOrCreateFolder(dmtg, 'Documents');
        ddoc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        docFolderUrl = ddoc.getUrl();
        data.documents.forEach(function(doc, idx) {
          var decoded = Utilities.base64Decode(doc.data);
          var blob = Utilities.newBlob(decoded, doc.type || 'application/octet-stream',
                       doc.name || (mtgId + '_doc' + (idx+1)));
          var f = ddoc.createFile(blob);
          f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        });
      }
    } catch(docErr) { docFolderUrl = ''; }

    var row = [
      mtgId,                    // A  Meeting ID
      data.district     || '',  // B  District
      data.employeeName || '',  // C  Employee Name
      data.designation  || '',  // D  Post
      data.email        || '',  // E  Email
      data.meetingDate  || '',  // F  Meeting Date
      data.meetingTime  || '',  // G  Meeting Time
      data.duration     || '',  // H  Duration
      data.meetingType  || '',  // I  Meeting Type
      data.adhikariName || '',  // J  Stakeholder Name
      data.adhikariPost || '',  // K  Stakeholder Post
      data.purpose      || '',  // L  Meeting Purpose
      data.agenda       || '',  // M  Meeting Agenda
      'Planned',                // N  Status
      '',                       // O  Start Time (filled on update)
      '',                       // P  End Time   (filled on update)
      '',                       // Q  Reason     (filled on update)
      data.colleagueName|| '',  // R  Colleague Name
      data.colleaguePost   || '',  // S  Colleague Post
      now.toLocaleString('en-IN'), // T  Submitted At
      data.parentMeetingId || '',  // U  Parent Meeting ID (for follow-ups)
      docFolderUrl                 // V  Documents folder URL
    ];

    sheet.appendRow(row);
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 7).setNumberFormat('@'); // keep Meeting Time as text
    // Only touched when a block was actually picked, so the usual save costs
    // neither the header read nor the extra write.
    if (data.adhikariBlock) {
      if (!sheet.getRange(1, COL_PLAN_SKBLOCK).getValue()) sheet.getRange(1, COL_PLAN_SKBLOCK).setValue('Stakeholder Block');
      sheet.getRange(newRow, COL_PLAN_SKBLOCK).setValue(data.adhikariBlock);
    }

    // ── Colleague email notification ──────────────────────────
    if (data.colleagueName && data.colleagueName.trim()) {
      try { sendColleagueNotification(data, mtgId); } catch(mailErr) { /* don't fail save if mail fails */ }
    }

    invalidateUser((data.email || '').trim().toLowerCase(), data.district);
    return { success: true, meetingId: mtgId, docUrl: docFolderUrl };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// The moment encoded in a meeting id (MTG-YYYYMMDD-HHMMSS); 0 if it will not parse
function mtgIdTime_(id) {
  var m = /^MTG-(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})(\d{2})$/.exec((id || '').toString().trim());
  if (!m) return 0;
  return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime();
}

// A plan by the same officer, for the same official, the same date and the same
// purpose, saved minutes ago: that is the same save arriving twice, not a second
// meeting. Returns its meeting id, or '' when this really is a new plan.
var DUP_PLAN_WINDOW_MS = 10 * 60 * 1000;
function findRecentPlan_(sheet, data) {
  try {
    var last = sheet.getLastRow();
    if (last < 2) return '';
    var from = Math.max(2, last - 200);   // a resend is always among the newest rows
    var rows = sheet.getRange(from, 1, last - from + 1, 14).getValues();
    function key(email, date, name, purpose) {
      return [(email || '').toString().trim().toLowerCase(),
              fmtDateVal(date),
              (name || '').toString().trim().toLowerCase(),
              (purpose || '').toString().trim().toLowerCase()].join('|');
    }
    var mine   = key(data.email, data.meetingDate, data.adhikariName, data.purpose);
    var cutoff = Date.now() - DUP_PLAN_WINDOW_MS;
    for (var i = rows.length - 1; i >= 0; i--) {
      var r = rows[i];
      if ((r[13] || '').toString().trim() !== 'Planned') continue;      // N  Status
      if (mtgIdTime_(r[0]) < cutoff) continue;                          // A  Meeting ID
      if (key(r[4], r[5], r[9], r[11]) === mine) return (r[0] || '').toString().trim();
    }
  } catch (e) { /* a guard that fails must never block a genuine save */ }
  return '';
}

// ------------------------------------------------------------
//  MEETINGS - GET (for logged-in employee)
// ------------------------------------------------------------
function getMyMeetings(email) {
  try {
    // Reading the whole plan sheet to pull out one officer's handful of rows
    // is the slow part of opening Manage Meetings.
    var _ck = 'planmtg_' + (email || '').trim().toLowerCase();   // mymtg_ belongs to getAllMyMeetings
    var _hit = cGet(_ck);
    if (_hit) return _hit;

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MEETINGS_SHEET);
    if (!sheet) return [];

    var sheetData = (sheetRows_(MEETINGS_SHEET) || []);
    var meetings  = [];
    var tz        = Session.getScriptTimeZone();
    for (var i = 1; i < sheetData.length; i++) {
      var rowEmail = sheetData[i][4] ? sheetData[i][4].toString().trim().toLowerCase() : '';
      if (rowEmail === email.trim().toLowerCase()) {
        var rawDate = sheetData[i][5];
        meetings.push({
          meetingId:    (sheetData[i][0]  || '').toString(),
          district:     (sheetData[i][1]  || '').toString(),
          employeeName: (sheetData[i][2]  || '').toString(),
          date:         fmtDateVal(rawDate),
          meetingTime:  fmtTimeVal(sheetData[i][6]),
          duration:     (sheetData[i][7]  || '').toString(),  // H
          type:         (sheetData[i][8]  || '').toString(),  // I
          adhikariName: (sheetData[i][9]  || '').toString(),  // J
          adhikariPost: (sheetData[i][10] || '').toString(),  // K
          adhikariBlock: (sheetData[i][COL_PLAN_SKBLOCK-1] || '').toString(),  // X
          purpose:      (sheetData[i][11] || '').toString(),  // L
          agenda:       (sheetData[i][12] || '').toString(),  // M
          status:       (sheetData[i][13] || '').toString(),  // N
          startTime:    (sheetData[i][14] || '').toString(),  // O
          endTime:      (sheetData[i][15] || '').toString(),  // P
          reason:       (sheetData[i][16] || '').toString(),  // Q
          colleagueName:(sheetData[i][17] || '').toString(),  // R
          colleaguePost:(sheetData[i][18] || '').toString(),  // S
          parentMeetingId: (sheetData[i][20] || '').toString(), // U
          docUrl:       (sheetData[i][21] || '').toString()   // V  Documents folder
        });
      }
    }
    cPut(_ck, meetings, 600);   // 10 min; every write path clears it anyway
    return meetings;
  } catch (err) {
    return [];
  }
}

// ------------------------------------------------------------
//  DISTRICT EMPLOYEES - for colleague picker
// ------------------------------------------------------------
function getDistrictEmployees(district, currentEmail) {
  var distKey = 'distEmp_' + district.trim().toLowerCase();
  var cur     = currentEmail.trim().toLowerCase();
  var cached  = cGet(distKey);
  if (cached) return cached.filter(function(r){ return r._email !== cur; });

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var all  = [];
  var distL = district.trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
    var e = data[i][4] ? data[i][4].toString().trim().toLowerCase() : '';
    if (d === distL) {
      all.push({
        name:        (data[i][2] || '').toString().trim(),
        designation: (data[i][3] || '').toString().trim(),
        district:    (data[i][0] || '').toString().trim(),
        block:       (data[i][1] || '').toString().trim(),
        _email:      e
      });
    }
  }
  cPut(distKey, all, C_TTL_DROP);
  return all.filter(function(r){ return r._email !== cur; });
}

// ------------------------------------------------------------
//  ALL EMPLOYEES - for colleague picker (no district filter)
// ------------------------------------------------------------
function getAllEmployees(currentEmail) {
  var cur    = currentEmail.trim().toLowerCase();
  var cached = cGet('allEmp');
  if (cached) return cached.filter(function(r){ return r._email !== cur; });

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return [];
  var data  = sheet.getDataRange().getValues();
  var all   = [];
  for (var i = 1; i < data.length; i++) {
    var emp = data[i][4] ? data[i][4].toString().trim().toLowerCase() : '';
    if (!emp) continue;
    all.push({
      name:        (data[i][2] || '').toString().trim(),
      designation: (data[i][3] || '').toString().trim(),
      district:    (data[i][0] || '').toString().trim(),
      block:       (data[i][1] || '').toString().trim(),
      _email:      emp
    });
  }
  all.sort(function(a, b) { return a.name.localeCompare(b.name); });
  cPut('allEmp', all, C_TTL_DROP);
  return all.filter(function(r){ return r._email !== cur; });
}

// ------------------------------------------------------------
//  ZONE TEAM EMPLOYEES - for colleague picker (Zone role)
//  All employees in the user's zone + all State-team members
// ------------------------------------------------------------
function getZoneTeamEmployees(zone, currentEmail) {
  var zkey   = findZoneKey_(zone);
  var cur    = (currentEmail || '').trim().toLowerCase();
  var cacheKey = 'zoneEmp_' + zkey;
  var cached = cGet(cacheKey);
  if (cached) return cached.filter(function(r){ return r._email !== cur; });

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return [];
  var data  = sheet.getDataRange().getValues();
  // Cols: District(0), Block(1), Name(2), Designation(3), Email(4), Role(5), Zone(6)
  var all = [];
  for (var i = 1; i < data.length; i++) {
    var emp = data[i][4] ? data[i][4].toString().trim().toLowerCase() : '';
    if (!emp) continue;
    // Zone membership = the employee's (admin) district's zone, or their own
    // Zone-column value (covers the zone lead whose district is blank).
    var empZone = districtToZone_(data[i][0]) || findZoneKey_(data[i][6]);
    var role    = (data[i][5] || '').toString().trim().toLowerCase();
    if ((zkey && empZone === zkey) || role === 'state') {
      all.push({
        name:        (data[i][2] || '').toString().trim(),
        designation: (data[i][3] || '').toString().trim(),
        district:    (data[i][0] || '').toString().trim(),
        block:       (data[i][1] || '').toString().trim(),
        _email:      emp
      });
    }
  }
  all.sort(function(a, b) { return a.name.localeCompare(b.name); });
  cPut(cacheKey, all, C_TTL_DROP);
  return all.filter(function(r){ return r._email !== cur; });
}

// ------------------------------------------------------------
//  MEETING - UPDATE STATUS (from Manage Meetings)
// ------------------------------------------------------------
function updateMeetingStatus(meetingId, updateData) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MEETINGS_SHEET);
    if (!sheet) return { success: false, message: 'Sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === meetingId) {
        var reason = updateData.reason || '';
        if (updateData.postponedDate) {
          reason = (reason ? reason + ' | ' : '') + 'New Date: ' + updateData.postponedDate;
        }
        sheet.getRange(i + 1, 14).setValue(updateData.status    || ''); // N Status
        sheet.getRange(i + 1, 15).setValue(updateData.startTime || ''); // O Start
        sheet.getRange(i + 1, 16).setValue(updateData.endTime   || ''); // P End
        sheet.getRange(i + 1, 17).setValue(reason);                     // Q Reason
        return { success: true };
      }
    }
    return { success: false, message: 'Meeting ID not found.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  TIME HELPER - Sheets stores time as Dec-30-1899 Date objects.
//  Historical timezone offset for Asia/Kolkata is NOT +5:30,
//  so Utilities.formatDate gives wrong hour. Use UTC directly.
// ------------------------------------------------------------
function fmtTimeVal(t) {
  if (!(t instanceof Date)) return (t || '').toString();
  var h = t.getUTCHours(), mn = t.getUTCMinutes();
  var ap = h >= 12 ? 'PM' : 'AM';
  return (h % 12 || 12) + ':' + (mn < 10 ? '0' : '') + mn + ' ' + ap;
}

// Formats a date cell value as "15 Apr 2026"
// Handles: Date object | "YYYY-MM-DD" | "DD-MM-YYYY" | already formatted string
function fmtDateVal(d) {
  if (!d) return '';
  var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  if (d instanceof Date) {
    return d.getDate() + ' ' + MN[d.getMonth()] + ' ' + d.getFullYear();
  }
  var s = d.toString().trim();
  // YYYY-MM-DD → "15 Apr 2026"
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p = s.split('-');
    return parseInt(p[2]) + ' ' + MN[parseInt(p[1])-1] + ' ' + p[0];
  }
  // DD-MM-YYYY → "15 Apr 2026"
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) {
    var p2 = s.split('-');
    return parseInt(p2[0]) + ' ' + MN[parseInt(p2[1])-1] + ' ' + p2[2];
  }
  return s; // already formatted or unknown - return as-is
}

// ------------------------------------------------------------
//  DRIVE - get or create folder by name under parent
// ------------------------------------------------------------
function getOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// Auto-find or create root meetings folder in script owner's Drive
function getRootMeetingsFolder() {
  var name = 'EG-GR-Meetings';
  var root = DriveApp.getRootFolder();
  var folder = getOrCreateFolder(root, name);
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return folder;
}

// ------------------------------------------------------------
//  CONDUCT MEETING - saves to Conducted sheet, Drive, MoM
// ------------------------------------------------------------
// Does this read like a record of a real meeting, or like someone filling the
// box to get past it? Deterministic on purpose: it runs before anything is
// written, and before the browser sends anything at all, so a poor note costs
// nobody a round trip and, above all, never reaches the sheet. Returns '' when
// the note is fine, or a short reason when it is not.
var NOTE_PLACEHOLDERS = {
  test:1, testing:1, tests:1, abc:1, xyz:1, asdf:1, qwerty:1, sample:1, dummy:1,
  na:1, none:1, nil:1, nothing:1, ok:1, okay:1, done:1, yes:1, no:1, good:1,
  check:1, checking:1, demo:1, blah:1, kuch:1, nahi:1, koi:1
};
// Turned off with NOTE_gate_off() and back on with NOTE_gate_on(). It was
// asked for in the afternoon and got in the way the same night, so it is a
// switch rather than a decision taken once.
function noteGateOn_() {
  try { return PropertiesService.getScriptProperties().getProperty('NOTE_GATE') !== 'off'; }
  catch (e) { return true; }
}
function NOTE_gate_off() { PropertiesService.getScriptProperties().setProperty('NOTE_GATE', 'off');
                           Logger.log('Note check OFF. Whatever is written will save.'); return 'off'; }
function NOTE_gate_on()  { PropertiesService.getScriptProperties().deleteProperty('NOTE_GATE');
                           Logger.log('Note check ON again.'); return 'on'; }

function noteLooksFake_(text) {
  if (!noteGateOn_()) return '';
  var t = (text || '').toString().toLowerCase();
  // The form's own labels are not the officer's words
  t = t.replace(/discussed:/g, ' ').replace(/official said:/g, ' ').replace(/next step:/g, ' ');
  var words = t.split(/\s+/)
    .map(function(w) { return w.replace(/[.,;:!?()\[\]'"\/\_-]+/g, ''); })
    .filter(function(w) { return w.length > 1; });
  if (words.length < 4) return 'it is only a few words';

  var count = {}, uniq = 0, top = 0;
  words.forEach(function(w) {
    if (!count[w]) { count[w] = 0; uniq++; }
    count[w]++;
    if (count[w] > top) top = count[w];
  });
  if (uniq < 4) return 'it is the same few words over and over';
  if (top >= 4 && (top / words.length) > 0.45) return 'one word is repeated over and over';

  var real = 0;
  for (var w2 in count) { if (!NOTE_PLACEHOLDERS[w2]) real++; }
  if (real < 3) return 'it is filler words rather than what happened';
  return '';
}

function conductMeeting(payload) {
  try {
    if (sheetBreakerTripped_()) return { success: false, message: SHEET_BUSY_MSG };
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tz  = Session.getScriptTimeZone();
    var now = new Date();

    // The note quality check was removed on 17 Sep 2026 at Alok's request: it
    // was refusing real conducts and costing more than it caught. Judging the
    // notes is back to being a training matter, not the software's job.

    // A note reused word for word from an earlier meeting tells us nothing
    // about this one, and no length rule catches it: a pasted template can run
    // to several hundred characters. Checked here rather than in the browser so
    // it cannot be skipped. Compares only this officer's own notes.
    try {
      var incoming = (payload.keyPoints || '').toString().trim().toLowerCase().replace(/\s+/g, ' ');
      if (incoming.length >= 40) {
        var cSh = ss.getSheetByName(CONDUCTED_SHEET);
        if (cSh) {
          var cdAll = cSh.getDataRange().getValues();
          var mine = (payload.email || '').toString().trim().toLowerCase();
          for (var dI = cdAll.length - 1; dI >= 1 && dI > cdAll.length - 200; dI--) {
            if ((cdAll[dI][4] || '').toString().trim().toLowerCase() !== mine) continue;
            var prev = (cdAll[dI][15] || '').toString().trim().toLowerCase().replace(/\s+/g, ' ');
            if (prev && prev === incoming) {
              return { success:false, message:'DUPLICATE_NOTE' };
            }
          }
        }
      }
    } catch (dupErr) { /* never block a genuine conduct because this check failed */ }

    // ── Already recorded ──────────────────────────────────────
    // One meeting, one conducted record. When the front door was slow the
    // browser gave up, said it had failed and let the officer save again, so
    // the same meeting landed in the sheet twice and was counted twice in
    // every report. Checked before the photos are uploaded, so a second
    // attempt does not leave stray files in Drive either.
    try {
      var cDone = ss.getSheetByName(CONDUCTED_SHEET);
      if (cDone && cDone.getLastRow() > 1) {
        var doneIds = cDone.getRange(2, 1, cDone.getLastRow() - 1, 1).getValues();
        var wantId  = (payload.meetingId || '').toString().trim();
        for (var aI = doneIds.length - 1; aI >= 0; aI--) {
          if ((doneIds[aI][0] || '').toString().trim() === wantId) {
            return { success: false, message: 'ALREADY_CONDUCTED' };
          }
        }
      }
    } catch (acErr) { /* never block a genuine conduct because this check failed */ }

    // 1. Find row in Plan Meetings (we'll delete it after saving)
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    var momUrl = '', photoFolderUrl = '';
    var planRowIdx = -1;
    if (planSheet) {
      var pd = planSheet.getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        if ((pd[i][0] || '').toString() === payload.meetingId) {
          planRowIdx = i;
          break;
        }
      }
    }

    // 2. Save photos to Drive - wrapped in try-catch so sheet save always happens
    var photoError = '';
    try {
      if (payload.photos && payload.photos.length > 0) {
        var root  = getRootMeetingsFolder();
        var distF = getOrCreateFolder(root, payload.district || 'General');
        var mtgF  = getOrCreateFolder(distF, payload.meetingId);
        mtgF.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoFolderUrl = mtgF.getUrl();
        payload.photos.forEach(function(p, idx) {
          var ext  = (p.type || 'image/jpeg').split('/')[1] || 'jpg';
          var decoded = Utilities.base64Decode(p.data);
          var blob = Utilities.newBlob(decoded, p.type || 'image/jpeg',
                       payload.meetingId + '_' + (idx+1) + '.' + ext);
          var f = mtgF.createFile(blob);
          f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        });
      }
    } catch(de) { photoFolderUrl = ''; photoError = de.message || 'Drive error'; }

    // 3. Create follow-up meeting first (so followUpId goes into MoM)
    var followUpId = '';
    if (payload.followUp && payload.followUp.date) {
      try {
        var fuNow2 = new Date();
        followUpId = 'MTG-' + fuNow2.getFullYear() +
                     ('0'+(fuNow2.getMonth()+1)).slice(-2) +
                     ('0'+fuNow2.getDate()).slice(-2) + '-' +
                     ('0'+fuNow2.getHours()).slice(-2) + ('0'+fuNow2.getMinutes()).slice(-2) + ('0'+fuNow2.getSeconds()).slice(-2);
        var fuPlanSheet = ss.getSheetByName(MEETINGS_SHEET);
        if (fuPlanSheet) {
          fuPlanSheet.appendRow([
            followUpId, payload.district||'', payload.employeeName||'',
            payload.designation||'', payload.email||'',
            payload.followUp.date, payload.followUp.time||'',
            payload.duration||'', payload.meetingType||'',
            payload.adhikariName||'', payload.adhikariPost||'',
            payload.purpose||'', '',
            'Follow-up', '', '', '',
            payload.colleagueName||'', payload.colleaguePost||'',
            fuNow2.toLocaleString('en-IN'), payload.meetingId
          ]);
          fuPlanSheet.getRange(fuPlanSheet.getLastRow(), 7).setNumberFormat('@'); // keep time as text
        }
      } catch(fe) { followUpId = ''; }
    }

    // 4. Create MoM Google Doc (includes follow-up info)
    payload.followUpId = followUpId;
    // Writing the document is the slowest thing this request does, four to six
    // seconds of Google Docs calls, so the officer gets to say whether this
    // meeting needs one. The column simply stays empty when it does not.
    if (!payload.skipMom) {
      try { momUrl = createMoMDoc(payload, photoFolderUrl); } catch(e) { momUrl = ''; }
    }

    // 5. Save to Conducted Meetings sheet
    var cSheet = ss.getSheetByName(CONDUCTED_SHEET);
    if (!cSheet) {
      cSheet = ss.insertSheet(CONDUCTED_SHEET);
      var ch = ['Meeting ID','District','Employee Name','Post','Email',
                'Original Date','Original Time','Duration','Meeting Type',
                'Stakeholder Name','Stakeholder Post','Purpose','Agenda',
                'Conduct Date','Conduct Time','Key Points',
                'Photos Folder','MoM Doc','Colleague Name','Colleague Post','Conducted At','Govt MoM'];
      cSheet.appendRow(ch);
      cSheet.getRange(1,1,1,ch.length).setBackground('#166534').setFontColor('#fff').setFontWeight('bold');
      cSheet.setFrozenRows(1);
    }
    cSheet.appendRow([
      payload.meetingId,    payload.district,      payload.employeeName,
      payload.designation,  payload.email,
      payload.originalDate, payload.originalTime,  payload.duration,      payload.meetingType,
      payload.adhikariName, payload.adhikariPost,  payload.purpose,       payload.agenda,
      payload.conductDate,  payload.conductTime,   payload.keyPoints,
      photoFolderUrl,       momUrl,
      payload.colleagueName || '', payload.colleaguePost || '',
      now.toLocaleString('en-IN')
    ]);

    // force time columns to text so Sheets doesn't reparse them
    var clr = cSheet.getLastRow();
    cSheet.getRange(clr, 7).setNumberFormat('@');  // G Original Time
    cSheet.getRange(clr, 15).setNumberFormat('@'); // O Conduct Time

    // Carry the official's block over from the plan row rather than trusting
    // the client to send it back unchanged.
    try {
      if (!cSheet.getRange(1, COL_CON_SKBLOCK).getValue()) cSheet.getRange(1, COL_CON_SKBLOCK).setValue('Stakeholder Block');
      var skBlock = (payload.adhikariBlock || '').toString().trim();
      if (!skBlock && planSheet && planRowIdx > 0) {
        skBlock = (planSheet.getRange(planRowIdx + 1, COL_PLAN_SKBLOCK).getValue() || '').toString().trim();
      }
      if (skBlock) cSheet.getRange(clr, COL_CON_SKBLOCK).setValue(skBlock);
      if (!cSheet.getRange(1, COL_CON_OUTCOME).getValue()) cSheet.getRange(1, COL_CON_OUTCOME).setValue('Outcome');
      var oc = (payload.outcome || '').toString().trim();
      if (oc) cSheet.getRange(clr, COL_CON_OUTCOME).setValue(oc);
    } catch (sbErr) { /* both are optional, never fail the conduct over them */ }

    // 5. Update status in Plan Meetings to "Conducted" - NEVER delete, keeps master ledger intact for dashboard reporting
    if (planSheet && planRowIdx > 0) {
      planSheet.getRange(planRowIdx + 1, 14).setValue('Conducted');
    }

    // 6. Send MoM email to colleague
    if (payload.colleagueName && payload.colleagueName.trim()) {
      try { sendMOMNotification(payload, momUrl, photoFolderUrl, followUpId); } catch(mailErr) { /* don't fail conduct if mail fails */ }
    }

    invalidateUser((payload.email || '').trim().toLowerCase(), payload.district);
    return { success: true, momUrl: momUrl, photoFolderUrl: photoFolderUrl, followUpId: followUpId, photoError: photoError };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  UPLOAD GOVT MoM - the meeting's conductor attaches the official
//  government-issued MoM (PDF only) to an already-conducted meeting.
//  Stored in CONDUCTED_SHEET column V ("Govt MoM"), comma-separated URLs.
// ------------------------------------------------------------
var GOVT_MOM_COL = 22;   // Column V (1-based) in the Conducted Meetings sheet

function uploadGovtMom(payload, session) {
  try {
    var meetingId = (payload && payload.meetingId || '').toString().trim();
    var files     = (payload && payload.files) || [];
    if (!meetingId)    return { success:false, message:'Missing meeting id' };
    if (!files.length) return { success:false, message:'No file provided' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var cS = ss.getSheetByName(CONDUCTED_SHEET);
    if (!cS) return { success:false, message:'No conducted meetings found' };

    // Find the conducted row + verify the caller is the conductor
    var data = cS.getDataRange().getValues();
    var rowIdx = -1, rowEmail = '', district = '';
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString() === meetingId) {
        rowIdx   = i;
        rowEmail = (data[i][4] || '').toString().trim().toLowerCase();  // E = conductor email
        district = (data[i][1] || '').toString();                       // B = district
        break;
      }
    }
    if (rowIdx < 0) return { success:false, message:'Meeting not found or not yet conducted' };

    var caller = (session && session.email || '').toString().trim().toLowerCase();
    if (!caller || caller !== rowEmail) {
      return { success:false, message:'Only the person who conducted this meeting can upload its Govt MoM.' };
    }

    // Validate (PDF only) + save to the meeting's Drive folder
    var root  = getRootMeetingsFolder();
    var distF = getOrCreateFolder(root, district || 'General');
    var mtgF  = getOrCreateFolder(distF, meetingId);
    try { mtgF.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}

    var newUrls = [];
    for (var f = 0; f < files.length; f++) {
      var file = files[f];
      var type = (file.type || '').toString().toLowerCase();
      var name = (file.name || '').toString();
      var isPdf = type.indexOf('pdf') !== -1 || /\.pdf$/i.test(name);
      if (!isPdf) return { success:false, message:'Only PDF files are allowed for the Govt MoM.' };
      var decoded = Utilities.base64Decode(file.data);
      var blob = Utilities.newBlob(decoded, 'application/pdf',
                   meetingId + '_GovtMoM_' + (new Date().getTime()) + '_' + (f+1) + '.pdf');
      var saved = mtgF.createFile(blob);
      try { saved.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
      newUrls.push(saved.getUrl());
    }

    // Append to column V (keep anything previously uploaded)
    var existing = (data[rowIdx][GOVT_MOM_COL - 1] || '').toString().trim();
    var all = (existing ? existing.split(/\s*,\s*/) : []).concat(newUrls).filter(function(u){ return u; });
    cS.getRange(rowIdx + 1, GOVT_MOM_COL).setValue(all.join(', '));

    invalidateUser(rowEmail);   // clears mymtg_<email> + reportData so it shows immediately
    return { success:true, govtMom: all.join(', '), count: all.length, added: newUrls.length };
  } catch(err) {
    return { success:false, message: err.message };
  }
}

// ------------------------------------------------------------
//  MoM - auto Google Doc creation
// ------------------------------------------------------------
function createMoMDoc(d, photoFolderUrl) {
  var title = 'MoM | ' + d.meetingId + ' | ' + d.adhikariName + ' | ' + d.conductDate;
  var doc   = DocumentApp.create(title);
  var body  = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // Title
  var h = body.appendParagraph('MINUTES OF MEETING');
  h.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  h.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  h.editAsText().setForegroundColor('#6B0F0F').setFontSize(18);

  body.appendParagraph('Educate Girls - Meeting Management System')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(10).setForegroundColor('#888888').setItalic(true);
  body.appendHorizontalRule();

  function sec(t) {
    var p = body.appendParagraph(t);
    p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    p.editAsText().setForegroundColor('#1F4E79').setFontSize(12);
    return p;
  }

  // Meeting Details table
  sec('Meeting Details');
  var tblData = [
    ['Meeting ID',    d.meetingId   || '-'],
    ['Conduct Date',  d.conductDate + (d.conductTime ? '   ' + d.conductTime : '')],
    ['Planned Date',  d.originalDate || '-'],
    ['Meeting Type',  d.meetingType || '-'],
    ['Duration',      d.duration    || '-'],
    ['District',      d.district    || '-'],
    ['Purpose',       d.purpose     || '-']
  ];
  var tbl = body.appendTable(tblData);
  tbl.setBorderWidth(0.5);
  for (var r = 0; r < tblData.length; r++) {
    tbl.getCell(r, 0).editAsText().setBold(true).setForegroundColor('#374151');
    tbl.getCell(r, 0).setBackgroundColor('#F3F4F6');
  }

  // Stakeholder
  sec('Stakeholder / Official');
  body.appendParagraph((d.adhikariName || '-') + '   |   ' + (d.adhikariPost || '-'));

  // Attendees
  sec('Attended By (EG Team)');
  body.appendParagraph((d.employeeName || '-') + '   (' + (d.designation || '-') + ')');
  if (d.colleagueName) {
    body.appendParagraph((d.colleagueName || '') + '   (' + (d.colleaguePost || '') + ')');
  }

  // Agenda
  sec('Agenda');
  body.appendParagraph(d.agenda || '-').editAsText().setItalic(true).setForegroundColor('#4B5563');

  // Meeting Documents (attached at plan time) - look up by meeting ID
  try {
    var _docUrl = d.docUrl || (getDocUrlMap_()[d.meetingId] || '');
    if (_docUrl) {
      sec('Meeting Documents');
      var docP = body.appendParagraph('');
      docP.appendText('Documents Folder Link: ').setBold(true);
      docP.appendText(_docUrl);
    }
  } catch(e) {}

  // Key Discussion Points
  sec('Key Discussion Points');
  var points = (d.keyPoints || '').split('\n').filter(function(p) { return p.trim(); });
  if (points.length) {
    points.forEach(function(pt) { body.appendListItem(pt.trim()); });
  } else {
    body.appendParagraph('-');
  }

  // Photos
  if (photoFolderUrl) {
    sec('Meeting Photos');
    var photoP = body.appendParagraph('');
    photoP.appendText('Drive Folder Link: ').setBold(true);
    photoP.appendText(photoFolderUrl);
  }

  // Follow-up
  if (d.followUp && d.followUp.date) {
    sec('Follow-up Meeting');
    var fuTbl = body.appendTable([
      ['Follow-up ID',   d.followUpId || '-'],
      ['Scheduled Date', d.followUp.date + (d.followUp.time ? '   ' + d.followUp.time : '')],
      ['With',           (d.adhikariName || '') + '   (' + (d.adhikariPost || '') + ')']
    ]);
    fuTbl.setBorderWidth(0.5);
    for (var fr = 0; fr < 3; fr++) {
      fuTbl.getCell(fr, 0).editAsText().setBold(true).setForegroundColor('#1D4ED8');
      fuTbl.getCell(fr, 0).setBackgroundColor('#EFF6FF');
    }
  }

  body.appendHorizontalRule();
  body.appendParagraph('Generated: ' + new Date().toLocaleString('en-IN') + '   |   EG Meeting Management System')
      .editAsText().setFontSize(9).setForegroundColor('#9CA3AF').setItalic(true);

  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  if (photoFolderUrl) {
    try {
      var root  = getRootMeetingsFolder();
      var distF = getOrCreateFolder(root, d.district || 'General');
      var mtgF  = getOrCreateFolder(distF, d.meetingId);
      mtgF.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch(e) {}
  }
  return doc.getUrl();
}

// ------------------------------------------------------------
//  POSTPONE MEETING - same ID, new date, history in sheet
// ------------------------------------------------------------
// ------------------------------------------------------------
//  WRITE ONCE
//  Google's front door can hold a request for a minute (measured 27 Sep 2026:
//  7 to 63 seconds for the same request). In that time an officer can close
//  the form, open the meeting again and send it a second time, and both
//  copies used to be written: MTG-20260927-142753 was recorded as not held
//  twice. postponeMeeting and cancelMeeting now check under a lock whether
//  the thing is already done, so a second copy finds it and writes nothing.
//  conductMeeting has had its own check (ALREADY_CONDUCTED) for a while.
// ------------------------------------------------------------
function withScriptLock_(fn) {
  var lock = LockService.getScriptLock(), got = false;
  // Never block a genuine save because the lock was busy: after 15 s it goes
  // ahead without one, which is no worse than before the lock existed.
  try { got = lock.tryLock(15000); } catch (le) {}
  try {
    var out = fn();
    // Written before the lock is let go, or the next copy could read the old status.
    if (got) { try { SpreadsheetApp.flush(); } catch (fe) {} }
    return out;
  } finally {
    if (got) { try { lock.releaseLock(); } catch (re) {} }
  }
}
// A date cell and the "2026-10-03" the page sends compare equal through this.
function dayKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return (v == null ? '' : v).toString().trim();
}

function postponeMeeting(payload) {
  return withScriptLock_(function() { return postponeMeeting_(payload); });
}
function postponeMeeting_(payload) {
  try {
    if (payload.email && !getEmployeeByEmail(payload.email.trim().toLowerCase())) {
      return { success: false, message: 'ACCESS_REVOKED' };
    }

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var now = new Date();

    // Update Plan Meetings: new date, status back to Planned
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    if (planSheet) {
      var pd = planSheet.getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        if ((pd[i][0] || '').toString() === payload.meetingId) {
          // Already moved to this very date: this is the same request again.
          if ((pd[i][13] || '').toString().trim() === 'Postponed' &&
              dayKey_(pd[i][5]) === dayKey_(payload.newDate)) {
            return { success: true, already: true };
          }
          planSheet.getRange(i+1, 6).setValue(payload.newDate);    // F new date
          planSheet.getRange(i+1, 14).setValue('Postponed');       // N keep as Postponed so user sees it was rescheduled
          planSheet.getRange(i+1, 17).setValue('Postponed from ' + payload.originalDate + (payload.reason ? ': ' + payload.reason : '')); // Q reason
          break;
        }
      }
    }

    // Save to Postponed Meetings sheet for history
    var pSheet = ss.getSheetByName(POSTPONED_SHEET);
    if (!pSheet) {
      pSheet = ss.insertSheet(POSTPONED_SHEET);
      var ph = ['Meeting ID','District','Employee Name','Email',
                'Stakeholder Name','Stakeholder Post','Purpose',
                'Original Date','New Date','Reason','Postponed At'];
      pSheet.appendRow(ph);
      pSheet.getRange(1,1,1,ph.length).setBackground('#B45309').setFontColor('#fff').setFontWeight('bold');
      pSheet.setFrozenRows(1);
    }
    pSheet.appendRow([
      payload.meetingId,    payload.district,    payload.employeeName, payload.email,
      payload.adhikariName, payload.adhikariPost, payload.purpose,
      payload.originalDate, payload.newDate,      payload.reason || '',
      now.toLocaleString('en-IN')
    ]);

    invalidateUser((payload.email || '').trim().toLowerCase());
    return { success: true };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  CANCEL / NO-CONDUCT MEETING
// ------------------------------------------------------------
function cancelMeeting(payload) {
  return withScriptLock_(function() { return cancelMeeting_(payload); });
}
function cancelMeeting_(payload) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tz    = Session.getScriptTimeZone();
    var now   = new Date();
    var sheet = ss.getSheetByName(MEETINGS_SHEET);
    if (!sheet) return { success: false, message: 'Sheet not found.' };

    // Find the row
    var data   = sheet.getDataRange().getValues();
    var rowData = null, rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString() === payload.meetingId) {
        rowData = data[i]; rowIdx = i; break;
      }
    }
    if (!rowData) return { success: false, message: 'Meeting not found.' };
    if (!getEmployeeByEmail((rowData[4] || '').toString().trim().toLowerCase())) {
      return { success: false, message: 'ACCESS_REVOKED' };
    }
    // Already recorded as not held: the same request again. Nothing to write.
    if ((rowData[13] || '').toString().trim() === 'Cancelled') return { success: true, already: true };

    // Save to Cancelled Meetings sheet
    var cSheet = ss.getSheetByName(CANCELLED_SHEET);
    if (!cSheet) {
      cSheet = ss.insertSheet(CANCELLED_SHEET);
      var ch = ['Meeting ID','District','Employee Name','Post','Email',
                'Meeting Date','Meeting Time','Duration','Meeting Type',
                'Stakeholder Name','Stakeholder Post','Meeting Purpose','Meeting Agenda',
                'Colleague Name','Colleague Post',
                'Reason','Cancelled At'];
      cSheet.appendRow(ch);
      cSheet.getRange(1,1,1,ch.length).setBackground('#DC2626').setFontColor('#fff').setFontWeight('bold');
      cSheet.setFrozenRows(1);
    }

    var rawDate     = rowData[5];
    var meetingDate = fmtDateVal(rawDate);
    var meetingTime = fmtTimeVal(rowData[6]);

    cSheet.appendRow([
      (rowData[0]  || '').toString(),  // Meeting ID
      (rowData[1]  || '').toString(),  // District
      (rowData[2]  || '').toString(),  // Employee Name
      (rowData[3]  || '').toString(),  // Post
      (rowData[4]  || '').toString(),  // Email
      meetingDate,                     // Meeting Date
      meetingTime,                     // Meeting Time
      (rowData[7]  || '').toString(),  // Duration
      (rowData[8]  || '').toString(),  // Meeting Type
      (rowData[9]  || '').toString(),  // Stakeholder Name
      (rowData[10] || '').toString(),  // Stakeholder Post
      (rowData[11] || '').toString(),  // Meeting Purpose
      (rowData[12] || '').toString(),  // Meeting Agenda
      (rowData[17] || '').toString(),  // Colleague Name
      (rowData[18] || '').toString(),  // Colleague Post
      payload.reason || '',            // Reason
      now.toLocaleString('en-IN')      // Cancelled At
    ]);

    // Update status in Plan Meetings to "Cancelled" - NEVER delete, keeps master ledger intact for dashboard reporting
    sheet.getRange(rowIdx + 1, 14).setValue('Cancelled');

    invalidateUser((payload.email || '').trim().toLowerCase());
    return { success: true };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  DELETE MEETING
// ------------------------------------------------------------
function deleteMeeting(meetingId, email) {
  try {
    // Field role cannot delete - prevents fake-meeting create-then-delete (full audit trail)
    if (email) {
      var emp = getEmployeeByEmail(email.trim().toLowerCase());
      var role = emp && emp.role ? emp.role.toString().trim().toLowerCase() : '';
      if (role === 'field') {
        return { success: false, message: 'Delete not allowed for Field role.' };
      }
    }
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MEETINGS_SHEET);
    if (!sheet) return { success: false, message: 'Sheet not found.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString() === meetingId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Meeting not found.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  ACTIONED MEETINGS - for My Meetings view
//  Reads Conducted + Cancelled sheets (Postponed stays in Plan Meetings)
// ------------------------------------------------------------
function getActionedMeetings(email) {
  try {
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var meetings = [];
    var emailKey = email.trim().toLowerCase();

    // From Conducted Meetings
    var cSheet = ss.getSheetByName(CONDUCTED_SHEET);
    if (cSheet && cSheet.getLastRow() > 1) {
      var cd = cSheet.getDataRange().getValues();
      for (var i = 1; i < cd.length; i++) {
        if ((cd[i][4] || '').toString().trim().toLowerCase() !== emailKey) continue;
        meetings.push({
          meetingId:    (cd[i][0]  || '').toString(),
          district:     (cd[i][1]  || '').toString(),
          date:         fmtDateVal(cd[i][5]),            // Original Date
          meetingTime:  fmtTimeVal(cd[i][6]),           // Original Time
          duration:     (cd[i][7]  || '').toString(),
          type:         (cd[i][8]  || '').toString(),
          adhikariName: (cd[i][9]  || '').toString(),
          adhikariPost: (cd[i][10] || '').toString(),
          purpose:      (cd[i][11] || '').toString(),
          agenda:       (cd[i][12] || '').toString(),
          status:       'Conducted',
          conductDate:  fmtDateVal(cd[i][13]),
          conductTime:  fmtTimeVal(cd[i][14]),
          keyPoints:    (cd[i][15] || '').toString(),
          photoLink:    (cd[i][16] || '').toString(),
          momLink:      (cd[i][17] || '').toString(),
          colleagueName:(cd[i][18] || '').toString(),
          colleaguePost:(cd[i][19] || '').toString(),
          reason:       ''
        });
      }
    }

    // From Cancelled Meetings
    var xSheet = ss.getSheetByName(CANCELLED_SHEET);
    if (xSheet && xSheet.getLastRow() > 1) {
      var xd = xSheet.getDataRange().getValues();
      for (var j = 1; j < xd.length; j++) {
        if ((xd[j][4] || '').toString().trim().toLowerCase() !== emailKey) continue;
        meetings.push({
          meetingId:    (xd[j][0]  || '').toString(),
          district:     (xd[j][1]  || '').toString(),
          date:         fmtDateVal(xd[j][5]),            // Meeting Date
          meetingTime:  fmtTimeVal(xd[j][6]),
          duration:     (xd[j][7]  || '').toString(),
          type:         (xd[j][8]  || '').toString(),
          adhikariName: (xd[j][9]  || '').toString(),
          adhikariPost: (xd[j][10] || '').toString(),
          purpose:      (xd[j][11] || '').toString(),
          agenda:       (xd[j][12] || '').toString(),
          status:       'Cancelled',
          reason:       (xd[j][15] || '').toString(),
          colleagueName:(xd[j][13] || '').toString(),
          colleaguePost:(xd[j][14] || '').toString(),
          conductDate:  '',
          conductTime:  '',
          keyPoints:    '',
          photoLink:    '',
          momLink:      ''
        });
      }
    }

    return meetings;
  } catch(err) {
    return [];
  }
}

// ------------------------------------------------------------
//  ALL MEETINGS - combined view for My Meetings tab
//  Returns Plan Meetings (all statuses) + Conducted + Cancelled
// ------------------------------------------------------------
function getAllMyMeetings(email) {
  try {
    var emailKey = email.trim().toLowerCase();
    var cacheKey = 'mymtg_' + emailKey;
    var hit = cGet(cacheKey);
    if (hit) return hit;

    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tz       = Session.getScriptTimeZone();
    var meetings = [];

    // 1. Postponed Meetings sheet - history of all reschedules
    var pSheet = ss.getSheetByName(POSTPONED_SHEET);
    if (pSheet && pSheet.getLastRow() > 1) {
      var phd = (sheetRows_(POSTPONED_SHEET) || []);
      // Columns: MeetingID(0) District(1) EmployeeName(2) Email(3)
      //          StakeholderName(4) StakeholderPost(5) Purpose(6)
      //          OriginalDate(7) NewDate(8) Reason(9) PostponedAt(10)
      var seenP = {};   // a reschedule sent twice shows once
      for (var i = 1; i < phd.length; i++) {
        if ((phd[i][3] || '').toString().trim().toLowerCase() !== emailKey) continue;
        var pKey = (phd[i][0] || '') + '|' + dayKey_(phd[i][7]) + '|' + dayKey_(phd[i][8]);
        if (seenP[pKey]) continue; seenP[pKey] = 1;
        meetings.push({
          meetingId:       (phd[i][0] || '').toString(),
          district:        (phd[i][1] || '').toString(),
          date:            fmtDateVal(phd[i][7]),            // Original Date
          meetingTime:     '',
          duration:        '',
          type:            '',
          adhikariName:    (phd[i][4] || '').toString(),
          adhikariPost:    (phd[i][5] || '').toString(),
          purpose:         (phd[i][6] || '').toString(),
          agenda:          '',
          status:          'Postponed',
          reason:          (phd[i][9] || '') + (phd[i][8] ? ' → New: ' + phd[i][8] : ''),
          colleagueName:   '',
          colleaguePost:   '',
          parentMeetingId: '',
          conductDate: '', conductTime: '', keyPoints: '',
          photoLink:   '', momLink:      ''
        });
      }
    }

    // 2. Conducted Meetings
    var cSheet = ss.getSheetByName(CONDUCTED_SHEET);
    if (cSheet && cSheet.getLastRow() > 1) {
      var cd = (sheetRows_(CONDUCTED_SHEET) || []);
      for (var j = 1; j < cd.length; j++) {
        if ((cd[j][4] || '').toString().trim().toLowerCase() !== emailKey) continue;
        meetings.push({
          meetingId:    (cd[j][0]  || '').toString(),
          district:     (cd[j][1]  || '').toString(),
          date:         fmtDateVal(cd[j][5]),
          meetingTime:  fmtTimeVal(cd[j][6]),
          duration:     (cd[j][7]  || '').toString(),
          type:         (cd[j][8]  || '').toString(),
          adhikariName: (cd[j][9]  || '').toString(),
          adhikariPost: (cd[j][10] || '').toString(),
          purpose:      (cd[j][11] || '').toString(),
          agenda:       (cd[j][12] || '').toString(),
          status:       'Conducted',
          conductDate:  fmtDateVal(cd[j][13]),
          conductTime:  fmtTimeVal(cd[j][14]),
          keyPoints:    (cd[j][15] || '').toString(),
          photoLink:    (cd[j][16] || '').toString(),
          momLink:      (cd[j][17] || '').toString(),
          govtMom:      (cd[j][21] || '').toString(),   // V = Govt MoM (comma-separated PDF urls)
          outcome:      (cd[j][COL_CON_OUTCOME-1] || '').toString(),   // AF
          priority:     (cd[j][22] || '').toString(),   // W..AB = AI tags
          flag:         (cd[j][23] || '').toString(),
          nextAction:   (cd[j][24] || '').toString(),
          escalate:     (cd[j][25] || '').toString(),
          category:     (cd[j][26] || '').toString(),
          momSummary:   (cd[j][27] || '').toString(),
          colleagueName:(cd[j][18] || '').toString(),
          colleaguePost:(cd[j][19] || '').toString(),
          reason: '', parentMeetingId: ''
        });
      }
    }

    // 3. Cancelled Meetings
    var xSheet = ss.getSheetByName(CANCELLED_SHEET);
    if (xSheet && xSheet.getLastRow() > 1) {
      var xd = (sheetRows_(CANCELLED_SHEET) || []), seenX = {};   // one meeting, one row
      for (var k = 1; k < xd.length; k++) {
        if ((xd[k][4] || '').toString().trim().toLowerCase() !== emailKey) continue;
        var xId = (xd[k][0] || '').toString().trim();
        if (seenX[xId]) continue; seenX[xId] = 1;
        meetings.push({
          meetingId:    (xd[k][0]  || '').toString(),
          district:     (xd[k][1]  || '').toString(),
          date:         fmtDateVal(xd[k][5]),
          meetingTime:  fmtTimeVal(xd[k][6]),
          duration:     (xd[k][7]  || '').toString(),
          type:         (xd[k][8]  || '').toString(),
          adhikariName: (xd[k][9]  || '').toString(),
          adhikariPost: (xd[k][10] || '').toString(),
          purpose:      (xd[k][11] || '').toString(),
          agenda:       (xd[k][12] || '').toString(),
          status:       'Cancelled',
          reason:       (xd[k][15] || '').toString(),
          colleagueName:(xd[k][13] || '').toString(),
          colleaguePost:(xd[k][14] || '').toString(),
          conductDate: '', conductTime: '', keyPoints: '',
          photoLink:   '', momLink:      '', parentMeetingId: ''
        });
      }
    }

    var _dm = getDocUrlMap_();
    meetings.forEach(function(m){ m.docUrl = _dm[m.meetingId] || ''; });
    cPut(cacheKey, meetings, C_TTL_LIVE);
    return meetings;
  } catch(err) {
    return [];
  }
}

// ------------------------------------------------------------
//  DOC URL MAP - meetingId → Documents folder URL (from Plan sheet
//  col V). Plan rows are never deleted, so this resolves docs for a
//  meeting in any later state (conducted/postponed/cancelled).
// ------------------------------------------------------------
function getDocUrlMap_() {
  var cacheKey = 'docUrlMap';
  var hit = cGet(cacheKey);
  if (hit) return hit;
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MEETINGS_SHEET);
  var map   = {};
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var id  = (data[i][0]  || '').toString();
      var doc = (data[i][21] || '').toString(); // V
      if (id && doc) map[id] = doc;
    }
  }
  cPut(cacheKey, map, C_TTL_LIVE);
  return map;
}

// ------------------------------------------------------------
//  DISTRICT ALL MEETINGS - for District Meetings view
//  All statuses: Planned, Conducted, Postponed, Cancelled
// ------------------------------------------------------------
function getDistrictAllMeetings(district) {
  try {
    var distL    = district.trim().toLowerCase();
    var cacheKey = 'distMtg_' + distL;
    var hit      = cGet(cacheKey);
    if (hit) return hit;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var meetings = [];

    // 1. Plan Meetings - Planned / Follow-up only (Cancelled/Postponed go to their own sheets)
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    if (planSheet) {
      var pd = (sheetRows_(MEETINGS_SHEET) || []);
      for (var i = 1; i < pd.length; i++) {
        if ((pd[i][1] || '').toString().trim().toLowerCase() !== distL) continue;
        var st = (pd[i][13] || 'Planned').toString();
        if (st !== 'Planned' && st !== 'Follow-up') continue;
        meetings.push({
          meetingId:    (pd[i][0]  || '').toString(),
          employeeName: (pd[i][2]  || '').toString(),
          designation:  (pd[i][3]  || '').toString(),
          date:         fmtDateVal(pd[i][5]),
          type:         (pd[i][8]  || '').toString(),
          adhikariName: (pd[i][9]  || '').toString(),
          adhikariPost: (pd[i][10] || '').toString(),
          purpose:      (pd[i][11] || '').toString(),
          colleagueName:(pd[i][17] || '').toString(),
          colleaguePost:(pd[i][18] || '').toString(),
          status: st, conductDate: '', keyPoints: '', photoLink: '', momLink: '', reason: ''
        });
      }
    }

    // 2. Conducted
    var condSheet = ss.getSheetByName(CONDUCTED_SHEET);
    if (condSheet) {
      var cd = (sheetRows_(CONDUCTED_SHEET) || []);
      for (var i = 1; i < cd.length; i++) {
        if ((cd[i][1] || '').toString().trim().toLowerCase() !== distL) continue;
        meetings.push({
          meetingId:    (cd[i][0]  || '').toString(),
          employeeName: (cd[i][2]  || '').toString(),
          designation:  (cd[i][3]  || '').toString(),
          date:         fmtDateVal(cd[i][5]),
          conductDate:  fmtDateVal(cd[i][13]),
          type:         (cd[i][8]  || '').toString(),
          adhikariName: (cd[i][9]  || '').toString(),
          adhikariPost: (cd[i][10] || '').toString(),
          purpose:      (cd[i][11] || '').toString(),
          keyPoints:    (cd[i][15] || '').toString(),
          photoLink:    (cd[i][16] || '').toString(),
          momLink:      (cd[i][17] || '').toString(),
          colleagueName:(cd[i][18] || '').toString(),
          colleaguePost:(cd[i][19] || '').toString(),
          status: 'Conducted', reason: ''
        });
      }
    }

    // 3. Postponed
    var postSheet = ss.getSheetByName(POSTPONED_SHEET);
    if (postSheet) {
      var xd = (sheetRows_(POSTPONED_SHEET) || []), seenP = {};   // a reschedule sent twice shows once
      for (var i = 1; i < xd.length; i++) {
        if ((xd[i][1] || '').toString().trim().toLowerCase() !== distL) continue;
        var pKey = (xd[i][0] || '') + '|' + dayKey_(xd[i][7]) + '|' + dayKey_(xd[i][8]);
        if (seenP[pKey]) continue; seenP[pKey] = 1;
        meetings.push({
          meetingId:    (xd[i][0] || '').toString(),
          employeeName: (xd[i][2] || '').toString(),
          designation:  '',
          date:         fmtDateVal(xd[i][7]),
          conductDate:  '',
          type:         '',
          adhikariName: (xd[i][4] || '').toString(),
          adhikariPost: (xd[i][5] || '').toString(),
          purpose:      (xd[i][6] || '').toString(),
          reason:       (xd[i][9] || '').toString(),
          colleagueName:'', colleaguePost:'', keyPoints:'', photoLink:'', momLink:'',
          status: 'Postponed'
        });
      }
    }

    // 4. Cancelled
    var cancelSheet = ss.getSheetByName(CANCELLED_SHEET);
    if (cancelSheet) {
      var xc = (sheetRows_(CANCELLED_SHEET) || []), seenX = {};   // one meeting, one row
      for (var i = 1; i < xc.length; i++) {
        if ((xc[i][1] || '').toString().trim().toLowerCase() !== distL) continue;
        var xId = (xc[i][0] || '').toString().trim();
        if (seenX[xId]) continue; seenX[xId] = 1;
        meetings.push({
          meetingId:    (xc[i][0]  || '').toString(),
          employeeName: (xc[i][2]  || '').toString(),
          designation:  (xc[i][3]  || '').toString(),
          date:         fmtDateVal(xc[i][5]),
          conductDate:  '',
          type:         (xc[i][8]  || '').toString(),
          adhikariName: (xc[i][9]  || '').toString(),
          adhikariPost: (xc[i][10] || '').toString(),
          purpose:      (xc[i][11] || '').toString(),
          reason:       (xc[i][15] || '').toString(),
          colleagueName:(xc[i][13] || '').toString(),
          colleaguePost:(xc[i][14] || '').toString(),
          keyPoints:'', photoLink:'', momLink:'',
          status: 'Cancelled'
        });
      }
    }

    var _dm = getDocUrlMap_();
    meetings.forEach(function(m){ m.docUrl = _dm[m.meetingId] || ''; });
    cPut(cacheKey, meetings, C_TTL_LIVE);
    return meetings;
  } catch(err) { return []; }
}

// ------------------------------------------------------------
//  STATE ALL MEETINGS - for State Meetings view
//  All districts, all statuses
// ------------------------------------------------------------
function getStateAllMeetings() {
  try {
    var cacheKey = 'stateMtg_all';
    var hit      = cGet(cacheKey);
    if (hit) return hit;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var meetings = [];

    // 1. Plan Meetings - Planned / Follow-up only
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    if (planSheet) {
      var pd = (sheetRows_(MEETINGS_SHEET) || []);
      for (var i = 1; i < pd.length; i++) {
        if (!pd[i][0]) continue;
        var st = (pd[i][13] || 'Planned').toString();
        if (st !== 'Planned' && st !== 'Follow-up') continue;
        meetings.push({
          meetingId:    (pd[i][0]  || '').toString(),
          district:     (pd[i][1]  || '').toString(),
          employeeName: (pd[i][2]  || '').toString(),
          designation:  (pd[i][3]  || '').toString(),
          date:         fmtDateVal(pd[i][5]),
          type:         (pd[i][8]  || '').toString(),
          adhikariName: (pd[i][9]  || '').toString(),
          adhikariPost: (pd[i][10] || '').toString(),
          purpose:      (pd[i][11] || '').toString(),
          colleagueName:(pd[i][17] || '').toString(),
          colleaguePost:(pd[i][18] || '').toString(),
          status: st, conductDate: '', keyPoints: '', photoLink: '', momLink: '', reason: ''
        });
      }
    }

    // 2. Conducted
    var condSheet = ss.getSheetByName(CONDUCTED_SHEET);
    if (condSheet) {
      var cd = (sheetRows_(CONDUCTED_SHEET) || []);
      for (var i = 1; i < cd.length; i++) {
        if (!cd[i][0]) continue;
        meetings.push({
          meetingId:    (cd[i][0]  || '').toString(),
          district:     (cd[i][1]  || '').toString(),
          employeeName: (cd[i][2]  || '').toString(),
          designation:  (cd[i][3]  || '').toString(),
          date:         fmtDateVal(cd[i][5]),
          conductDate:  fmtDateVal(cd[i][13]),
          type:         (cd[i][8]  || '').toString(),
          adhikariName: (cd[i][9]  || '').toString(),
          adhikariPost: (cd[i][10] || '').toString(),
          purpose:      (cd[i][11] || '').toString(),
          keyPoints:    (cd[i][15] || '').toString(),
          photoLink:    (cd[i][16] || '').toString(),
          momLink:      (cd[i][17] || '').toString(),
          colleagueName:(cd[i][18] || '').toString(),
          colleaguePost:(cd[i][19] || '').toString(),
          status: 'Conducted', reason: ''
        });
      }
    }

    // 3. Postponed
    var postSheet = ss.getSheetByName(POSTPONED_SHEET);
    if (postSheet) {
      var xd = (sheetRows_(POSTPONED_SHEET) || []), seenP = {};   // a reschedule sent twice shows once
      for (var i = 1; i < xd.length; i++) {
        if (!xd[i][0]) continue;
        var pKey = xd[i][0] + '|' + dayKey_(xd[i][7]) + '|' + dayKey_(xd[i][8]);
        if (seenP[pKey]) continue; seenP[pKey] = 1;
        meetings.push({
          meetingId:    (xd[i][0] || '').toString(),
          district:     (xd[i][1] || '').toString(),
          employeeName: (xd[i][2] || '').toString(),
          designation:  '',
          date:         fmtDateVal(xd[i][7]),
          conductDate:  '',
          type:         '',
          adhikariName: (xd[i][4] || '').toString(),
          adhikariPost: (xd[i][5] || '').toString(),
          purpose:      (xd[i][6] || '').toString(),
          reason:       (xd[i][9] || '').toString(),
          colleagueName:'', colleaguePost:'', keyPoints:'', photoLink:'', momLink:'',
          status: 'Postponed'
        });
      }
    }

    // 4. Cancelled
    var cancelSheet = ss.getSheetByName(CANCELLED_SHEET);
    if (cancelSheet) {
      var xc = (sheetRows_(CANCELLED_SHEET) || []), seenX = {};   // one meeting, one row
      for (var i = 1; i < xc.length; i++) {
        if (!xc[i][0]) continue;
        var xId = xc[i][0].toString().trim();
        if (seenX[xId]) continue; seenX[xId] = 1;
        meetings.push({
          meetingId:    (xc[i][0]  || '').toString(),
          district:     (xc[i][1]  || '').toString(),
          employeeName: (xc[i][2]  || '').toString(),
          designation:  (xc[i][3]  || '').toString(),
          date:         fmtDateVal(xc[i][5]),
          conductDate:  '',
          type:         (xc[i][8]  || '').toString(),
          adhikariName: (xc[i][9]  || '').toString(),
          adhikariPost: (xc[i][10] || '').toString(),
          purpose:      (xc[i][11] || '').toString(),
          reason:       (xc[i][15] || '').toString(),
          colleagueName:(xc[i][13] || '').toString(),
          colleaguePost:(xc[i][14] || '').toString(),
          keyPoints:'', photoLink:'', momLink:'',
          status: 'Cancelled'
        });
      }
    }

    var _dm = getDocUrlMap_();
    meetings.forEach(function(m){ m.docUrl = _dm[m.meetingId] || ''; });
    cPut(cacheKey, meetings, C_TTL_LIVE);
    return meetings;
  } catch(err) { return []; }
}

// ------------------------------------------------------------
//  ZONE ALL MEETINGS - for Zone Meetings view (Zone role)
//  A zone spans multiple districts; show every meeting whose
//  district belongs to this zone (mapping comes from Employee_DB).
// ------------------------------------------------------------
// ------------------------------------------------------------
//  ZONE STRUCTURE - fixed org mapping: Zone → Admin Districts.
//  A meeting/employee's zone is derived from its (admin) district;
//  space/case differences are normalised (e.g. "BARA BANKI" == "BARABANKI").
// ------------------------------------------------------------
var ZONE_DISTRICTS = {
  'UP ZONE-1': ['BANDA','CHITRAKOOT','FATEHPUR','RAE BARELI','UNNAO','BHADOHI','KAUSHAMBI','MIRZAPUR','PRAYAGRAJ','SONBHADRA'],
  'UP ZONE-2': ['BAHRAICH','SHRAVASTI','BALRAMPUR','GONDA','KUSHINAGAR','MAHARAJGANJ'],
  'UP ZONE-3': ['BUDAUN','FARRUKHABAD','HARDOI','LAKHIMPUR KHERI','SHAHJAHANPUR','BARA BANKI','SITAPUR']
};
// State-level-only districts - selectable by State users when planning a meeting,
// but not part of any zone (zone leads don't see them; not zone-grouped in analytics).
var STATE_EXTRA_DISTRICTS = ['LUCKNOW'];
// The same district written two ways in different sheets. The employee
// master has MAHRAJGANJ and SANT RAVIDAS NAGAR (BHADOHI) where the zone list
// has MAHARAJGANJ and BHADOHI, so meetings filed there belonged to no zone at
// all: missing from Zone Meetings, from the zone report, and counted as zero
// by the chatbot. Found by checking every district in the employee list and
// the meetings against the zone list on 27 Sep 2026; these were the only two.
// normDist_ is only ever a comparison and cache key, never shown or written,
// so mapping an alias here changes which names match and nothing else.
var DISTRICT_ALIASES = {
  'MAHRAJGANJ': 'MAHARAJGANJ',
  'SANTRAVIDASNAGAR(BHADOHI)': 'BHADOHI',
  'SANTRAVIDASNAGAR': 'BHADOHI',
  'BHADOHI(SANTRAVIDASNAGAR)': 'BHADOHI'
};
function normDist_(d) {
  var k = (d || '').toString().trim().toUpperCase().replace(/\s+/g, '');
  return DISTRICT_ALIASES[k] || k;
}
// The Zone column is typed by hand, as "UP ZONE-1", "UP Zone 1" or
// "upzone1". normDist_ only drops spaces and case, so the dash still had to
// match exactly, and a lead whose column had no dash got no zone at all:
// an empty Zone Meetings, no feedback button, no zone in their report.
// Letters and digits only here. normDist_ itself is left alone because every
// district comparison in the file goes through it.
function zoneNorm_(z) { return (z || '').toString().toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function findZoneKey_(zone) {
  var zn = zoneNorm_(zone);
  if (!zn) return '';
  for (var z in ZONE_DISTRICTS) { if (zoneNorm_(z) === zn) return z; }
  return '';
}
function districtToZone_(district) {
  var nd = normDist_(district);
  if (!nd) return '';
  for (var z in ZONE_DISTRICTS) {
    for (var i = 0; i < ZONE_DISTRICTS[z].length; i++) {
      if (normDist_(ZONE_DISTRICTS[z][i]) === nd) return z;
    }
  }
  return '';
}

// districts of a zone as { NORMDIST: true } - for membership checks
function getDistrictsInZone_(zone) {
  var zkey = findZoneKey_(zone);
  var set  = {};
  if (zkey) ZONE_DISTRICTS[zkey].forEach(function(d){ set[normDist_(d)] = true; });
  return set;
}

// meetingId → ZONE, from the meeting's (admin) district. Falls back to
// the creator's district/zone for legacy blank-district meetings.
function getMeetingZoneMap_() {
  var cacheKey = 'meetingZoneMap';
  var hit = cGet(cacheKey);
  if (hit) return hit;
  var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var emp  = ss.getSheetByName(EMPLOYEE_SHEET);
  var plan = ss.getSheetByName(MEETINGS_SHEET);
  var emailDist = {}, emailZone = {};
  if (emp) {
    var ed = emp.getDataRange().getValues();
    // Employee_DB: District(0) Block(1) Name(2) Desig(3) Email(4) Role(5) Zone(6)
    for (var i = 1; i < ed.length; i++) {
      var em = (ed[i][4] || '').toString().trim().toLowerCase();
      if (!em) continue;
      emailDist[em] = (ed[i][0] || '').toString();
      emailZone[em] = findZoneKey_(ed[i][6]);   // lead's Zone column (UP ZONE-X)
    }
  }
  var map = {};
  if (plan) {
    var pd = plan.getDataRange().getValues();
    // Plan Meetings: MeetingID(0) District(1) ... Email(4)
    for (var j = 1; j < pd.length; j++) {
      var id = (pd[j][0] || '').toString();
      if (!id) continue;
      var z = districtToZone_(pd[j][1]);
      if (!z) {
        var em2 = (pd[j][4] || '').toString().trim().toLowerCase();
        z = districtToZone_(emailDist[em2]) || emailZone[em2] || '';
      }
      map[id] = z;
    }
  }
  cPut(cacheKey, map, C_TTL_LIVE);
  return map;
}

function getZoneAllMeetings(zone) {
  try {
    var zkey = findZoneKey_(zone);
    if (!zkey) return [];
    var cacheKey = 'zoneMtg_' + zkey;
    var hit = cGet(cacheKey);
    if (hit) return hit;

    var zmap = getMeetingZoneMap_();          // meetingId → ZONE
    var all  = getStateAllMeetings();         // reuse (cached, all meetings)
    var filtered = all.filter(function(m) {
      return (zmap[m.meetingId] || '') === zkey;
    });
    cPut(cacheKey, filtered, C_TTL_LIVE);
    return filtered;
  } catch(err) { return []; }
}

// ------------------------------------------------------------
//  COLLEAGUE MEETING NOTIFICATION EMAIL
// ------------------------------------------------------------
function sendColleagueNotification(data, mtgId) {
  if (!data || !data.colleagueName || !data.colleagueName.trim()) return;

  // Find colleague email by name in Employee_DB
  var colleague = getEmployeeByName(data.colleagueName.trim());
  if (!colleague || !colleague.email) return; // not found, skip

  var subject = 'Meeting Invitation | ' + mtgId + ' | ' + data.adhikariPost + ', ' + (data.district || '');

  var body =
    '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #E5E7EB;border-radius:10px;overflow:hidden;">' +

    // Header
    '<div style="background:linear-gradient(135deg,#7B1010,#9B1C1C);padding:24px 28px;">' +
      '<img src="https://www.educategirls.ngo/wp-content/themes/egindia/static/images/eg-logo.png" ' +
           'style="height:34px;filter:brightness(0) invert(1);opacity:0.9;margin-bottom:12px;display:block;" />' +
      '<h2 style="color:#fff;margin:0 0 4px;font-size:19px;font-weight:700;">Meeting Invitation</h2>' +
      '<p style="color:rgba(255,255,255,0.7);margin:0;font-size:12px;letter-spacing:0.4px;">EG Meeting Management System &nbsp;|&nbsp; Government Relations</p>' +
    '</div>' +

    // Greeting
    '<div style="padding:28px 28px 0;background:#fff;">' +
      '<p style="font-size:14px;color:#111827;margin:0 0 6px;">Dear <strong>' + data.colleagueName + '</strong>,</p>' +
      '<p style="font-size:13.5px;color:#374151;line-height:1.7;margin:0 0 22px;">' +
        'You have been designated as the <strong>Accompanying Colleague</strong> for an upcoming stakeholder meeting ' +
        'organized by <strong>' + data.employeeName + '</strong>. Kindly make a note of the following details and ensure ' +
        'your availability on the scheduled date.' +
      '</p>' +
    '</div>' +

    // Meeting details card
    '<div style="padding:0 28px 22px;background:#fff;">' +
      '<div style="background:#FAFAFA;border:1px solid #E5E7EB;border-left:4px solid #7B1010;border-radius:8px;padding:18px 20px;">' +
        '<p style="margin:0 0 14px;font-size:11px;font-weight:700;color:#7B1010;text-transform:uppercase;letter-spacing:1px;">Meeting Details</p>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;color:#374151;">' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;width:38%;vertical-align:top;">Meeting ID</td>' +
            '<td style="padding:8px 0;font-weight:700;color:#111827;">' + mtgId + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;vertical-align:top;">Organized By</td>' +
            '<td style="padding:8px 0;">' + data.employeeName + '<br><span style="font-size:11.5px;color:#6B7280;">' + (data.designation || data.role || '') + ' &nbsp;|&nbsp; ' + (data.district || '') + '</span></td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;vertical-align:top;">Stakeholder</td>' +
            '<td style="padding:8px 0;font-weight:600;">' + data.adhikariName + '<br><span style="font-size:11.5px;color:#6B7280;font-weight:400;">' + data.adhikariPost + '</span></td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Date</td>' +
            '<td style="padding:8px 0;font-weight:600;color:#111827;">' + data.meetingDate + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Time</td>' +
            '<td style="padding:8px 0;">' + (data.meetingTime || 'To be confirmed') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Duration</td>' +
            '<td style="padding:8px 0;">' + (data.duration || '-') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Meeting Type</td>' +
            '<td style="padding:8px 0;">' + (data.meetingType || '-') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Purpose</td>' +
            '<td style="padding:8px 0;">' + data.purpose + '</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="padding:8px 0;color:#6B7280;vertical-align:top;">Agenda</td>' +
            '<td style="padding:8px 0;line-height:1.6;">' + data.agenda + '</td>' +
          '</tr>' +
        '</table>' +
      '</div>' +
    '</div>' +

    // Closing note
    '<div style="padding:0 28px 28px;background:#fff;">' +
      '<p style="font-size:13px;color:#6B7280;line-height:1.7;margin:0;">' +
        'Please treat this as an official communication and plan your schedule accordingly. ' +
        'For any clarification or rescheduling, please contact <strong>' + data.employeeName + '</strong> directly.' +
      '</p>' +
    '</div>' +

    // Footer
    '<div style="background:#7B1010;padding:14px 28px;text-align:center;">' +
      '<p style="color:rgba(255,255,255,0.65);font-size:11.5px;margin:0;">' +
        'This is a system-generated notification from <strong style="color:#fff;">EG Meeting Management System</strong>.<br>' +
        'Educate Girls &nbsp;|&nbsp; Government Relations Team' +
      '</p>' +
    '</div>' +

  '</div>';

  MailApp.sendEmail({
    to:       colleague.email,
    subject:  subject,
    htmlBody: body
  });
}

// ------------------------------------------------------------
//  COLLEAGUE MOM EMAIL - sent after meeting is conducted
// ------------------------------------------------------------
function sendMOMNotification(data, momUrl, photoFolderUrl, followUpId) {
  if (!data || !data.colleagueName || !data.colleagueName.trim()) return;

  var colleague = getEmployeeByName(data.colleagueName.trim());
  if (!colleague || !colleague.email) return;

  var subject = 'Minutes of Meeting | ' + data.meetingId + ' | ' + data.adhikariPost + ', ' + (data.district || '');

  // Format key points as bullet list
  var kpLines = (data.keyPoints || '').split('\n').filter(function(l){ return l.trim(); });
  var kpHtml = kpLines.map(function(l){
    return '<tr><td style="padding:5px 0 5px 8px;color:#374151;font-size:13px;border-bottom:1px solid #F3F4F6;">• ' + l.trim() + '</td></tr>';
  }).join('');
  if (!kpHtml) kpHtml = '<tr><td style="padding:5px 0;color:#6B7280;font-size:13px;">-</td></tr>';

  var body =
    '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #E5E7EB;border-radius:10px;overflow:hidden;">' +

    // Header
    '<div style="background:linear-gradient(135deg,#7B1010,#9B1C1C);padding:24px 28px;">' +
      '<img src="https://www.educategirls.ngo/wp-content/themes/egindia/static/images/eg-logo.png" ' +
           'style="height:34px;filter:brightness(0) invert(1);opacity:0.9;margin-bottom:12px;display:block;" />' +
      '<h2 style="color:#fff;margin:0 0 4px;font-size:19px;font-weight:700;">Minutes of Meeting (MoM)</h2>' +
      '<p style="color:rgba(255,255,255,0.7);margin:0;font-size:12px;letter-spacing:0.4px;">EG Meeting Management System &nbsp;|&nbsp; Government Relations</p>' +
    '</div>' +

    // Greeting
    '<div style="padding:28px 28px 0;background:#fff;">' +
      '<p style="font-size:14px;color:#111827;margin:0 0 6px;">Dear <strong>' + data.colleagueName + '</strong>,</p>' +
      '<p style="font-size:13.5px;color:#374151;line-height:1.7;margin:0 0 22px;">' +
        'Please find below the Minutes of Meeting (MoM) for the stakeholder meeting you attended alongside ' +
        '<strong>' + data.employeeName + '</strong>. Kindly review the key discussion points and take note of any follow-up actions.' +
      '</p>' +
    '</div>' +

    // Meeting details card
    '<div style="padding:0 28px 18px;background:#fff;">' +
      '<div style="background:#FAFAFA;border:1px solid #E5E7EB;border-left:4px solid #7B1010;border-radius:8px;padding:18px 20px;">' +
        '<p style="margin:0 0 14px;font-size:11px;font-weight:700;color:#7B1010;text-transform:uppercase;letter-spacing:1px;">Meeting Details</p>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;color:#374151;">' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;width:38%;vertical-align:top;">Meeting ID</td>' +
            '<td style="padding:8px 0;font-weight:700;color:#111827;">' + data.meetingId + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;vertical-align:top;">Organized By</td>' +
            '<td style="padding:8px 0;">' + data.employeeName + '<br><span style="font-size:11.5px;color:#6B7280;">' + (data.designation || '') + ' &nbsp;|&nbsp; ' + (data.district || '') + '</span></td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;vertical-align:top;">Stakeholder</td>' +
            '<td style="padding:8px 0;font-weight:600;">' + data.adhikariName + '<br><span style="font-size:11.5px;color:#6B7280;font-weight:400;">' + data.adhikariPost + '</span></td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Meeting Type</td>' +
            '<td style="padding:8px 0;">' + (data.meetingType || '-') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Purpose</td>' +
            '<td style="padding:8px 0;">' + (data.purpose || '-') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Conducted On</td>' +
            '<td style="padding:8px 0;font-weight:600;color:#111827;">' + (data.conductDate || '-') + (data.conductTime ? ' &nbsp;at&nbsp; ' + data.conductTime : '') + '</td>' +
          '</tr>' +
          (followUpId ? '<tr><td style="padding:8px 0;color:#6B7280;vertical-align:top;">Follow-up</td>' +
            '<td style="padding:8px 0;font-weight:600;color:#1D4ED8;">Meeting Scheduled &nbsp;|&nbsp; ' + (data.followUp && data.followUp.date ? data.followUp.date : '') + '</td></tr>' : '') +
        '</table>' +
      '</div>' +
    '</div>' +

    // Key Discussion Points
    '<div style="padding:0 28px 18px;background:#fff;">' +
      '<div style="background:#F0FDF4;border:1px solid #BBF7D0;border-left:4px solid #16A34A;border-radius:8px;padding:18px 20px;">' +
        '<p style="margin:0 0 12px;font-size:11px;font-weight:700;color:#15803D;text-transform:uppercase;letter-spacing:1px;">Key Discussion Points</p>' +
        '<table style="width:100%;border-collapse:collapse;">' + kpHtml + '</table>' +
      '</div>' +
    '</div>' +

    // MoM Doc & Photos links
    (momUrl || photoFolderUrl ?
    '<div style="padding:0 28px 18px;background:#fff;">' +
      '<div style="background:#EFF6FF;border:1px solid #BFDBFE;border-left:4px solid #2563EB;border-radius:8px;padding:16px 20px;">' +
        '<p style="margin:0 0 12px;font-size:11px;font-weight:700;color:#1D4ED8;text-transform:uppercase;letter-spacing:1px;">Documents & Resources</p>' +
        (momUrl ? '<p style="margin:0 0 8px;font-size:13px;color:#374151;">📄 &nbsp;<a href="' + momUrl + '" style="color:#2563EB;font-weight:600;text-decoration:none;">View Full MoM Document</a></p>' : '') +
        (photoFolderUrl ? '<p style="margin:0;font-size:13px;color:#374151;">📷 &nbsp;<a href="' + photoFolderUrl + '" style="color:#2563EB;font-weight:600;text-decoration:none;">View Meeting Photos</a></p>' : '') +
      '</div>' +
    '</div>' : '') +

    // Closing note
    '<div style="padding:0 28px 28px;background:#fff;">' +
      '<p style="font-size:13px;color:#6B7280;line-height:1.7;margin:0;">' +
        'Please retain this MoM for your records. For any discrepancies or additional inputs, ' +
        'kindly reach out to <strong>' + data.employeeName + '</strong> at the earliest.' +
      '</p>' +
    '</div>' +

    // Footer
    '<div style="background:#7B1010;padding:14px 28px;text-align:center;">' +
      '<p style="color:rgba(255,255,255,0.65);font-size:11.5px;margin:0;">' +
        'This is a system-generated notification from <strong style="color:#fff;">EG Meeting Management System</strong>.<br>' +
        'Educate Girls &nbsp;|&nbsp; Government Relations Team' +
      '</p>' +
    '</div>' +

  '</div>';

  MailApp.sendEmail({
    to:       colleague.email,
    subject:  subject,
    htmlBody: body
  });
}

// ------------------------------------------------------------
//  EMPLOYEE LOOKUP BY NAME
// ------------------------------------------------------------
function getEmployeeByName(name) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var nameLower = name.toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowName = data[i][2] ? data[i][2].toString().trim().toLowerCase() : '';
    if (rowName === nameLower) {
      return {
        name:        data[i][2] || '',
        designation: data[i][3] || '',
        email:       data[i][4] ? data[i][4].toString().trim() : ''
      };
    }
  }
  return null;
}

// ------------------------------------------------------------
//  DISTRICT REPORT - detailed breakdown for one district
// ------------------------------------------------------------
// ------------------------------------------------------------
//  REPORT DATA (public) - every meeting with district + block +
//  status + date, for the open Analytics Portal. Block is resolved
//  from the creator's email via Employee_DB. Client filters by
//  District × Block × Month (any combination, incl. "All").
// ------------------------------------------------------------
// ------------------------------------------------------------
//  MONTHLY REPORT - role-scoped (State→all, Zone→zone, District/Field→district).
//  All numbers computed here (exact); narrative is a template (Phase 1, no AI).
// ------------------------------------------------------------
var _RPT_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function monthKeyOf_(m) {
  var s = (m.status === 'Conducted' ? (m.conductDate || m.date) : m.date) || '';
  var p = s.toString().trim().split(' ');
  return p.length >= 3 ? (p[1] + ' ' + p[2]) : '';
}
function monthSortVal_(k) { var p = (k||'').split(' '); return (parseInt(p[1], 10) || 0) * 12 + _RPT_MONTHS.indexOf(p[0]); }
// The just-completed calendar month, e.g. run on 1 Sep -> "Aug 2026".
function prevMonthKey_() { var d = new Date(); d.setDate(1); d.setDate(0); return _RPT_MONTHS[d.getMonth()] + ' ' + d.getFullYear(); }

function getMonthlyReport(session, monthParam) {
  try {
    var role = (session && session.role || '').toString();

    // ── Scope from role ──
    var scopeDistricts = null, scopeKind = 'state', scopeLabel = 'Uttar Pradesh';
    if (role === 'Zone') {
      var zk = findZoneKey_(session.zone);
      scopeDistricts = zk ? ZONE_DISTRICTS[zk].slice() : [];
      scopeKind = 'zone'; scopeLabel = (session.zone || 'Zone');
    } else if (role === 'District' || role === 'Field') {
      scopeDistricts = (session.districts && session.districts.length) ? session.districts.slice() : [session.district];
      scopeKind = 'district'; scopeLabel = scopeDistricts.filter(String).join(', ');
    }
    function inScope(d) {
      if (!scopeDistricts) return true;
      for (var i = 0; i < scopeDistricts.length; i++) if (normDist_(scopeDistricts[i]) === normDist_(d)) return true;
      return false;
    }

    // ── Data (cached) ──
    var rd = getReportData();      var allM = (rd && rd.meetings)  ? rd.meetings  : [];
    var em = getEmployeeMaster();  var allE = (em && em.employees) ? em.employees : [];

    // available months (newest last)
    var mset = {};
    allM.forEach(function(m){ var k = monthKeyOf_(m); if (k) mset[k] = 1; });
    var months = Object.keys(mset).sort(function(a,b){ return monthSortVal_(a) - monthSortVal_(b); });
    var month = monthParam || months[months.length - 1] || '';

    // month + scope slice
    var mm = allM.filter(function(m){ return monthKeyOf_(m) === month && inScope(m.district); });
    var conducted = mm.filter(function(m){ return m.status === 'Conducted'; });
    var emps = allE.filter(function(e){ return e.district && inScope(e.district); });

    // active staff = distinct conductors this month
    var activeNames = {};
    conducted.forEach(function(m){ if (m.employeeName) activeNames[m.employeeName.trim().toLowerCase()] = 1; });
    var activeStaff = emps.filter(function(e){ return activeNames[(e.name||'').trim().toLowerCase()]; }).length;

    var pct = function(n, d){ return d ? Math.round(n / d * 100) : 0; };
    var success = pct(conducted.length, mm.length);
    var distsActive = {}, distsScope = {};
    conducted.forEach(function(m){ if (m.district) distsActive[normDist_(m.district)] = 1; });
    emps.forEach(function(e){ if (e.district) distsScope[normDist_(e.district)] = 1; });
    var pending = mm.filter(function(m){ return ['Planned','Follow-up','Postponed'].indexOf(m.status) !== -1; }).length;
    var govtMom = conducted.filter(function(m){ return (m.govtMom || '').toString().trim(); }).length;

    // ── Primary breakdown (state→zone, zone→district, district→block) ──
    function activeIn(list){ var s={}; list.forEach(function(m){ if(m.status==='Conducted' && m.employeeName) s[m.employeeName.trim().toLowerCase()]=1; }); return Object.keys(s).length; }
    function groupBy(keyFn){
      var g = {};
      mm.forEach(function(m){ var k = keyFn(m) || '-'; if (!g[k]) g[k] = { name:k, list:[], planned:0, conducted:0 }; g[k].planned++; g[k].list.push(m); if (m.status==='Conducted') g[k].conducted++; });
      return Object.keys(g).map(function(k){ var r=g[k]; return { name:r.name, planned:r.planned, conducted:r.conducted, pct:pct(r.conducted,r.planned), activeStaff:activeIn(r.list) }; });
    }
    var breakdown = { by:'', rows:[], leaderboard:[] };
    if (scopeKind === 'state') {
      breakdown.by = 'zone';
      var zg = {};
      mm.forEach(function(m){ var z = districtToZone_(m.district) || 'State'; if (!zg[z]) zg[z] = { name:z, planned:0, conducted:0, list:[] }; zg[z].planned++; zg[z].list.push(m); if (m.status==='Conducted') zg[z].conducted++; });
      breakdown.rows = Object.keys(ZONE_DISTRICTS).map(function(z){
        var r = zg[z] || { planned:0, conducted:0, list:[] };
        var dcount = (ZONE_DISTRICTS[z] ? ZONE_DISTRICTS[z].length : 0);
        return { name:z, districts:dcount, planned:r.planned, conducted:r.conducted, pct:pct(r.conducted,r.planned), activeStaff:activeIn(r.list) };
      }).sort(function(a,b){ return b.conducted - a.conducted; });
      // state-level row (zone-less districts like Lucknow), shown only when present
      if (zg['State']) {
        var uz = zg['State'];
        breakdown.rows.push({ name:'State', districts:1, planned:uz.planned, conducted:uz.conducted, pct:pct(uz.conducted,uz.planned), activeStaff:activeIn(uz.list) });
      }
      // district leaderboard: full list of districts with any conducted meeting (zero-conducted go to Attention Needed)
      breakdown.leaderboard = groupBy(function(m){ return m.district || '-'; })
        .map(function(r){ r.zone = districtToZone_(r.name) || 'State'; return r; })
        .filter(function(r){ return r.conducted > 0 && r.name !== '-'; })
        .sort(function(a,b){ return b.conducted - a.conducted; });
    } else if (scopeKind === 'zone') {
      breakdown.by = 'district';
      breakdown.rows = groupBy(function(m){ return m.district || '-'; }).sort(function(a,b){ return b.conducted - a.conducted; });
    } else {
      breakdown.by = 'block';
      breakdown.rows = groupBy(function(m){ return (m.block || '').trim() || 'District-level'; }).sort(function(a,b){ return b.conducted - a.conducted; });
    }

    // zero-activity areas (in scope, original names) - districts (or blocks for district scope)
    var zeroAreas = [];
    if (scopeKind === 'district') {
      var blockSet = {};
      emps.forEach(function(e){ if (e.block) blockSet[e.block] = 1; });
      var activeBlocks = {}; conducted.forEach(function(m){ if (m.block) activeBlocks[normDist_(m.block)] = 1; });
      zeroAreas = Object.keys(blockSet).filter(function(b){ return !activeBlocks[normDist_(b)]; });
    } else {
      var distNames = {};
      emps.forEach(function(e){ if (e.district) distNames[normDist_(e.district)] = e.district; });
      zeroAreas = Object.keys(distNames).filter(function(k){ return !distsActive[k]; }).map(function(k){ return distNames[k]; });
    }

    // focus
    function tally(list, key){ var o={}; list.forEach(function(m){ var k=(m[key]||'-').toString().trim()||'-'; o[k]=(o[k]||0)+1; }); return Object.keys(o).map(function(k){ return { name:k, count:o[k] }; }).sort(function(a,b){ return b.count-a.count; }).slice(0,6); }
    var byPurpose = tally(conducted, 'purpose');
    var byStakeholder = tally(conducted, 'stakeholderPost');

    // ── Outcomes ──
    // Every report so far counts activity: how many meetings happened. This
    // counts what came out of them, which is the actual job. The officer picks
    // the outcome on the conduct form, so this is reported, not inferred.
    var OUTCOME_CONCRETE = { Commitment:1, Information:1, Permission:1 };
    var ocCounts = {}, ocAnswered = 0, ocConcrete = 0;
    conducted.forEach(function(m){
      var o = (m.outcome || '').toString().trim();
      if (!o) return;
      ocAnswered++;
      ocCounts[o] = (ocCounts[o] || 0) + 1;
      if (OUTCOME_CONCRETE[o]) ocConcrete++;
    });
    var outcomes = {
      answered: ocAnswered,                    // older meetings predate the question
      concrete: ocConcrete,
      rate: pct(ocConcrete, ocAnswered),
      counts: Object.keys(ocCounts).map(function(k){ return { name:k, count:ocCounts[k] }; })
                    .sort(function(a,b){ return b.count - a.count; })
    };

    // ── Relationship health ──
    // Government relations is repeat contact, so each official is tracked as an
    // ongoing relationship rather than as a pile of separate meetings. This part
    // reads the whole history, not just the report month, but measures it as of
    // the END of the report month so an old report keeps reading the same way.
    var COLD_DAYS = 60;
    function relMonthEdge_(k, atEnd) {
      var p = (k||'').split(' '), mi = _RPT_MONTHS.indexOf(p[0]), yr = parseInt(p[1], 10);
      if (mi < 0 || isNaN(yr)) return 0;
      return atEnd ? new Date(yr, mi + 1, 0, 23, 59, 59).getTime() : new Date(yr, mi, 1).getTime();
    }
    function relTs_(m) {
      var s = ((m.conductDate || m.date) || '').toString().trim().split(' ');
      if (s.length < 3) return 0;
      var d = parseInt(s[0], 10), mi = _RPT_MONTHS.indexOf(s[1]), y = parseInt(s[2], 10);
      return (isNaN(d) || mi < 0 || isNaN(y)) ? 0 : new Date(y, mi, d).getTime();
    }
    var relRefTs   = relMonthEdge_(month, true)  || Date.now();
    var relStartTs = relMonthEdge_(month, false);
    var relMap = {};
    allM.forEach(function(m){
      if (m.status !== 'Conducted' || !inScope(m.district)) return;
      // Keyed on the POST, not the person. Officers get transferred, but the
      // office carries the relationship: the BSA of a district stays the BSA.
      // Falls back to the person's name only when no post was recorded.
      var post = (m.stakeholderPost||'').toString().trim();
      var person = (m.stakeholderName||'').toString().trim();
      var pk = _prepNorm_(post) || _prepNorm_(person);
      if (!pk) return;
      var ts = relTs_(m);
      if (!ts || ts > relRefTs) return;                 // nothing after this report's month
      var k = normDist_(m.district) + '|' + pk;
      var r = relMap[k] || (relMap[k] = { post:(post || person), district:(m.district||'').toString().trim(),
                                          lastPerson:'', people:{}, count:0, first:ts, last:0 });
      r.count++;
      if (person) r.people[_prepNorm_(person)] = 1;
      if (ts < r.first) r.first = ts;
      if (ts > r.last) { r.last = ts; if (person) r.lastPerson = person; }
    });
    Object.keys(relMap).forEach(function(k){ var r = relMap[k]; r.peopleCount = Object.keys(r.people).length; delete r.people; });
    var relAll  = Object.keys(relMap).map(function(k){ return relMap[k]; });
    var relCold = relAll.filter(function(r){ return r.last && (relRefTs - r.last) > COLD_DAYS * 86400000; })
                        .sort(function(a,b){ return a.last - b.last; });
    var relSeen = {};
    conducted.forEach(function(m){
      var pk = _prepNorm_(m.stakeholderPost) || _prepNorm_(m.stakeholderName);
      if (pk) relSeen[normDist_(m.district) + '|' + pk] = 1;
    });
    var relationships = {
      totalOffices:   relAll.length,
      metThisMonth:   Object.keys(relSeen).length,
      newThisMonth:   relStartTs ? relAll.filter(function(r){ return r.first >= relStartTs; }).length : 0,
      onlyOnce:       relAll.filter(function(r){ return r.count === 1; }).length,
      coldDays:       COLD_DAYS,
      coldCount:      relCold.length,
      cold:           relCold.slice(0, 10).map(function(r){
                        return { post:r.post, district:r.district, lastPerson:r.lastPerson,
                                 peopleCount:r.peopleCount, count:r.count,
                                 daysSince: Math.round((relRefTs - r.last) / 86400000) };
                      })
    };

    // ── Attention (rules) ──
    var attention = [];
    var areaWord = scopeKind === 'district' ? 'blocks' : 'districts';
    if (zeroAreas.length) attention.push({ level:'crit', title:zeroAreas.length + ' ' + areaWord + ' with no activity', detail:zeroAreas.slice(0,8).join(', ') + (zeroAreas.length>8?' +more':'') });
    if (pending) attention.push({ level:'warn', title:pending + ' follow-ups / planned meetings pending', detail:'Open in ' + month });
    if (conducted.length && govtMom < conducted.length) attention.push({ level:'warn', title:'Govt MoM pending on ' + (conducted.length - govtMom) + ' of ' + conducted.length, detail:'Only ' + govtMom + ' conducted meetings have official minutes uploaded' });
    if (relationships.coldCount) attention.push({ level:'warn', title:relationships.coldCount + ' offices with no contact in ' + COLD_DAYS + '+ days', detail:relationships.cold.slice(0,4).map(function(r){ return r.post + ' ' + r.district + ' (' + r.daysSince + 'd)'; }).join(', ') + (relationships.coldCount > 4 ? ' +more' : '') });
    if (scopeKind === 'state' && breakdown.rows.length) {
      var worst = breakdown.rows.slice().sort(function(a,b){ return a.pct - b.pct; })[0];
      if (worst) attention.push({ level:'warn', title:worst.name + ' is the lowest-performing zone (' + worst.pct + '%)', detail:worst.conducted + ' of ' + worst.planned + ' conducted' });
    }

    // ── Narrative (Phase 1 template - grounded in the numbers) ──
    var best = breakdown.rows[0];
    var topPerf = breakdown.leaderboard[0] || breakdown.rows[0];
    var summary = 'In ' + month + ', ' + scopeLabel + ' conducted ' + conducted.length + ' of ' + mm.length +
      ' planned meetings (' + success + '% success rate)' + (best ? ', led by ' + best.name + ' (' + best.pct + '%)' : '') + '. ' +
      'Staff participation stood at ' + pct(activeStaff, emps.length) + '% (' + activeStaff + ' of ' + emps.length + ' active)' +
      (zeroAreas.length ? ', and ' + zeroAreas.length + ' ' + areaWord + ' recorded no activity' : '') + '.';
    var highlights = [];
    if (topPerf) highlights.push({ h: topPerf.name + ' led with ' + topPerf.conducted + ' conducted (' + topPerf.pct + '%)', d: 'Strongest ' + (scopeKind==='state'?'district':breakdown.by) + ' this month.' });
    if (byPurpose[0]) highlights.push({ h: 'Focus on ' + byPurpose[0].name + ' (' + byPurpose[0].count + ' meetings)', d: 'Most common meeting purpose.' });
    if (byStakeholder[0]) highlights.push({ h: 'Most engaged: ' + byStakeholder[0].name + ' (' + byStakeholder[0].count + ')', d: 'Top government stakeholder met.' });
    var recs = [];
    if (zeroAreas.length) recs.push({ h:'Activate the ' + zeroAreas.length + ' inactive ' + areaWord + ' first.', d:zeroAreas.slice(0,5).join(', ') + ' had no conducted meetings.' });
    if (pending) recs.push({ h:'Close the ' + pending + ' pending follow-ups.', d:'Convert planned/postponed meetings before month-end.' });
    if (conducted.length && govtMom < conducted.length) recs.push({ h:'Push Govt MoM collection.', d:'Only ' + pct(govtMom, conducted.length) + '% of meetings have official minutes.' });
    if (relationships.coldCount) recs.push({ h:'Re-engage the ' + relationships.coldCount + ' offices that have gone quiet.', d:'No contact for over ' + COLD_DAYS + ' days. Longest gap: ' + (relationships.cold[0] ? relationships.cold[0].post + ' ' + relationships.cold[0].district + ', ' + relationships.cold[0].daysSince + ' days' : '') + '.' });
    if (relationships.onlyOnce && relationships.totalOffices) recs.push({ h:'Build depth, not just reach.', d:relationships.onlyOnce + ' of ' + relationships.totalOffices + ' offices have been met only once. Repeat contact is what moves government work.' });

    var resp = {
      success: true,
      scope: { kind:scopeKind, label:scopeLabel, role:role, month:month, generatedAt:new Date().toISOString() },
      months: months,
      kpis: { total:mm.length, conducted:conducted.length, success:success,
              totalStaff:emps.length, activeStaff:activeStaff, participation:pct(activeStaff, emps.length),
              distsActive:Object.keys(distsActive).length, distsInScope:Object.keys(distsScope).length,
              pending:pending, govtMom:govtMom },
      breakdown: breakdown,
      byPurpose: byPurpose, byStakeholder: byStakeholder,
      relationships: relationships, outcomes: outcomes,
      zeroAreas: zeroAreas, attention: attention,
      narrative: { ai:false, summary:summary, highlights:highlights, recommendations:recs }
    };

    // ── AI narrative (Phase 2) - aggregated numbers only; cached; template fallback ──
    try {
      var aiKey = 'aiNarr_' + scopeKind + '_' + normDist_(scopeLabel).slice(0,40) + '_' + month.replace(/\s/g,'');
      var cached = cGet(aiKey);
      if (cached && cached.summary) {
        resp.narrative = { ai:true, summary:cached.summary, highlights:cached.highlights, recommendations:cached.recommendations };
      } else {
        var ai = aiReportNarrative(resp);
        if (ai && ai.summary) {
          resp.narrative = { ai:true, summary:ai.summary, highlights:ai.highlights, recommendations:ai.recommendations };
          cPut(aiKey, ai, 21600);   // 6 h - regenerated a few times/day at most
        }
      }
    } catch (aiErr) { /* keep template narrative */ }

    return resp;
  } catch (err) {
    return { success:false, message: err.message };
  }
}

// ------------------------------------------------------------
//  AI NARRATIVE - Mistral (primary) → Gemini (fallback) → null.
//  Keys live in Script Properties (MISTRAL_KEY / GEMINI_KEY), never in code.
//  Prompt contains ONLY aggregated numbers + area/purpose names (no person
//  names, no meeting notes) - the agreed privacy stance.
// ------------------------------------------------------------
// Pull the answer out of a Gemini response. A thinking model can return parts
// marked as thoughts, or an empty content block when it ran out of budget while
// reasoning, so take the first real text part and never assume parts[0].
function _geminiText_(j) {
  try {
    var c = j && j.candidates && j.candidates[0];
    if (!c || !c.content || !c.content.parts) return '';
    for (var i = 0; i < c.content.parts.length; i++) {
      var p = c.content.parts[i];
      if (p && p.text && !p.thought) return p.text;
    }
  } catch(e) {}
  return '';
}

// callLLM hides every failure on purpose so features degrade quietly. This does
// the opposite: it sends a trivial prompt to each provider and reports exactly
// what came back, so a dead key, a wrong model name or a quota block is visible.
function LLM_probe() {
  var props = PropertiesService.getScriptProperties(), out = {};

  var mk = props.getProperty('MISTRAL_KEY');
  out.mistral_key = mk ? ('present, length ' + mk.length) : 'MISSING';
  if (mk) {
    try {
      var r = UrlFetchApp.fetch('https://api.mistral.ai/v1/chat/completions', {
        method:'post', contentType:'application/json', muteHttpExceptions:true,
        headers:{ Authorization:'Bearer ' + mk },
        payload: JSON.stringify({ model:'mistral-small-latest', messages:[{role:'user', content:'Reply with the single word OK.'}], max_tokens:10 })
      });
      out.mistral_httpCode = r.getResponseCode();
      out.mistral_reply    = r.getContentText().substring(0, 400);
    } catch(e) { out.mistral_threw = e.message; }
  }

  var gk = props.getProperty('GEMINI_KEY');
  out.gemini_key = gk ? ('present, length ' + gk.length) : 'MISSING';
  if (gk) {
    try {
      var r2 = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + encodeURIComponent(gk), {
        method:'post', contentType:'application/json', muteHttpExceptions:true,
        payload: JSON.stringify({ contents:[{parts:[{text:'Reply with the single word OK.'}]}],
                                  generationConfig:{ maxOutputTokens:2000, temperature:0.2, thinkingConfig:{ thinkingBudget:0 } } })
      });
      out.gemini_httpCode    = r2.getResponseCode();
      out.gemini_extractedText = _geminiText_(JSON.parse(r2.getContentText() || '{}')) || '(nothing extracted)';
      out.gemini_reply       = r2.getContentText().substring(0, 400);
    } catch(e) { out.gemini_threw = e.message; }
  }

  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// Both providers fail temporarily under load: Mistral answers 429 (free-tier
// rate limit) and Gemini answers 503 (model busy). Both clear within seconds,
// so try each provider, wait, and go round again before giving up.
function callLLM(prompt) {
  var props = PropertiesService.getScriptProperties();
  var mk = props.getProperty('MISTRAL_KEY');
  var gk = props.getProperty('GEMINI_KEY');

  function mistral() {
    if (!mk) return '';
    try {
      var r = UrlFetchApp.fetch('https://api.mistral.ai/v1/chat/completions', {
        method:'post', contentType:'application/json', muteHttpExceptions:true,
        headers:{ Authorization:'Bearer ' + mk },
        payload: JSON.stringify({ model:'mistral-small-latest', messages:[{role:'user', content:prompt}], max_tokens:900, temperature:0.3 })
      });
      if (r.getResponseCode() === 200) {
        var j = JSON.parse(r.getContentText());
        var t = j && j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content;
        if (t) return t;
      }
    } catch(e) {}
    return '';
  }

  function gemini() {
    if (!gk) return '';
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + encodeURIComponent(gk);
    // gemini-3.6-flash reasons before it answers, and that reasoning is charged
    // against maxOutputTokens. Left alone it can spend the whole budget thinking
    // and return an empty content block with finishReason MAX_TOKENS, which is
    // exactly why this fallback used to look dead. Turn thinking off and leave
    // plenty of room for the answer itself.
    function ask(useThinkingConfig) {
      var cfg = { maxOutputTokens:8000, temperature:0.3 };
      if (useThinkingConfig) cfg.thinkingConfig = { thinkingBudget: 0 };
      try {
        var r2 = UrlFetchApp.fetch(url, {
          method:'post', contentType:'application/json', muteHttpExceptions:true,
          payload: JSON.stringify({ contents:[{parts:[{text:prompt}]}], generationConfig:cfg })
        });
        if (r2.getResponseCode() !== 200) return null;   // null = try the other shape
        return _geminiText_(JSON.parse(r2.getContentText()));
      } catch(e) { return null; }
    }
    var t = ask(true);
    if (t === null) t = ask(false);   // older API shape rejects thinkingConfig
    return t || '';
  }

  // Two rounds only: enough to ride out a brief limit, bounded enough that a
  // batch job cannot run into the 6 minute Apps Script ceiling.
  var waits = [0, 2500];
  for (var a = 0; a < waits.length; a++) {
    if (waits[a]) Utilities.sleep(waits[a]);
    var g = gemini();  if (g) return g;   // Gemini first: Mistral's free tier runs out
    var t = mistral(); if (t) return t;
  }
  return '';
}

function buildReportPrompt(r) {
  var k = r.kpis, sc = r.scope, b = r.breakdown, L = [];
  L.push('Scope: ' + sc.label + ' (' + sc.kind + ' level), Month: ' + sc.month);
  L.push('Meetings: ' + k.total + ' planned, ' + k.conducted + ' conducted, ' + k.success + '% success rate');
  L.push('Staff participation: ' + k.activeStaff + ' active of ' + k.totalStaff + ' (' + k.participation + '%)');
  var byLabel = { zone:'zone', district:'district', block:'block' }[b.by] || 'area';
  if (b.rows && b.rows.length) L.push('By ' + byLabel + ': ' + b.rows.slice(0,12).map(function(x){ return x.name + ' ' + x.planned + '/' + x.conducted + '/' + x.pct + '%'; }).join('; '));
  if (b.leaderboard && b.leaderboard.length) L.push('Top districts by conducted: ' + b.leaderboard.slice(0,5).map(function(x){ return x.name + ' ' + x.conducted + ' (' + x.pct + '%)'; }).join(', '));
  if (r.byPurpose && r.byPurpose.length) L.push('Meeting purposes: ' + r.byPurpose.map(function(x){ return x.name + ' ' + x.count; }).join(', '));
  if (r.byStakeholder && r.byStakeholder.length) L.push('Stakeholder types met: ' + r.byStakeholder.map(function(x){ return x.name + ' ' + x.count; }).join(', '));
  if (r.outcomes && r.outcomes.answered) {
    var oc = r.outcomes;
    L.push('Outcomes (reported by the officer, not inferred): ' + oc.concrete + ' of ' + oc.answered + ' meetings produced something concrete (' + oc.rate + '%), broken down as ' + oc.counts.map(function(x){ return x.name + ' ' + x.count; }).join(', '));
  }
  // counts only, never the officials' names: this prompt leaves the org
  if (r.relationships) {
    var rl = r.relationships;
    L.push('Relationships (tracked by office, not by person, since officials transfer): ' + rl.totalOffices + ' offices engaged to date, ' + rl.metThisMonth + ' met this month, ' +
           rl.newThisMonth + ' engaged for the first time, ' + rl.onlyOnce + ' have been met only once, ' +
           rl.coldCount + ' have had no contact for over ' + rl.coldDays + ' days');
  }
  var att = [];
  if (r.zeroAreas && r.zeroAreas.length) att.push(r.zeroAreas.length + ' ' + byLabel + 's with no activity');
  att.push(k.pending + ' follow-ups pending');
  att.push('Govt MoM received on ' + k.govtMom + ' of ' + k.conducted + ' conducted');
  L.push('Attention: ' + att.join('; '));
  return "You are writing a concise monthly report for Educate Girls' government-relations meeting tracker. Use ONLY the data below. Write in professional Indian English. Do not invent numbers or names. Do not use em dashes; use commas or hyphens. Return STRICT JSON only (no markdown fences), exactly this shape: {\"summary\":\"2 to 3 sentences\",\"highlights\":[{\"h\":\"short headline\",\"d\":\"one detail sentence\"}],\"recommendations\":[{\"h\":\"action\",\"d\":\"why or how\"}]}. Give 2 to 3 highlights and 2 to 3 recommendations.\nDATA:\n" + L.join('\n');
}

function aiReportNarrative(r) {
  var raw = callLLM(buildReportPrompt(r));
  if (!raw) return null;
  raw = raw.replace(/```json/gi, '').replace(/```/g, '').trim();
  var s = raw.indexOf('{'), e = raw.lastIndexOf('}');
  if (s < 0 || e < 0) return null;
  try {
    var o = JSON.parse(raw.slice(s, e + 1));
    if (!o || !o.summary) return null;
    function norm(a){ return (a||[]).map(function(x){ return (typeof x === 'string') ? { h:x, d:'' } : { h:(x.h||''), d:(x.d||'') }; }).filter(function(x){ return x.h; }); }
    return { summary: String(o.summary), highlights: norm(o.highlights), recommendations: norm(o.recommendations) };
  } catch(e) { return null; }
}

// ============================================================
//  TIER 2 - AI MEETING TAGGING (notes + Govt MoM)
//  Reads each conducted meeting's key-points (and Govt MoM PDF) and writes
//  Priority / Flag / Next Action / Escalate / Category / MoM-summary into
//  the Conducted sheet. Free (Mistral for text, Gemini for the PDF).
//  Conducted sheet new columns (1-based): W..AC = 23..29.
// ============================================================
var COL_TAG_PRIORITY=23, COL_TAG_FLAG=24, COL_TAG_NEXT=25, COL_TAG_ESC=26, COL_TAG_CAT=27, COL_TAG_MOMSUM=28, COL_TAG_AT=29;

function _parseJson_(raw){ if(!raw) return null; raw=raw.replace(/```json/gi,'').replace(/```/g,'').trim(); var s=raw.indexOf('{'),e=raw.lastIndexOf('}'); if(s<0||e<0)return null; try{return JSON.parse(raw.slice(s,e+1));}catch(err){return null;} }

// Read a Govt MoM PDF from Drive with Gemini (multimodal). Returns a short summary or ''.
// Same Gemini call as callLLM after the September fix. gemini-3.6-flash charges
// its reasoning to maxOutputTokens, so a small budget with thinking left on can
// come back HTTP 200 with an empty content block (finishReason MAX_TOKENS), and
// parts[0] is not always the answer. Thinking off, room to answer, a second try
// without thinkingConfig if it is refused, and _geminiText_ to pick the text.
function readGovtMomPdf_(url) {
  try {
    var m = (url||'').toString().match(/[-\w]{25,}/); if (!m) return '';
    var gk = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY'); if (!gk) return '';
    var b64 = Utilities.base64Encode(DriveApp.getFileById(m[0]).getBlob().getBytes());
    var prompt = 'This is an official Government Minutes of Meeting, possibly Hindi, English, scanned or handwritten. In 2 short lines plus up to 3 action items with any deadlines, summarize the key government commitments. Plain text, do not use em dashes.';
    var api = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + encodeURIComponent(gk);
    function ask(useThinkingConfig) {
      var cfg = { maxOutputTokens:8000, temperature:0.2 };
      if (useThinkingConfig) cfg.thinkingConfig = { thinkingBudget: 0 };
      try {
        var res = UrlFetchApp.fetch(api, {
          method:'post', contentType:'application/json', muteHttpExceptions:true,
          payload: JSON.stringify({ contents:[{parts:[{text:prompt},{inline_data:{mime_type:'application/pdf',data:b64}}]}], generationConfig:cfg })
        });
        if (res.getResponseCode() !== 200) return null;   // null = try the other shape
        return _geminiText_(JSON.parse(res.getContentText()));
      } catch(e) { return null; }
    }
    var t = ask(true);
    if (t === null) t = ask(false);   // older API shape rejects thinkingConfig
    return (t||'').trim();
  } catch(e){}
  return '';
}

// How many conducted meetings carry a Govt MoM link, and how many were tagged
// but still have an empty MoM summary. Editor only: writes nothing, calls no
// model, logs counts only. The empty ones are split by cause, because only one
// cause is the Gemini bug above. A MoM uploaded AFTER its row was tagged was
// never read at all, since tagUntaggedMeetings skips rows already tagged, and
// fixing readGovtMomPdf_ does nothing for those.
function MOM_measure() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'no conducted sheet' };
  var data = sh.getDataRange().getValues();
  var res = { conductedMeetings:0, withGovtMomLink:0, notTaggedYet:0, tagged:0,
              taggedWithSummary:0, taggedEmptySummary:0,
              emptyBecause:{ noFileIdInLink:0, fileCannotBeOpened:0, uploadedAfterTagging:0, presentWhenTagged:0 },
              presentWhenTaggedOver15MB:0 };
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    res.conductedMeetings++;
    var link = (data[i][21]||'').toString().trim();              // V = Govt MoM
    if (!link) continue;
    res.withGovtMomLink++;
    var at = data[i][COL_TAG_AT-1];
    if (!(at||'').toString().trim()) { res.notTaggedYet++; continue; }
    res.tagged++;
    if ((data[i][COL_TAG_MOMSUM-1]||'').toString().trim()) { res.taggedWithSummary++; continue; }
    res.taggedEmptySummary++;
    // the same first link and the same id pattern readGovtMomPdf_ uses
    var m = link.split(/\s*,\s*/)[0].match(/[-\w]{25,}/);
    if (!m) { res.emptyBecause.noFileIdInLink++; continue; }
    var f;
    try { f = DriveApp.getFileById(m[0]); f.getDateCreated(); }
    catch(e) { res.emptyBecause.fileCannotBeOpened++; continue; }
    var taggedAt = (at instanceof Date) ? at : new Date(at);
    if (!isNaN(taggedAt.getTime()) && f.getDateCreated() > taggedAt) { res.emptyBecause.uploadedAfterTagging++; continue; }
    res.emptyBecause.presentWhenTagged++;
    if (f.getSize() > 15 * 1024 * 1024) res.presentWhenTaggedOver15MB++;
  }
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// One live check of readGovtMomPdf_ after the fix: the newest meeting with a
// Govt MoM link and an empty summary (or, if none, the newest with a link).
// One Gemini call, writes nothing, logs only the length of what came back.
function MOM_tryOne() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'no conducted sheet' };
  var data = sh.getDataRange().getValues(), pick = -1, fallback = -1;
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][0] || !(data[i][21]||'').toString().trim()) continue;
    if (fallback < 0) fallback = i;
    if (!(data[i][COL_TAG_MOMSUM-1]||'').toString().trim()) { pick = i; break; }
  }
  if (pick < 0) pick = fallback;
  if (pick < 0) { Logger.log('No conducted meeting has a Govt MoM link.'); return { success:false }; }
  var t0 = Date.now();
  var s = readGovtMomPdf_((data[pick][21]||'').toString().split(/\s*,\s*/)[0]);
  var res = { row:pick + 1, hadEmptySummary:!(data[pick][COL_TAG_MOMSUM-1]||'').toString().trim(),
              summaryChars:s.length, seconds:Math.round((Date.now() - t0) / 100) / 10 };
  Logger.log(JSON.stringify(res, null, 2) + (s ? '' : '\nNothing came back. Run LLM_probe to see what Gemini is answering.'));
  return res;
}

// ============================================================
//  VOICE NOTE TEST - editor only, writes nothing
// ============================================================
// Before any mic button goes on the conduct form, find out whether Gemini can
// actually hear a UP field officer: local accent, Hinglish, traffic behind
// them. Put a few phone recordings in a Drive folder called
// VOICE TEST, pick VOICE_test in the editor and press Run, then read the log.
// (A folder link can be passed too, but the Run button cannot pass one.) Each file is sent with the same instruction the real
// feature would use. The log shows the word-for-word transcript, the three
// answers it would put in the boxes, how long it took, and how much of each
// answer can be found in the transcript. A low figure there means the model
// wrote something that was never said, which is the thing to watch for.
// Nothing is saved, no sheet is touched, and the live app does not change.
var VOICE_TEST_FOLDER = 'VOICE TEST';
var VOICE_MIMES = {
  m4a:['audio/mp4','audio/aac','audio/x-m4a'], mp4:['audio/mp4','audio/aac'],
  aac:['audio/aac'], mp3:['audio/mp3','audio/mpeg'], wav:['audio/wav'],
  ogg:['audio/ogg'], oga:['audio/ogg'], opus:['audio/ogg','audio/opus'],
  webm:['audio/webm','audio/ogg'], flac:['audio/flac'], aiff:['audio/aiff'],
  amr:['audio/amr'], '3gp':['audio/3gpp']
};
var VOICE_PROMPT = [
  'You will hear a field officer from Uttar Pradesh describing a meeting they just had with a government official.',
  'They may speak Hindi, English or a mix of both, with a local accent, and there may be traffic or other noise.',
  'Return STRICT JSON only, no markdown:',
  '{"transcript":"...","what":"...","said":"...","next":"...","unclear":"..."}',
  'transcript: exactly what was said, word for word, in the language and script it was spoken. Hindi words in Devanagari, English words in English. Do not translate and do not tidy it into better sentences.',
  'what: what was discussed in the meeting.',
  'said: what the official said, agreed to or refused.',
  'next: what the officer or the team will do next.',
  'unclear: any words you could not make out, or an empty string.',
  'Rules: use only what was actually said. If something was not mentioned, leave that field as an empty string.',
  'Never add a detail, date, number or name that was not spoken. Write what, said and next in the language the officer used.',
  'For what, said and next, keep to the officer\'s own words as far as possible rather than rephrasing them.',
  'Do not use em dashes.'
].join(String.fromCharCode(10));

function voiceWords_(s) {
  return (s || '').toString().toLowerCase()
    .split(/[^a-z0-9ऀ-ॿ]+/).filter(function(w){ return w.length >= 3; });
}
// Share of an answer's words that also occur in the transcript, 0 to 100,
// or -1 when the answer is empty.
function voicePct_(part, transcript) {
  var w = voiceWords_(part);
  if (!w.length) return -1;
  var have = {};
  voiceWords_(transcript).forEach(function(x){ have[x] = 1; });
  return Math.round(100 * w.filter(function(x){ return have[x]; }).length / w.length);
}
function voiceFound_(part, transcript) {
  var p = voicePct_(part, transcript);
  return p < 0 ? '-' : p + '%';
}

// Same shape as callLLM's Gemini call after the September fix: thinking off,
// room to answer, and a second try without thinkingConfig if it is refused.
function voiceAsk_(gk, b64, mime) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + encodeURIComponent(gk);
  var last = '';
  for (var t = 0; t < 2; t++) {
    var cfg = { maxOutputTokens:8000, temperature:0.2, responseMimeType:'application/json' };
    if (t === 0) cfg.thinkingConfig = { thinkingBudget:0 };
    var r = UrlFetchApp.fetch(url, {
      method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({ contents:[{ parts:[{ text:VOICE_PROMPT }, { inline_data:{ mime_type:mime, data:b64 } }] }],
                                generationConfig:cfg })
    });
    if (r.getResponseCode() === 200) return { ok:true, text:_geminiText_(JSON.parse(r.getContentText())) };
    last = r.getResponseCode() + ' ' + r.getContentText().substring(0, 200);
  }
  return { ok:false, error:last };
}

function VOICE_test(folder) {
  var gk = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!gk) { Logger.log('GEMINI_KEY is not set in Script Properties.'); return; }
  // With no link, look for the folder by name in the Drive of whoever pressed
  // Run, because the editor's Run button cannot pass an argument.
  var dir = null, id = ((folder || '').toString().match(/[-\w]{25,}/) || [])[0];
  if (id) {
    try { dir = DriveApp.getFolderById(id); }
    catch (e) { Logger.log('Cannot open that folder from the account running this script. Share the folder with it and run again. (' + e.message + ')'); return; }
  } else {
    var found = DriveApp.getFoldersByName(VOICE_TEST_FOLDER);
    if (!found.hasNext()) {
      Logger.log('No folder called "' + VOICE_TEST_FOLDER + '" in the Drive of the account running this script. Make one there, put the recordings in it, and run VOICE_test again.');
      return;
    }
    dir = found.next();
  }

  var files = dir.getFiles(), done = 0, NL = String.fromCharCode(10);
  while (files.hasNext() && done < 8) {
    var f = files.next(), name = f.getName();
    var ext = (name.split('.').pop() || '').toLowerCase();
    var own = (f.getMimeType() || '').toLowerCase();
    var tries = (VOICE_MIMES[ext] || []).slice();
    if (own.indexOf('audio/') === 0 && tries.indexOf(own) === -1) tries.unshift(own);
    if (!tries.length) { Logger.log('SKIP ' + name + ' (not an audio file: ' + own + ')'); continue; }
    if (f.getSize() > 15 * 1024 * 1024) { Logger.log('SKIP ' + name + ' (over 15MB)'); continue; }
    done++;

    var b64 = Utilities.base64Encode(f.getBlob().getBytes());
    var res = null, used = '', t0 = Date.now(), errors = [];
    for (var i = 0; i < tries.length; i++) {
      res = voiceAsk_(gk, b64, tries[i]);
      if (res.ok) { used = tries[i]; break; }
      errors.push(tries[i] + ': ' + res.error);
    }
    var secs = Math.round((Date.now() - t0) / 100) / 10;
    var head = '===== ' + name + '  (' + Math.round(f.getSize() / 1024) + ' KB, ' + secs + ' s)';
    if (!res || !res.ok) { Logger.log(head + NL + 'FORMAT NOT ACCEPTED' + NL + errors.join(NL)); continue; }

    var o = null;
    try { o = JSON.parse((res.text || '').replace(/^```(?:json)?\s*|\s*```$/g, '')); } catch (pe) {}
    if (!o) { Logger.log(head + '  sent as ' + used + NL + 'The answer was not JSON:' + NL + (res.text || '(empty)')); continue; }
    Logger.log([
      head + '  sent as ' + used,
      'TRANSCRIPT: ' + (o.transcript || '(empty)'),
      '',
      'WHAT  [' + voiceFound_(o.what, o.transcript) + ' found in transcript]  ' + (o.what || '-'),
      'SAID  [' + voiceFound_(o.said, o.transcript) + ' found in transcript]  ' + (o.said || '-'),
      'NEXT  [' + voiceFound_(o.next, o.transcript) + ' found in transcript]  ' + (o.next || '-'),
      'UNCLEAR: ' + (o.unclear || '-')
    ].join(NL));
  }
  if (!done) Logger.log('No audio files found in that folder.');
}

// ------------------------------------------------------------
//  VOICE NOTE - speak the three answers instead of typing them
// ------------------------------------------------------------
// The officer records on the conduct form. The page turns the recording into
// 16 kHz mono WAV first, because WAV is a format Gemini documents and WebM,
// which Android Chrome records in, is not; that removes the guess. This hands
// it to Gemini with the same instruction VOICE_test uses and returns the three
// answers, the word-for-word transcript, and how much of each answer can be
// found in that transcript, so the page can warn when something was written
// that was never said.
//
// NOTHING IS KEPT. The audio is not put in Drive or anywhere else, and nothing
// is written to the sheet. The answers only land in the boxes; the officer
// still reads them and presses Save. The outcome question is deliberately not
// answered here: the officer was in the room, and they are the one asked.
var VOICE_MAX_B64 = 8 * 1024 * 1024;   // a little over two minutes of 16 kHz WAV

function transcribeVoice(session, audio, mime) {
  audio = (audio || '').toString();
  mime  = (mime || 'audio/wav').toString();
  if (!audio) return { success:false, message:'No recording arrived. Please try again, or type it.' };
  if (audio.length > VOICE_MAX_B64) return { success:false, message:'That recording is too long. Please keep it under two minutes.' };
  if (!/^audio\/[a-z0-9.+-]+$/i.test(mime)) mime = 'audio/wav';

  var gk = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!gk) return { success:false, message:'Voice is not set up yet. Please type it for now.' };

  var r;
  try { r = voiceAsk_(gk, audio, mime); }
  catch (e) { return { success:false, message:'Could not reach the voice service. Please type it for now.' }; }
  if (!r.ok) return { success:false, message:'The voice service did not answer. Please try again in a minute, or type it.' };

  var o = null;
  try { o = JSON.parse((r.text || '').replace(/^```(?:json)?\s*|\s*```$/g, '')); } catch (pe) {}
  if (!o || !(o.transcript || '').toString().trim()) {
    return { success:false, message:'Could not make out the recording. Try again somewhere quieter, or type it.' };
  }

  var EM = new RegExp(String.fromCharCode(8212), 'g');   // house style: no em dashes
  function clean(x) { return (x || '').toString().replace(EM, ',').trim(); }
  var t = clean(o.transcript);
  var out = { success:true, transcript:t, what:clean(o.what), said:clean(o.said),
              next:clean(o.next), unclear:clean(o.unclear) };
  out.found = { what:voicePct_(out.what, t), said:voicePct_(out.said, t), next:voicePct_(out.next, t) };
  return out;
}

// ------------------------------------------------------------
//  ASK THE MEETINGS - the chatbot on the analytics portal
// ------------------------------------------------------------
// A question in plain Hindi, English or Hinglish, answered from the meetings
// the person is allowed to see. Only for people who are signed in: the portal
// itself is public, and this reads the notes, which have deliberately never
// been published.
//
// Two jobs are kept apart on purpose.
//  * COUNTING is done here, in code, and handed to the model as exact figures.
//    A model asked to count across three hundred rows gets it wrong, and a
//    wrong number is the answer leadership will repeat.
//  * READING, what was discussed and what was promised, is left to the model,
//    with each meeting laid out on one line and its ID in front.
//
// Retrieval: while the scoped data fits the budget, every question sees every
// meeting, so nothing can be missed. Once it outgrows the budget (twelve states
// will), the most relevant meetings are picked by the words, districts, people
// and posts in the question, most recent first, and the exact summary still
// covers the rest.
//
// Scope matches the app: State sees everything, Zone its zone, District its
// districts, Field only their own meetings. Emails never reach the model.
var ASK_BUDGET_CHARS = 150000;
var ASK_QUIET_DAYS   = 60;
var ASK_PER_HOUR     = 40;
var ASK_STATUS_RANK  = { Conducted:4, Cancelled:3, Postponed:2, 'Follow-up':1, Planned:1 };

function askTs_(v) {
  if (v instanceof Date) return v.getTime();
  var t = Date.parse((v || '').toString().trim());
  return isNaN(t) ? 0 : t;
}
function askMonth_(ts) { return ts ? Utilities.formatDate(new Date(ts), 'Asia/Kolkata', 'MMM yyyy') : ''; }
function askOneLine_(x) { return (x || '').toString().replace(/\s+/g, ' ').trim(); }

// Every meeting, one record each, the furthest stage winning when an ID
// appears in more than one sheet, the later row winning at equal stage.
function askAllMeetings_() {
  var byId = {}, emailOf = {};
  function keep(m) {
    var old = byId[m.id];
    if (!old || (ASK_STATUS_RANK[m.status] || 0) >= (ASK_STATUS_RANK[old.status] || 0)) byId[m.id] = m;
  }
  var pd = sheetRows_(MEETINGS_SHEET) || [];
  for (var i = 1; i < pd.length; i++) {
    var id = (pd[i][0] || '').toString().trim(); if (!id) continue;
    emailOf[id] = (pd[i][4] || '').toString().trim().toLowerCase();
    var st = (pd[i][13] || 'Planned').toString().trim();
    if (st !== 'Planned' && st !== 'Follow-up') continue;
    keep({ id:id, status:st, ts:askTs_(pd[i][5]), district:(pd[i][1]||'').toString().trim(),
           block:(pd[i][COL_PLAN_SKBLOCK-1]||'').toString().trim(), officer:(pd[i][2]||'').toString().trim(),
           name:(pd[i][9]||'').toString().trim(), post:(pd[i][10]||'').toString().trim(),
           purpose:(pd[i][11]||'').toString().trim(), agenda:askOneLine_(pd[i][12]) });
  }
  var cd = sheetRows_(CONDUCTED_SHEET) || [];
  for (var j = 1; j < cd.length; j++) {
    var cid = (cd[j][0] || '').toString().trim(); if (!cid) continue;
    if (!emailOf[cid]) emailOf[cid] = (cd[j][4] || '').toString().trim().toLowerCase();
    keep({ id:cid, status:'Conducted', ts:askTs_(cd[j][13]) || askTs_(cd[j][5]),
           district:(cd[j][1]||'').toString().trim(), block:(cd[j][COL_CON_SKBLOCK-1]||'').toString().trim(),
           officer:(cd[j][2]||'').toString().trim(), name:(cd[j][9]||'').toString().trim(),
           post:(cd[j][10]||'').toString().trim(), purpose:(cd[j][11]||'').toString().trim(),
           notes:askOneLine_(cd[j][15]), outcome:(cd[j][COL_CON_OUTCOME-1]||'').toString().trim(),
           next:askOneLine_(cd[j][COL_TAG_NEXT-1]), category:(cd[j][COL_TAG_CAT-1]||'').toString().trim(),
           govtMom:askOneLine_(cd[j][COL_TAG_MOMSUM-1]) });
  }
  var xd = sheetRows_(POSTPONED_SHEET) || [];
  for (var k = 1; k < xd.length; k++) {
    var pid = (xd[k][0] || '').toString().trim(); if (!pid) continue;
    keep({ id:pid, status:'Postponed', ts:askTs_(xd[k][7]), district:(xd[k][1]||'').toString().trim(),
           officer:(xd[k][2]||'').toString().trim(), name:(xd[k][4]||'').toString().trim(),
           post:(xd[k][5]||'').toString().trim(), purpose:(xd[k][6]||'').toString().trim(),
           reason:askOneLine_(xd[k][9]) });
  }
  var xc = sheetRows_(CANCELLED_SHEET) || [];
  for (var c = 1; c < xc.length; c++) {
    var kid = (xc[c][0] || '').toString().trim(); if (!kid) continue;
    keep({ id:kid, status:'Cancelled', ts:askTs_(xc[c][5]), district:(xc[c][1]||'').toString().trim(),
           officer:(xc[c][2]||'').toString().trim(), name:(xc[c][9]||'').toString().trim(),
           post:(xc[c][10]||'').toString().trim(), purpose:(xc[c][11]||'').toString().trim(),
           reason:askOneLine_(xc[c][15]) });
  }
  var out = [];
  for (var key in byId) { byId[key].email = emailOf[key] || ''; out.push(byId[key]); }
  return out;
}

// Which meetings this person may ask about, and how to describe that to them.
function askScope_(session, all) {
  var role = (session.role || '').toString(), me = (session.email || '').toLowerCase();
  if (role === 'State') {
    var every = [];
    for (var z in ZONE_DISTRICTS) every = every.concat(ZONE_DISTRICTS[z]);
    return { label:'all districts', rows:all, districts:every.concat(STATE_EXTRA_DISTRICTS) };
  }
  if (role === 'Zone') {
    var zk = findZoneKey_(session.zone || '');
    var set = {}; (ZONE_DISTRICTS[zk] || []).forEach(function(d){ set[normDist_(d)] = 1; });
    return { label:(zk || 'your zone') + ' (' + (ZONE_DISTRICTS[zk] || []).join(', ') + ')',
             rows:all.filter(function(m){ return set[normDist_(m.district)]; }), districts:(ZONE_DISTRICTS[zk] || []).slice() };
  }
  if (role === 'District') {
    var ds = (session.districts && session.districts.length) ? session.districts : [session.district];
    var dset = {}; ds.forEach(function(d){ if (d) dset[normDist_(d)] = 1; });
    return { label:ds.join(', '), rows:all.filter(function(m){ return dset[normDist_(m.district)]; }), districts:ds.filter(Boolean) };
  }
  return { label:'your own meetings', rows:all.filter(function(m){ return m.email && m.email === me; }) };
}

// The exact figures. Everything numeric the model says should come from here.
function askSummary_(rows, now, scopeDistricts) {
  var L = [], st = {}, byMonth = {}, byDist = {}, byOfficer = {}, outc = {}, lastOffice = {}, overdue = 0, oldest = 0;
  rows.forEach(function(m){
    st[m.status] = (st[m.status] || 0) + 1;
    var d = m.district || 'State level';
    var r = byDist[d] || (byDist[d] = { Conducted:0, Planned:0, Postponed:0, Cancelled:0 });
    r[m.status === 'Follow-up' ? 'Planned' : m.status] = (r[m.status === 'Follow-up' ? 'Planned' : m.status] || 0) + 1;
    if (m.status === 'Conducted') {
      var mo = askMonth_(m.ts); if (mo) byMonth[mo] = (byMonth[mo] || 0) + 1;
      if (m.officer) byOfficer[m.officer] = (byOfficer[m.officer] || 0) + 1;
      var o = m.outcome || 'Not recorded'; outc[o] = (outc[o] || 0) + 1;
      var ok = d + '|' + (_prepNorm_(m.post) || _prepNorm_(m.name));
      if (!lastOffice[ok] || m.ts > lastOffice[ok].ts) lastOffice[ok] = { ts:m.ts, district:d, post:m.post || m.name, name:m.name };
    } else if ((m.status === 'Planned' || m.status === 'Follow-up') && m.ts && m.ts < now - 86400000) {
      overdue++; if (!oldest || m.ts < oldest) oldest = m.ts;
    }
  });
  L.push('Meetings in scope: ' + rows.length + ' (' + Object.keys(st).sort().map(function(k){ return k + ' ' + st[k]; }).join(', ') + ').');
  var months = Object.keys(byMonth).sort(function(a, b){ return Date.parse('1 ' + a) - Date.parse('1 ' + b); });
  L.push('Conducted by month: ' + (months.map(function(k){ return k + ' ' + byMonth[k]; }).join(', ') || 'none') + '.');
  L.push('By district (conducted / planned / postponed / cancelled):');
  Object.keys(byDist).sort().forEach(function(k){
    var r = byDist[k]; L.push('  ' + k + ': ' + r.Conducted + ' / ' + r.Planned + ' / ' + r.Postponed + ' / ' + r.Cancelled);
  });
  var offs = Object.keys(byOfficer).sort(function(a, b){ return byOfficer[b] - byOfficer[a]; });
  L.push('Conducted meetings by officer, all time: ' + (offs.map(function(k){ return k + ' ' + byOfficer[k]; }).join('; ') || 'none') + '.');

  // The last four months, oldest first, this month last.
  var recent = [], base = new Date(now);
  for (var b = 3; b >= 0; b--) recent.push(askMonth_(new Date(base.getFullYear(), base.getMonth() - b, 15).getTime()));
  var thisM = recent[3], lastM = recent[2];
  var dm = {}, om = {}, stM = {};
  (scopeDistricts || []).forEach(function(x){ if (x) dm[normDist_(x)] = { name:x, n:{} }; });
  rows.forEach(function(m){
    var mo = askMonth_(m.ts);
    if (mo === thisM) stM[m.status] = (stM[m.status] || 0) + 1;
    if (m.status !== 'Conducted' || recent.indexOf(mo) < 0) return;
    var dk = normDist_(m.district || 'State level');
    var e = dm[dk] || (dm[dk] = { name:m.district || 'State level', n:{} });
    e.n[mo] = (e.n[mo] || 0) + 1;
    if (m.officer && (mo === thisM || mo === lastM)) {
      var oe = om[m.officer] || (om[m.officer] = {}); oe[mo] = (oe[mo] || 0) + 1;
    }
  });
  L.push('THIS MONTH is ' + thisM + '. LAST MONTH is ' + lastM + '.');
  L.push('This month by status: ' + (Object.keys(stM).sort().map(function(k){ return k + ' ' + stM[k]; }).join(', ') || 'nothing yet') + '.');
  L.push('Conducted by district and month (' + recent.join(', ') + '). Every district in scope is listed, 0 means none:');
  Object.keys(dm).sort().forEach(function(k){
    L.push('  ' + dm[k].name + ' by month: ' + recent.map(function(mo){ return mo + ' ' + (dm[k].n[mo] || 0); }).join(', '));
  });
  var ofr = Object.keys(om).sort(function(a, b2){ return ((om[b2][thisM] || 0) - (om[a][thisM] || 0)) || a.localeCompare(b2); });
  L.push('Conducted by officer, this month and last (' + thisM + ' / ' + lastM + '): ' +
         (ofr.map(function(k){ return k + ' ' + (om[k][thisM] || 0) + ' / ' + (om[k][lastM] || 0); }).join('; ') || 'none') + '.');
  L.push('Outcome of conducted meetings: ' + Object.keys(outc).sort().map(function(k){ return k + ' ' + outc[k]; }).join(', ') + '.');
  var quiet = [];
  for (var q in lastOffice) {
    var days = Math.floor((now - lastOffice[q].ts) / 86400000);
    if (lastOffice[q].ts && days > ASK_QUIET_DAYS) quiet.push({ days:days, o:lastOffice[q] });
  }
  quiet.sort(function(a, b){ return b.days - a.days; });
  L.push('Offices not met for over ' + ASK_QUIET_DAYS + ' days: ' + quiet.length + (quiet.length ? '. Longest first: ' +
    quiet.slice(0, 40).map(function(x){ return x.o.post + ', ' + x.o.district + ' (' + x.days + ' days, last ' +
      Utilities.formatDate(new Date(x.o.ts), 'Asia/Kolkata', 'd MMM yyyy') + ')'; }).join('; ') : '') + '.');
  L.push('Planned meetings whose date has passed and were never closed: ' + overdue +
         (oldest ? ' (oldest from ' + Utilities.formatDate(new Date(oldest), 'Asia/Kolkata', 'd MMM yyyy') + ')' : '') + '.');
  return L.join('\n');
}

function askRow_(m) {
  var when = m.ts ? Utilities.formatDate(new Date(m.ts), 'Asia/Kolkata', 'd MMM yyyy') : 'no date';
  var p = ['[' + m.id + '] ' + (m.status === 'Conducted' ? 'Conducted ' + when :
           m.status === 'Postponed' || m.status === 'Cancelled' ? m.status + ' (was ' + when + ')' : m.status + ' for ' + when),
           (m.district || 'State level') + (m.block ? ' / ' + m.block : ''),
           'Officer: ' + (m.officer || '-'),
           'Met: ' + (m.name || '-') + (m.post ? ', ' + m.post : '')];
  if (m.purpose)  p.push('Purpose: ' + m.purpose);
  if (m.outcome)  p.push('Outcome: ' + m.outcome);
  if (m.notes)    p.push('Notes: ' + m.notes);
  if (m.next)     p.push('Next: ' + m.next);
  if (m.govtMom)  p.push('Govt MoM: ' + m.govtMom);
  if (m.agenda)   p.push('Agenda: ' + m.agenda);
  if (m.reason)   p.push('Reason: ' + m.reason);
  return p.join(' | ');
}

function askWords_(s) {
  return (s || '').toString().toLowerCase().split(/[^a-z0-9ऀ-ॿ]+/).filter(function(w){ return w.length >= 3; });
}
// Most relevant first: rarer question words count for more, then recency.
function askRank_(question, lines) {
  var q = {}; askWords_(question).forEach(function(w){ q[w] = 1; });
  var df = {}, toks = lines.map(function(l){
    var seen = {}; askWords_(l.text).forEach(function(w){ if (q[w]) seen[w] = 1; });
    for (var w in seen) df[w] = (df[w] || 0) + 1;
    return seen;
  });
  lines.forEach(function(l, i){
    var sc = 0; for (var w in toks[i]) sc += Math.log(1 + lines.length / df[w]);
    l.score = sc;
  });
  return lines.slice().sort(function(a, b){ return (b.score - a.score) || (b.ts - a.ts); });
}

function askMeetings(session, question, history) {
  question = askOneLine_(question).substring(0, 1000);
  if (question.length < 3) return { success:false, message:'Please type a question.' };

  var me = (session.email || '').toLowerCase(), rk = 'askq_' + me, used = 0;
  try { used = +(CacheService.getScriptCache().get(rk) || 0); } catch (e) {}
  if (used >= ASK_PER_HOUR) return { success:false, message:'That is ' + ASK_PER_HOUR + ' questions in the last hour. Please try again a little later.' };

  var hist = [];
  (Array.isArray(history) ? history : []).slice(-6).forEach(function(h){
    if (h && h.text) hist.push((h.who === 'me' ? 'Q: ' : 'A: ') + askOneLine_(h.text).substring(0, 1500));
  });

  var rows, scope;
  try {
    scope = askScope_(session, askAllMeetings_());
    rows = scope.rows;
  } catch (e) {
    return { success:false, message:'The meeting data is busy right now. Please try again in a minute.' };
  }
  if (!rows.length) return { success:true, answer:'There are no meetings recorded yet in what you can see (' + scope.label + ').', sources:[] };

  var ck = 'aska_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
             me + '|' + scope.label + '|' + question + '|' + hist.join('|'), Utilities.Charset.UTF_8));
  var hit = cGet(ck);
  if (hit) { hit.cached = true; return hit; }

  var now = Date.now();
  var lines = rows.map(function(m){ return { id:m.id, ts:m.ts, text:askRow_(m), m:m }; });
  var total = 0; lines.forEach(function(l){ total += l.text.length + 1; });
  var chosen, note;
  if (total <= ASK_BUDGET_CHARS) {
    chosen = lines.slice().sort(function(a, b){ return b.ts - a.ts; });
    note = 'all ' + lines.length + ' meetings, newest first';
  } else {
    chosen = []; var size = 0, ranked = askRank_(question + ' ' + hist.join(' '), lines);
    for (var i = 0; i < ranked.length && size < ASK_BUDGET_CHARS; i++) { chosen.push(ranked[i]); size += ranked[i].text.length + 1; }
    note = chosen.length + ' of ' + lines.length + ' meetings, the most relevant to the question; the summary covers them all';
  }
  var known = {}; chosen.forEach(function(l){ known[l.id] = l.m; });

  var prompt = [
    'You are the EG-MMS assistant for the Educate Girls Government Relations team in Uttar Pradesh, India.',
    'You answer questions about their meetings with government officials, using ONLY the data below.',
    'Today is ' + Utilities.formatDate(new Date(now), 'Asia/Kolkata', 'd MMM yyyy') + '. The person asking is ' +
      (session.name || 'a team member') + ' (' + (session.role || 'Field') + '), who can see ' + scope.label + '.',
    '',
    'RULES',
    '- Use only the data below. If the answer is not there, say so plainly. Never add a name, date, number or promise that is not in the data.',
    '- For any count or total, use the SUMMARY figures, which are exact. Do not count rows yourself. "This month" and "last month" are defined in the SUMMARY.',
    '- When you rely on particular meetings, cite their IDs in square brackets, for example [MTG-20260812-101010]. Cite only IDs that appear below.',
    '- Reply in the language the question is written in: Hindi, English, or Hinglish. Lead with the direct answer, then detail if it helps. Use short bullet points for lists.',
    '- Meeting notes are written by field officers, in Hindi, English or a mix, and may be brief.',
    '- Do not use em dashes.',
    '',
    'SUMMARY (exact figures)',
    askSummary_(rows, now, scope.districts),
    '',
    'MEETINGS (' + note + ')',
    chosen.map(function(l){ return l.text; }).join('\n'),
    '',
    hist.length ? 'CONVERSATION SO FAR\n' + hist.join('\n') + '\n' : '',
    'QUESTION',
    question
  ].join('\n');

  var raw = callLLM(prompt);
  if (!raw) return { success:false, message:'The assistant did not answer just now. Please try again in a minute.' };

  try { CacheService.getScriptCache().put(rk, String(used + 1), 3600); } catch (e) {}

  var EM = new RegExp(String.fromCharCode(8212), 'g');
  var answer = raw.replace(EM, ',').trim(), sources = [], seen = {};
  // Any meeting ID the model cites that is not in the data it was given is
  // made up. Take it out of the answer rather than show it as a source.
  answer = answer.replace(/\[?(MTG-\d{8}-\d{6})\]?/g, function(all, id){
    if (!known[id]) return '';
    if (!seen[id] && sources.length < 12) {
      seen[id] = 1; var m = known[id];
      sources.push({ id:id, status:m.status,
                     date:m.ts ? Utilities.formatDate(new Date(m.ts), 'Asia/Kolkata', 'd MMM yyyy') : '',
                     district:m.district || 'State level', met:(m.name || '') + (m.post ? ', ' + m.post : '') });
    }
    return '[' + id + ']';
  }).replace(/[ \t]+\n/g, '\n').replace(/ {2,}/g, ' ');

  var res = { success:true, answer:answer, sources:sources, scope:scope.label, seen:chosen.length, total:lines.length };
  try { cPut(ck, res, 600); } catch (e) {}
  return res;
}

function tagOneMeeting_(d) {
  var note = (d.keyPoints||'').toString().trim();
  var out = { priority:'', flag:'', nextAction:'', escalate:false, category:'None', momSummary:'' };
  if (note) {
    var prompt = 'You are tagging a government-relations field meeting note (may be Hindi, English or mixed). Return STRICT JSON only, no markdown: '+
      '{"priority":"High|Medium|Low","flag":"Follow-up needed|Resolved|Blocked","nextAction":"one short action line in English","escalate":true|false,"category":"Document/Data request|Quality issue|Blocker|Resource needed|Commitment|None"}. '+
      'METHOD (read this first): read every sentence of the note to the end and decide from the FULL meaning of the sentence. Never react to a single keyword. Words like "block", "report", "issue", "problem", "pending" carry no meaning on their own. What matters is what the sentence says actually happened to that thing: was it done, asked for, refused, or is it just naming a place. Judge the sentence, not the word.\n'+
      'GLOSSARY (India, important): "block" (also "khand", "vikas khand") is an administrative unit BELOW a district, and BEO, BDO, BRC, CRC, BSA are block level officers or offices. A note saying the meeting was held in a block, at block level, in a named block, or with a block officer is an ordinary LOCATION detail. It NEVER means the work is blocked. '+
      'Use flag "Blocked" or category "Blocker" ONLY when the note actually describes work being stuck, refused, delayed, denied or obstructed. If in doubt, prefer "Follow-up needed".\n'+
      'COMPLETED vs PENDING (important): these notes are usually past tense Hinglish. "report submit ki", "jama kar di", "de di", "kar diya", "ho gaya", "bhej diya", "share kar diya" all mean the work is ALREADY DONE. Report that as a completed update: escalate false, flag "Resolved", category "None" unless something else is open. '+
      'Use escalate true, or a request category, ONLY when something is still OUTSTANDING right now: the officer is still asking for it, it was refused, or it is pending. Never turn a finished action into a fresh request.\n'+
      "Escalate true only for a real ask, request, quality issue, complaint, blocker or problem needing a senior's attention; a positive or normal update is false. Do not use em dashes.\n"+
      'NOTE: '+note+'\n(Purpose: '+(d.purpose||'')+'; Stakeholder: '+(d.stakeholder||'')+'; Type: '+(d.type||'')+')';
    var o = _parseJson_(callLLM(prompt));
    if (o) {
      out.priority=(o.priority||''); out.flag=(o.flag||''); out.nextAction=(o.nextAction||'');
      out.escalate=(o.escalate===true||o.escalate==='true'); out.category=(o.category||'None');
    }
  }
  if (d.govtMom && d.govtMom.toString().trim()) {
    out.momSummary = readGovtMomPdf_(d.govtMom.toString().split(/\s*,\s*/)[0]);
  }
  return out;
}

// Batch: tag conducted meetings that have no tag yet (limit per run to fit the 6-min cap).
function tagUntaggedMeetings(limit) {
  limit = limit || 15;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'No conducted sheet' };
  var data = sh.getDataRange().getValues();
  var done = 0, results = [];
  for (var i = 1; i < data.length && done < limit; i++) {
    if (!data[i][0]) continue;                                  // no meeting id
    if ((data[i][COL_TAG_AT-1]||'').toString().trim()) continue; // already tagged
    var keyPoints = (data[i][15]||'').toString().trim();
    var govtMom   = (data[i][21]||'').toString().trim();
    if (!keyPoints && !govtMom) continue;                       // nothing to read
    if (done) Utilities.sleep(1500);   // stay under the free-tier rate limit
    var t = tagOneMeeting_({ keyPoints:keyPoints, purpose:data[i][11], stakeholder:data[i][10], type:data[i][8], govtMom:govtMom });
    var r = i + 1;
    sh.getRange(r, COL_TAG_PRIORITY).setValue(t.priority);
    sh.getRange(r, COL_TAG_FLAG).setValue(t.flag);
    sh.getRange(r, COL_TAG_NEXT).setValue(t.nextAction);
    sh.getRange(r, COL_TAG_ESC).setValue(t.escalate ? 'Yes' : 'No');
    sh.getRange(r, COL_TAG_CAT).setValue(t.category);
    sh.getRange(r, COL_TAG_MOMSUM).setValue(t.momSummary);
    sh.getRange(r, COL_TAG_AT).setValue(new Date());
    done++;
    results.push(data[i][0] + ': ' + t.priority + '/' + t.flag + (t.escalate ? ' [ESCALATE: ' + t.category + ']' : '') + (t.momSummary ? ' +MoM' : ''));
  }
  try { cDel('reportData'); } catch(e){}
  Logger.log('Tagged ' + done + ' meeting(s).');
  Logger.log(results.join('\n'));
  return { success:true, tagged:done, details:results };
}

// Which meetings are currently marked Blocked/Blocker, and does the note even
// mention a block? The old prompt had no glossary, so an ordinary
// "meeting in X block" note could be read as work being obstructed.
// ------------------------------------------------------------
//  COMMITMENT EXTRACTION (dry run only, nothing is written yet)
//  Pulls out what a government official said they would do. Every
//  commitment must carry the exact sentence it came from, which makes an
//  invented one easy to catch: the quote simply will not be in the note.
// ------------------------------------------------------------
function extractCommitments_(d) {
  var note = (d.keyPoints || '').toString().trim();
  if (note.length < 25) return [];                 // nothing meaningful to read
  var prompt =
    'You are reading a note from a government relations field meeting in India. The note may be Hindi, English or mixed. '+
    'Find every COMMITMENT the government official made, meaning something they said they would do.\n'+
    'METHOD: read each sentence to the end and judge from its full meaning, never from a single keyword.\n'+
    'GLOSSARY: "block" (khand) is an administrative area in India, and BEO, BDO, BRC, CRC, BSA are block level offices or officers. A block is a place. It never means something is obstructed.\n'+
    'ALREADY DONE vs STILL PROMISED: past tense such as "de di", "kar diya", "jaari kar diya", "submit ki", "bhej diya", "ho gaya" means the thing is ALREADY DONE, so it is NOT an open commitment. Only record what the official still has to do.\n'+
    'RULES:\n'+
    '1. Record a commitment ONLY when the note clearly says the official agreed, promised, assured, sanctioned or directed that something will be done. Never infer one from a general discussion or from a request our own team made.\n'+
    '2. Record it whenever the official names a SPECIFIC thing they will provide or do: a document, letter, permission, order, list, data, quantity, visit, instruction to a junior officer, or any named action. Record these readily.\n'+
    '   Leave out only a bare expression of goodwill with no specific thing attached, such as "assured full cooperation" or "call me if there is any problem", because nobody can ever tick those off.\n'+
    '3. For each commitment you MUST copy, word for word, the sentence from the note that shows it, into "evidence". If you cannot copy such a sentence, do not record that commitment.\n'+
    '4. Never invent a deadline. Leave "due" empty unless the note states one.\n'+
    '5. If the note truly contains no such promise, return an empty list.\n'+
    'Return STRICT JSON only, no markdown: {"commitments":[{"what":"one short line in English","by":"who promised, as written in the note","due":"as written, or empty","evidence":"exact sentence copied from the note"}]}\n'+
    'Do not use em dashes.\n'+
    'OFFICIAL: ' + (d.stakeholderName||'') + ' (' + (d.stakeholderPost||'') + '), ' + (d.district||'') + '\n'+
    'PURPOSE: ' + (d.purpose||'') + '\n'+
    'NOTE: ' + note;
  var o = _parseJson_(callLLM(prompt));
  var list = (o && Array.isArray(o.commitments)) ? o.commitments : [];
  var flat = note.toLowerCase().replace(/\s+/g, ' ');
  return list.map(function(c){
    var ev = (c.evidence || '').toString().trim();
    // Does the quote actually exist in the note? A hallucination fails this.
    var probe = ev.toLowerCase().replace(/\s+/g, ' ').substring(0, 45);
    return { what:(c.what||'').toString().trim(), by:(c.by||'').toString().trim(),
             due:(c.due||'').toString().trim(), evidence:ev,
             evidenceFound: (probe.length >= 12 && flat.indexOf(probe) >= 0) };
  }).filter(function(c){ return c.what; });
}

// Try the extractor on real meetings and PRINT the result. Writes nothing,
// changes nothing, so it is always safe to run.
function COMMIT_dryRun(limit) {
  limit = limit || 15;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'no conducted sheet' };
  var data = sh.getDataRange().getValues();

  // Take the meetings with the most substantial notes: that is where the
  // extractor has something to work with and where a mistake would matter.
  var cand = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var note = (data[i][15] || '').toString().trim();
    if (note.length < 25) continue;
    cand.push({ row:i+1, id:data[i][0], district:data[i][1], officer:data[i][2],
                stakeholderName:data[i][9], stakeholderPost:data[i][10],
                purpose:data[i][11], date:fmtDateVal(data[i][13]), keyPoints:note });
  }
  cand.sort(function(a,b){ return b.keyPoints.length - a.keyPoints.length; });
  var pick = cand.slice(0, limit);

  var out = [], totals = { meetings:0, found:0, evidenceOk:0, evidenceBad:0, none:0 };
  pick.forEach(function(m, idx){
    if (idx) Utilities.sleep(1500);                // stay under the rate limit
    var cs = extractCommitments_(m);
    totals.meetings++;
    if (!cs.length) totals.none++;
    cs.forEach(function(c){
      totals.found++;
      if (c.evidenceFound) totals.evidenceOk++; else totals.evidenceBad++;
    });
    out.push({ date:m.date, district:m.district, officer:m.officer,
               official:m.stakeholderName + ' (' + m.stakeholderPost + ')',
               noteLength:m.keyPoints.length,
               note:m.keyPoints.substring(0, 220),
               commitments: cs.length ? cs : '(none found)' });
  });

  var res = { summary:totals, results:out };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// ------------------------------------------------------------
//  NOTE QUALITY
//  Counts only, never the text itself, so this can be read and shared
//  without exposing what was written in any meeting.
// ------------------------------------------------------------
// Find rows in Plan Meetings that render as a blank line in Manage Meetings:
// no id, or an id with the core fields missing. Reads only, writes nothing.
// Editor helper: rows that look like one plan saved more than once (same officer,
// official, date and purpose). Reports only, it deletes nothing; remove the extras
// from Manage Meetings so the app keeps its own record straight.
// Editor helper: meeting ids written more than once into the Conducted sheet.
// These are the old duplicate-conduct rows, from before conductMeeting refused a
// second record. Reports only. Delete the row holding the weaker note and keep
// the real one.
// Editor helper: every row any sheet holds for one meeting, with its row number.
// Before deleting anything, this is how you confirm you are looking at the right
// sheet and the right row. Row numbers are per sheet: row 127 of Conducted
// Meetings is a different meeting from row 127 of Plan Meetings.
//   MTG_trace('MTG-20260807-195551')
function MTG_trace(meetingId) {
  var want = (meetingId || '').toString().trim().toUpperCase();
  if (!want) return 'Pass a meeting id, e.g. MTG_trace("MTG-20260807-195551")';
  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = ['Tracing ' + want, ''];
  [MEETINGS_SHEET, CONDUCTED_SHEET, POSTPONED_SHEET, CANCELLED_SHEET].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { log.push(name + '  -  sheet not found'); log.push(''); return; }
    var d = sh.getDataRange().getValues(), hits = 0;
    for (var i = 1; i < d.length; i++) {
      if ((d[i][0] || '').toString().trim().toUpperCase() !== want) continue;
      hits++;
      log.push(name + '   ROW ' + (i + 1));
      log.push('     district: ' + (d[i][1] || '') + '    officer: ' + (d[i][2] || ''));
      log.push('     official: ' + (d[i][9] || '') + '    purpose: ' + (d[i][11] || ''));
      if (name === MEETINGS_SHEET) {
        log.push('     status:   ' + (d[i][13] || '(blank)') + '    planned for: ' + (d[i][5] || ''));
      }
      if (name === CONDUCTED_SHEET) {
        log.push('     conducted on: ' + (d[i][13] || '') + '    filed at: ' + (d[i][20] || ''));
        log.push('     note: ' + (d[i][15] || '').toString().replace(/\s+/g, ' ').slice(0, 120));
      }
      log.push('');
    }
    if (!hits) { log.push(name + '  -  not present'); log.push(''); }
  });
  Logger.log(log.join(String.fromCharCode(10)));
  return 'done';
}

// Editor helper: clears the exact copies out of the Conducted sheet.
//
// CONDUCT_cleanDupes()          shows what it would do, changes nothing
// CONDUCT_cleanDupes('DELETE')  actually deletes
//
// It only removes a row whose note is character for character the same as
// another row for the same meeting, so nothing anyone actually wrote is lost.
// Where two rows for one meeting hold genuinely different notes it deletes
// neither and says so: choosing which account of a meeting is the real one is
// not a decision code should make. Deletes from the bottom up, so the row
// numbers above each deletion stay valid while it works.
function CONDUCT_cleanDupes(confirm) {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONDUCTED_SHEET);
  if (!sh) return 'No conducted sheet';
  var rows = sh.getDataRange().getValues();
  var groups = {};
  for (var i = 1; i < rows.length; i++) {
    var id = (rows[i][0] || '').toString().trim();
    if (!id) continue;
    if (!groups[id]) groups[id] = [];
    groups[id].push({ row: i + 1, note: (rows[i][15] || '').toString().trim().replace(/\s+/g, ' ') });
  }

  var copies = [], log = [], yours = 0;
  for (var id2 in groups) {
    var g = groups[id2];
    if (g.length < 2) continue;
    var kept = {}, distinct = [];
    for (var j = 0; j < g.length; j++) {
      if (kept[g[j].note]) { copies.push(g[j]); continue; }
      kept[g[j].note] = 1;
      distinct.push(g[j]);
    }
    if (distinct.length > 1) {
      yours++;
      log.push(id2 + '  -  ' + distinct.length + ' different notes, this one is your call:');
      for (var k = 0; k < distinct.length; k++) {
        log.push('     row ' + distinct[k].row + '   ' + distinct[k].note.slice(0, 100));
      }
      log.push('');
    }
  }

  copies.sort(function(a, b) { return a.row - b.row; });
  log.push('Exact copies, safe to remove: ' + copies.length);
  for (var c = 0; c < copies.length; c++) {
    log.push('     row ' + copies[c].row + '   ' + copies[c].note.slice(0, 70));
  }
  log.push('');

  if (confirm === 'DELETE') {
    for (var dI = copies.length - 1; dI >= 0; dI--) sh.deleteRow(copies[dI].row);
    cDel('reportData', 'stateMtg_all');
    log.push('DELETED ' + copies.length + ' rows. Caches cleared.');
    log.push(yours + ' meeting(s) still need you to pick which note to keep.');
  } else {
    log.push('Preview only, nothing was deleted.');
    log.push('Run CONDUCT_cleanDupes("DELETE") to remove the ' + copies.length + ' exact copies.');
  }
  Logger.log(log.join(String.fromCharCode(10)));
  return copies.length;
}

function CONDUCT_findDupes() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONDUCTED_SHEET);
  if (!sh) return 0;
  var rows = sh.getDataRange().getValues();
  var seen = {}, out = [], extra = 0;
  for (var i = 1; i < rows.length; i++) {
    var id = (rows[i][0] || '').toString().trim();
    if (!id) continue;
    if (!seen[id]) { seen[id] = []; }
    seen[id].push({ row: i + 1, note: (rows[i][15] || '').toString().slice(0, 70) });
  }
  for (var k in seen) {
    if (seen[k].length < 2) continue;
    extra += seen[k].length - 1;
    out.push(k + '  (' + seen[k].length + ' rows)');
    seen[k].forEach(function(x) { out.push('     row ' + x.row + '   ' + x.note); });
  }
  Logger.log(out.length ? out.join('\n') : 'No duplicate conducted rows found.');
  Logger.log('Extra rows beyond one per meeting: ' + extra);
  return extra;
}

function PLAN_findDupes() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MEETINGS_SHEET);
  var rows  = sheet.getDataRange().getValues();
  var groups = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!(r[0] || '').toString().trim()) continue;
    var k = [(r[4] || '').toString().trim().toLowerCase(),
             fmtDateVal(r[5]),
             (r[9] || '').toString().trim().toLowerCase(),
             (r[11] || '').toString().trim().toLowerCase()].join('|');
    if (!groups[k]) groups[k] = [];
    groups[k].push({ row: i + 1, id: (r[0] || '').toString().trim(), status: (r[13] || '').toString().trim() });
  }
  var out = [], total = 0;
  for (var g in groups) {
    if (groups[g].length < 2) continue;
    total += groups[g].length - 1;
    var p = g.split('|');
    out.push(p[2] + '  |  ' + p[1] + '  |  ' + p[0] + '  |  ' + p[3]);
    groups[g].forEach(function(x) { out.push('     row ' + x.row + '   ' + x.id + '   ' + x.status); });
  }
  Logger.log(out.length ? out.join('\n') : 'No duplicate plans found.');
  Logger.log('Extra rows beyond one per meeting: ' + total);
  return total;
}

function PLAN_findBlank() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(MEETINGS_SHEET);
  if (!sh) return { success:false, message:'no plan sheet' };
  var data = sh.getDataRange().getValues();
  var bad = [];
  for (var i = 1; i < data.length; i++) {
    var id = (data[i][0] || '').toString().trim();
    var dist = (data[i][1] || '').toString().trim();
    var email = (data[i][4] || '').toString().trim();
    var date = (data[i][5] || '').toString().trim();
    var name = (data[i][9] || '').toString().trim();
    var status = (data[i][13] || '').toString().trim();
    var filled = data[i].filter(function(c){ return (c || '').toString().trim(); }).length;
    if (!id || !dist || !date || !name) {
      bad.push({ sheetRow:i + 1, id:id || '(none)', district:dist || '(none)',
                 email:email || '(none)', date:date || '(none)', stakeholder:name || '(none)',
                 status:status || '(none)', nonEmptyCells:filled });
    }
  }
  var res = { totalRows: data.length - 1, columnsInSheet: data[0].length, suspectRows: bad.length, rows: bad };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// Remove the half rows PLAN_findBlank lists: a meeting id and the session
// fields, but no date and no stakeholder, so nothing was ever really planned.
// Deletes bottom up so the row numbers stay valid while it works. Run
// PLAN_findBlank first and read what it found.
function PLAN_deleteBlank() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(MEETINGS_SHEET);
  if (!sh) return { success:false, message:'no plan sheet' };
  var data = sh.getDataRange().getValues();
  var kill = [];
  for (var i = 1; i < data.length; i++) {
    if (!(data[i][0] || '').toString().trim()) continue;              // keep truly empty lines alone
    var date = (data[i][5] || '').toString().trim();
    var name = (data[i][9] || '').toString().trim();
    var status = (data[i][13] || '').toString().trim();
    // Only ever the never-started ones: no date, no stakeholder, still Planned.
    if (!date && !name && status === 'Planned') kill.push({ row:i + 1, id:data[i][0] });
  }
  for (var k = kill.length - 1; k >= 0; k--) sh.deleteRow(kill[k].row);
  try { cDel('reportData'); } catch(e){}
  Logger.log('Deleted ' + kill.length + ' half rows: ' + kill.map(function(x){ return x.id; }).join(', '));
  return { deleted: kill.length, rows: kill };
}

function NOTE_quality() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'no conducted sheet' };
  var data = sh.getDataRange().getValues();

  function bucket(n) {
    return n === 0 ? '0 empty' : n < 30 ? '1 under 30' : n < 80 ? '2 30 to 79' :
           n < 200 ? '3 80 to 199' : n < 500 ? '4 200 to 499' : '5 500 plus';
  }
  var kp = {}, ag = {}, byOfficer = {}, seen = {}, dupes = 0, dupOfficers = {};
  var total = 0, kpSum = 0;

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    total++;
    var officer = (data[i][2] || '?').toString().trim();
    var note = (data[i][15] || '').toString().trim();
    var agenda = (data[i][12] || '').toString().trim();
    kp[bucket(note.length)] = (kp[bucket(note.length)] || 0) + 1;
    ag[bucket(agenda.length)] = (ag[bucket(agenda.length)] || 0) + 1;
    kpSum += note.length;

    var o = byOfficer[officer] || (byOfficer[officer] = { meetings:0, totalChars:0, thin:0 });
    o.meetings++; o.totalChars += note.length;
    if (note.length < 80) o.thin++;

    // the same note text reused across meetings, which a length rule cannot catch
    if (note.length >= 40) {
      var k = note.toLowerCase().replace(/\s+/g, ' ').substring(0, 160);
      if (seen[k]) { dupes++; dupOfficers[officer] = (dupOfficers[officer] || 0) + 1; }
      else seen[k] = 1;
    }
  }

  var officers = Object.keys(byOfficer).map(function(n){
    var o = byOfficer[n];
    return { officer:n, meetings:o.meetings, avgChars:Math.round(o.totalChars / o.meetings), thinNotes:o.thin };
  }).sort(function(a,b){ return a.avgChars - b.avgChars; });

  var res = {
    conductedMeetings: total,
    avgKeyPointsChars: total ? Math.round(kpSum / total) : 0,
    keyPointsLengths: kp,
    agendaLengths: ag,
    repeatedNotes: dupes,
    repeatedNotesByOfficer: dupOfficers,
    shortestWriters: officers.slice(0, 10),
    longestWriters: officers.slice(-5).reverse()
  };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// Would a keyword rule ("urgent", "immediately", "denied" -> Priority High) do
// the job the model is doing? Read only, writes nothing. It scores the keyword
// rule against the tags already on the sheet and, more usefully, prints the
// notes where the keyword fires but the model disagreed, because a false
// positive is something you have to read to judge. Counts alone cannot settle
// it. See tagOneMeeting_, whose prompt says in as many words: never react to a
// single keyword.
function NOTE_keywordCheck() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'no conducted sheet' };
  var data = sh.getDataRange().getValues();

  // Latin terms are matched on word boundaries so "blocked" cannot be found
  // inside "block", which in these notes is a place, not a problem.
  var WORD = ['urgent','urgently','immediately','immediate','denied','deny',
              'refused','refuse','rejected','reject','delayed','pending',
              'blocked','objection','complaint','turant','jaldi','inkar',
              'mana','lambit','dikkat','samasya'];
  // Devanagari has no useful word boundary in this regex engine, so substring.
  var SUB  = ['तत्काल','शीघ्र','अविलम्ब','मना','इनकार','अस्वीकार',
              'लंबित','विलंब','समस्या','आपत्ति','शिकायत'];

  // Whole words only, by splitting on anything that is not a letter or a
  // digit: "blocked" must not be found inside "block", which in these notes
  // is a place and not a problem. Devanagari is matched as a substring
  // because it does not sit inside that Latin word split.
  function hits(note) {
    var low = ' ' + note.toLowerCase().replace(/[^a-z0-9]+/g, ' ') + ' ';
    var found = [];
    for (var w = 0; w < WORD.length; w++) {
      if (low.indexOf(' ' + WORD[w] + ' ') !== -1) found.push(WORD[w]);
    }
    for (var d = 0; d < SUB.length; d++) {
      if (note.indexOf(SUB[d]) !== -1) found.push(SUB[d]);
    }
    return found;
  }
  function around(note, term) {
    var i = note.toLowerCase().indexOf(term.toLowerCase());
    if (i < 0) i = note.indexOf(term);
    if (i < 0) return note.substring(0, 110);
    return (i > 40 ? '...' : '') + note.substring(Math.max(0, i - 40), i + 70) + '...';
  }

  var perWord = {}, tagged = 0, kwAny = 0, aiHigh = 0;
  var kwAndHigh = 0, kwNotHigh = 0, highNoKw = 0;
  var wouldPromote = [], wouldMiss = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var note = (data[i][15] || '').toString().trim();
    if (!note) continue;
    var prio = (data[i][COL_TAG_PRIORITY-1] || '').toString().trim();
    if (!prio) continue;                      // untagged, nothing to compare against
    tagged++;
    var found = hits(note), isHigh = (prio === 'High');
    if (isHigh) aiHigh++;

    for (var f = 0; f < found.length; f++) {
      var w = perWord[found[f]] || (perWord[found[f]] = { notes:0, High:0, Medium:0, Low:0 });
      w.notes++; w[prio] = (w[prio] || 0) + 1;
    }
    if (found.length) {
      kwAny++;
      if (isHigh) kwAndHigh++;
      else {
        kwNotHigh++;
        // The keyword rule would raise this to High. Read it and decide.
        if (wouldPromote.length < 25) {
          wouldPromote.push({ id:data[i][0], aiSaid:prio, word:found[0], text:around(note, found[0]) });
        }
      }
    } else if (isHigh) {
      highNoKw++;
      if (wouldMiss.length < 12) {
        wouldMiss.push({ id:data[i][0], text:note.substring(0, 110) + (note.length > 110 ? '...' : '') });
      }
    }
  }

  var res = {
    taggedNotes: tagged,
    aiSaysHigh: aiHigh,
    keywordWouldFire: kwAny,
    agreeBothHigh: kwAndHigh,
    keywordFiresButAiSaysMediumOrLow: kwNotHigh,
    aiSaysHighButNoKeyword: highNoKw,
    ofKeywordHitsPercentActuallyHigh: kwAny ? Math.round(100 * kwAndHigh / kwAny) : 0,
    ofHighPercentTheKeywordWouldCatch: aiHigh ? Math.round(100 * kwAndHigh / aiHigh) : 0,
    perWord: perWord,
    readThese_keywordFiredAnywayLowOrMedium: wouldPromote,
    readThese_highWithNoKeywordAtAll: wouldMiss
  };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function TAG_showBlocked() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'No conducted sheet' };
  var data = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var flag = (data[i][COL_TAG_FLAG-1]||'').toString();
    var cat  = (data[i][COL_TAG_CAT-1]||'').toString();
    if (flag !== 'Blocked' && cat !== 'Blocker') continue;
    var note = (data[i][15]||'').toString();
    out.push({ row:i+1, id:data[i][0], flag:flag, category:cat,
               escalate:(data[i][COL_TAG_ESC-1]||'').toString(),
               noteMentionsBlock:/block|khand/i.test(note) ? 'YES' : 'no',
               note:note.substring(0,120) });
  }
  Logger.log(JSON.stringify(out, null, 2));
  return { count:out.length, rows:out };
}

// Clear the tag columns on Blocked/Blocker rows so the hourly job re-tags them
// with the corrected prompt. Does NOT touch the escalation-sent flag, so this
// cannot trigger a fresh wave of escalation mails on its own.
function TAG_recheckBlocked() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'No conducted sheet' };
  var data = sh.getDataRange().getValues(), n = 0, ids = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var flag = (data[i][COL_TAG_FLAG-1]||'').toString();
    var cat  = (data[i][COL_TAG_CAT-1]||'').toString();
    if (flag !== 'Blocked' && cat !== 'Blocker') continue;
    sh.getRange(i+1, COL_TAG_AT).setValue('');                  // marks it untagged again
    n++; ids.push(data[i][0]);
  }
  try { cDel('reportData'); } catch(e){}
  Logger.log('Cleared ' + n + ' row(s) for re-tagging: ' + ids.join(', '));
  return { success:true, cleared:n, ids:ids,
           note:'Run TAG_run (or wait for the hourly job) to re-tag these with the corrected prompt.' };
}

// Ensure the Conducted sheet has the tag column headers (run once; also safe to re-run).
function ensureTagHeaders() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONDUCTED_SHEET);
  if (!sh) return 'no sheet';
  var hdr = ['Priority','Flag','Next Action','Escalate','Category','Govt MoM Summary','Tagged At','Escalation Sent At'];
  sh.getRange(1, COL_TAG_PRIORITY, 1, hdr.length).setValues([hdr]);
  return 'headers set';
}

function taggingJob() { return tagUntaggedMeetings(12); }   // small batch: sleeps + retries must fit in 6 min
function installTaggingTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='taggingJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('taggingJob').timeBased().everyHours(1).create();
  return 'Tagging trigger installed: taggingJob runs hourly and tags newly conducted meetings.';
}

// ---- Run from the editor ----
function TAG_run()         { ensureTagHeaders(); return tagUntaggedMeetings(10); }   // manual test (tags up to 10)
function TAG_installAuto() { ensureTagHeaders(); return installTaggingTrigger(); }   // hourly auto-tagging

// ============================================================
//  TIER 2 - ESCALATION EMAILS (senior CC by hierarchy)
//  When a conducted meeting is Escalate=Yes / High / Blocked, email the
//  officer and CC their senior (Field->District lead, District->Zone lead,
//  Zone->State lead). Sent once per meeting (tracked in col AD).
// ============================================================
var COL_ESC_SENT = 30;   // AD = Escalation Sent At

// Copied on every escalation, whoever filed the meeting, so the state sees
// what is being asked for across all districts. One named person rather than
// everyone holding the State role: there are several, and most of them do not
// need ten of these a month. Blank this out to switch it off.
var ESC_CC_STATE = 'nitinkumar.jha@educategirls.ngo';

// Senior email(s) one level up, from the officer's role + geography.
function findSenior_(emp, recips) {
  var role = (emp && emp.role) || '', me = ((emp && emp.email)||'').toLowerCase(), out = [];
  function push(r){ if (r.email && r.email.toLowerCase() !== me && out.indexOf(r.email) === -1) out.push(r.email); }
  if (role === 'Field') {
    var dists = (emp.districts && emp.districts.length) ? emp.districts : [emp.district];
    recips.forEach(function(r){ if (r.role==='District' && dists.some(function(d){ return normDist_(d)===normDist_(r.district); })) push(r); });
    if (!out.length) { var zk=districtToZone_(emp.district); recips.forEach(function(r){ if(r.role==='Zone'&&findZoneKey_(r.zone)===zk) push(r); }); }
  } else if (role === 'District') {
    var zk2 = districtToZone_(emp.district);
    recips.forEach(function(r){ if (r.role==='Zone' && findZoneKey_(r.zone)===zk2) push(r); });
    if (!out.length) recips.forEach(function(r){ if (r.role==='State') push(r); });
  } else if (role === 'Zone') {
    recips.forEach(function(r){ if (r.role==='State') push(r); });
  }
  return out;
}

function buildEscalationEmail_(o) {
  var pc = o.priority==='High' ? '#B91C1C' : '#9a5b0e';
  return '<div style="margin:0;padding:20px 12px;background:#f4f2ef;font-family:Arial,Helvetica,sans-serif;color:#1f2937;">'+
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:620px;margin:0 auto;background:#fff;border:1px solid #e5e7eb;border-radius:12px;">'+
    '<tr><td style="padding:22px 28px 14px;border-bottom:2px solid '+pc+';">'+
      '<div style="font-size:11px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:'+pc+';">Escalation - Needs attention</div>'+
      '<h1 style="font-family:Georgia,serif;font-size:20px;margin:8px 0 3px;color:#1f2937;">'+_emailEsc(o.category||'Follow-up needed')+'</h1>'+
      '<div style="font-size:13px;color:#6b7280;">'+_emailEsc(o.district)+' &middot; '+_emailEsc(o.conductDate)+'</div></td></tr>'+
    '<tr><td style="padding:16px 28px 0;font-size:14px;line-height:1.6;">'+
      'Dear <b>'+_emailEsc(o.officerName)+'</b>, one of your meetings needs attention.'+
      '<div style="background:#f7f2ee;border:1px solid #e5e7eb;border-radius:8px;padding:12px 16px;margin-top:12px;font-size:13.5px;">'+
        '<b>Meeting:</b> '+_emailEsc(o.stakeholder)+(o.purpose?' &middot; '+_emailEsc(o.purpose):'')+'<br>'+
        '<b>Priority:</b> <span style="color:'+pc+';font-weight:700;">'+_emailEsc(o.priority||'-')+'</span> &nbsp; <b>Status:</b> '+_emailEsc(o.flag||'-')+'<br>'+
        '<b>Next action:</b> '+_emailEsc(o.nextAction||'-')+
      '</div>'+
      (o.keyPoints?'<div style="font-size:12.5px;color:#6b7280;margin-top:10px;"><b style="color:#1f2937;">Note:</b> '+_emailEsc(o.keyPoints.slice(0,300))+'</div>':'')+
    '</td></tr>'+
    '<tr><td style="padding:18px 28px 24px;"><div style="border-top:1px solid #e5e7eb;padding-top:12px;font-size:11px;color:#9ca3af;">Auto-flagged by EG-MMS from the meeting note. dataimpact.in</div></td></tr>'+
    '</table></div>';
}

// ------------------------------------------------------------
//  MEETING PREP BRIEF
//  On demand only (officer taps "Prep"), so nothing here runs on page load
//  and no existing path is touched. Result is cached for 30 min per meeting.
// ------------------------------------------------------------

// Officers type the stakeholder freehand, so compare loosely.
function _prepNorm_(s) {
  return (s||'').toString().toLowerCase()
    .replace(/\b(shri|smt|mr|mrs|ms|dr|sir|madam|maam|ma'am|ji)\b/g, '')
    .replace(/[^a-z0-9ऀ-ॿ]/g, '')
    .trim();
}

function getMeetingPrep(session, meetingId) {
  meetingId = (meetingId||'').toString().trim();
  if (!meetingId) return { success:false, message:'No meeting id' };

  var cacheKey = 'prep_' + meetingId;
  var hit = cGet(cacheKey);
  if (hit) { hit.cached = true; return hit; }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var plan = ss.getSheetByName(MEETINGS_SHEET);
  if (!plan) return { success:false, message:'No plan sheet' };

  // 1. Find the planned meeting, and make sure it belongs to this officer.
  var pd = plan.getDataRange().getValues(), me = null;
  for (var i = 1; i < pd.length; i++) {
    if ((pd[i][0]||'').toString().trim() !== meetingId) continue;
    me = { district:(pd[i][1]||'').toString(), officer:(pd[i][2]||'').toString(),
           email:(pd[i][4]||'').toString(), date:fmtDateVal(pd[i][5]),
           time:(pd[i][6]||'').toString(), type:(pd[i][8]||'').toString(),
           name:(pd[i][9]||'').toString(), post:(pd[i][10]||'').toString(),
           purpose:(pd[i][11]||'').toString(), agenda:(pd[i][12]||'').toString() };
    break;
  }
  if (!me) return { success:false, message:'Meeting not found' };
  if (me.email.toLowerCase() !== (session.email||'').toLowerCase()) {
    return { success:false, message:'This meeting belongs to another officer' };
  }

  // 2. Past conducted meetings in the same district, matched by person or by post.
  var con = ss.getSheetByName(CONDUCTED_SHEET);
  var cd = con ? con.getDataRange().getValues() : [];
  var wantName = _prepNorm_(me.name), wantPost = _prepNorm_(me.post), wantDist = normDist_(me.district);
  var byPerson = [], byPost = [], myNotes = [];
  for (var j = 1; j < cd.length; j++) {
    if (!cd[j][0]) continue;
    var note = (cd[j][15]||'').toString().trim();
    // language sample: this officer's own writing, from any meeting
    if (note && (cd[j][4]||'').toString().toLowerCase() === me.email.toLowerCase()) myNotes.push(note);
    if (normDist_((cd[j][1]||'').toString()) !== wantDist) continue;
    var nm = _prepNorm_(cd[j][9]), pt = _prepNorm_(cd[j][10]);
    var rec = { date:fmtDateVal(cd[j][13]), by:(cd[j][2]||'').toString(),
                stakeholder:(cd[j][9]||'').toString(), post:(cd[j][10]||'').toString(),
                purpose:(cd[j][11]||'').toString(), note:note,
                flag:(cd[j][COL_TAG_FLAG-1]||'').toString(),
                nextAction:(cd[j][COL_TAG_NEXT-1]||'').toString(),
                category:(cd[j][COL_TAG_CAT-1]||'').toString() };
    if (wantName && nm === wantName)      byPerson.push(rec);
    else if (wantPost && pt === wantPost) byPost.push(rec);
  }
  byPerson.sort(function(a,b){ return monthSortVal_(b.date) - monthSortVal_(a.date); });
  byPost.sort(function(a,b){ return monthSortVal_(b.date) - monthSortVal_(a.date); });

  var hist = byPerson.concat(byPost).slice(0, 8);    // keep the prompt small
  var res = {
    success: true,
    meeting: { id:meetingId, stakeholder:me.name, post:me.post, district:me.district,
               date:me.date, time:me.time, purpose:me.purpose },
    counts: { withPerson:byPerson.length, withPost:byPost.length },
    history: hist.map(function(h){ return { date:h.date, by:h.by, purpose:h.purpose, flag:h.flag }; }),
    brief: ''
  };

  if (!hist.length) {
    res.brief = '';
    res.firstMeeting = true;
    cPut(cacheKey, res, C_TTL_EMP);
    return res;
  }

  // 3. Ask the model for the brief, written in the officer's own language.
  var sample = myNotes.slice(-4).join(' | ').substring(0, 500);
  var lines = hist.map(function(h){
    return '- ' + h.date + ' (' + h.purpose + '; status: ' + (h.flag||'-') +
           (h.nextAction ? '; next: ' + h.nextAction : '') + '): ' + h.note.substring(0, 300);
  }).join('\n');

  var prompt =
    'You are briefing a government-relations field officer who is about to meet a government official again. '+
    'Write a SHORT prep brief from the past meeting records below.\n'+
    'LANGUAGE (important): write the brief in the SAME language and script the officer himself uses in his own notes. '+
    'If his notes are Roman-script Hinglish, reply in Roman-script Hinglish. If Devanagari Hindi, reply in Devanagari. '+
    'If English, reply in English. Match his style, do not translate him into another language. '+
    'Here is a sample of his own writing: "' + sample + '"\n'+
    'METHOD: read each past record fully and judge from the whole sentence, never from a single keyword. '+
    'Note that "block" (khand) is an administrative area in India, not an obstruction. '+
    'Past tense like "kar diya", "submit ki", "ho gaya" means that thing is ALREADY DONE, so put it under done, not under pending.\n'+
    'Return STRICT JSON only, no markdown: '+
    '{"summary":"1 or 2 lines on the relationship so far","done":["things already achieved, short lines"],'+
    '"pending":["things still open or promised but not delivered, short lines"],'+
    '"talkingPoints":["2 to 4 things to raise in this meeting, short lines"]}. '+
    'Use [] for any list with nothing to report. At most 3 items per list, each under 15 words. Do not invent anything that is not in the records. Do not use em dashes.\n'+
    'UPCOMING MEETING: ' + me.name + ' (' + me.post + '), ' + me.district + ', purpose: ' + me.purpose + '\n'+
    'PAST RECORDS (newest first):\n' + lines + '\n\n'+
    'Reply with the JSON object only. No explanation before or after it.';

  var rawLLM = callLLM(prompt);
  var o = _parseJson_(rawLLM);
  if (o) {
    res.brief = {
      summary: (o.summary||'').toString(),
      done: Array.isArray(o.done) ? o.done : [],
      pending: Array.isArray(o.pending) ? o.pending : [],
      talkingPoints: Array.isArray(o.talkingPoints) ? o.talkingPoints : []
    };
  } else {
    res.brief = '';
    res.aiFailed = true;    // the raw history is still returned, so the popup is never empty
    res.rawLen = (rawLLM||'').length;              // 0 means the API call itself failed
    res.rawSample = (rawLLM||'').substring(0, 400); // what the model actually replied
  }
  // Only cache a brief that actually came out. Caching a failure would keep
  // serving the same empty result for half an hour after the model recovers.
  if (!res.aiFailed) cPut(cacheKey, res, C_TTL_EMP);
  return res;
}

// Run from the editor to see exactly why a prep brief failed: the model's raw
// reply, or length 0 if the API call itself did not come back.
function PREP_debug() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var plan = ss.getSheetByName(MEETINGS_SHEET);
  var pd = plan.getDataRange().getValues(), id = '', email = '';
  for (var i = pd.length - 1; i >= 1; i--) {                    // newest first
    var st = (pd[i][13]||'').toString();
    if (st !== 'Planned' && st !== 'Follow-up') continue;
    if ((pd[i][4]||'').toString().toLowerCase() !== REPORT_TEST_EMAIL.toLowerCase()) continue;
    id = (pd[i][0]||'').toString(); email = (pd[i][4]||'').toString(); break;
  }
  if (!id) return { message:'No planned meeting found for ' + REPORT_TEST_EMAIL };
  cDel('prep_' + id);                                          // force a fresh model call
  var r = getMeetingPrep({ email:email }, id);
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}

// Dry run: who WOULD get an escalation right now, and why every other row is
// skipped. Sends nothing and writes nothing, so it is always safe to run.
function ESC_preview() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'no sheet' };
  var data = sh.getDataRange().getValues();
  var skip = { notTagged:0, alreadySent:0, resolved:0, nothingToEscalate:0 };
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (!(data[i][COL_TAG_AT-1]||'').toString().trim())   { skip.notTagged++;   continue; }
    if ((data[i][COL_ESC_SENT-1]||'').toString().trim())  { skip.alreadySent++; continue; }
    var prio = (data[i][COL_TAG_PRIORITY-1]||'').toString();
    var flag = (data[i][COL_TAG_FLAG-1]||'').toString();
    var escY = (data[i][COL_TAG_ESC-1]||'').toString();
    if (flag === 'Resolved') { skip.resolved++; continue; }
    if (!(prio === 'High' || flag === 'Blocked' || escY === 'Yes')) { skip.nothingToEscalate++; continue; }
    list.push({ row:i+1, id:data[i][0], officer:(data[i][2]||'').toString(),
                district:(data[i][1]||'').toString(), priority:prio, flag:flag, escalate:escY,
                nextAction:(data[i][COL_TAG_NEXT-1]||'').toString(),
                note:(data[i][15]||'').toString().substring(0,140) });
  }
  var res = { wouldSendNow:list.length, skipped:skip, rows:list };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// mode 'test' sends all to REPORT_TEST_EMAIL; 'live' emails the officer + CC senior.
function sendEscalations(mode, limit) {
  mode = mode || 'test'; limit = limit || 25;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(CONDUCTED_SHEET);
  if (!sh) return { success:false, message:'no sheet' };
  var data = sh.getDataRange().getValues();
  var recips = getReportRecipients();
  var done = 0, out = [];
  for (var i = 1; i < data.length && done < limit; i++) {
    if (!data[i][0]) continue;
    if (!(data[i][COL_TAG_AT-1]||'').toString().trim()) continue;                 // not tagged
    if ((data[i][COL_ESC_SENT-1]||'').toString().trim()) continue;                // already escalated
    var prio = (data[i][COL_TAG_PRIORITY-1]||'').toString();
    var flag = (data[i][COL_TAG_FLAG-1]||'').toString();
    var escY = (data[i][COL_TAG_ESC-1]||'').toString();
    // Resolved means the work is already done. A finished item never needs a
    // senior's attention, however high its priority was rated.
    if (flag === 'Resolved') continue;
    var esc = prio === 'High' || flag === 'Blocked' || escY === 'Yes';
    if (!esc) continue;
    var email = (data[i][4]||'').toString().trim();
    if (!email) continue;
    var emp = getEmployeeByEmail(email.toLowerCase());
    if (!emp) continue;
    var seniors = findSenior_(emp, recips);
    // findSenior_ answers who sits above someone, and above a zone lead is the
    // whole state team. That is the right answer to that question and the
    // wrong list to copy: only one of them wants ten of these a month. So any
    // state address it returns is dropped here, and the named one added back
    // below. Doing it this way keeps findSenior_ honest about the hierarchy
    // and puts the mailing policy in one place.
    var stateSet = {};
    recips.forEach(function(r){ if (r.role === 'State' && r.email) stateSet[r.email.toLowerCase()] = 1; });
    seniors = seniors.filter(function(e){ return !stateSet[e.toLowerCase()]; });

    // findSenior_ goes exactly one level up, which means the state lead only
    // ever saw a zone lead's escalations. Almost all of these come from field
    // officers and stopped at the district lead, so the state had no sight of
    // what was being asked for across the state.
    if (ESC_CC_STATE &&
        ESC_CC_STATE.toLowerCase() !== email.toLowerCase() &&      // their own meeting
        seniors.map(function(x){ return x.toLowerCase(); }).indexOf(ESC_CC_STATE.toLowerCase()) === -1) {
      seniors.push(ESC_CC_STATE);
    }
    var html = buildEscalationEmail_({
      officerName:(data[i][2]||'').toString(), district:(data[i][1]||'').toString(),
      stakeholder:(data[i][9]||'').toString(), purpose:(data[i][11]||'').toString(),
      conductDate:fmtDateVal(data[i][13]), meetingType:(data[i][8]||'').toString(),
      priority:(data[i][22]||'').toString(), flag:(data[i][23]||'').toString(),
      category:(data[i][26]||'').toString(), nextAction:(data[i][24]||'').toString(),
      keyPoints:(data[i][15]||'').toString()
    });
    var to = (mode==='live') ? email : REPORT_TEST_EMAIL;
    var cc = (mode==='live') ? seniors.join(',') : '';
    var subj = 'Escalation: '+(data[i][26]||'Follow-up')+' - '+(data[i][1]||'')+' meeting';
    if (mode !== 'live') subj = '[TEST -> officer:'+email+' | CC senior:'+(seniors.join(',')||'NONE FOUND')+'] '+subj;
    var opts = { to:to, subject:subj, htmlBody:html, name:'EG-MMS Alerts' };
    if (cc) opts.cc = cc;
    try {
      MailApp.sendEmail(opts);
      if (mode === 'live') sh.getRange(i+1, COL_ESC_SENT).setValue(new Date());   // test never marks
      done++; out.push(to + (cc?(' cc '+cc):'') + ' [' + (data[i][26]||'') + ']');
    } catch(e){ out.push('FAIL '+email+' '+e.message); }
  }
  Logger.log('Escalations sent: '+done); Logger.log(out.join('\n'));
  return { success:true, mode:mode, sent:done, details:out };
}

function escalationJob() { return sendEscalations('live', 40); }
function installEscalationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='escalationJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('escalationJob').timeBased().everyHours(1).create();
  return 'Escalation trigger installed: escalationJob runs hourly.';
}
// ---- Run from the editor ----
function ESC_step1_TEST()      { return sendEscalations('test', 25); }   // all to admin (review)
function ESC_step2_LIVE()      { return sendEscalations('live', 40); }   // officer + senior CC
function ESC_installAuto()     { return installEscalationTrigger(); }    // hourly auto
function ESC_stopAuto()        { var n=0; ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='escalationJob'){ ScriptApp.deleteTrigger(t); n++; } }); return 'Escalation paused: removed '+n+' trigger(s). Re-enable later with ESC_installAuto.'; }
function ESC_reset()           { var sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONDUCTED_SHEET); var n=sh.getLastRow(); if(n>1) sh.getRange(2,COL_ESC_SENT,n-1,1).clearContent(); return 'Cleared escalation-sent flags on '+(n-1)+' rows.'; }

// ============================================================
//  TIER 2 - GOOGLE CALENDAR (planned meetings -> officer calendar)
//  Creates a Calendar event for each future Planned/Follow-up meeting and
//  invites the officer (and colleague/stakeholder if an email is present).
//  Event id stored in Plan Meetings col W to avoid duplicates.
// ============================================================
var COL_CAL_EVENT = 23;   // W in Plan Meetings
// The block a block-level official sits in. Kept separate from the officer's
// own block, which is what the sheets already carried and is only a proxy.
var COL_PLAN_SKBLOCK = 24;   // X in Plan Meetings
var COL_CON_SKBLOCK  = 31;   // AE in Conducted Meetings
// What the officer says came out of the meeting. Asked rather than inferred:
// they were there, and this number will end up judging them, so a model's
// guess is the wrong thing to build it on.
var COL_CON_OUTCOME  = 32;   // AF in Conducted Meetings

function parseStart_(dateStr, timeStr) {
  var p = (dateStr||'').toString().trim().split(' '); if (p.length < 3) return null;
  var day = parseInt(p[0],10), mon = _RPT_MONTHS.indexOf(p[1]), yr = parseInt(p[2],10);
  if (isNaN(day) || mon < 0 || isNaN(yr)) return null;
  var h = 10, mi = 0;
  var mt = (timeStr||'').toString().trim().match(/(\d{1,2}):(\d{2})\s*(am|pm)?/i);
  if (mt) { h = parseInt(mt[1],10); mi = parseInt(mt[2],10); var ap=(mt[3]||'').toLowerCase(); if(ap==='pm'&&h<12)h+=12; if(ap==='am'&&h===12)h=0; }
  return new Date(yr, mon, day, h, mi, 0);
}
function durMin_(s){ s=(s||'').toString().toLowerCase(); var m=s.match(/(\d+)/); var n=m?parseInt(m[1],10):0; if(s.indexOf('hour')>=0||s.indexOf('hr')>=0) return (n||1)*60; if(s.indexOf('min')>=0) return n||30; return 60; }

// mode 'test' invites only the admin (review); 'live' invites the officer + stores the event id.
function syncCalendarEvents(mode, limit) {
  mode = mode || 'test'; limit = limit || 20;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(MEETINGS_SHEET);
  if (!sh) return { success:false, message:'no plan sheet' };
  if (!sh.getRange(1, COL_CAL_EVENT).getValue()) sh.getRange(1, COL_CAL_EVENT).setValue('Calendar Event ID');
  var data = sh.getDataRange().getValues();
  var cal = CalendarApp.getDefaultCalendar();
  var done = 0, out = [], now = Date.now();
  for (var i = 1; i < data.length && done < limit; i++) {
    if (!data[i][0]) continue;
    var status = (data[i][13]||'Planned').toString();
    if (status !== 'Planned' && status !== 'Follow-up') continue;
    if ((data[i][COL_CAL_EVENT-1]||'').toString().trim()) continue;   // already synced
    var timeStr = (data[i][6]||'').toString();
    var hasTime = /(\d{1,2}):(\d{2})/.test(timeStr);                  // no time -> all-day event, never a fake 10am
    var start = parseStart_(fmtDateVal(data[i][5]), timeStr);
    if (!start) continue;
    if (start.getTime() < now - 3600000) continue;                    // skip past meetings
    var end = new Date(start.getTime() + durMin_(data[i][7]) * 60000);
    var officer = (data[i][4]||'').toString().trim();
    var title = 'GR Meeting: ' + (data[i][9]||'Stakeholder') + (data[i][11] ? ' (' + data[i][11] + ')' : '');
    var desc = 'Stakeholder: ' + (data[i][9]||'') + ' ' + (data[i][10]||'') +
               '\nPurpose: ' + (data[i][11]||'') + '\nAgenda: ' + (data[i][12]||'') +
               '\nType: ' + (data[i][8]||'') + (data[i][17] ? '\nColleague: ' + data[i][17] : '') + '\nvia EG-MMS';
    var guests = (mode==='live') ? officer : REPORT_TEST_EMAIL;
    try {
      var opts = { description:desc + (hasTime ? '' : '\nTime not specified when planned'), location:(data[i][1]||'').toString(), guests:guests, sendInvites:true };
      var ev = hasTime ? cal.createEvent(title, start, end, opts)
                       : cal.createAllDayEvent(title, start, opts);
      if (mode==='live') sh.getRange(i+1, COL_CAL_EVENT).setValue(ev.getId());
      done++; out.push(title + ' @ ' + start + ' -> ' + guests);
    } catch(e){ out.push('FAIL ' + data[i][0] + ' ' + e.message); }
  }
  Logger.log('Calendar events created: ' + done); Logger.log(out.join('\n'));
  return { success:true, mode:mode, created:done, details:out };
}
function calendarJob() {
  // Keeps the sign-in copy of the employee master current, so a Sheets outage
  // can never stop anyone logging in. Wrapped, because refreshing the copy is
  // not worth failing the calendar sync over.
  try { EMP_refreshMirror(); } catch (e) {}
  try { PURPOSE_refresh(); }   catch (e) {}
  return syncCalendarEvents('live', 30);
}
function installCalendarTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='calendarJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('calendarJob').timeBased().everyHours(1).create();
  return 'Calendar trigger installed: calendarJob runs hourly.';
}
// ---- Run from the editor ----
// Which automatic triggers are actually installed right now?
// ── Letting a besieged document breathe ─────────────────────────────────
// SHEET_probe showed the Sheets service healthy and only our document refusing
// to open. Everything that opens it holds it for up to eight minutes while it
// waits, and the hourly jobs, the half hourly snapshot and every app request
// kept arriving before the last ones had died, so the document never came free.
// These two take the hourly jobs off for a while and put them back exactly as
// they were. The monthly report and the weekly nudge are left alone.
var HOURLY_JOBS = ['taggingJob', 'escalationJob', 'calendarJob'];

function TRIGGERS_pauseHourly() {
  var removed = [];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (HOURLY_JOBS.indexOf(fn) >= 0) { ScriptApp.deleteTrigger(t); removed.push(fn); }
  });
  PropertiesService.getScriptProperties().setProperty('PAUSED_HOURLY', JSON.stringify(removed));
  Logger.log('Paused: ' + (removed.join(', ') || 'nothing was installed') +
             String.fromCharCode(10) + 'Put them back later with TRIGGERS_restoreHourly().');
  return removed;
}

function TRIGGERS_restoreHourly() {
  var was = [];
  try { was = JSON.parse(PropertiesService.getScriptProperties().getProperty('PAUSED_HOURLY') || '[]'); } catch (e) {}
  var have = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
  var back = [];
  if (was.indexOf('taggingJob')    >= 0 && have.indexOf('taggingJob')    < 0) { installTaggingTrigger();    back.push('taggingJob'); }
  if (was.indexOf('escalationJob') >= 0 && have.indexOf('escalationJob') < 0) { installEscalationTrigger(); back.push('escalationJob'); }
  if (was.indexOf('calendarJob')   >= 0 && have.indexOf('calendarJob')   < 0) { installCalendarTrigger();   back.push('calendarJob'); }
  PropertiesService.getScriptProperties().deleteProperty('PAUSED_HOURLY');
  Logger.log('Restored: ' + (back.join(', ') || 'nothing to restore'));
  return back;
}

// ── Plan B: a fresh copy of the document ────────────────────────────────
// Goes through Drive rather than Sheets, so it can copy a document that Sheets
// itself will not open. Prints the new id and url. NOTHING is switched over by
// this: the script keeps using the old document until SPREADSHEET_ID in the code
// is changed to the new id and redeployed. The old file stays where it is.
function SHEET_makeCopy() {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var src   = DriveApp.getFileById(SPREADSHEET_ID);
  var copy  = src.makeCopy(src.getName() + ' (copy ' + stamp + ')');
  Logger.log('Copied.' + String.fromCharCode(10) +
             'New id : ' + copy.getId() + String.fromCharCode(10) +
             'New url: ' + copy.getUrl() + String.fromCharCode(10) +
             'Send the id back to switch the script over to it.');
  return copy.getId();
}

function TRIGGER_status() {
  var t = ScriptApp.getProjectTriggers().map(function(x){ return x.getHandlerFunction(); });
  var want = ['monthlyReportJob','taggingJob','escalationJob','calendarJob','nudgeJob'];
  var out = {};
  want.forEach(function(w){ out[w] = t.indexOf(w) >= 0 ? 'INSTALLED' : 'not installed'; });
  out._allTriggers = t;
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// Full picture: is the hourly job on, which calendar do events land on,
// and which upcoming meetings are still unsynced?
function CAL_status() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(MEETINGS_SHEET);
  if (!sh) return { error:'no plan sheet' };
  var data = sh.getDataRange().getValues(), now = Date.now(), up = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var status = (data[i][13]||'Planned').toString();
    if (status !== 'Planned' && status !== 'Follow-up') continue;
    var timeStr = (data[i][6]||'').toString();
    var start = parseStart_(fmtDateVal(data[i][5]), timeStr);
    if (!start || start.getTime() < now - 3600000) continue;          // upcoming only
    up.push({ id:data[i][0], date:fmtDateVal(data[i][5]), time:timeStr || '(no time)',
              officer:(data[i][4]||'').toString(),
              synced:(data[i][COL_CAL_EVENT-1]||'').toString().trim() ? 'YES' : 'NO' });
  }
  var trig = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
  var res = {
    hourlyTrigger: trig.indexOf('calendarJob') >= 0 ? 'INSTALLED' : 'NOT INSTALLED',
    eventsLandOnCalendar: CalendarApp.getDefaultCalendar().getName(),
    upcomingCount: up.length,
    upcoming: up
  };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function CAL_test()        { return syncCalendarEvents('test', 5); }    // 5 events, invite admin only (review)
function CAL_live()        { return syncCalendarEvents('live', 30); }   // invite officers, store ids
function CAL_installAuto() { return installCalendarTrigger(); }         // hourly auto

// ============================================================
//  WEEKLY REMINDER (Monday, about 8am)
//  One email to each officer listing the meetings they have in the coming
//  seven days: when, with whom, where, about what. Nothing else. Defined by
//  Alok on 27 Sep 2026 as "is week me kab koun si meeting hai", in English.
//  The earlier version also carried pending follow-ups, last week's meetings
//  and a team summary for leads; it was never switched on and is gone.
//
//  Officers with no meeting that week get nothing.
//  Rescheduled meetings are included. Postponing moves the plan row to its new
//  date with status Postponed, and the old version, which looked only at
//  Planned and Follow-up, would have left every rescheduled meeting out.
// ============================================================
var NUDGE_STATUSES = { 'Planned':1, 'Follow-up':1, 'Postponed':1 };

function weeklyMeetingsByOfficer_(now) {
  var pd = sheetRows_(MEETINGS_SHEET) || [];
  var until = now + 7 * 86400000, tz = Session.getScriptTimeZone(), by = {};
  for (var j = 1; j < pd.length; j++) {
    if (!pd[j][0]) continue;
    var st = (pd[j][13] || 'Planned').toString().trim();
    if (!NUDGE_STATUSES[st]) continue;
    var email = (pd[j][4] || '').toString().trim().toLowerCase();
    if (!email) continue;
    var time = fmtTimeVal(pd[j][6]);
    var start = parseStart_(fmtDateVal(pd[j][5]), time);
    if (!start) continue;
    var t = start.getTime();
    if (t < now - 3600000 || t > until) continue;
    var o = by[email] || (by[email] = { email:email, name:(pd[j][2] || '').toString().trim(), meetings:[] });
    o.meetings.push({
      t: t,
      when: Utilities.formatDate(start, tz, 'EEE d MMM'),
      time: /\d/.test(time || '') ? time : '',
      official: (pd[j][9] || '').toString().trim(),
      post: (pd[j][10] || '').toString().trim(),
      district: (pd[j][1] || '').toString().trim(),
      block: (pd[j][COL_PLAN_SKBLOCK - 1] || '').toString().trim(),
      purpose: (pd[j][11] || '').toString().trim(),
      colleague: (pd[j][17] || '').toString().trim(),
      moved: st === 'Postponed'
    });
  }
  Object.keys(by).forEach(function(k){ by[k].meetings.sort(function(a, b){ return a.t - b.t; }); });
  return by;
}

function buildWeeklyReminder_(o, range) {
  var first = ((o.name || '').split(' ')[0]) || 'there';
  var n = o.meetings.length;
  var rows = o.meetings.map(function(m){
    return '<tr style="border-top:1px solid #f0ebe5;">' +
      '<td style="padding:10px 12px;white-space:nowrap;vertical-align:top;"><b>' + _emailEsc(m.when) + '</b>' +
        (m.time ? '<br><span style="color:#6b7280;">' + _emailEsc(m.time) + '</span>' : '') + '</td>' +
      '<td style="padding:10px 12px;vertical-align:top;"><b>' + _emailEsc(m.official || m.post || 'Official') + '</b>' +
        (m.official && m.post ? ', ' + _emailEsc(m.post) : '') +
        '<br><span style="color:#6b7280;">' + _emailEsc(m.district || 'State level') + (m.block ? ' / ' + _emailEsc(m.block) : '') +
        (m.purpose ? ' &middot; ' + _emailEsc(m.purpose) : '') + '</span>' +
        (m.colleague ? '<br><span style="color:#6b7280;">With ' + _emailEsc(m.colleague) + '</span>' : '') +
        (m.moved ? '<br><span style="color:#92400E;font-size:12px;font-weight:600;">Rescheduled</span>' : '') +
      '</td></tr>';
  }).join('');
  return '<div style="margin:0;padding:20px 12px;background:#f4f2ef;font-family:Arial,Helvetica,sans-serif;color:#1f2937;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:620px;margin:0 auto;background:#fff;border:1px solid #e5e7eb;border-radius:12px;">' +
    '<tr><td style="padding:22px 28px 12px;border-bottom:2px solid #7B1010;">' +
      '<div style="font-size:11px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#7B1010;">Your meetings this week</div>' +
      '<h1 style="font-family:Georgia,serif;font-size:20px;margin:8px 0 3px;">' + n + ' meeting' + (n === 1 ? '' : 's') + ', ' + _emailEsc(range) + '</h1></td></tr>' +
    '<tr><td style="padding:14px 28px 4px;font-size:14px;color:#374151;">Good morning ' + _emailEsc(first) +
      '. Here is what you have planned for the week ahead.</td></tr>' +
    '<tr><td style="padding:8px 28px 0;"><table width="100%" cellpadding="0" cellspacing="0" ' +
      'style="border:1px solid #e5e7eb;border-collapse:collapse;font-size:13px;">' + rows + '</table></td></tr>' +
    '<tr><td style="padding:16px 28px 0;font-size:13px;color:#6b7280;">If a meeting has moved or will not happen, ' +
      'update it in Manage Meetings so the record stays right.</td></tr>' +
    '<tr><td style="padding:16px 28px 24px;"><div style="border-top:1px solid #e5e7eb;padding-top:12px;font-size:11px;color:#9ca3af;">' +
      'EG-MMS &middot; https://dataimpact.in</div></td></tr>' +
    '</table></div>';
}

// mode 'preview': sends nothing, logs who would get one.
// mode 'test':    up to three sample emails, all to REPORT_TEST_EMAIL.
// mode 'live':    every officer with a meeting in the next seven days.
function sendWeeklyNudges(mode) {
  mode = mode || 'test';
  var now = Date.now(), tz = Session.getScriptTimeZone();
  var range = Utilities.formatDate(new Date(now), tz, 'd MMM') + ' to ' +
              Utilities.formatDate(new Date(now + 6 * 86400000), tz, 'd MMM');
  var by = weeklyMeetingsByOfficer_(now), out = [], done = 0, officers = 0, total = 0;
  Object.keys(by).sort().forEach(function(em){
    var o = by[em];
    if (!o.meetings.length) return;
    officers++; total += o.meetings.length;
    if (mode === 'preview') { out.push(o.name + ' <' + o.email + '>: ' + o.meetings.length); return; }
    if (mode === 'test' && done >= 3) return;
    var to = (mode === 'live') ? o.email : REPORT_TEST_EMAIL;
    try {
      MailApp.sendEmail({
        to: to, name: 'EG-MMS', htmlBody: buildWeeklyReminder_(o, range),
        subject: (mode !== 'live' ? '[TEST -> ' + o.email + '] ' : '') +
                 'Your meetings this week: ' + o.meetings.length + ' (' + range + ')'
      });
      out.push(to + ' (' + o.meetings.length + ')'); done++;
    } catch (e) { out.push('FAIL ' + o.email + ': ' + e.message); }
  });
  Logger.log(officers + ' officer(s) have ' + total + ' meeting(s) between ' + range + '. ' +
             (mode === 'preview' ? 'Nothing sent.' : done + ' email(s) sent (' + mode + ').'));
  Logger.log(out.join('\n'));
  return { success:true, mode:mode, officers:officers, meetings:total, sent:done, details:out };
}
function nudgeJob() { return sendWeeklyNudges('live'); }
function installNudgeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='nudgeJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('nudgeJob').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  return 'Weekly reminder switched on: nudgeJob runs every Monday around 8am.';
}
// ---- Run from the editor ----
function NUDGE_preview()     { return sendWeeklyNudges('preview'); } // who would get one, sends nothing
function NUDGE_test()        { return sendWeeklyNudges('test'); }    // up to 3 samples, all to REPORT_TEST_EMAIL
function NUDGE_live()        { return sendWeeklyNudges('live'); }    // every officer with a meeting, right now
function NUDGE_installAuto() { return installNudgeTrigger(); }       // every Monday around 8am
function NUDGE_stopAuto() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='nudgeJob') { ScriptApp.deleteTrigger(t); n++; } });
  return 'Weekly reminder switched off: removed ' + n + ' trigger(s).';
}

// ============================================================
//  MONTHLY REPORT EMAIL DELIVERY
//  Recipients = State / Zone / District leads (each their own scope).
//  Sent from gr@educategirls.ngo via MailApp. Run installMonthlyTrigger()
//  once to schedule for the 1st of each month.
// ============================================================
var REPORT_TEST_EMAIL = 'alok.mohan@educategirls.ngo';   // used by mode 'test'

function getReportRecipients() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(EMPLOYEE_SHEET);
  var data = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    var email = (data[i][4] || '').toString().trim();
    if (!email) continue;
    var role = normalizeRole_(data[i][5]);
    if (role !== 'State' && role !== 'Zone' && role !== 'District') continue;
    var district = (data[i][0] || '').toString().trim();
    var districts = [district];
    (data[i][7] || '').toString().split(/[,;]/).forEach(function(x){
      var d = x.trim();
      if (d && districts.map(function(z){ return z.toLowerCase(); }).indexOf(d.toLowerCase()) === -1) districts.push(d);
    });
    out.push({ name:(data[i][2]||'').toString().trim(), email:email, role:role,
               zone:(data[i][6]||'').toString().trim(), district:district, districts:districts });
  }
  return out;
}

// Preview who would receive the monthly email (names/roles/emails). Admin only.
function previewReportRecipients() {
  var r = getReportRecipients();
  return { success:true, count:r.length,
    recipients: r.map(function(x){ return { name:x.name, email:x.email, role:x.role, scope:(x.role==='State'?'Uttar Pradesh':x.role==='Zone'?x.zone:x.districts.join(', ')) }; }) };
}

function _emailEsc(s){ return (s==null?'':String(s)).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function _pctColor(p){ return p>=80?'#166534':p>=50?'#166534':p>=30?'#9a5b0e':'#991b1b'; }

function buildReportEmailHtml(rep, recipientName) {
  var k = rep.kpis, sc = rep.scope, b = rep.breakdown;
  var byLabel = { zone:'Zone', district:'District', block:'Block' }[b.by] || 'Area';
  var SERIF = "font-family:'Spectral',Georgia,'Times New Roman',serif;";
  function sech(title, tag, cls){
    var ts = (cls==='ai') ? 'color:#7B1010;background:#f6e9e9;' : 'color:#6b7280;background:#f7f2ee;border:1px solid #e5e7eb;';
    return '<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 12px;"><tr>'+
      '<td style="'+SERIF+'font-size:17px;font-weight:700;color:#1f2937;">'+title+'</td>'+
      '<td align="right"><span style="font-size:10px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;padding:3px 9px;border-radius:20px;'+ts+'">'+tag+'</span></td>'+
      '</tr></table>';
  }
  function pill(planned, p){
    var c,bg,t;
    if(!planned){c='#991b1b';bg='#fbe9e9';t='No activity';}
    else if(p>=80){c='#166534';bg='#e6f0e8';t='On track';}
    else if(p>=50){c='#166534';bg='#e6f0e8';t='Steady';}
    else if(p>=30){c='#9a5b0e';bg='#faf0de';t='Watch';}
    else {c='#991b1b';bg='#fbe9e9';t='Attention';}
    return '<span style="font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;color:'+c+';background:'+bg+';white-space:nowrap;">'+t+'</span>';
  }
  function tile(lbl, val, color, sub){
    return '<td width="33%" style="background:#fafafa;border:1px solid #e5e7eb;border-radius:10px;padding:12px 14px;">'+
      '<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#6b7280;letter-spacing:.5px;">'+lbl+'</div>'+
      '<div style="'+SERIF+'font-size:23px;font-weight:700;color:'+(color||'#1f2937')+';margin-top:5px;">'+val+'</div>'+
      (sub?'<div style="font-size:11px;color:#6b7280;margin-top:2px;">'+sub+'</div>':'')+'</td>';
  }
  function sec(inner){ return '<tr><td style="padding:24px 30px 0;">'+inner+'</td></tr>'; }

  // Performance table
  var zsum = 0; (b.rows||[]).forEach(function(r){ zsum += (r.districts||0); });
  var rows = (b.rows||[]).slice(0,15).map(function(r){
    return '<tr style="border-top:1px solid #f0ebe5;"><td style="padding:10px 12px;font-weight:700;">'+_emailEsc(r.name)+'</td>'+
      (b.by==='zone'?'<td align="right" style="padding:10px 12px;">'+(r.districts||0)+'</td>':'')+
      '<td align="right" style="padding:10px 12px;">'+r.planned+'</td><td align="right" style="padding:10px 12px;">'+r.conducted+'</td>'+
      '<td align="right" style="padding:10px 12px;color:'+_pctColor(r.pct)+';font-weight:700;">'+(r.planned?r.pct+'%':'-')+'</td>'+
      '<td style="padding:10px 12px;">'+pill(r.planned,r.pct)+'</td></tr>';
  }).join('');
  var zoneTotal = (b.by==='zone') ? '<tr style="border-top:1px solid #e5e7eb;background:#f7f2ee;font-weight:700;font-size:13px;"><td style="padding:10px 12px;">State total</td><td align="right" style="padding:10px 12px;">'+zsum+'</td><td align="right" style="padding:10px 12px;">'+k.total+'</td><td align="right" style="padding:10px 12px;">'+k.conducted+'</td><td align="right" style="padding:10px 12px;">'+k.success+'%</td><td></td></tr>' : '';
  var perfTable = sec(sech('Performance by '+byLabel,'Computed','calc')+
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;border:1px solid #e5e7eb;border-radius:8px;">'+
    '<tr style="background:#f7f2ee;color:#6b7280;font-size:11px;text-transform:uppercase;"><th align="left" style="padding:10px 12px;">'+byLabel+'</th>'+(b.by==='zone'?'<th align="right" style="padding:10px 12px;">Dist</th>':'')+'<th align="right" style="padding:10px 12px;">Planned</th><th align="right" style="padding:10px 12px;">Conducted</th><th align="right" style="padding:10px 12px;">Success</th><th align="left" style="padding:10px 12px;">Status</th></tr>'+
    rows + zoneTotal + '</table>');

  // Leaderboard (state)
  var lb = '';
  if (b.leaderboard && b.leaderboard.length) {
    var lrows = b.leaderboard.map(function(x,i){
      return '<tr style="border-top:1px solid #f0ebe5;"><td style="padding:10px 12px;color:#a8a29e;font-weight:700;">'+(i+1)+'</td><td style="padding:10px 12px;font-weight:700;">'+_emailEsc(x.name)+'</td><td style="padding:10px 12px;color:#6b7280;">'+_emailEsc((x.zone||'').replace('UP ',''))+'</td><td align="right" style="padding:10px 12px;color:#6b7280;">'+x.planned+'</td><td align="right" style="padding:10px 12px;">'+x.conducted+'</td><td align="right" style="padding:10px 12px;color:'+_pctColor(x.pct)+';font-weight:700;">'+x.pct+'%</td></tr>';
    }).join('');
    lb = sec(sech('District Leaderboard',b.leaderboard.length+' active','calc')+
      '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;border:1px solid #e5e7eb;border-radius:8px;">'+
      '<tr style="background:#f7f2ee;color:#6b7280;font-size:11px;text-transform:uppercase;"><th align="left" style="padding:10px 12px;">#</th><th align="left" style="padding:10px 12px;">District</th><th align="left" style="padding:10px 12px;">Zone</th><th align="right" style="padding:10px 12px;">Planned</th><th align="right" style="padding:10px 12px;">Conducted</th><th align="right" style="padding:10px 12px;">Success</th></tr>'+
      lrows + '</table>');
  }

  // Participation
  var part = sec(sech('Team Participation','Computed','calc')+
    '<div style="background:#f7f2ee;border:1px solid #e5e7eb;border-radius:10px;padding:16px 18px;">'+
    '<table width="100%" cellpadding="0" cellspacing="0"><tr>'+
    '<td width="92" valign="top"><div style="'+SERIF+'font-size:30px;font-weight:700;color:#166534;">'+k.participation+'%</div><div style="font-size:11px;color:#6b7280;">participation</div></td>'+
    '<td valign="middle" style="padding-left:14px;"><div style="height:9px;background:#e5e7eb;border-radius:6px;"><div style="height:9px;width:'+Math.max(2,Math.min(100,k.participation))+'%;background:#166534;border-radius:6px;font-size:1px;">&nbsp;</div></div>'+
    '<div style="font-size:12px;color:#6b7280;margin-top:8px;"><b style="color:#166534;">'+k.activeStaff+'</b> active &nbsp;&middot;&nbsp; <b style="color:#991b1b;">'+(k.totalStaff-k.activeStaff)+'</b> inactive of '+k.totalStaff+' staff</div></td></tr></table></div>');

  // Meeting focus
  function focusCol(title, arr){
    var body = (arr||[]).map(function(x){ return '<tr><td style="padding:5px 0;font-weight:600;">'+_emailEsc(x.name)+'</td><td align="right" style="padding:5px 0;font-weight:700;">'+x.count+'</td></tr>'; }).join('') || '<tr><td style="color:#a8a29e;padding:5px 0;">No data</td></tr>';
    return '<td width="50%" valign="top" style="padding:0 8px;"><div style="border:1px solid #e5e7eb;border-radius:10px;padding:10px 14px;"><div style="font-size:12px;font-weight:700;color:#1f2937;margin-bottom:4px;">'+title+'</div><table width="100%" style="font-size:13px;">'+body+'</table></div></td>';
  }
  var focus = (rep.byPurpose&&rep.byPurpose.length || rep.byStakeholder&&rep.byStakeholder.length) ?
    sec(sech('Meeting Focus','Computed','calc')+'<table width="100%" cellpadding="0" cellspacing="0"><tr>'+focusCol('By Purpose',rep.byPurpose)+focusCol('By Stakeholder',rep.byStakeholder)+'</tr></table>') : '';

  // Outcomes: what came out of the meetings, not just how many happened.
  var oc = rep.outcomes;
  var ocSec = '';
  if (oc && oc.answered) {
    var ocRows = (oc.counts || []).map(function(x){
      return '<tr style="border-top:1px solid #f0ebe5;"><td style="padding:8px 12px;font-weight:600;">'+_emailEsc(x.name)+'</td>'+
        '<td align="right" style="padding:8px 12px;font-weight:700;">'+x.count+'</td></tr>';
    }).join('');
    ocSec = sec(sech('Meeting Outcomes','Reported','calc')+
      '<table width="100%" cellpadding="0" cellspacing="0"><tr>'+
        '<td width="50%" style="padding:0 5px;"><div style="background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;padding:12px 10px;text-align:center;">'+
          '<div style="font-family:Georgia,serif;font-size:24px;font-weight:700;color:#166534;">'+oc.rate+'%</div>'+
          '<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#6b7280;margin-top:2px;">Produced something</div></div></td>'+
        '<td width="50%" style="padding:0 5px;"><div style="background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;padding:12px 10px;text-align:center;">'+
          '<div style="font-family:Georgia,serif;font-size:24px;font-weight:700;color:#1f2937;">'+oc.concrete+' / '+oc.answered+'</div>'+
          '<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#6b7280;margin-top:2px;">Meetings</div></div></td>'+
      '</tr></table>'+
      '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;border:1px solid #e5e7eb;border-radius:8px;margin-top:12px;">'+
      '<tr style="background:#f7f2ee;color:#6b7280;font-size:11px;text-transform:uppercase;"><th align="left" style="padding:10px 12px;">Outcome</th><th align="right" style="padding:10px 12px;">Meetings</th></tr>'+
      ocRows+'</table>'+
      '<div style="font-size:11px;color:#9ca3af;margin-top:8px;">Reported by the officer on the conduct form. A courtesy visit is a normal part of the work, not a poor result.</div>');
  }

  // Relationship health: repeat contact is the actual work, so this reads the
  // whole history of each official rather than this month's meeting count.
  var rel = rep.relationships;
  var relSec = '';
  if (rel && rel.totalOffices) {
    function relTile(label, val, colour) {
      return '<td width="25%" style="padding:0 5px;"><div style="background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;padding:12px 10px;text-align:center;">'+
        '<div style="font-family:Georgia,serif;font-size:24px;font-weight:700;color:'+colour+';">'+val+'</div>'+
        '<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#6b7280;margin-top:2px;">'+label+'</div></div></td>';
    }
    var coldRows = (rel.cold||[]).map(function(r){
      return '<tr style="border-top:1px solid #f0ebe5;">'+
        '<td style="padding:8px 12px;"><b>'+_emailEsc(r.post)+'</b>'+(r.lastPerson?'<br><span style="color:#6b7280;font-size:12px;">last met: '+_emailEsc(r.lastPerson)+'</span>':'')+'</td>'+
        '<td style="padding:8px 12px;color:#6b7280;font-size:12px;">'+_emailEsc(r.district)+'</td>'+
        '<td align="right" style="padding:8px 12px;color:#6b7280;font-size:12px;">'+r.count+'</td>'+
        '<td align="right" style="padding:8px 12px;color:#991b1b;font-weight:700;font-size:12px;white-space:nowrap;">'+r.daysSince+' days</td></tr>';
    }).join('');
    var coldTable = coldRows ?
      '<div style="font-size:12px;font-weight:700;color:#991b1b;margin:14px 0 6px;">No contact in '+rel.coldDays+'+ days ('+rel.coldCount+')</div>'+
      '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;border:1px solid #e5e7eb;border-radius:8px;">'+
      '<tr style="background:#f7f2ee;color:#6b7280;font-size:11px;text-transform:uppercase;"><th align="left" style="padding:10px 12px;">Office</th><th align="left" style="padding:10px 12px;">District</th><th align="right" style="padding:10px 12px;">Meetings</th><th align="right" style="padding:10px 12px;">Last contact</th></tr>'+
      coldRows+'</table>'+
      (rel.coldCount > (rel.cold||[]).length ? '<div style="font-size:11px;color:#9ca3af;margin-top:6px;">Showing the '+(rel.cold||[]).length+' longest gaps of '+rel.coldCount+'.</div>' : '') : '';
    relSec = sec(sech('Relationship Health','Computed','calc')+
      '<table width="100%" cellpadding="0" cellspacing="0"><tr>'+
        relTile('Offices engaged', rel.totalOffices, '#1f2937')+
        relTile('Met this month',    rel.metThisMonth,   '#166534')+
        relTile('First time',        rel.newThisMonth,   '#1D4ED8')+
        relTile('Gone quiet',        rel.coldCount,      '#991b1b')+
      '</tr></table>'+
      '<div style="font-size:12px;color:#6b7280;margin-top:10px;">'+rel.onlyOnce+' of '+rel.totalOffices+' offices have been met only once. Tracked by office, since officials get transferred.</div>'+
      coldTable);
  }

  // Dot lists
  function dotList(items, dotColorFn){
    var body = items.map(function(it){
      return '<tr><td width="16" valign="top" style="padding:9px 0;"><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:'+dotColorFn(it)+';">&nbsp;</span></td>'+
        '<td style="padding:9px 0;"><b>'+_emailEsc(it.h)+'</b>'+(it.d?'<br><span style="color:#6b7280;font-size:13px;">'+_emailEsc(it.d)+'</span>':'')+'</td></tr>';
    }).join('');
    return '<table width="100%" cellpadding="0" cellspacing="0" style="font-size:13.5px;">'+body+'</table>';
  }
  var attItems = (rep.attention||[]).slice(0,4).map(function(a){ return { h:a.title, d:a.detail, level:a.level }; });
  var attention = attItems.length ? sec(sech('Attention Needed','Rules','calc')+dotList(attItems, function(it){ return it.level==='crit'?'#991b1b':it.level==='good'?'#166534':'#9a5b0e'; })) : '';
  var hiItems = (rep.narrative&&rep.narrative.highlights)||[];
  var highlights = hiItems.length ? sec(sech('Highlights','AI-written','ai')+dotList(hiItems, function(){ return '#7B1010'; })) : '';

  // Recommendations (numbered)
  var recItems = (rep.narrative&&rep.narrative.recommendations)||[];
  var recBody = recItems.map(function(x,i){
    return '<tr><td width="34" valign="top" style="padding:9px 0;"><span style="display:inline-block;width:24px;height:24px;line-height:24px;text-align:center;border-radius:7px;background:#f6e9e9;color:#7B1010;font-weight:700;font-size:13px;">'+(i+1)+'</span></td>'+
      '<td style="padding:9px 0 9px 6px;"><b>'+_emailEsc(x.h)+'</b>'+(x.d?' <span style="color:#6b7280;font-size:13px;">'+_emailEsc(x.d)+'</span>':'')+'</td></tr>';
  }).join('');
  var recommendations = recItems.length ? sec(sech('Recommendations','AI-written','ai')+'<table width="100%" cellpadding="0" cellspacing="0" style="font-size:13.5px;">'+recBody+'</table>') : '';

  return '<div style="margin:0;padding:24px 12px;background:#f4f2ef;font-family:Arial,Helvetica,sans-serif;color:#1f2937;">'+
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:680px;margin:0 auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;">'+
    '<tr><td style="padding:28px 30px 16px;border-bottom:2px solid #7B1010;">'+
      '<div style="font-size:11px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;color:#7B1010;">Educate Girls &middot; Government Relations</div>'+
      '<h1 style="'+SERIF+'font-size:25px;line-height:1.12;margin:9px 0 5px;color:#1f2937;">Monthly GR Meetings Report</h1>'+
      '<div style="font-size:14px;color:#6b7280;"><b style="color:#1f2937;">'+_emailEsc(sc.label)+'</b> &middot; '+_emailEsc(sc.month)+'</div></td></tr>'+
    '<tr><td style="padding:18px 30px 0;font-size:13px;color:#6b7280;">Dear '+_emailEsc(recipientName||'Colleague')+', here is your '+_emailEsc(sc.kind)+'-level summary for '+_emailEsc(sc.month)+'.</td></tr>'+
    sec(sech('Executive Summary','AI-written','ai')+'<div style="background:#f7f2ee;border:1px solid #e5e7eb;border-left:3px solid #7B1010;border-radius:10px;padding:16px 20px;font-size:14.5px;line-height:1.6;">'+_emailEsc(rep.narrative.summary)+'</div>')+
    sec(sech('At a Glance','Computed','calc')+
      '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:8px;">'+
      '<tr>'+tile('Total',k.total,'','Planned in month')+tile('Conducted',k.conducted,'#166534','This month')+tile('Success',k.success+'%','#7B1010','Conducted vs planned')+'</tr>'+
      '<tr>'+tile('Active Staff',k.activeStaff+' / '+k.totalStaff,'',k.participation+'% participation')+tile('Pending',k.pending,'#9a5b0e','Open in month')+tile('Govt MoM',k.govtMom+' / '+k.conducted,'','Official minutes')+'</tr></table>')+
    perfTable + lb + part + focus + ocSec + relSec + attention + highlights + recommendations +
    '<tr><td style="padding:24px 30px 26px;"><div style="border-top:1px solid #e5e7eb;padding-top:14px;font-size:11px;color:#9ca3af;line-height:1.6;">Numbers computed from records; summary and recommendations written by AI. Full analytics portal: https://dataimpact.in/report.html<br>EG-MMS &middot; automated monthly report.</div></td></tr>'+
    '</table></div>';
}

// mode: 'test' sends every report to REPORT_TEST_EMAIL; 'live' sends to each lead.
function sendMonthlyReports(mode, monthOverride, roleFilter) {
  mode = mode || 'test';
  var recips = getReportRecipients();
  if (roleFilter) recips = recips.filter(function(r){ return r.role === roleFilter; });
  var month = monthOverride || prevMonthKey_();   // default: the just-completed month (Aug on 1 Sep)
  var sent = [], failed = [];
  recips.forEach(function(r){
    try {
      var rep = getMonthlyReport({ role:r.role, zone:r.zone, district:r.district, districts:r.districts, email:r.email, name:r.name }, month);
      if (!rep || !rep.success) { failed.push(r.email + ' (no report)'); return; }
      var html = buildReportEmailHtml(rep, r.name);
      var to = (mode === 'live') ? r.email : REPORT_TEST_EMAIL;
      var attach = [];
      try {
        attach.push(Utilities.newBlob('<html><head><meta charset="utf-8"></head><body>' + html + '</body></html>', 'text/html', 'report.html')
          .getAs('application/pdf')
          .setName('GR-Report-' + rep.scope.label.replace(/[^A-Za-z0-9]+/g,'-') + '-' + rep.scope.month.replace(/\s/g,'') + '.pdf'));
      } catch(pe) { /* PDF optional - send without it if conversion fails */ }
      MailApp.sendEmail({ to:to, subject:'Monthly GR Report - ' + rep.scope.label + ' - ' + rep.scope.month, htmlBody:html, name:'EG-MMS Reports', attachments:attach });
      sent.push(to + ' [' + r.role + ': ' + (r.role==='State'?'UP':r.role==='Zone'?r.zone:r.district) + ']');
    } catch(e){ failed.push(r.email + ' ' + e.message); }
  });
  Logger.log('MODE=' + mode + ' | sent=' + sent.length + ' failed=' + failed.length);
  Logger.log(sent.join('\n'));
  if (failed.length) Logger.log('FAILED:\n' + failed.join('\n'));
  return { success:true, mode:mode, sentCount:sent.length, failedCount:failed.length, sent:sent, failed:failed };
}

function monthlyReportJob() { return sendMonthlyReports('live'); }   // uses prevMonthKey_() = just-completed month

// ---- Run these from the editor, in order ----
function REPORT_step1_TEST()        { return sendMonthlyReports('test'); }   // all reports to admin only (review)
function REPORT_step2_SEND_LIVE()   { return sendMonthlyReports('live'); }   // real send to all 33 leads
function REPORT_step3_INSTALL_AUTO(){ return installMonthlyTrigger(); }      // schedule for the 1st of each month

// ---- Re-send only to STATE-level people (e.g. after adding a new state name) ----
function REPORT_STATE_preview() {   // see the State recipients (confirm the new name is here)
  var r = getReportRecipients().filter(function(x){ return x.role === 'State'; });
  return { count:r.length, recipients:r.map(function(x){ return { name:x.name, email:x.email }; }) };
}
function REPORT_STATE_test()   { return sendMonthlyReports('test', null, 'State'); }   // State reports to admin only (review)
function REPORT_STATE_live()   { return sendMonthlyReports('live', null, 'State'); }   // real send to State people only

function installMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction() === 'monthlyReportJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('monthlyReportJob').timeBased().onMonthDay(1).atHour(7).create();
  return 'Monthly trigger installed: monthlyReportJob runs on the 1st of every month at ~7am, sending live reports to all leads.';
}

function removeMonthlyTrigger() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction() === 'monthlyReportJob') { ScriptApp.deleteTrigger(t); n++; } });
  return 'Removed ' + n + ' monthly trigger(s).';
}

// ------------------------------------------------------------
//  EMPLOYEE MASTER (public) - name/designation/district/block only
//  (no email/role). Powers the coverage & active/inactive reports.
// ------------------------------------------------------------
function getEmployeeMaster() {
  try {
    var cacheKey = 'empMaster';
    var hit = cGet(cacheKey);
    if (hit) return hit;
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
    var out   = { success: true, employees: [] };
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      // Cols: District(0) Block(1) Name(2) Designation(3) Email(4) Role(5) Zone(6)
      for (var i = 1; i < data.length; i++) {
        var name = (data[i][2] || '').toString().trim();
        if (!name) continue;
        out.employees.push({
          name:        name,
          designation: (data[i][3] || '').toString().trim(),
          district:    (data[i][0] || '').toString().trim(),
          block:       (data[i][1] || '').toString().trim()
        });
      }
    }
    cPut(cacheKey, out, C_TTL_DROP);
    return out;
  } catch(err) { return { success: false, message: err.message }; }
}

function getReportData() {
  try {
    var cacheKey = 'reportData';
    var hit = cGet(cacheKey);
    if (hit) return hit;

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);

    // email → block map
    var emp = ss.getSheetByName(EMPLOYEE_SHEET);
    var blockMap = {};
    if (emp) {
      var ed = (sheetRows_(EMPLOYEE_SHEET) || []);
      for (var i = 1; i < ed.length; i++) {
        var em = (ed[i][4] || '').toString().trim().toLowerCase();
        if (em) blockMap[em] = (ed[i][1] || '').toString().trim(); // B = Block
      }
    }
    function blk(email) { return blockMap[(email || '').toString().trim().toLowerCase()] || ''; }

    var meetings = [];

    // 1. Plan Meetings - Planned / Follow-up only
    var plan = ss.getSheetByName(MEETINGS_SHEET);
    if (plan) {
      var pd = (sheetRows_(MEETINGS_SHEET) || []);
      for (var a = 1; a < pd.length; a++) {
        if (!pd[a][0]) continue;
        var st = (pd[a][13] || 'Planned').toString();
        if (st !== 'Planned' && st !== 'Follow-up') continue;
        meetings.push({
          meetingId:(pd[a][0]||'').toString(), district:(pd[a][1]||'').toString(),
          block: blk(pd[a][4]), employeeName:(pd[a][2]||'').toString(), post:(pd[a][3]||'').toString(),
          status: st, date: fmtDateVal(pd[a][5]), conductDate:'',
          meetingType:(pd[a][8]||'').toString(), stakeholderName:(pd[a][9]||'').toString(),
          stakeholderPost:(pd[a][10]||'').toString(), stakeholderBlock:(pd[a][COL_PLAN_SKBLOCK-1]||'').toString(), purpose:(pd[a][11]||'').toString(),
          momUrl:'', photoUrl:'', colleagueName:(pd[a][17]||'').toString()
        });
      }
    }

    // 2. Conducted
    var cS = ss.getSheetByName(CONDUCTED_SHEET);
    if (cS) {
      var cd = (sheetRows_(CONDUCTED_SHEET) || []);
      for (var b = 1; b < cd.length; b++) {
        if (!cd[b][0]) continue;
        meetings.push({
          meetingId:(cd[b][0]||'').toString(), district:(cd[b][1]||'').toString(),
          block: blk(cd[b][4]), employeeName:(cd[b][2]||'').toString(), post:(cd[b][3]||'').toString(),
          status:'Conducted', date: fmtDateVal(cd[b][13]), conductDate: fmtDateVal(cd[b][13]),
          meetingType:(cd[b][8]||'').toString(), stakeholderName:(cd[b][9]||'').toString(),
          stakeholderPost:(cd[b][10]||'').toString(), stakeholderBlock:(cd[b][COL_CON_SKBLOCK-1]||'').toString(), outcome:(cd[b][COL_CON_OUTCOME-1]||'').toString(), purpose:(cd[b][11]||'').toString(),
          momUrl:(cd[b][17]||'').toString(), photoUrl:(cd[b][16]||'').toString(),
          govtMom:(cd[b][21]||'').toString(),
          priority:(cd[b][22]||'').toString(), flag:(cd[b][23]||'').toString(), nextAction:(cd[b][24]||'').toString(),
          escalate:(cd[b][25]||'').toString(), category:(cd[b][26]||'').toString(),
          colleagueName:(cd[b][18]||'').toString()
        });
      }
    }

    // 3. Postponed
    var xS = ss.getSheetByName(POSTPONED_SHEET);
    if (xS) {
      var xd = (sheetRows_(POSTPONED_SHEET) || []);
      for (var c = 1; c < xd.length; c++) {
        if (!xd[c][0]) continue;
        meetings.push({
          meetingId:(xd[c][0]||'').toString(), district:(xd[c][1]||'').toString(),
          block: blk(xd[c][3]), employeeName:(xd[c][2]||'').toString(), post:'',
          status:'Postponed', date: fmtDateVal(xd[c][8] || xd[c][7]), conductDate:'',
          meetingType:'', stakeholderName:(xd[c][4]||'').toString(),
          stakeholderPost:(xd[c][5]||'').toString(), purpose:(xd[c][6]||'').toString(),
          momUrl:'', photoUrl:'', colleagueName:''
        });
      }
    }

    // 4. Cancelled
    var zS = ss.getSheetByName(CANCELLED_SHEET);
    if (zS) {
      var zd = (sheetRows_(CANCELLED_SHEET) || []);
      for (var d = 1; d < zd.length; d++) {
        if (!zd[d][0]) continue;
        meetings.push({
          meetingId:(zd[d][0]||'').toString(), district:(zd[d][1]||'').toString(),
          block: blk(zd[d][4]), employeeName:(zd[d][2]||'').toString(), post:(zd[d][3]||'').toString(),
          status:'Cancelled', date: fmtDateVal(zd[d][5]), conductDate:'',
          meetingType:(zd[d][8]||'').toString(), stakeholderName:(zd[d][9]||'').toString(),
          stakeholderPost:(zd[d][10]||'').toString(), stakeholderBlock:(zd[d][COL_PLAN_SKBLOCK-1]||'').toString(), purpose:(zd[d][11]||'').toString(),
          momUrl:'', photoUrl:'', colleagueName:(zd[d][13]||'').toString()
        });
      }
    }

    // One meeting, one row. These four sheets are event logs, not one ledger:
    // a meeting postponed and then held is written in both, and until the
    // duplicate guards went in, a save that never reported back could leave two
    // conducted rows for the same meeting. Counting them flat made the portal's
    // district, team and stakeholder pages disagree with its own Overview,
    // which reads the current state of each meeting from the plan sheet.
    var STATUS_RANK = { 'Conducted':4, 'Cancelled':3, 'Postponed':2, 'Follow-up':1, 'Planned':1 };
    var byId = {}, idOrder = [];
    meetings.forEach(function(m) {
      var id = (m.meetingId || '').toString().trim();
      if (!id) return;
      if (!byId[id]) { byId[id] = m; idOrder.push(id); return; }
      // The furthest a meeting actually got wins. Between two rows at the same
      // stage the later one wins, which is the corrected note when an officer
      // wrote it twice.
      if ((STATUS_RANK[m.status] || 0) >= (STATUS_RANK[byId[id].status] || 0)) byId[id] = m;
    });
    meetings = idOrder.map(function(id) { return byId[id]; });

    var out = { success: true, meetings: meetings };
    cPut(cacheKey, out, C_TTL_LIVE);
    return out;
  } catch(err) { return { success:false, message: err.message }; }
}

function getDistrictReport(district) {
  try {
    if (!district) return { success: false, message: 'District required.' };
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var distLow  = district.trim().toLowerCase();

    // ── Totals from Plan Meetings ──────────────────────────────
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    var totals    = { total:0, conducted:0, planned:0, cancelled:0, postponed:0 };
    if (planSheet && planSheet.getLastRow() > 1) {
      var pd = planSheet.getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        if ((pd[i][1]||'').toString().trim().toLowerCase() !== distLow) continue;
        totals.total++;
        var st = (pd[i][13]||'Planned').toString().trim().toLowerCase();
        if      (st === 'conducted') totals.conducted++;
        else if (st === 'planned')   totals.planned++;
        else if (st === 'cancelled') totals.cancelled++;
        else if (st === 'postponed') totals.postponed++;
      }
    }

    // ── Conducted Meetings (full detail) ──────────────────────
    var cSheet   = ss.getSheetByName(CONDUCTED_SHEET);
    var conducted = [];
    var postMap   = {}; // post → { count, empMap: { name→count } }

    if (cSheet && cSheet.getLastRow() > 1) {
      var cd = cSheet.getDataRange().getValues();
      for (var ci = 1; ci < cd.length; ci++) {
        if ((cd[ci][1]||'').toString().trim().toLowerCase() !== distLow) continue;

        var emp  = (cd[ci][2]  || '').toString().trim();
        var post = (cd[ci][3]  || '').toString().trim();

        // post-wise map
        if (!postMap[post]) postMap[post] = { count:0, empMap:{} };
        postMap[post].count++;
        postMap[post].empMap[emp] = (postMap[post].empMap[emp] || 0) + 1;

        conducted.push({
          meetingId:       (cd[ci][0]  || '').toString(),
          employeeName:    emp,
          post:            post,
          conductDate:     fmtDateVal(cd[ci][13]),
          stakeholderName: (cd[ci][9]  || '').toString(),
          stakeholderPost: (cd[ci][10] || '').toString(),
          purpose:         (cd[ci][11] || '').toString(),
          meetingType:     (cd[ci][8]  || '').toString(),
          momUrl:          (cd[ci][17] || '').toString(),
          photoUrl:        (cd[ci][16] || '').toString(),
          colleagueName:   (cd[ci][18] || '').toString()
        });
      }
    }

    // sort conducted: newest first
    conducted.sort(function(a,b){ return b.conductDate.localeCompare(a.conductDate); });

    // build byPost array
    var byPost = [];
    for (var p in postMap) {
      var emps = [];
      for (var en in postMap[p].empMap) emps.push({ name:en, count:postMap[p].empMap[en] });
      emps.sort(function(a,b){ return b.count - a.count; });
      byPost.push({ post:p, count:postMap[p].count, employees:emps });
    }
    byPost.sort(function(a,b){ return b.count - a.count; });

    return {
      success:   true,
      district:  district,
      totals:    totals,
      byPost:    byPost,
      conducted: conducted
    };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  ALL CONDUCTED REPORTS - paginated list for dashboard
//  State user sees all districts; others see own district only
// ------------------------------------------------------------
function getAllReports(email) {
  try {
    var repKey = 'rep_' + email.trim().toLowerCase();
    var repHit = cGet(repKey);
    if (repHit) return repHit;

    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var userEmail = email.trim().toLowerCase();

    var cSheet = ss.getSheetByName(CONDUCTED_SHEET);
    if (!cSheet || cSheet.getLastRow() <= 1) return { success: true, reports: [] };

    var cd      = cSheet.getDataRange().getValues();
    var reports = [];
    for (var i = 1; i < cd.length; i++) {
      // Show only this user's own conducted meetings (col 4 = employee email)
      var rowEmail = (cd[i][4] || '').toString().trim().toLowerCase();
      if (rowEmail !== userEmail) continue;
      var dist = (cd[i][1] || '').toString().trim();
      reports.push({
        meetingId:       (cd[i][0]  || '').toString(),
        district:        dist,
        employeeName:    (cd[i][2]  || '').toString(),
        post:            (cd[i][3]  || '').toString(),
        originalDate:    fmtDateVal(cd[i][5]),
        meetingType:     (cd[i][8]  || '').toString(),
        stakeholderName: (cd[i][9]  || '').toString(),
        stakeholderPost: (cd[i][10] || '').toString(),
        purpose:         (cd[i][11] || '').toString(),
        conductDate:     fmtDateVal(cd[i][13]),
        keyPoints:       (cd[i][15] || '').toString(),
        momUrl:          (cd[i][17] || '').toString(),
        photoUrl:        (cd[i][16] || '').toString(),
        colleagueName:   (cd[i][18] || '').toString(),
        colleaguePost:   (cd[i][19] || '').toString()
      });
    }
    // Newest first
    reports.sort(function(a, b) { return b.conductDate.localeCompare(a.conductDate); });
    var repResult = { success: true, reports: reports };
    cPut(repKey, repResult, C_TTL_LIVE);
    return repResult;
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  EMPLOYEE LOOKUP
// ------------------------------------------------------------
// ------------------------------------------------------------
//  DASHBOARD STATS - cards & reports data
// ------------------------------------------------------------
function getDashboardStats(email, allDistricts, activeDistrict) {
  try {
    var statKey = 'stats_' + email.trim().toLowerCase() + '_' + (allDistricts ? '1' : '0') + '_' + (activeDistrict || '').toString().trim().toLowerCase();
    var statHit = cGet(statKey);
    if (statHit) return statHit;

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var emp = getEmployeeByEmail(email);
    var userRole     = emp ? (emp.role     || 'Field') : 'Field';
    // activeDistrict (from the switcher) overrides the user's primary when filtering
    var userDistrict = (activeDistrict || (emp ? emp.district : '') || '').toString();
    var isState      = allDistricts || (userRole === 'State');

    // ── Plan Meetings ──────────────────────────────────────────
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    var planData  = (planSheet && planSheet.getLastRow() > 1) ? (sheetRows_(MEETINGS_SHEET) || []) : [];

    var distMap = {};   // district → {total,conducted,planned,cancelled,postponed}
    var typeMap = {};
    var purpMap = {};
    var monthMap= {};   // "MMM YYYY" → conducted count

    for (var i = 1; i < planData.length; i++) {
      var row    = planData[i];
      var dist   = (row[1]  || '').toString().trim();
      var status = (row[13] || 'Planned').toString().trim().toLowerCase();
      var type   = (row[8]  || '').toString().trim();
      var purp   = (row[11] || '').toString().trim();
      var dateV  = row[5];

      if (!isState && dist.toUpperCase() !== userDistrict.toUpperCase()) continue;

      var dKey = dist.charAt(0).toUpperCase() + dist.slice(1).toLowerCase();
      if (!distMap[dKey]) distMap[dKey] = {total:0,conducted:0,planned:0,cancelled:0,postponed:0};
      distMap[dKey].total++;
      if      (status === 'conducted') distMap[dKey].conducted++;
      // A follow-up is a planned meeting that came out of an earlier one. It was
      // counted in the total but in none of the four buckets, so the cards on
      // the portal added up to eleven less than the total they sat under.
      else if (status === 'planned' || status === 'follow-up') distMap[dKey].planned++;
      else if (status === 'cancelled') distMap[dKey].cancelled++;
      else if (status === 'postponed') distMap[dKey].postponed++;

      if (type) typeMap[type] = (typeMap[type] || 0) + 1;
      if (purp) purpMap[purp] = (purpMap[purp] || 0) + 1;

      if (status === 'conducted' && dateV) {
        var d = new Date(dateV);
        if (!isNaN(d)) {
          var mk = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()] + ' ' + d.getFullYear();
          monthMap[mk] = (monthMap[mk] || 0) + 1;
        }
      }
    }

    // totals
    var totals = {total:0, conducted:0, planned:0, cancelled:0, postponed:0};
    var distArr = [];
    for (var d2 in distMap) {
      var dm = distMap[d2];
      totals.total     += dm.total;
      totals.conducted += dm.conducted;
      totals.planned   += dm.planned;
      totals.cancelled += dm.cancelled;
      totals.postponed += dm.postponed;
      distArr.push({name:d2, total:dm.total, conducted:dm.conducted, planned:dm.planned, cancelled:dm.cancelled, postponed:dm.postponed});
    }
    distArr.sort(function(a,b){ return a.name.localeCompare(b.name); });

    // type array
    var typeArr = [];
    for (var t in typeMap) typeArr.push({name:t, count:typeMap[t]});
    typeArr.sort(function(a,b){ return b.count - a.count; });

    // purpose array
    var purpArr = [];
    for (var p in purpMap) purpArr.push({name:p, count:purpMap[p]});
    purpArr.sort(function(a,b){ return b.count - a.count; });

    // month trend (last 6)
    var monthArr = [];
    for (var m in monthMap) monthArr.push({month:m, count:monthMap[m]});

    // ── Conducted Meetings - enriched stats + recent ──────────────
    var cSheet      = ss.getSheetByName(CONDUCTED_SHEET);
    var recent      = [];
    var empSet      = {};     // unique employee names
    var stkPostMap  = {};     // stakeholder post → count
    var dtfCount    = 0;      // Group Meeting conducted
    var momReady    = 0;      // meetings with MoM doc link

    if (cSheet && cSheet.getLastRow() > 1) {
      var cd = (sheetRows_(CONDUCTED_SHEET) || []);
      // Forward pass - collect stats
      for (var ci = 1; ci < cd.length; ci++) {
        var cr      = cd[ci];
        var cdist   = (cr[1]||'').toString().trim();
        if (!isState && cdist.toUpperCase() !== userDistrict.toUpperCase()) continue;
        var cEmp    = (cr[2] ||'').toString().trim();
        var cType   = (cr[8] ||'').toString().trim();
        var cStkP   = (cr[10]||'').toString().trim();
        var cMom    = (cr[17]||'').toString().trim();
        if (cEmp)  empSet[cEmp] = true;
        if (cStkP) stkPostMap[cStkP] = (stkPostMap[cStkP] || 0) + 1;
        if (cType.toLowerCase().indexOf('group') !== -1) dtfCount++;
        if (cMom)  momReady++;
      }
      // Reverse pass - collect recent 8
      for (var ri = cd.length - 1; ri >= 1 && recent.length < 8; ri--) {
        var rr    = cd[ri];
        if (!(rr[0] || '').toString().trim()) continue;   // trailing blank rows are not meetings
        var rdist = (rr[1]||'').toString().trim();
        if (!isState && rdist.toUpperCase() !== userDistrict.toUpperCase()) continue;
        recent.push({
          meetingId:       (rr[0] ||'').toString(),
          district:        (rr[1] ||'').toString(),
          employeeName:    (rr[2] ||'').toString(),
          post:            (rr[3] ||'').toString(),
          stakeholderName: (rr[9] ||'').toString(),
          stakeholderPost: (rr[10]||'').toString(),
          purpose:         (rr[11]||'').toString(),
          meetingType:     (rr[8] ||'').toString(),
          conductDate:     fmtDateVal(rr[13])
        });
      }
    }

    // Top stakeholder post
    var topStkPost = '-'; var topStkCount = 0;
    for (var sp in stkPostMap) {
      if (stkPostMap[sp] > topStkCount) { topStkCount = stkPostMap[sp]; topStkPost = sp; }
    }

    var statsResult = {
      success:             true,
      role:                userRole,
      district:            userDistrict,
      totals:              totals,
      districts:           distArr,
      byType:              typeArr,
      byPurpose:           purpArr,
      monthTrend:          monthArr,
      recentConducted:     recent,
      activeEmployees:     Object.keys(empSet).length,
      dtfSessions:         dtfCount,
      momReady:            momReady,
      topStakeholderPost:  topStkPost,
      topStakeholderCount: topStkCount
    };
    cPut(statKey, statsResult, C_TTL_LIVE);
    return statsResult;
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  INSERT SAMPLE DATA - run once from GAS editor
// ------------------------------------------------------------
function insertSampleData() {
  var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  var planSheet = ss.getSheetByName(MEETINGS_SHEET);
  var condSheet = ss.getSheetByName(CONDUCTED_SHEET);
  var cancSheet = ss.getSheetByName('Cancelled Meetings');

  if (!planSheet || !condSheet) { Logger.log('Sheets not found'); return; }

  var now = new Date();
  function ts(d) { return d.toLocaleString('en-IN'); }

  // ─── SAMPLE PLAN MEETINGS ─────────────────────────────────────
  // Cols: MtgID, District, EmpName, Post, Email, Date, Time, Duration,
  //       Type, StkName, StkPost, Purpose, Agenda, Status, ...SubmittedAt
  var planRows = [
    // HARDOI
    ['MTG-S-H01','Hardoi','Uday Raj','District Impact Specialist','uday.raj@educategirls.ngo','2026-04-10','10:00 AM','1 hr','One-on-One','Rajesh Kumar Verma','BSA','Review Meeting','Quarterly review of enrollment and retention data','Conducted','','','','','','01/04/2026, 9:00:00 am'],
    ['MTG-S-H02','Hardoi','Rahul Kumar','District Program Officer','rahul.kumar3@educategirls.ngo','2026-04-15','11:00 AM','45 min','One-on-One','Dr. Sunita Pathak','DIET Principal','Enrollment','Discuss strategies for out-of-school girl enrollment','Conducted','','','','Manvendra Mishra','District Program Officer','02/04/2026, 10:00:00 am'],
    ['MTG-S-H03','Hardoi','Manvendra Mishra','District Program Officer','manvendra.mishra@educategirls.ngo','2026-05-05','3:00 PM','30 min','One-on-One','Anil Tiwari','District Collector','Introductory Meeting','Initial introduction and EG program briefing','Conducted','','','','','','20/04/2026, 2:00:00 pm'],
    ['MTG-S-H04','Hardoi','Shivangi Verma','District Program Training Officer','shivangi.verma1@educategirls.ngo','2026-05-28','10:00 AM','2 hr','Group Meeting','Smt. Priya Agarwal','ABSA','DTF','Block-level training facilitation with ABSAs','Planned','','','','Rahul Kumar','District Program Officer','22/04/2026, 9:00:00 am'],
    ['MTG-S-H05','Hardoi','Uday Raj','District Impact Specialist','uday.raj@educategirls.ngo','2026-04-20','2:00 PM','1 hr','One-on-One','Vinod Sharma','CDO','Retention','Retention strategies for upper primary girls - Cancelled due to officer unavailability','Cancelled','','','Officer on leave','','','10/04/2026, 3:00:00 pm'],
    ['MTG-S-H06','Hardoi','Manvendra Mishra','District Program Officer','manvendra.mishra@educategirls.ngo','2026-05-18','11:30 AM','1 hr','Dept. Review','Ram Kishore','JD','MPR Submission','Submit monthly progress report and discuss targets','Conducted','','','','Uday Raj','District Impact Specialist','15/05/2026, 10:00:00 am'],

    // FATEHPUR
    ['MTG-S-F01','Fatehpur','Shubham Yadav','District Impact Specialist','shubham.yadav@educategirls.ngo','2026-04-12','10:30 AM','1 hr','One-on-One','Pramod Srivastava','BSA','MPR Submission','Monthly progress report submission and follow-up','Conducted','','','','Deepak Dixit','District Program Officer','05/04/2026, 9:00:00 am'],
    ['MTG-S-F02','Fatehpur','Deepak Dixit','District Program Officer','deepak.dixit@educategirls.ngo','2026-04-18','11:00 AM','1 hr','One-on-One','Dr. Kavita Mishra','DIET Principal','Review Meeting','Mid-year review of learning outcomes and teacher training','Conducted','','','','','','08/04/2026, 10:00:00 am'],
    ['MTG-S-F03','Fatehpur','Pushpendra Singh','District Program Training Officer','pushpendra.singh@educategirls.ngo','2026-04-25','9:00 AM','3 hr','Group Meeting','Anil Jaiswal','ABSA','DTF','Cluster-level training on learning assessment tools','Conducted','','','','Shubham Yadav','District Impact Specialist','12/04/2026, 8:00:00 am'],
    ['MTG-S-F04','Fatehpur','Ashish Rai','District Operational Assistant Lead','ashish.rai@educategirls.ngo','2026-05-30','10:00 AM','45 min','One-on-One','Smt. Rekha Devi','District Collector','Enrollment','Enrollment drive planning for new academic year','Planned','','','','','','18/04/2026, 9:00:00 am'],
    ['MTG-S-F05','Fatehpur','Shubham Yadav','District Impact Specialist','shubham.yadav@educategirls.ngo','2026-05-02','4:00 PM','30 min','One-on-One','Ajay Tripathi','CDO','Invitation','Invite CDO for EG annual review event','Cancelled','','','Event postponed','','','25/04/2026, 3:00:00 pm'],
    ['MTG-S-F06','Fatehpur','Deepak Dixit','District Program Officer','deepak.dixit@educategirls.ngo','2026-05-15','11:00 AM','1 hr','One-on-One','Suresh Patel','DC-Training','Learning','Discussion on training calendar and capacity building','Conducted','','','','','','05/05/2026, 10:00:00 am'],

    // GONDA
    ['MTG-S-G01','Gonda','Atul Pandey','District Impact Specialist','atul.pandey@educategirls.ngo','2026-04-08','10:00 AM','1 hr','One-on-One','Krishna Nand Yadav','BSA','Review Meeting','Review of EG program KPIs and district targets','Conducted','','','','Ashish Kumar Singh','District Program Officer','01/04/2026, 9:00:00 am'],
    ['MTG-S-G02','Gonda','Ashish Kumar Singh','District Program Officer','ashishkumar.singh1@educategirls.ngo','2026-04-22','11:30 AM','45 min','One-on-One','Dr. Reena Verma','DIET Principal','Enrollment','DIET-EG collaboration for out-of-school girl data','Conducted','','','','','','10/04/2026, 10:00:00 am'],
    ['MTG-S-G03','Gonda','Vedprakash Yadav','District Program Officer','vedprakash.Yadav@educategirls.ngo','2026-05-27','3:00 PM','1 hr','One-on-One','Suresh Prasad','District Collector','Introductory Meeting','EG program introduction and support request','Planned','','','','','','20/04/2026, 2:00:00 pm'],
    ['MTG-S-G04','Gonda','Arvind Kumar Yadav','Training Senior Specialist','arvind.yadav@educategirls.ngo','2026-05-01','9:00 AM','4 hr','Group Meeting','Ramesh Misra','ABSA','DTF','Training on NIPUN assessment and learning level improvement','Conducted','','','','Vedprakash Yadav','District Program Officer','22/04/2026, 8:00:00 am'],
    ['MTG-S-G05','Gonda','Atul Pandey','District Impact Specialist','atul.pandey@educategirls.ngo','2026-05-08','2:00 PM','30 min','One-on-One','Hari Om Mishra','JD','Retention','Discuss retention challenges at upper primary level','Cancelled','','','Meeting cancelled by stakeholder','','','30/04/2026, 1:00:00 pm'],
    ['MTG-S-G06','Gonda','Ashish Kumar Singh','District Program Officer','ashishkumar.singh1@educategirls.ngo','2026-05-20','10:30 AM','1 hr','Dept. Review','Om Prakash Tiwari','DC- Gender','MPR Submission','Gender data review and MPR submission','Conducted','','','','','','12/05/2026, 9:00:00 am'],

    // SITAPUR
    ['MTG-S-S01','Sitapur','Sumit Kumar','District Impact Specialist','sumit.kumar3@educategirls.ngo','2026-04-14','11:00 AM','1 hr','One-on-One','Awadhesh Yadav','BSA','MPR Submission','Submit district MPR and review block-wise progress','Conducted','','','','Vikrant Kumar','District Program Officer','06/04/2026, 10:00:00 am'],
    ['MTG-S-S02','Sitapur','Vikrant Kumar','District Program Officer','vikrant.kumar@educategirls.ngo','2026-05-06','10:00 AM','1 hr','One-on-One','Dr. Shashi Bala','DIET Principal','Review Meeting','Review of DIET training effectiveness on EG teachers','Conducted','','','','','','28/04/2026, 9:00:00 am'],
    ['MTG-S-S03','Sitapur','Mohd Shadab Ansari','District Program Officer','shadab.ansari@educategirls.ngo','2026-05-12','3:30 PM','45 min','One-on-One','Vinay Kumar Gupta','District Collector','Enrollment','Enrollment campaign planning for 2026-27','Conducted','','','','','','05/05/2026, 3:00:00 pm'],
    ['MTG-S-S04','Sitapur','Sashi Prakash','District Program Training Officer','shashi.patel@educategirls.ngo','2026-05-29','9:00 AM','3 hr','Group Meeting','Geeta Devi','ABSA','DTF','Pre-session training for ABSAs on new learning tools','Planned','','','','Sumit Kumar','District Impact Specialist','15/05/2026, 8:00:00 am'],
    ['MTG-S-S05','Sitapur','Sumit Kumar','District Impact Specialist','sumit.kumar3@educategirls.ngo','2026-04-28','4:00 PM','30 min','One-on-One','Ajeet Singh','CDO','Courtesy Meeting','Courtesy visit and program update to CDO','Cancelled','','','CDO transferred to another district','','','20/04/2026, 3:00:00 pm'],

    // BAHRAICH
    ['MTG-S-B01','Bahraich','Buddh Vilas','District Impact Specialist','buddh.vilas@educategirls.ngo','2026-04-16','10:00 AM','1 hr','One-on-One','Shyam Lal Gupta','BSA','Review Meeting','Annual review meeting - enrollment, retention, learning','Conducted','','','','Balwant Singh','District Operational Lead','08/04/2026, 9:00:00 am'],
    ['MTG-S-B02','Bahraich','Shyam Narayan Nath','District Program Officer','shyamnarayan.nath@educategirls.ngo','2026-05-26','11:00 AM','1 hr','One-on-One','Dr. Alka Jain','DIET Principal','Enrollment','Out-of-school girls data sharing with DIET','Planned','','','','','','18/04/2026, 10:00:00 am'],
    ['MTG-S-B03','Bahraich','Sanwara Vaishnav','District Program Training Officer','sanwara.vaishnav@educategirls.ngo','2026-05-04','9:00 AM','3 hr','Group Meeting','Deepak Kumar','ABSA','DTF','Training session on EG methodology and community mobilization','Conducted','','','','','','25/04/2026, 8:00:00 am'],
    ['MTG-S-B04','Bahraich','Balwant Singh','District Operational Lead','balwant.singh@educategirls.ngo','2026-04-30','2:00 PM','1 hr','One-on-One','Mohd. Azam Khan','District Collector','Introductory Meeting','Introductory meeting with new District Collector','Cancelled','','','New DC not yet joined charge','','','22/04/2026, 1:00:00 pm'],

    // SHAHJAHANPUR
    ['MTG-S-SJ1','Shahjahanpur','Indra Dev Tiwari','District Program Officer','indradev.tiwari@educategirls.ngo','2026-04-10','11:00 AM','1 hr','One-on-One','Surendra Bahadur Singh','BSA','MPR Submission','Monthly progress report submission - April','Conducted','','','','Ankit Kumar Dixit','District Program Officer','03/04/2026, 10:00:00 am'],
    ['MTG-S-SJ2','Shahjahanpur','Ankit Kumar Dixit','District Program Officer','ankit.dixit@educategirls.ngo','2026-04-23','10:30 AM','1 hr','One-on-One','Dr. Rama Kant','DIET Principal','Enrollment','Discuss enrollment targets and DIET support for EG program','Conducted','','','','','','14/04/2026, 9:00:00 am'],
    ['MTG-S-SJ3','Shahjahanpur','Chandra Mohan Sharma','District Program Training Officer','chandramohan.sharma@educategirls.ngo','2026-05-28','9:00 AM','4 hr','Group Meeting','Smt. Pushpa Singh','ABSA','DTF','District Training Facilitation - refresher session','Planned','','','','Indra Dev Tiwari','District Program Officer','20/04/2026, 8:00:00 am'],
    ['MTG-S-SJ4','Shahjahanpur','Vikas Kumar Tiwari','District Operational Assistant Lead','vikash.tiwari@educategirls.ngo','2026-05-07','3:00 PM','45 min','One-on-One','Ashutosh Verma','District Collector','Retention','Retention drive support request from district administration','Cancelled','','','Meeting rescheduled to next month','','','30/04/2026, 2:00:00 pm']
  ];

  // ─── SAMPLE CONDUCTED MEETINGS ────────────────────────────────
  // Cols: MtgID, Dist, EmpName, Post, Email, OrigDate, OrigTime, Duration,
  //       Type, StkName, StkPost, Purpose, Agenda, ConductDate, ConductTime,
  //       KeyPoints, PhotosFolder, MoMDoc, ColleagueName, ColleaguePost, ConductedAt
  var condRows = [
    ['MTG-S-H01','Hardoi','Uday Raj','District Impact Specialist','uday.raj@educategirls.ngo','2026-04-10','10:00 AM','1 hr','One-on-One','Rajesh Kumar Verma','BSA','Review Meeting','Quarterly review of enrollment and retention data','2026-04-10','10:45 AM','• Reviewed Q4 enrollment data - 87% target achieved\n• Discussed block-wise retention gaps in KPTG and SANDI\n• BSA agreed to issue circular for ABSA attendance in DTF sessions\n• Follow-up scheduled for May 15','','','','','10/04/2026, 11:50:00 am'],
    ['MTG-S-H02','Hardoi','Rahul Kumar','District Program Officer','rahul.kumar3@educategirls.ngo','2026-04-15','11:00 AM','45 min','One-on-One','Dr. Sunita Pathak','DIET Principal','Enrollment','Discuss strategies for out-of-school girl enrollment','2026-04-15','11:30 AM','• DIET will share block-wise OOS data by April 20\n• Principal agreed to conduct school-wise sensitization\n• EG to provide resource materials for DIET faculty\n• Joint visit to 3 schools planned for May','','','Manvendra Mishra','District Program Officer','15/04/2026, 12:05:00 pm'],
    ['MTG-S-H03','Hardoi','Manvendra Mishra','District Program Officer','manvendra.mishra@educategirls.ngo','2026-05-05','3:00 PM','30 min','One-on-One','Anil Tiwari','District Collector','Introductory Meeting','Initial introduction and EG program briefing','2026-05-05','3:15 PM','• DC appreciated EG program outcomes in Hardoi\n• Requested monthly update sheet for DC office\n• Discussed upcoming enrollment campaign - DC agreed to flag-off\n• Next meeting scheduled post elections','','','','','05/05/2026, 4:00:00 pm'],
    ['MTG-S-H06','Hardoi','Manvendra Mishra','District Program Officer','manvendra.mishra@educategirls.ngo','2026-05-18','11:30 AM','1 hr','Dept. Review','Ram Kishore','JD','MPR Submission','Submit monthly progress report and discuss targets','2026-05-18','12:00 PM','• April MPR submitted - 92% targets achieved\n• JD directed to improve learning outcomes data quality\n• EG team to share school-wise learning data by May 25\n• Monthly review mechanism to be strengthened','','','Uday Raj','District Impact Specialist','18/05/2026, 1:15:00 pm'],

    ['MTG-S-F01','Fatehpur','Shubham Yadav','District Impact Specialist','shubham.yadav@educategirls.ngo','2026-04-12','10:30 AM','1 hr','One-on-One','Pramod Srivastava','BSA','MPR Submission','Monthly progress report submission and follow-up','2026-04-12','11:00 AM','• March MPR accepted - strong enrollment numbers\n• BSA highlighted teacher absenteeism as key challenge\n• EG team to document school-wise attendance data\n• Follow-up on ABSA deployment in 3 blocks','','','Deepak Dixit','District Program Officer','12/04/2026, 12:00:00 pm'],
    ['MTG-S-F02','Fatehpur','Deepak Dixit','District Program Officer','deepak.dixit@educategirls.ngo','2026-04-18','11:00 AM','1 hr','One-on-One','Dr. Kavita Mishra','DIET Principal','Review Meeting','Mid-year review of learning outcomes and teacher training','2026-04-18','11:45 AM','• Learning assessment data reviewed - improvement in Grade 3-5\n• DIET agreed to include EG module in next BTC training\n• Principal to depute 2 DIET faculty for EG school visits\n• Collaborative workshop planned for June','','','','','18/04/2026, 12:30:00 pm'],
    ['MTG-S-F03','Fatehpur','Pushpendra Singh','District Program Training Officer','pushpendra.singh@educategirls.ngo','2026-04-25','9:00 AM','3 hr','Group Meeting','Anil Jaiswal','ABSA','DTF','Cluster-level training on learning assessment tools','2026-04-25','12:00 PM','• 18 ABSAs trained on NIPUN learning tools\n• Hands-on practice on assessment rubrics completed\n• All participants committed to weekly school monitoring\n• Next DTF scheduled for June','','','Shubham Yadav','District Impact Specialist','25/04/2026, 12:30:00 pm'],
    ['MTG-S-F06','Fatehpur','Deepak Dixit','District Program Officer','deepak.dixit@educategirls.ngo','2026-05-15','11:00 AM','1 hr','One-on-One','Suresh Patel','DC-Training','Learning','Discussion on training calendar and capacity building','2026-05-15','11:50 AM','• Training calendar for 2026-27 shared with DC Training\n• Three EG-specific modules approved for inclusion\n• Resource persons list to be shared by May 20\n• Joint review after first training cycle','','','','','15/05/2026, 12:20:00 pm'],

    ['MTG-S-G01','Gonda','Atul Pandey','District Impact Specialist','atul.pandey@educategirls.ngo','2026-04-08','10:00 AM','1 hr','One-on-One','Krishna Nand Yadav','BSA','Review Meeting','Review of EG program KPIs and district targets','2026-04-08','10:50 AM','• KPI review: enrollment 91%, retention 84%, learning 78%\n• BSA committed to resolve ABSA vacancy in 2 blocks\n• EG team to provide block-wise dashboard monthly\n• Next review in May with data from all 12 blocks','','','Ashish Kumar Singh','District Program Officer','08/04/2026, 11:55:00 am'],
    ['MTG-S-G02','Gonda','Ashish Kumar Singh','District Program Officer','ashishkumar.singh1@educategirls.ngo','2026-04-22','11:30 AM','45 min','One-on-One','Dr. Reena Verma','DIET Principal','Enrollment','DIET-EG collaboration for out-of-school girl data','2026-04-22','12:00 PM','• DIET OOS data for 8 blocks shared\n• Joint verification exercise to be conducted in May\n• EG and DIET to co-develop household survey tool\n• DIET faculty to support community mobilization','','','','','22/04/2026, 12:30:00 pm'],
    ['MTG-S-G04','Gonda','Arvind Kumar Yadav','Training Senior Specialist','arvind.yadav@educategirls.ngo','2026-05-01','9:00 AM','4 hr','Group Meeting','Ramesh Misra','ABSA','DTF','Training on NIPUN assessment and learning level improvement','2026-05-01','1:00 PM','• 22 ABSAs trained across 4 blocks\n• Practical sessions on NIPUN tools completed\n• Block-wise action plans prepared by each ABSA\n• Follow-up classroom observation scheduled for June','','','Vedprakash Yadav','District Program Officer','01/05/2026, 1:30:00 pm'],
    ['MTG-S-G06','Gonda','Ashish Kumar Singh','District Program Officer','ashishkumar.singh1@educategirls.ngo','2026-05-20','10:30 AM','1 hr','Dept. Review','Om Prakash Tiwari','DC- Gender','MPR Submission','Gender data review and MPR submission','2026-05-20','11:20 AM','• Gender-disaggregated data reviewed for April\n• Drop-out rate among girls in Class 6-8 flagged as concern\n• DC Gender to raise in DISE data meeting\n• EG to provide school-wise risk analysis','','','','','20/05/2026, 11:45:00 am'],

    ['MTG-S-S01','Sitapur','Sumit Kumar','District Impact Specialist','sumit.kumar3@educategirls.ngo','2026-04-14','11:00 AM','1 hr','One-on-One','Awadhesh Yadav','BSA','MPR Submission','Submit district MPR and review block-wise progress','2026-04-14','11:45 AM','• March MPR submitted - 89% enrollment, 82% retention\n• BSA requested EG data in Excel format for compilation\n• Block-wise performance matrix to be shared weekly\n• ABSA meeting to be organized in May','','','Vikrant Kumar','District Program Officer','14/04/2026, 12:30:00 pm'],
    ['MTG-S-S02','Sitapur','Vikrant Kumar','District Program Officer','vikrant.kumar@educategirls.ngo','2026-05-06','10:00 AM','1 hr','One-on-One','Dr. Shashi Bala','DIET Principal','Review Meeting','Review of DIET training effectiveness on EG teachers','2026-05-06','10:55 AM','• DIET training impact study data shared\n• Significant improvement in teacher facilitation skills noted\n• 3 best-practice schools identified for documentation\n• Exposure visit for DIET faculty to EG schools planned','','','','','06/05/2026, 11:20:00 am'],
    ['MTG-S-S03','Sitapur','Mohd Shadab Ansari','District Program Officer','shadab.ansari@educategirls.ngo','2026-05-12','3:30 PM','45 min','One-on-One','Vinay Kumar Gupta','District Collector','Enrollment','Enrollment campaign planning for 2026-27','2026-05-12','4:05 PM','• DC approved EG-led enrollment campaign for June\n• Gram Pradhan mobilization to be done via BDO circulars\n• EG team to prepare campaign material by May 20\n• DC office to share support letter for schools','','','','','12/05/2026, 4:30:00 pm'],

    ['MTG-S-B01','Bahraich','Buddh Vilas','District Impact Specialist','buddh.vilas@educategirls.ngo','2026-04-16','10:00 AM','1 hr','One-on-One','Shyam Lal Gupta','BSA','Review Meeting','Annual review meeting - enrollment, retention, learning','2026-04-16','10:50 AM','• Annual data reviewed - targets met in 7 of 9 blocks\n• Learning outcomes below benchmark in 2 blocks - plan needed\n• BSA agreed to depute resource persons for those blocks\n• EG to submit action plan by April 25','','','Balwant Singh','District Operational Lead','16/04/2026, 11:50:00 am'],
    ['MTG-S-B03','Bahraich','Sanwara Vaishnav','District Program Training Officer','sanwara.vaishnav@educategirls.ngo','2026-05-04','9:00 AM','3 hr','Group Meeting','Deepak Kumar','ABSA','DTF','Training session on EG methodology and community mobilization','2026-05-04','12:00 PM','• 16 ABSAs trained on EG community mobilization approach\n• Role-play exercises on parent engagement conducted\n• Commitments taken for monthly school-community meets\n• Refresher session scheduled for July','','','','','04/05/2026, 12:30:00 pm'],

    ['MTG-S-SJ1','Shahjahanpur','Indra Dev Tiwari','District Program Officer','indradev.tiwari@educategirls.ngo','2026-04-10','11:00 AM','1 hr','One-on-One','Surendra Bahadur Singh','BSA','MPR Submission','Monthly progress report submission - April','2026-04-10','11:50 AM','• March MPR submitted - 85% overall target achievement\n• BSA appreciated improvement in retention data quality\n• New data collection format to be piloted in 2 blocks\n• Follow-up meeting for April data in first week of May','','','Ankit Kumar Dixit','District Program Officer','10/04/2026, 12:00:00 pm'],
    ['MTG-S-SJ2','Shahjahanpur','Ankit Kumar Dixit','District Program Officer','ankit.dixit@educategirls.ngo','2026-04-23','10:30 AM','1 hr','One-on-One','Dr. Rama Kant','DIET Principal','Enrollment','Discuss enrollment targets and DIET support for EG program','2026-04-23','11:20 AM','• Enrollment targets for 2026-27 discussed and agreed\n• DIET to provide training support for 45 EG schools\n• Resource material library to be set up at DIET\n• Joint visit to 5 EG schools planned for May','','','','','23/04/2026, 11:45:00 am']
  ];

  // ─── SAMPLE CANCELLED MEETINGS ────────────────────────────────
  // Cols: MtgID, Dist, EmpName, Post, Email, Date, Time, Duration,
  //       Type, StkName, StkPost, Purpose, Agenda, ColleagueName, ColleaguePost, Reason, CancelledAt
  var cancRows = [
    ['MTG-S-H05','Hardoi','Uday Raj','District Impact Specialist','uday.raj@educategirls.ngo','2026-04-20','2:00 PM','1 hr','One-on-One','Vinod Sharma','CDO','Retention','Retention strategies for upper primary girls','','','Officer on leave - rescheduled','20/04/2026, 2:30:00 pm'],
    ['MTG-S-F05','Fatehpur','Shubham Yadav','District Impact Specialist','shubham.yadav@educategirls.ngo','2026-05-02','4:00 PM','30 min','One-on-One','Ajay Tripathi','CDO','Invitation','Invite CDO for EG annual review event','','','Event postponed by organizers','02/05/2026, 4:15:00 pm'],
    ['MTG-S-G05','Gonda','Atul Pandey','District Impact Specialist','atul.pandey@educategirls.ngo','2026-05-08','2:00 PM','30 min','One-on-One','Hari Om Mishra','JD','Retention','Discuss retention challenges at upper primary level','','','Meeting cancelled by stakeholder - national duty','08/05/2026, 2:20:00 pm'],
    ['MTG-S-S05','Sitapur','Sumit Kumar','District Impact Specialist','sumit.kumar3@educategirls.ngo','2026-04-28','4:00 PM','30 min','One-on-One','Ajeet Singh','CDO','Courtesy Meeting','Courtesy visit and program update to CDO','','','CDO transferred to another district','28/04/2026, 4:10:00 pm'],
    ['MTG-S-B04','Bahraich','Balwant Singh','District Operational Lead','balwant.singh@educategirls.ngo','2026-04-30','2:00 PM','1 hr','One-on-One','Mohd. Azam Khan','District Collector','Introductory Meeting','Introductory meeting with new District Collector','','','New DC not yet joined charge','30/04/2026, 2:15:00 pm'],
    ['MTG-S-SJ4','Shahjahanpur','Vikas Kumar Tiwari','District Operational Assistant Lead','vikash.tiwari@educategirls.ngo','2026-05-07','3:00 PM','45 min','One-on-One','Ashutosh Verma','District Collector','Retention','Retention drive support request from district administration','','','Meeting rescheduled to next month - DC tour','07/05/2026, 3:20:00 pm']
  ];

  // ─── INSERT ROWS ──────────────────────────────────────────────
  planRows.forEach(function(r) { planSheet.appendRow(r); });
  condRows.forEach(function(r) { condSheet.appendRow(r); });

  if (cancSheet) {
    var cancHeader = cancSheet.getLastRow();
    if (cancHeader < 1) {
      cancSheet.appendRow(['Meeting ID','District','Employee Name','Post','Email','Meeting Date','Meeting Time','Duration','Meeting Type','Stakeholder Name','Stakeholder Post','Meeting Purpose','Meeting Agenda','Colleague Name','Colleague Post','Reason','Cancelled At']);
      cancSheet.getRange(1,1,1,17).setBackground('#7F1D1D').setFontColor('#fff').setFontWeight('bold');
      cancSheet.setFrozenRows(1);
    }
    cancRows.forEach(function(r) { cancSheet.appendRow(r); });
  }

  Logger.log('✅ Sample data inserted: ' + planRows.length + ' planned, ' + condRows.length + ' conducted, ' + cancRows.length + ' cancelled.');
}

// ── Employee master, with a copy that does not depend on Sheets ──────────
// Signing in has to know who you are, which meant opening the employee sheet.
// On the evening of 17 Sep 2026 the Sheets service began timing out on this
// document, every OTP request hung for its full six minutes, and nobody could
// sign in at all. A trimmed copy of the master now lives in Script Properties,
// which has nothing to do with Sheets, and the sign-in path reads that. The
// Sheet is opened only to refresh the copy: hourly from calendarJob, and on
// demand from EMP_refreshMirror(). Run that after adding or changing anyone.
var EMP_MIRROR_N     = 'EMP_MIRROR_N';
var EMP_MIRROR_AT    = 'EMP_MIRROR_AT';
var EMP_MIRROR_CHUNK = 8000;            // Script Properties hold 9KB per value

function empMirrorWrite_(map) {
  try {
    var s = JSON.stringify(map), out = {}, parts = 0;
    for (var i = 0; i < s.length; i += EMP_MIRROR_CHUNK) {
      out['EMP_MIRROR_' + parts] = s.slice(i, i + EMP_MIRROR_CHUNK);
      parts++;
    }
    out[EMP_MIRROR_N]  = String(parts);
    out[EMP_MIRROR_AT] = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperties(out);
    return parts;
  } catch (e) { return 0; }
}

function empMirrorRead_() {
  try {
    var all = PropertiesService.getScriptProperties().getProperties();
    var n = parseInt(all[EMP_MIRROR_N], 10);
    if (!n) return null;
    var s = '';
    for (var i = 0; i < n; i++) {
      if (all['EMP_MIRROR_' + i] == null) return null;   // a half written copy is no copy
      s += all['EMP_MIRROR_' + i];
    }
    return JSON.parse(s);
  } catch (e) { return null; }
}

// The one place that opens the employee sheet. Returns null when Sheets will not
// answer, so callers can fall back instead of treating an outage as "no staff".
function empReadFromSheet_() {
  if (sheetBreakerTripped_()) return null;
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EMPLOYEE_SHEET);
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues(), map = {};
    for (var i = 1; i < data.length; i++) {
      var em = data[i][4] ? data[i][4].toString().trim().toLowerCase() : '';
      if (!em) continue;
      var primaryDist = (data[i][0] || '').toString().trim();
      // Col H = "Additional Districts", extra charge, comma or semicolon separated
      var districts = [primaryDist];
      (data[i][7] || '').toString().split(/[,;]/).forEach(function(x) {
        var d = x.toString().trim();
        if (d && districts.map(function(z){ return z.toLowerCase(); }).indexOf(d.toLowerCase()) === -1) districts.push(d);
      });
      map[em] = {
        district:    primaryDist,
        districts:   districts,
        block:       (data[i][1] || '').toString().trim(),
        name:        (data[i][2] || '').toString().trim(),
        designation: (data[i][3] || '').toString().trim(),
        email:       (data[i][4] || '').toString().trim(),
        role:        normalizeRole_(data[i][5]),
        zone:        (data[i][6] || '').toString().trim()
      };
    }
    return map;
  } catch (e) { sheetBreakerTrip_(); return null; }
}

// Editor helper: take a fresh copy of the employee master. One attempt only,
// because a Sheets timeout eats the whole six minutes an execution is allowed.
// If it says the service will not answer, just run it again in a minute.
function EMP_refreshMirror() {
  var map = empReadFromSheet_();
  if (!map) {
    Logger.log('Could not read the employee sheet - the Sheets service is not answering. Run this again in a minute.');
    return 0;
  }
  var people = 0;
  for (var k in map) people++;
  var chunks = empMirrorWrite_(map);
  Logger.log('Employee copy refreshed: ' + people + ' people, ' + chunks + ' chunk(s). Sign-in no longer needs the Sheet.');
  return people;
}

// Editor helper: write one person into the sign-in copy WITHOUT opening the
// Sheet. For the case this was written in: the Sheets service is refusing the
// document, so EMP_refreshMirror() cannot run, and nobody can sign in because
// sign-in has to look the person up. This puts one known person in by hand so
// they are not locked out. It adds to the copy, it does not replace it, and the
// hourly refresh overwrites the lot with the real sheet as soon as Sheets is
// answering again. Nothing here grants anything the sheet does not already say.
function EMP_seedOne(email, name, role, district, designation, zone) {
  email = (email || '').toString().trim().toLowerCase();
  if (!email) { Logger.log('Pass an email.'); return 0; }
  var map = empMirrorRead_() || {};
  map[email] = {
    district:    (district || '').toString().trim(),
    districts:   (district || '').toString().trim() ? [(district || '').toString().trim()] : [],
    block:       '',
    name:        (name || email.split('@')[0]).toString().trim(),
    designation: (designation || role || '').toString().trim(),
    email:       email,
    role:        normalizeRole_(role || 'State'),
    zone:        (zone || '').toString().trim()
  };
  var chunks = empMirrorWrite_(map);
  var n = 0; for (var k in map) n++;
  Logger.log('Seeded ' + email + String.fromCharCode(10) +
             JSON.stringify(map[email]) + String.fromCharCode(10) +
             'Copy now holds ' + n + ' person(s) in ' + chunks + ' chunk(s). Sign-in will use this.');
  return n;
}

// The one this was needed for. Run EMP_seedMe() and sign in.
function EMP_seedMe() {
  return EMP_seedOne('alok.mohan@educategirls.ngo', 'Alok Mohan', 'State', '', 'State Lead', '');
}

function EMP_mirrorStatus() {
  var all = PropertiesService.getScriptProperties().getProperties();
  var m = empMirrorRead_();
  var n = 0; for (var k in (m || {})) n++;
  Logger.log('Copy taken at: ' + (all[EMP_MIRROR_AT] || 'never') + String.fromCharCode(10) + 'People in the copy: ' + n);
  return n;
}

function getEmployeeByEmail(email) {
  email = (email || '').toString().trim().toLowerCase();
  if (!email) return null;
  var key = 'emp_' + email;
  var hit = cGet(key);
  if (hit !== null) return hit;   // null-employee cached as JSON null

  // The copy first. A Sheets outage must not be able to stop anyone signing in.
  var mirror = empMirrorRead_();
  if (mirror && Object.prototype.hasOwnProperty.call(mirror, email)) {
    cPut(key, mirror[email], C_TTL_EMP);
    return mirror[email];
  }

  // Not in the copy, so this is someone new: the Sheet is the only answer.
  var map = empReadFromSheet_();
  if (!map) return null;
  empMirrorWrite_(map);
  var res = map[email] || null;
  cPut(key, res, C_TTL_EMP);
  return res;
}

function bulkUpdateEmployeeDB(rows) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
    if (!sheet) return { success: false, message: 'Employee_DB sheet not found' };

    // Clear existing data rows (keep header at row 1)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }

    // Write new data
    if (rows && rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }

    return { success: true, count: rows ? rows.length : 0, message: 'Employee_DB updated successfully' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  PEEK SOURCE SHEET - returns sheet names + first 3 data rows
//  action=peekSourceSheet&sourceId=SPREADSHEET_ID&sheetIndex=0
// ------------------------------------------------------------
function peekSourceSheet(sourceId, sheetIndex) {
  try {
    var src    = SpreadsheetApp.openById(sourceId);
    var sheets = src.getSheets();
    var info   = sheets.map(function(s){
      return { name: s.getName(), gid: s.getSheetId(), rows: s.getLastRow() - 1 };
    });
    var ws = findSheet_(sheets, sheetIndex);
    var sample = ws.getRange(1, 1, Math.min(4, ws.getLastRow()), ws.getLastColumn()).getValues();
    return { success: true, sheets: info, selectedSheet: ws.getName(), selectedGid: ws.getSheetId(), sample: sample };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// Helper: find sheet by gid (if sheetIndex > 100) or by array index
function findSheet_(sheets, sheetIndex) {
  var n = parseInt(sheetIndex) || 0;
  if (n > 100) {
    // treat as gid
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === n) return sheets[i];
    }
  }
  return sheets[n] || sheets[0];
}

// ------------------------------------------------------------
//  IMPORT FROM SOURCE SHEET → Employee_DB
//  action=importFromSource&sourceId=SPREADSHEET_ID&sheetIndex=0
//  Reads first tab of source spreadsheet and overwrites Employee_DB
// ------------------------------------------------------------
function importFromSource(sourceId, sheetIndex) {
  try {
    // 1. Read source
    var src    = SpreadsheetApp.openById(sourceId);
    var sheets = src.getSheets();
    var srcWs  = findSheet_(sheets, sheetIndex);
    if (!srcWs) return { success: false, message: 'Source sheet not found for: ' + sheetIndex };

    var srcData = srcWs.getDataRange().getValues();
    if (srcData.length <= 1) return { success: false, message: 'Source sheet has no data rows' };

    // Data rows (skip header row 0)
    // Filter: skip vacant/empty rows
    // Employee_DB format: col0=District, col1=Block, col2=Name, col3=Designation, col4=Email, col5=Role
    var dataRows = srcData.slice(1).filter(function(row) {
      var name  = (row[2] || '').toString().trim().toLowerCase();
      var email = (row[4] || '').toString().trim().toLowerCase();
      // Skip if row is entirely empty
      var allEmpty = row.every(function(c){ return !c || c.toString().trim() === ''; });
      if (allEmpty) return false;
      // Skip if name is vacant/empty
      if (!name || name === 'vacant' || name.indexOf('vacant') === 0) return false;
      // Skip if email is vacant/missing/error
      if (!email || email === 'vacant' || email === '#n/a' || email === 'n/a') return false;
      return true;
    });

    // 2. Write to target Employee_DB
    var tgt   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tgtWs = tgt.getSheetByName(EMPLOYEE_SHEET);
    if (!tgtWs) return { success: false, message: 'Employee_DB sheet not found in target' };

    var lastRow = tgtWs.getLastRow();
    if (lastRow > 1) {
      tgtWs.deleteRows(2, lastRow - 1);
    }

    var colCount = srcData[0].length;
    tgtWs.getRange(2, 1, dataRows.length, colCount).setValues(dataRows);

    return {
      success: true,
      sourceSheet: srcWs.getName(),
      rowsImported: dataRows.length,
      columns: colCount,
      message: 'Import complete - vacant rows skipped'
    };
  } catch(err) {
    return { success: false, message: err.message };
  }
}


// ------------------------------------------------------------
//  CLEAR MY CACHE - call after role/data change to force refresh
//  action=clearMyCache&email=user@educategirls.ngo
// ------------------------------------------------------------
function clearMyCache(email) {
  try {
    var e = email.trim().toLowerCase();
    invalidateUser(e);
    // Also clear district employee lists
    CacheService.getScriptCache().removeAll([
      'allEmp', 'emp_' + e,
      'distEmp_bahraich','distEmp_fatehpur','distEmp_gonda',
      'distEmp_hardoi','distEmp_shahjahanpur','distEmp_sitapur','distEmp_prayagraj'
    ]);
    return { success: true, message: 'Cache cleared for ' + email };
  } catch(err) {
    return { success: false, message: err.message };
  }
}