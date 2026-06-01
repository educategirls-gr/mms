// ============================================================
//  EG GR Reporting System — Main Backend
//  Google Apps Script | Bound to Employee_DB Spreadsheet
// ============================================================

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
//  CACHE HELPERS  (GAS CacheService — script-level, 6 hr max)
// ============================================================
var C_TTL_EMP    = 1800;  // 30 min — employee data (rarely changes)
var C_TTL_LIVE   = 90;    // 90 sec — dashboard stats & reports
var C_TTL_DROP   = 900;   // 15 min — dropdown / colleague lists

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
function invalidateUser(email) {
  cDel('emp_' + email,
       'stats_' + email + '_0', 'stats_' + email + '_1',
       'rep_' + email,
       'mymt_' + email, 'allmymt_' + email,
       'mymtg_' + email,
       'stateMtg_all');
}

// ------------------------------------------------------------
//  DEMO / PRESENTATION HELPER — Run from GAS editor
//  Sheet mein role change karne ke baad ye run karo
//  Turant cache clear hoga — logout/login ke baad naya role dikhega
// ------------------------------------------------------------
function clearCacheForDemo() {
  var email = 'alok.mohan@educategirls.ngo'; // ← apna email yahan rakho
  invalidateUser(email);
  CacheService.getScriptCache().remove('distMtg_sitapur');
  CacheService.getScriptCache().remove('stateMtg_all');
  Logger.log('✅ Cache cleared for: ' + email + ' — ab logout karke login karo');
}

// ------------------------------------------------------------
//  DEMO ROLE SWITCHERS — Run ONE of these from the GAS editor,
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
//  NORMALIZE DISTRICTS — Run ONCE from GAS editor
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
//  DIAGNOSTIC — district column audit across all meeting sheets
//  Returns a summary of district values so we can see why a
//  district filter (e.g. SITAPUR) shows fewer meetings than expected.
// ------------------------------------------------------------
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
//  AUTHORIZE ALL SERVICES — Run this once from GAS editor
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
//  API HANDLER — called from GitHub Pages frontend via fetch()
// ------------------------------------------------------------
function doPost(e) {
  return apiResponse(e, 'POST');
}

// Admin allowlist — only these emails can call destructive/import actions
var ADMIN_EMAILS = ['gr@educategirls.ngo', 'alok.mohan@educategirls.ngo'];
function isAdmin_(email) {
  email = (email || '').toString().trim().toLowerCase();
  for (var i = 0; i < ADMIN_EMAILS.length; i++) {
    if (ADMIN_EMAILS[i].toLowerCase() === email) return true;
  }
  return false;
}

function apiResponse(e, method) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var result;
  try {
    var body = {};
    if (method === 'POST' && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch(pe) { body = {}; }
    }
    var token  = (e && e.parameter && e.parameter.token) ? e.parameter.token : '';
    var PUBLIC = { sendOTP: 1, verifyOTP: 1 };
    var ADMIN  = { bulkUpdateEmployeeDB: 1, importFromSource: 1, peekSourceSheet: 1 };

    if (PUBLIC[action]) {
      // ── No auth required ──────────────────────────────────────
      if (action === 'sendOTP') result = sendOTP(e.parameter.email || '');
      else                      result = verifyOTP(e.parameter.email || '', e.parameter.otp || '');
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

        if      (action === 'getDropdownData')      result = getDropdownData(session.email);
        else if (action === 'getMyMeetings')        result = getMyMeetings(session.email);
        else if (action === 'getAllMyMeetings')     result = getAllMyMeetings(session.email);
        else if (action === 'getDistrictEmployees') result = getDistrictEmployees(session.district, session.email);
        else if (action === 'getAllEmployees')      result = getAllEmployees(session.email);
        else if (action === 'getDistrictAllMeetings') {
          // District role locked to own district; State may query any
          var dist = (role === 'State') ? (e.parameter.district || session.district) : session.district;
          result = getDistrictAllMeetings(dist);
        }
        else if (action === 'getStateAllMeetings')  result = (role === 'State')
                                                       ? getStateAllMeetings()
                                                       : { success: false, message: 'FORBIDDEN' };
        else if (action === 'getDashboardStats')    result = getDashboardStats(session.email, e.parameter.all === '1');
        else if (action === 'getDistrictReport') {
          var dr = (role === 'State') ? (e.parameter.district || session.district) : session.district;
          result = getDistrictReport(dr);
        }
        else if (action === 'getAllReports')        result = getAllReports(session.email);
        else if (action === 'deleteMeeting')        result = deleteMeeting(e.parameter.meetingId || '', session.email);
        else if (action === 'saveMeeting') {
          // Stamp identity from session — fixes attribution + blank district
          body.email = session.email; body.employeeName = session.name;
          body.district = session.district; body.designation = session.designation; body.block = session.block;
          result = saveMeeting(body);
        }
        else if (action === 'conductMeeting')   { body.email = session.email; result = conductMeeting(body); }
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
//  TEST FUNCTION — Run this once from GAS editor to authorize MailApp
// ------------------------------------------------------------
function authorizeMailApp() {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'EG MMS — MailApp Authorization Successful',
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
//  OTP — SEND
// ------------------------------------------------------------
function sendOTP(email) {
  email = email.trim().toLowerCase();

  // Only allow office domain
  var domain = email.split('@')[1] || '';
  if (domain !== ALLOWED_DOMAIN) {
    return { success: false, message: 'Only @' + ALLOWED_DOMAIN + ' email addresses are allowed.' };
  }

  var employee = getEmployeeByEmail(email);
  if (!employee) {
    return { success: false, message: 'Your email is not registered in the system. Please contact HR.' };
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
//  OTP — VERIFY
// ------------------------------------------------------------
function verifyOTP(email, otp) {
  email = email.trim().toLowerCase();
  otp   = otp.trim();

  var cache     = CacheService.getScriptCache();
  var storedOTP = cache.get('OTP_' + email);

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
    block:       emp.block,
    designation: emp.designation,
    role:        emp.role,
    loginTime:   new Date().toISOString()
  });
  cache.put('SESSION_' + token, session, 3600); // 1 hour

  return {
    success:     true,
    token:       token,
    role:        emp.role,
    name:        emp.name,
    district:    emp.district,
    block:       emp.block,
    designation: emp.designation,
    email:       emp.email
  };
}

// ------------------------------------------------------------
//  SESSION — GET
// ------------------------------------------------------------
function getSession(token) {
  var data = CacheService.getScriptCache().get('SESSION_' + token);
  if (!data) return null;
  return JSON.parse(data);
}

// ------------------------------------------------------------
//  ACCESS CHECK — re-verify employee still active in sheet
//  Returns emp object if active, null if removed/not found
// ------------------------------------------------------------
function checkAccess(email) {
  return getEmployeeByEmail(email.trim().toLowerCase());
}

// ------------------------------------------------------------
//  DROPDOWN DATA — Stakeholder Type (hardcoded) + Meeting Purpose (sheet)
// ------------------------------------------------------------
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

  var purposes = [];
  var cache = CacheService.getScriptCache();
  var cached = cache.get('EG_PURPOSES');
  if (cached) {
    purposes = JSON.parse(cached);
  } else {
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ws2 = ss.getSheetByName('Meeting Purpose');
    if (ws2) {
      var d2 = ws2.getDataRange().getValues();
      for (var i = 1; i < d2.length; i++) {
        if (d2[i][0]) purposes.push(d2[i][0].toString().trim());
      }
    }
    cache.put('EG_PURPOSES', JSON.stringify(purposes), 600);
  }

  return { stakeholders: stakeholders, purposes: purposes };
}

// ------------------------------------------------------------
//  MEETINGS — SAVE
// ------------------------------------------------------------
function saveMeeting(data) {
  try {
    // Re-verify employee is still active before saving
    if (!getEmployeeByEmail((data.email || '').trim().toLowerCase())) {
      return { success: false, message: 'ACCESS_REVOKED' };
    }

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MEETINGS_SHEET);
    if (!sheet) return { success: false, message: 'Meetings sheet not found.' };

    var now   = new Date();
    var mtgId = 'MTG-' + now.getFullYear() +
                ('0'+(now.getMonth()+1)).slice(-2) +
                ('0'+now.getDate()).slice(-2) + '-' +
                ('0'+now.getHours()).slice(-2) + ('0'+now.getMinutes()).slice(-2) + ('0'+now.getSeconds()).slice(-2);

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
      data.parentMeetingId || ''   // U  Parent Meeting ID (for follow-ups)
    ];

    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), 7).setNumberFormat('@'); // keep Meeting Time as text

    // ── Colleague email notification ──────────────────────────
    if (data.colleagueName && data.colleagueName.trim()) {
      try { sendColleagueNotification(data, mtgId); } catch(mailErr) { /* don't fail save if mail fails */ }
    }

    invalidateUser((data.email || '').trim().toLowerCase());
    return { success: true, meetingId: mtgId };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  MEETINGS — GET (for logged-in employee)
// ------------------------------------------------------------
function getMyMeetings(email) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MEETINGS_SHEET);
    if (!sheet) return [];

    var sheetData = sheet.getDataRange().getValues();
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
          purpose:      (sheetData[i][11] || '').toString(),  // L
          agenda:       (sheetData[i][12] || '').toString(),  // M
          status:       (sheetData[i][13] || '').toString(),  // N
          startTime:    (sheetData[i][14] || '').toString(),  // O
          endTime:      (sheetData[i][15] || '').toString(),  // P
          reason:       (sheetData[i][16] || '').toString(),  // Q
          colleagueName:(sheetData[i][17] || '').toString(),  // R
          colleaguePost:(sheetData[i][18] || '').toString(),  // S
          parentMeetingId: (sheetData[i][20] || '').toString() // U
        });
      }
    }
    return meetings;
  } catch (err) {
    return [];
  }
}

// ------------------------------------------------------------
//  DISTRICT EMPLOYEES — for colleague picker
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
//  ALL EMPLOYEES — for colleague picker (no district filter)
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
//  MEETING — UPDATE STATUS (from Manage Meetings)
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
//  TIME HELPER — Sheets stores time as Dec-30-1899 Date objects.
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
  return s; // already formatted or unknown — return as-is
}

// ------------------------------------------------------------
//  DRIVE — get or create folder by name under parent
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
//  CONDUCT MEETING — saves to Conducted sheet, Drive, MoM
// ------------------------------------------------------------
function conductMeeting(payload) {
  try {
    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tz  = Session.getScriptTimeZone();
    var now = new Date();

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

    // 2. Save photos to Drive — wrapped in try-catch so sheet save always happens
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
    try { momUrl = createMoMDoc(payload, photoFolderUrl); } catch(e) { momUrl = ''; }

    // 5. Save to Conducted Meetings sheet
    var cSheet = ss.getSheetByName(CONDUCTED_SHEET);
    if (!cSheet) {
      cSheet = ss.insertSheet(CONDUCTED_SHEET);
      var ch = ['Meeting ID','District','Employee Name','Post','Email',
                'Original Date','Original Time','Duration','Meeting Type',
                'Stakeholder Name','Stakeholder Post','Purpose','Agenda',
                'Conduct Date','Conduct Time','Key Points',
                'Photos Folder','MoM Doc','Colleague Name','Colleague Post','Conducted At'];
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

    // 5. Update status in Plan Meetings to "Conducted" — NEVER delete, keeps master ledger intact for dashboard reporting
    if (planSheet && planRowIdx > 0) {
      planSheet.getRange(planRowIdx + 1, 14).setValue('Conducted');
    }

    // 6. Send MoM email to colleague
    if (payload.colleagueName && payload.colleagueName.trim()) {
      try { sendMOMNotification(payload, momUrl, photoFolderUrl, followUpId); } catch(mailErr) { /* don't fail conduct if mail fails */ }
    }

    invalidateUser((payload.email || '').trim().toLowerCase());
    return { success: true, momUrl: momUrl, photoFolderUrl: photoFolderUrl, followUpId: followUpId, photoError: photoError };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  MoM — auto Google Doc creation
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

  // Key Discussion Points
  sec('Key Discussion Points');
  var points = (d.keyPoints || '').split('\n').filter(function(p) { return p.trim(); });
  if (points.length) {
    points.forEach(function(pt) { body.appendListItem(pt.trim()); });
  } else {
    body.appendParagraph('—');
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
//  POSTPONE MEETING — same ID, new date, history in sheet
// ------------------------------------------------------------
function postponeMeeting(payload) {
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

    // Update status in Plan Meetings to "Cancelled" — NEVER delete, keeps master ledger intact for dashboard reporting
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
    // Field role cannot delete — prevents fake-meeting create-then-delete (full audit trail)
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
//  ACTIONED MEETINGS — for My Meetings view
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
//  ALL MEETINGS — combined view for My Meetings tab
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

    // 1. Postponed Meetings sheet — history of all reschedules
    var pSheet = ss.getSheetByName(POSTPONED_SHEET);
    if (pSheet && pSheet.getLastRow() > 1) {
      var phd = pSheet.getDataRange().getValues();
      // Columns: MeetingID(0) District(1) EmployeeName(2) Email(3)
      //          StakeholderName(4) StakeholderPost(5) Purpose(6)
      //          OriginalDate(7) NewDate(8) Reason(9) PostponedAt(10)
      for (var i = 1; i < phd.length; i++) {
        if ((phd[i][3] || '').toString().trim().toLowerCase() !== emailKey) continue;
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
      var cd = cSheet.getDataRange().getValues();
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
          colleagueName:(cd[j][18] || '').toString(),
          colleaguePost:(cd[j][19] || '').toString(),
          reason: '', parentMeetingId: ''
        });
      }
    }

    // 3. Cancelled Meetings
    var xSheet = ss.getSheetByName(CANCELLED_SHEET);
    if (xSheet && xSheet.getLastRow() > 1) {
      var xd = xSheet.getDataRange().getValues();
      for (var k = 1; k < xd.length; k++) {
        if ((xd[k][4] || '').toString().trim().toLowerCase() !== emailKey) continue;
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

    cPut(cacheKey, meetings, C_TTL_LIVE);
    return meetings;
  } catch(err) {
    return [];
  }
}

// ------------------------------------------------------------
//  DISTRICT ALL MEETINGS — for District Meetings view
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

    // 1. Plan Meetings — Planned / Follow-up only (Cancelled/Postponed go to their own sheets)
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    if (planSheet) {
      var pd = planSheet.getDataRange().getValues();
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
      var cd = condSheet.getDataRange().getValues();
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
      var xd = postSheet.getDataRange().getValues();
      for (var i = 1; i < xd.length; i++) {
        if ((xd[i][1] || '').toString().trim().toLowerCase() !== distL) continue;
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
      var xc = cancelSheet.getDataRange().getValues();
      for (var i = 1; i < xc.length; i++) {
        if ((xc[i][1] || '').toString().trim().toLowerCase() !== distL) continue;
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

    cPut(cacheKey, meetings, C_TTL_LIVE);
    return meetings;
  } catch(err) { return []; }
}

// ------------------------------------------------------------
//  STATE ALL MEETINGS — for State Meetings view
//  All districts, all statuses
// ------------------------------------------------------------
function getStateAllMeetings() {
  try {
    var cacheKey = 'stateMtg_all';
    var hit      = cGet(cacheKey);
    if (hit) return hit;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var meetings = [];

    // 1. Plan Meetings — Planned / Follow-up only
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    if (planSheet) {
      var pd = planSheet.getDataRange().getValues();
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
      var cd = condSheet.getDataRange().getValues();
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
      var xd = postSheet.getDataRange().getValues();
      for (var i = 1; i < xd.length; i++) {
        if (!xd[i][0]) continue;
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
      var xc = cancelSheet.getDataRange().getValues();
      for (var i = 1; i < xc.length; i++) {
        if (!xc[i][0]) continue;
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

    cPut(cacheKey, meetings, C_TTL_LIVE);
    return meetings;
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
            '<td style="padding:8px 0;">' + (data.duration || '—') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Meeting Type</td>' +
            '<td style="padding:8px 0;">' + (data.meetingType || '—') + '</td>' +
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
//  COLLEAGUE MOM EMAIL — sent after meeting is conducted
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
  if (!kpHtml) kpHtml = '<tr><td style="padding:5px 0;color:#6B7280;font-size:13px;">—</td></tr>';

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
            '<td style="padding:8px 0;">' + (data.meetingType || '—') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Purpose</td>' +
            '<td style="padding:8px 0;">' + (data.purpose || '—') + '</td>' +
          '</tr>' +
          '<tr style="border-bottom:1px solid #F3F4F6;">' +
            '<td style="padding:8px 0;color:#6B7280;">Conducted On</td>' +
            '<td style="padding:8px 0;font-weight:600;color:#111827;">' + (data.conductDate || '—') + (data.conductTime ? ' &nbsp;at&nbsp; ' + data.conductTime : '') + '</td>' +
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
//  DISTRICT REPORT — detailed breakdown for one district
// ------------------------------------------------------------
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
//  ALL CONDUCTED REPORTS — paginated list for dashboard
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
//  DASHBOARD STATS — cards & reports data
// ------------------------------------------------------------
function getDashboardStats(email, allDistricts) {
  try {
    var statKey = 'stats_' + email.trim().toLowerCase() + '_' + (allDistricts ? '1' : '0');
    var statHit = cGet(statKey);
    if (statHit) return statHit;

    var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    var emp = getEmployeeByEmail(email);
    var userRole     = emp ? (emp.role     || 'Field') : 'Field';
    var userDistrict = emp ? (emp.district || '')      : '';
    var isState      = allDistricts || (userRole === 'State');

    // ── Plan Meetings ──────────────────────────────────────────
    var planSheet = ss.getSheetByName(MEETINGS_SHEET);
    var planData  = (planSheet && planSheet.getLastRow() > 1) ? planSheet.getDataRange().getValues() : [];

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
      else if (status === 'planned')   distMap[dKey].planned++;
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

    // ── Conducted Meetings — enriched stats + recent ──────────────
    var cSheet      = ss.getSheetByName(CONDUCTED_SHEET);
    var recent      = [];
    var empSet      = {};     // unique employee names
    var stkPostMap  = {};     // stakeholder post → count
    var dtfCount    = 0;      // Group Meeting conducted
    var momReady    = 0;      // meetings with MoM doc link

    if (cSheet && cSheet.getLastRow() > 1) {
      var cd = cSheet.getDataRange().getValues();
      // Forward pass — collect stats
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
      // Reverse pass — collect recent 8
      for (var ri = cd.length - 1; ri >= 1 && recent.length < 8; ri--) {
        var rr    = cd[ri];
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
    var topStkPost = '—'; var topStkCount = 0;
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
//  INSERT SAMPLE DATA — run once from GAS editor
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
    ['MTG-S-H05','Hardoi','Uday Raj','District Impact Specialist','uday.raj@educategirls.ngo','2026-04-20','2:00 PM','1 hr','One-on-One','Vinod Sharma','CDO','Retention','Retention strategies for upper primary girls — Cancelled due to officer unavailability','Cancelled','','','Officer on leave','','','10/04/2026, 3:00:00 pm'],
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
    ['MTG-S-B01','Bahraich','Buddh Vilas','District Impact Specialist','buddh.vilas@educategirls.ngo','2026-04-16','10:00 AM','1 hr','One-on-One','Shyam Lal Gupta','BSA','Review Meeting','Annual review meeting — enrollment, retention, learning','Conducted','','','','Balwant Singh','District Operational Lead','08/04/2026, 9:00:00 am'],
    ['MTG-S-B02','Bahraich','Shyam Narayan Nath','District Program Officer','shyamnarayan.nath@educategirls.ngo','2026-05-26','11:00 AM','1 hr','One-on-One','Dr. Alka Jain','DIET Principal','Enrollment','Out-of-school girls data sharing with DIET','Planned','','','','','','18/04/2026, 10:00:00 am'],
    ['MTG-S-B03','Bahraich','Sanwara Vaishnav','District Program Training Officer','sanwara.vaishnav@educategirls.ngo','2026-05-04','9:00 AM','3 hr','Group Meeting','Deepak Kumar','ABSA','DTF','Training session on EG methodology and community mobilization','Conducted','','','','','','25/04/2026, 8:00:00 am'],
    ['MTG-S-B04','Bahraich','Balwant Singh','District Operational Lead','balwant.singh@educategirls.ngo','2026-04-30','2:00 PM','1 hr','One-on-One','Mohd. Azam Khan','District Collector','Introductory Meeting','Introductory meeting with new District Collector','Cancelled','','','New DC not yet joined charge','','','22/04/2026, 1:00:00 pm'],

    // SHAHJAHANPUR
    ['MTG-S-SJ1','Shahjahanpur','Indra Dev Tiwari','District Program Officer','indradev.tiwari@educategirls.ngo','2026-04-10','11:00 AM','1 hr','One-on-One','Surendra Bahadur Singh','BSA','MPR Submission','Monthly progress report submission — April','Conducted','','','','Ankit Kumar Dixit','District Program Officer','03/04/2026, 10:00:00 am'],
    ['MTG-S-SJ2','Shahjahanpur','Ankit Kumar Dixit','District Program Officer','ankit.dixit@educategirls.ngo','2026-04-23','10:30 AM','1 hr','One-on-One','Dr. Rama Kant','DIET Principal','Enrollment','Discuss enrollment targets and DIET support for EG program','Conducted','','','','','','14/04/2026, 9:00:00 am'],
    ['MTG-S-SJ3','Shahjahanpur','Chandra Mohan Sharma','District Program Training Officer','chandramohan.sharma@educategirls.ngo','2026-05-28','9:00 AM','4 hr','Group Meeting','Smt. Pushpa Singh','ABSA','DTF','District Training Facilitation — refresher session','Planned','','','','Indra Dev Tiwari','District Program Officer','20/04/2026, 8:00:00 am'],
    ['MTG-S-SJ4','Shahjahanpur','Vikas Kumar Tiwari','District Operational Assistant Lead','vikash.tiwari@educategirls.ngo','2026-05-07','3:00 PM','45 min','One-on-One','Ashutosh Verma','District Collector','Retention','Retention drive support request from district administration','Cancelled','','','Meeting rescheduled to next month','','','30/04/2026, 2:00:00 pm']
  ];

  // ─── SAMPLE CONDUCTED MEETINGS ────────────────────────────────
  // Cols: MtgID, Dist, EmpName, Post, Email, OrigDate, OrigTime, Duration,
  //       Type, StkName, StkPost, Purpose, Agenda, ConductDate, ConductTime,
  //       KeyPoints, PhotosFolder, MoMDoc, ColleagueName, ColleaguePost, ConductedAt
  var condRows = [
    ['MTG-S-H01','Hardoi','Uday Raj','District Impact Specialist','uday.raj@educategirls.ngo','2026-04-10','10:00 AM','1 hr','One-on-One','Rajesh Kumar Verma','BSA','Review Meeting','Quarterly review of enrollment and retention data','2026-04-10','10:45 AM','• Reviewed Q4 enrollment data — 87% target achieved\n• Discussed block-wise retention gaps in KPTG and SANDI\n• BSA agreed to issue circular for ABSA attendance in DTF sessions\n• Follow-up scheduled for May 15','','','','','10/04/2026, 11:50:00 am'],
    ['MTG-S-H02','Hardoi','Rahul Kumar','District Program Officer','rahul.kumar3@educategirls.ngo','2026-04-15','11:00 AM','45 min','One-on-One','Dr. Sunita Pathak','DIET Principal','Enrollment','Discuss strategies for out-of-school girl enrollment','2026-04-15','11:30 AM','• DIET will share block-wise OOS data by April 20\n• Principal agreed to conduct school-wise sensitization\n• EG to provide resource materials for DIET faculty\n• Joint visit to 3 schools planned for May','','','Manvendra Mishra','District Program Officer','15/04/2026, 12:05:00 pm'],
    ['MTG-S-H03','Hardoi','Manvendra Mishra','District Program Officer','manvendra.mishra@educategirls.ngo','2026-05-05','3:00 PM','30 min','One-on-One','Anil Tiwari','District Collector','Introductory Meeting','Initial introduction and EG program briefing','2026-05-05','3:15 PM','• DC appreciated EG program outcomes in Hardoi\n• Requested monthly update sheet for DC office\n• Discussed upcoming enrollment campaign — DC agreed to flag-off\n• Next meeting scheduled post elections','','','','','05/05/2026, 4:00:00 pm'],
    ['MTG-S-H06','Hardoi','Manvendra Mishra','District Program Officer','manvendra.mishra@educategirls.ngo','2026-05-18','11:30 AM','1 hr','Dept. Review','Ram Kishore','JD','MPR Submission','Submit monthly progress report and discuss targets','2026-05-18','12:00 PM','• April MPR submitted — 92% targets achieved\n• JD directed to improve learning outcomes data quality\n• EG team to share school-wise learning data by May 25\n• Monthly review mechanism to be strengthened','','','Uday Raj','District Impact Specialist','18/05/2026, 1:15:00 pm'],

    ['MTG-S-F01','Fatehpur','Shubham Yadav','District Impact Specialist','shubham.yadav@educategirls.ngo','2026-04-12','10:30 AM','1 hr','One-on-One','Pramod Srivastava','BSA','MPR Submission','Monthly progress report submission and follow-up','2026-04-12','11:00 AM','• March MPR accepted — strong enrollment numbers\n• BSA highlighted teacher absenteeism as key challenge\n• EG team to document school-wise attendance data\n• Follow-up on ABSA deployment in 3 blocks','','','Deepak Dixit','District Program Officer','12/04/2026, 12:00:00 pm'],
    ['MTG-S-F02','Fatehpur','Deepak Dixit','District Program Officer','deepak.dixit@educategirls.ngo','2026-04-18','11:00 AM','1 hr','One-on-One','Dr. Kavita Mishra','DIET Principal','Review Meeting','Mid-year review of learning outcomes and teacher training','2026-04-18','11:45 AM','• Learning assessment data reviewed — improvement in Grade 3-5\n• DIET agreed to include EG module in next BTC training\n• Principal to depute 2 DIET faculty for EG school visits\n• Collaborative workshop planned for June','','','','','18/04/2026, 12:30:00 pm'],
    ['MTG-S-F03','Fatehpur','Pushpendra Singh','District Program Training Officer','pushpendra.singh@educategirls.ngo','2026-04-25','9:00 AM','3 hr','Group Meeting','Anil Jaiswal','ABSA','DTF','Cluster-level training on learning assessment tools','2026-04-25','12:00 PM','• 18 ABSAs trained on NIPUN learning tools\n• Hands-on practice on assessment rubrics completed\n• All participants committed to weekly school monitoring\n• Next DTF scheduled for June','','','Shubham Yadav','District Impact Specialist','25/04/2026, 12:30:00 pm'],
    ['MTG-S-F06','Fatehpur','Deepak Dixit','District Program Officer','deepak.dixit@educategirls.ngo','2026-05-15','11:00 AM','1 hr','One-on-One','Suresh Patel','DC-Training','Learning','Discussion on training calendar and capacity building','2026-05-15','11:50 AM','• Training calendar for 2026-27 shared with DC Training\n• Three EG-specific modules approved for inclusion\n• Resource persons list to be shared by May 20\n• Joint review after first training cycle','','','','','15/05/2026, 12:20:00 pm'],

    ['MTG-S-G01','Gonda','Atul Pandey','District Impact Specialist','atul.pandey@educategirls.ngo','2026-04-08','10:00 AM','1 hr','One-on-One','Krishna Nand Yadav','BSA','Review Meeting','Review of EG program KPIs and district targets','2026-04-08','10:50 AM','• KPI review: enrollment 91%, retention 84%, learning 78%\n• BSA committed to resolve ABSA vacancy in 2 blocks\n• EG team to provide block-wise dashboard monthly\n• Next review in May with data from all 12 blocks','','','Ashish Kumar Singh','District Program Officer','08/04/2026, 11:55:00 am'],
    ['MTG-S-G02','Gonda','Ashish Kumar Singh','District Program Officer','ashishkumar.singh1@educategirls.ngo','2026-04-22','11:30 AM','45 min','One-on-One','Dr. Reena Verma','DIET Principal','Enrollment','DIET-EG collaboration for out-of-school girl data','2026-04-22','12:00 PM','• DIET OOS data for 8 blocks shared\n• Joint verification exercise to be conducted in May\n• EG and DIET to co-develop household survey tool\n• DIET faculty to support community mobilization','','','','','22/04/2026, 12:30:00 pm'],
    ['MTG-S-G04','Gonda','Arvind Kumar Yadav','Training Senior Specialist','arvind.yadav@educategirls.ngo','2026-05-01','9:00 AM','4 hr','Group Meeting','Ramesh Misra','ABSA','DTF','Training on NIPUN assessment and learning level improvement','2026-05-01','1:00 PM','• 22 ABSAs trained across 4 blocks\n• Practical sessions on NIPUN tools completed\n• Block-wise action plans prepared by each ABSA\n• Follow-up classroom observation scheduled for June','','','Vedprakash Yadav','District Program Officer','01/05/2026, 1:30:00 pm'],
    ['MTG-S-G06','Gonda','Ashish Kumar Singh','District Program Officer','ashishkumar.singh1@educategirls.ngo','2026-05-20','10:30 AM','1 hr','Dept. Review','Om Prakash Tiwari','DC- Gender','MPR Submission','Gender data review and MPR submission','2026-05-20','11:20 AM','• Gender-disaggregated data reviewed for April\n• Drop-out rate among girls in Class 6-8 flagged as concern\n• DC Gender to raise in DISE data meeting\n• EG to provide school-wise risk analysis','','','','','20/05/2026, 11:45:00 am'],

    ['MTG-S-S01','Sitapur','Sumit Kumar','District Impact Specialist','sumit.kumar3@educategirls.ngo','2026-04-14','11:00 AM','1 hr','One-on-One','Awadhesh Yadav','BSA','MPR Submission','Submit district MPR and review block-wise progress','2026-04-14','11:45 AM','• March MPR submitted — 89% enrollment, 82% retention\n• BSA requested EG data in Excel format for compilation\n• Block-wise performance matrix to be shared weekly\n• ABSA meeting to be organized in May','','','Vikrant Kumar','District Program Officer','14/04/2026, 12:30:00 pm'],
    ['MTG-S-S02','Sitapur','Vikrant Kumar','District Program Officer','vikrant.kumar@educategirls.ngo','2026-05-06','10:00 AM','1 hr','One-on-One','Dr. Shashi Bala','DIET Principal','Review Meeting','Review of DIET training effectiveness on EG teachers','2026-05-06','10:55 AM','• DIET training impact study data shared\n• Significant improvement in teacher facilitation skills noted\n• 3 best-practice schools identified for documentation\n• Exposure visit for DIET faculty to EG schools planned','','','','','06/05/2026, 11:20:00 am'],
    ['MTG-S-S03','Sitapur','Mohd Shadab Ansari','District Program Officer','shadab.ansari@educategirls.ngo','2026-05-12','3:30 PM','45 min','One-on-One','Vinay Kumar Gupta','District Collector','Enrollment','Enrollment campaign planning for 2026-27','2026-05-12','4:05 PM','• DC approved EG-led enrollment campaign for June\n• Gram Pradhan mobilization to be done via BDO circulars\n• EG team to prepare campaign material by May 20\n• DC office to share support letter for schools','','','','','12/05/2026, 4:30:00 pm'],

    ['MTG-S-B01','Bahraich','Buddh Vilas','District Impact Specialist','buddh.vilas@educategirls.ngo','2026-04-16','10:00 AM','1 hr','One-on-One','Shyam Lal Gupta','BSA','Review Meeting','Annual review meeting — enrollment, retention, learning','2026-04-16','10:50 AM','• Annual data reviewed — targets met in 7 of 9 blocks\n• Learning outcomes below benchmark in 2 blocks — plan needed\n• BSA agreed to depute resource persons for those blocks\n• EG to submit action plan by April 25','','','Balwant Singh','District Operational Lead','16/04/2026, 11:50:00 am'],
    ['MTG-S-B03','Bahraich','Sanwara Vaishnav','District Program Training Officer','sanwara.vaishnav@educategirls.ngo','2026-05-04','9:00 AM','3 hr','Group Meeting','Deepak Kumar','ABSA','DTF','Training session on EG methodology and community mobilization','2026-05-04','12:00 PM','• 16 ABSAs trained on EG community mobilization approach\n• Role-play exercises on parent engagement conducted\n• Commitments taken for monthly school-community meets\n• Refresher session scheduled for July','','','','','04/05/2026, 12:30:00 pm'],

    ['MTG-S-SJ1','Shahjahanpur','Indra Dev Tiwari','District Program Officer','indradev.tiwari@educategirls.ngo','2026-04-10','11:00 AM','1 hr','One-on-One','Surendra Bahadur Singh','BSA','MPR Submission','Monthly progress report submission — April','2026-04-10','11:50 AM','• March MPR submitted — 85% overall target achievement\n• BSA appreciated improvement in retention data quality\n• New data collection format to be piloted in 2 blocks\n• Follow-up meeting for April data in first week of May','','','Ankit Kumar Dixit','District Program Officer','10/04/2026, 12:00:00 pm'],
    ['MTG-S-SJ2','Shahjahanpur','Ankit Kumar Dixit','District Program Officer','ankit.dixit@educategirls.ngo','2026-04-23','10:30 AM','1 hr','One-on-One','Dr. Rama Kant','DIET Principal','Enrollment','Discuss enrollment targets and DIET support for EG program','2026-04-23','11:20 AM','• Enrollment targets for 2026-27 discussed and agreed\n• DIET to provide training support for 45 EG schools\n• Resource material library to be set up at DIET\n• Joint visit to 5 EG schools planned for May','','','','','23/04/2026, 11:45:00 am']
  ];

  // ─── SAMPLE CANCELLED MEETINGS ────────────────────────────────
  // Cols: MtgID, Dist, EmpName, Post, Email, Date, Time, Duration,
  //       Type, StkName, StkPost, Purpose, Agenda, ColleagueName, ColleaguePost, Reason, CancelledAt
  var cancRows = [
    ['MTG-S-H05','Hardoi','Uday Raj','District Impact Specialist','uday.raj@educategirls.ngo','2026-04-20','2:00 PM','1 hr','One-on-One','Vinod Sharma','CDO','Retention','Retention strategies for upper primary girls','','','Officer on leave — rescheduled','20/04/2026, 2:30:00 pm'],
    ['MTG-S-F05','Fatehpur','Shubham Yadav','District Impact Specialist','shubham.yadav@educategirls.ngo','2026-05-02','4:00 PM','30 min','One-on-One','Ajay Tripathi','CDO','Invitation','Invite CDO for EG annual review event','','','Event postponed by organizers','02/05/2026, 4:15:00 pm'],
    ['MTG-S-G05','Gonda','Atul Pandey','District Impact Specialist','atul.pandey@educategirls.ngo','2026-05-08','2:00 PM','30 min','One-on-One','Hari Om Mishra','JD','Retention','Discuss retention challenges at upper primary level','','','Meeting cancelled by stakeholder — national duty','08/05/2026, 2:20:00 pm'],
    ['MTG-S-S05','Sitapur','Sumit Kumar','District Impact Specialist','sumit.kumar3@educategirls.ngo','2026-04-28','4:00 PM','30 min','One-on-One','Ajeet Singh','CDO','Courtesy Meeting','Courtesy visit and program update to CDO','','','CDO transferred to another district','28/04/2026, 4:10:00 pm'],
    ['MTG-S-B04','Bahraich','Balwant Singh','District Operational Lead','balwant.singh@educategirls.ngo','2026-04-30','2:00 PM','1 hr','One-on-One','Mohd. Azam Khan','District Collector','Introductory Meeting','Introductory meeting with new District Collector','','','New DC not yet joined charge','30/04/2026, 2:15:00 pm'],
    ['MTG-S-SJ4','Shahjahanpur','Vikas Kumar Tiwari','District Operational Assistant Lead','vikash.tiwari@educategirls.ngo','2026-05-07','3:00 PM','45 min','One-on-One','Ashutosh Verma','District Collector','Retention','Retention drive support request from district administration','','','Meeting rescheduled to next month — DC tour','07/05/2026, 3:20:00 pm']
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

function getEmployeeByEmail(email) {
  var key = 'emp_' + email;
  var hit = cGet(key);
  if (hit !== null) return hit;   // null-employee cached as JSON null → re-fetch only on miss

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  // Columns: District(0), Block(1), Employee Name(2), Designation(3), Email(4), Role(5)
  var result = null;
  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][4] ? data[i][4].toString().trim().toLowerCase() : '';
    if (rowEmail === email) {
      result = {
        district:    (data[i][0] || '').toString().trim(),
        block:       (data[i][1] || '').toString().trim(),
        name:        (data[i][2] || '').toString().trim(),
        designation: (data[i][3] || '').toString().trim(),
        email:       (data[i][4] || '').toString().trim(),
        role:        (data[i][5] || 'Field').toString().trim()
      };
      break;
    }
  }
  cPut(key, result, C_TTL_EMP);
  return result;
}

// ------------------------------------------------------------
//  BULK UPDATE EMPLOYEE_DB — POST action: bulkUpdateEmployeeDB
//  Accepts { rows: [[District,Block,Name,Designation,Email,Role], ...] }
//  Clears existing data (except header) and writes new rows
// ------------------------------------------------------------
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
//  PEEK SOURCE SHEET — returns sheet names + first 3 data rows
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
      message: 'Import complete — vacant rows skipped'
    };
  } catch(err) {
    return { success: false, message: err.message };
  }
}


// ------------------------------------------------------------
//  CLEAR MY CACHE — call after role/data change to force refresh
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