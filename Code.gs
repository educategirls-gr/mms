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
var DRIVE_ROOT_ID    = '1TlayZ0R6xCMwmumqPbQFxJ24eqp71dB1'; // EG-GR-Meetings Drive folder
var OTP_EXPIRY_SEC   = 600;
var ALLOWED_DOMAIN   = 'educategirls.ngo';

// ------------------------------------------------------------
//  ENTRY POINT
// ------------------------------------------------------------
// ------------------------------------------------------------
//  API HANDLER — called from GitHub Pages frontend via fetch()
// ------------------------------------------------------------
function doPost(e) {
  return apiResponse(e, 'POST');
}

function apiResponse(e, method) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var result;
  try {
    var body = {};
    if (method === 'POST' && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch(pe) { body = {}; }
    }
    if      (action === 'sendOTP')              result = sendOTP(e.parameter.email || '');
    else if (action === 'verifyOTP')            result = verifyOTP(e.parameter.email || '', e.parameter.otp || '');
    else if (action === 'getDropdownData')      result = getDropdownData(e.parameter.email || '');
    else if (action === 'getMyMeetings')        result = getMyMeetings(e.parameter.email || '');
    else if (action === 'getAllMyMeetings')      result = getAllMyMeetings(e.parameter.email || '');
    else if (action === 'getDistrictEmployees') result = getDistrictEmployees(e.parameter.district || '', e.parameter.email || '');
    else if (action === 'deleteMeeting')        result = deleteMeeting(e.parameter.meetingId || '');
    else if (action === 'saveMeeting')          result = saveMeeting(body);
    else if (action === 'conductMeeting')       result = conductMeeting(body);
    else if (action === 'postponeMeeting')      result = postponeMeeting(body);
    else if (action === 'cancelMeeting')        result = cancelMeeting(body);
    else if (action === 'updateMeetingStatus')  result = updateMeetingStatus(body.meetingId || '', body);
    else                                        result = { success: false, message: 'Unknown action: ' + action };
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
  DriveApp.getFolderById(DRIVE_ROOT_ID).getName();
  MailApp.getRemainingDailyQuota();
  Logger.log('Authorization successful.');
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
          date:         rawDate instanceof Date
                          ? Utilities.formatDate(rawDate, tz, 'dd-MM-yyyy')
                          : (rawDate || '').toString(),
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
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
    var e = data[i][4] ? data[i][4].toString().trim().toLowerCase() : '';
    if (d === district.trim().toLowerCase() && e !== currentEmail.trim().toLowerCase()) {
      list.push({
        name:        data[i][2] || '',
        designation: data[i][3] || ''
      });
    }
  }
  return list;
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

// ------------------------------------------------------------
//  DRIVE — get or create folder by name under parent
// ------------------------------------------------------------
function getOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
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
    try {
      if (payload.photos && payload.photos.length > 0) {
        var root  = DriveApp.getFolderById(DRIVE_ROOT_ID);
        var distF = getOrCreateFolder(root, payload.district || 'General');
        var mtgF  = getOrCreateFolder(distF, payload.meetingId);
        mtgF.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoFolderUrl = mtgF.getUrl();
        payload.photos.forEach(function(p, idx) {
          var ext  = (p.type || 'image/jpeg').split('/')[1];
          var blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.type,
                       payload.meetingId + '_' + (idx+1) + '.' + ext);
          var f = mtgF.createFile(blob);
          f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        });
      }
    } catch(de) { photoFolderUrl = ''; }

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

    return { success: true, momUrl: momUrl, photoFolderUrl: photoFolderUrl, followUpId: followUpId };
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
      var root  = DriveApp.getFolderById(DRIVE_ROOT_ID);
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
    var meetingDate = rawDate instanceof Date ? Utilities.formatDate(rawDate, tz, 'dd-MM-yyyy') : (rawDate || '').toString();
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

    return { success: true };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ------------------------------------------------------------
//  DELETE MEETING
// ------------------------------------------------------------
function deleteMeeting(meetingId) {
  try {
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
          date:         (cd[i][5]  || '').toString(),  // Original Date
          meetingTime:  fmtTimeVal(cd[i][6]),           // Original Time
          duration:     (cd[i][7]  || '').toString(),
          type:         (cd[i][8]  || '').toString(),
          adhikariName: (cd[i][9]  || '').toString(),
          adhikariPost: (cd[i][10] || '').toString(),
          purpose:      (cd[i][11] || '').toString(),
          agenda:       (cd[i][12] || '').toString(),
          status:       'Conducted',
          conductDate:  (cd[i][13] || '').toString(),
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
          date:         (xd[j][5]  || '').toString(),  // Meeting Date
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
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var emailKey = email.trim().toLowerCase();
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
          date:            (phd[i][7] || '').toString(),   // Original Date
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
          date:         (cd[j][5]  || '').toString(),
          meetingTime:  fmtTimeVal(cd[j][6]),
          duration:     (cd[j][7]  || '').toString(),
          type:         (cd[j][8]  || '').toString(),
          adhikariName: (cd[j][9]  || '').toString(),
          adhikariPost: (cd[j][10] || '').toString(),
          purpose:      (cd[j][11] || '').toString(),
          agenda:       (cd[j][12] || '').toString(),
          status:       'Conducted',
          conductDate:  (cd[j][13] || '').toString(),
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
          date:         (xd[k][5]  || '').toString(),
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

    return meetings;
  } catch(err) {
    return [];
  }
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
//  EMPLOYEE LOOKUP
// ------------------------------------------------------------
function getEmployeeByEmail(email) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  // Columns: District(0), Block(1), Employee Name(2), Designation(3), Email(4), Role(5)
  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][4] ? data[i][4].toString().trim().toLowerCase() : '';
    if (rowEmail === email) {
      return {
        district:    data[i][0] || '',
        block:       data[i][1] || '',
        name:        data[i][2] || '',
        designation: data[i][3] || '',
        email:       data[i][4] || '',
        role:        data[i][5] || 'Field'
      };
    }
  }
  return null;
}
