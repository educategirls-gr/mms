// ============================================================
//  Setup.gs - One-time sheet setup
//  Run setupSheets() manually once from Apps Script editor
// ============================================================

function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  createPlanMeetingsSheet(ss);
  createConductedSheet(ss);
  createPostponedSheet(ss);
  createCancelledSheet(ss);

  SpreadsheetApp.flush();
  Logger.log('Setup complete - 4 sheets created.');
}

// ── Helper ────────────────────────────────────────────────────
function styleHeader(sheet, headers, bg) {
  var hdr = sheet.getRange(1, 1, 1, headers.length);
  hdr.setBackground(bg || '#1F4E79');
  hdr.setFontColor('#FFFFFF');
  hdr.setFontWeight('bold');
  hdr.setFontSize(10);
  hdr.setWrap(false);
  sheet.setFrozenRows(1);
}
function setWidths(sheet, widths) {
  for (var i = 0; i < widths.length; i++) sheet.setColumnWidth(i + 1, widths[i]);
}
function getOrCreate(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  else s.clearContents();
  return s;
}

// ------------------------------------------------------------
//  Plan Meetings  (A-T = 20 columns - only Planned status rows)
//  Matches saveMeeting() in Code.gs exactly
// ------------------------------------------------------------
function createPlanMeetingsSheet(ss) {
  var sheet = getOrCreate(ss, 'Plan Meetings');

  var headers = [
    'Meeting ID',       // A  0
    'District',         // B  1
    'Employee Name',    // C  2
    'Post',             // D  3
    'Email',            // E  4
    'Meeting Date',     // F  5
    'Meeting Time',     // G  6
    'Duration',         // H  7
    'Meeting Type',     // I  8
    'Stakeholder Name', // J  9
    'Stakeholder Post', // K  10
    'Meeting Purpose',  // L  11
    'Meeting Agenda',   // M  12
    'Status',           // N  13  always Planned (or empty if postponed back)
    'Start Time',       // O  14  reserved (not used)
    'End Time',         // P  15  reserved (not used)
    'Reason',           // Q  16  postpone reason if re-scheduled
    'Colleague Name',   // R  17
    'Colleague Post',   // S  18
    'Submitted At'      // T  19
  ];

  sheet.appendRow(headers);

  // Color groups
  sheet.getRange(1, 1, 1, 5).setBackground('#1F4E79');   // ID + employee → dark blue
  sheet.getRange(1, 6, 1, 9).setBackground('#2E75B6');   // meeting details → blue
  sheet.getRange(1, 15, 1, 3).setBackground('#843C0C');  // reserved/reason → dark red
  sheet.getRange(1, 18, 1, 2).setBackground('#7030A0'); // colleague → purple
  sheet.getRange(1, 20, 1, 1).setBackground('#595959'); // meta → gray
  sheet.getRange(1, 14, 1, 1).setBackground('#166534'); // status → green
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);

  setWidths(sheet, [
    130, 120, 160, 160, 200,  // A-E
    110, 90,  100, 120,        // F-I
    160, 180, 200, 250,        // J-M
    100, 80,  80,  200,        // N-Q
    160, 160, 160              // R-T
  ]);

  Logger.log('Plan Meetings sheet created (20 columns).');
}

// ------------------------------------------------------------
//  Conducted Meetings  (21 columns)
//  Matches conductMeeting() in Code.gs exactly
// ------------------------------------------------------------
function createConductedSheet(ss) {
  var sheet = getOrCreate(ss, 'Conducted Meetings');

  var headers = [
    'Meeting ID',       // A  0
    'District',         // B  1
    'Employee Name',    // C  2
    'Post',             // D  3
    'Email',            // E  4
    'Original Date',    // F  5
    'Original Time',    // G  6
    'Duration',         // H  7
    'Meeting Type',     // I  8
    'Stakeholder Name', // J  9
    'Stakeholder Post', // K  10
    'Meeting Purpose',  // L  11
    'Meeting Agenda',   // M  12
    'Conduct Date',     // N  13
    'Conduct Time',     // O  14
    'Key Discussion Points', // P  15
    'Photos Folder Link',    // Q  16
    'MoM Doc Link',          // R  17
    'Colleague Name',   // S  18
    'Colleague Post',   // T  19
    'Conducted At'      // U  20
  ];

  sheet.appendRow(headers);

  sheet.getRange(1, 1, 1, 5).setBackground('#166534');   // ID + employee
  sheet.getRange(1, 6, 1, 8).setBackground('#2E7D32');   // original meeting info
  sheet.getRange(1, 14, 1, 3).setBackground('#0F766E');  // conduct info
  sheet.getRange(1, 17, 1, 2).setBackground('#1F4E79');  // links
  sheet.getRange(1, 19, 1, 2).setBackground('#7030A0'); // colleague
  sheet.getRange(1, 21, 1, 1).setBackground('#595959'); // meta
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);

  setWidths(sheet, [
    130, 120, 160, 160, 200,   // A-E
    110, 90,  100, 120,         // F-I
    160, 180, 200, 250,         // J-M
    110, 90,  300,              // N-P
    300, 300,                   // Q-R
    160, 160, 160               // S-U
  ]);

  Logger.log('Conducted Meetings sheet created (21 columns).');
}

// ------------------------------------------------------------
//  Postponed Meetings  (11 columns - history log)
//  Matches postponeMeeting() in Code.gs exactly
// ------------------------------------------------------------
function createPostponedSheet(ss) {
  var sheet = getOrCreate(ss, 'Postponed Meetings');

  var headers = [
    'Meeting ID',       // A  0
    'District',         // B  1
    'Employee Name',    // C  2
    'Email',            // D  3
    'Stakeholder Name', // E  4
    'Stakeholder Post', // F  5
    'Meeting Purpose',  // G  6
    'Original Date',    // H  7
    'New Date',         // I  8
    'Reason',           // J  9
    'Postponed At'      // K  10
  ];

  sheet.appendRow(headers);
  styleHeader(sheet, headers, '#B45309');

  setWidths(sheet, [
    130, 120, 160, 200,
    160, 180, 200,
    110, 110, 250, 160
  ]);

  Logger.log('Postponed Meetings sheet created (11 columns).');
}

// ------------------------------------------------------------
//  Cancelled Meetings  (17 columns)
//  Matches cancelMeeting() in Code.gs exactly
// ------------------------------------------------------------
function createCancelledSheet(ss) {
  var sheet = getOrCreate(ss, 'Cancelled Meetings');

  var headers = [
    'Meeting ID',       // A  0
    'District',         // B  1
    'Employee Name',    // C  2
    'Post',             // D  3
    'Email',            // E  4
    'Meeting Date',     // F  5
    'Meeting Time',     // G  6
    'Duration',         // H  7
    'Meeting Type',     // I  8
    'Stakeholder Name', // J  9
    'Stakeholder Post', // K  10
    'Meeting Purpose',  // L  11
    'Meeting Agenda',   // M  12
    'Colleague Name',   // N  13
    'Colleague Post',   // O  14
    'Reason',           // P  15
    'Cancelled At'      // Q  16
  ];

  sheet.appendRow(headers);
  styleHeader(sheet, headers, '#DC2626');

  setWidths(sheet, [
    130, 120, 160, 160, 200,
    110, 90,  100, 120,
    160, 180, 200, 250,
    160, 160, 250, 160
  ]);

  Logger.log('Cancelled Meetings sheet created (17 columns).');
}
