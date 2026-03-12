/**
 * Code.gs — Google Apps Script backend for Y7 Worksheets submission system
 *
 * SETUP INSTRUCTIONS FOR TEACHER:
 * 1. Create a new Google Sheet (name it "Y7 Worksheet Submissions")
 * 2. Open Extensions → Apps Script
 * 3. Paste this entire file into the script editor (replace any existing code)
 * 4. Click the ⚙ Project Settings (gear icon) → add a Script Property:
 *       Property name:  TEACHER_KEY
 *       Value:          (make up a password, e.g. "mySecret2026")
 * 5. Click Deploy → New deployment
 *       Type:       Web app
 *       Execute as: Me
 *       Who has access: Anyone
 * 6. Click Deploy → copy the URL
 * 7. Paste the URL into submission.js (APPS_SCRIPT_URL constant)
 *    and into teacher-dashboard.html (APPS_SCRIPT_URL constant)
 *
 * SHEET STRUCTURE (auto-created on first submission):
 * Sheet "Submissions":
 *   A: Timestamp | B: StudentName | C: StudentClass | D: WorksheetId |
 *   E: WorksheetTitle | F: SubmissionCount | G: DataJSON |
 *   H: FeedbackText | I: FeedbackDate
 *
 * PRIVACY:
 *   - All data stored in YOUR Google Drive (school-managed Google Workspace)
 *   - No third-party services involved
 *   - No student email addresses collected
 *   - TEACHER_KEY protects dashboard-only endpoints
 */

// ===== CORS + ROUTING =====

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    var result;
    switch (action) {
      case 'submit':
        result = handleSubmit(payload);
        break;
      case 'giveFeedback':
        result = handleGiveFeedback(payload);
        break;
      default:
        result = { status: 'error', message: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action;

    var result;
    switch (action) {
      case 'feedback':
        result = handleGetFeedback(e.parameter);
        break;
      case 'students':
        result = handleGetStudents(e.parameter);
        break;
      case 'detail':
        result = handleGetDetail(e.parameter);
        break;
      default:
        result = { status: 'error', message: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===== SHEET HELPERS =====

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Submissions');
  if (!sheet) {
    sheet = ss.insertSheet('Submissions');
    // Set headers
    sheet.getRange('A1:I1').setValues([[
      'Timestamp', 'StudentName', 'StudentClass', 'WorksheetId',
      'WorksheetTitle', 'SubmissionCount', 'DataJSON',
      'FeedbackText', 'FeedbackDate'
    ]]);
    sheet.getRange('A1:I1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Set column widths
    sheet.setColumnWidth(1, 160);  // Timestamp
    sheet.setColumnWidth(2, 150);  // Name
    sheet.setColumnWidth(3, 80);   // Class
    sheet.setColumnWidth(4, 140);  // WorksheetId
    sheet.setColumnWidth(5, 160);  // WorksheetTitle
    sheet.setColumnWidth(6, 50);   // Count
    sheet.setColumnWidth(7, 300);  // DataJSON
    sheet.setColumnWidth(8, 300);  // FeedbackText
    sheet.setColumnWidth(9, 160);  // FeedbackDate
  }
  return sheet;
}

/**
 * Find existing row for a student + worksheet combo.
 * Returns row number (1-based) or -1 if not found.
 */
function findStudentRow(sheet, name, cls, worksheetId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {  // skip header
    if (data[i][1] === name && data[i][2] === cls && data[i][3] === worksheetId) {
      return i + 1;  // 1-based row number
    }
  }
  return -1;
}

function checkTeacherKey(params) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('TEACHER_KEY');
  if (!key) return true;  // no key set = no protection (for initial testing)
  return params.key === key;
}

// ===== HANDLERS =====

/**
 * POST: Student submits their work
 * Upserts by (name + class + worksheetId)
 */
function handleSubmit(payload) {
  var sheet = getOrCreateSheet();
  var name = (payload.studentName || '').trim();
  var cls = (payload.studentClass || '').trim();
  var wsId = payload.worksheetId || '';
  var wsTitle = payload.worksheetTitle || '';
  var data = payload.data || {};

  if (!name || !cls) {
    return { status: 'error', message: 'Name and class are required.' };
  }

  var row = findStudentRow(sheet, name, cls, wsId);
  var now = new Date().toISOString();
  var dataJson = JSON.stringify(data);

  if (row > 0) {
    // Update existing row
    var currentCount = sheet.getRange(row, 6).getValue() || 0;
    sheet.getRange(row, 1).setValue(now);                    // Timestamp
    sheet.getRange(row, 6).setValue(currentCount + 1);       // SubmissionCount
    sheet.getRange(row, 7).setValue(dataJson);               // DataJSON
  } else {
    // New row
    sheet.appendRow([
      now,        // Timestamp
      name,       // StudentName
      cls,        // StudentClass
      wsId,       // WorksheetId
      wsTitle,    // WorksheetTitle
      1,          // SubmissionCount
      dataJson,   // DataJSON
      '',         // FeedbackText (empty)
      ''          // FeedbackDate (empty)
    ]);
  }

  return { status: 'ok', message: 'Submission received.' };
}

/**
 * GET: Student checks for teacher feedback
 * Params: name, class, ws (worksheetId)
 * No teacher key required — students need this
 */
function handleGetFeedback(params) {
  var sheet = getOrCreateSheet();
  var name = (params.name || '').trim();
  var cls = (params['class'] || '').trim();
  var wsId = params.ws || '';

  if (!name || !cls) {
    return { status: 'ok', hasFeedback: false };
  }

  var row = findStudentRow(sheet, name, cls, wsId);
  if (row < 0) {
    return { status: 'ok', hasFeedback: false };
  }

  var feedbackText = sheet.getRange(row, 8).getValue();
  var feedbackDate = sheet.getRange(row, 9).getValue();

  if (!feedbackText) {
    return { status: 'ok', hasFeedback: false };
  }

  return {
    status: 'ok',
    hasFeedback: true,
    feedbackText: feedbackText,
    feedbackDate: feedbackDate
  };
}

/**
 * GET: Teacher dashboard — list all students for a worksheet
 * Params: key (teacher key), ws (worksheetId, optional), class (optional)
 */
function handleGetStudents(params) {
  if (!checkTeacherKey(params)) {
    return { status: 'error', message: 'Invalid teacher key.' };
  }

  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var students = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Filter by worksheet if specified
    if (params.ws && row[3] !== params.ws) continue;
    // Filter by class if specified
    if (params['class'] && row[2] !== params['class']) continue;

    // Parse data to extract summary info
    var appData = {};
    try { appData = JSON.parse(row[6] || '{}'); } catch (e) {}

    students.push({
      name: row[1],
      class: row[2],
      worksheetId: row[3],
      worksheetTitle: row[4],
      submissionCount: row[5],
      lastSubmitted: row[0],
      band: appData.band || 'unknown',
      diagnosticScore: appData.diagnosticScore || 0,
      modulesCompleted: countCompleted(appData.moduleCompleted),
      totalModules: 8,
      hasFeedback: !!row[7],
      feedbackText: row[7] || '',
      feedbackDate: row[8] || ''
    });
  }

  return { status: 'ok', students: students };
}

/**
 * GET: Teacher dashboard — get full submission data for one student
 * Params: key, name, class, ws
 */
function handleGetDetail(params) {
  if (!checkTeacherKey(params)) {
    return { status: 'error', message: 'Invalid teacher key.' };
  }

  var sheet = getOrCreateSheet();
  var name = (params.name || '').trim();
  var cls = (params['class'] || '').trim();
  var wsId = params.ws || '';

  var row = findStudentRow(sheet, name, cls, wsId);
  if (row < 0) {
    return { status: 'error', message: 'Student not found.' };
  }

  var dataJson = sheet.getRange(row, 7).getValue();
  var appData = {};
  try { appData = JSON.parse(dataJson || '{}'); } catch (e) {}

  return {
    status: 'ok',
    name: name,
    class: cls,
    worksheetId: wsId,
    submissionCount: sheet.getRange(row, 6).getValue(),
    lastSubmitted: sheet.getRange(row, 1).getValue(),
    data: appData,
    feedbackText: sheet.getRange(row, 8).getValue() || '',
    feedbackDate: sheet.getRange(row, 9).getValue() || ''
  };
}

/**
 * POST: Teacher gives feedback to a student
 * Payload: key, studentName, studentClass, worksheetId, feedbackText
 */
function handleGiveFeedback(payload) {
  if (!checkTeacherKey({ key: payload.key })) {
    return { status: 'error', message: 'Invalid teacher key.' };
  }

  var sheet = getOrCreateSheet();
  var name = (payload.studentName || '').trim();
  var cls = (payload.studentClass || '').trim();
  var wsId = payload.worksheetId || '';
  var feedback = (payload.feedbackText || '').trim();

  if (!name || !cls || !feedback) {
    return { status: 'error', message: 'Name, class, and feedback are required.' };
  }

  var row = findStudentRow(sheet, name, cls, wsId);
  if (row < 0) {
    return { status: 'error', message: 'Student submission not found.' };
  }

  sheet.getRange(row, 8).setValue(feedback);              // FeedbackText
  sheet.getRange(row, 9).setValue(new Date().toISOString()); // FeedbackDate

  return { status: 'ok', message: 'Feedback saved.' };
}

// ===== UTILITIES =====

function countCompleted(moduleCompleted) {
  if (!moduleCompleted) return 0;
  var count = 0;
  for (var key in moduleCompleted) {
    if (moduleCompleted[key] === true) count++;
  }
  return count;
}
