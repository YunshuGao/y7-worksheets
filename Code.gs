/**
 * Code.gs — Google Apps Script backend for Write It Right (Multi-Teacher Edition)
 *
 * VERSION: 2.0 — Expanded for TAS faculty-wide deployment
 *
 * WHAT'S NEW IN V2:
 *   - Passcode auth via Teachers tab (replaces single TEACHER_KEY)
 *   - Subject column in Submissions for multi-subject support
 *   - CommentBank tab for categorised teacher feedback
 *   - Server-side DataJSON parsing for dashboard summaries
 *   - Heatmap endpoint, CSV export, batch feedback
 *   - Backward-compatible: student endpoints (submit, feedback) unchanged
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your existing "Y7 Worksheet Submissions" Google Sheet
 * 2. Open Extensions > Apps Script
 * 3. Paste this file (replace old Code.gs)
 * 4. The script auto-creates Teachers + CommentBank tabs on first use
 * 5. Add yourself to the Teachers tab:
 *      Passcode: gao2026 | DisplayName: Ms Gao | Classes: 7ECE,7ECI,7TECJ,8TECI
 * 6. Deploy > Manage deployments > edit existing > bump version > Deploy
 * 7. No changes needed in submission.js or Write-It-Right.html
 *
 * SHEET STRUCTURE:
 *
 * Tab "Submissions" (columns A-J, auto-migrated from 9 to 10 columns):
 *   A: Timestamp | B: StudentName | C: StudentClass | D: WorksheetId |
 *   E: WorksheetTitle | F: SubmissionCount | G: DataJSON |
 *   H: FeedbackText | I: FeedbackDate | J: Subject
 *
 * Tab "Teachers" (you manage manually):
 *   A: Passcode | B: DisplayName | C: Classes (comma-separated)
 *
 * Tab "CommentBank" (you manage manually):
 *   A: Category | B: Subject | C: CommentText
 *
 * PRIVACY:
 *   - All data stored in YOUR Google Drive
 *   - No third-party services involved
 *   - No student email addresses collected
 *   - Teacher passcodes are simple strings (not hashed — acceptable for
 *     low-stakes school use; DET accounts can't use Google OAuth anyway)
 */

// ===== ROUTING =====

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
      case 'batchFeedback':
        result = handleBatchFeedback(payload);
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
      // --- Student-facing (no auth) ---
      case 'feedback':
        result = handleGetFeedback(e.parameter);
        break;
      // --- Teacher-facing (passcode auth) ---
      case 'login':
        result = handleLogin(e.parameter);
        break;
      case 'students':
        result = handleGetStudents(e.parameter);
        break;
      case 'detail':
        result = handleGetDetail(e.parameter);
        break;
      case 'commentBank':
        result = handleGetCommentBank(e.parameter);
        break;
      case 'heatmap':
        result = handleGetHeatmap(e.parameter);
        break;
      case 'export':
        result = handleExport(e.parameter);
        break;
      case 'diag':
        result = handleDiag(e.parameter);
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


// ===================================================================
//  SHEET HELPERS
// ===================================================================

/**
 * Get or create the Submissions tab.
 * V2 adds column J (Subject). Auto-migrates existing 9-column sheets.
 */
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Submissions');

  if (!sheet) {
    // Brand new — create with all 10 columns
    sheet = ss.insertSheet('Submissions');
    sheet.getRange('A1:J1').setValues([[
      'Timestamp', 'StudentName', 'StudentClass', 'WorksheetId',
      'WorksheetTitle', 'SubmissionCount', 'DataJSON',
      'FeedbackText', 'FeedbackDate', 'Subject'
    ]]);
    sheet.getRange('A1:J1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 160);
    sheet.setColumnWidth(6, 50);
    sheet.setColumnWidth(7, 300);
    sheet.setColumnWidth(8, 300);
    sheet.setColumnWidth(9, 160);
    sheet.setColumnWidth(10, 100);
  } else {
    // Migrate: add Subject column if missing (existing sheets have 9 cols)
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('Subject') === -1) {
      var nextCol = headers.length + 1;
      sheet.getRange(1, nextCol).setValue('Subject').setFontWeight('bold');
      sheet.setColumnWidth(nextCol, 100);
    }
  }

  return sheet;
}

/**
 * Build a column map from header names → 0-based indices.
 * Handles BOTH old (6-col) and new (10-col) sheet layouts.
 *
 * Old layout: Timestamp | Student Name | Class | Worksheet | Data JSON | Subject
 * New layout: Timestamp | StudentName | StudentClass | WorksheetId | WorksheetTitle | SubmissionCount | DataJSON | FeedbackText | FeedbackDate | Subject
 *
 * The map always returns { timestamp, name, class, wsId, wsTitle, count, data, feedback, feedbackDate, subject }
 * with -1 for columns that don't exist.
 */
function buildColMap(headers) {
  var map = { timestamp: 0, name: -1, class: -1, wsId: -1, wsTitle: -1, count: -1, data: -1, feedback: -1, feedbackDate: -1, subject: -1 };

  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || '').toString().trim().toLowerCase().replace(/[\s_]+/g, '');
    // Match flexibly: "Student Name", "StudentName", "student_name" all → name
    if (h === 'studentname' || h === 'name') map.name = i;
    else if (h === 'studentclass' || h === 'class') map.class = i;
    else if (h === 'worksheetid' || h === 'worksheet') map.wsId = i;
    else if (h === 'worksheettitle') map.wsTitle = i;
    else if (h === 'submissioncount') map.count = i;
    else if (h === 'datajson' || h === 'data json' || h === 'datajson') map.data = i;
    else if (h === 'feedbacktext') map.feedback = i;
    else if (h === 'feedbackdate') map.feedbackDate = i;
    else if (h === 'subject') map.subject = i;
    else if (h === 'timestamp') map.timestamp = i;
  }

  // In old 6-col layout, "Worksheet" column holds the title (no separate ID)
  // and "Data JSON" holds the appData blob
  if (map.wsTitle === -1 && map.wsId >= 0) {
    // Old layout: wsId column actually holds the worksheet title
    // wsId and wsTitle are the same column
    map.wsTitle = map.wsId;
  }

  return map;
}

/** Helper: safely read a row value by column map index */
function colVal(row, idx) {
  if (idx < 0 || idx >= row.length) return '';
  return row[idx];
}

/**
 * Get or create the Teachers tab.
 * Columns: A: Passcode | B: DisplayName | C: Classes
 */
function getOrCreateTeachersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Teachers');
  if (!sheet) {
    sheet = ss.insertSheet('Teachers');
    sheet.getRange('A1:C1').setValues([['Passcode', 'DisplayName', 'Classes']]);
    sheet.getRange('A1:C1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 250);
  }
  return sheet;
}

/**
 * Get or create the CommentBank tab.
 * Columns: A: Category | B: Subject | C: CommentText
 */
function getOrCreateCommentBankSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CommentBank');
  if (!sheet) {
    sheet = ss.insertSheet('CommentBank');
    sheet.getRange('A1:C1').setValues([['Category', 'Subject', 'CommentText']]);
    sheet.getRange('A1:C1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 400);
    // Seed with starter comments
    var seed = [
      ['teel',       'all', 'Your topic sentence clearly states the main idea.'],
      ['teel',       'all', 'Add a specific example from class to support your point.'],
      ['teel',       'all', 'Your explanation connects the evidence back to your topic sentence — well done.'],
      ['teel',       'all', 'Try adding a linking sentence that returns to the original question.'],
      ['vocabulary', 'all', 'Great use of technical vocabulary in your response.'],
      ['vocabulary', 'all', 'Try replacing everyday words with Tier 2 or Tier 3 vocabulary.'],
      ['grammar',    'all', 'Check your verb tenses — keep them consistent throughout.'],
      ['grammar',    'all', 'Watch your subject-verb agreement (e.g. "the design was" not "the design were").'],
      ['general',    'all', 'Well done — you have completed all modules!'],
      ['general',    'all', 'Keep practising your sentence structure — you are improving.'],
      ['general',    'all', 'Please complete the remaining modules before the end of term.'],
      ['revision',   'all', 'This response does not sound like your own writing. Please rewrite it in your own words — I want to hear YOUR ideas.'],
      ['revision',   'all', 'Your answer is too advanced for what we have covered in class. Please rewrite using the vocabulary and sentence patterns from the modules.'],
      ['revision',   'all', 'I can see you understand the topic, but I need you to express it in your own words. Use the sentence starters from the word bank to help you.'],
      ['revision',   'all', 'This section needs to be rewritten. Think about what YOU learned in class and write about that. It is okay to keep it simple.'],
      ['revision',   'all', 'Good effort on the other sections! This part needs to be in your own words. Try starting with "I think..." or "In class, we learned..."']
    ];
    sheet.getRange(2, 1, seed.length, 3).setValues(seed);
  }
  return sheet;
}

/**
 * Find existing row for a student + worksheet combo.
 * Returns row number (1-based) or -1 if not found.
 */
function findStudentRow(sheet, name, cls, worksheetId) {
  var data = sheet.getDataRange().getValues();
  var cm = buildColMap(data[0]);
  for (var i = 1; i < data.length; i++) {
    if (colVal(data[i], cm.name) === name &&
        colVal(data[i], cm.class) === cls &&
        colVal(data[i], cm.wsId) === worksheetId) {
      return i + 1;
    }
  }
  return -1;
}


// ===================================================================
//  AUTHENTICATION — Passcode-based via Teachers tab
// ===================================================================

/**
 * Authenticate a teacher by passcode.
 * Returns { authenticated, name, classes[] } or { authenticated: false }.
 *
 * Falls back to legacy TEACHER_KEY if Teachers tab is empty
 * (for backward compatibility during migration).
 */
function authenticateTeacher(passcode) {
  if (!passcode) return { authenticated: false };

  // Try Teachers tab first
  var teacherSheet = getOrCreateTeachersSheet();
  var data = teacherSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === passcode) {
      var classesRaw = (data[i][2] || '').toString();
      var classes = classesRaw.split(',').map(function(c) { return c.trim(); }).filter(Boolean);
      return {
        authenticated: true,
        name: data[i][1] || 'Teacher',
        classes: classes
      };
    }
  }

  // Fallback: legacy TEACHER_KEY from Script Properties
  var props = PropertiesService.getScriptProperties();
  var legacyKey = props.getProperty('TEACHER_KEY');
  if (legacyKey && passcode === legacyKey) {
    return {
      authenticated: true,
      name: 'Teacher (legacy key)',
      classes: []  // no class filtering — sees everything
    };
  }

  return { authenticated: false };
}

/**
 * Quick auth check — returns true/false.
 * Accepts either 'passcode' param (new) or 'key' param (legacy).
 */
function checkAuth(params) {
  var passcode = params.passcode || params.key || '';
  return authenticateTeacher(passcode).authenticated;
}

/**
 * Get authenticated teacher's class list. Empty = no filtering (sees all).
 */
function getTeacherClasses(params) {
  var passcode = params.passcode || params.key || '';
  var auth = authenticateTeacher(passcode);
  return auth.authenticated ? auth.classes : null;
}


// ===================================================================
//  STUDENT-FACING HANDLERS (no auth required)
// ===================================================================

/**
 * POST: Student submits their work.
 * Upserts by (name + class + worksheetId).
 * V2: writes Subject column (defaults to "technology").
 */
function handleSubmit(payload) {
  var sheet = getOrCreateSheet();
  var name = (payload.studentName || '').trim();
  var cls = (payload.studentClass || '').trim();
  var wsId = payload.worksheetId || '';
  var wsTitle = payload.worksheetTitle || '';
  var data = payload.data || {};
  var subject = (payload.subject || 'technology').toLowerCase().trim();

  if (!name || !cls) {
    return { status: 'error', message: 'Name and class are required.' };
  }

  var row = findStudentRow(sheet, name, cls, wsId);
  var now = new Date().toISOString();
  var dataJson = JSON.stringify(data);
  var subjectCol = getSubjectColIndex(sheet) + 1;  // 1-based for Range

  if (row > 0) {
    var currentCount = sheet.getRange(row, 6).getValue() || 0;
    sheet.getRange(row, 1).setValue(now);
    sheet.getRange(row, 6).setValue(currentCount + 1);
    sheet.getRange(row, 7).setValue(dataJson);
    sheet.getRange(row, subjectCol).setValue(subject);
  } else {
    // Build row array — pad to 10 columns
    var newRow = [
      now, name, cls, wsId, wsTitle, 1, dataJson, '', '', subject
    ];
    sheet.appendRow(newRow);
  }

  return { status: 'ok', message: 'Submission received.' };
}

/**
 * GET: Student checks for teacher feedback.
 * Params: name, class, ws (worksheetId)
 * No auth required.
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


// ===================================================================
//  TEACHER-FACING HANDLERS (passcode auth required)
// ===================================================================

/**
 * GET: Teacher login.
 * Params: passcode
 * Returns: { name, classes }
 */
function handleLogin(params) {
  var passcode = params.passcode || '';
  var auth = authenticateTeacher(passcode);

  if (!auth.authenticated) {
    return { status: 'error', message: 'Invalid passcode.' };
  }

  return {
    status: 'ok',
    name: auth.name,
    classes: auth.classes
  };
}

/**
 * GET: List students with parsed summary data.
 * Params: passcode, class (optional), ws (optional), subject (optional)
 *
 * Class filtering: if the teacher's Teachers tab entry lists specific classes,
 * they ONLY see those classes. The 'class' param further narrows within that set.
 */
function handleGetStudents(params) {
  if (!checkAuth(params)) {
    return { status: 'error', message: 'Invalid passcode.' };
  }

  var teacherClasses = getTeacherClasses(params);
  var filterClass = params['class'] || '';
  var filterWs = params.ws || '';
  var filterSubject = (params.subject || '').toLowerCase();

  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var cm = buildColMap(data[0]);
  var students = [];

  // Normalise teacher class list for comparison
  var normalClasses = (teacherClasses || []).map(function(c) { return c.toString().trim().toUpperCase(); });

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowClass = colVal(row, cm.class).toString().trim();
    var rowClassUpper = rowClass.toUpperCase();
    var rowSubject = (colVal(row, cm.subject) || 'technology').toString().toLowerCase();

    // Teacher can only see their assigned classes
    if (normalClasses.length > 0) {
      if (normalClasses.indexOf(rowClassUpper) === -1) continue;
    }
    // Additional filters
    if (filterClass && rowClassUpper !== filterClass.toUpperCase()) continue;
    if (filterWs && colVal(row, cm.wsId) !== filterWs) continue;
    if (filterSubject && rowSubject !== filterSubject) continue;

    // Parse DataJSON for summary — try data column, then wsTitle column (old layout puts JSON there)
    var dataStr = colVal(row, cm.data) || '';
    if (!dataStr && cm.wsTitle >= 0) {
      // In old layout, the DataJSON might be in what we called wsTitle
      var candidate = colVal(row, cm.wsTitle) || '';
      if (candidate.toString().charAt(0) === '{') dataStr = candidate;
    }
    var appData = {};
    try { appData = JSON.parse(dataStr || '{}'); } catch (e) {}

    var summary = buildStudentSummary(appData);

    students.push({
      name: colVal(row, cm.name),
      class: rowClass,
      worksheetId: colVal(row, cm.wsId),
      worksheetTitle: colVal(row, cm.wsTitle),
      submissionCount: colVal(row, cm.count) || 1,
      lastSubmitted: colVal(row, cm.timestamp),
      subject: rowSubject,
      hasFeedback: !!colVal(row, cm.feedback),
      feedbackDate: colVal(row, cm.feedbackDate) || '',
      // Parsed summary fields
      band: summary.band,
      diagnosticScore: summary.diagnosticScore,
      modulesCompleted: summary.modulesCompleted,
      totalModules: summary.totalModules,
      wordCount: summary.totalWordCount,
      vocabScore: summary.vocabScore,
      teelScore: summary.teelScore,
      grammarScore: summary.grammarScore,
      evidenceScore: summary.evidenceScore
    });
  }

  return { status: 'ok', students: students };
}

/**
 * GET: Full detail for one student submission.
 * Params: passcode, name, class, ws
 * Returns full appData + computed metrics.
 */
function handleGetDetail(params) {
  if (!checkAuth(params)) {
    return { status: 'error', message: 'Invalid passcode.' };
  }

  var sheet = getOrCreateSheet();
  var name = (params.name || '').trim();
  var cls = (params['class'] || '').trim();
  var wsId = params.ws || '';

  var data = sheet.getDataRange().getValues();
  var cm = buildColMap(data[0]);

  // Find student row using column map
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (colVal(data[i], cm.name) === name && colVal(data[i], cm.class).toString().trim() === cls) {
      rowIdx = i;
      break;
    }
  }
  if (rowIdx < 0) {
    return { status: 'error', message: 'Student not found.' };
  }

  var rowData = data[rowIdx];

  // Find the DataJSON — scan ALL columns for a JSON blob
  // This is the most reliable approach for both old and new layouts
  var dataStr = '';
  for (var c = 0; c < rowData.length; c++) {
    var cell = (rowData[c] || '').toString();
    if (cell.charAt(0) === '{' && cell.length > 50) { dataStr = cell; break; }
  }
  var appData = {};
  try { appData = JSON.parse(dataStr || '{}'); } catch (e) {}

  var parseOk = !!(appData && appData.band);

  // Parse all writing responses into readable format
  var writings = extractWritings(appData);
  var scores = extractScores(appData);

  return {
    status: 'ok',
    name: name,
    class: cls,
    worksheetId: colVal(rowData, cm.wsId),
    submissionCount: colVal(rowData, cm.count) || 1,
    lastSubmitted: colVal(rowData, cm.timestamp),
    feedbackText: colVal(rowData, cm.feedback) || '',
    feedbackDate: colVal(rowData, cm.feedbackDate) || '',
    // Full data for dashboard to display
    data: appData,
    // Pre-parsed for easy dashboard rendering
    writings: writings,
    scores: scores,
    summary: buildStudentSummary(appData),
    // Debug: remove after confirming it works
    _debug: { dataFound: parseOk, dataLen: dataStr.length, colMap: cm, rowLen: rowData.length }
  };
}

/**
 * GET: Comment bank filtered by subject.
 * Params: passcode, subject (optional, defaults to "all")
 */
function handleGetCommentBank(params) {
  if (!checkAuth(params)) {
    return { status: 'error', message: 'Invalid passcode.' };
  }

  var filterSubject = (params.subject || '').toLowerCase();
  var bankSheet = getOrCreateCommentBankSheet();
  var data = bankSheet.getDataRange().getValues();
  var comments = [];

  for (var i = 1; i < data.length; i++) {
    var category = (data[i][0] || '').toString().toLowerCase();
    var subject = (data[i][1] || 'all').toString().toLowerCase();
    var text = (data[i][2] || '').toString();

    if (!text) continue;
    // Include 'all' comments + subject-specific
    if (filterSubject && subject !== 'all' && subject !== filterSubject) continue;

    comments.push({
      category: category,
      subject: subject,
      text: text
    });
  }

  return { status: 'ok', comments: comments };
}

/**
 * GET: Heatmap data for a class.
 * Params: passcode, class, subject (optional)
 * Returns per-student scores across 4 criteria: vocabulary, teel, grammar, evidence.
 * Score scale: 0 (not attempted), 1 (emerging), 2 (developing), 3 (proficient).
 */
function handleGetHeatmap(params) {
  if (!checkAuth(params)) {
    return { status: 'error', message: 'Invalid passcode.' };
  }

  var teacherClasses = getTeacherClasses(params);
  var filterClass = params['class'] || '';
  var filterSubject = (params.subject || '').toLowerCase();

  if (!filterClass) {
    return { status: 'error', message: 'Class parameter is required for heatmap.' };
  }

  // Verify teacher has access to this class
  if (teacherClasses && teacherClasses.length > 0) {
    if (teacherClasses.indexOf(filterClass) === -1) {
      return { status: 'error', message: 'You do not have access to class ' + filterClass + '.' };
    }
  }

  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var cm = buildColMap(data[0]);
  var heatmap = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (colVal(row, cm.class).toString().trim().toUpperCase() !== filterClass.toUpperCase()) continue;
    var rowSubject = (colVal(row, cm.subject) || 'technology').toString().toLowerCase();
    if (filterSubject && rowSubject !== filterSubject) continue;

    // Find DataJSON
    var dataStr = colVal(row, cm.data) || '';
    if (!dataStr) {
      for (var c = 0; c < row.length; c++) {
        var cell = (row[c] || '').toString();
        if (cell.charAt(0) === '{' && cell.length > 50) { dataStr = cell; break; }
      }
    }
    var appData = {};
    try { appData = JSON.parse(dataStr || '{}'); } catch (e) {}

    var scores = computeHeatmapScores(appData);

    heatmap.push({
      name: colVal(row, cm.name),
      vocabulary: scores.vocabulary,
      teel: scores.teel,
      grammar: scores.grammar,
      evidence: scores.evidence
    });
  }

  // Sort alphabetically by name
  heatmap.sort(function(a, b) { return a.name.localeCompare(b.name); });

  return { status: 'ok', heatmap: heatmap, class: filterClass };
}

/**
 * GET: Export class data as CSV text.
 * Params: passcode, class, subject (optional)
 */
function handleExport(params) {
  if (!checkAuth(params)) {
    return { status: 'error', message: 'Invalid passcode.' };
  }

  var teacherClasses = getTeacherClasses(params);
  var filterClass = params['class'] || '';
  var filterSubject = (params.subject || '').toLowerCase();

  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var subjectIdx = getSubjectColIndex(sheet);

  // CSV header
  var csvRows = ['Name,Class,Band,Modules Completed,Total Modules,Word Count,Vocabulary,TEEL,Grammar,Evidence,Submissions,Last Submitted,Has Feedback'];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowClass = row[2];
    var rowSubject = (row[subjectIdx] || 'technology').toString().toLowerCase();

    if (teacherClasses && teacherClasses.length > 0) {
      if (teacherClasses.indexOf(rowClass) === -1) continue;
    }
    if (filterClass && rowClass !== filterClass) continue;
    if (filterSubject && rowSubject !== filterSubject) continue;

    var appData = {};
    try { appData = JSON.parse(row[6] || '{}'); } catch (e) {}
    var s = buildStudentSummary(appData);

    csvRows.push([
      csvEscape(row[1]),
      csvEscape(rowClass),
      csvEscape(s.band),
      s.modulesCompleted,
      s.totalModules,
      s.totalWordCount,
      s.vocabScore,
      s.teelScore,
      s.grammarScore,
      s.evidenceScore,
      row[5],
      csvEscape(row[0]),
      row[7] ? 'Yes' : 'No'
    ].join(','));
  }

  return { status: 'ok', csv: csvRows.join('\n'), filename: 'write-it-right-' + (filterClass || 'all') + '.csv' };
}

/**
 * POST: Give feedback to a single student.
 * Payload: passcode, studentName, studentClass, worksheetId, feedbackText, criterionRatings (optional)
 */
function handleGiveFeedback(payload) {
  var passcode = payload.passcode || payload.key || '';
  if (!authenticateTeacher(passcode).authenticated) {
    return { status: 'error', message: 'Invalid passcode.' };
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

  // Find or create feedback columns
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var fbCol = headers.indexOf('FeedbackText');
  var fdCol = headers.indexOf('FeedbackDate');
  if (fbCol === -1) {
    fbCol = headers.length;
    sheet.getRange(1, fbCol + 1).setValue('FeedbackText').setFontWeight('bold');
    fdCol = fbCol + 1;
    sheet.getRange(1, fdCol + 1).setValue('FeedbackDate').setFontWeight('bold');
  }
  if (fdCol === -1) {
    fdCol = fbCol + 1;
    sheet.getRange(1, fdCol + 1).setValue('FeedbackDate').setFontWeight('bold');
  }

  // Store feedback with optional revision flag
  var revisionRequired = payload.revisionRequired ? 'yes' : '';
  var feedbackObj = feedback;
  if (revisionRequired) {
    // Prefix with [REVISION] tag so the student client can detect it
    feedbackObj = '[REVISION] ' + feedback;
  }

  sheet.getRange(row, fbCol + 1).setValue(feedbackObj);
  sheet.getRange(row, fdCol + 1).setValue(new Date().toISOString());

  return { status: 'ok', message: revisionRequired ? 'Revision request sent.' : 'Feedback saved.' };
}

/**
 * POST: Give the same feedback to multiple students.
 * Payload: passcode, students: [{ name, class, worksheetId }], feedbackText
 */
function handleBatchFeedback(payload) {
  var passcode = payload.passcode || payload.key || '';
  if (!authenticateTeacher(passcode).authenticated) {
    return { status: 'error', message: 'Invalid passcode.' };
  }

  var sheet = getOrCreateSheet();
  var feedback = (payload.feedbackText || '').trim();
  var targets = payload.students || [];

  if (!feedback || targets.length === 0) {
    return { status: 'error', message: 'Feedback text and at least one student are required.' };
  }

  var now = new Date().toISOString();
  var updated = 0;
  var notFound = [];

  for (var i = 0; i < targets.length; i++) {
    var t = targets[i];
    var name = (t.name || '').trim();
    var cls = (t['class'] || '').trim();
    var wsId = t.worksheetId || '';

    var row = findStudentRow(sheet, name, cls, wsId);
    if (row > 0) {
      sheet.getRange(row, 8).setValue(feedback);
      sheet.getRange(row, 9).setValue(now);
      updated++;
    } else {
      notFound.push(name);
    }
  }

  return {
    status: 'ok',
    message: updated + ' student(s) updated.',
    notFound: notFound
  };
}


// ===================================================================
//  DATA PARSING — Extract human-readable content from DataJSON
// ===================================================================

/**
 * Build a summary object from appData for the student list view.
 * Keeps it lightweight — no full writing text, just metrics.
 */
function buildStudentSummary(appData) {
  if (!appData || typeof appData !== 'object') {
    return {
      band: 'unknown',
      diagnosticScore: 0,
      modulesCompleted: 0,
      totalModules: 8,
      totalWordCount: 0,
      vocabScore: 0,
      teelScore: 0,
      grammarScore: 0,
      evidenceScore: 0
    };
  }

  var band = appData.band || 'unknown';
  var diagScore = appData.diagnosticScore || appData.diagTotalScore || 0;
  var modulesCompleted = countCompleted(appData.moduleCompleted, appData.moduleScores);

  // Aggregate word count from all writing fields
  var totalWords = 0;
  var writingFields = [
    'diagWrite', 'mod1Write', 'mod2Write', 'mod5Compare', 'mod5Analysis',
    'capstoneParagraph', 'teelT', 'teelE1', 'teelE2', 'teelL',
    'sb1compound', 'sb1complex', 'sb2compoundComplex', 'sb3passive',
    'sb3conditional', 'sb3participial', 'sb3appositive', 'sb3inversion'
  ];
  for (var i = 0; i < writingFields.length; i++) {
    var text = appData[writingFields[i]];
    if (text && typeof text === 'string') {
      totalWords += text.trim().split(/\s+/).filter(Boolean).length;
    }
  }

  // Derive criterion scores (0-3 scale) for heatmap compatibility
  var hScores = computeHeatmapScores(appData);

  return {
    band: band,
    diagnosticScore: diagScore,
    modulesCompleted: modulesCompleted,
    totalModules: 8,
    totalWordCount: totalWords,
    vocabScore: hScores.vocabulary,
    teelScore: hScores.teel,
    grammarScore: hScores.grammar,
    evidenceScore: hScores.evidence
  };
}

/**
 * Extract all student writing responses from appData.
 * Returns { label, text } pairs for the detail view.
 */
function extractWritings(appData) {
  if (!appData) return [];

  var writings = [];
  var fields = [
    { key: 'diagWrite', label: 'Diagnostic Writing (Module 0)',
      prompt: 'Write 2 sentences about something you made in Technology. What did you make? What worked well?' },
    { key: 'mod1Write', label: 'Word Power — Use It (Module 1)',
      prompt: 'Write a sentence using at least one key term from the Technology word bank.' },
    { key: 'mod2Write', label: 'Text Types — Mini-Write (Module 2)',
      prompt: 'Write a short Information Report about a topic from Technology class.' },
    { key: 'teelT', label: 'TEEL — Topic Sentence (Module 4)',
      prompt: 'State your main idea. Frame: "[Topic] is important/useful because ___."' },
    { key: 'teelE1', label: 'TEEL — Evidence (Module 4)',
      prompt: 'Give a fact or example. Frame: "For example, in Technology class we ___."' },
    { key: 'teelE2', label: 'TEEL — Explanation (Module 4)',
      prompt: 'Say why your evidence matters. Frame: "This shows that ___ because ___."' },
    { key: 'teelL', label: 'TEEL — Linking Sentence (Module 4)',
      prompt: 'Connect back to your main idea. Frame: "Therefore, [topic] is important because ___."' },
    { key: 'anal1', label: 'Design Analysis — Observe (Module 5)',
      prompt: 'What do you see? Describe the design (a phone stand made from cardboard).' },
    { key: 'anal2', label: 'Design Analysis — Identify (Module 5)',
      prompt: 'What materials and features does it have?' },
    { key: 'anal3', label: 'Design Analysis — Analyse (Module 5)',
      prompt: 'Why was it designed this way? Explain the design choices.' },
    { key: 'anal4', label: 'Design Analysis — Evaluate (Module 5)',
      prompt: 'How effective is it? What are strengths and weaknesses?' },
    { key: 'mod5Compare', label: 'Design Analysis — Compare (Module 5)',
      prompt: 'Compare Design A and Design B. Which is better and why?' },
    { key: 'capPlanT', label: 'Capstone — Plan: Topic (Module 6)',
      prompt: 'Plan your topic sentence for the capstone paragraph.' },
    { key: 'capPlanE1', label: 'Capstone — Plan: Evidence (Module 6)',
      prompt: 'Plan what evidence/example you will use.' },
    { key: 'capPlanE2', label: 'Capstone — Plan: Explanation (Module 6)',
      prompt: 'Plan how you will explain why your evidence matters.' },
    { key: 'capPlanL', label: 'Capstone — Plan: Link (Module 6)',
      prompt: 'Plan your linking/concluding sentence.' },
    { key: 'capstoneParagraph', label: 'Capstone Paragraph (Module 6)',
      prompt: 'Write a full TEEL paragraph using all the skills from Modules 1-5.' },
    { key: 'capGrade', label: 'Self-assessed Grade (Module 6)',
      prompt: 'What grade (A-E) do you think your capstone paragraph deserves?' },
    { key: 'www', label: 'Reflection — What Went Well',
      prompt: 'What part of Write It Right did you enjoy or do well in?' },
    { key: 'ebi', label: 'Reflection — Even Better If',
      prompt: 'What could you improve next time?' },
    { key: 'yns', label: 'Reflection — Your Next Step',
      prompt: 'What is one specific thing you will practise?' },
    // Module 7 sentence building (existing)
    { key: 'sb1compound', label: 'Compound Sentence (Module 7)',
      prompt: 'Turn a simple sentence into a compound sentence using FANBOYS.' },
    { key: 'sb1complex', label: 'Complex Sentence (Module 7)',
      prompt: 'Turn a simple sentence into a complex sentence using because/although/when.' },
    { key: 'sb2compound', label: 'Compound Sentence 2 (Module 7)',
      prompt: 'Build a compound sentence from the given starter.' },
    { key: 'sb2complex', label: 'Complex Sentence 2 (Module 7)',
      prompt: 'Build a complex sentence from the given starter.' },
    { key: 'sb3compound', label: 'Compound Sentence 3 (Module 7)',
      prompt: 'Make "The design failed the test" into a compound sentence.' },
    { key: 'sb3complex', label: 'Complex Sentence 3 (Module 7)',
      prompt: 'Make "The design failed the test" into a complex sentence.' },
    // Module 8 sentence builder
    { key: 'sbL1write1', label: 'Simple Sentence 1 (Module 8)',
      prompt: 'Write a simple sentence using Technology vocabulary.' },
    { key: 'sbL1write2', label: 'Simple Sentence 2 (Module 8)',
      prompt: 'Write another simple sentence using Technology vocabulary.' },
    { key: 'sbL2write1', label: 'Compound — Join (Module 8)',
      prompt: 'Join two simple sentences with a FANBOYS conjunction.' },
    { key: 'sbL2write2', label: 'Compound — Join 2 (Module 8)',
      prompt: 'Join "Cardboard is easy to cut" + "It is not very strong".' },
    { key: 'sbL3write1', label: 'Complex — Because (Module 8)',
      prompt: 'Complete: "We chose recycled cardboard because..."' },
    { key: 'sbL3write2', label: 'Complex — Although (Module 8)',
      prompt: 'Complete: "Although the first prototype failed, ..."' },
    { key: 'sbL3write3', label: 'Complex — If (Module 8)',
      prompt: 'Complete: "If we had used a stronger base, ..."' },
    { key: 'sbL4write1', label: 'Compound-Complex 1 (Module 8)',
      prompt: 'Combine 3 ideas into one compound-complex sentence.' },
    { key: 'sbL4write2', label: 'Compound-Complex 2 (Module 8)',
      prompt: 'Write a compound-complex sentence about sustainability and design criteria.' },
    { key: 'sbL5passive1', label: 'Passive Voice 1 (Module 8)',
      prompt: 'Rewrite "The team measured the bridge" in passive voice.' },
    { key: 'sbL5passive2', label: 'Passive Voice 2 (Module 8)',
      prompt: 'Rewrite "We selected recycled cardboard for the model" in passive voice.' },
    { key: 'sbL5cond1', label: 'Conditional Type 1 (Module 8)',
      prompt: 'Write a Type 1 conditional (real possibility) about your design.' },
    { key: 'sbL5cond2', label: 'Conditional Type 2 (Module 8)',
      prompt: 'Write a Type 2 conditional (hypothetical) about your design.' },
    { key: 'sbL5cond3', label: 'Conditional Type 3 (Module 8)',
      prompt: 'Write a Type 3 conditional (past hypothetical) about your design.' },
    { key: 'sbL5part', label: 'Participial Phrase (Module 8)',
      prompt: 'Write a sentence starting with a participial phrase (-ing word).' },
    { key: 'sbL5appos', label: 'Appositive (Module 8)',
      prompt: 'Write a sentence with an appositive that defines a technical term.' }
  ];

  // Responses can be at appData.responses.KEY or appData.KEY (depends on version)
  var responses = appData.responses || {};

  for (var i = 0; i < fields.length; i++) {
    var val = responses[fields[i].key] || appData[fields[i].key];
    if (val !== undefined && val !== null && val !== '') {
      writings.push({
        key: fields[i].key,
        label: fields[i].label,
        prompt: fields[i].prompt || '',
        text: String(val),
        wordCount: typeof val === 'string' ? val.trim().split(/\s+/).filter(Boolean).length : 0
      });
    }
  }

  return writings;
}

/**
 * Extract quiz/activity scores from appData.
 * Returns { label, score, maxScore } list.
 */
function extractScores(appData) {
  if (!appData) return [];

  var scores = [];

  // Module scores are at appData.moduleScores: { "0": 10, "1": 5, ... }
  var ms = appData.moduleScores || {};
  var moduleNames = [
    'Diagnostic', 'Word Power', 'Text Types', 'Read the Question',
    'Build a Paragraph', 'Analyse a Design', 'Put It All Together',
    'Grammar Gym', 'Sentence Builder'
  ];
  for (var m = 0; m <= 8; m++) {
    if (ms[m] !== undefined && ms[m] !== null) {
      scores.push({
        key: 'module' + m,
        label: 'Module ' + m + ' — ' + (moduleNames[m] || ''),
        score: Number(ms[m]) || 0,
        maxScore: m === 0 ? 12 : 15
      });
    }
  }

  // Quiz scores from appData.quizState
  var qs = appData.quizState || {};
  if (qs.sortScore !== null && qs.sortScore !== undefined) {
    scores.push({ key: 'wordSort', label: 'Word Sort', score: qs.sortScore, maxScore: qs.sortTotal || 12 });
  }
  if (qs.jigsawScore !== null && qs.jigsawScore !== undefined) {
    scores.push({ key: 'jigsawScore', label: 'TEEL Jigsaw', score: qs.jigsawScore, maxScore: 4 });
  }

  // Diagnostic score
  if (appData.diagnosticScore !== undefined) {
    scores.push({ key: 'diagnostic', label: 'Diagnostic Total', score: appData.diagnosticScore, maxScore: 12 });
  }

  return scores;
}

/**
 * Compute heatmap criterion scores from appData.
 * Scale: 0 = not attempted, 1 = emerging, 2 = developing, 3 = proficient.
 */
function computeHeatmapScores(appData) {
  if (!appData) return { vocabulary: 0, teel: 0, grammar: 0, evidence: 0 };

  // Vocabulary: based on word sort + gap fill scores
  var vocabRaw = (appData.mod1SortScore || 0) + (appData.mod1GapScore || 0);
  var vocabMax = 16;  // 12 + 4
  var vocabulary = vocabMax > 0 ? Math.round((vocabRaw / vocabMax) * 3) : 0;

  // TEEL: based on TEEL paragraph completeness
  var teelParts = 0;
  if (appData.teelT) teelParts++;
  if (appData.teelE1) teelParts++;
  if (appData.teelE2) teelParts++;
  if (appData.teelL) teelParts++;
  var capstonLen = (appData.capstoneParagraph || '').trim().split(/\s+/).filter(Boolean).length;
  var teel = 0;
  if (teelParts === 0 && capstonLen === 0) teel = 0;
  else if (teelParts <= 2 || capstonLen < 20) teel = 1;
  else if (teelParts <= 3 || capstonLen < 40) teel = 2;
  else teel = 3;

  // Grammar: based on tense quiz + sentence type quiz
  var grammarRaw = (appData.tenseQuizScore || 0) + (appData.sentenceTypeScore || 0);
  var grammarMax = 9;  // 4 + 5
  var grammar = grammarMax > 0 ? Math.round((grammarRaw / grammarMax) * 3) : 0;

  // Evidence: based on analysis responses + compare paragraph
  var evidenceParts = 0;
  if (appData.mod5Analysis) evidenceParts++;
  if (appData.mod5Compare) evidenceParts++;
  var evidenceLen = 0;
  if (appData.mod5Compare) evidenceLen = appData.mod5Compare.trim().split(/\s+/).filter(Boolean).length;
  var evidence = 0;
  if (evidenceParts === 0) evidence = 0;
  else if (evidenceParts === 1 && evidenceLen < 20) evidence = 1;
  else if (evidenceLen < 40) evidence = 2;
  else evidence = 3;

  return {
    vocabulary: Math.min(vocabulary, 3),
    teel: Math.min(teel, 3),
    grammar: Math.min(grammar, 3),
    evidence: Math.min(evidence, 3)
  };
}


// ===================================================================
//  UTILITIES
// ===================================================================

function countCompleted(moduleCompleted, moduleScores) {
  if (!moduleCompleted) return 0;
  var count = 0;
  for (var key in moduleCompleted) {
    if (moduleCompleted[key] !== true) continue;
    // Module 0 (diagnostic) counts if flagged complete
    if (key === '0') { count++; continue; }
    // All other modules: only count if they have a real score >= 1
    // This prevents "click-through" students from showing false progress
    if (moduleScores && moduleScores[key] && Number(moduleScores[key]) >= 1) {
      count++;
    }
  }
  return count;
}

/**
 * Escape a value for CSV output.
 */
function csvEscape(val) {
  if (val === null || val === undefined) return '';
  var str = String(val);
  if (str.indexOf(',') >= 0 || str.indexOf('"') >= 0 || str.indexOf('\n') >= 0) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * TEMPORARY DIAGNOSTIC — remove after debugging.
 * Returns raw sheet info so we can see exactly what's in the data.
 */
function handleDiag(params) {
  if (!checkAuth(params)) {
    return { status: 'error', message: 'Invalid passcode.' };
  }
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0] || [];
  var teacherClasses = getTeacherClasses(params);
  var normalClasses = (teacherClasses || []).map(function(c) { return c.toString().trim().toUpperCase(); });

  // Sample first 3 data rows
  var samples = [];
  for (var i = 1; i < Math.min(data.length, 4); i++) {
    var row = data[i];
    samples.push({
      rowIndex: i,
      colA_timestamp: String(row[0]),
      colB_name: String(row[1]),
      colC_class: String(row[2]),
      colC_type: typeof row[2],
      colC_trimUpper: (row[2] || '').toString().trim().toUpperCase(),
      colD_wsId: String(row[3]),
      colE_wsTitle: String(row[4]),
      colF_count: String(row[5]),
      colG_dataLen: String(row[6] || '').length,
      matchesTeacher: normalClasses.indexOf((row[2] || '').toString().trim().toUpperCase()) >= 0
    });
  }

  return {
    status: 'ok',
    totalRows: data.length - 1,
    headers: headers.map(String),
    headerCount: headers.length,
    teacherClasses: teacherClasses,
    normalClasses: normalClasses,
    samples: samples
  };
}
