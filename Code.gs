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
 * Get the column index (0-based) for Subject in the Submissions sheet.
 * Returns 9 for new sheets (column J), or finds it dynamically for migrated sheets.
 */
function getSubjectColIndex(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = headers.indexOf('Subject');
  return idx >= 0 ? idx : 9;  // default to column J (index 9)
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
      ['general',    'all', 'Please complete the remaining modules before the end of term.']
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
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === name && data[i][2] === cls && data[i][3] === worksheetId) {
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
  var subjectIdx = getSubjectColIndex(sheet);
  var students = [];

  // Normalise teacher class list for comparison
  var normalClasses = (teacherClasses || []).map(function(c) { return c.toString().trim().toUpperCase(); });

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowClass = (row[2] || '').toString().trim();
    var rowClassUpper = rowClass.toUpperCase();
    var rowSubject = (row[subjectIdx] || 'technology').toString().toLowerCase();

    // Teacher can only see their assigned classes
    if (normalClasses.length > 0) {
      if (normalClasses.indexOf(rowClassUpper) === -1) continue;
    }
    // Additional filters
    if (filterClass && rowClassUpper !== filterClass.toUpperCase()) continue;
    if (filterWs && row[3] !== filterWs) continue;
    if (filterSubject && rowSubject !== filterSubject) continue;

    // Parse DataJSON for summary
    var appData = {};
    try { appData = JSON.parse(row[6] || '{}'); } catch (e) {}

    var summary = buildStudentSummary(appData);

    students.push({
      name: row[1],
      class: rowClass,
      worksheetId: row[3],
      worksheetTitle: row[4],
      submissionCount: row[5],
      lastSubmitted: row[0],
      subject: rowSubject,
      hasFeedback: !!row[7],
      feedbackDate: row[8] || '',
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

  var row = findStudentRow(sheet, name, cls, wsId);
  if (row < 0) {
    return { status: 'error', message: 'Student not found.' };
  }

  var dataJson = sheet.getRange(row, 7).getValue();
  var appData = {};
  try { appData = JSON.parse(dataJson || '{}'); } catch (e) {}

  // Parse all writing responses into readable format
  var writings = extractWritings(appData);
  var scores = extractScores(appData);

  return {
    status: 'ok',
    name: name,
    class: cls,
    worksheetId: wsId,
    submissionCount: sheet.getRange(row, 6).getValue(),
    lastSubmitted: sheet.getRange(row, 1).getValue(),
    feedbackText: sheet.getRange(row, 8).getValue() || '',
    feedbackDate: sheet.getRange(row, 9).getValue() || '',
    // Full data for dashboard to display
    data: appData,
    // Pre-parsed for easy dashboard rendering
    writings: writings,
    scores: scores,
    summary: buildStudentSummary(appData)
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
  var subjectIdx = getSubjectColIndex(sheet);
  var heatmap = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[2] !== filterClass) continue;
    var rowSubject = (row[subjectIdx] || 'technology').toString().toLowerCase();
    if (filterSubject && rowSubject !== filterSubject) continue;

    var appData = {};
    try { appData = JSON.parse(row[6] || '{}'); } catch (e) {}

    var scores = computeHeatmapScores(appData);

    heatmap.push({
      name: row[1],
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

  sheet.getRange(row, 8).setValue(feedback);
  sheet.getRange(row, 9).setValue(new Date().toISOString());

  return { status: 'ok', message: 'Feedback saved.' };
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
  var modulesCompleted = countCompleted(appData.moduleCompleted);

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
    { key: 'diagWrite', label: 'Diagnostic Writing (Module 0)' },
    { key: 'mod1Write', label: 'Word Power — Use It Challenge (Module 1)' },
    { key: 'mod2Write', label: 'Text Types — Mini-Write (Module 2)' },
    { key: 'teelT', label: 'TEEL — Topic Sentence (Module 4)' },
    { key: 'teelE1', label: 'TEEL — Evidence (Module 4)' },
    { key: 'teelE2', label: 'TEEL — Explanation (Module 4)' },
    { key: 'teelL', label: 'TEEL — Linking Sentence (Module 4)' },
    { key: 'mod5Analysis', label: 'Design Analysis — Guided (Module 5)' },
    { key: 'mod5Compare', label: 'Design Analysis — Compare (Module 5)' },
    { key: 'capstoneParagraph', label: 'Capstone Paragraph (Module 6)' },
    { key: 'capstoneGrade', label: 'Self-assessed Grade (Module 6)' },
    { key: 'capstoneJustification', label: 'Grade Justification (Module 6)' },
    { key: 'www', label: 'Reflection — What Went Well' },
    { key: 'ebi', label: 'Reflection — Even Better If' },
    { key: 'yns', label: 'Reflection — Your Next Step' },
    // Sentence building fields
    { key: 'sb1compound', label: 'Sentence Building — Compound' },
    { key: 'sb1complex', label: 'Sentence Building — Complex' },
    { key: 'sb2compoundComplex', label: 'Sentence Building — Compound-Complex' },
    { key: 'sb3passive', label: 'Sentence Building — Passive Voice' },
    { key: 'sb3conditional', label: 'Sentence Building — Conditionals' },
    { key: 'sb3participial', label: 'Sentence Building — Participial Phrases' },
    { key: 'sb3appositive', label: 'Sentence Building — Appositives' },
    { key: 'sb3inversion', label: 'Sentence Building — Inversion' },
    // Verb tenses
    { key: 'tenseT1', label: 'Verb Tense — Task 1' },
    { key: 'tenseT2', label: 'Verb Tense — Task 2' },
    { key: 'tenseT3', label: 'Verb Tense — Task 3' },
    { key: 'tenseT4', label: 'Verb Tense — Task 4' }
  ];

  for (var i = 0; i < fields.length; i++) {
    var val = appData[fields[i].key];
    if (val !== undefined && val !== null && val !== '') {
      writings.push({
        key: fields[i].key,
        label: fields[i].label,
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
  var fields = [
    { key: 'diagWordScore', label: 'Diagnostic — Word Match', max: 6 },
    { key: 'diagSentenceScore', label: 'Diagnostic — Sentence Check', max: 3 },
    { key: 'diagWriteScore', label: 'Diagnostic — Writing', max: 3 },
    { key: 'diagTotalScore', label: 'Diagnostic — Total', max: 12 },
    { key: 'mod1SortScore', label: 'Module 1 — Word Sort', max: 12 },
    { key: 'mod1GapScore', label: 'Module 1 — Gap Fill', max: 4 },
    { key: 'mod1WriteScore', label: 'Module 1 — Writing', max: 2 },
    { key: 'mod2TypeScore', label: 'Module 2 — Text Type ID', max: 4 },
    { key: 'mod2FeatureScore', label: 'Module 2 — Feature Spotter', max: 4 },
    { key: 'mod3VSCCScore', label: 'Module 3 — VSCC Highlighter', max: 3 },
    { key: 'mod3ALARMScore', label: 'Module 3 — ALARM Quiz', max: 4 },
    { key: 'sentenceTypeScore', label: 'Sentence Type Quiz', max: 5 },
    { key: 'tenseQuizScore', label: 'Verb Tense Quiz', max: 4 }
  ];

  for (var i = 0; i < fields.length; i++) {
    var val = appData[fields[i].key];
    if (val !== undefined && val !== null) {
      scores.push({
        key: fields[i].key,
        label: fields[i].label,
        score: Number(val) || 0,
        maxScore: fields[i].max
      });
    }
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

function countCompleted(moduleCompleted) {
  if (!moduleCompleted) return 0;
  var count = 0;
  for (var key in moduleCompleted) {
    if (moduleCompleted[key] === true) count++;
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
