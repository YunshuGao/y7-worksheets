/**
 * submission.js — Student Submission & Teacher Feedback Module
 *
 * Include in any Y7 worksheet to enable:
 *   - "Submit to Teacher" button
 *   - Submission status tracking
 *   - Teacher feedback display
 *
 * Usage:
 *   <script src="submission.js"></script>
 *   <script>initSubmission({ storageKey: 'y7wir-data', worksheetTitle: 'Write It Right' });</script>
 *
 * Privacy: Data sent ONLY to teacher's Google Sheet via Apps Script.
 *          No third-party services. No student email addresses collected.
 */

(function () {
  'use strict';

  // ===== CONFIGURATION =====
  // Teacher: replace this URL after deploying Google Apps Script
  const APPS_SCRIPT_URL = '';  // e.g. 'https://script.google.com/macros/s/XXXXX/exec'

  // ===== MODULE STATE =====
  let config = {};
  let submissionMeta = {};  // cached locally: { lastSubmitted, submissionCount }

  // ===== INITIALISATION =====
  window.initSubmission = function (opts) {
    config = {
      storageKey: opts.storageKey || 'worksheet-data',
      worksheetTitle: opts.worksheetTitle || 'Worksheet',
      getStudentName: opts.getStudentName || function () {
        var el = document.getElementById('f-name');
        return el ? el.value.trim() : '';
      },
      getStudentClass: opts.getStudentClass || function () {
        var el = document.getElementById('f-class');
        return el ? el.value.trim() : '';
      },
      getData: opts.getData || function () {
        try { return JSON.parse(localStorage.getItem(config.storageKey)); }
        catch (e) { return null; }
      },
      containerSelector: opts.containerSelector || '.btn-bar',
      insertMode: opts.insertMode || 'append'  // 'append' = inside container, 'after' = after container
    };

    // Load cached submission metadata
    var metaKey = config.storageKey + '-submission-meta';
    try {
      var saved = localStorage.getItem(metaKey);
      if (saved) submissionMeta = JSON.parse(saved);
    } catch (e) { submissionMeta = {}; }

    // Inject CSS
    injectStyles();

    // Inject UI elements
    injectSubmitButton();
    injectStatusBadge();
    injectFeedbackPanel();

    // Cache student name across worksheets
    restoreCachedName();

    // Check for feedback on load (non-blocking)
    if (APPS_SCRIPT_URL && config.getStudentName()) {
      checkForFeedback();
    }

    // Retry any pending submissions
    retryPendingSubmission();

    // Listen for online/offline
    window.addEventListener('online', function () { updateOnlineStatus(true); });
    window.addEventListener('offline', function () { updateOnlineStatus(false); });
    updateOnlineStatus(navigator.onLine);
  };

  // ===== STYLES =====
  function injectStyles() {
    var style = document.createElement('style');
    style.textContent = [
      '.sub-btn { background: #1A8A6E; color: #fff; border: 2px solid #16755d; border-radius: 6px;',
      '  padding: 6px 14px; font-size: 13px; font-weight: 600; cursor: pointer; display: inline-flex;',
      '  align-items: center; gap: 6px; transition: all 0.2s; font-family: inherit; }',
      '.sub-btn:hover { background: #16755d; transform: translateY(-1px); }',
      '.sub-btn:disabled { background: #999; border-color: #888; cursor: not-allowed; transform: none; }',
      '.sub-btn .spinner { display: none; width: 14px; height: 14px; border: 2px solid #fff;',
      '  border-top-color: transparent; border-radius: 50%; animation: sub-spin 0.6s linear infinite; }',
      '.sub-btn.loading .spinner { display: inline-block; }',
      '.sub-btn.loading .btn-label { display: none; }',
      '@keyframes sub-spin { to { transform: rotate(360deg); } }',
      '',
      '.sub-status { font-size: 11px; color: #666; padding: 4px 8px; display: inline-block; }',
      '.sub-status.has-submitted { color: #1A8A6E; }',
      '',
      '.sub-feedback-panel { background: #f0faf7; border-left: 4px solid #1A8A6E; border-radius: 0 8px 8px 0;',
      '  padding: 14px 18px; margin: 10px 0; display: none; position: relative; }',
      '.sub-feedback-panel.visible { display: block; }',
      '.sub-feedback-panel h4 { margin: 0 0 8px; color: #1A8A6E; font-size: 14px; }',
      '.sub-feedback-panel p { margin: 0; font-size: 13px; line-height: 1.5; color: #333; }',
      '.sub-feedback-panel .fb-date { font-size: 11px; color: #888; margin-top: 6px; }',
      '.sub-feedback-badge { background: #E8712B; color: #fff; font-size: 10px; padding: 2px 6px;',
      '  border-radius: 10px; margin-left: 6px; animation: sub-pulse 1.5s ease-in-out 3; }',
      '@keyframes sub-pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.5; } }',
      '',
      '.sub-toast { position: fixed; bottom: 20px; right: 20px; background: #1A8A6E; color: #fff;',
      '  padding: 12px 20px; border-radius: 8px; font-size: 13px; z-index: 9999;',
      '  transform: translateY(100px); opacity: 0; transition: all 0.3s ease; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }',
      '.sub-toast.show { transform: translateY(0); opacity: 1; }',
      '.sub-toast.error { background: #c0392b; }',
      '',
      '.sub-privacy { font-size: 10px; color: #999; text-align: center; padding: 4px; }',
      '',
      '.sub-offline-dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block;',
      '  margin-right: 4px; vertical-align: middle; }',
      '.sub-offline-dot.online { background: #27ae60; }',
      '.sub-offline-dot.offline { background: #e74c3c; }',
      '',
      '@media print { .sub-btn, .sub-status, .sub-feedback-panel, .sub-toast, .sub-privacy { display: none !important; } }'
    ].join('\n');
    document.head.appendChild(style);
  }

  // ===== UI INJECTION =====
  function injectSubmitButton() {
    var container = document.querySelector(config.containerSelector);
    if (!container) return;

    var btn = document.createElement('button');
    btn.className = 'sub-btn';
    btn.id = 'subSubmitBtn';
    btn.title = 'Submit your work to Yunshu Gao(Ms Gao)';
    btn.innerHTML = '<span class="btn-label">\uD83D\uDCE4 Submit to Teacher</span><span class="spinner"></span>';
    btn.onclick = handleSubmit;

    if (config.insertMode === 'append') {
      container.appendChild(btn);
    } else {
      container.parentNode.insertBefore(btn, container.nextSibling);
    }
  }

  function injectStatusBadge() {
    var btn = document.getElementById('subSubmitBtn');
    if (!btn) return;

    var badge = document.createElement('span');
    badge.className = 'sub-status';
    badge.id = 'subStatusBadge';
    btn.parentNode.insertBefore(badge, btn.nextSibling);
    updateStatusBadge();
  }

  function injectFeedbackPanel() {
    // Insert after the topbar/btn-bar area, before worksheet content
    var container = document.querySelector(config.containerSelector);
    if (!container) return;

    var panel = document.createElement('div');
    panel.className = 'sub-feedback-panel';
    panel.id = 'subFeedbackPanel';
    panel.innerHTML = '<h4>\uD83D\uDCAC Teacher Feedback</h4>' +
      '<p id="subFeedbackText"></p>' +
      '<div class="fb-date" id="subFeedbackDate"></div>';

    // Insert after the container's parent section (topbar or btn-bar)
    var ws = document.querySelector('.worksheet, #ws, .tab-bar, .topbar');
    if (ws) {
      ws.parentNode.insertBefore(panel, ws.nextSibling);
    } else {
      container.parentNode.insertBefore(panel, container.nextSibling);
    }
  }

  // ===== SUBMISSION =====
  function handleSubmit() {
    var name = config.getStudentName();
    var cls = config.getStudentClass();

    if (!name) {
      showToast('Please enter your name before submitting.', 'error');
      return;
    }
    if (!cls) {
      showToast('Please select your class before submitting.', 'error');
      return;
    }

    if (!APPS_SCRIPT_URL) {
      showToast('Submission not configured yet. Your work is saved locally.', 'error');
      return;
    }

    // Show privacy notice on first submit
    if (!submissionMeta.privacyAcknowledged) {
      if (!confirm('Privacy Notice: Your name and answers will be sent to your teacher\'s Google account. No other service receives your data.\n\nClick OK to submit.')) {
        return;
      }
      submissionMeta.privacyAcknowledged = true;
      saveSubmissionMeta();
    }

    var btn = document.getElementById('subSubmitBtn');
    btn.classList.add('loading');
    btn.disabled = true;

    var data = config.getData();
    var payload = {
      action: 'submit',
      studentName: name,
      studentClass: cls,
      worksheetId: config.storageKey,
      worksheetTitle: config.worksheetTitle,
      data: data
    };

    // Cache student name for other worksheets
    localStorage.setItem('y7-student-name', name);
    localStorage.setItem('y7-student-class', cls);

    sendToAppsScript(payload)
      .then(function (result) {
        btn.classList.remove('loading');
        btn.disabled = false;
        submissionMeta.lastSubmitted = new Date().toISOString();
        submissionMeta.submissionCount = (submissionMeta.submissionCount || 0) + 1;
        saveSubmissionMeta();
        updateStatusBadge();
        showToast('Submitted! Your teacher can now see your work. (' +
          submissionMeta.submissionCount + ' submission' +
          (submissionMeta.submissionCount > 1 ? 's' : '') + ')');
        // Clear any pending submission
        localStorage.removeItem(config.storageKey + '-pending-submission');
      })
      .catch(function (err) {
        btn.classList.remove('loading');
        btn.disabled = false;
        // Queue for retry
        localStorage.setItem(config.storageKey + '-pending-submission', JSON.stringify(payload));
        showToast('Could not submit right now \u2014 saved for retry. Check your internet connection.', 'error');
        console.warn('Submission error:', err);
      });
  }

  // ===== FEEDBACK =====
  function checkForFeedback() {
    var name = config.getStudentName();
    var cls = config.getStudentClass();
    if (!name || !cls || !APPS_SCRIPT_URL) return;

    var url = APPS_SCRIPT_URL + '?action=feedback' +
      '&name=' + encodeURIComponent(name) +
      '&class=' + encodeURIComponent(cls) +
      '&ws=' + encodeURIComponent(config.storageKey);

    fetch(url, { method: 'GET' })
      .then(function (r) { return r.json(); })
      .then(function (result) {
        if (result && result.hasFeedback) {
          showFeedback(result.feedbackText, result.feedbackDate);
        }
      })
      .catch(function () { /* silently fail — feedback check is non-critical */ });
  }

  function showFeedback(text, date) {
    var panel = document.getElementById('subFeedbackPanel');
    var textEl = document.getElementById('subFeedbackText');
    var dateEl = document.getElementById('subFeedbackDate');
    if (!panel || !textEl) return;

    textEl.textContent = text;
    if (dateEl && date) {
      dateEl.textContent = 'Feedback given: ' + formatDate(date);
    }
    panel.classList.add('visible');

    // Show badge on submit button
    var btn = document.getElementById('subSubmitBtn');
    if (btn && !document.getElementById('subFbBadge')) {
      var badge = document.createElement('span');
      badge.className = 'sub-feedback-badge';
      badge.id = 'subFbBadge';
      badge.textContent = 'New feedback!';
      btn.parentNode.insertBefore(badge, btn.nextSibling);
    }
  }

  // ===== APPS SCRIPT COMMUNICATION =====
  function sendToAppsScript(payload) {
    return fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },  // Apps Script needs text/plain for CORS
      body: JSON.stringify(payload)
    })
    .then(function (response) {
      if (!response.ok) throw new Error('HTTP ' + response.status);
      return response.json();
    })
    .then(function (result) {
      if (result.status !== 'ok') throw new Error(result.message || 'Unknown error');
      return result;
    });
  }

  // ===== RETRY PENDING =====
  function retryPendingSubmission() {
    var pendingKey = config.storageKey + '-pending-submission';
    var pending = localStorage.getItem(pendingKey);
    if (!pending || !APPS_SCRIPT_URL || !navigator.onLine) return;

    try {
      var payload = JSON.parse(pending);
      sendToAppsScript(payload)
        .then(function () {
          localStorage.removeItem(pendingKey);
          submissionMeta.lastSubmitted = new Date().toISOString();
          submissionMeta.submissionCount = (submissionMeta.submissionCount || 0) + 1;
          saveSubmissionMeta();
          updateStatusBadge();
          showToast('Previous submission sent successfully!');
        })
        .catch(function () { /* keep pending for next retry */ });
    } catch (e) {
      localStorage.removeItem(pendingKey);
    }
  }

  // ===== HELPERS =====
  function saveSubmissionMeta() {
    localStorage.setItem(config.storageKey + '-submission-meta', JSON.stringify(submissionMeta));
  }

  function updateStatusBadge() {
    var badge = document.getElementById('subStatusBadge');
    if (!badge) return;

    if (submissionMeta.lastSubmitted) {
      badge.className = 'sub-status has-submitted';
      badge.innerHTML = '<span class="sub-offline-dot online"></span>' +
        'Submitted ' + submissionMeta.submissionCount + ' time' +
        (submissionMeta.submissionCount > 1 ? 's' : '') +
        ' \u2014 last: ' + formatDate(submissionMeta.lastSubmitted);
    } else {
      badge.textContent = 'Not yet submitted';
    }
  }

  function updateOnlineStatus(isOnline) {
    var btn = document.getElementById('subSubmitBtn');
    if (!btn) return;

    if (isOnline) {
      btn.disabled = false;
      btn.title = 'Submit your work to Yunshu Gao(Ms Gao)';
    } else {
      btn.disabled = true;
      btn.title = 'You are offline \u2014 your work is saved locally';
    }

    // Update dot indicator
    var dot = document.querySelector('.sub-offline-dot');
    if (dot) {
      dot.className = 'sub-offline-dot ' + (isOnline ? 'online' : 'offline');
    }
  }

  function restoreCachedName() {
    // Auto-fill student name from cross-worksheet cache if current field is empty
    var nameEl = null;
    try { nameEl = config.getStudentName ? null : null; } catch (e) {}

    var cachedName = localStorage.getItem('y7-student-name');
    var cachedClass = localStorage.getItem('y7-student-class');

    if (cachedName) {
      var nameInput = document.getElementById('f-name') || document.getElementById('studentName');
      if (nameInput && !nameInput.value.trim()) {
        nameInput.value = cachedName;
        // Trigger save if the worksheet has a saveAll function
        if (typeof window.saveAll === 'function') window.saveAll();
      }
    }

    if (cachedClass) {
      var classInput = document.getElementById('f-class') || document.getElementById('studentClass');
      if (classInput && !classInput.value) {
        classInput.value = cachedClass;
        if (typeof window.saveAll === 'function') window.saveAll();
      }
    }
  }

  function formatDate(iso) {
    try {
      var d = new Date(iso);
      return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short' }) +
        ', ' + d.toLocaleTimeString('en-AU', { hour: 'numeric', minute: '2-digit' });
    } catch (e) {
      return iso;
    }
  }

  function showToast(msg, type) {
    // Remove existing toast
    var old = document.querySelector('.sub-toast');
    if (old) old.remove();

    var toast = document.createElement('div');
    toast.className = 'sub-toast' + (type === 'error' ? ' error' : '');
    toast.textContent = msg;
    document.body.appendChild(toast);

    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        toast.classList.add('show');
      });
    });

    setTimeout(function () {
      toast.classList.remove('show');
      setTimeout(function () { toast.remove(); }, 300);
    }, 4000);
  }

})();
