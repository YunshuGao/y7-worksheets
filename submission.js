/**
 * submission.js — Student Submission & Feedback Coach
 *
 * VERSION 2.0 — Redesigned feedback experience
 *
 * Features:
 *   - "Submit to Teacher" button with offline retry
 *   - "Feedback Coach" floating widget (bottom-right)
 *     → new feedback badge → expand to read → checklist → resubmit prompt
 *   - Cross-worksheet name caching
 *
 * Privacy: Data sent ONLY to teacher's Google Sheet via Apps Script.
 */

(function () {
  'use strict';

  // ===== CONFIGURATION =====
  var APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwq7dx3INqCCVukx2Xvufi8ak_d-WwwExG-R1zlGO2-JE2bTFGX66w2LaGVskAuF6_K6g/exec';

  // ===== STATE =====
  var config = {};
  var submissionMeta = {};
  var feedbackState = 'none'; // 'none' | 'new' | 'read' | 'working' | 'resubmitted'
  var feedbackData = null;    // { text, date }

  // ===== INIT =====
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
      insertMode: opts.insertMode || 'append'
    };

    var metaKey = config.storageKey + '-submission-meta';
    try {
      var saved = localStorage.getItem(metaKey);
      if (saved) submissionMeta = JSON.parse(saved);
    } catch (e) { submissionMeta = {}; }

    // Restore feedback state from this session
    try {
      var fbState = sessionStorage.getItem(config.storageKey + '-fb-state');
      if (fbState) feedbackState = fbState;
    } catch (e) {}

    injectStyles();
    injectSubmitButton();
    injectStatusBadge();
    injectFeedbackCoach();
    restoreCachedName();

    if (APPS_SCRIPT_URL && config.getStudentName()) {
      checkForFeedback();
    }

    retryPendingSubmission();

    window.addEventListener('online', function () { updateOnlineStatus(true); retryPendingSubmission(); });
    window.addEventListener('offline', function () { updateOnlineStatus(false); });
    updateOnlineStatus(navigator.onLine);

    // Re-check feedback when student fills in name (may not be available at init time)
    var _feedbackChecked = !!(APPS_SCRIPT_URL && config.getStudentName());
    if (!_feedbackChecked) {
      var _nameEl = document.getElementById('f-name') || document.getElementById('studentName');
      var _classEl = document.getElementById('f-class') || document.getElementById('studentClass');
      function _tryFeedbackCheck() {
        if (_feedbackChecked) return;
        if (config.getStudentName() && config.getStudentClass()) {
          _feedbackChecked = true;
          checkForFeedback();
        }
      }
      if (_nameEl) _nameEl.addEventListener('blur', _tryFeedbackCheck);
      if (_classEl) _classEl.addEventListener('change', _tryFeedbackCheck);
    }
  };

  // ===== STYLES =====
  function injectStyles() {
    var style = document.createElement('style');
    style.textContent = [
      /* Submit button */
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
      '.sub-status { font-size: 11px; color: #666; padding: 4px 8px; display: inline-block; }',
      '.sub-status.has-submitted { color: #1A8A6E; }',
      '.sub-offline-dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block;',
      '  margin-right: 4px; vertical-align: middle; }',
      '.sub-offline-dot.online { background: #27ae60; }',
      '.sub-offline-dot.offline { background: #e74c3c; }',
      '',
      '/* ===== FEEDBACK COACH WIDGET ===== */',
      '.fb-coach { position: fixed; bottom: 20px; right: 20px; z-index: 9998;',
      '  font-family: "Segoe UI", system-ui, sans-serif; }',
      '',
      /* Floating trigger button */
      '.fb-trigger { display: none; align-items: center; gap: 8px; padding: 10px 16px;',
      '  border-radius: 24px; border: none; cursor: pointer; font-size: 14px;',
      '  font-weight: 600; font-family: inherit; box-shadow: 0 4px 16px rgba(0,0,0,0.15);',
      '  transition: all 0.2s; }',
      '.fb-trigger:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.2); }',
      '.fb-trigger.state-new { display: flex; background: #e67e22; color: #fff;',
      '  animation: fb-bounce 1s ease-in-out 3; }',
      '.fb-trigger.state-working { display: flex; background: #2980b9; color: #fff; }',
      '.fb-trigger.state-resubmitted { display: flex; background: #27ae60; color: #fff; }',
      '@keyframes fb-bounce { 0%,100% { transform: translateY(0); }',
      '  50% { transform: translateY(-6px); } }',
      '',
      /* Expanded card */
      '.fb-card { display: none; position: fixed; bottom: 20px; right: 20px;',
      '  width: 340px; max-height: 80vh; background: #fff; border-radius: 16px;',
      '  box-shadow: 0 8px 32px rgba(0,0,0,0.18); overflow: hidden;',
      '  flex-direction: column; z-index: 9999; animation: fb-slideUp 0.25s ease-out; }',
      '.fb-card.visible { display: flex; }',
      '@keyframes fb-slideUp { from { opacity: 0; transform: translateY(20px); }',
      '  to { opacity: 1; transform: translateY(0); } }',
      '',
      /* Card header */
      '.fb-card-header { background: linear-gradient(135deg, #1a3d5c, #2c5f8a); color: #fff;',
      '  padding: 14px 16px; display: flex; align-items: center; justify-content: space-between; }',
      '.fb-card-header h3 { font-size: 14px; margin: 0; display: flex; align-items: center; gap: 6px; }',
      '.fb-card-close { background: none; border: none; color: rgba(255,255,255,0.7);',
      '  font-size: 20px; cursor: pointer; padding: 0 4px; }',
      '.fb-card-close:hover { color: #fff; }',
      '',
      /* Card body */
      '.fb-card-body { padding: 16px; overflow-y: auto; flex: 1; }',
      '.fb-message { background: #f0f7ff; border-left: 4px solid #2980b9;',
      '  border-radius: 0 8px 8px 0; padding: 12px 14px; margin-bottom: 12px; }',
      '.fb-message .fb-quote { font-size: 14px; line-height: 1.6; color: #333;',
      '  font-style: italic; }',
      '.fb-message .fb-date { font-size: 11px; color: #888; margin-top: 6px; }',
      '',
      /* Action checklist */
      '.fb-checklist { margin-bottom: 12px; }',
      '.fb-checklist h4 { font-size: 13px; color: #1a3d5c; margin-bottom: 8px;',
      '  display: flex; align-items: center; gap: 6px; }',
      '.fb-check-item { display: flex; align-items: flex-start; gap: 8px; padding: 6px 0;',
      '  font-size: 13px; color: #333; cursor: pointer; }',
      '.fb-check-item input[type="checkbox"] { margin-top: 2px; accent-color: #1A8A6E;',
      '  width: 16px; height: 16px; cursor: pointer; }',
      '.fb-check-item.checked label { color: #888; text-decoration: line-through; }',
      '',
      /* Progress bar */
      '.fb-progress { background: #eee; border-radius: 10px; height: 6px;',
      '  margin: 8px 0 12px; overflow: hidden; }',
      '.fb-progress-fill { background: linear-gradient(90deg, #1A8A6E, #27ae60);',
      '  height: 100%; border-radius: 10px; transition: width 0.3s ease; }',
      '',
      /* Action buttons */
      '.fb-actions { display: flex; flex-direction: column; gap: 6px; }',
      '.fb-action-btn { padding: 10px 14px; border-radius: 8px; border: none;',
      '  font-size: 13px; font-weight: 600; cursor: pointer; font-family: inherit;',
      '  transition: all 0.15s; text-align: center; }',
      '.fb-action-primary { background: #1A8A6E; color: #fff; }',
      '.fb-action-primary:hover { background: #16755d; }',
      '.fb-action-secondary { background: #f0f4f8; color: #2c3e50; border: 1px solid #ddd; }',
      '.fb-action-secondary:hover { background: #e4e8ec; }',
      '',
      /* Done state */
      '.fb-done { text-align: center; padding: 8px 0; }',
      '.fb-done .fb-done-icon { font-size: 32px; margin-bottom: 4px; }',
      '.fb-done .fb-done-text { font-size: 14px; color: #1A8A6E; font-weight: 600; }',
      '.fb-done .fb-done-sub { font-size: 12px; color: #888; margin-top: 2px; }',
      '',
      /* Toast */
      '.sub-toast { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%) translateY(100px);',
      '  background: #1A8A6E; color: #fff; padding: 12px 24px; border-radius: 10px;',
      '  font-size: 13px; z-index: 10000; opacity: 0; transition: all 0.3s ease;',
      '  box-shadow: 0 4px 16px rgba(0,0,0,0.15); text-align: center; max-width: 90vw; }',
      '.sub-toast.show { transform: translateX(-50%) translateY(0); opacity: 1; }',
      '.sub-toast.error { background: #c0392b; }',
      '',
      /* Print + mobile */
      '@media print { .sub-btn, .sub-status, .fb-coach, .sub-toast { display: none !important; } }',
      '@media (max-width: 500px) {',
      '  .fb-card { width: calc(100vw - 24px); right: 12px; bottom: 12px; }',
      '  .fb-trigger { bottom: 12px; right: 12px; }',
      '}'
    ].join('\n');
    document.head.appendChild(style);
  }

  // ===== SUBMIT BUTTON =====
  function injectSubmitButton() {
    var container = document.querySelector(config.containerSelector);
    if (!container) return;
    var btn = document.createElement('button');
    btn.className = 'sub-btn';
    btn.id = 'subSubmitBtn';
    btn.title = 'Submit your work to your teacher';
    btn.innerHTML = '<span class="btn-label">\uD83D\uDCE4 Submit to Teacher</span><span class="spinner"></span>';
    btn.onclick = handleSubmit;
    if (config.insertMode === 'append') {
      container.appendChild(btn);
    } else if (config.insertMode === 'prepend') {
      container.insertBefore(btn, container.firstChild);
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

  // ===== FEEDBACK COACH WIDGET =====
  function injectFeedbackCoach() {
    var coach = document.createElement('div');
    coach.className = 'fb-coach';
    coach.id = 'fbCoach';
    coach.innerHTML =
      /* Floating trigger */
      '<button class="fb-trigger" id="fbTrigger" onclick="window._fbToggle()">' +
        '<span id="fbTriggerIcon">\uD83D\uDCAC</span>' +
        '<span id="fbTriggerText">New feedback!</span>' +
      '</button>' +
      /* Expanded card */
      '<div class="fb-card" id="fbCard">' +
        '<div class="fb-card-header">' +
          '<h3>\uD83C\uDFAF Feedback Coach</h3>' +
          '<button class="fb-card-close" onclick="window._fbClose()">&times;</button>' +
        '</div>' +
        '<div class="fb-card-body" id="fbCardBody"></div>' +
      '</div>';
    document.body.appendChild(coach);
  }

  // Toggle card open/closed
  window._fbToggle = function () {
    var card = document.getElementById('fbCard');
    var trigger = document.getElementById('fbTrigger');
    if (card.classList.contains('visible')) {
      card.classList.remove('visible');
      trigger.style.display = 'flex';
    } else {
      renderFeedbackCard();
      card.classList.add('visible');
      trigger.style.display = 'none';
      if (feedbackState === 'new') {
        feedbackState = 'read';
        saveFbState();
      }
    }
  };

  window._fbClose = function () {
    document.getElementById('fbCard').classList.remove('visible');
    document.getElementById('fbTrigger').style.display = 'flex';
    updateTriggerAppearance();
  };

  // Check all items clicked
  window._fbCheck = function (cb, idx) {
    var item = cb.closest('.fb-check-item');
    item.classList.toggle('checked', cb.checked);
    updateCheckProgress();
    // Save check state
    try {
      var checks = JSON.parse(sessionStorage.getItem(config.storageKey + '-fb-checks') || '{}');
      checks[idx] = cb.checked;
      sessionStorage.setItem(config.storageKey + '-fb-checks', JSON.stringify(checks));
    } catch (e) {}
  };

  window._fbStartWorking = function () {
    feedbackState = 'working';
    saveFbState();
    document.getElementById('fbCard').classList.remove('visible');
    updateTriggerAppearance();
    showToast('Great! Take your time to improve your work. \uD83D\uDCAA');
  };

  window._fbResubmit = function () {
    document.getElementById('fbCard').classList.remove('visible');
    document.getElementById('fbTrigger').style.display = 'none';
    // Trigger the submit button
    var btn = document.getElementById('subSubmitBtn');
    if (btn) {
      btn.scrollIntoView({ behavior: 'smooth', block: 'center' });
      btn.style.animation = 'fb-bounce 0.5s ease-in-out 2';
      setTimeout(function () { btn.style.animation = ''; }, 1200);
    }
  };

  function renderFeedbackCard() {
    var body = document.getElementById('fbCardBody');
    if (!feedbackData) { body.innerHTML = '<p style="color:#888;text-align:center;padding:20px;">No feedback yet.</p>'; return; }

    var isRevision = feedbackData.revision;
    var checks = {};
    try { checks = JSON.parse(sessionStorage.getItem(config.storageKey + '-fb-checks') || '{}'); } catch (e) {}

    // Different checklists for regular feedback vs revision
    var checkItems = isRevision ? [
      'Read the feedback carefully',
      'Go to the section that needs rewriting',
      'Delete the old text and write YOUR OWN words',
      'Check: does it sound like something YOU would say?',
      'Resubmit to teacher'
    ] : [
      'Read the feedback carefully',
      'Find the section to improve',
      'Make your changes',
      'Resubmit to teacher'
    ];

    var html = '';

    // Header colour and icon differ for revision
    var headerEl = document.querySelector('.fb-card-header');
    if (headerEl) {
      headerEl.style.background = isRevision
        ? 'linear-gradient(135deg, #922b21, #c0392b)'
        : 'linear-gradient(135deg, #1a3d5c, #2c5f8a)';
    }
    var headerTitle = document.querySelector('.fb-card-header h3');
    if (headerTitle) {
      headerTitle.innerHTML = isRevision
        ? '\u270D\uFE0F Revision Coach'
        : '\uD83C\uDFAF Feedback Coach';
    }

    // Revision banner
    if (isRevision) {
      html += '<div style="background:#fce4e4;border:1px solid #f5c6cb;border-radius:8px;' +
        'padding:10px 12px;margin-bottom:12px;font-size:13px;color:#922b21;">' +
        '<strong>\u270D\uFE0F Revision needed</strong> — Your teacher wants you to rewrite this section ' +
        'in your own words. That is okay! Use the sentence starters below to help you.' +
        '</div>';
    }

    // Teacher's message
    var msgBorder = isRevision ? '#c0392b' : '#2980b9';
    var msgBg = isRevision ? '#fef5f5' : '#f0f7ff';
    html += '<div class="fb-message" style="border-left-color:' + msgBorder + ';background:' + msgBg + ';">' +
      '<div class="fb-quote">\u201C' + esc(feedbackData.text) + '\u201D</div>' +
      '<div class="fb-date">From your teacher \u00b7 ' + formatDate(feedbackData.date) + '</div>' +
      '</div>';

    // Sentence starters (only for revision mode)
    if (isRevision) {
      html += '<div style="background:#f0f7ff;border-radius:8px;padding:10px 14px;margin-bottom:12px;">' +
        '<div style="font-size:12px;font-weight:700;color:#1a3d5c;margin-bottom:6px;">' +
        '\uD83D\uDCA1 Sentence starters to help you</div>' +
        '<div style="font-size:13px;color:#333;line-height:1.7;">' +
        '\u2022 "In class, we learned that..."<br>' +
        '\u2022 "I think this is important because..."<br>' +
        '\u2022 "One example from our activity is..."<br>' +
        '\u2022 "The design process helped me to..."<br>' +
        '\u2022 "I found that [material] works well because..."<br>' +
        '\u2022 "If I could improve my design, I would..."' +
        '</div></div>';
    }

    // Checklist
    html += '<div class="fb-checklist">' +
      '<h4>\u2705 What to do next</h4>';
    for (var i = 0; i < checkItems.length; i++) {
      var checked = checks[i] ? ' checked' : '';
      var cls = checks[i] ? ' checked' : '';
      html += '<div class="fb-check-item' + cls + '">' +
        '<input type="checkbox" id="fbChk' + i + '"' + checked + ' onchange="window._fbCheck(this,' + i + ')">' +
        '<label for="fbChk' + i + '">' + checkItems[i] + '</label></div>';
    }
    html += '</div>';

    // Progress bar
    var done = 0;
    for (var j = 0; j < checkItems.length; j++) { if (checks[j]) done++; }
    var pct = Math.round((done / checkItems.length) * 100);
    var barColour = isRevision ? 'linear-gradient(90deg, #e74c3c, #c0392b)' : 'linear-gradient(90deg, #1A8A6E, #27ae60)';
    html += '<div class="fb-progress"><div class="fb-progress-fill" style="width:' + pct + '%;background:' + barColour + '"></div></div>';

    // Action buttons
    var readyThreshold = checkItems.length - 1; // all but last = ready to resubmit
    if (feedbackState === 'resubmitted') {
      html += '<div class="fb-done">' +
        '<div class="fb-done-icon">\uD83C\uDF1F</div>' +
        '<div class="fb-done-text">' + (isRevision ? 'Revision submitted!' : 'Feedback addressed!') + '</div>' +
        '<div class="fb-done-sub">Your teacher will review your updated work.</div>' +
        '</div>';
    } else if (done >= readyThreshold) {
      html += '<div class="fb-actions">' +
        '<button class="fb-action-btn fb-action-primary" onclick="window._fbResubmit()"' +
          (isRevision ? ' style="background:#c0392b;"' : '') + '>' +
          '\uD83D\uDCE4 I\'m ready \u2014 resubmit my work</button>' +
        '</div>';
    } else if (feedbackState === 'read' || feedbackState === 'working') {
      html += '<div class="fb-actions">' +
        '<button class="fb-action-btn fb-action-primary" onclick="window._fbStartWorking()"' +
          (isRevision ? ' style="background:#c0392b;"' : '') + '>' +
          (isRevision ? '\u270D\uFE0F I\'ll rewrite it now' : '\uD83D\uDCAA Got it \u2014 I\'ll work on it') + '</button>' +
        '</div>';
    } else {
      html += '<div class="fb-actions">' +
        '<button class="fb-action-btn fb-action-primary" onclick="window._fbStartWorking()"' +
          (isRevision ? ' style="background:#c0392b;"' : '') + '>' +
          (isRevision ? '\u270D\uFE0F I\'ll rewrite it now' : '\uD83D\uDCAA Got it \u2014 I\'ll work on it') + '</button>' +
        '</div>';
    }

    body.innerHTML = html;
  }

  function updateCheckProgress() {
    var checks = {};
    try { checks = JSON.parse(sessionStorage.getItem(config.storageKey + '-fb-checks') || '{}'); } catch (e) {}
    var done = 0;
    for (var j = 0; j < 4; j++) { if (checks[j]) done++; }
    var pct = Math.round((done / 4) * 100);
    var fill = document.querySelector('.fb-progress-fill');
    if (fill) fill.style.width = pct + '%';

    // If 3+ checked, show resubmit button
    if (done >= 3) renderFeedbackCard();
  }

  function updateTriggerAppearance() {
    var trigger = document.getElementById('fbTrigger');
    var icon = document.getElementById('fbTriggerIcon');
    var text = document.getElementById('fbTriggerText');
    if (!trigger) return;

    trigger.className = 'fb-trigger';
    var isRevision = feedbackData && feedbackData.revision;

    if (feedbackState === 'new') {
      trigger.classList.add('state-new');
      if (isRevision) {
        trigger.style.background = '#c0392b';
        icon.textContent = '\u270D\uFE0F';
        text.textContent = 'Revision needed!';
      } else {
        icon.textContent = '\uD83D\uDCAC';
        text.textContent = 'New feedback!';
      }
      trigger.style.display = 'flex';
    } else if (feedbackState === 'read' || feedbackState === 'working') {
      trigger.classList.add('state-working');
      if (isRevision) {
        trigger.style.background = '#c0392b';
        icon.textContent = '\u270D\uFE0F';
        text.textContent = 'Rewriting...';
      } else {
        icon.textContent = '\uD83D\uDCDD';
        text.textContent = 'Working on feedback';
      }
      trigger.style.display = 'flex';
    } else if (feedbackState === 'resubmitted') {
      trigger.classList.add('state-resubmitted');
      icon.textContent = '\u2705';
      text.textContent = 'Feedback done!';
      trigger.style.display = 'flex';
      setTimeout(function () { trigger.style.display = 'none'; }, 5000);
    } else {
      trigger.style.display = 'none';
    }
  }

  function saveFbState() {
    try { sessionStorage.setItem(config.storageKey + '-fb-state', feedbackState); } catch (e) {}
  }

  // ===== FEEDBACK CHECK =====
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
          var rawText = result.feedbackText || '';
          var isRevision = rawText.indexOf('[REVISION]') === 0;
          var cleanText = isRevision ? rawText.replace('[REVISION] ', '').replace('[REVISION]', '') : rawText;
          feedbackData = { text: cleanText, date: result.feedbackDate, revision: isRevision };

          // Check if this is new feedback (different from what we last saw)
          var lastSeenFb = null;
          try { lastSeenFb = localStorage.getItem(config.storageKey + '-fb-last-seen'); } catch (e) {}

          if (lastSeenFb !== result.feedbackText) {
            // New or updated feedback
            feedbackState = 'new';
            saveFbState();
            // Clear old checklist
            try { sessionStorage.removeItem(config.storageKey + '-fb-checks'); } catch (e) {}
          }

          // Save what we've seen
          try { localStorage.setItem(config.storageKey + '-fb-last-seen', result.feedbackText); } catch (e) {}

          updateTriggerAppearance();
        }
      })
      .catch(function () { /* silently fail */ });
  }

  // ===== SUBMISSION =====
  function handleSubmit() {
    var name = config.getStudentName();
    var cls = config.getStudentClass();

    if (!name) { showToast('Please enter your name first.', 'error'); return; }
    if (!cls) { showToast('Please select your class first.', 'error'); return; }
    if (!APPS_SCRIPT_URL) { showToast('Submission not configured. Your work is saved locally.', 'error'); return; }

    if (!submissionMeta.privacyAcknowledged) {
      if (!confirm('Privacy Notice: Your name and answers will be sent to your teacher\'s Google account. No other service receives your data.\n\nClick OK to submit.')) return;
      submissionMeta.privacyAcknowledged = true;
      saveSubmissionMeta();
    }

    var btn = document.getElementById('subSubmitBtn');
    btn.classList.add('loading');
    btn.disabled = true;

    var payload = {
      action: 'submit',
      studentName: name,
      studentClass: cls,
      worksheetId: config.storageKey,
      worksheetTitle: config.worksheetTitle,
      data: config.getData()
    };

    localStorage.setItem('y7-student-name', name);
    localStorage.setItem('y7-student-class', cls);

    sendToAppsScript(payload)
      .then(function () {
        btn.classList.remove('loading');
        btn.disabled = false;
        submissionMeta.lastSubmitted = new Date().toISOString();
        submissionMeta.submissionCount = (submissionMeta.submissionCount || 0) + 1;
        saveSubmissionMeta();
        updateStatusBadge();

        // Update feedback state if they had feedback
        if (feedbackData && (feedbackState === 'working' || feedbackState === 'read')) {
          feedbackState = 'resubmitted';
          saveFbState();
          updateTriggerAppearance();
          showToast('\u2705 Submitted! Your teacher will see your improvements.');
        } else {
          showToast('\uD83D\uDCE4 Submitted! Your teacher can now see your work.');
        }

        localStorage.removeItem(config.storageKey + '-pending-submission');
      })
      .catch(function () {
        btn.classList.remove('loading');
        btn.disabled = false;
        localStorage.setItem(config.storageKey + '-pending-submission', JSON.stringify(payload));
        showToast('Could not submit \u2014 saved for retry when you\'re back online.', 'error');
      });
  }

  // ===== NETWORK =====
  function sendToAppsScript(payload) {
    return fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload)
    })
    .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
    .then(function (result) { if (result.status !== 'ok') throw new Error(result.message); return result; });
  }

  function retryPendingSubmission() {
    var pendingKey = config.storageKey + '-pending-submission';
    var pending = localStorage.getItem(pendingKey);
    if (!pending || !APPS_SCRIPT_URL || !navigator.onLine) return;
    try {
      sendToAppsScript(JSON.parse(pending))
        .then(function () {
          localStorage.removeItem(pendingKey);
          submissionMeta.lastSubmitted = new Date().toISOString();
          submissionMeta.submissionCount = (submissionMeta.submissionCount || 0) + 1;
          saveSubmissionMeta();
          updateStatusBadge();
          showToast('Previous submission sent!');
        })
        .catch(function () {});
    } catch (e) { console.warn('submission.js: could not parse pending submission, keeping for next try:', e.message); }
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
      badge.innerHTML = '<span class="sub-offline-dot online"></span>Submitted ' +
        submissionMeta.submissionCount + ' time' + (submissionMeta.submissionCount > 1 ? 's' : '') +
        ' \u2014 last: ' + formatDate(submissionMeta.lastSubmitted);
    } else {
      badge.textContent = 'Not yet submitted';
    }
  }

  function updateOnlineStatus(isOnline) {
    var btn = document.getElementById('subSubmitBtn');
    if (!btn) return;
    btn.disabled = !isOnline;
    btn.title = isOnline ? 'Submit your work to your teacher' : 'You are offline \u2014 work is saved locally';
    var dot = document.querySelector('.sub-offline-dot');
    if (dot) dot.className = 'sub-offline-dot ' + (isOnline ? 'online' : 'offline');
  }

  function restoreCachedName() {
    var cachedName = localStorage.getItem('y7-student-name');
    var cachedClass = localStorage.getItem('y7-student-class');
    if (cachedName) {
      var el = document.getElementById('f-name') || document.getElementById('studentName');
      if (el && !el.value.trim()) { el.value = cachedName; if (typeof window.saveAll === 'function') window.saveAll(); }
    }
    if (cachedClass) {
      var el2 = document.getElementById('f-class') || document.getElementById('studentClass');
      if (el2 && !el2.value) { el2.value = cachedClass; if (typeof window.saveAll === 'function') window.saveAll(); }
    }
  }

  function formatDate(iso) {
    try {
      var d = new Date(iso);
      return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short' }) +
        ', ' + d.toLocaleTimeString('en-AU', { hour: 'numeric', minute: '2-digit' });
    } catch (e) { return iso; }
  }

  function esc(str) {
    if (!str) return '';
    var d = document.createElement('div'); d.textContent = String(str); return d.innerHTML;
  }

  function showToast(msg, type) {
    var old = document.querySelector('.sub-toast'); if (old) old.remove();
    var toast = document.createElement('div');
    toast.className = 'sub-toast' + (type === 'error' ? ' error' : '');
    toast.textContent = msg;
    document.body.appendChild(toast);
    requestAnimationFrame(function () { requestAnimationFrame(function () { toast.classList.add('show'); }); });
    setTimeout(function () { toast.classList.remove('show'); setTimeout(function () { toast.remove(); }, 300); }, 4000);
  }

})();
