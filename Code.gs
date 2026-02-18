/**
 * ============================================================================
 *  Game Lab Autograder v2  —  Google Apps Script
 * ============================================================================
 *
 *  Grades Code.org Game Lab projects against rubric criteria using an LLM.
 *
 *  Setup:
 *    1. Paste this entire file into Extensions → Apps Script → Code.gs
 *    2. Set GEMINI_API_KEY in Project Settings → Script Properties
 *    3. Run Autograder → Initial Setup… from the spreadsheet menu
 *    4. See README.md or Autograder → Help / Setup Guide for full instructions
 *
 *  Default LLM: Gemini 2.0 Flash
 *  Optional:    Set LLM_PROVIDER=openai and OPENAI_API_KEY for OpenAI
 * ============================================================================
 */

// ═══════════════════════════════════════════════════════════════════════════════
//  1. CONFIG & CONSTANTS
// ═══════════════════════════════════════════════════════════════════════════════

var SHEET_SUB    = 'Submissions';
var SHEET_LEVELS = 'Levels';
var SHEET_CRIT   = 'Criteria';
var GRADE_VIEW_PREFIX = 'Grade View P';
var MAX_PERIODS  = 8;

var DEFAULT_PROVIDER = 'gemini';
var DEFAULT_MODEL_BY_PROVIDER = {
  gemini: 'gemini-2.0-flash',
  openai: 'gpt-4o'
};

var SUB_HEADERS = [
  'Timestamp','First','Last','Period','Email','LevelID','ShareURL',
  'ChannelID','Score','MaxScore','Status','Notes','EmailedAt'
];

var GRADE_VIEW_HEADERS = [
  'LevelID','First','Last','Score','MaxScore','Status','Email','ShareURL','Timestamp','Notes'
];

// Column indices in Submissions (0-based) — kept in sync with SUB_HEADERS
var SC = {};
SUB_HEADERS.forEach(function(h, i) { SC[h] = i; });

// ═══════════════════════════════════════════════════════════════════════════════
//  2. MENU
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Autograder')
    .addItem('Initial Setup\u2026',         'showSetupDialog')
    .addSeparator()
    .addItem('Grade New Submissions',        'gradeNewRows')
    .addItem('Re-grade Selected Rows',       'gradeSelectedRows')
    .addItem('Re-grade All Rows\u2026',      'gradeAllRows')
    .addSeparator()
    .addItem('Grade & Email All New',        'gradeAndEmailAllNew')
    .addItem('Email Selected Rows',          'emailSelectedRows')
    .addSeparator()
    .addItem('Test API Connection',          'testAPIConnection')
    .addSeparator()
    .addItem('Help / Setup Guide',           'showHelp')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════════
//  3. SETUP WIZARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Shows the HTML dialog for picking periods and creating sheets.
 */
function showSetupDialog() {
  var ss = SpreadsheetApp.getActive();
  var existing = [];
  for (var p = 1; p <= MAX_PERIODS; p++) {
    if (ss.getSheetByName(GRADE_VIEW_PREFIX + p)) existing.push(p);
  }

  var html = HtmlService.createHtmlOutput(buildSetupHtml_(existing))
    .setWidth(460)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Game Lab Autograder — Initial Setup');
}

function buildSetupHtml_(existingPeriods) {
  var checkboxes = '';
  for (var p = 1; p <= MAX_PERIODS; p++) {
    var exists = existingPeriods.indexOf(p) >= 0;
    var disabled = exists ? ' disabled' : '';
    var checked  = exists ? ' checked' : '';
    var label    = exists ? 'Period ' + p + ' \u2713' : 'Period ' + p;
    var style    = exists ? 'color:#888;' : '';
    checkboxes +=
      '<label style="margin:4px 0;' + style + '">' +
      '<input type="checkbox" value="' + p + '"' + disabled + checked + '> ' + label +
      '</label>';
  }

  return '' +
    '<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.5;">' +
    '<p style="margin:0 0 12px 0;">Select class periods to create Grade View sheets for:</p>' +
    '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:2px 0;margin:0 0 16px 4px;">' + checkboxes + '</div>' +
    '<p style="margin:0 0 8px 0;color:#666;font-size:11px;">' +
    'Periods with \u2713 already exist and will not be modified.</p>' +
    '<div style="display:flex;gap:8px;margin-top:12px;">' +
    '<button id="btnCreate" onclick="doCreate()" style="padding:8px 20px;font-size:13px;cursor:pointer;' +
    'background:#4285f4;color:#fff;border:none;border-radius:4px;">Create Sheets</button>' +
    '<button id="btnReset" onclick="doReset()" style="padding:8px 20px;font-size:13px;cursor:pointer;' +
    'background:#ea4335;color:#fff;border:none;border-radius:4px;">Reset Everything</button>' +
    '<button id="btnCancel" onclick="google.script.host.close()" style="padding:8px 20px;font-size:13px;' +
    'cursor:pointer;border:1px solid #ccc;border-radius:4px;background:#fff;">Cancel</button>' +
    '</div>' +
    '<div id="status" style="margin-top:10px;color:#1a73e8;font-size:12px;min-height:18px;"></div>' +
    '</div>' +
    '<script>' +
    'function setWorking(msg){' +
    '  document.querySelectorAll("button").forEach(function(b){b.disabled=true;b.style.opacity="0.6";b.style.cursor="wait";});' +
    '  document.getElementById("status").textContent=msg||"Working\u2026";' +
    '}' +
    'function setReady(){' +
    '  document.querySelectorAll("button").forEach(function(b){b.disabled=false;b.style.opacity="1";b.style.cursor="pointer";});' +
    '  document.getElementById("status").textContent="";' +
    '}' +
    'function getChecked(){' +
    '  var cbs=document.querySelectorAll("input[type=checkbox]:checked:not(:disabled)");' +
    '  var arr=[];cbs.forEach(function(cb){arr.push(parseInt(cb.value));});return arr;' +
    '}' +
    'function doCreate(){' +
    '  var periods=getChecked();' +
    '  setWorking("\u23F3 Creating sheets\u2026 please wait.");' +
    '  google.script.run.withSuccessHandler(function(msg){' +
    '    alert(msg);google.script.host.close();' +
    '  }).withFailureHandler(function(e){' +
    '    setReady();alert("Error: "+e.message);' +
    '  }).createSheetsFromSetup(periods);' +
    '}' +
    'function doReset(){' +
    '  if(!confirm("This will DELETE all Autograder sheets and data.\\n\\nAre you sure?"))return;' +
    '  setWorking("\u23F3 Resetting\u2026 please wait.");' +
    '  google.script.run.withSuccessHandler(function(){' +
    '    alert("All autograder sheets deleted. Re-opening setup...");' +
    '    google.script.host.close();' +
    '    google.script.run.showSetupDialog();' +
    '  }).withFailureHandler(function(e){' +
    '    setReady();alert("Error: "+e.message);' +
    '  }).resetEverything();' +
    '}' +
    '</script>';
}

/**
 * Called from the setup dialog. Creates Submissions, Levels, Criteria (if missing)
 * and any new Grade View P# sheets for the selected periods.
 */
function createSheetsFromSetup(newPeriods) {
  var ss = SpreadsheetApp.getActive();

  // --- Submissions ---
  var sub = ss.getSheetByName(SHEET_SUB);
  if (!sub) {
    sub = ss.insertSheet(SHEET_SUB);
    sub.clear();
    sub.getRange(1, 1, 1, SUB_HEADERS.length).setValues([SUB_HEADERS]).setFontWeight('bold');
    sub.setFrozenRows(1);

    // LevelID dropdown validation in column F (index 6)
    var lvlRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getSheetByName(SHEET_LEVELS)
        ? ss.getSheetByName(SHEET_LEVELS).getRange('A2:A')
        : sub.getRange('F2:F'), // fallback; will fix after Levels created
        true)
      .setAllowInvalid(false)
      .setHelpText('Pick a LevelID from the Levels sheet')
      .build();
    // We'll set this after Levels sheet exists (see below)
  }

  // --- Levels ---
  var lev = ss.getSheetByName(SHEET_LEVELS);
  if (!lev) {
    lev = ss.insertSheet(SHEET_LEVELS);
    lev.clear();
    var levHeaders = ['LevelID', 'LevelURL', 'Enabled', 'Model'];
    lev.getRange(1, 1, 1, levHeaders.length).setValues([levHeaders]).setFontWeight('bold');

    var rubric = parseCriteriaTableCsv_();
    var levelIds = Object.keys(rubric.byLevel).sort();
    var levelRows = levelIds.map(function(id) {
      return [id, levelIdToUrl_(id), true, ''];
    });
    if (levelRows.length) {
      lev.getRange(2, 1, levelRows.length, 4).setValues(levelRows);
      // Checkbox for Enabled column
      lev.getRange(2, 3, levelRows.length, 1)
        .insertCheckboxes();
    }
    lev.setFrozenRows(1);
  }

  // Now set LevelID validation on Submissions if it was just created
  if (sub && lev) {
    var lvlValRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(lev.getRange('A2:A'), true)
      .setAllowInvalid(false)
      .setHelpText('Pick a LevelID from the Levels sheet')
      .build();
    sub.getRange('F2:F').setDataValidation(lvlValRule);
  }

  // --- Criteria ---
  var crit = ss.getSheetByName(SHEET_CRIT);
  if (!crit) {
    crit = ss.insertSheet(SHEET_CRIT);
    crit.clear();
    var critHeaders = ['LevelID', 'CriterionID', 'Points', 'Type', 'Description', 'Notes', 'Teacher Notes'];
    crit.getRange(1, 1, 1, critHeaders.length).setValues([critHeaders]).setFontWeight('bold');
    crit.setFrozenRows(1);
    var rubric2 = parseCriteriaTableCsv_();
    if (rubric2.rows.length) {
      crit.getRange(2, 1, rubric2.rows.length, critHeaders.length).setValues(rubric2.rows);
    }
  }

  // --- Grade View sheets ---
  var createdCount = 0;
  if (newPeriods && newPeriods.length) {
    newPeriods.forEach(function(p) {
      var name = GRADE_VIEW_PREFIX + p;
      if (ss.getSheetByName(name)) return; // already exists
      var gv = ss.insertSheet(name);
      gv.clear();
      gv.getRange(1, 1, 1, GRADE_VIEW_HEADERS.length)
        .setValues([GRADE_VIEW_HEADERS])
        .setFontWeight('bold');
      gv.setFrozenRows(1);

      // Array formula that filters+sorts from Submissions
      var formula = buildGradeViewFormula_(p);
      gv.getRange(2, 1).setFormula(formula);

      // Protect the sheet
      var protection = gv.protect().setDescription('Auto-generated grade view — do not edit');
      protection.setWarningOnly(true);

      // Grade View column widths
      setColumnWidths_(gv, {
        LevelID: 180, First: 120, Last: 120, Score: 55, MaxScore: 75,
        Status: 80, Email: 180, ShareURL: 160, Timestamp: 140, Notes: 300
      });

      createdCount++;
    });
  }

  // --- Clean up leftover junk sheets ---
  ['_autograder_reset_temp_', 'Sheet1'].forEach(function(name) {
    var junk = ss.getSheetByName(name);
    if (junk && ss.getSheets().length > 1) {
      try { ss.deleteSheet(junk); } catch (_) {}
    }
  });

  // --- Conditional formatting on Submissions Status column ---
  applyStatusFormatting_(sub);

  // --- Set column widths ---
  setColumnWidths_(sub, {
    Timestamp: 140, First: 120, Last: 120, Period: 60, Email: 180,
    LevelID: 180, ShareURL: 160, ChannelID: 100,
    Score: 55, MaxScore: 75, Status: 80, Notes: 300, EmailedAt: 140
  });
  setColumnWidths_(lev, { LevelID: 200, LevelURL: 400, Enabled: 70, Model: 160 });
  // Criteria — auto-resize is fine for the rubric descriptions
  if (crit) {
    var critCols = crit.getLastColumn();
    for (var c = 1; c <= critCols; c++) crit.autoResizeColumn(c);
  }

  var msg = 'Setup complete!\n\n';
  msg += '\u2022 Submissions sheet: ready\n';
  msg += '\u2022 Levels: ' + (lev ? lev.getLastRow() - 1 : 0) + ' levels loaded\n';
  msg += '\u2022 Criteria: ' + (crit ? crit.getLastRow() - 1 : 0) + ' rubric rows loaded\n';
  if (createdCount) msg += '\u2022 Created ' + createdCount + ' new Grade View sheet(s)\n';
  msg += '\nNext steps:\n';
  msg += '1. Set GEMINI_API_KEY in Extensions \u2192 Apps Script \u2192 Project Settings \u2192 Script Properties\n';
  msg += '2. Use "Test API Connection" from the Autograder menu to verify\n';
  msg += '3. See "Help / Setup Guide" for full instructions';
  return msg;
}

/**
 * Builds the SORT(FILTER(...)) formula for a Grade View sheet.
 * Submissions columns (1-indexed): A=Timestamp B=First C=Last D=Period E=Email
 *   F=LevelID G=ShareURL H=ChannelID I=Score J=MaxScore K=Status L=Notes M=EmailedAt
 *
 * Grade View order: LevelID, First, Last, Score, MaxScore, Status, Email, ShareURL, Timestamp, Notes
 */
function buildGradeViewFormula_(periodNum) {
  // We use curly-brace array notation to reorder columns:
  // {F,B,C,I,J,K,E,G,A,L} filtered where D = periodNum, sorted by col1 (LevelID) then col3 (Last)
  // REGEXEXTRACT handles both numeric (7) and text ("Period 7") values in column D
  return '=IFERROR(SORT(FILTER({' +
    'Submissions!F:F,' +   // LevelID  → col 1
    'Submissions!B:B,' +   // First    → col 2
    'Submissions!C:C,' +   // Last     → col 3
    'Submissions!I:I,' +   // Score    → col 4
    'Submissions!J:J,' +   // MaxScore → col 5
    'Submissions!K:K,' +   // Status   → col 6
    'Submissions!E:E,' +   // Email    → col 7
    'Submissions!G:G,' +   // ShareURL → col 8
    'Submissions!A:A,' +   // Timestamp→ col 9
    'Submissions!L:L' +    // Notes    → col 10
    '},IFERROR(VALUE(REGEXEXTRACT(TO_TEXT(Submissions!D:D),"\\d+")),0)=' + periodNum +
    '),1,TRUE,3,TRUE),"")';
}

/**
 * Deletes all autograder sheets so the teacher can start fresh.
 */
function resetEverything() {
  var ss = SpreadsheetApp.getActive();
  var toDelete = [SHEET_SUB, SHEET_LEVELS, SHEET_CRIT];
  for (var p = 1; p <= MAX_PERIODS; p++) toDelete.push(GRADE_VIEW_PREFIX + p);

  // Make sure there's always at least one sheet (Sheets requires it)
  var tempName = '_autograder_reset_temp_';
  var temp = ss.getSheetByName(tempName) || ss.insertSheet(tempName);

  toDelete.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  // Clear grade cache
  try { CacheService.getDocumentCache().removeAll([]); } catch(_) {}

  // The temp sheet is auto-deleted when Initial Setup runs next.
}

/**
 * Applies conditional formatting to the Status column (K) of Submissions.
 */
function applyStatusFormatting_(sh) {
  if (!sh) return;
  var statusCol = SC.Status + 1; // 1-indexed
  var range = sh.getRange(2, statusCol, sh.getMaxRows() - 1, 1);

  var rules = sh.getConditionalFormatRules();

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('OK')
    .setBackground('#d9ead3')
    .setRanges([range])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Error')
    .setBackground('#f4cccc')
    .setRanges([range])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Invalid')
    .setBackground('#fce5cd')
    .setRanges([range])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('disabled')
    .setBackground('#fce5cd')
    .setRanges([range])
    .build());

  sh.setConditionalFormatRules(rules);
}


// ═══════════════════════════════════════════════════════════════════════════════
//  4. GRADING ENGINE
// ═══════════════════════════════════════════════════════════════════════════════

function gradeNewRows() {
  var ss = SpreadsheetApp.getActive();
  ss.toast('Checking for new submissions\u2026', 'Autograder', -1);

  // Step 1: Import any new form responses (if a form is linked)
  var importCount = importFormResponses_();

  // Step 2: Find all ungraded rows
  var sh = getSheet_(SHEET_SUB);
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);
  var targets = [];
  for (var r = 1; r < data.length; r++) {
    var score = data[r][head.Score];
    var url   = data[r][head.ShareURL];
    var lvl   = data[r][head.LevelID];
    if (url && lvl && (score === '' || score === null || score === undefined)) {
      targets.push(r + 1);
    }
  }

  if (!targets.length && !importCount) {
    ss.toast('', 'Autograder', 1);
    SpreadsheetApp.getUi().alert('No new submissions found.\n\nAll rows with a LevelID and ShareURL already have a Score.');
    return;
  }

  // Step 3: Grade them
  if (targets.length) gradeRows_(targets);

  ss.toast('', 'Autograder', 1);
  var msg = 'Done!\n\n';
  if (importCount) msg += '\u2022 Imported ' + importCount + ' new form response(s)\n';
  msg += '\u2022 Graded ' + targets.length + ' submission(s)';
  SpreadsheetApp.getUi().alert(msg);
}

function gradeSelectedRows() {
  var ss = SpreadsheetApp.getActive();
  var sh = getSheet_(SHEET_SUB);
  if (ss.getActiveSheet().getName() !== SHEET_SUB) {
    SpreadsheetApp.getUi().alert('Please switch to the Submissions sheet and select the rows you want to re-grade.');
    return;
  }
  var sel = sh.getActiveRange();
  if (!sel) {
    SpreadsheetApp.getUi().alert('Please select one or more rows in the Submissions sheet first.');
    return;
  }
  var rows = [];
  for (var r = sel.getRow(); r < sel.getRow() + sel.getNumRows(); r++) {
    if (r >= 2) rows.push(r); // skip header
  }
  if (!rows.length) {
    SpreadsheetApp.getUi().alert('No data rows selected (row 1 is the header).');
    return;
  }
  ss.toast('Re-grading ' + rows.length + ' row(s)\u2026', 'Autograder', -1);
  gradeRows_(rows);
  ss.toast('', 'Autograder', 1);
  SpreadsheetApp.getUi().alert('Graded ' + rows.length + ' row(s).');
}

function gradeAllRows() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var res = ui.alert(
    'Re-grade ALL rows?',
    'This will re-grade every submission. It can be slow and may use significant API credits.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  var sh = getSheet_(SHEET_SUB);
  var last = sh.getLastRow();
  var rows = [];
  for (var r = 2; r <= last; r++) rows.push(r);
  if (!rows.length) {
    ui.alert('No submissions to grade.');
    return;
  }
  ss.toast('Re-grading ' + rows.length + ' row(s)\u2026', 'Autograder', -1);
  gradeRows_(rows);
  ss.toast('', 'Autograder', 1);
  ui.alert('Re-graded ' + rows.length + ' submission(s).');
}

/**
 * Core grading loop. Grades the specified row numbers (1-indexed) in Submissions.
 */
function gradeRows_(rowNums) {
  if (!rowNums || !rowNums.length) return;

  var sh = getSheet_(SHEET_SUB);
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);
  var critByLevel = loadCriteriaByLevel_();
  var enabledLevels = loadEnabledLevels_();
  var total = rowNums.length;

  rowNums.forEach(function(rowNum, idx) {
    // Progress toast
    SpreadsheetApp.getActive().toast(
      'Grading row ' + (idx + 1) + ' of ' + total + '\u2026',
      'Autograder', 3
    );

    try {
      var r = rowNum - 1;
      if (r < 0 || r >= data.length) return;

      var url     = String(data[r][head.ShareURL] || '').trim();
      var levelId = String(data[r][head.LevelID]  || '').trim();

      if (!url || !levelId) {
        writeRow_(sh, rowNum, head, { Status: 'No URL/LevelID', Score: 0, Notes: '' });
        return;
      }

      if (!enabledLevels[levelId]) {
        writeRow_(sh, rowNum, head, { Status: 'Level disabled/unknown', Score: 0, Notes: '' });
        return;
      }

      var crits = critByLevel[levelId] || [];
      var maxPts = crits.reduce(function(s, c) { return s + (Number(c.Points) || 0); }, 0);
      if (!crits.length) {
        writeRow_(sh, rowNum, head, { Status: 'No criteria found', Score: 0, MaxScore: 0, Notes: '' });
        return;
      }

      var channelId = extractChannelId_(url);
      if (!channelId) {
        writeRow_(sh, rowNum, head, {
          ChannelID: '', Score: 0, MaxScore: maxPts,
          Status: 'Invalid share link (no ChannelID)',
          Notes: 'Expected a studio.code.org/projects/gamelab/<id> share URL'
        });
        return;
      }

      var fetched = fetchGameLabSource_(channelId);
      if (!fetched || !fetched.ok) {
        writeRow_(sh, rowNum, head, {
          ChannelID: channelId, Score: 0, MaxScore: maxPts,
          Status: 'Invalid share link or unreadable project',
          Notes: (fetched && fetched.msg) ? fetched.msg : 'Fetch failed'
        });
        return;
      }

      // Check cache
      var cacheKey = 'grade:' + sha256_(levelId + '|' + fetched.src);
      var cached = getGradeCache_(cacheKey);
      if (cached) {
        try {
          var cachedResult = JSON.parse(cached);
          writeRow_(sh, rowNum, head, {
            ChannelID: channelId,
            Score: cachedResult.score,
            MaxScore: cachedResult.max,
            Status: 'OK',
            Notes: cachedResult.notes.join(' | ')
          });
          return;
        } catch (_) { /* cache corrupt; re-grade */ }
      }

      // Grade via LLM
      var res = runCriteria_(fetched.src, crits, levelId);
      var patch = {
        ChannelID: channelId,
        Score: res.score,
        MaxScore: res.max,
        Status: 'OK',
        Notes: res.notes.join(' | ')
      };

      writeRow_(sh, rowNum, head, patch);

      // Cache the result (6-hour TTL = 21600 seconds)
      setGradeCache_(cacheKey, JSON.stringify({ score: res.score, max: res.max, notes: res.notes }));

    } catch (e) {
      writeRow_(sh, rowNum, head, { Status: 'Error', Notes: String(e) });
    }
  });
}


// ═══════════════════════════════════════════════════════════════════════════════
//  5. LLM ENGINE
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Runs all criteria (local + LLM) against the given source code.
 * Returns { score, max, notes[] }.
 */
function runCriteria_(src, crits, levelIdOpt) {
  var localCrits = [], llmCrits = [];
  crits.forEach(function(c) {
    var t = String(c.Type || '').trim().toLowerCase();
    if (t === 'code_nonempty' || t === 'contains' || t === 'regex_present' || t === 'regex_absent') {
      localCrits.push(c);
    } else if (t === 'llm_check') {
      llmCrits.push(c);
    }
  });

  var total = crits.reduce(function(s, c) { return s + (Number(c.Points) || 0); }, 0);
  var got = 0, notes = [];

  // --- Local checks ---
  localCrits.forEach(function(c, i) {
    var id   = String(c.CriterionID || ('L' + i));
    var pts  = Number(c.Points) || 0;
    var t    = String(c.Type || '').trim().toLowerCase();
    var desc = String(c.Description || '');
    var pass = false, reason = '';

    if (t === 'code_nonempty') {
      pass = !!(src && String(src).trim().length >= 10);
      reason = pass ? '' : 'Code appears empty/too short.';
    } else if (t === 'contains') {
      var needle = desc.trim();
      pass = !!(needle && String(src || '').toLowerCase().indexOf(needle.toLowerCase()) >= 0);
      reason = pass ? '' : ('Missing text: ' + needle);
    } else if (t === 'regex_present') {
      try { pass = new RegExp(desc, 'i').test(String(src || '')); reason = pass ? '' : 'Pattern not found.'; }
      catch (e) { pass = false; reason = 'Bad regex.'; }
    } else if (t === 'regex_absent') {
      try { pass = !new RegExp(desc, 'i').test(String(src || '')); reason = pass ? '' : 'Forbidden pattern present.'; }
      catch (e) { pass = false; reason = 'Bad regex.'; }
    }

    if (pass) got += pts;
    notes.push((pass ? '\u2705 ' : '\u274C ') + (c.Description || id) + (pass ? '' : (reason ? ' \u2014 ' + reason : '')));
  });

  // --- LLM checks ---
  if (llmCrits.length) {
    var levelId = levelIdOpt || (llmCrits[0] && llmCrits[0].LevelID) || '';
    var res = llmGrade_(levelId, src, llmCrits);
    var byId = res.byId || {};
    llmCrits.forEach(function(c, i) {
      var id  = String(c.CriterionID || ('C' + i));
      var pts = Number(c.Points) || 0;
      var r   = byId[id] || { pass: false, reason: '' };
      if (r.pass) got += pts;
      notes.push((r.pass ? '\u2705 ' : '\u274C ') + (c.Description || id) + (r.pass ? '' : (r.reason ? ' \u2014 ' + r.reason : '')));
    });
  }

  return { score: got, max: total, notes: notes };
}

function buildRubricPrompt_(levelId, src, llmCrits) {
  var checks = llmCrits.map(function(c, i) {
    return {
      id: String(c.CriterionID || ('C' + i)),
      description: String(c.Description || '').trim(),
      points: Number(c.Points) || 0
    };
  });

  var system =
    'You are a strict, consistent autograder for Code.org Game Lab (p5.js-style JavaScript). ' +
    'Given student code and rubric checks, decide PASS/FAIL for each check. ' +
    'If the code is empty/unreadable, mark all FAIL and set unreadable=true. ' +
    'Output JSON only.';

  var scoringRules = [
    'Treat code order as draw order: later shapes appear on top.',
    'rect(x,y) or rect(x,y,w,h) are both valid.',
    'Color can be literal ("purple") or a variable assigned that literal.',
    'Whitespace, comments, and semicolons are irrelevant.',
    'If unsure, mark FAIL (false).'
  ].join('\n- ');

  var user =
    'LEVEL: ' + levelId + '\n' +
    'SCORING RULES:\n- ' + scoringRules + '\n\n' +
    'Return ONLY JSON with this shape: {"unreadable":boolean,"checks":[{"id":string,"pass":boolean,"reason":string}]}.\n' +
    'CHECKS (IDs and descriptions):\n' +
    checks.map(function(x) { return '- ' + x.id + ': ' + x.description + ' (points ' + x.points + ')'; }).join('\n') +
    '\n\nCODE (fenced):\n```javascript\n' + (src || '') + '\n```';

  return { system: system, user: user, expectedIds: checks.map(function(x) { return x.id; }) };
}

function llmGrade_(levelId, src, llmCrits) {
  var provider = getLLMProvider_();
  if (provider === 'openai') return openaiGrade_(levelId, src, llmCrits);
  return geminiGrade_(levelId, src, llmCrits);
}

// ── Gemini ──

function geminiGrade_(levelId, src, llmCrits) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Missing GEMINI_API_KEY in Script properties');

  var built = buildRubricPrompt_(levelId, src, llmCrits);
  var expectedIds = built.expectedIds;
  var model = getModelForLevel_(levelId) || getDefaultModel_();

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(key);

  var body = {
    contents: [{ role: 'user', parts: [{ text: built.system + '\n\n' + built.user }] }],
    generationConfig: { temperature: 0, topP: 1 }
  };

  var resp = fetchWithRetry_(url, {
    method: 'post', contentType: 'application/json',
    muteHttpExceptions: true, payload: JSON.stringify(body)
  });

  var code = resp.getResponseCode();
  var txt  = resp.getContentText();
  if (code >= 400) throw new Error('Gemini HTTP ' + code + ': ' + txt.substring(0, 300));

  var outText = extractGeminiText_(txt);
  var parsed  = normalizeAutogradeJson_(outText, expectedIds);

  var byId = {};
  (parsed.checks || []).forEach(function(ch) {
    byId[String(ch.id)] = { pass: !!ch.pass, reason: ch.reason || '' };
  });
  return { byId: byId, raw: parsed, provider: 'gemini', model: model };
}

function extractGeminiText_(txt) {
  var obj; try { obj = JSON.parse(txt); } catch (e) { return ''; }
  var c = obj && obj.candidates && obj.candidates[0];
  var parts = c && c.content && c.content.parts;
  if (parts && parts.length) {
    return parts.map(function(p) { return p.text || ''; }).join('');
  }
  return (obj && obj.text) ? String(obj.text) : '';
}

// ── OpenAI ──

function openaiGrade_(levelId, src, llmCrits) {
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('Missing OPENAI_API_KEY in Script properties');

  var built = buildRubricPrompt_(levelId, src, llmCrits);
  var expectedIds = built.expectedIds;

  var schema = {
    name: 'autograde_result',
    schema: {
      type: 'object', additionalProperties: false,
      properties: {
        unreadable: { type: 'boolean' },
        checks: {
          type: 'array', items: {
            type: 'object', additionalProperties: false,
            required: ['id', 'pass'],
            properties: { id: { type: 'string' }, pass: { type: 'boolean' }, reason: { type: 'string' } }
          }
        }
      },
      required: ['checks']
    },
    strict: true
  };

  var model  = getModelForLevel_(levelId) || getDefaultModel_();
  var result = callResponsesStructured_(model, key, built.system, built.user, schema);
  var parsed = normalizeAutogradeJson_(result.text, expectedIds);

  var byId = {};
  (parsed.checks || []).forEach(function(ch) {
    byId[String(ch.id)] = { pass: !!ch.pass, reason: ch.reason || '' };
  });
  return { byId: byId, raw: parsed, provider: 'openai', model: model };
}

function extractResponsesText_(txt) {
  var obj; try { obj = JSON.parse(txt); } catch (e) { return ''; }
  return (
    (obj && obj.output_text) ||
    (obj && obj.output && obj.output[0] && obj.output[0].content &&
     obj.output[0].content[0] && obj.output[0].content[0].text) ||
    (obj && obj.choices && obj.choices[0] && obj.choices[0].message &&
     obj.choices[0].message.content) ||
    ''
  );
}

/**
 * Robust OpenAI Responses API call with 3-tier fallback:
 *   json_schema → json_object → plain "ONLY JSON"
 */
function callResponsesStructured_(model, key, system, user, schema) {
  function fetchBody(body) {
    return fetchWithRetry_('https://api.openai.com/v1/responses', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify(body)
    });
  }

  var base = {
    model: model,
    input: [{ role: 'system', content: system }, { role: 'user', content: user }],
    temperature: 0, top_p: 1
  };

  // Attempt 1: json_schema
  var b1 = JSON.parse(JSON.stringify(base));
  b1.response_format = { type: 'json_schema', json_schema: schema };
  var resp1 = fetchBody(b1);
  if (resp1.getResponseCode() < 400) {
    var t1 = extractResponsesText_(resp1.getContentText());
    try { JSON.parse(t1); return { code: resp1.getResponseCode(), text: t1, usedModel: model }; } catch (_) {}
  }

  // Attempt 2: json_object
  var b2 = JSON.parse(JSON.stringify(base));
  b2.response_format = { type: 'json_object' };
  b2.input[1].content = user + '\n\nReturn a JSON object with this exact shape: ' +
    '{"unreadable":boolean,"checks":[{"id":string,"pass":boolean,"reason":string}]}';
  var resp2 = fetchBody(b2);
  if (resp2.getResponseCode() < 400) {
    var t2 = extractResponsesText_(resp2.getContentText());
    try { JSON.parse(t2); return { code: resp2.getResponseCode(), text: t2, usedModel: model }; } catch (_) {}
  }

  // Attempt 3: plain
  var b3 = JSON.parse(JSON.stringify(base));
  b3.input[1].content = user + '\n\nReturn ONLY JSON, no prose.';
  var resp3 = fetchBody(b3);
  var t3 = extractResponsesText_(resp3.getContentText());
  return { code: resp3.getResponseCode(), text: t3, usedModel: model };
}


// ═══════════════════════════════════════════════════════════════════════════════
//  6. CODE.ORG FETCH
// ═══════════════════════════════════════════════════════════════════════════════

function extractChannelId_(url) {
  var m = String(url).match(/https?:\/\/studio\.code\.org\/projects\/gamelab\/([A-Za-z0-9\-_]+)/i);
  return m ? m[1] : '';
}

function fetchGameLabSource_(channelId) {
  var u = 'https://studio.code.org/v3/sources/' + encodeURIComponent(channelId) + '/main.json';
  var res = UrlFetchApp.fetch(u, { muteHttpExceptions: true });
  var code = res.getResponseCode(), body = res.getContentText();
  if (code >= 400) return { ok: false, src: '', msg: 'HTTP ' + code + ' from Code.org' };
  try {
    var parsed = JSON.parse(body);
    var src = (typeof parsed === 'string') ? parsed :
              (parsed && (parsed.source || parsed.code)) ? (parsed.source || parsed.code) : '';
    if (!src || src.trim().length < 10) return { ok: false, src: '', msg: 'Empty or too-short source' };
    return { ok: true, src: src, msg: 'OK' };
  } catch (e) {
    return { ok: false, src: '', msg: 'Non-JSON response (likely invalid share link)' };
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
//  7. EMAIL
// ═══════════════════════════════════════════════════════════════════════════════

function emailSelectedRows() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_SUB);
  if (!sh) { SpreadsheetApp.getUi().alert('Submissions sheet not found. Run Initial Setup first.'); return; }
  if (ss.getActiveSheet().getName() !== SHEET_SUB) {
    SpreadsheetApp.getUi().alert('Please switch to the Submissions sheet and select the rows you want to email.');
    return;
  }
  var sel = sh.getActiveRange();
  if (!sel) { SpreadsheetApp.getUi().alert('Please select rows in the Submissions sheet.'); return; }
  ss.toast('Sending emails\u2026', 'Autograder', -1);
  var count = 0;
  for (var r = sel.getRow(); r < sel.getRow() + sel.getNumRows(); r++) {
    if (r >= 2 && sendEmailForRow_(r)) count++;
  }
  ss.toast('', 'Autograder', 1);
  SpreadsheetApp.getUi().alert('Sent ' + count + ' email(s).');
}

/**
 * Sends a results email for a single row. Returns true if sent, false if skipped.
 * Skips rows with Status "Error" (e.g., API failures) unless force=true.
 */
function sendEmailForRow_(rowNum, force) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUB);
  if (!sh) return false;
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);

  if (head.Email === undefined) return false;
  if (head.EmailedAt === undefined) return false;

  var row   = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  var email = String(row[head.Email] || '').trim();
  if (!email) return false;
  if (row[head.EmailedAt]) return false; // already emailed

  // Don't email students about internal errors (429, timeouts, etc.)
  var status = String(row[head.Status] || '').trim();
  if (!force && status === 'Error') return false;

  var first  = row[head.First] || '';
  var last   = row[head.Last]  || '';
  var level  = row[head.LevelID] || '';
  var url    = row[head.ShareURL] || '';
  var score  = row[head.Score]    || 0;
  var max    = row[head.MaxScore] || 0;
  var notes  = String(row[head.Notes] || '');
  var who    = [first, last].filter(Boolean).join(' ').trim() || 'Student';

  var subject = '[Autograder] ' + level + ' \u2014 ' + score + '/' + max + (who ? (' \u2014 ' + who) : '');
  var items   = notes ? notes.split(' | ') : [];
  var htmlNotes = items.length
    ? '<ul>' + items.map(function(x) { return '<li>' + esc_(x) + '</li>'; }).join('') + '</ul>'
    : '<em>No detailed notes.</em>';
  var statusMsg = (status === 'OK')
    ? 'Your submission was graded automatically.'
    : 'Your submission could not be fully graded: <strong>' + esc_(status) + '</strong>.';

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;">' +
    '<p>Hi ' + esc_(who) + ',</p>' +
    '<p>' + statusMsg + '</p>' +
    '<p><strong>Level:</strong> ' + esc_(level) + '<br>' +
    '<strong>Score:</strong> ' + esc_(score + '/' + max) + '<br>' +
    (url ? '<strong>Link:</strong> <a href="' + esc_(url) + '">your project</a>' : '') +
    '</p>' +
    '<p><strong>Checks:</strong></p>' + htmlNotes +
    '<p style="color:#666;">This email was generated by the class autograder.</p>' +
    '</div>';

  var text =
    'Hi ' + who + ',\n\n' +
    ((status === 'OK') ? 'Your submission was graded automatically.\n' : 'Your submission could not be fully graded: ' + status + '\n') +
    '\nLevel: ' + level +
    '\nScore: ' + score + '/' + max +
    (url ? '\nLink: ' + url : '') +
    '\n\nChecks:\n- ' + (items.length ? items.join('\n- ') : 'No detailed notes.') +
    '\n\n(This email was generated by the class autograder.)';

  GmailApp.sendEmail(email, subject, text, { htmlBody: html });
  sh.getRange(rowNum, head.EmailedAt + 1).setValue(new Date());
  return true;
}


// ═══════════════════════════════════════════════════════════════════════════════
//  8. FORM INTEGRATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Trigger function — set up as: From spreadsheet → On form submit.
 * Appends the response to Submissions, grades it, and emails.
 */
function onFormSubmit(e) {
  try {
    var subSh = getSheet_(SHEET_SUB);
    var subHeaders = subSh.getRange(1, 1, 1, subSh.getLastColumn()).getValues()[0];
    var subMap = headers_(subHeaders);

    // Source = the tab the Form writes to (e.g., "Form Responses 1")
    var srcSh = e.range.getSheet();
    var srcHeaders = srcSh.getRange(1, 1, 1, srcSh.getLastColumn()).getValues()[0];
    var srcMap = headersSmart_(srcHeaders);

    var values = (e.values && e.values.length === srcHeaders.length)
      ? e.values
      : srcSh.getRange(e.range.getRow(), 1, 1, srcSh.getLastColumn()).getValues()[0];

    function getField(name, fallback) {
      var idx = srcMap[name];
      return (idx !== undefined) ? values[idx] : (fallback || '');
    }

    var out = new Array(subHeaders.length).fill('');
    if (subMap.Timestamp !== undefined) out[subMap.Timestamp] = getField('Timestamp', new Date());
    if (subMap.Email     !== undefined) out[subMap.Email]     = getField('Email', '');
    if (subMap.First     !== undefined) out[subMap.First]     = getField('First', '');
    if (subMap.Last      !== undefined) out[subMap.Last]      = getField('Last', '');
    if (subMap.Period    !== undefined) out[subMap.Period]     = toNumber_(getField('Period', ''));
    if (subMap.LevelID   !== undefined) out[subMap.LevelID]   = getField('LevelID', '');
    if (subMap.ShareURL  !== undefined) out[subMap.ShareURL]  = getField('ShareURL', '');

    subSh.appendRow(out);
    var newRow = subSh.getLastRow();

    gradeRows_([newRow]);

    try { sendEmailForRow_(newRow); } catch (_) {}

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err);
  }
}

// ── Sync (Backfill) ──

/**
 * Imports new form responses into Submissions (de-duplicated). No grading or emailing.
 * Returns the number of new rows imported. Returns 0 if no form responses sheet found.
 */
function importFormResponses_() {
  var ss = SpreadsheetApp.getActive();
  var subSh = ss.getSheetByName(SHEET_SUB);
  if (!subSh) return 0;
  var subHeaders = subSh.getRange(1, 1, 1, subSh.getLastColumn()).getValues()[0];
  var subMap = headers_(subHeaders);

  var srcSh = findFormResponsesSheet_();
  if (!srcSh) return 0;

  var srcValues = srcSh.getDataRange().getValues();
  if (srcValues.length <= 1) return 0;
  var srcHead = srcValues[0];
  var srcMap = headersSmart_(srcHead);

  // Build dedup key set from Submissions
  var existing = {};
  var subValues = subSh.getDataRange().getValues();
  for (var i = 1; i < subValues.length; i++) {
    var ts  = normalizeTimestamp_(subValues[i][subMap.Timestamp]);
    var em  = (subMap.Email !== undefined) ? String(subValues[i][subMap.Email] || '').trim().toLowerCase() : '';
    var lvl = (subMap.LevelID !== undefined) ? String(subValues[i][subMap.LevelID] || '').trim() : '';
    existing[[ts, em, lvl].join('|')] = true;
  }

  var count = 0;
  for (var r = 1; r < srcValues.length; r++) {
    var row = srcValues[r];
    var tsVal  = (srcMap.Timestamp !== undefined) ? row[srcMap.Timestamp] : new Date();
    var emVal  = (srcMap.Email     !== undefined) ? row[srcMap.Email]     : '';
    var first  = (srcMap.First     !== undefined) ? row[srcMap.First]     : '';
    var last   = (srcMap.Last      !== undefined) ? row[srcMap.Last]      : '';
    var period = (srcMap.Period    !== undefined) ? row[srcMap.Period]    : '';
    var level  = (srcMap.LevelID   !== undefined) ? row[srcMap.LevelID]  : '';
    var share  = (srcMap.ShareURL  !== undefined) ? row[srcMap.ShareURL] : '';

    var key = [
      normalizeTimestamp_(tsVal),
      String(emVal || '').trim().toLowerCase(),
      String(level || '').trim()
    ].join('|');
    if (existing[key]) continue;

    var out = new Array(subHeaders.length).fill('');
    if (subMap.Timestamp !== undefined) out[subMap.Timestamp] = tsVal || new Date();
    if (subMap.Email     !== undefined) out[subMap.Email]     = emVal;
    if (subMap.First     !== undefined) out[subMap.First]     = first;
    if (subMap.Last      !== undefined) out[subMap.Last]      = last;
    if (subMap.Period    !== undefined) out[subMap.Period]     = toNumber_(period);
    if (subMap.LevelID   !== undefined) out[subMap.LevelID]   = level;
    if (subMap.ShareURL  !== undefined) out[subMap.ShareURL]  = share;

    subSh.appendRow(out);
    count++;
    existing[key] = true;
  }

  return count;
}

/**
 * One-click workflow: import form responses, grade all ungraded, email all un-emailed.
 */
function gradeAndEmailAllNew() {
  var ss = SpreadsheetApp.getActive();
  ss.toast('Syncing, grading & emailing\u2026', 'Autograder', -1);

  // Step 1: Import any new form responses
  var importCount = importFormResponses_();

  // Step 2: Grade all ungraded rows
  var sh = getSheet_(SHEET_SUB);
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);
  var targets = [];
  for (var r = 1; r < data.length; r++) {
    var score = data[r][head.Score];
    var url   = data[r][head.ShareURL];
    var lvl   = data[r][head.LevelID];
    if (url && lvl && (score === '' || score === null || score === undefined)) {
      targets.push(r + 1);
    }
  }
  if (targets.length) gradeRows_(targets);

  // Step 3: Email all un-emailed OK rows
  data = sh.getDataRange().getValues();
  var emailCount = 0;
  for (var r = 1; r < data.length; r++) {
    var status  = String(data[r][head.Status] || '');
    var emailed = data[r][head.EmailedAt];
    var email   = String(data[r][head.Email] || '').trim();
    if (status === 'OK' && !emailed && email) {
      try { if (sendEmailForRow_(r + 1)) emailCount++; } catch (_) {}
    }
  }

  ss.toast('', 'Autograder', 1);
  var msg = 'Done!\n\n';
  if (importCount) msg += '\u2022 Imported ' + importCount + ' new form response(s)\n';
  msg += '\u2022 Graded ' + targets.length + ' submission(s)\n';
  msg += '\u2022 Emailed ' + emailCount + ' student(s)';
  SpreadsheetApp.getUi().alert(msg);
}


// ═══════════════════════════════════════════════════════════════════════════════
//  9. DIAGNOSTICS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Combined API test: checks basic connectivity then structured JSON grading.
 */
function testAPIConnection() {
  var p = getLLMProvider_();
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Testing API connection\u2026', 'Autograder', -1);

  var model = getDefaultModel_();
  var lines = [];
  var allOk = true;

  // --- Test 1: Basic connectivity ---
  try {
    var basicOk = false, basicText = '';
    if (p === 'openai') {
      var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
      if (!key) throw new Error('OPENAI_API_KEY is not set.\n\nGo to Extensions \u2192 Apps Script \u2192 Project Settings \u2192 Script Properties and add it.');
      var resp = fetchWithRetry_('https://api.openai.com/v1/responses', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify({ model: model, input: [{ role: 'user', content: 'Reply with the single word: pong.' }] })
      });
      basicOk = resp.getResponseCode() < 400;
      basicText = basicOk ? extractResponsesText_(resp.getContentText()).substring(0, 80).trim() : ('HTTP ' + resp.getResponseCode());
    } else {
      var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!key) throw new Error('GEMINI_API_KEY is not set.\n\nGo to Extensions \u2192 Apps Script \u2192 Project Settings \u2192 Script Properties and add it.');
      var gUrl = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(key);
      var resp = fetchWithRetry_(gUrl, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: 'Reply with the single word: pong.' }] }], generationConfig: { temperature: 0, topP: 1 } })
      });
      basicOk = resp.getResponseCode() < 400;
      basicText = basicOk ? extractGeminiText_(resp.getContentText()).substring(0, 80).trim() : ('HTTP ' + resp.getResponseCode());
    }
    if (basicOk) lines.push('\u2705 Connection OK (' + basicText + ')');
    else { allOk = false; lines.push('\u274C Connection failed: ' + basicText); }
  } catch (e) {
    allOk = false; lines.push('\u274C ' + String(e));
  }

  // --- Test 2: Structured JSON grading (only if basic passed) ---
  if (allOk) {
    try {
      var checks = [
        { id: 'has_purple', description: 'Code sets fill("purple") before drawing a rectangle.' },
        { id: 'has_draw',   description: 'Code defines a draw() function.' }
      ];
      var checkIds = checks.map(function(c) { return c.id; });
      var system = 'You are a strict autograder. Decide PASS/FAIL per check. Output JSON only.';
      var user = 'Return ONLY JSON: {"unreadable":boolean,"checks":[{"id":string,"pass":boolean,"reason":string}]}\n\n' +
        'CHECKS:\n' + checks.map(function(c) { return '- ' + c.id + ': ' + c.description; }).join('\n') +
        '\n\nCODE:\n```javascript\nfill("purple"); rect(10,10,20,20);\n```';

      var structOk = false, numChecks = 0;
      if (p === 'openai') {
        var oaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
        var schema = {
          name: 'autograde_result',
          schema: {
            type: 'object', additionalProperties: false,
            properties: {
              unreadable: { type: 'boolean' },
              checks: { type: 'array', items: { type: 'object', additionalProperties: false,
                required: ['id', 'pass'],
                properties: { id: { type: 'string' }, pass: { type: 'boolean' }, reason: { type: 'string' } }
              }}
            }, required: ['checks']
          }, strict: true
        };
        var result = callResponsesStructured_(model, oaiKey, system, user, schema);
        var parsed = normalizeAutogradeJson_(result.text, checkIds);
        structOk = parsed && Array.isArray(parsed.checks) && parsed.checks.length > 0;
        numChecks = structOk ? parsed.checks.length : 0;
      } else {
        var gemKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
        var gemUrl = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(gemKey);
        var gResp = fetchWithRetry_(gemUrl, {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: system + '\n\n' + user }] }], generationConfig: { temperature: 0, topP: 1 } })
        });
        var gText = extractGeminiText_(gResp.getContentText());
        var parsed = normalizeAutogradeJson_(gText, checkIds);
        structOk = parsed && Array.isArray(parsed.checks) && parsed.checks.length > 0;
        numChecks = structOk ? parsed.checks.length : 0;
        Logger.log('Structured test raw:\n%s', gText);
      }
      if (structOk) lines.push('\u2705 Structured grading OK (' + numChecks + ' checks parsed)');
      else { allOk = false; lines.push('\u274C Structured grading: could not parse JSON response \u2014 see Logs'); }
    } catch (e) {
      allOk = false; lines.push('\u274C Structured grading: ' + String(e));
    }
  }

  ss.toast('', 'Autograder', 1);
  ui.alert(
    (allOk ? '\u2705 All tests passed!' : '\u274C Test failed') + '\n\n' +
    'Provider: ' + p + '\nModel: ' + model + '\n\n' +
    lines.join('\n')
  );
}


// ═══════════════════════════════════════════════════════════════════════════════
// 10. UTILITIES
// ═══════════════════════════════════════════════════════════════════════════════

// ── Sheet helpers ──

function getSheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: "' + name + '". Run Autograder \u2192 Initial Setup first.');
  return sh;
}

function headers_(row1) {
  var m = {};
  for (var i = 0; i < row1.length; i++) m[row1[i]] = i;
  return m;
}

function writeRow_(sh, rowNum, head, patch) {
  var row = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(patch).forEach(function(k) {
    if (head[k] !== undefined) row[head[k]] = patch[k];
  });
  sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
}

/**
 * Sets column widths on a sheet by header name → pixel width map.
 */
function setColumnWidths_(sh, widthMap) {
  if (!sh) return;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var c = 0; c < head.length; c++) {
    var name = String(head[c] || '').trim();
    if (widthMap[name]) sh.setColumnWidth(c + 1, widthMap[name]);
  }
}

// ── Levels & Criteria ──

function loadCriteriaByLevel_() {
  var sh = getSheet_(SHEET_CRIT);
  var values = sh.getDataRange().getValues();
  var head = headers_(values[0]);
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var levelId = String(row[head.LevelID] || '').trim();
    if (!levelId) continue;
    (map[levelId] = map[levelId] || []).push({
      LevelID:     levelId,
      CriterionID: row[head.CriterionID],
      Points:      row[head.Points],
      Type:        row[head.Type],
      Description: row[head.Description],
      Notes:       (head.Notes !== undefined ? row[head.Notes] : ''),
      TeacherNotes:(head['Teacher Notes'] !== undefined ? row[head['Teacher Notes']] : '')
    });
  }
  return map;
}

function loadEnabledLevels_() {
  var sh = getSheet_(SHEET_LEVELS);
  var vals = sh.getDataRange().getValues();
  var head = headers_(vals[0]);
  var set = {};
  for (var i = 1; i < vals.length; i++) {
    var id = String(vals[i][head.LevelID] || '').trim();
    var en = String(vals[i][head.Enabled]).toUpperCase() !== 'FALSE';
    if (id) set[id] = en;
  }
  return set;
}

function getModelForLevel_(levelId) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LEVELS);
  if (!sh) return '';
  var vals = sh.getDataRange().getValues();
  var head = headers_(vals[0]);
  for (var i = 1; i < vals.length; i++) {
    var id = String(vals[i][head.LevelID] || '').trim();
    if (id === levelId) {
      return (head.Model !== undefined ? String(vals[i][head.Model] || '').trim() : '');
    }
  }
  return '';
}

/**
 * Converts a LevelID (e.g. "Lesson-05-Level-07") into the Code.org level URL.
 * Special case: Lesson-21-Side-Scroller → lessons/21/levels/2
 */
function levelIdToUrl_(levelId) {
  var base = 'https://studio.code.org/courses/csd-2025/units/3/';
  // Lesson-21-Side-Scroller
  if (/Lesson-21/i.test(levelId)) return base + 'lessons/21/levels/2';
  // Standard pattern: Lesson-NN-Level-NN
  var m = String(levelId).match(/Lesson-(\d+)-Level-(\d+)/i);
  if (m) return base + 'lessons/' + parseInt(m[1], 10) + '/levels/' + parseInt(m[2], 10);
  return '';
}

// ── LLM config ──

function getLLMProvider_() {
  var p = PropertiesService.getScriptProperties().getProperty('LLM_PROVIDER');
  p = String(p || DEFAULT_PROVIDER).trim().toLowerCase();
  if (p !== 'openai' && p !== 'gemini') p = DEFAULT_PROVIDER;
  return p;
}

function getDefaultModel_() {
  return DEFAULT_MODEL_BY_PROVIDER[getLLMProvider_()] || DEFAULT_MODEL_BY_PROVIDER.gemini;
}

// ── Header mapping for form responses (verbose → short names) ──

function headersSmart_(row1) {
  var aliases = {
    Timestamp: ['Timestamp', 'Response Timestamp', 'Submitted at'],
    Email:     ['Email', 'Email Address', 'Email address'],
    First:     ['First', 'First Name', 'Given Name'],
    Last:      ['Last', 'Last Name', 'Family Name', 'Surname'],
    Period:    ['Period', 'Class Period', 'Class', 'Section'],
    LevelID:   ['LevelID', 'Level ID', 'Which assessment level',
                'Which assessment level are you submitting'],
    ShareURL:  ['ShareURL', 'Share URL', 'URL', 'Project URL', 'Project Link',
                'Paste the share URL', 'Paste the URL']
  };
  var map = {};
  for (var c = 0; c < row1.length; c++) {
    var h  = String(row1[c] || '').trim();
    var hl = h.toLowerCase();
    Object.keys(aliases).forEach(function(key) {
      if (map[key] !== undefined) return;
      aliases[key].some(function(alias) {
        if (hl === alias.toLowerCase() || hl.indexOf(alias.toLowerCase()) === 0) {
          map[key] = c;
          return true;
        }
        return false;
      });
    });
  }
  return map;
}

function findFormResponsesSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Form Responses 1') || ss.getSheetByName('Form Responses');
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (/^Form Responses/i.test(all[i].getName())) return all[i];
  }
  return null;
}

function normalizeTimestamp_(v) {
  try {
    var tz = Session.getScriptTimeZone() || 'UTC';
    var d  = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return String(v || '').trim();
    return Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ss");
  } catch (e) {
    return String(v || '').trim();
  }
}

// ── CSV parser ──

function parseCriteriaTableCsv_() {
  var csvText = getCriteriaTableCsvText_();
  if (!csvText || !String(csvText).trim()) throw new Error('Embedded criteria table CSV is empty.');

  csvText = String(csvText).replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  var values = parseCsvText_(csvText);
  if (!values || values.length < 2) throw new Error('criteria-table.csv has no data rows.');

  var header = values[0].map(function(h) { return String(h || '').trim(); });
  var idx = {};
  header.forEach(function(h, i) { idx[h] = i; });
  function col(name) { return idx[name]; }

  var byLevel = {};
  var outRows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row || !row.length) continue;
    var levelId = String(row[col('LevelID')] || '').trim();
    if (!levelId) continue;
    byLevel[levelId] = true;
    outRows.push([
      levelId,
      String(row[col('CriterionID')] || '').trim(),
      row[col('Points')],
      String(row[col('Type')] || '').trim(),
      String(row[col('Description')] || '').trim(),
      String((col('Notes') !== undefined ? (row[col('Notes')] || '') : '') || '').trim(),
      String((col('Teacher Notes') !== undefined ? (row[col('Teacher Notes')] || '') : '') || '').trim()
    ]);
  }
  return { rows: outRows, byLevel: byLevel };
}

function parseCsvText_(text) {
  var rows = [], row = [], field = '', inQuotes = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else { field += ch; }
      continue;
    }
    if (ch === '"') { inQuotes = true; continue; }
    if (ch === ',') { row.push(field); field = ''; continue; }
    if (ch === '\n') {
      row.push(field); field = '';
      if (row.length > 1 || (row.length === 1 && String(row[0]).trim() !== '')) rows.push(row);
      row = []; continue;
    }
    field += ch;
  }
  row.push(field);
  if (row.length > 1 || (row.length === 1 && String(row[0]).trim() !== '')) rows.push(row);
  if (inQuotes) throw new Error('CSV parse error: unmatched quote.');
  return rows;
}

// ── JSON normalization ──

function normalizeAutogradeJson_(text, expectedIds) {
  var out = { unreadable: false, checks: [] };
  text = stripCodeFences_(text);
  var obj;
  try { obj = JSON.parse(text); } catch (e) { return out; }

  if (obj && Array.isArray(obj.checks)) {
    out.unreadable = !!obj.unreadable;
    obj.checks.forEach(function(ch) {
      if (!ch) return;
      out.checks.push({ id: String(ch.id), pass: toBool_(ch.pass), reason: ch.reason ? String(ch.reason) : '' });
    });
    return out;
  }

  var source = (obj && typeof obj.results === 'object' && obj.results) ? obj.results : obj;
  var ids = expectedIds && expectedIds.length ? expectedIds : Object.keys(source || {});
  ids.forEach(function(id) {
    if (!source || !(id in source)) return;
    var v = source[id], pass = false, reason = '';
    if (v && typeof v === 'object') {
      if ('pass' in v) pass = toBool_(v.pass);
      else if ('result' in v) pass = toBool_(v.result);
      else if ('ok' in v) pass = toBool_(v.ok);
      if (v.reason) reason = String(v.reason);
    } else {
      pass = toBool_(v);
    }
    out.checks.push({ id: String(id), pass: pass, reason: reason });
  });
  return out;
}

function stripCodeFences_(s) {
  s = String(s || '').trim();
  if (s.substring(0, 3) === '```') s = s.replace(/^```[\w-]*\s*/i, '').replace(/\s*```$/, '');
  return s.trim();
}

function toBool_(v) {
  if (typeof v === 'boolean') return v;
  var s = String(v || '').trim().toLowerCase();
  return s === 'pass' || s === 'passed' || s === 'true' || s === 'yes' || s === 'y' || s === '1';
}

function toNumber_(v) {
  if (typeof v === 'number') return v;
  var s = String(v).trim();
  var n = Number(s);
  if (!isNaN(n)) return n;
  // Extract trailing number from strings like "Period 7"
  var m = s.match(/(\d+)\s*$/);
  return m ? Number(m[1]) : v;
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function(c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

// ── HTTP fetch with retry (handles 429 rate limits) ──

/**
 * Wrapper around UrlFetchApp.fetch that retries on 429 (rate limit) and 503
 * with exponential backoff. Up to 4 retries (waits ~2s, ~4s, ~8s, ~16s).
 * Total max wait ≈ 30s, well within Apps Script's 6-minute execution limit.
 */
function fetchWithRetry_(url, options) {
  var MAX_RETRIES = 4;
  var baseDelay   = 2000; // 2 seconds

  for (var attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();

    if (code !== 429 && code !== 503) return resp; // success or non-retryable error
    if (attempt === MAX_RETRIES) return resp;       // out of retries, return last response

    // Exponential backoff with jitter: 2s, 4s, 8s, 16s (±25%)
    var delay = baseDelay * Math.pow(2, attempt);
    var jitter = delay * 0.25 * (Math.random() - 0.5); // ±12.5%
    Utilities.sleep(Math.round(delay + jitter));
  }
  return resp; // shouldn't reach here, but just in case
}

// ── Cache helpers (CacheService, 6-hour TTL) ──

function getGradeCache_(key) {
  try { return CacheService.getDocumentCache().get(key); }
  catch (_) { return null; }
}

function setGradeCache_(key, value) {
  try { CacheService.getDocumentCache().put(key, value, 21600); } // 6 hours
  catch (_) {}
}

function sha256_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return raw.map(function(b) { var v = (b < 0 ? b + 256 : b); return ('0' + v.toString(16)).slice(-2); }).join('');
}


// ═══════════════════════════════════════════════════════════════════════════════
// 11. HELP DIALOG
// ═══════════════════════════════════════════════════════════════════════════════

function showHelp() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:1.5;max-width:500px;">' +

    '<h2 style="margin:0 0 8px 0;font-size:16px;">\uD83C\uDFAE Game Lab Autograder v2</h2>' +
    '<p style="margin:0 0 12px 0;color:#555;">Automatically grades Code.org Game Lab projects using AI.</p>' +

    '<h3 style="margin:12px 0 6px 0;font-size:14px;">\uD83D\uDE80 Getting Started</h3>' +
    '<ol style="margin:0 0 12px 18px;padding:0;">' +

    '<li><b>Run Initial Setup</b> from the Autograder menu. Check the periods you teach.</li>' +

    '<li><b>Set your API key:</b><br>' +
    'Go to <b>Extensions \u2192 Apps Script \u2192 \u2699\uFE0F Project Settings \u2192 Script Properties</b><br>' +
    'Add: <code>GEMINI_API_KEY</code> = your key<br>' +
    '<span style="color:#666;">Get a free key at <a href="https://aistudio.google.com" target="_blank">aistudio.google.com</a></span><br>' +
    '<span style="color:#666;">(Optional: set <code>LLM_PROVIDER</code> = <code>openai</code> and add <code>OPENAI_API_KEY</code>)</span></li>' +

    '<li><b>Test your connection:</b> Use <b>Test API Connection</b> from the Autograder menu.</li>' +

    '<li><b>Create a Google Form</b> with these fields:<br>' +
    '<table style="font-size:12px;border-collapse:collapse;margin:4px 0 4px 0;">' +
    '<tr><td style="padding:2px 8px;border:1px solid #ddd;">Email Address</td><td style="padding:2px 8px;border:1px solid #ddd;">Settings \u2192 Collect email addresses</td></tr>' +
    '<tr><td style="padding:2px 8px;border:1px solid #ddd;">First Name</td><td style="padding:2px 8px;border:1px solid #ddd;">Short answer</td></tr>' +
    '<tr><td style="padding:2px 8px;border:1px solid #ddd;">Last Name</td><td style="padding:2px 8px;border:1px solid #ddd;">Short answer</td></tr>' +
    '<tr><td style="padding:2px 8px;border:1px solid #ddd;">Class Period</td><td style="padding:2px 8px;border:1px solid #ddd;">Dropdown: 1, 2, 3\u2026 (match your periods)</td></tr>' +
    '<tr><td style="padding:2px 8px;border:1px solid #ddd;">Assessment Level</td><td style="padding:2px 8px;border:1px solid #ddd;">Dropdown: Lesson-03-Level-08, etc.</td></tr>' +
    '<tr><td style="padding:2px 8px;border:1px solid #ddd;">Share URL</td><td style="padding:2px 8px;border:1px solid #ddd;">Short answer (the Code.org share link)</td></tr>' +
    '</table></li>' +

    '<li><b>Link the form to this spreadsheet:</b><br>' +
    'In the Form editor \u2192 <b>Responses</b> tab \u2192 click the green Sheets icon \u2192 <b>Select existing spreadsheet</b> \u2192 pick this spreadsheet.</li>' +

    '<li><b>Set up the auto-grade trigger:</b><br>' +
    'In <b>Extensions \u2192 Apps Script</b>, click the \u23F0 <b>Triggers</b> icon (left sidebar)<br>' +
    'Click <b>+ Add Trigger</b><br>' +
    'Function: <code>onFormSubmit</code> | Source: <b>From spreadsheet</b> | Event: <b>On form submit</b><br>' +
    'Leave "Which deployment should run" set to <b>Head</b><br>' +
    'Click Save and authorize when prompted.</li>' +

    '<li><b>Done!</b> When a student submits the form, their code is automatically graded and they receive an email with their score.</li>' +
    '</ol>' +

    '<h3 style="margin:12px 0 6px 0;font-size:14px;">\uD83D\uDCCB Menu Reference</h3>' +
    '<ul style="margin:0 0 12px 18px;padding:0;">' +
    '<li><b>Grade New Submissions</b> \u2014 imports new form responses (if any), then grades all ungraded rows in Submissions</li>' +
    '<li><b>Re-grade Selected Rows</b> \u2014 re-grades the rows you highlight in Submissions (e.g., after editing criteria)</li>' +
    '<li><b>Re-grade All Rows</b> \u2014 re-grades every row in Submissions (slow, uses API credits)</li>' +
    '<li><b>Grade & Email All New</b> \u2014 imports, grades, and emails results in one step</li>' +
    '<li><b>Email Selected Rows</b> \u2014 sends result emails for rows you highlight in Submissions</li>' +
    '</ul>' +

    '<h3 style="margin:12px 0 6px 0;font-size:14px;">\uD83D\uDCC4 Sheet Reference</h3>' +
    '<ul style="margin:0 0 12px 18px;padding:0;">' +
    '<li><b>Submissions</b> \u2014 all student submissions and grades (the main data sheet)</li>' +
    '<li><b>Grade View P#</b> \u2014 read-only views filtered by period, sorted by level then name</li>' +
    '<li><b>Levels</b> \u2014 enable/disable levels, set per-level model overrides</li>' +
    '<li><b>Criteria</b> \u2014 rubric criteria (auto-populated; you can edit descriptions/points)</li>' +
    '</ul>' +

    '</div>'
  ).setWidth(560).setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Autograder Help');
}


// ═══════════════════════════════════════════════════════════════════════════════
// 12. EMBEDDED CRITERIA CSV
// ═══════════════════════════════════════════════════════════════════════════════
// This mirrors criteria-table.csv so setupSheets() can populate rubrics
// without needing an external file. Keep in sync with the repo CSV.

function getCriteriaTableCsvText_() {
  return (
"LevelID,CriterionID,Points,Type,Description,Notes,Teacher Notes\n" +
"Lesson-03-Level-08,ellipses_present,1,llm_check,\"The four orange ellipses are present in a 2x2 grid with 50 pixel spacing. Count it as OK if each ellipse is drawn while the current fill is orange (literal \"\"orange\"\" or a variable set to \"\"orange\"\").\",,\n" +
"Lesson-03-Level-08,purple_rect_after_fill,4,llm_check,\"A rectangle is drawn after a call that sets fill to purple (either the literal \"\"purple\"\" or a variable assigned \"\"purple\"\"). The rectangle call may be rect(x,y) or rect(x,y,width,height); rectangle x and y coordinates must be within 10 pixels of the coordinates of the first ellipse (the ellipse in the top-left corner of the 2x2 grid)\",,\n" +
"Lesson-03-Level-08,rect_after_ellipses,4,llm_check,The (purple) rectangle's draw call appears later in the code than the four orange ellipse calls (so it is on top by draw order). It does not have to be the last shape in the program.,,\n" +
"Lesson-03-Level-08,one_rect_present,1,llm_check,There is exactly one rectangle present (a single call of the rect() function),,\n" +
"Lesson-04-Level-08,cloud_wider_than_tall,9,llm_check,\"After setting fill(\"\"white\"\") for the cloud, the cloud ellipse's width is greater than its height (any numeric values OK as long as clearly wider).\",,\n" +
"Lesson-04-Level-08,no_old_tall_cloud,1,llm_check,\"The original tall cloud ellipse at (150,100,100,200) is not left unchanged.\",,\n" +
"Lesson-05-Level-07,left_uses_eyeSize,1,llm_check,The left eye draws an ellipse whose width and height are controlled by the variable eyeSize (not numeric literals).,,\n" +
"Lesson-05-Level-07,right_uses_eyeSize,8,llm_check,The right eye draws an ellipse whose width and height are controlled by the variable eyeSize (not numeric literals).,,\n" +
"Lesson-05-Level-07,eyeSize_declared,1,llm_check,The program declares a variable named eyeSize and assigns it a number.,,\n" +
"Lesson-05-Level-07,pupil_bonus_scale,1,llm_check,The eyes contain pupils that consist of a second pair of ellipses is drawn in the same location as the eyes. The width and height of these pupil ellipses are always smaller than the eyes (either by dividing eyeSize by a number greater than 1 or multiplying eyeSize by a number less than 1). ,,\n" +
"Lesson-05-Level-07,pupil_bonus_fill,1,llm_check,The pupil ellipses (the smaller ones drawn in the same location as the eyes) have a fill color different from the eyes.,,\n" +
"Lesson-06-Level-07,at_least_3_new_circles,2,llm_check,The program draws at least six circles total (ellipse calls where width == height). This implies the student added at least three new circles beyond the original three.,,\n" +
"Lesson-06-Level-07,all_unique_colors,2,llm_check,\"All circles (the original three plus the new three that the student added) have distinct fill colors. Accept named colors (e.g., \"\"green\"\") or color codes (e.g., \"\"#00ff00\"\", rgb(), hsl()).\",,\n" +
"Lesson-06-Level-07,new_circles_random_y,3,llm_check,\"For every added circle, the ellipse's Y coordinate is controlled by randomNumber(...).  Matching the numeric range of 190\u2013210 is ideal, but credit should be awarded as long as the low value is not less than180 and the high value is not greater than 220.\",,\n" +
"Lesson-06-Level-07,new_circles_correct_x,3,llm_check,\"Each added circle is exactly 40 pixels to the right of the previous one. The proper x-coordinates of all circles, from left to right, should be: 100, 140, 180, 220, 260, 300. If additional circles are optionally added, the pattern should continue (340, 380, 420, etc.)\",,\n" +
"Lesson-08-Level-10,at_least_2_sprites,4,llm_check,At least two sprites are present in the program.,,\n" +
"Lesson-08-Level-10,different_locations,2,llm_check,The sprites are in different locations (they don't have the exact same x and y coordinates).,,\n" +
"Lesson-08-Level-10,sprite_animations,4,llm_check,All sprites have been assigned an animation using the sprite.setAnimation() method. The animation name (the string parameter used with the setAnimation method) should be different for each sprite.,,\n" +
"Lesson-09-Level-05,burger_size,4,llm_check,burger sprite has been assigned a .scale property value of less than 1,,\n" +
"Lesson-09-Level-05,fries_size,3,llm_check,fries sprite has been assigned a .scale property value of less than 1,,\n" +
"Lesson-09-Level-05,dessert_size,3,llm_check,dessert sprite has been assigned a .scale property value of less than 1,,\n" +
"Lesson-10-Level-05,two_texts,5,llm_check,\"The program draws at least two separate text elements using text(string, x, y), and each string is non-empty. Positions/colors are flexible.\",,\n" +
"Lesson-10-Level-05,two_positions,5,llm_check,The two text elements appear at different positions (they don't have the same x / y coordinates).,,\n" +
"Lesson-10-Level-05,text_stroke,2,llm_check,\"At least one text element is rendered in a high-contrast outlined style: either fill set to black with stroke set to white OR fill set to white with stroke set to black. Accept any equivalent color notation (e.g., \"\"black\"\"/\"\"white\"\", #000/#fff, #000000/#ffffff, rgb(0,0,0)/rgb(255,255,255), case-insensitive). The stroke/fill must be in effect when text(...) is drawn; strokeWeight \u2265 1 may be explicit or default.\",,\n" +
"Lesson-12-Level-07,has_draw_loop,3,llm_check,The program defines a draw() function that runs continuously.,,\n" +
"Lesson-12-Level-07,salt_y_randomized,3,llm_check,\"Inside draw(), the salt sprite's Y position is randomized each frame (e.g., salt.y = randomNumber(...)). Any numeric range is fine.\",,\n" +
"Lesson-12-Level-07,drawSprites_in_draw,2,llm_check,drawSprites() is called within draw() (not only once before the loop).,,\n" +
"Lesson-12-Level-07,background_in_draw,2,llm_check,background(...) is called within draw() so the screen is cleared each frame (no trails).,,\n" +
"Lesson-13-Level-07,all_fish_moving,4,llm_check,\"All three sprites (orangeFish, blueFish, greenFish) are moving (their x positions are changing over time within the draw() function)\",,\n" +
"Lesson-13-Level-07,fish_different_speeds,4,llm_check,\"The three sprites (orangeFish, blueFish, greenFish) are all moving at different speeds (their x positions are changing by different values within the draw() function)\",,\n" +
"Lesson-13-Level-07,blue_fish_faster,1,llm_check,The blueFish sprite is moving faster than the other two fish (its x position is changing more per frame than the other two),,\n" +
"Lesson-13-Level-07,green_fish_slower,1,llm_check,The greenFish sprite is moving slower than the other two fish (its x position is changing less per frame than the other two),,\n" +
"Lesson-15-Level-07,dinosaur_transforms,10,llm_check,\"There is an if-statement that checks to see whether the y position of the dinosaur sprite is less than a certain value (any number between 50 and 300 is acceptable), and if so, uses dinosaur.setAnimation(...) to change the dinosaur's animation\",,\n" +
"Lesson-16-Level-06,flyer_movement_left,2.5,llm_check,\"There is an if-statement that checks to see whether the left arrow is being pressed (the letter 'a' is also acceptable) and if so, moves the sprite to the left by reducing its x position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"Lesson-16-Level-06,flyer_movement_right,2.5,llm_check,\"There is an if-statement that checks to see whether the right arrow is being pressed (the letter 'd' is also acceptable) and if so, moves the sprite to the right by increasing its x position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"Lesson-16-Level-06,flyer_movement_up,2.5,llm_check,\"There is an if-statement that checks to see whether the up arrow is being pressed (the letter 'w' is also acceptable) and if so, moves the sprite up by reducing its y position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"Lesson-16-Level-06,flyer_movement_down,2.5,llm_check,\"There is an if-statement that checks to see whether the down arrow is being pressed (the letter 's' is also acceptable) and if so, moves the sprite down by increasing its y position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"Lesson-17-Level-07,shake_when_clicking,4,llm_check,The creature only shakes while the mouse is being pressed (the line of code that randomizes the rotation property of the creature sprite should be inside of an if (mouseDown(...) conditional statement).,,\n" +
"Lesson-17-Level-07,text_in_else_statement,4,llm_check,\"The text function that writes \"\"Press the mouse to shake the creature.\"\" to the screen runs only while the mouse is NOT being pressed (the text function should be placed in an else statement that runs when mouseDown(...) returns false)\",,\n" +
"Lesson-17-Level-07,drawSprites_first,2,llm_check,The drawSprites() function must be called in the draw() loop BEFORE the if-else statement (otherwise the sprite background will cover the text and you won't be able to see it),,\n" +
"Lesson-19-Level-09,fish_starts_with_key,2,llm_check,\"When the right arrow key is pressed, the fish starts moving to the right (there is an if-statement with keyWentDown(...) that sets fish.velocityX to a positive number. It's ok if the code responds to the left arrow key instead, or if both options are present.\",,\n" +
"Lesson-19-Level-09,fish_moves_right_to_left,3,llm_check,\"When the fish reaches the right side of the screen, it starts moving left (there is an if-statement that checks whether fish.x is greater than 400, and if so, sets fish.velocityX to a negative number)\",,\n" +
"Lesson-19-Level-09,fish_moves_left_to_right,3,llm_check,\"When the fish reaches the left side of the screen, it starts moving right (there is an if-statement that checks whether fish.x is less than 0, and if so, sets fish.velocityX to a positive number)\",,\n" +
"Lesson-19-Level-09,fish_faces_move_dir,2,llm_check,\"The fish is always facing the direction that it's swimming (when fish.velocityX changes to a negative number, the fish's animation is set to \"\"fishL\"\", and when fish.velocityX changes to a positive number, the fish's animation is set to \"\"fishR\"\")\",,\n" +
"Lesson-20-Level-07,horse_changes_to_unicorn,10,llm_check,\"When the rainbow sprite collides with the horse sprite, the horse's animation changes to a unicorn.\",,\n" +
"Lesson-21-Side-Scroller,background_in_draw,3,llm_check,\"Game has a background. This can be either a solid color using the background() function OR a sprite with a full-screen animation image assigned that is created before all others. If a sprite background is used, the background() function does NOT need to be present to receive credit.\",,\n" +
"Lesson-21-Side-Scroller,player_sprite,5,llm_check,\"Game has a player sprite. Ideally, the sprite is named 'player', but it may be called something else (e.g.: frog, alien, etc.)\",,\n" +
"Lesson-21-Side-Scroller,player_jump,4,llm_check,\"This is a jump button that cause the player to move upwards. Button can be up-arrow, spacebar, or w.\",,\n" +
"Lesson-21-Side-Scroller,player_ceiling,3,llm_check,The player sprite does not fly off of the top of the screen,,\n" +
"Lesson-21-Side-Scroller,player_floor,3,llm_check,The player sprite does not fall off of the bottom of the screen,,\n" +
"Lesson-21-Side-Scroller,obstacle_sprite,2,llm_check,\"Game has an obstacle sprite. Ideally, the sprite is named 'obstacle', but it may be called something else (e.g.: mushroom, spike, enemy, etc.)\",,\n" +
"Lesson-21-Side-Scroller,obstacle_movement,2,llm_check,\"The obstacle sprite has a negative velocityX value, causing it to constantly move from right to left\",,\n" +
"Lesson-21-Side-Scroller,obstacle_looping,2,llm_check,\"When the obstacle sprite reaches the left edge of the screen, it jumps back to the right edge of the screen, causing it to \"\"loop\"\" indefinitely\",,\n" +
"Lesson-21-Side-Scroller,obstacle_collision,2,llm_check,\"When the player touches the obstacle, the health variable goes down\",,\n" +
"Lesson-21-Side-Scroller,target_sprite,2,llm_check,\"Game has a collectible item sprite. Ideally, the sprite is named 'target', but it may be called something else (e.g.: fly, coin, etc.)\",,\n" +
"Lesson-21-Side-Scroller,target_sprite_movement,2,llm_check,\"The target sprite has a negative velocityX value, causing it to constantly move from right to left\",,\n" +
"Lesson-21-Side-Scroller,target_looping,2,llm_check,\"When the target sprite reaches the left edge of the screen, it jumps back to the right edge of the screen, causing it to \"\"loop\"\" indefinitely\",,\n" +
"Lesson-21-Side-Scroller,target_collision,2,llm_check,\"When the player touches the target, the score variable goes up\",,\n" +
"Lesson-21-Side-Scroller,score_display,2,llm_check,There is a score-counter displayed somewhere on the screen that properly displays the value of the 'score' variable,,\n" +
"Lesson-21-Side-Scroller,health_display,2,llm_check,There is a health-counter displayed somewhere on the screen that properly displays the value of the 'health' variable,,\n" +
"Lesson-21-Side-Scroller,game_over,2,llm_check,\"When the 'health' variable reaches 0, a Game Over message is displayed. This can be either a background() function with a text() function message, or a sprite-based image\",,\n" +
"Lesson-22-Level-06,rocks_falls_back_down,10,llm_check,\"The rock, which is initially moving upward, slows its ascent and then falls back down. This is accomplished by increasing rock.velocityY each frame.\",,\n"
  );
}
