// ============================================================
//  IELTS Simulation — Google Apps Script Backend v4.0
// ============================================================

const SHEET_CANDIDATES = 'Candidates';
const SHEET_SCORES     = 'Scores';
const SHEET_SETS       = 'TestSets';
const SHEET_SETTINGS   = 'Settings';
const SHEET_OTP        = 'OTP_Tokens';

const API_SECRET = 'ielts-sim-x7k2m9-change-me';

// ============================================================
//  CANONICAL SCHEMAS
// ============================================================
const CANDIDATE_HEADERS = [
  'CandidateID',     // 0
  'FullName',        // 1
  'DateOfBirth',     // 2
  'Gender',          // 3
  'Nationality',     // 4
  'Email',           // 5
  'Password',        // 6
  'CreatedAt',       // 7
  'LastLogin',       // 8
  'PhotoBase64'      // 9
];

const SCORE_HEADERS = [
  'SubmissionID',     // 0
  'Timestamp',        // 1
  'CandidateID',      // 2
  'FullName',         // 3
  'TestSetID',        // 4
  'TestSetName',      // 5
  'ListeningScore',   // 6
  'ListeningBand',    // 7
  'ReadingScore',     // 8
  'ReadingBand',      // 9
  'WritingTask1Band', // 10
  'WritingTask2Band', // 11
  'WritingBand',      // 12
  'SpeakingBand',     // 13
  'OverallBand',      // 14
  'WritingTask1WC',   // 15
  'WritingTask2WC',   // 16
  'WritingTask1Text', // 17
  'WritingTask2Text', // 18
  'ExaminerNotes',    // 19
  'GradedBy',         // 20
  'GradedAt',         // 21
  'ListeningAnswers', // 22
  'ReadingAnswers'    // 23
];

const TESTSET_HEADERS = [
  'SetID',
  'SetName',
  'ConfigURL',
  'Status',
  'CreatedAt',
  'Description'
];

const SETTINGS_HEADERS = ['Key', 'Value'];
const OTP_HEADERS      = ['Email', 'OTP', 'Purpose', 'CandidateID', 'Expires'];

const CIDX = {
  ID: 0,
  FULL_NAME: 1,
  DOB: 2,
  GENDER: 3,
  NATIONALITY: 4,
  EMAIL: 5,
  PASSWORD: 6,
  CREATED_AT: 7,
  LAST_LOGIN: 8,
  PHOTO: 9
};

const SIDX = {
  SUBMISSION_ID: 0,
  TIMESTAMP: 1,
  CANDIDATE_ID: 2,
  FULL_NAME: 3,
  TESTSET_ID: 4,
  TESTSET_NAME: 5,
  LISTENING_SCORE: 6,
  LISTENING_BAND: 7,
  READING_SCORE: 8,
  READING_BAND: 9,
  WRITING_TASK1_BAND: 10,
  WRITING_TASK2_BAND: 11,
  WRITING_BAND: 12,
  SPEAKING_BAND: 13,
  OVERALL_BAND: 14,
  WRITING_TASK1_WC: 15,
  WRITING_TASK2_WC: 16,
  WRITING_TASK1_TEXT: 17,
  WRITING_TASK2_TEXT: 18,
  EXAMINER_NOTES: 19,
  GRADED_BY: 20,
  GRADED_AT: 21,
  LISTENING_ANSWERS: 22,
  READING_ANSWERS: 23
};

// ============================================================
//  SETUP
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders_(ss, SHEET_CANDIDATES, CANDIDATE_HEADERS, '#002f5f');
  ensureCandidateMigration_(ss.getSheetByName(SHEET_CANDIDATES));

  ensureSheetWithHeaders_(ss, SHEET_SCORES, SCORE_HEADERS, '#002f5f');
  ensureSheetWithHeaders_(ss, SHEET_SETS, TESTSET_HEADERS, '#002f5f');
  ensureSheetWithHeaders_(ss, SHEET_SETTINGS, SETTINGS_HEADERS, '#002f5f');
  ensureSheetWithHeaders_(ss, SHEET_OTP, OTP_HEADERS, '#002f5f');

  seedDefaultSettings_();
  return 'All sheets ready! (v4.0)';
}

function ensureSheetWithHeaders_(ss, name, headers, color) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else {
    const lastCol = Math.max(sheet.getLastColumn(), headers.length);
    const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    } else {
      for (let i = 0; i < headers.length; i++) {
        if (existing[i] !== headers[i]) {
          sheet.getRange(1, i + 1).setValue(headers[i]);
        }
      }
    }
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(color || '#002f5f')
    .setFontColor('#ffffff');

  return sheet;
}

function ensureCandidateMigration_(sheet) {
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 10)).getValues()[0].map(String);

  // Old format: A=CandidateID B=FullName C=DOB D=Gender E=Nationality F=Password...
  if (headers[5] === 'Password') {
    sheet.insertColumnAfter(5);
    sheet.getRange(1, 6).setValue('Email');
    sheet.getRange(1, 7).setValue('Password');
    sheet.getRange(1, 8).setValue('CreatedAt');
    sheet.getRange(1, 9).setValue('LastLogin');
    sheet.getRange(1, 10).setValue('PhotoBase64');
  }
}

function seedDefaultSettings_() {
  const defaults = {
    institution_name: 'IELTS Simulation Center',
    institution_address: '',
    active_set_id: '',
    trf_footer_note: 'This is a simulation test result for preparation purposes only.',
    notify_email: Session.getActiveUser().getEmail() || ''
  };

  Object.keys(defaults).forEach(key => {
    if (!getSettingValue(key)) saveSettingValue(key, defaults[key]);
  });

  if (!getSettingValue('admin_password')) {
    saveSettingValue('admin_password', hashPassword('admin1234'));
  }
}

// ============================================================
//  ROUTER
// ============================================================
function doGet(e) {
  const p = e.parameter || {};
  const callback = p.callback || '';
  const action = p.action || '';
  let result;

  try {
    const publicActions = [
      'ping',
      'checkCandidate',
      'registerCandidate',
      'loginCandidate',
      'saveScore',
      'getTestSets',
      'getActiveSet',
      'getConfig',
      'getCandidateScores',
      'verifyAdmin',
      'updateProfile',
      'updateProfilePhoto',
      'sendOTP',
      'verifyOTPAndRegister',
      'verifyOTPAndResetPassword',
      'changePassword',
      'verifyOTPAndUpdateEmail'
    ];

    if (!publicActions.includes(action) && !checkApiKey(p)) {
      result = { success: false, message: 'Unauthorized — invalid or missing API key.' };
    } else {
      switch (action) {
        case 'ping':                      result = { success: true, message: 'pong', time: new Date().toISOString() }; break;
        case 'setup':                     result = { success: true, message: setupSheets() }; break;

        case 'checkCandidate':            result = checkCandidate(p); break;
        case 'registerCandidate':         result = registerCandidate(p); break;
        case 'loginCandidate':            result = loginCandidate(p); break;
        case 'updateProfile':             result = updateProfile(p); break;
        case 'updateProfilePhoto':        result = updateProfilePhoto(p); break;

        case 'saveScore':                 result = saveScore(p); break;
        case 'getScores':                 result = getScores(p); break;
        case 'getCandidateScores':        result = getCandidateScores(p); break;
        case 'gradeWriting':              result = gradeWriting(p); break;

        case 'getCandidates':             result = getCandidates(); break;

        case 'getTestSets':               result = getTestSets(); break;
        case 'saveTestSet':               result = saveTestSet(p); break;
        case 'deleteTestSet':             result = deleteTestSet(p); break;
        case 'setActiveSet':              result = setActiveSet(p); break;
        case 'getActiveSet':              result = getActiveSet(); break;

        case 'getSettings':               result = getSettings(); break;
        case 'saveSettings':              result = saveSettings(p); break;

        case 'verifyAdmin':               result = verifyAdmin(p); break;
        case 'changeAdminPassword':       result = changeAdminPassword(p); break;

        case 'sendOTP':                   result = handleSendOTP(p); break;
        case 'verifyOTPAndRegister':      result = handleVerifyOTPAndRegister(p); break;
        case 'verifyOTPAndResetPassword': result = handleVerifyOTPAndResetPassword(p); break;
        case 'changePassword':            result = handleChangePassword(p); break;
        case 'verifyOTPAndUpdateEmail':   result = handleVerifyOTPAndUpdateEmail(p); break;

        default:
          result = { success: false, message: 'Unknown action: ' + action };
      }
    }
  } catch (err) {
    result = { success: false, message: 'Error: ' + err.toString() };
  }

  const json = JSON.stringify(result);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

function checkApiKey(p) {
  return p.key === API_SECRET;
}

// ============================================================
//  CANDIDATE ACTIONS
// ============================================================
function checkCandidate(p) {
  const row = findCandidate(p.candidateID);
  if (row) {
    return {
      success: true,
      exists: true,
      name: String(row[CIDX.FULL_NAME] || ''),
      nationality: String(row[CIDX.NATIONALITY] || '')
    };
  }
  return { success: true, exists: false };
}

function registerCandidate(p) {
  if (!p.candidateID || !p.fullName || !p.password) {
    return { success: false, message: 'Missing required fields.' };
  }

  if (findCandidate(p.candidateID)) {
    return { success: false, message: 'Candidate ID already registered.' };
  }

  const sheet = getOrCreateSheet_(SHEET_CANDIDATES, CANDIDATE_HEADERS);
  const now = new Date().toISOString();

  sheet.appendRow([
    String(p.candidateID),
    String(p.fullName),
    String(p.dob || ''),
    String(p.gender || ''),
    String(p.nationality || ''),
    String(p.email || '').trim().toLowerCase(),
    hashPassword(String(p.password)),
    now,
    now,
    ''
  ]);

  return {
    success: true,
    candidateID: p.candidateID,
    fullName: p.fullName
  };
}

function loginCandidate(p) {
  if (!p.candidateID || !p.password) {
    return { success: false, message: 'Missing credentials.' };
  }

  const row = findCandidate(p.candidateID);
  if (!row) return { success: false, message: 'Candidate ID not found.' };

  if (String(row[CIDX.PASSWORD] || '') !== hashPassword(String(p.password))) {
    return { success: false, message: 'Incorrect password.' };
  }

  updateLastLogin(p.candidateID);

  return {
    success: true,
    candidateID: String(row[CIDX.ID] || ''),
    fullName: String(row[CIDX.FULL_NAME] || ''),
    dob: String(row[CIDX.DOB] || ''),
    gender: String(row[CIDX.GENDER] || ''),
    nationality: String(row[CIDX.NATIONALITY] || ''),
    email: String(row[CIDX.EMAIL] || ''),
    photoBase64: String(row[CIDX.PHOTO] || '')
  };
}

function updateProfile(p) {
  if (!p.candidateID) return { success: false, message: 'No candidateID.' };

  const sheet = getSheet_(SHEET_CANDIDATES);
  if (!sheet) return { success: false, message: 'Candidates sheet not found.' };

  const data = sheet.getDataRange().getValues();
  const id = normalize_(p.candidateID);

  for (let i = 1; i < data.length; i++) {
    if (normalize_(data[i][CIDX.ID]) === id) {
      const row = i + 1;

      if (p.fullName !== undefined)    sheet.getRange(row, CIDX.FULL_NAME + 1).setValue(String(p.fullName));
      if (p.dob !== undefined)         sheet.getRange(row, CIDX.DOB + 1).setValue(String(p.dob));
      if (p.gender !== undefined)      sheet.getRange(row, CIDX.GENDER + 1).setValue(String(p.gender));
      if (p.nationality !== undefined) sheet.getRange(row, CIDX.NATIONALITY + 1).setValue(String(p.nationality));
      if (p.email !== undefined)       sheet.getRange(row, CIDX.EMAIL + 1).setValue(String(p.email).trim().toLowerCase());

      return { success: true, message: 'Profile updated.' };
    }
  }

  return { success: false, message: 'Candidate not found.' };
}

function updateProfilePhoto(p) {
  if (!p.candidateID) return { success: false, message: 'No candidateID.' };
  if (p.photoBase64 && String(p.photoBase64).length > 45000) {
    return { success: false, message: 'Image too large even after compression. Try a smaller photo.' };
  }

  const sheet = getSheet_(SHEET_CANDIDATES);
  if (!sheet) return { success: false, message: 'Candidates sheet not found.' };

  const data = sheet.getDataRange().getValues();
  const id = normalize_(p.candidateID);

  for (let i = 1; i < data.length; i++) {
    if (normalize_(data[i][CIDX.ID]) === id) {
      sheet.getRange(i + 1, CIDX.PHOTO + 1).setValue(String(p.photoBase64 || ''));
      return { success: true, message: 'Photo updated.' };
    }
  }

  return { success: false, message: 'Candidate not found.' };
}

// ============================================================
//  SCORE ACTIONS
// ============================================================
function saveScore(p) {
  const sheet = getOrCreateSheet_(SHEET_SCORES, SCORE_HEADERS);
  const subID = 'SUB-' + Date.now();

  sheet.appendRow([
    subID,
    new Date().toISOString(),
    String(p.candidateID || ''),
    String(p.fullName || ''),
    String(p.testSetID || ''),
    String(p.testSetName || ''),
    toNumber_(p.listeningScore, 0),
    toNumber_(p.listeningBand, 0),
    toNumber_(p.readingScore, 0),
    toNumber_(p.readingBand, 0),
    '',
    '',
    '',
    '',
    '',
    toNumber_(p.writingTask1WC, 0),
    toNumber_(p.writingTask2WC, 0),
    String(p.writingTask1 || '').substring(0, 2000),
    String(p.writingTask2 || '').substring(0, 2000),
    '',
    '',
    '',
    String(p.listeningAnswers || ''),
    String(p.readingAnswers || '')
  ]);

  return { success: true, submissionID: subID };
}

function getScores(p) {
  const sheet = getSheet_(SHEET_SCORES);
  if (!sheet) return { success: true, rows: [], headers: SCORE_HEADERS };

  const data = sheet.getDataRange().getValues();
  let rows = data.slice(1).map(stringifyRow_);

  if (p && p.testSetID) {
    rows = rows.filter(r => String(r[SIDX.TESTSET_ID]) === String(p.testSetID));
  }

  return { success: true, rows, headers: SCORE_HEADERS };
}

function getCandidateScores(p) {
  if (!p.candidateID) return { success: false, message: 'No candidateID.' };

  const sheet = getSheet_(SHEET_SCORES);
  if (!sheet) return { success: true, rows: [] };

  const data = sheet.getDataRange().getValues();
  const id = normalize_(p.candidateID);

  const rows = data
    .slice(1)
    .filter(r => normalize_(r[SIDX.CANDIDATE_ID]) === id)
    .map(stringifyRow_);

  return { success: true, rows };
}

function gradeWriting(p) {
  if (!p.submissionID) return { success: false, message: 'No submissionID.' };

  const sheet = getSheet_(SHEET_SCORES);
  if (!sheet) return { success: false, message: 'Scores sheet not found.' };

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][SIDX.SUBMISSION_ID]) === String(p.submissionID)) {
      const rowNum = i + 1;

      const w1Band = toFloat_(p.writingTask1Band, 0);
      const w2Band = toFloat_(p.writingTask2Band, 0);
      const sBand  = toFloat_(p.speakingBand, 0);

      const wBand = roundHalf_((w1Band + w2Band) / 2);

      sheet.getRange(rowNum, SIDX.WRITING_TASK1_BAND + 1).setValue(w1Band || '');
      sheet.getRange(rowNum, SIDX.WRITING_TASK2_BAND + 1).setValue(w2Band || '');
      sheet.getRange(rowNum, SIDX.WRITING_BAND + 1).setValue(wBand || '');
      sheet.getRange(rowNum, SIDX.SPEAKING_BAND + 1).setValue(sBand || '');

      const lBand = toFloat_(data[i][SIDX.LISTENING_BAND], 0);
      const rBand = toFloat_(data[i][SIDX.READING_BAND], 0);

      const components = [lBand, rBand, wBand];
      if (sBand > 0) components.push(sBand);

      const avg = components.reduce((a, b) => a + b, 0) / components.length;
      const overall = roundHalf_(avg);

      sheet.getRange(rowNum, SIDX.OVERALL_BAND + 1).setValue(overall);
      sheet.getRange(rowNum, SIDX.EXAMINER_NOTES + 1).setValue(String(p.examinerNotes || ''));
      sheet.getRange(rowNum, SIDX.GRADED_BY + 1).setValue(String(p.gradedBy || 'Admin'));
      sheet.getRange(rowNum, SIDX.GRADED_AT + 1).setValue(new Date().toISOString());

      return {
        success: true,
        overallBand: overall,
        writingBand: wBand
      };
    }
  }

  return { success: false, message: 'Submission not found.' };
}

// ============================================================
//  TEST SETS
// ============================================================
function getTestSets() {
  const sheet = getSheet_(SHEET_SETS);
  if (!sheet) return { success: true, sets: [] };

  const data = sheet.getDataRange().getValues();
  const sets = data.slice(1).map(r => ({
    setID: String(r[0] || ''),
    setName: String(r[1] || ''),
    configURL: String(r[2] || ''),
    status: String(r[3] || ''),
    createdAt: String(r[4] || ''),
    description: String(r[5] || '')
  }));

  return { success: true, sets };
}

function saveTestSet(p) {
  const sheet = getOrCreateSheet_(SHEET_SETS, TESTSET_HEADERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.setID)) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([[
        String(p.setID),
        String(p.setName || ''),
        String(p.configURL || ''),
        String(p.status || 'Active'),
        String(data[i][4] || ''),
        String(p.description || '')
      ]]);
      return { success: true, message: 'Updated.' };
    }
  }

  const setID = p.setID || ('SET-' + Date.now());
  sheet.appendRow([
    setID,
    String(p.setName || ''),
    String(p.configURL || ''),
    String(p.status || 'Active'),
    new Date().toISOString(),
    String(p.description || '')
  ]);

  return { success: true, setID, message: 'Created.' };
}

function deleteTestSet(p) {
  if (!p.setID) return { success: false, message: 'No setID.' };

  const sheet = getSheet_(SHEET_SETS);
  if (!sheet) return { success: false, message: 'TestSets sheet not found.' };

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.setID)) {
      sheet.deleteRow(i + 1);

      if (getSettingValue('active_set_id') === String(p.setID)) {
        saveSettingValue('active_set_id', '');
      }

      return { success: true, message: 'Deleted.' };
    }
  }

  return { success: false, message: 'Test set not found.' };
}

function setActiveSet(p) {
  if (!p.setID) return { success: false, message: 'No setID.' };
  saveSettingValue('active_set_id', String(p.setID));
  return { success: true };
}

function getActiveSet() {
  const id = getSettingValue('active_set_id');
  if (!id) return { success: true, set: null };

  const all = getTestSets();
  const found = (all.sets || []).find(s => s.setID === id);

  return { success: true, set: found || null };
}

// ============================================================
//  SETTINGS
// ============================================================
function getSettings() {
  const sheet = getSheet_(SHEET_SETTINGS);
  if (!sheet) return { success: true, settings: {} };

  const data = sheet.getDataRange().getValues();
  const settings = {};

  data.slice(1).forEach(r => {
    const key = String(r[0] || '');
    if (key === 'admin_password') return;
    settings[key] = String(r[1] || '');
  });

  return { success: true, settings };
}

function saveSettings(p) {
  const keys = [
    'institution_name',
    'institution_address',
    'trf_footer_note',
    'notify_email'
  ];

  keys.forEach(k => {
    if (p[k] !== undefined) saveSettingValue(k, p[k]);
  });

  return { success: true };
}

function getSettingValue(key) {
  const sheet = getSheet_(SHEET_SETTINGS);
  if (!sheet) return '';

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      return String(data[i][1] || '');
    }
  }

  return '';
}

function saveSettingValue(key, value) {
  const sheet = getOrCreateSheet_(SHEET_SETTINGS, SETTINGS_HEADERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      sheet.getRange(i + 1, 2).setValue(String(value));
      return;
    }
  }

  sheet.appendRow([String(key), String(value)]);
}

// ============================================================
//  ADMIN AUTH
// ============================================================
function verifyAdmin(p) {
  const stored = getSettingValue('admin_password');
  const given = hashPassword(String(p.password || ''));
  const ok = stored && stored === given;

  notifyAdminLoginAttempt(ok, p);

  if (!ok) return { success: false, message: 'Incorrect password.' };
  return { success: true, message: 'Verified.' };
}

function changeAdminPassword(p) {
  const stored = getSettingValue('admin_password');
  const current = hashPassword(String(p.currentPassword || ''));

  if (!stored || stored !== current) {
    return { success: false, message: 'Current password is incorrect.' };
  }

  if (!p.newPassword || String(p.newPassword).length < 6) {
    return { success: false, message: 'New password must be at least 6 characters.' };
  }

  saveSettingValue('admin_password', hashPassword(String(p.newPassword)));
  notifyAdminPasswordChanged();

  return { success: true, message: 'Password changed.' };
}

function notifyAdminLoginAttempt(success, p) {
  try {
    const email = getSettingValue('notify_email');
    if (!email) return;

    const subject = success
      ? '✅ IELTS Admin Panel — Successful Login'
      : '⚠️ IELTS Admin Panel — Failed Login Attempt';

    const body =
      (success
        ? 'Someone successfully logged into your IELTS admin panel.'
        : 'Someone attempted to log in with an incorrect password.') +
      '\n\nTime: ' + new Date().toString() +
      '\nUser-Agent / Source: ' + (p.ua || 'Unknown') +
      '\n\nIf this was not you, change your admin password immediately from the admin panel -> Settings -> Change Password.';

    MailApp.sendEmail(email, subject, body);
  } catch (e) {}
}

function notifyAdminPasswordChanged() {
  try {
    const email = getSettingValue('notify_email');
    if (!email) return;

    MailApp.sendEmail(
      email,
      '🔑 IELTS Admin Panel — Password Changed',
      'Your IELTS admin panel password was just changed.\n\nTime: ' + new Date().toString() +
      '\n\nIf you did not make this change, your system may be compromised — review access immediately.'
    );
  } catch (e) {}
}

// ============================================================
//  CONFIG PROXY
// ============================================================
function getConfig(p) {
  if (!p.configURL) return { success: false, message: 'No configURL.' };

  try {
    const res = UrlFetchApp.fetch(p.configURL, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    const code = res.getResponseCode();
    if (code !== 200) return { success: false, message: 'HTTP ' + code };

    const json = JSON.parse(res.getContentText());
    return { success: true, config: json };
  } catch (e) {
    return { success: false, message: 'Fetch error: ' + e.toString() };
  }
}

// ============================================================
//  HELPERS
// ============================================================
function getCandidates() {
  const sheet = getSheet_(SHEET_CANDIDATES);
  if (!sheet) return { success: true, rows: [] };

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).map(r => [
    String(r[CIDX.ID] || ''),
    String(r[CIDX.FULL_NAME] || ''),
    String(r[CIDX.DOB] || ''),
    String(r[CIDX.GENDER] || ''),
    String(r[CIDX.NATIONALITY] || ''),
    '',
    String(r[CIDX.EMAIL] || ''),
    String(r[CIDX.CREATED_AT] || ''),
    String(r[CIDX.LAST_LOGIN] || '')
  ]);

  return { success: true, rows };
}

function findCandidate(candidateID) {
  if (!candidateID) return null;

  const sheet = getSheet_(SHEET_CANDIDATES);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const id = normalize_(candidateID);

  for (let i = 1; i < data.length; i++) {
    if (normalize_(data[i][CIDX.ID]) === id) return data[i];
  }

  return null;
}

function updateLastLogin(candidateID) {
  const sheet = getSheet_(SHEET_CANDIDATES);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const id = normalize_(candidateID);

  for (let i = 1; i < data.length; i++) {
    if (normalize_(data[i][CIDX.ID]) === id) {
      sheet.getRange(i + 1, CIDX.LAST_LOGIN + 1).setValue(new Date().toISOString());
      return;
    }
  }
}

function hashPassword(str) {
  str = String(str || '');
  let hash = 5381;
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) + hash) + str.charCodeAt(i);
    hash = hash & hash;
  }
  return 'h_' + Math.abs(hash).toString(16) + '_' + str.length;
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#002f5f')
      .setFontColor('#ffffff');
  }
  return sheet;
}

function normalize_(v) {
  return String(v || '').trim().toLowerCase();
}

function stringifyRow_(row) {
  return row.map(v => v === null || v === undefined ? '' : String(v));
}

function toNumber_(v, fallback) {
  const n = Number(v);
  return isNaN(n) ? fallback : n;
}

function toFloat_(v, fallback) {
  const n = parseFloat(v);
  return isNaN(n) ? fallback : n;
}

function roundHalf_(n) {
  return Math.round(n * 2) / 2;
}

// ============================================================
//  OTP — EMAIL VERIFICATION & PASSWORD RESET
// ============================================================
function generateOTP() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

function getOTPSheet() {
  return getOrCreateSheet_(SHEET_OTP, OTP_HEADERS);
}

function saveOTP(email, otp, purpose, candidateID) {
  const sheet = getOTPSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (
      normalize_(data[i][0]) === normalize_(email) &&
      String(data[i][2]) === String(purpose)
    ) {
      sheet.deleteRow(i + 1);
    }
  }

  const expires = new Date(Date.now() + 10 * 60 * 1000);
  sheet.appendRow([
    String(email).trim().toLowerCase(),
    String(otp),
    String(purpose),
    String(candidateID || ''),
    expires
  ]);
}

function verifyOTP(email, otp, purpose) {
  const sheet = getOTPSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (
      normalize_(data[i][0]) === normalize_(email) &&
      String(data[i][1]) === String(otp) &&
      String(data[i][2]) === String(purpose)
    ) {
      const expires = new Date(data[i][4]);

      if (now > expires) {
        sheet.deleteRow(i + 1);
        return { valid: false, message: 'Code has expired. Please request a new one.' };
      }

      const candidateID = String(data[i][3] || '');
      sheet.deleteRow(i + 1);

      return { valid: true, candidateID: candidateID };
    }
  }

  return { valid: false, message: 'Incorrect code. Please check and try again.' };
}

function sendOTPEmail(toEmail, otp, purpose) {
  const siteName = 'IELTS Academic Simulation';
  let subject, body;

  if (purpose === 'register') {
    subject = '[' + siteName + '] Your Registration Code: ' + otp;
    body =
      'Hello,\n\n' +
      'Your email verification code is:\n\n' +
      '    ' + otp + '\n\n' +
      'This code expires in 10 minutes.\n\n' +
      'If you did not register on ' + siteName + ', you can ignore this email.\n\n' +
      '— ' + siteName;
  } else if (purpose === 'emailchange') {
    subject = '[' + siteName + '] Confirm Your New Email: ' + otp;
    body =
      'Hello,\n\n' +
      'You requested to update your email address. Your confirmation code is:\n\n' +
      '    ' + otp + '\n\n' +
      'This code expires in 10 minutes.\n\n' +
      'If you did not request this change, please ignore this email.\n\n' +
      '— ' + siteName;
  } else {
    subject = '[' + siteName + '] Password Reset Code: ' + otp;
    body =
      'Hello,\n\n' +
      'You requested to reset your password. Your code is:\n\n' +
      '    ' + otp + '\n\n' +
      'This code expires in 10 minutes.\n\n' +
      'If you did not request this, please ignore this email.\n\n' +
      '— ' + siteName;
  }

  MailApp.sendEmail({
    to: String(toEmail).trim().toLowerCase(),
    subject: subject,
    body: body
  });
}

function handleSendOTP(p) {
  const email = String(p.email || '').trim().toLowerCase();
  const purpose = p.purpose || 'register';
  const candidateID = p.candidateID || '';

  if (!email || !email.includes('@')) {
    return { success: false, message: 'Invalid email address.' };
  }

  if (purpose === 'reset') {
    const sheet = getSheet_(SHEET_CANDIDATES);
    if (!sheet) return { success: false, message: 'No account found with that email address.' };

    const data = sheet.getDataRange().getValues();
    const found = data.slice(1).some(row =>
      normalize_(row[CIDX.EMAIL]) === normalize_(email)
    );

    if (!found) {
      return { success: false, message: 'No account found with that email address.' };
    }
  }

  try {
    const otp = generateOTP();
    saveOTP(email, otp, purpose, candidateID);
    sendOTPEmail(email, otp, purpose);
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Could not send email: ' + e.message };
  }
}

function handleVerifyOTPAndRegister(p) {
  const otp         = String(p.otp || '').trim();
  const email       = String(p.email || '').trim().toLowerCase();
  const candidateID = String(p.candidateID || '');
  const fullName    = String(p.fullName || '');
  const dob         = String(p.dob || '');
  const gender      = String(p.gender || '');
  const nationality = String(p.nationality || '');
  const password    = String(p.password || '');

  const otpResult = verifyOTP(email, otp, 'register');
  if (!otpResult.valid) {
    return { success: false, message: otpResult.message };
  }

  if (findCandidate(candidateID)) {
    return { success: false, message: 'This ID is already registered.' };
  }

  try {
    const sheet = getOrCreateSheet_(SHEET_CANDIDATES, CANDIDATE_HEADERS);
    const now = new Date().toISOString();

    sheet.appendRow([
      candidateID,
      fullName,
      dob,
      gender,
      nationality,
      email,
      hashPassword(password),
      now,
      now,
      ''
    ]);

    return { success: true, candidateID: candidateID, fullName: fullName };
  } catch (e) {
    return { success: false, message: 'Registration failed: ' + e.message };
  }
}

function handleVerifyOTPAndResetPassword(p) {
  const otp = String(p.otp || '').trim();
  const email = String(p.email || '').trim().toLowerCase();
  const newPass = String(p.newPassword || '');

  if (newPass.length < 6) {
    return { success: false, message: 'New password must be at least 6 characters.' };
  }

  const otpResult = verifyOTP(email, otp, 'reset');
  if (!otpResult.valid) {
    return { success: false, message: otpResult.message };
  }

  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    if (!sheet) return { success: false, message: 'Candidates sheet not found.' };

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (normalize_(data[i][CIDX.EMAIL]) === normalize_(email)) {
        sheet.getRange(i + 1, CIDX.PASSWORD + 1).setValue(hashPassword(newPass));
        return {
          success: true,
          fullName: String(data[i][CIDX.FULL_NAME] || ''),
          candidateID: String(data[i][CIDX.ID] || '')
        };
      }
    }

    return { success: false, message: 'Email not found.' };
  } catch (e) {
    return { success: false, message: 'Password reset failed: ' + e.message };
  }
}

function handleChangePassword(p) {
  const candidateID = String(p.candidateID || '');
  const currentPass = String(p.currentPassword || '');
  const newPass = String(p.newPassword || '');

  if (!candidateID || !currentPass || !newPass) {
    return { success: false, message: 'Missing required fields.' };
  }

  if (newPass.length < 6) {
    return { success: false, message: 'New password must be at least 6 characters.' };
  }

  const row = findCandidate(candidateID);
  if (!row) return { success: false, message: 'Candidate not found.' };

  if (String(row[CIDX.PASSWORD] || '') !== hashPassword(currentPass)) {
    return { success: false, message: 'Current password is incorrect.' };
  }

  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const data = sheet.getDataRange().getValues();
    const id = normalize_(candidateID);

    for (let i = 1; i < data.length; i++) {
      if (normalize_(data[i][CIDX.ID]) === id) {
        sheet.getRange(i + 1, CIDX.PASSWORD + 1).setValue(hashPassword(newPass));
        return { success: true };
      }
    }

    return { success: false, message: 'Candidate not found.' };
  } catch (e) {
    return { success: false, message: 'Password change failed: ' + e.message };
  }
}

function handleVerifyOTPAndUpdateEmail(p) {
  const otp = String(p.otp || '').trim();
  const newEmail = String(p.email || '').trim().toLowerCase();
  const candidateID = String(p.candidateID || '');

  if (!otp || !newEmail || !candidateID) {
    return { success: false, message: 'Missing required fields.' };
  }

  const otpResult = verifyOTP(newEmail, otp, 'emailchange');
  if (!otpResult.valid) {
    return { success: false, message: otpResult.message };
  }

  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    if (!sheet) return { success: false, message: 'Candidates sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const id = normalize_(candidateID);

    for (let i = 1; i < data.length; i++) {
      if (normalize_(data[i][CIDX.ID]) === id) {
        sheet.getRange(i + 1, CIDX.EMAIL + 1).setValue(newEmail);
        return { success: true };
      }
    }

    return { success: false, message: 'Candidate not found.' };
  } catch (e) {
    return { success: false, message: 'Email update failed: ' + e.message };
  }
}
