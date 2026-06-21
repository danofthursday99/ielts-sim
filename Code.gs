// ============================================================
//  IELTS Simulation — Google Apps Script Backend v3.0
//  
//  SETUP:
//  1. Paste this in Apps Script editor
//  2. Run setupSheets() once
//  3. Deploy → New Deployment → Web App
//     Execute as: Me | Who has access: Anyone
// ============================================================

const SHEET_CANDIDATES = 'Candidates';
const SHEET_SCORES     = 'Scores';
const SHEET_SETS       = 'TestSets';
const SHEET_SETTINGS   = 'Settings';

// ============================================================
//  SETUP
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Candidates sheet
  let cSheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (!cSheet) {
    cSheet = ss.insertSheet(SHEET_CANDIDATES);
    cSheet.appendRow(['CandidateID','FullName','DateOfBirth','Gender','Nationality','Password','CreatedAt','LastLogin','PhotoBase64']);
    cSheet.setFrozenRows(1);
    cSheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#002f5f').setFontColor('#ffffff');
  }

  // Scores sheet
  let sSheet = ss.getSheetByName(SHEET_SCORES);
  if (!sSheet) {
    sSheet = ss.insertSheet(SHEET_SCORES);
    sSheet.appendRow([
      'SubmissionID','Timestamp','CandidateID','FullName',
      'TestSetID','TestSetName',
      'ListeningScore','ListeningBand',
      'ReadingScore','ReadingBand',
      'WritingTask1Band','WritingTask2Band','WritingBand',
      'SpeakingBand','OverallBand',
      'WritingTask1WC','WritingTask2WC',
      'WritingTask1Text','WritingTask2Text',
      'ExaminerNotes','GradedBy','GradedAt',
      'ListeningAnswers','ReadingAnswers'
    ]);
    sSheet.setFrozenRows(1);
    sSheet.getRange(1,1,1,24).setFontWeight('bold').setBackground('#002f5f').setFontColor('#ffffff');
  }

  // Test Sets sheet
  let tSheet = ss.getSheetByName(SHEET_SETS);
  if (!tSheet) {
    tSheet = ss.insertSheet(SHEET_SETS);
    tSheet.appendRow(['SetID','SetName','ConfigURL','Status','CreatedAt','Description']);
    tSheet.setFrozenRows(1);
    tSheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#002f5f').setFontColor('#ffffff');
  }

  // Settings sheet
  let stSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!stSheet) {
    stSheet = ss.insertSheet(SHEET_SETTINGS);
    stSheet.appendRow(['Key','Value']);
    stSheet.appendRow(['institution_name','IELTS Simulation Center']);
    stSheet.appendRow(['institution_address','']);
    stSheet.appendRow(['active_set_id','']);
    stSheet.appendRow(['trf_footer_note','This is a simulation test result for preparation purposes only.']);
    stSheet.appendRow(['admin_password', hashPassword('admin1234')]); // CHANGE THIS after first login!
    stSheet.appendRow(['notify_email', Session.getActiveUser().getEmail() || '']); // defaults to script owner
    stSheet.setFrozenRows(1);
    stSheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#002f5f').setFontColor('#ffffff');
  }

  return 'All sheets ready!';
}

// ============================================================
//  SECURITY
//  This key adds a basic barrier against random/bot traffic.
//  It is visible in client-side code (unavoidable for a static
//  site with no real backend), so it does NOT stop a determined
//  attacker — but it blocks casual scrapers and accidental hits.
//  Change this to your own random string.
// ============================================================
const API_SECRET = 'ielts-sim-x7k2m9-change-me';

function checkApiKey(p) {
  return p.key === API_SECRET;
}

// ============================================================
//  ROUTER
// ============================================================
function doGet(e) {
  const p        = e.parameter || {};
  const callback = p.callback  || '';
  const action   = p.action    || '';
  let result;

  try {
    // Public actions — no key required (candidate-facing, must work for everyone)
    const publicActions = ['ping','checkCandidate','registerCandidate','loginCandidate','saveScore','getTestSets','getActiveSet','getConfig','getCandidateScores','verifyAdmin','updateProfile','updateProfilePhoto'];

    if (!publicActions.includes(action) && !checkApiKey(p)) {
      result = { success:false, message:'Unauthorized — invalid or missing API key.' };
    } else {
      switch(action) {
        case 'ping':              result = { success:true, message:'pong', time:new Date().toISOString() }; break;
        case 'setup':             result = { success:true, message:setupSheets() }; break;
        case 'checkCandidate':    result = checkCandidate(p);    break;
        case 'registerCandidate': result = registerCandidate(p); break;
        case 'loginCandidate':    result = loginCandidate(p);    break;
        case 'updateProfile':     result = updateProfile(p);     break;
        case 'updateProfilePhoto':result = updateProfilePhoto(p);break;
        case 'saveScore':         result = saveScore(p);         break;
        case 'getCandidates':     result = getCandidates();      break;
        case 'getScores':         result = getScores(p);         break;
        case 'getCandidateScores':result = getCandidateScores(p);break;
        case 'gradeWriting':      result = gradeWriting(p);      break;
        case 'getTestSets':       result = getTestSets();        break;
        case 'saveTestSet':       result = saveTestSet(p);       break;
        case 'deleteTestSet':     result = deleteTestSet(p);     break;
        case 'setActiveSet':      result = setActiveSet(p);      break;
        case 'getActiveSet':      result = getActiveSet();       break;
        case 'getSettings':       result = getSettings();        break;
        case 'saveSettings':      result = saveSettings(p);      break;
        case 'getConfig':         result = getConfig(p);         break;
        case 'verifyAdmin':       result = verifyAdmin(p);       break;
        case 'changeAdminPassword': result = changeAdminPassword(p); break;
        default: result = { success:false, message:'Unknown action: '+action };
      }
    }
  } catch(err) {
    result = { success:false, message:'Error: '+err.toString() };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback+'('+json+');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) { return doGet(e); }

// ============================================================
//  CANDIDATE ACTIONS
// ============================================================
function checkCandidate(p) {
  const row = findCandidate(p.candidateID);
  if (row) return { success:true, exists:true, name:String(row[1]), nationality:String(row[4]) };
  return { success:true, exists:false };
}

function registerCandidate(p) {
  if (!p.candidateID||!p.fullName||!p.password)
    return { success:false, message:'Missing required fields.' };
  if (findCandidate(p.candidateID))
    return { success:false, message:'Candidate ID already registered.' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_CANDIDATES);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SHEET_CANDIDATES); }
  const now   = new Date().toISOString();
  sheet.appendRow([
    String(p.candidateID), String(p.fullName),
    String(p.dob||''), String(p.gender||''), String(p.nationality||''),
    hashPassword(String(p.password)), now, now, ''
  ]);
  return { success:true, candidateID:p.candidateID, fullName:p.fullName };
}

function loginCandidate(p) {
  if (!p.candidateID||!p.password)
    return { success:false, message:'Missing credentials.' };
  const row = findCandidate(p.candidateID);
  if (!row) return { success:false, message:'Candidate ID not found.' };
  if (String(row[5]) !== hashPassword(String(p.password)))
    return { success:false, message:'Incorrect password.' };
  updateLastLogin(p.candidateID);
  return {
    success:true,
    candidateID:  String(row[0]),
    fullName:     String(row[1]),
    dob:          String(row[2]),
    gender:       String(row[3]),
    nationality:  String(row[4]),
    photoBase64:  String(row[8]||'')
  };
}

// ============================================================
//  PROFILE UPDATES
// ============================================================
function updateProfile(p) {
  if (!p.candidateID) return { success:false, message:'No candidateID.' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (!sheet) return { success:false, message:'Candidates sheet not found.' };
  const data  = sheet.getDataRange().getValues();
  const id    = String(p.candidateID).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === id) {
      const row = i + 1;
      if (p.fullName     !== undefined) sheet.getRange(row, 2).setValue(String(p.fullName));
      if (p.dob          !== undefined) sheet.getRange(row, 3).setValue(String(p.dob));
      if (p.gender       !== undefined) sheet.getRange(row, 4).setValue(String(p.gender));
      if (p.nationality  !== undefined) sheet.getRange(row, 5).setValue(String(p.nationality));
      return { success:true, message:'Profile updated.' };
    }
  }
  return { success:false, message:'Candidate not found.' };
}

function updateProfilePhoto(p) {
  if (!p.candidateID) return { success:false, message:'No candidateID.' };
  // Rough size guard — Sheets cells max out around 50,000 chars
  if (p.photoBase64 && p.photoBase64.length > 45000) {
    return { success:false, message:'Image too large even after compression. Try a smaller photo.' };
  }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (!sheet) return { success:false, message:'Candidates sheet not found.' };
  const data  = sheet.getDataRange().getValues();
  const id    = String(p.candidateID).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === id) {
      sheet.getRange(i+1, 9).setValue(String(p.photoBase64 || ''));
      return { success:true, message:'Photo updated.' };
    }
  }
  return { success:false, message:'Candidate not found.' };
}

// ============================================================
//  SCORE ACTIONS
// ============================================================
function saveScore(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_SCORES);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SHEET_SCORES); }
  const subID = 'SUB-' + Date.now();
  sheet.appendRow([
    subID,
    new Date().toISOString(),
    String(p.candidateID   || ''),
    String(p.fullName      || ''),
    String(p.testSetID     || ''),
    String(p.testSetName   || ''),
    Number(p.listeningScore|| 0),
    Number(p.listeningBand || 0),
    Number(p.readingScore  || 0),
    Number(p.readingBand   || 0),
    '', '', '',  // W1 band, W2 band, Writing band (filled by examiner)
    '',          // Speaking band
    '',          // Overall (calculated after grading)
    Number(p.writingTask1WC|| 0),
    Number(p.writingTask2WC|| 0),
    String(p.writingTask1  || '').substring(0,2000),
    String(p.writingTask2  || '').substring(0,2000),
    '', '', '',  // Examiner notes, graded by, graded at
    String(p.listeningAnswers || ''),  // Raw L answers (JSON) — admin checks manually
    String(p.readingAnswers   || '')   // Raw R answers (JSON) — admin checks manually
  ]);
  return { success:true, submissionID:subID };
}

function getScores(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SCORES);
  if (!sheet) return { success:true, rows:[] };
  const data  = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  let rows = data.slice(1).map(r => r.map(v => v === null || v === undefined ? '' : String(v)));
  // Filter by testSetID if provided
  if (p && p.testSetID) {
    const setIdx = headers.indexOf('TestSetID');
    if (setIdx >= 0) rows = rows.filter(r => r[setIdx] === p.testSetID);
  }
  return { success:true, rows, headers };
}

function getCandidateScores(p) {
  if (!p.candidateID) return { success:false, message:'No candidateID.' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SCORES);
  if (!sheet) return { success:true, rows:[] };
  const data  = sheet.getDataRange().getValues();
  const id    = String(p.candidateID).toLowerCase();
  const rows  = data.slice(1)
    .filter(r => String(r[2]).toLowerCase() === id)
    .map(r => r.map(v => v === null || v === undefined ? '' : String(v)));
  return { success:true, rows };
}

function gradeWriting(p) {
  if (!p.submissionID) return { success:false, message:'No submissionID.' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SCORES);
  if (!sheet) return { success:false, message:'Scores sheet not found.' };
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.submissionID)) {
      const row    = i + 1;
      const w1Band = parseFloat(p.writingTask1Band) || 0;
      const w2Band = parseFloat(p.writingTask2Band) || 0;
      const wBand  = ((w1Band + w2Band) / 2).toFixed(1);
      const sBand  = parseFloat(p.speakingBand)     || 0;
      // Writing bands
      sheet.getRange(row, 11).setValue(w1Band);
      sheet.getRange(row, 12).setValue(w2Band);
      sheet.getRange(row, 13).setValue(parseFloat(wBand));
      sheet.getRange(row, 14).setValue(sBand || '');
      // Calculate overall band (average of available components)
      const lBand  = parseFloat(data[i][7]) || 0;
      const rBand  = parseFloat(data[i][9]) || 0;
      const components = [lBand, rBand, parseFloat(wBand)];
      if (sBand) components.push(sBand);
      const avg     = components.reduce((a,b)=>a+b,0) / components.length;
      const overall = Math.round(avg * 2) / 2; // Round to nearest 0.5
      sheet.getRange(row, 15).setValue(overall);
      // Examiner info
      sheet.getRange(row, 20).setValue(String(p.examinerNotes || ''));
      sheet.getRange(row, 21).setValue(String(p.gradedBy      || 'Admin'));
      sheet.getRange(row, 22).setValue(new Date().toISOString());
      return { success:true, overallBand:overall, writingBand:parseFloat(wBand) };
    }
  }
  return { success:false, message:'Submission not found.' };
}

// ============================================================
//  TEST SETS
// ============================================================
function getTestSets() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETS);
  if (!sheet) return { success:true, sets:[] };
  const data  = sheet.getDataRange().getValues();
  const sets  = data.slice(1).map(r => ({
    setID:       String(r[0]),
    setName:     String(r[1]),
    configURL:   String(r[2]),
    status:      String(r[3]),
    createdAt:   String(r[4]),
    description: String(r[5])
  }));
  return { success:true, sets };
}

function saveTestSet(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_SETS);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SHEET_SETS); }
  // Check if set already exists — update it
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.setID)) {
      sheet.getRange(i+1,1,1,6).setValues([[
        String(p.setID), String(p.setName), String(p.configURL||''),
        String(p.status||'Active'), String(data[i][4]), String(p.description||'')
      ]]);
      return { success:true, message:'Updated.' };
    }
  }
  // New set
  const setID = p.setID || ('SET-'+Date.now());
  sheet.appendRow([
    setID, String(p.setName), String(p.configURL||''),
    String(p.status||'Active'), new Date().toISOString(), String(p.description||'')
  ]);
  return { success:true, setID, message:'Created.' };
}

function deleteTestSet(p) {
  if (!p.setID) return { success:false, message:'No setID.' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETS);
  if (!sheet) return { success:false, message:'TestSets sheet not found.' };
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.setID)) {
      sheet.deleteRow(i + 1);
      // If this was the active set, clear the setting
      if (getSettingValue('active_set_id') === String(p.setID)) {
        saveSettingValue('active_set_id', '');
      }
      return { success:true, message:'Deleted.' };
    }
  }
  return { success:false, message:'Test set not found.' };
}

function setActiveSet(p) {
  if (!p.setID) return { success:false, message:'No setID.' };
  saveSettingValue('active_set_id', p.setID);
  return { success:true };
}

function getActiveSet() {
  const id    = getSettingValue('active_set_id');
  if (!id) return { success:true, set:null };
  const all   = getTestSets();
  const found = all.sets.find(s => s.setID === id);
  return { success:true, set: found || null };
}

// ============================================================
//  SETTINGS
// ============================================================
function getSettings() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return { success:true, settings:{} };
  const data  = sheet.getDataRange().getValues();
  const settings = {};
  data.slice(1).forEach(r => {
    const key = String(r[0]);
    if (key === 'admin_password') return; // never expose the password hash to the client
    settings[key] = String(r[1]);
  });
  return { success:true, settings };
}

function saveSettings(p) {
  const keys = ['institution_name','institution_address','trf_footer_note','notify_email'];
  keys.forEach(k => { if (p[k] !== undefined) saveSettingValue(k, p[k]); });
  return { success:true };
}

function getSettingValue(key) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return '';
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return String(data[i][1]);
  }
  return '';
}

function saveSettingValue(key, value) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(SHEET_SETTINGS); }
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sheet.getRange(i+1, 2).setValue(String(value));
      return;
    }
  }
  sheet.appendRow([key, String(value)]);
}

// ============================================================
//  ADMIN AUTH — password verified server-side, never shipped
//  to the client. Login attempts trigger an email alert.
// ============================================================
function verifyAdmin(p) {
  const stored = getSettingValue('admin_password');
  const given  = hashPassword(String(p.password || ''));
  const ok     = stored && stored === given;

  notifyAdminLoginAttempt(ok, p);

  if (!ok) return { success:false, message:'Incorrect password.' };
  return { success:true, message:'Verified.' };
}

function changeAdminPassword(p) {
  // Require the CURRENT password before allowing a change
  const stored  = getSettingValue('admin_password');
  const current = hashPassword(String(p.currentPassword || ''));
  if (!stored || stored !== current) {
    return { success:false, message:'Current password is incorrect.' };
  }
  if (!p.newPassword || String(p.newPassword).length < 6) {
    return { success:false, message:'New password must be at least 6 characters.' };
  }
  saveSettingValue('admin_password', hashPassword(String(p.newPassword)));
  notifyAdminPasswordChanged();
  return { success:true, message:'Password changed.' };
}

function notifyAdminLoginAttempt(success, p) {
  try {
    const email = getSettingValue('notify_email');
    if (!email) return; // no email configured — skip silently
    const subject = success
      ? '✅ IELTS Admin Panel — Successful Login'
      : '⚠️ IELTS Admin Panel — Failed Login Attempt';
    const body =
      (success ? 'Someone successfully logged into your IELTS admin panel.' : 'Someone attempted to log in with an incorrect password.') +
      '\n\nTime: ' + new Date().toString() +
      '\nUser-Agent / Source: ' + (p.ua || 'Unknown') +
      '\n\nIf this was not you, change your admin password immediately from the admin panel → Settings → Change Password.';
    MailApp.sendEmail(email, subject, body);
  } catch(e) {
    // Email failures should never break login — fail silently
  }
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
  } catch(e) {}
}

// ============================================================
//  CONFIG PROXY (fetch JSON from GitHub/Drive)
// ============================================================
function getConfig(p) {
  if (!p.configURL) return { success:false, message:'No configURL.' };
  try {
    const res  = UrlFetchApp.fetch(p.configURL, { muteHttpExceptions:true, followRedirects:true });
    const code = res.getResponseCode();
    if (code !== 200) return { success:false, message:'HTTP '+code };
    const json = JSON.parse(res.getContentText());
    return { success:true, config:json };
  } catch(e) {
    return { success:false, message:'Fetch error: '+e.toString() };
  }
}

// ============================================================
//  HELPERS
// ============================================================
function getCandidates() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (!sheet) return { success:true, rows:[] };
  const data  = sheet.getDataRange().getValues();
  const rows  = data.slice(1).map(r => [
    String(r[0]), String(r[1]), String(r[2]),
    String(r[3]), String(r[4]), '',
    String(r[6]), String(r[7])
  ]);
  return { success:true, rows };
}

function findCandidate(candidateID) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (!sheet) return null;
  const data  = sheet.getDataRange().getValues();
  const id    = String(candidateID).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === id) return data[i];
  }
  return null;
}

function updateLastLogin(candidateID) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CANDIDATES);
  if (!sheet) return;
  const data  = sheet.getDataRange().getValues();
  const id    = String(candidateID).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === id) {
      sheet.getRange(i+1, 8).setValue(new Date().toISOString());
      break;
    }
  }
}

function hashPassword(str) {
  let hash = 5381;
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) + hash) + str.charCodeAt(i);
    hash = hash & hash;
  }
  return 'h_' + Math.abs(hash).toString(16) + '_' + str.length;
}
