/**
 * Chennai Residency Lodge – guest register API (Google Apps Script)
 * Paste this into: Google Sheet > Extensions > Apps Script.
 * Only the Firebase API key below needs editing.
 */
const FIREBASE_API_KEY = 'PASTE_FIREBASE_API_KEY';   // same key as in index.html
const SHEET_NAME = 'Guest Register';
const TZ = 'Asia/Kolkata';

const HEADERS = [
  'Ref No', 'Date & Time', 'Type', 'Name', 'Gender', 'Age', 'Mobile (OTP verified)', 'Email',
  'Address', 'Coming From', 'Purpose', 'Nationality', 'ID Type', 'ID Number',
  'Passport No', 'Visa No', 'Visa Expiry', 'Persons', 'Room / Visiting', 'Expected Check-out',
  'Vehicle No', 'Actual Check-out (reception)'
];

function doGet() {
  return out_({ ok: true, msg: 'Guest register API is running' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);

    // 1. Server-side OTP check: the Firebase token must be valid and carry a phone number
    const phone = verifyToken_(d.idToken);
    if (!phone) return out_({ ok: false, error: 'Mobile verification failed. Please verify your number again.' });

    // 2. Basic validation
    if (!c_(d.name) || !c_(d.idNo) || !c_(d.address) || !c_(d.roomOrVisiting))
      return out_({ ok: false, error: 'Some required details are missing.' });

    // 3. Save
    const sh = getSheet_();
    const now = new Date();
    const serial = sh.getLastRow();   // header row counts as 1, so first guest = 1
    const ref = 'CRL-' + Utilities.formatDate(now, TZ, 'yyMMdd') + '-' + ('0000' + serial).slice(-4);

    const row = [
      ref, now, c_(d.type), c_(d.name), c_(d.gender), c_(d.age), phone, c_(d.email),
      c_(d.address), c_(d.comingFrom), c_(d.purpose), c_(d.nationality), c_(d.idType), c_(d.idNo),
      c_(d.passport), c_(d.visa), c_(d.visaExp), c_(d.persons), c_(d.roomOrVisiting), c_(d.checkout),
      c_(d.vehicle), ''
    ];
    const r = sh.getLastRow() + 1;
    const range = sh.getRange(r, 1, 1, row.length);
    range.setNumberFormat('@');                                   // keep phone / ID numbers as text
    sh.getRange(r, 2).setNumberFormat('dd-mmm-yyyy hh:mm');
    range.setValues([row]);

    return out_({ ok: true, ref: ref, time: Utilities.formatDate(now, TZ, 'dd MMM yyyy, hh:mm a') });
  } catch (err) {
    return out_({ ok: false, error: 'Server error. Please tell the reception desk.' });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function verifyToken_(idToken) {
  if (!idToken) return null;
  const res = UrlFetchApp.fetch(
    'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) return null;
  const u = (JSON.parse(res.getContentText()).users || [])[0];
  return u && u.phoneNumber ? u.phoneNumber : null;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#0f5c73').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Trim, cap length, and stop spreadsheet formula injection (= + - @ at the start)
function c_(v) {
  let s = (v === null || v === undefined) ? '' : String(v).trim().slice(0, 300);
  if (/^[=+\-@]/.test(s) && !/^\+\d[\d\s]*$/.test(s)) s = "'" + s;
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
