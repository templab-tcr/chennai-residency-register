// ============================================================
// Guest Register + Property Checks + Hotel Sync - Google Apps Script (full merged code, fixed)
// One script, one /exec URL, one spreadsheet for:
//  - index.html           (guest / visitor registration, Firebase phone OTP)
//  - admin.html           (visitor timer, ADMIN_PIN)
//  - property-check.html  (room inspections, ADMIN_PIN)
//  - hotel.html           (hotel software sync, HOTEL_TOKEN)
//
// FIRST-TIME SETUP (do these in order):
//  1. Edit the two values inside setSecrets() below, then run setSecrets() once.
//  2. Run setupHeaders() once (creates all tabs).
//  3. Run hotelDiagnose() once and read the Execution log: every line must say OK.
//  4. Deploy > Manage deployments > pencil icon > Version: "New version" > Deploy.
//     Execute as: Me.   Who has access: Anyone.   (Editing code alone does NOT update the live URL.)
//  5. In hotel.html > Guest Records, type the same HOTEL_TOKEN in "Secret token", click Save & Sync Now.
// ============================================================

const FIREBASE_API_KEY = "AIzaSyA-Cn6-j_iZi7oN_yaaKjHsCG4-_JxHyiE"; // fallback; Script property FIREBASE_API_KEY takes priority

const SPREADSHEET_ID = "1At8aEYhMUnFjVB8Oe7veAqKgWN-nzzPItXlC01rYyUY";
const SHEET_NAME = "Guest Register";
const REGISTRATION_SHEET_NAME = "Guest Registrations";
const LODGE_PREFIX = "CRL";
const PHOTO_FOLDER_NAME = "Guest ID Photos";
const PROPERTY_PHOTO_FOLDER_NAME = "Property Check Photos";
const PROPERTY_CHECK_SHEET_NAME = "Property Checks";
const PROPERTY_ITEMS_SHEET_NAME = "Property Check Items";
const MAX_SHEET_PHOTOS = 8;

const HOTEL_DATA_SHEET = "Hotel Data";
const HOTEL_CHUNK = 45000; // a cell holds max 50,000 characters

const HEADERS = [
  "Reference Number", "Registration Time", "Registration Type", "Full Name", "Gender", "Age",
  "Email", "Home Address", "Coming From", "Purpose of Visit", "Nationality", "ID Type",
  "ID Number", "Passport Number", "Visa Number", "Visa Expiry", "Total Persons",
  "Room Number / Meeting Guest", "Expected Check-out", "Vehicle Number", "Verified Mobile Number",
  "ID Front Photo", "ID Back Photo", "Status", "Stopped At"
];

const REGISTRATION_HEADERS = [
  "Submitted At", "Type", "Full Name", "Gender", "Age",
  "Phone", "Email", "Address", "Nationality", "ID Type",
  "ID Number", "Purpose", "Room / Visiting", "Persons",
  "Expected Checkout", "Vehicle"
];

const PROPERTY_CHECK_HEADERS = [
  "Reference Number", "Submitted At", "Room Number", "Room Type",
  "Inspection Type", "Inspection Date", "Inspection Time", "Staff Name",
  "Result", "Issues", "Notes", "Photo Count",
  "Photo 1", "Photo 2", "Photo 3", "Photo 4",
  "Photo 5", "Photo 6", "Photo 7", "Photo 8",
  "All Photo Links", "Client ID", "Raw Items JSON"
];
const PC = { PHOTO_FIRST: 13, LINKS: 21, CLIENT: 22, RAW: 23 };

const PROPERTY_ITEM_HEADERS = [
  "Reference Number", "Inspection Date", "Room Number", "Inspection Type",
  "Staff Name", "Item", "Expected", "Status", "Quantity", "Note"
];

/* ================= Entry points ================= */

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === "status") return jsonResponse(getStatus_(p.ref));

  // hotel.html: ?action=rev  /  ?action=get  (&token=...)
  if (p.action === "rev" || p.action === "get") return hotelGet_(p);

  return jsonResponse({ ok: true, message: "Guest Register service is running." });
}

function getFirebaseKey_() {
  const fromProps = PropertiesService.getScriptProperties().getProperty("FIREBASE_API_KEY");
  return fromProps || FIREBASE_API_KEY;
}

function lastTenDigits_(s) {
  return String(s || "").replace(/\D/g, "").slice(-10);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("No registration data received.");

    const d = JSON.parse(e.postData.contents);

    // hotel.html sends { token, data } with no "action".
    if (d.data && typeof d.data === "object" && !d.action) return hotelPush_(d);

    // Property checks use the Staff PIN only.
    if (d.action === "propertyCheck") return savePropertyCheck_(d);

    // Admin actions
    if (d.action === "adminList") return adminList_(d);
    if (d.action === "adminStop") return adminStop_(d);

    // Guest/visitor registration requires Firebase phone verification.
    const apiKey = getFirebaseKey_();
    if (!apiKey || apiKey === "PASTE_FIREBASE_API_KEY_HERE") throw new Error("Firebase API key is not configured.");
    if (!d.idToken) throw new Error("Phone verification is required.");

    const verifyUrl = "https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=" + encodeURIComponent(apiKey);
    const verification = UrlFetchApp.fetch(verifyUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ idToken: d.idToken }),
      muteHttpExceptions: true
    });

    const verifyStatus = verification.getResponseCode();
    let verifyResult = {};
    try { verifyResult = JSON.parse(verification.getContentText()); } catch (parseErr) { verifyResult = {}; }

    if (verifyStatus !== 200 || !verifyResult.users || verifyResult.users.length === 0) {
      const reason = verifyResult && verifyResult.error && verifyResult.error.message
        ? " (" + verifyResult.error.message + ")" : "";
      throw new Error("Phone verification failed. Please verify again." + reason);
    }

    const user = verifyResult.users[0];
    if (!user.phoneNumber) throw new Error("A verified phone number is required.");

    if (d.phone && lastTenDigits_(d.phone) !== lastTenDigits_(user.phoneNumber)) {
      throw new Error("Verified phone does not match the form.");
    }

    if (!["Guest", "Visitor"].includes(d.type)) throw new Error("Invalid registration type.");
    if (!d.name || String(d.name).trim().length < 2) throw new Error("A valid full name is required.");
    if (!d.address || String(d.address).trim().length < 5) throw new Error("A valid home address is required.");
    if (!d.idType || !d.idNo) throw new Error("ID proof details are required.");
    if (d.nationality !== "Indian" && (!d.passport || !d.visa || !d.visaExp)) {
      throw new Error("Passport and visa details are required.");
    }
    if (!d.purpose || !d.roomOrVisiting) throw new Error("Purpose and room or meeting details are required.");
    if (d.type === "Guest" && !d.checkout) throw new Error("Expected check-out date is required.");

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);

    try {
      const ss = getSpreadsheet_();
      let sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

      if (sheet.getLastRow() === 0 ||
          sheet.getRange(1, HEADERS.length).getValue() !== HEADERS[HEADERS.length - 1]) {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      }

      const now = new Date();
      const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyMMdd");
      const reference = getNextReference(sheet, datePart);

      let frontUrl = "";
      let backUrl = "";
      if (d.idFront || d.idBack) {
        const folder = getPhotoFolder();
        if (d.idFront) frontUrl = saveImage(d.idFront, reference + "_front.jpg", folder);
        if (d.idBack) backUrl = saveImage(d.idBack, reference + "_back.jpg", folder);
      }

      const values = [
        reference, now, d.type, d.name, d.gender, d.age, d.email, d.address, d.comingFrom,
        d.purpose, d.nationality, d.idType, d.idNo, d.passport, d.visa, d.visaExp, d.persons,
        d.roomOrVisiting, d.checkout, d.vehicle, user.phoneNumber, frontUrl, backUrl,
        d.type === "Visitor" ? "Active" : "Checked-in", ""
      ];

      const safeValues = values.map(function(value, index) {
        if (index === 1) return value;
        if (value === null || value === undefined) return "";
        const s = String(value);
        return /^[=+\-@]/.test(s) ? "'" + s : s;
      });

      sheet.appendRow(safeValues);

      appendGuestRegistration_(ss, {
        type: d.type, name: d.name, gender: d.gender, age: d.age, phone: user.phoneNumber,
        email: d.email, address: d.address, nationality: d.nationality, idType: d.idType,
        idNo: d.idNo, purpose: d.purpose, roomOrVisiting: d.roomOrVisiting,
        persons: d.persons, checkout: d.checkout, vehicle: d.vehicle
      }, now);

      return jsonResponse({
        ok: true,
        ref: reference,
        time: Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy hh:mm a")
      });
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    console.error("doPost error: " + (err && err.stack ? err.stack : err));
    return jsonResponse({ ok: false, error: err.message || "Registration could not be saved." });
  }
}

/* ================= Property checks ================= */

function getPropertyPhotoFolder_() {
  const it = DriveApp.getFoldersByName(PROPERTY_PHOTO_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PROPERTY_PHOTO_FOLDER_NAME);
}

function savePropertyCheck_(d) {
  adminAuth_(d);

  if (!d.room || !d.roomType || !d.kind || !d.date || !d.staff || String(d.staff).trim().length < 2) {
    throw new Error("Room, room type, inspection type, date and staff name are required.");
  }
  if (!Array.isArray(d.items)) throw new Error("Inspection checklist data is missing.");
  if (!Array.isArray(d.photos) || d.photos.length === 0) throw new Error("At least one room photo is required.");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const ss = getSpreadsheet_();
    const sheets = ensurePropertySheets_(ss);
    const now = new Date();
    const stamp = function() {
      return Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy hh:mm a");
    };

    if (d.clientId && sheets.main.getLastRow() >= 2) {
      const found = sheets.main
        .getRange(2, PC.CLIENT, sheets.main.getLastRow() - 1, 1)
        .createTextFinder(String(d.clientId))
        .matchEntireCell(true)
        .findNext();
      if (found) {
        return jsonResponse({
          ok: true,
          ref: sheets.main.getRange(found.getRow(), 1).getDisplayValue(),
          time: stamp(),
          duplicate: true
        });
      }
    }

    const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyMMdd");
    const reference = getNextPropertyCheckReference_(sheets.main, datePart);

    const folder = getPropertyPhotoFolder_();
    const photos = [];

    d.photos.forEach(function(photo, index) {
      if (!photo || !photo.data) return;
      const label = String(photo.label || ("Photo " + (index + 1)))
        .replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 40);
      const url = saveImage(
        photo.data,
        reference + "_Room" + String(d.room).replace(/[^a-zA-Z0-9]/g, "") + "_" + (index + 1) + "_" + label + ".jpg",
        folder
      );
      photos.push({ label: String(photo.label || ""), item: String(photo.item || ""), url: url, data: photo.data });
    });

    if (photos.length === 0) throw new Error("No valid room photos were received.");

    appendPropertyRow_(sheets, {
      ref: reference, submittedAt: now, clientId: String(d.clientId || ""),
      room: d.room, roomType: d.roomType, kind: d.kind, date: d.date, time: d.time,
      staff: d.staff, notes: d.notes, items: d.items, photos: photos
    });

    return jsonResponse({ ok: true, ref: reference, time: stamp() });
  } finally {
    lock.releaseLock();
  }
}

function ensurePropertySheets_(ss) {
  let main = ss.getSheetByName(PROPERTY_CHECK_SHEET_NAME);

  if (main && main.getLastRow() > 0 && String(main.getRange(1, 3).getValue()) !== "Room Number") {
    main.setName(uniqueSheetName_(ss, PROPERTY_CHECK_SHEET_NAME + " (old layout)"));
    main = null;
  }

  if (!main) main = ss.insertSheet(PROPERTY_CHECK_SHEET_NAME);
  if (main.getLastRow() === 0) formatPropertyMain_(main);

  let items = ss.getSheetByName(PROPERTY_ITEMS_SHEET_NAME);
  if (!items) items = ss.insertSheet(PROPERTY_ITEMS_SHEET_NAME);
  if (items.getLastRow() === 0) formatPropertyItems_(items);

  return { main: main, items: items };
}

function uniqueSheetName_(ss, base) {
  let name = base, n = 2;
  while (ss.getSheetByName(name)) { name = base + " " + n; n++; }
  return name;
}

function formatPropertyMain_(sheet) {
  const cols = PROPERTY_CHECK_HEADERS.length;
  sheet.getRange(1, 1, 1, cols).setValues([PROPERTY_CHECK_HEADERS])
    .setFontWeight("bold").setBackground("#642b27").setFontColor("#f2bb4e")
    .setWrap(true).setVerticalAlignment("middle").setHorizontalAlignment("center");
  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  const widths = [150, 150, 80, 85, 130, 105, 85, 120, 110, 300, 220, 70];
  for (let i = 0; i < MAX_SHEET_PHOTOS; i++) widths.push(130);
  widths.push(280, 140, 100);
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.hideColumns(PC.RAW);

  const result = sheet.getRange("I2:I");
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LEFT($I2,3)="All"').setBackground("#d9ead3").setRanges([result]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($I2<>"",LEFT($I2,3)<>"All")').setBackground("#f4cccc").setRanges([result]).build()
  ]);
}

function formatPropertyItems_(sheet) {
  const cols = PROPERTY_ITEM_HEADERS.length;
  sheet.getRange(1, 1, 1, cols).setValues([PROPERTY_ITEM_HEADERS])
    .setFontWeight("bold").setBackground("#642b27").setFontColor("#f2bb4e").setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  [150, 105, 80, 130, 120, 150, 170, 140, 80, 300].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  const status = sheet.getRange("H2:H");
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($H2<>"",$H2<>"OK",$H2<>"Not recorded")')
      .setBackground("#f4cccc").setRanges([status]).build()
  ]);
}

function safeText_(value) {
  if (value === null || value === undefined) return "";
  const s = String(value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function plainText_(value, fmt) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), fmt);
  return String(value === null || value === undefined ? "" : value);
}

function appendPropertyRow_(sheets, r) {
  const main = sheets.main;
  const items = Array.isArray(r.items) ? r.items : [];
  const photos = Array.isArray(r.photos) ? r.photos : [];

  const issues = items.filter(function(i) {
    return i && i.status && i.status !== "ok" && i.status !== "not recorded";
  });
  const issueText = issues.map(function(i) {
    let line = "• " + i.name + " – " + (i.label || i.status);
    if (i.status === "diff" && i.qty !== null && i.qty !== undefined) line += " (count " + i.qty + ")";
    if (i.note) line += ": " + i.note;
    return line;
  }).join("\n");
  const result = issues.length
    ? issues.length + (issues.length === 1 ? " issue found" : " issues found")
    : "All OK";

  const dateText = plainText_(r.date, "yyyy-MM-dd");
  const dm = dateText.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const dateValue = dm ? new Date(Number(dm[1]), Number(dm[2]) - 1, Number(dm[3])) : dateText;

  const links = photos.map(function(p) {
    return (p.label || "Photo") + (p.item ? " (" + p.item + ")" : "") + ": " + (p.url || "");
  }).join("\n");

  const rawJson = JSON.stringify({
    items: items,
    photos: photos.map(function(p) { return { label: p.label || "", item: p.item || "", url: p.url || "" }; })
  });

  const row = main.getLastRow() + 1;
  const values = [
    r.ref, r.submittedAt, safeText_(r.room), safeText_(r.roomType), safeText_(r.kind),
    dateValue, plainText_(r.time, "hh:mm a"), safeText_(r.staff), result, issueText,
    safeText_(r.notes), photos.length
  ];
  while (values.length < PC.PHOTO_FIRST - 1 + MAX_SHEET_PHOTOS) values.push("");
  values.push(links, safeText_(r.clientId), rawJson);

  main.getRange(row, 7).setNumberFormat("@");
  main.getRange(row, 1, 1, values.length).setValues([values]);
  main.getRange(row, 2).setNumberFormat("dd-mmm-yyyy hh:mm");
  main.getRange(row, 6).setNumberFormat("dd-mmm-yyyy");
  main.getRange(row, 1, 1, PROPERTY_CHECK_HEADERS.length).setVerticalAlignment("middle").setWrap(true);
  main.getRange(row, 3, 1, 10).setHorizontalAlignment("left");
  main.setRowHeight(row, 110);

  for (let i = 0; i < Math.min(photos.length, MAX_SHEET_PHOTOS); i++) {
    embedPhotoCell_(main, row, PC.PHOTO_FIRST + i, photos[i], r.ref + " Room " + r.room);
  }

  if (items.length) {
    const ir = sheets.items;
    const start = ir.getLastRow() + 1;
    const rows = items.map(function(i) {
      return [
        r.ref, dateValue, safeText_(r.room), safeText_(r.kind), safeText_(r.staff),
        safeText_(i.name), safeText_(i.expected), safeText_(i.label || i.status),
        (i.qty === null || i.qty === undefined) ? "" : i.qty, safeText_(i.note)
      ];
    });
    ir.getRange(start, 2, rows.length, 1).setNumberFormat("dd-mmm-yyyy");
    ir.getRange(start, 1, rows.length, PROPERTY_ITEM_HEADERS.length).setValues(rows);
  }
}

function embedPhotoCell_(sheet, row, col, photo, altTitle) {
  const cell = sheet.getRange(row, col);
  const url = String((photo && photo.url) || "");
  const label = String((photo && photo.label) || "Photo");

  try {
    const uri = photoDataUri_(photo);
    if (!uri) throw new Error("No image data available.");

    const image = SpreadsheetApp.newCellImage()
      .setSourceUrl(uri)
      .setAltTextTitle(altTitle + " – " + label)
      .setAltTextDescription(String((photo && (photo.item || photo.label)) || "Property inspection photo"))
      .build();
    cell.setValue(image);
  } catch (err) {
    if (url) cell.setFormula('=HYPERLINK("' + url.replace(/"/g, '""') + '","Open ' + label.replace(/"/g, '""') + '")');
    cell.setNote("Photo could not be shown inside the cell: " + String(err.message || err) + (url ? "\nDrive file: " + url : ""));
    console.warn("Property photo embedding failed at row " + row + ", column " + col + ": " + String(err));
  }
}

function photoDataUri_(photo) {
  const LIMIT = 2 * 1024 * 1024;
  if (photo && photo.data && /^data:image\//i.test(photo.data)) {
    if (photo.data.length * 0.75 > LIMIT) throw new Error("Image is larger than 2 MB.");
    return photo.data;
  }
  const match = String((photo && photo.url) || "").match(/[-\w]{25,}/);
  if (!match) return "";

  const file = DriveApp.getFileById(match[0]);
  let blob = file.getBlob();
  if (blob.getBytes().length > LIMIT) {
    const thumb = file.getThumbnail();
    if (thumb) blob = thumb;
  }
  const bytes = blob.getBytes();
  if (bytes.length > LIMIT) throw new Error("Image is larger than 2 MB.");
  return "data:" + (blob.getContentType() || "image/jpeg") + ";base64," + Utilities.base64Encode(bytes);
}

function migratePropertyChecksToNewLayout() {
  const ss = getSpreadsheet_();
  const sheets = ensurePropertySheets_(ss);
  const done = {};

  if (sheets.main.getLastRow() > 1) {
    sheets.main.getRange(2, 1, sheets.main.getLastRow() - 1, 1)
      .getDisplayValues().forEach(function(r) { done[r[0]] = true; });
  }

  let copied = 0;
  ss.getSheets().forEach(function(old) {
    if (old.getName().indexOf(PROPERTY_CHECK_SHEET_NAME + " (old layout") !== 0) return;
    if (old.getLastRow() < 2) return;

    old.getRange(2, 1, old.getLastRow() - 1, 13).getValues().forEach(function(r) {
      const ref = String(r[0] || "");
      if (!ref || done[ref]) return;

      let items = [], photos = [];
      try { items = JSON.parse(r[11] || "[]"); } catch (e) {}
      try { photos = JSON.parse(r[12] || "[]"); } catch (e) {}

      appendPropertyRow_(sheets, {
        ref: ref,
        submittedAt: r[1] instanceof Date ? r[1] : new Date(),
        clientId: String(r[2] || ""),
        room: r[3], roomType: r[4], kind: r[5],
        date: r[6], time: r[7], staff: r[8], notes: r[9],
        items: items, photos: photos
      });
      done[ref] = true;
      copied++;
    });
  });

  Logger.log("Copied " + copied + " inspection(s) into the new layout.");
}

function getNextPropertyCheckReference_(sheet, datePart) {
  const prefix = "PC-" + datePart + "-";
  let sequence = 1;
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().forEach(function(row) {
      const ref = row[0];
      if (ref && ref.indexOf(prefix) === 0) {
        const number = Number(ref.slice(prefix.length));
        if (Number.isInteger(number) && number >= sequence) sequence = number + 1;
      }
    });
  }

  return prefix + String(sequence).padStart(4, "0");
}

/* ================= Photos / references ================= */

function getPhotoFolder() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function saveImage(dataUrl, fileName, folder) {
  const m = String(dataUrl).match(/^data:(image\/[a-z+.-]+);base64,(.+)$/i);
  if (!m) throw new Error("The photo could not be read. Please retake it.");

  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], fileName);
  if (blob.getBytes().length > 4 * 1024 * 1024) throw new Error("The photo is too large. Please retake it.");

  return folder.createFile(blob).getUrl();
}

function getNextReference(sheet, datePart) {
  const lastRow = sheet.getLastRow();
  let sequence = 1;

  if (lastRow >= 2) {
    const refs = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    const prefix = LODGE_PREFIX + "-" + datePart + "-";
    refs.forEach(function(row) {
      const ref = row[0];
      if (ref && ref.indexOf(prefix) === 0) {
        const number = Number(ref.slice(prefix.length));
        if (Number.isInteger(number) && number >= sequence) sequence = number + 1;
      }
    });
  }

  return LODGE_PREFIX + "-" + datePart + "-" + String(sequence).padStart(4, "0");
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Visitor timer: status + admin ================= */

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet_() {
  return getSpreadsheet_().getSheetByName(SHEET_NAME);
}

function appendGuestRegistration_(ss, d, submittedAt) {
  let sheet = ss.getSheetByName(REGISTRATION_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(REGISTRATION_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setValues([REGISTRATION_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setFontWeight("bold");
  }

  const values = [
    submittedAt, d.type || "", d.name || "", d.gender || "", d.age || "", d.phone || "",
    d.email || "", d.address || "", d.nationality || "", d.idType || "", d.idNo || "",
    d.purpose || "", d.roomOrVisiting || "", d.persons || "", d.checkout || "", d.vehicle || ""
  ];

  const safeValues = values.map(function(value, index) {
    if (index === 0) return value;
    if (value === null || value === undefined) return "";
    const s = String(value);
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  });

  sheet.appendRow(safeValues);
}

function getStatus_(ref) {
  const sheet = getSheet_();

  if (!sheet || sheet.getLastRow() < 2 || !ref) return { ok: true, status: "", now: Date.now() };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 25).getValues();

  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) === String(ref)) {
      return {
        ok: true,
        ref: String(ref),
        status: String(rows[i][23] || ""),
        start: rows[i][1] instanceof Date ? rows[i][1].getTime() : 0,
        now: Date.now()
      };
    }
  }

  return { ok: true, status: "", now: Date.now() };
}

function adminAuth_(d) {
  const pin = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");
  if (!pin) throw new Error("Admin PIN is not set in Script properties.");

  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("pin_fails") || 0);
  if (fails >= 5) throw new Error("Too many wrong PIN attempts. Try again after 10 minutes.");

  if (String(d.pin || "") !== pin) {
    cache.put("pin_fails", String(fails + 1), 600);
    Utilities.sleep(1500);
    throw new Error("Wrong PIN.");
  }

  cache.remove("pin_fails");
}

function adminList_(d) {
  adminAuth_(d);

  const sheet = getSheet_();
  const out = [];

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 25).getValues();
    const cutoff = Date.now() - 12 * 3600 * 1000;

    rows.forEach(function(r) {
      if (r[2] !== "Visitor") return;

      const status = r[23];
      const start = r[1] instanceof Date ? r[1].getTime() : 0;
      const stopped = r[24] instanceof Date ? r[24].getTime() : 0;

      if (status === "Active" || (status === "Stopped" && stopped > cutoff)) {
        out.push({ ref: r[0], name: r[3], meeting: r[17], mobile: r[20], start: start, status: status, stopped: stopped });
      }
    });
  }

  return jsonResponse({ ok: true, now: Date.now(), visitors: out.reverse() });
}

function adminStop_(d) {
  adminAuth_(d);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sheet = getSheet_();
    let row = 0;

    if (sheet && sheet.getLastRow() >= 2) {
      const refs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (let i = 0; i < refs.length; i++) {
        if (refs[i][0] === String(d.ref)) { row = i + 2; break; }
      }
    }

    if (!row) throw new Error("Registration not found.");

    sheet.getRange(row, 24, 1, 2).setValues([["Stopped", new Date()]]);
    return jsonResponse({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/* ================= HOTEL SYNC (hotel.html) =================
   hotel.html sends its full data; this script
   1) keeps an exact copy in the hidden tab "Hotel Data" (used by "Restore from Sheet")
   2) writes readable tabs: Hotel Guests, Hotel GST Bills, Hotel Non-GST Bills,
      Hotel Check-in-out Audit, Hotel Accounts, Hotel Rooms, Hotel Housekeeping, Hotel Cash Closings
   Each tab is written independently: if one fails, the others are still written and the
   problem is reported back to the app (shown in the status line) and written to the log.
   ============================================================ */

function hotelAuth_(token) {
  const real = PropertiesService.getScriptProperties().getProperty("HOTEL_TOKEN");
  if (!real) throw new Error("HOTEL_TOKEN is not set in Script properties (run setSecrets()).");
  if (String(token || "") !== real) throw new Error("Wrong token. Type the same HOTEL_TOKEN in the app's Secret token box.");
}

function hotelGet_(p) {
  try {
    hotelAuth_(p.token);
    const st = hotelReadStored_();
    if (p.action === "rev") return jsonResponse({ ok: true, rev: st.rev });
    if (!st.json) return jsonResponse({ ok: false, error: "No data saved in the Sheet yet." });
    return jsonResponse({ ok: true, rev: st.rev, data: JSON.parse(st.json) });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message || "Hotel read failed." });
  }
}

function hotelPush_(d) {
  hotelAuth_(d.token);
  const data = d.data;
  if (!data || typeof data !== "object") throw new Error("No hotel data received.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet_();
    const stored = hotelReadStored_();
    const rev = Number(data.rev) || 0;

    // Never overwrite newer Sheet data with an older copy from another computer.
    if (rev < stored.rev) {
      return jsonResponse({
        ok: false,
        error: "Sheet already has a newer version (" + stored.rev + ") than this computer (" + rev +
               "). Click 'Restore from Sheet' on this computer first."
      });
    }

    const warnings = hotelWriteReadable_(ss, data);   // readable tabs first
    hotelWriteData_(ss, rev, JSON.stringify(data));   // version number is saved LAST

    console.log("Hotel sync OK: rev=" + rev + ", guests=" + (data.guests || []).length +
                (warnings.length ? ", warnings=" + warnings.join(" | ") : ""));

    return jsonResponse({ ok: true, rev: rev, warnings: warnings.join(" | ") });
  } finally {
    lock.releaseLock();
  }
}

function hotelReadStored_() {
  const sh = getSpreadsheet_().getSheetByName(HOTEL_DATA_SHEET);
  if (!sh || sh.getLastRow() < 2) return { rev: 0, json: "" };
  const rev = Number(sh.getRange(1, 2).getValue()) || 0;
  const parts = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]); });
  return { rev: rev, json: parts.join("") };
}

function hotelWriteData_(ss, rev, json) {
  let sh = ss.getSheetByName(HOTEL_DATA_SHEET);
  const isNew = !sh;
  if (!sh) sh = ss.insertSheet(HOTEL_DATA_SHEET);
  sh.clear();

  const chunks = [];
  for (let i = 0; i < json.length; i += HOTEL_CHUNK) chunks.push([json.slice(i, i + HOTEL_CHUNK)]);

  sh.getRange(1, 1, 1, 2).setValues([["rev (do not edit this tab)", rev]]);
  sh.getRange(2, 1, chunks.length, 1).setNumberFormat("@").setValues(chunks);
  if (isNew) sh.hideSheet();
}

function hotelCell_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? "Yes" : "No";
  if (typeof v === "object") return safeText_(JSON.stringify(v));
  let s = String(v);
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}.*Z$/.test(s)) {
    const dt = new Date(s);
    if (!isNaN(dt.getTime())) return Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s)) s = s.replace("T", " ");
  return safeText_(s);
}

function hotelWriteTable_(ss, name, headers, rows, textCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#642b27").setFontColor("#f2bb4e");
  sh.setFrozenRows(1);
  if (rows.length) {
    (textCols || []).forEach(function(c) { sh.getRange(2, c, rows.length, 1).setNumberFormat("@"); });
    sh.getRange(2, 1, rows.length, headers.length)
      .setValues(rows.map(function(r) { return r.map(hotelCell_); }));
  }
}

// Returns an array of warning strings (empty = everything written).
function hotelWriteReadable_(ss, db) {
  const warnings = [];
  const step = function(name, fn) {
    try { fn(); }
    catch (err) {
      const msg = name + ": " + (err && err.message ? err.message : err);
      warnings.push(msg);
      console.error("Hotel tab failed - " + msg);
    }
  };

  const guests = Array.isArray(db.guests) ? db.guests : [];

  step("Hotel Guests", function() {
    hotelWriteTable_(ss, "Hotel Guests",
      ["Guest ID", "Name", "Mobile", "Room", "Room Type / Plan", "Persons", "ID Type", "ID Reference",
       "Check-in", "Expected Check-out", "Check-out", "Rate", "GST Bill", "Booking Amount", "Booking Mode",
       "Advance Total", "Source", "Remarks", "Status", "Bill Total", "Tax", "Balance", "Payment Mode",
       "Invoice No", "Additional Guests"],
      guests.map(function(g) {
        const b = g.bill || {};
        const extra = (g.additionalGuests || []).map(function(a) {
          return a.personNo + ". " + (a.name || "") + (a.idType ? " (" + a.idType + " " + (a.idRef || "") + ")" : "");
        }).join("\n");
        return [g.id, g.name, g.mobile, g.room, g.plan, g.persons, g.idType, g.idRef,
          g.checkin, g.expected, g.checkout, g.rate, g.gst, g.bookingAmount, g.bookingAmountMode,
          g.advance, g.source, g.remarks, g.status, b.total, b.tax, b.balance, b.payment,
          b.invoiceNo, extra];
      }), [3, 8]);
  });

  const billHeaders = ["Invoice No", "Checkout Date", "Guest ID", "Guest Name", "Mobile", "Room", "Room Type",
    "Persons", "Nights", "Taxable Value", "GST Rate %", "CGST", "SGST", "Invoice Total", "Booking Received",
    "Advance Received", "Checkout Received", "Total Received", "Balance", "Payment Mode", "Invoice Type"];
  const billKeys = ["invoiceNo", "date", "guestId", "guestName", "mobile", "room", "roomType", "persons", "nights",
    "taxableValue", "gstRate", "cgst", "sgst", "total", "bookingReceived", "advanceReceived", "checkoutReceived",
    "totalReceived", "balance", "paymentMode", "invoiceType"];
  const billRows = function(list) {
    return (Array.isArray(list) ? list : []).map(function(b) {
      return billKeys.map(function(k) { return b[k]; });
    });
  };
  step("Hotel GST Bills", function() { hotelWriteTable_(ss, "Hotel GST Bills", billHeaders, billRows(db.auditBills), [5]); });
  step("Hotel Non-GST Bills", function() { hotelWriteTable_(ss, "Hotel Non-GST Bills", billHeaders, billRows(db.nonGstBills), [5]); });

  step("Hotel Check-in-out Audit", function() {
    const billByGuest = {};
    (Array.isArray(db.nonGstBills) ? db.nonGstBills : []).forEach(function(b) { billByGuest[b.guestId] = b; });
    (Array.isArray(db.auditBills) ? db.auditBills : []).forEach(function(b) { billByGuest[b.guestId] = b; });
    hotelWriteTable_(ss, "Hotel Check-in-out Audit",
      ["Guest ID", "Guest Name", "Mobile", "Room", "Persons", "Check-in", "Expected Check-out", "Actual Check-out",
       "Nights", "Stay Status", "Invoice No", "Invoice Type", "Taxable Value", "CGST", "SGST", "Invoice Total",
       "Booking Received", "Advance Received", "Checkout Received", "Total Received", "Balance", "Payment Mode"],
      guests.slice().sort(function(a, b) {
        return String(a.checkin || "").localeCompare(String(b.checkin || ""));
      }).map(function(g) {
        const b = billByGuest[g.id] || {};
        const gb = g.bill || {};
        return [g.id, g.name, g.mobile, g.room, g.persons, g.checkin, g.expected, g.checkout,
          b.nights, g.status, b.invoiceNo || gb.invoiceNo, b.invoiceType || gb.invoiceType,
          b.taxableValue, b.cgst, b.sgst, b.total !== undefined ? b.total : gb.total,
          b.bookingReceived !== undefined ? b.bookingReceived : g.bookingAmount,
          b.advanceReceived !== undefined ? b.advanceReceived : g.advance,
          b.checkoutReceived, b.totalReceived,
          b.balance !== undefined ? b.balance : gb.balance, b.paymentMode || gb.payment];
      }), [3]);
  });

  step("Hotel Accounts", function() {
    hotelWriteTable_(ss, "Hotel Accounts",
      ["ID", "Date / Time", "Type", "Details", "Mode", "Amount", "Reference", "Staff", "Guest ID", "Auto"],
      (Array.isArray(db.accounts) ? db.accounts : []).map(function(a) {
        return [a.id, a.date, a.kind, a.details, a.mode, a.amount, a.ref, a.staff, a.guestId, a.auto];
      }));
  });

  step("Hotel Rooms", function() {
    const rooms = db.rooms && typeof db.rooms === "object" ? db.rooms : {};
    const nameById = {};
    guests.forEach(function(g) { nameById[g.id] = g.name; });
    hotelWriteTable_(ss, "Hotel Rooms",
      ["Room", "Type", "Status", "Current Guest ID", "Current Guest"],
      Object.keys(rooms).sort().map(function(n) {
        const r = rooms[n] || {};
        return [n, r.type || r.plan || "", r.status, r.guest || "", r.guest ? (nameById[r.guest] || "") : ""];
      }));
  });

  step("Hotel Housekeeping", function() {
    hotelWriteTable_(ss, "Hotel Housekeeping",
      ["Date / Time", "Room", "Status", "Staff", "Notes"],
      (Array.isArray(db.logs) ? db.logs : []).map(function(l) {
        return [l.date, l.room, l.status, l.staff, l.notes];
      }));
  });

  step("Hotel Cash Closings", function() {
    const closings = db.cashClosings && typeof db.cashClosings === "object" ? db.cashClosings : {};
    const opening = db.openingCash && typeof db.openingCash === "object" ? db.openingCash : {};
    const keys = Object.keys(Object.assign({}, opening, closings)).sort();
    hotelWriteTable_(ss, "Hotel Cash Closings",
      ["Date_Shift", "Opening Cash", "Actual Cash", "Expected Cash", "Variance", "Notes", "Saved At"],
      keys.map(function(k) {
        const c = closings[k] || {};
        return [k, opening[k], c.actualCash, c.expectedCash, c.variance, c.notes, c.savedAt];
      }));
  });

  return warnings;
}

/* ================= Run once from the editor ================= */

// STEP 1: edit the two values, run this once. Leave a value as "" to keep what is already saved.
function setSecrets() {
  const HOTEL_TOKEN_VALUE = "TCR_Hotel@2026-Secret";   // e.g. "Hotel@2026-Secret"  (type the SAME text in hotel.html)
  const ADMIN_PIN_VALUE = "1212";     // e.g. "4821"               (staff PIN for admin / property check)

  const props = PropertiesService.getScriptProperties();
  if (HOTEL_TOKEN_VALUE) props.setProperty("HOTEL_TOKEN", HOTEL_TOKEN_VALUE);
  if (ADMIN_PIN_VALUE) props.setProperty("ADMIN_PIN", ADMIN_PIN_VALUE);

  Logger.log("HOTEL_TOKEN set: " + !!props.getProperty("HOTEL_TOKEN") +
             " | ADMIN_PIN set: " + !!props.getProperty("ADMIN_PIN"));
}

// STEP 2: creates every tab.
function setupHeaders() {
  const ss = getSpreadsheet_();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");

  let registrationSheet = ss.getSheetByName(REGISTRATION_SHEET_NAME);
  if (!registrationSheet) registrationSheet = ss.insertSheet(REGISTRATION_SHEET_NAME);
  registrationSheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setValues([REGISTRATION_HEADERS]);
  registrationSheet.setFrozenRows(1);
  registrationSheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setFontWeight("bold");

  ensurePropertySheets_(ss);
  hotelWriteReadable_(ss, { guests: [], rooms: {}, accounts: [], logs: [], auditBills: [], nonGstBills: [] });
}

function setupHotelSheets() {
  hotelWriteReadable_(getSpreadsheet_(), { guests: [], rooms: {}, accounts: [], logs: [], auditBills: [], nonGstBills: [] });
  Logger.log("Hotel tabs created.");
}

// STEP 3: checks everything the hotel sync needs. Every line should start with OK.
function hotelDiagnose() {
  const props = PropertiesService.getScriptProperties();
  Logger.log((props.getProperty("HOTEL_TOKEN") ? "OK" : "PROBLEM") + " - HOTEL_TOKEN property " +
             (props.getProperty("HOTEL_TOKEN") ? "is set" : "is MISSING (run setSecrets)"));
  Logger.log((props.getProperty("ADMIN_PIN") ? "OK" : "PROBLEM") + " - ADMIN_PIN property " +
             (props.getProperty("ADMIN_PIN") ? "is set" : "is MISSING (run setSecrets)"));

  try {
    const ss = getSpreadsheet_();
    Logger.log("OK - spreadsheet opens: " + ss.getName());

    const warnings = hotelWriteReadable_(ss, {
      guests: [{ id: "TEST00001", name: "Test Guest", mobile: "0987654321", room: "101", plan: "Double Non-A/C",
                 persons: 1, checkin: "2026-11-02T10:00", status: "Occupied", bookingAmount: 500, advance: 0 }],
      rooms: { "101": { number: "101", type: "Double Non-A/C", status: "Occupied", guest: "TEST00001" } },
      accounts: [], logs: [], auditBills: [], nonGstBills: []
    });
    Logger.log(warnings.length ? "PROBLEM - " + warnings.join(" | ")
                               : "OK - all 8 hotel tabs written (a test guest row was added; the next real sync replaces it)");
  } catch (err) {
    Logger.log("PROBLEM - cannot open/write spreadsheet: " + err.message);
  }

  Logger.log("Stored sheet version: " + hotelReadStored_().rev);
  Logger.log("Now deploy a NEW VERSION (Deploy > Manage deployments > edit > New version).");
}

function hotelSelfTest() { hotelDiagnose(); }

function authorizeDrive() {
  Logger.log("Drive OK: " + getPhotoFolder().getUrl() + " | " + getPropertyPhotoFolder_().getUrl());
}

function testFirebaseKey() {
  const res = UrlFetchApp.fetch(
    "https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=" + encodeURIComponent(getFirebaseKey_()),
    { method: "post", contentType: "application/json", payload: JSON.stringify({ idToken: "test" }), muteHttpExceptions: true }
  );
  Logger.log(res.getResponseCode() + " " + res.getContentText());
}
