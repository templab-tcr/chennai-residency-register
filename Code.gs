const FIREBASE_API_KEY = "AIzaSyA-Cn6-j_iZi7oN_yaaKjHsCG4-_JxHyiE";

const SHEET_NAME = "Guest Register";
const LODGE_PREFIX = "CRL";
const PHOTO_FOLDER_NAME = "Guest ID Photos";   // created automatically in your Drive (private)

const HEADERS = [
  "Reference Number",
  "Registration Time",
  "Registration Type",
  "Full Name",
  "Gender",
  "Age",
  "Email",
  "Home Address",
  "Coming From",
  "Purpose of Visit",
  "Nationality",
  "ID Type",
  "ID Number",
  "Passport Number",
  "Visa Number",
  "Visa Expiry",
  "Total Persons",
  "Room Number / Meeting Guest",
  "Expected Check-out",
  "Vehicle Number",
  "Verified Mobile Number",
  "ID Front Photo",
  "ID Back Photo",
  "Status",
  "Stopped At"
];

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === "status") {
    return jsonResponse(getStatus_(p.ref));
  }

  return jsonResponse({
    ok: true,
    message: "Guest Register service is running."
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No registration data received.");
    }

    const d = JSON.parse(e.postData.contents);

    // Admin actions (protected by ADMIN_PIN in Script properties)
    if (d.action === "adminList") return adminList_(d);
    if (d.action === "adminStop") return adminStop_(d);

    if (!FIREBASE_API_KEY ||
        FIREBASE_API_KEY === "PASTE_FIREBASE_API_KEY_HERE") {
      throw new Error("Firebase API key is not configured.");
    }

    if (!d.idToken) {
      throw new Error("Phone verification is required.");
    }

    // Verify the Firebase ID token with Firebase Authentication.
    const verifyUrl =
      "https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=" +
      encodeURIComponent(FIREBASE_API_KEY);

    const verification = UrlFetchApp.fetch(verifyUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ idToken: d.idToken }),
      muteHttpExceptions: true
    });

    const verifyStatus = verification.getResponseCode();
    const verifyResult = JSON.parse(verification.getContentText());

    if (verifyStatus !== 200 ||
        !verifyResult.users ||
        verifyResult.users.length === 0) {
      throw new Error("Phone verification failed. Please verify again.");
    }

    const user = verifyResult.users[0];

    if (!user.phoneNumber) {
      throw new Error("A verified phone number is required.");
    }

    // Ensure the submitted phone matches the verified Firebase user.
    const submittedPhone =
      String(d.phone || "").replace(/\s/g, "");

    if (submittedPhone &&
        submittedPhone !== user.phoneNumber) {
      throw new Error("Verified phone does not match the form.");
    }

    if (!["Guest", "Visitor"].includes(d.type)) {
      throw new Error("Invalid registration type.");
    }

    if (!d.name || String(d.name).trim().length < 2) {
      throw new Error("A valid full name is required.");
    }

    if (!d.address || String(d.address).trim().length < 5) {
      throw new Error("A valid home address is required.");
    }

    if (!d.idType || !d.idNo) {
      throw new Error("ID proof details are required.");
    }

    if (!d.idFront || !d.idBack) {
      throw new Error("Front and back photos of the ID are required.");
    }

    if (d.nationality !== "Indian" &&
        (!d.passport || !d.visa || !d.visaExp)) {
      throw new Error("Passport and visa details are required.");
    }

    if (!d.purpose || !d.roomOrVisiting) {
      throw new Error("Purpose and room or meeting details are required.");
    }

    if (d.type === "Guest" && !d.checkout) {
      throw new Error("Expected check-out date is required.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      if (!ss) {
        throw new Error("The script is not linked to a spreadsheet.");
      }

      let sheet = ss.getSheetByName(SHEET_NAME);

      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
      }

      // Write headers on a new sheet, or add the two new photo headers to an old one.
      if (sheet.getLastRow() === 0 ||
          sheet.getRange(1, HEADERS.length).getValue() !== HEADERS[HEADERS.length - 1]) {
        sheet.getRange(1, 1, 1, HEADERS.length)
          .setValues([HEADERS]);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, HEADERS.length)
          .setFontWeight("bold");
      }

      const now = new Date();
      const datePart = Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "yyMMdd"
      );

      const reference = getNextReference(sheet, datePart);

      // Save the ID photos to a private Drive folder and keep the links.
      const folder = getPhotoFolder();
      const frontUrl = saveImage(d.idFront, reference + "_front.jpg", folder);
      const backUrl = saveImage(d.idBack, reference + "_back.jpg", folder);

      const values = [
        reference,
        now,
        d.type,
        d.name,
        d.gender,
        d.age,
        d.email,
        d.address,
        d.comingFrom,
        d.purpose,
        d.nationality,
        d.idType,
        d.idNo,
        d.passport,
        d.visa,
        d.visaExp,
        d.persons,
        d.roomOrVisiting,
        d.checkout,
        d.vehicle,
        user.phoneNumber,
        frontUrl,
        backUrl,
        d.type === "Visitor" ? "Active" : "Checked-in",
        ""
      ];

      // Prevent user-entered text from being interpreted as formulas.
      const safeValues = values.map(function(value, index) {
        if (index === 1) return value;
        if (value === null || value === undefined) return "";

        const s = String(value);

        return /^[=+\-@]/.test(s) ? "'" + s : s;
      });

      sheet.appendRow(safeValues);

      return jsonResponse({
        ok: true,
        ref: reference,
        time: Utilities.formatDate(
          now,
          Session.getScriptTimeZone(),
          "dd-MMM-yyyy hh:mm a"
        )
      });

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    return jsonResponse({
      ok: false,
      error: err.message || "Registration could not be saved."
    });
  }
}

function getPhotoFolder() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function saveImage(dataUrl, fileName, folder) {
  const m = String(dataUrl).match(/^data:(image\/[a-z+.-]+);base64,(.+)$/i);

  if (!m) {
    throw new Error("The ID photo could not be read. Please retake it.");
  }

  const blob = Utilities.newBlob(
    Utilities.base64Decode(m[2]),
    m[1],
    fileName
  );

  if (blob.getBytes().length > 4 * 1024 * 1024) {
    throw new Error("The ID photo is too large. Please retake it.");
  }

  return folder.createFile(blob).getUrl();
}

function getNextReference(sheet, datePart) {
  const lastRow = sheet.getLastRow();
  let sequence = 1;

  if (lastRow >= 2) {
    const refs = sheet.getRange(2, 1, lastRow - 1, 1)
      .getDisplayValues();

    const prefix = LODGE_PREFIX + "-" + datePart + "-";

    refs.forEach(function(row) {
      const ref = row[0];

      if (ref && ref.indexOf(prefix) === 0) {
        const number = Number(ref.slice(prefix.length));

        if (Number.isInteger(number) && number >= sequence) {
          sequence = number + 1;
        }
      }
    });
  }

  return LODGE_PREFIX + "-" + datePart + "-" +
    String(sequence).padStart(4, "0");
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= Visitor timer: status + admin ================= */

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

// Called by the visitor's page every few seconds.
function getStatus_(ref) {
  const sheet = getSheet_();

  if (!sheet || sheet.getLastRow() < 2 || !ref) {
    return { ok: true, status: "" };
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 24).getDisplayValues();

  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === String(ref)) {
      return { ok: true, status: rows[i][23] };
    }
  }

  return { ok: true, status: "" };
}

function adminAuth_(d) {
  const pin = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");

  if (!pin) {
    throw new Error("Admin PIN is not set in Script properties.");
  }

  if (String(d.pin || "") !== pin) {
    Utilities.sleep(1500);
    throw new Error("Wrong PIN.");
  }
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
        out.push({
          ref: r[0],
          name: r[3],
          meeting: r[17],
          mobile: r[20],
          start: start,
          status: status,
          stopped: stopped
        });
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

/* ================= Run once from the editor ================= */

function setupHeaders() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
}

function authorizeDrive() {
  Logger.log("Drive OK: " + getPhotoFolder().getUrl());
}
