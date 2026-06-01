const SHEET_NAME = "Registrations";

const PAYMENT_PROOFS_FOLDER_ID = "1rnqd-HDnFykqHnoPF43FW4qXXT3kqwZj";

const HEADERS = [
  "Timestamp",
  "Full Name",
  "Phone Number",
  "Email",
  "Age",
  "Gender",
  "Team Registration",
  "Teammate Names",
  "How heard about event",
  "Referred by (name)",
  "Payment proof file name",
  "Payment proof link",
  "Terms accepted",
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Registration form")
    .addItem("Fix column layout", "fixRegistrationSheetLayout")
    .addItem("Test Drive upload", "testPaymentProofUploadToDrive")
    .addToUi();
}

function fixRegistrationSheetLayout() {
  const sheet = getOrCreateSheet_();
  ensureHeaders_(sheet);
  SpreadsheetApp.getUi().alert("Registrations tab headers refreshed.");
}

function doGet() {
  return ContentService.createTextOutput(
    "Web app is live. Submissions are saved to the sheet tab named: " + SHEET_NAME
  ).setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    Logger.log("doPost received");
    const sheet = getOrCreateSheet_();
    ensureHeaders_(sheet);

    const fullName           = clean_(e.parameter.fullName);
    const phoneNumber        = clean_(e.parameter.phoneNumber);
    const email              = clean_(e.parameter.email);
    const age                = clean_(e.parameter.age);
    const gender             = clean_(e.parameter.gender);
    const isTeam             = clean_(e.parameter.isTeam);
    const teammateNames      = clean_(e.parameter.teammateNames);
    const heardHow           = clean_(e.parameter.heardHow);
    const heardPersonName    = clean_(e.parameter.heardPersonName);
    const paymentProofName   = clean_(e.parameter.paymentProofName);
    const paymentProofBase64 = clean_(e.parameter.paymentProofBase64);
    const paymentProofMime   = clean_(e.parameter.paymentProofMime) || "image/png";
    const termsAccepted      = clean_(e.parameter.termsAccepted);

    if (!fullName || !phoneNumber || !email || !age || !gender) {
      return json_({ ok: false, message: "Missing required fields." });
    }
    if (!heardHow) {
      return json_({ ok: false, message: "Please tell us how you heard about the event." });
    }
    if (heardHow === "Through a person" && !heardPersonName) {
      return json_({ ok: false, message: "Please enter the name of the person who told you." });
    }
    if (!isTeam) {
      return json_({ ok: false, message: "Please indicate whether you are registering as a team." });
    }
    if (isTeam === "yes" && !teammateNames) {
      return json_({ ok: false, message: "Please enter your teammates' names." });
    }
    if (termsAccepted !== "yes") {
      return json_({ ok: false, message: "Please accept the event terms to continue." });
    }

    var paymentProofLink = "";
    if (paymentProofBase64 && paymentProofName) {
      try {
        paymentProofLink = savePaymentProofToDrive_(paymentProofBase64, paymentProofMime, paymentProofName, fullName);
      } catch (driveErr) {
        Logger.log("Drive upload error: " + driveErr);
        paymentProofLink = "Drive upload failed — " + String(driveErr).slice(0, 120) + " | Registrant: " + email;
      }
    } else {
      paymentProofLink = "No file received — ask registrant to resend a smaller screenshot.";
    }

    Logger.log("Appending row for: " + fullName);
    sheet.appendRow([
      new Date(),
      fullName,
      phoneNumber,
      email,
      age,
      gender,
      isTeam === "yes" ? "Yes" : "No",
      isTeam === "yes" ? teammateNames : "",
      heardHow,
      heardHow === "Through a person" ? heardPersonName : "",
      paymentProofName || "(none)",
      paymentProofLink,
      "Yes",
    ]);

    return json_({ ok: true });
  } catch (err) {
    Logger.log("doPost error: " + err);
    return json_({ ok: false, message: String(err) });
  }
}

function savePaymentProofToDrive_(base64, mime, fileName, registrantName) {
  const bytes = Utilities.base64Decode(base64);
  const safeName =
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm") +
    "_" +
    String(registrantName || "registrant").replace(/[^\w\s-]/g, "").slice(0, 40) +
    "_" +
    (fileName || "proof.png");
  const blob = Utilities.newBlob(bytes, mime, safeName);
  const folder = getPaymentProofsTargetFolder_();
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {
    Logger.log("setSharing skipped: " + shareErr);
  }
  return file.getUrl();
}

function testPaymentProofUploadToDrive() {
  const tinyPngBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
  const url = savePaymentProofToDrive_(tinyPngBase64, "image/png", "test-proof.png", "DriveTest");
  SpreadsheetApp.getUi().alert("Drive upload OK!\n\n" + url);
}

function getPaymentProofsTargetFolder_() {
  var folderId = String(PAYMENT_PROOFS_FOLDER_ID || "").trim();
  var m = folderId.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) folderId = m[1].split(/[?#]/)[0];
  if (!folderId) throw new Error("PAYMENT_PROOFS_FOLDER_ID is not set in Code.gs.");
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    throw new Error("Cannot open Drive folder id: " + folderId + ". Check the ID and Drive permissions. " + err);
  }
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0]
    .map(function(h) { return String(h || "").trim(); });
  const matches = HEADERS.every(function(h, i) { return existing[i] === h; });
  if (!matches) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function clean_(v) {
  return String(v || "").trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
