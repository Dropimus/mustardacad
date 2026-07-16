/**
 * Mustard Academy — Polymarket class application backend.
 *
 * What this does: records each application (name, username, email) in a
 * Google Sheet, prevents the same email from being recorded twice, and
 * issues each applicant an Application ID. Follow/community verification
 * itself is done manually by an admin (checking the X followers list and
 * the submitted email), not by this script — no serverless function can
 * check that without each platform's own API access. The Application ID
 * exists so the applicant can quote it when requesting to join Telegram,
 * letting the Telegram admin cross-reference it against this sheet.
 *
 * Deploy:
 *   1. Create a Google Sheet (or open an existing one).
 *   2. Extensions -> Apps Script, delete the placeholder code, paste this file.
 *   3. Deploy -> New deployment -> type "Web app".
 *   4. Execute as: Me. Who has access: Anyone.
 *   5. Deploy, then copy the "Web app URL" (ends in /exec).
 *   6. Paste that URL into CONFIG.endpoint in polymarket-apply.html.
 *
 * The script creates an "Applications" sheet automatically the first time
 * it runs, with header row: Timestamp | Name | Username | Email | Application ID
 */

const SHEET_NAME = "Applications";
const TELEGRAM_LINK = "https://t.me/MustardAcademy";
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond_({ status: "error", message: "Invalid request." });
  }

  const name = (body.name || "").toString().trim();
  const username = (body.username || "").toString().trim();
  const email = (body.email || "").toString().trim().toLowerCase();

  if (name.length < 2 || !EMAIL_RE.test(email)) {
    return respond_({ status: "error", message: "Missing or invalid name/email." });
  }

  const sheet = getSheet_();
  const existing = findByEmail_(sheet, email);
  if (existing) {
    return respond_({ status: "duplicate", appId: existing[4], telegramLink: TELEGRAM_LINK });
  }

  const appId = generateAppId_();
  sheet.appendRow([new Date(), name, username, email, appId]);
  return respond_({ status: "confirmed", appId: appId, telegramLink: TELEGRAM_LINK });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Name", "Username", "Email", "Application ID"]);
  }
  return sheet;
}

function findByEmail_(sheet, email) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowEmail = (values[i][3] || "").toString().trim().toLowerCase();
    if (rowEmail && rowEmail === email) return values[i];
  }
  return null;
}

function generateAppId_() {
  return "MA-" + Utilities.getUuid().split("-")[0].toUpperCase();
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
