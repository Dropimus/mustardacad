/**
 * Mustard Academy — Polymarket class application backend.
 *
 * What this does: records each application (name, username, phone) in a
 * Google Sheet and prevents the same phone number from being recorded
 * twice. It does NOT verify that an applicant actually followed on X or
 * joined the WhatsApp community — no serverless function can check that
 * without each platform's own API access, so this only tracks self-reported
 * checklist completion plus real submitted contact details.
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
 * it runs, with header row: Timestamp | Name | Username | Phone
 */

const SHEET_NAME = "Applications";
const TELEGRAM_LINK = "https://t.me/MustardAcademy";

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond_({ status: "error", message: "Invalid request." });
  }

  const name = (body.name || "").toString().trim();
  const username = (body.username || "").toString().trim();
  const phoneRaw = (body.phone || "").toString().trim();
  const phone = phoneRaw.replace(/\D/g, "");

  if (name.length < 2 || phone.length < 7) {
    return respond_({ status: "error", message: "Missing or invalid name/phone." });
  }

  const sheet = getSheet_();
  if (findByPhone_(sheet, phone)) {
    return respond_({ status: "duplicate", telegramLink: TELEGRAM_LINK });
  }

  sheet.appendRow([new Date(), name, username, phoneRaw]);
  return respond_({ status: "confirmed", telegramLink: TELEGRAM_LINK });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Name", "Username", "Phone"]);
  }
  return sheet;
}

function findByPhone_(sheet, phone) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowPhone = (values[i][3] || "").toString().replace(/\D/g, "");
    if (rowPhone && rowPhone === phone) return values[i];
  }
  return null;
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
