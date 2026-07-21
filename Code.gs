/**
 * Mustard Academy — Polymarket class application backend.
 *
 * What this does: records every lead submitted from any of the site's
 * forms (main site, Polymarket page, Polymarket application) in one
 * Google Sheet, prevents the same email from being recorded twice, and
 * issues each applicant an Application ID. Follow/community verification
 * (X follow, Telegram, Polymarket account) is done manually by an admin,
 * not by this script — no serverless function can check that without
 * each platform's own API access.
 *
 * The one exception is the "invite a friend" referral step: each
 * applicant gets a unique referral code baked into their personal
 * invite link (?ref=CODE). When someone submits the form after arriving
 * via that link, this script logs the code that referred them. The
 * "Invite a friend" checklist step only completes once a GET request
 * confirms at least one submission carries the applicant's code — that
 * part is real, verifiable, and needs no external API.
 *
 * Referral hardening: a submission can't credit itself — if the
 * "referredBy" code matches the submitter's own "refCode" (i.e. they
 * carried their own invite link's ?ref= param, typically by opening it
 * themselves), the referral is dropped rather than counted. Both codes
 * are also validated against the shape genRefCode() actually produces,
 * so junk/hand-typed values in ?ref= can't pollute the sheet or be used
 * to probe doGet. The Polymarket application flow additionally requires
 * a non-empty X handle and WhatsApp number server-side, mirroring the
 * checklist gating in polymarket-apply.html, so those steps can't be
 * skipped by calling this endpoint directly instead of using the page.
 * None of this proves the referred signup is a distinct real person —
 * that would need real identity/email verification — it only closes the
 * "open your own link and resubmit" and "call the API directly" gaps.
 *
 * Deploy:
 *   1. Create a Google Sheet (or open an existing one).
 *   2. Extensions -> Apps Script, delete the placeholder code, paste this file.
 *   3. Deploy -> New deployment -> type "Web app".
 *   4. Execute as: Me. Who has access: Anyone.
 *   5. Deploy, then copy the "Web app URL" (ends in /exec).
 *   6. Paste that URL into CONFIG.endpoint / ENDPOINT in index.html,
 *      polymarket.html, and polymarket-apply.html.
 *
 * The script creates an "Applications" sheet automatically the first
 * time it runs (and tops up the header row on an existing sheet), with
 * header row:
 *   Timestamp | Name | Username | Email | Track | Source |
 *   Application ID | Telegram | Polymarket | Referral Code | Referred By |
 *   WhatsApp
 *
 * New columns are appended after the original columns so existing rows
 * from before each change stay aligned.
 */

const SHEET_NAME = "Applications";
const TELEGRAM_LINK = "https://t.me/MustardAcademy";
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
// Matches what genRefCode() in the front-end actually produces (base-36 chars,
// uppercased). Anything outside this shape is treated as not a real code.
const CODE_RE = /^[A-Z0-9]{4,12}$/;
const HEADERS = ["Timestamp", "Name", "Username", "Email", "Track", "Source", "Application ID", "Telegram", "Polymarket", "Referral Code", "Referred By", "WhatsApp"];

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond_({ status: "error", message: "Invalid request." });
  }

  const name = (body.name || "").toString().trim();
  const xHandle = (body.xHandle || body.username || "").toString().trim();
  const telegram = (body.telegram || "").toString().trim();
  const polymarket = (body.polymarket || "").toString().trim();
  const whatsapp = (body.whatsapp || "").toString().trim();
  const email = (body.email || "").toString().trim().toLowerCase();
  const track = (body.track || "").toString().trim();
  const source = (body.source || "Polymarket Application").toString().trim();
  const refCode = normalizeCode_(body.refCode);
  let referredBy = normalizeCode_(body.referredBy);

  if (name.length < 2 || !EMAIL_RE.test(email)) {
    return respond_({ status: "error", message: "Missing or invalid name/email." });
  }

  // The apply page's checklist requires an X username and WhatsApp number
  // before submission unlocks — enforce that here too so it can't be
  // skipped by posting to this endpoint directly.
  if (source === "Polymarket Application") {
    if (xHandle.length < 2) {
      return respond_({ status: "error", message: "Missing X username." });
    }
    if (whatsapp.replace(/\D/g, "").length < 7) {
      return respond_({ status: "error", message: "Missing or invalid WhatsApp number." });
    }
  }

  // A ?ref= code that matches the submitter's own code is their own invite
  // link, not a friend's — don't let it count as a referral.
  if (referredBy && referredBy === refCode) {
    referredBy = "";
  }

  const sheet = getSheet_();
  const existing = findByEmail_(sheet, email);
  if (existing) {
    return respond_({ status: "duplicate", appId: existing[6], telegramLink: TELEGRAM_LINK });
  }

  const appId = generateAppId_();
  sheet.appendRow([new Date(), name, xHandle, email, track, source, appId, telegram, polymarket, refCode, referredBy, whatsapp]);
  return respond_({ status: "confirmed", appId: appId, telegramLink: TELEGRAM_LINK });
}

/**
 * GET ?code=REFCODE — used by the "invite a friend" step to check
 * whether anyone has submitted the application form with this referral
 * code, i.e. whether the applicant's invite link actually got used.
 */
function doGet(e) {
  const code = normalizeCode_(((e && e.parameter) || {}).code);
  if (!code) {
    return respond_({ used: false, count: 0 });
  }
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < values.length; i++) {
    const referredBy = (values[i][10] || "").toString().trim().toUpperCase();
    if (referredBy && referredBy === code) count++;
  }
  return respond_({ used: count > 0, count: count });
}

function normalizeCode_(v) {
  const code = (v || "").toString().trim().toUpperCase();
  return CODE_RE.test(code) ? code : "";
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  ensureHeaders_(sheet);
  return sheet;
}

function ensureHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < HEADERS.length) {
    sheet.getRange(1, lastCol + 1, 1, HEADERS.length - lastCol).setValues([HEADERS.slice(lastCol)]);
  }
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
