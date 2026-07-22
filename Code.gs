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
 * Referral signals (breaking the chicken-and-egg problem): the "Invite a
 * friend" checklist step is required to unlock submission, but the full
 * application form (name + email) is only shown after *all* checklist
 * steps — including "Invite a friend" — are done. If crediting a
 * referral required the friend to fully submit, nobody could ever be
 * first: submitting requires your own referral to be done, which
 * requires someone else to submit, which requires their own referral to
 * be done, forever. To break that, a referral is credited as soon as the
 * referred visitor fills in their X handle and WhatsApp number — the
 * first two checklist steps, which every visitor completes regardless of
 * referrals — via a lightweight "referralSignal" POST (see doPost),
 * logged to a separate "ReferralSignals" sheet rather than the main
 * Applications sheet, since no name/email exists yet at that point. This
 * is a materially weaker signal than a real submission (just two text
 * fields, no email dedupe), so it's still checked for the self-referral
 * and code-shape rules above, and lightly de-duped by refCode, but it
 * can't prove the "friend" is a distinct real person any more than the
 * fields above can. De-duping is keyed on the (refCode, referredBy) pair,
 * not refCode alone — a browser that previously sent a signal for one
 * referrer (e.g. an earlier test, or a self-referral that got nulled)
 * must still be able to credit a different, real referrer later.
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
 *   WhatsApp | Approved | X Verified | WhatsApp Verified |
 *   Polymarket Verified | Referral Verified
 *
 * New columns are appended after the original columns so existing rows
 * from before each change stay aligned.
 *
 * Manual review, per checklist step: each of the four apply-page tasks
 * (Follow X, Join WhatsApp, Open Polymarket, Invite a friend) gets its own
 * "* Verified" column, all starting blank/FALSE. As soon as a visitor
 * self-checks a task on the apply page (before they've even submitted
 * their name/email), a lightweight "taskProgress" POST creates or updates
 * a row for them keyed by their Referral Code — the one identifier the
 * front-end already has from the moment they land on the page — so an
 * admin can start reviewing (and typing TRUE into a Verified column)
 * without waiting for the full submission. When the applicant does submit
 * their full application, doPost fills in the rest of that same row
 * in place rather than creating a duplicate.
 *
 * "Referral Verified" is also auto-derived from real usage (see
 * countReferralUses_) — an admin can still force it TRUE manually, but
 * normally there's nothing to review since a real referral is already
 * server-verifiable. "Approved" (the field that unlocks the Telegram
 * button) is likewise auto-derived: true once all four Verified columns
 * read true, with a manual TRUE in Approved still available as an
 * override/escape hatch. See parseApproved_, statusForRow_.
 */

const SHEET_NAME = "Applications";
const PENDING_SHEET_NAME = "ReferralSignals";
const TELEGRAM_LINK = "https://t.me/MustardAcademy";
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
// Matches what genRefCode() in the front-end actually produces (base-36 chars,
// uppercased). Anything outside this shape is treated as not a real code.
const CODE_RE = /^[A-Z0-9]{4,12}$/;
const HEADERS = [
  "Timestamp", "Name", "Username", "Email", "Track", "Source",
  "Application ID", "Telegram", "Polymarket", "Referral Code", "Referred By", "WhatsApp",
  "Approved", "X Verified", "WhatsApp Verified", "Polymarket Verified", "Referral Verified"
];
const PENDING_HEADERS = ["Timestamp", "Ref Code", "Referred By", "X Handle", "WhatsApp"];
// 0-based column indexes into a HEADERS row, kept in sync with HEADERS above.
const COL_TIMESTAMP = 0;
const COL_NAME = 1;
const COL_USERNAME = 2;
const COL_EMAIL = 3;
const COL_TRACK = 4;
const COL_SOURCE = 5;
const COL_APP_ID = 6;
const COL_TELEGRAM = 7;
const COL_POLYMARKET = 8;
const COL_REF_CODE = 9;
const COL_REFERRED_BY = 10;
const COL_WHATSAPP = 11;
const COL_APPROVED = 12;
const COL_X_VERIFIED = 13;
const COL_WA_VERIFIED = 14;
const COL_PM_VERIFIED = 15;
const COL_REF_VERIFIED = 16;

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond_({ status: "error", message: "Invalid request." });
  }

  if ((body.action || "").toString().trim() === "referralSignal") {
    return handleReferralSignal_(body);
  }

  if ((body.action || "").toString().trim() === "taskProgress") {
    return handleTaskProgress_(body);
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
  const existingByEmail = findByEmail_(sheet, email);
  if (existingByEmail) {
    const dupStatus = statusForRow_(existingByEmail);
    return respond_({ status: "duplicate", appId: existingByEmail[COL_APP_ID], approved: dupStatus.approved, telegramLink: TELEGRAM_LINK });
  }

  const appId = generateAppId_();
  // A pre-submission taskProgress row may already exist for this Referral
  // Code — fill it in rather than appending a duplicate, so any Verified
  // columns an admin already set aren't lost.
  const rowNum = refCode ? findRowIndexByRefCode_(sheet, refCode) : -1;
  if (rowNum === -1) {
    const row = new Array(HEADERS.length).fill("");
    row[COL_TIMESTAMP] = new Date();
    row[COL_NAME] = name;
    row[COL_USERNAME] = xHandle;
    row[COL_EMAIL] = email;
    row[COL_TRACK] = track;
    row[COL_SOURCE] = source;
    row[COL_APP_ID] = appId;
    row[COL_TELEGRAM] = telegram;
    row[COL_POLYMARKET] = polymarket;
    row[COL_REF_CODE] = refCode;
    row[COL_REFERRED_BY] = referredBy;
    row[COL_WHATSAPP] = whatsapp;
    row[COL_APPROVED] = false;
    sheet.appendRow(row);
    return respond_({ status: "confirmed", appId: appId, approved: false, telegramLink: TELEGRAM_LINK });
  }

  sheet.getRange(rowNum, COL_TIMESTAMP + 1).setValue(new Date());
  sheet.getRange(rowNum, COL_NAME + 1).setValue(name);
  sheet.getRange(rowNum, COL_USERNAME + 1).setValue(xHandle);
  sheet.getRange(rowNum, COL_EMAIL + 1).setValue(email);
  sheet.getRange(rowNum, COL_TRACK + 1).setValue(track);
  sheet.getRange(rowNum, COL_SOURCE + 1).setValue(source);
  sheet.getRange(rowNum, COL_APP_ID + 1).setValue(appId);
  sheet.getRange(rowNum, COL_TELEGRAM + 1).setValue(telegram);
  sheet.getRange(rowNum, COL_POLYMARKET + 1).setValue(polymarket);
  sheet.getRange(rowNum, COL_REFERRED_BY + 1).setValue(referredBy);
  sheet.getRange(rowNum, COL_WHATSAPP + 1).setValue(whatsapp);
  const updatedRow = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
  const status = statusForRow_(updatedRow);
  return respond_({ status: "confirmed", appId: appId, approved: status.approved, telegramLink: TELEGRAM_LINK });
}

/**
 * POST {action:"taskProgress", refCode, task, xHandle?, whatsapp?} — sent
 * as soon as a visitor self-checks a checklist task on the apply page,
 * well before they submit their name/email. Creates or updates a row for
 * them (keyed by Referral Code, the one ID they already have) so an admin
 * can start manually verifying tasks without waiting for full submission.
 * Responds with that row's current status so the page can immediately
 * reflect any verification an admin already did.
 */
function handleTaskProgress_(body) {
  const refCode = normalizeCode_(body.refCode);
  const task = (body.task || "").toString().trim();
  const xHandle = (body.xHandle || "").toString().trim();
  const whatsapp = (body.whatsapp || "").toString().trim();
  if (!refCode) return respond_({ status: "ignored" });

  const sheet = getSheet_();
  const rowNum = findRowIndexByRefCode_(sheet, refCode);
  if (rowNum === -1) {
    const row = new Array(HEADERS.length).fill("");
    row[COL_TIMESTAMP] = new Date();
    row[COL_REF_CODE] = refCode;
    if (task === "x" && xHandle) row[COL_USERNAME] = xHandle;
    if (task === "wa" && whatsapp) row[COL_WHATSAPP] = whatsapp;
    row[COL_APPROVED] = false;
    sheet.appendRow(row);
    const newRow = sheet.getRange(sheet.getLastRow(), 1, 1, HEADERS.length).getValues()[0];
    return respond_(statusForRow_(newRow));
  }

  if (task === "x" && xHandle) sheet.getRange(rowNum, COL_USERNAME + 1).setValue(xHandle);
  if (task === "wa" && whatsapp) sheet.getRange(rowNum, COL_WHATSAPP + 1).setValue(whatsapp);
  const row = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
  return respond_(statusForRow_(row));
}

/**
 * POST {action:"referralSignal", refCode, referredBy, xHandle, whatsapp} —
 * sent as soon as a visitor fills in their X handle and WhatsApp number,
 * rather than waiting for them to finish their whole application. Logged
 * to a separate sheet since there's no name/email yet. See the file-level
 * comment on "referral signals" for why this exists.
 */
function handleReferralSignal_(body) {
  const refCode = normalizeCode_(body.refCode);
  let referredBy = normalizeCode_(body.referredBy);
  const xHandle = (body.xHandle || "").toString().trim();
  const whatsapp = (body.whatsapp || "").toString().trim();

  if (referredBy && referredBy === refCode) {
    referredBy = "";
  }
  if (!referredBy || xHandle.length < 2 || whatsapp.replace(/\D/g, "").length < 7) {
    return respond_({ status: "ignored" });
  }

  const sheet = getPendingSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const existingRefCode = (values[i][1] || "").toString().trim().toUpperCase();
    const existingReferredBy = (values[i][2] || "").toString().trim().toUpperCase();
    if (existingRefCode === refCode && existingReferredBy === referredBy) {
      return respond_({ status: "ok" }); // this exact referral pairing is already logged
    }
  }

  sheet.appendRow([new Date(), refCode, referredBy, xHandle, whatsapp]);
  return respond_({ status: "ok" });
}

/**
 * GET ?code=REFCODE — used by the "invite a friend" step to check
 * whether anyone has referred by this code, either by fully submitting
 * an application or by sending a referral signal (see above).
 *
 * GET ?action=leaderboard — used by referral-leaderboard.html to show
 * the top referrers site-wide. See leaderboard_() below.
 *
 * GET ?action=status&refCode=XXXXXX (or &appId=MA-XXXXXXXX) — used by the
 * apply page to poll per-task verification and overall approval. Keyed by
 * Referral Code so it works from the moment the visitor lands on the page,
 * long before an Application ID exists; appId is accepted too for
 * resuming a session where only the ID is known. See the "Manual review,
 * per checklist step" note in the file-level comment above.
 */
function doGet(e) {
  const params = (e && e.parameter) || {};

  if ((params.action || "").toString().trim() === "leaderboard") {
    return respond_(leaderboard_());
  }

  if ((params.action || "").toString().trim() === "status") {
    const refCode = normalizeCode_(params.refCode);
    const appId = (params.appId || "").toString().trim();
    return respond_(statusForKey_(refCode, appId));
  }

  const code = normalizeCode_(params.code);
  if (!code) {
    return respond_({ used: false, count: 0 });
  }

  const count = countReferralUses_(code);
  return respond_({ used: count > 0, count: count });
}

/**
 * Builds the site-wide referral leaderboard: every referral code that has
 * referred at least one person (via a full application or a referral
 * signal — same two sources doGet's per-code lookup counts), ranked by
 * how many people it referred.
 *
 * A code's owner is only knowable once *they* submit their own full
 * application (their Name/Username land in the Applications row that
 * carries their Referral Code) — referral signals only carry the
 * *referred* visitor's details, not the referrer's. Until then the entry
 * is shown anonymously so someone who's referred people mid-application
 * still shows up on the board.
 */
function leaderboard_() {
  const appValues = getSheet_().getDataRange().getValues();

  // Referral Code -> { name, xHandle } for codes whose owner has submitted.
  const owners = {};
  for (let i = 1; i < appValues.length; i++) {
    const ownCode = (appValues[i][9] || "").toString().trim().toUpperCase();
    if (!ownCode) continue;
    owners[ownCode] = {
      name: (appValues[i][1] || "").toString().trim(),
      xHandle: (appValues[i][2] || "").toString().trim()
    };
  }

  // Referred By -> count, tallied across full applications and signals.
  const counts = {};
  for (let i = 1; i < appValues.length; i++) {
    const referredBy = (appValues[i][10] || "").toString().trim().toUpperCase();
    if (referredBy) counts[referredBy] = (counts[referredBy] || 0) + 1;
  }
  const pendingValues = getPendingSheet_().getDataRange().getValues();
  for (let i = 1; i < pendingValues.length; i++) {
    const referredBy = (pendingValues[i][2] || "").toString().trim().toUpperCase();
    if (referredBy) counts[referredBy] = (counts[referredBy] || 0) + 1;
  }

  const rows = Object.keys(counts).map(function (code) {
    const owner = owners[code];
    return {
      refCode: code,
      name: owner ? owner.name : "",
      xHandle: owner ? owner.xHandle : "",
      count: counts[code]
    };
  });
  rows.sort(function (a, b) { return b.count - a.count; });

  const LIMIT = 50;
  return {
    updatedAt: new Date().toISOString(),
    totalReferrals: rows.reduce(function (sum, r) { return sum + r.count; }, 0),
    leaderboard: rows.slice(0, LIMIT)
  };
}

/**
 * Looks up a single application row by Referral Code (checked first) or
 * Application ID, and reports its per-task verification + overall
 * approval. { found: false } if neither key matches any row.
 */
function statusForKey_(refCode, appId) {
  const values = getSheet_().getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowRef = (values[i][COL_REF_CODE] || "").toString().trim().toUpperCase();
    const rowAppId = (values[i][COL_APP_ID] || "").toString().trim();
    if ((refCode && rowRef === refCode) || (appId && rowAppId === appId)) {
      return statusForRow_(values[i]);
    }
  }
  return { found: false, approved: false, xVerified: false, waVerified: false, pmVerified: false, refVerified: false };
}

/**
 * Derives the status object the apply page polls for from one sheet row.
 * Referral Verified is true if the admin typed TRUE OR a real referral
 * use is already on record (see countReferralUses_) — it doesn't need
 * both. Approved is true if the admin typed TRUE directly OR all four
 * task columns read true — admins normally never touch Approved at all;
 * it just falls out of the per-task columns.
 */
function statusForRow_(row) {
  const xVerified = parseApproved_(row[COL_X_VERIFIED]);
  const waVerified = parseApproved_(row[COL_WA_VERIFIED]);
  const pmVerified = parseApproved_(row[COL_PM_VERIFIED]);
  const refCode = (row[COL_REF_CODE] || "").toString().trim().toUpperCase();
  const refVerified = parseApproved_(row[COL_REF_VERIFIED]) || countReferralUses_(refCode) > 0;
  const approved = parseApproved_(row[COL_APPROVED]) || (xVerified && waVerified && pmVerified && refVerified);
  return {
    found: true,
    appId: row[COL_APP_ID] || null,
    approved: approved,
    xVerified: xVerified,
    waVerified: waVerified,
    pmVerified: pmVerified,
    refVerified: refVerified
  };
}

/** 1-based sheet row number for a Referral Code, or -1 if no row has it. */
function findRowIndexByRefCode_(sheet, refCode) {
  if (!refCode) return -1;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if ((values[i][COL_REF_CODE] || "").toString().trim().toUpperCase() === refCode) return i + 1;
  }
  return -1;
}

/** How many submissions/signals carry this code as their Referred By. */
function countReferralUses_(code) {
  if (!code) return 0;
  let count = 0;
  const appValues = getSheet_().getDataRange().getValues();
  for (let i = 1; i < appValues.length; i++) {
    const referredBy = (appValues[i][COL_REFERRED_BY] || "").toString().trim().toUpperCase();
    if (referredBy && referredBy === code) count++;
  }
  const pendingValues = getPendingSheet_().getDataRange().getValues();
  for (let i = 1; i < pendingValues.length; i++) {
    const referredBy = (pendingValues[i][2] || "").toString().trim().toUpperCase();
    if (referredBy && referredBy === code) count++;
  }
  return count;
}

/**
 * A cell is treated as "approved" only for an explicit true-ish value —
 * Sheets stores a typed TRUE as a real boolean, but admins sometimes type
 * "true"/"yes" as plain text instead. Blank, FALSE, or anything else means
 * still pending.
 */
function parseApproved_(v) {
  if (v === true) return true;
  const s = (v || "").toString().trim().toLowerCase();
  return s === "true" || s === "yes" || s === "1";
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
  ensureHeaders_(sheet, HEADERS);
  return sheet;
}

function getPendingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PENDING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PENDING_SHEET_NAME);
  }
  ensureHeaders_(sheet, PENDING_HEADERS);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < headers.length) {
    sheet.getRange(1, lastCol + 1, 1, headers.length - lastCol).setValues([headers.slice(lastCol)]);
  }
}

function findByEmail_(sheet, email) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowEmail = (values[i][COL_EMAIL] || "").toString().trim().toLowerCase();
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
