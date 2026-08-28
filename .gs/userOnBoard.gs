/**
 * Our Khata onboarding and ownership helpers.
 * Merge these functions into the Apps Script deployment used by the app.
 * Set GOOGLE_CLIENT_ID to the same web client ID used in index.html.
 */
const GOOGLE_CLIENT_ID = '938761019761-ur365n4lqqav25vs2eb38clqajvhb0f3.apps.googleusercontent.com';
const USER_ONBOARD_SPREADSHEET_ID = '1DjmVesr0grY9xDJwTRqVHNVvX4Fizn2EwLHPx2Pj0RI';
const USER_ONBOARD_SHEET = 'userOnBoard';
const MAX_BOOK_NAME_LENGTH = 100;
const OTP_SENDER_EMAIL = 'ganesh.tekne101@gmail.com';
const OTP_SENDER_NAME = 'Our Khata';

function getUserOnBoardSheet_() {
  const spreadsheet = SpreadsheetApp.openById(USER_ONBOARD_SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(USER_ONBOARD_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(USER_ONBOARD_SHEET);
    sheet.getRange(1, 1, 1, 11).setValues([[
      'email', 'name', 'workbookName', 'scriptUrl', 'SHEET_ID', 'createdAt',
      'created_OTP', 'userEntered_OTP', 'otpExpiresAt', 'onboardingVerified', 'otpAttempts'
    ]]);
    sheet.setFrozenRows(1);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let column = headers.length - 1; column >= 0; column--) {
    const header = String(headers[column]).trim().toLowerCase().replace(/[\s_-]+/g, '');
    if (header === 'workbooknamelink') sheet.deleteColumn(column + 1);
  }
  return sheet;
}

function getUserOnBoardColumns_() {
  const sheet = getUserOnBoardSheet_();
  let rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const normalizedHeaders = rawHeaders.map(function(header) {
    return String(header).trim().toLowerCase().replace(/[\s_-]+/g, '');
  });
  const scriptUrlIndex = normalizedHeaders.indexOf('scripturl');
  if (scriptUrlIndex === -1) throw new Error('Missing userOnBoard column: scripturl');
  const sheetIdIndex = normalizedHeaders.indexOf('sheetid');
  if (sheetIdIndex === -1) {
    sheet.insertColumnAfter(scriptUrlIndex + 1);
    sheet.getRange(1, scriptUrlIndex + 2).setValue('SHEET_ID');
    rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } else if (sheetIdIndex !== scriptUrlIndex + 1) {
    sheet.moveColumns(sheet.getRange(1, sheetIdIndex + 1, sheet.getMaxRows(), 1), scriptUrlIndex + 2);
    rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  const managedHeaders = ['created_OTP', 'userEntered_OTP', 'otpExpiresAt', 'onboardingVerified', 'otpAttempts'];
  managedHeaders.forEach(function(requiredHeader) {
    const normalized = requiredHeader.toLowerCase().replace(/[\s_-]+/g, '');
    const exists = rawHeaders.some(function(header) {
      return String(header).trim().toLowerCase().replace(/[\s_-]+/g, '') === normalized;
    });
    if (!exists) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(requiredHeader);
      rawHeaders.push(requiredHeader);
    }
  });
  const headers = rawHeaders
    .map(function(header) {
      return String(header).trim().toLowerCase().replace(/[\s_-]+/g, '');
    });
  const columns = {};
  headers.forEach(function(header, index) { columns[header] = index; });
  ['email', 'name', 'workbookname', 'scripturl', 'createdat', 'createdotp', 'userenteredotp', 'otpexpiresat', 'onboardingverified', 'otpattempts'].forEach(function(required) {
    if (columns[required] === undefined) throw new Error('Missing userOnBoard column: ' + required);
  });
  // workbookurl remains optional; SHEET_ID can hold the raw spreadsheet ID
  // to pin each user's data to a specific spreadsheet.
  // Until one is added, onboarding still works exactly as before.
  return columns;
}

function verifyGoogleCredential_(credential) {
  if (!credential) throw new Error('Sign-in is required');
  const response = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential),
    { muteHttpExceptions: true }
  );
  if (response.getResponseCode() !== 200) throw new Error('Invalid Google sign-in');
  const profile = JSON.parse(response.getContentText());
  const email = String(profile.email || '').toLowerCase();
  if (!email.endsWith('@gmail.com') || profile.email_verified !== 'true') {
    throw new Error('Only verified Gmail accounts are supported');
  }
  if (GOOGLE_CLIENT_ID.indexOf('REPLACE_') !== 0 && profile.aud !== GOOGLE_CLIENT_ID) {
    throw new Error('Google client ID does not match');
  }
  return { email: email, name: String(profile.name || email.split('@')[0]) };
}

function workbookNameForEmail_(email) {
  return email.substring(0, email.indexOf('@')).substring(0, MAX_BOOK_NAME_LENGTH);
}

// Accepts a full Google Sheets URL (e.g. .../spreadsheets/d/<ID>/edit?gid=0)
// or a bare ID, and returns just the spreadsheet ID. Returns '' if unparseable.
function parseSpreadsheetId_(urlOrId) {
  const value = String(urlOrId || '').trim();
  if (!value) return '';
  const match = value.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  // Looks like a bare ID already (no slashes, reasonable length) — use as-is.
  if (/^[a-zA-Z0-9-_]{20,}$/.test(value)) return value;
  return '';
}

function findUserOnBoard_(email) {
  const sheet = getUserOnBoardSheet_();
  const columns = getUserOnBoardColumns_();
  const values = sheet.getDataRange().getValues();
  const cleanEmail = String(email || '').trim().toLowerCase();
  let match = null;
  for (let row = 1; row < values.length; row++) {
    const rowEmail = String(values[row][columns.email] || '').trim().toLowerCase();
    if (rowEmail !== cleanEmail) continue;
    // Prefer a raw SHEET_ID column if present (no parsing needed); fall
    // back to workbookUrl (a full Sheets link) otherwise.
    const rawSheetId = columns.sheetid !== undefined
      ? String(values[row][columns.sheetid] || '').trim() : '';
    const rawWorkbookUrl = columns.workbookurl !== undefined
      ? String(values[row][columns.workbookurl] || '').trim() : '';
    const candidate = {
      row: row + 1,
      email: cleanEmail,
      name: String(values[row][columns.name] || '').trim(),
      workbookName: String(values[row][columns.workbookname] || '').trim(),
      scriptUrl: String(values[row][columns.scripturl] || '').trim(),
      workbookUrl: rawWorkbookUrl,
      spreadsheetId: parseSpreadsheetId_(rawSheetId) || parseSpreadsheetId_(rawWorkbookUrl),
      createdOtp: String(values[row][columns.createdotp] || '').trim(),
      userEnteredOtp: String(values[row][columns.userenteredotp] || '').trim(),
      otpExpiresAt: values[row][columns.otpexpiresat],
      onboardingVerified: String(values[row][columns.onboardingverified] || '').toLowerCase() === 'true',
      otpAttempts: parseInt(values[row][columns.otpattempts], 10) || 0
    };
    // If duplicate rows exist for this email (e.g. from a race before this
    // fix), prefer the one that has actually been configured with a
    // scriptUrl over an earlier blank row, so a manually-added script is
    // never shadowed by a stray duplicate.
    if (!match || (!match.scriptUrl && candidate.scriptUrl)) {
      match = candidate;
    }
  }
  return match;
}

function onboardUser_(credential, displayName) {
  const profile = verifyGoogleCredential_(credential);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // avoid two near-simultaneous first-logins creating duplicate rows
  try {
    const sheet = getUserOnBoardSheet_();
    const columns = getUserOnBoardColumns_();
    let user = findUserOnBoard_(profile.email);
    if (!user) {
      const row = new Array(sheet.getLastColumn()).fill('');
      row[columns.email] = profile.email;
      row[columns.name] = String(displayName || profile.name);
      row[columns.workbookname] = workbookNameForEmail_(profile.email);
      row[columns.createdat] = new Date();
      sheet.appendRow(row);
      user = findUserOnBoard_(profile.email);
    } else {
      sheet.getRange(user.row, columns.name + 1).setValue(String(displayName || profile.name));
      if (!user.workbookName) {
        user.workbookName = workbookNameForEmail_(profile.email);
        sheet.getRange(user.row, columns.workbookname + 1).setValue(user.workbookName);
      }
    }
    return { profile: profile, user: user };
  } finally {
    lock.releaseLock();
  }
}

function requireUserBook_(credential, requestedBook) {
  const profile = verifyGoogleCredential_(credential);
  const user = findUserOnBoard_(profile.email);
  if (!user) throw new Error('Complete onboarding first');
  if (!user.onboardingVerified) throw new Error('Verify your email before accessing the app');
  if (requestedBook && requestedBook !== user.workbookName) {
    throw new Error('You do not have access to this book');
  }
  return { profile: profile, user: user };
}

function renameUserBook_(credential, newWorkbookName) {
  const access = requireUserBook_(credential, accessWorkbookNameForUser_(credential));
  const cleanName = String(newWorkbookName || '').trim().substring(0, MAX_BOOK_NAME_LENGTH);
  if (!cleanName) throw new Error('Book name is required');
  const spreadsheet = SpreadsheetApp.openById(USER_ONBOARD_SPREADSHEET_ID);
  if (spreadsheet.getSheetByName(cleanName)) throw new Error('Book name is already in use');
  spreadsheet.getSheetByName(access.user.workbookName).setName(cleanName);
  const columns = getUserOnBoardColumns_();
  getUserOnBoardSheet_().getRange(access.user.row, columns.workbookname + 1).setValue(cleanName);
  return cleanName;
}

function accessWorkbookNameForUser_(credential) {
  const profile = verifyGoogleCredential_(credential);
  const user = findUserOnBoard_(profile.email);
  if (!user) throw new Error('Complete onboarding first');
  if (!user.onboardingVerified) throw new Error('Verify your email before accessing the app');
  return user.workbookName;
}

function generateOtp_() {
  return String(Math.floor(Math.random() * 10000)).padStart(4, '0');
}

function sendOtp_(user, force) {
  const sheet = getUserOnBoardSheet_();
  const columns = getUserOnBoardColumns_();
  const now = new Date();
  const currentExpiry = user.otpExpiresAt instanceof Date ? user.otpExpiresAt : null;
  if (!force && user.createdOtp && currentExpiry && currentExpiry.getTime() > now.getTime()) return;
  const cooldownKey = 'otp-sent-' + user.email;
  if (force && CacheService.getScriptCache().get(cooldownKey)) {
    throw new Error('Please wait a minute before requesting another OTP.');
  }
  const otp = generateOtp_();
  const expiresAt = new Date(now.getTime() + 10 * 60 * 1000);
  sheet.getRange(user.row, columns.createdotp + 1).setNumberFormat('@').setValue(otp);
  sheet.getRange(user.row, columns.userenteredotp + 1).setNumberFormat('@').setValue('');
  sheet.getRange(user.row, columns.otpexpiresat + 1).setValue(expiresAt);
  sheet.getRange(user.row, columns.otpattempts + 1).setValue(0);
  const subject = 'Your Our Khata verification code';
  const body = [
    'Hello ' + (user.name || 'there') + ',',
    '',
    'Your Our Khata verification code is: ' + otp,
    '',
    'This code expires in 10 minutes. Enter it in the Our Khata app to complete your first-time sign-in.',
    '',
    'If you did not request this code, you can safely ignore this email.',
    '',
    'Regards,',
    'Our Khata'
  ].join('\n');
  const htmlBody = body
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>')
    .replace(otp, '<strong style="font-size:24px;letter-spacing:4px;">' + otp + '</strong>');
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const aliases = GmailApp.getAliases().map(function(alias) { return String(alias).toLowerCase(); });
  const senderIsAlias = aliases.indexOf(OTP_SENDER_EMAIL) !== -1;
  if (effectiveEmail !== OTP_SENDER_EMAIL && !senderIsAlias) {
    throw new Error('Email sender is not configured. Add ' + OTP_SENDER_EMAIL + ' as a Gmail Send mail as address for the Apps Script account.');
  }
  const mailOptions = {
    name: OTP_SENDER_NAME,
    htmlBody: htmlBody
  };
  if (senderIsAlias) mailOptions.from = OTP_SENDER_EMAIL;
  GmailApp.sendEmail(user.email, subject, body, mailOptions);
}

function verifyOtp_(credential, submittedOtp) {
  const profile = verifyGoogleCredential_(credential);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getUserOnBoardSheet_();
    const columns = getUserOnBoardColumns_();
    const user = findUserOnBoard_(profile.email);
    if (!user) throw new Error('Complete onboarding first');
    const otp = String(submittedOtp || '').trim();
    sheet.getRange(user.row, columns.userenteredotp + 1).setNumberFormat('@').setValue(otp);
    if (user.onboardingVerified) return { profile: profile, user: user };
    if (user.otpAttempts >= 5) throw new Error('Too many incorrect attempts. Request a new OTP.');
    const expiry = user.otpExpiresAt instanceof Date ? user.otpExpiresAt : null;
    if (!expiry || expiry.getTime() < Date.now()) throw new Error('This OTP has expired. Request a new OTP.');
    sheet.getRange(user.row, columns.otpattempts + 1).setValue(user.otpAttempts + 1);
    if (!/^\d{4}$/.test(otp) || otp !== user.createdOtp) throw new Error('Incorrect OTP');
    sheet.getRange(user.row, columns.onboardingverified + 1).setValue(true);
    user.onboardingVerified = true;
    return { profile: profile, user: user };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Call this at the beginning of the existing doGet(e):
 *   const access = requireUserBook_(e.parameter.credential, e.parameter.book);
 * Then use access.user.workbookName instead of trusting e.parameter.book.
 */
function handleOnboardAction_(payload) {
  const result = onboardUser_(payload.credential, payload.name);
  if (!result.user.onboardingVerified) {
    sendOtp_(result.user, false);
    return { ok: true, email: result.profile.email, requiresOtp: true };
  }
  const configured = Boolean(
    String(result.user.workbookName || '').trim() &&
    String(result.user.scriptUrl || '').trim()
  );
  return {
    ok: true,
    email: result.profile.email,
    name: result.user.name,
    configured: configured,
    book: result.user.workbookName,
    scriptUrl: result.user.scriptUrl,
    // Optional: only populated once you add a 'SHEET_ID' (raw spreadsheet
    // ID) or 'workbookUrl' (full Sheets link) column and fill it in for
    // that user's row. Pass this through on every fetch/save call so the
    // entries script opens the correct spreadsheet explicitly instead of
    // relying on its own binding.
    spreadsheetId: result.user.spreadsheetId || '',
    workbookUrl: result.user.workbookUrl || '',
    books: configured ? [result.user.workbookName] : []
  };
}

function handleVerifyOtpAction_(payload) {
  const result = verifyOtp_(payload.credential, payload.otp);
  const configured = Boolean(
    String(result.user.workbookName || '').trim() &&
    String(result.user.scriptUrl || '').trim()
  );
  return {
    ok: true,
    verified: true,
    email: result.profile.email,
    name: result.user.name,
    configured: configured,
    book: result.user.workbookName,
    scriptUrl: result.user.scriptUrl,
    spreadsheetId: result.user.spreadsheetId || '',
    workbookUrl: result.user.workbookUrl || '',
    books: configured ? [result.user.workbookName] : []
  };
}

function handleResendOtpAction_(payload) {
  const profile = verifyGoogleCredential_(payload.credential);
  const user = findUserOnBoard_(profile.email);
  if (!user) throw new Error('Complete onboarding first');
  if (user.onboardingVerified) return { ok: true, verified: true };
  sendOtp_(user, true);
  return { ok: true, requiresOtp: true };
}

/**
 * Call this at the beginning of the existing doPost(e):
 *   const payload = JSON.parse(e.postData.contents);
 *   if (payload.action === 'onboard') return json_(handleOnboardAction_(payload));
 *   const access = requireUserBook_(payload.credential, payload.book);
 *   payload.book = access.user.workbookName;
 */
function setupUserOnBoardSheet_() {
  getUserOnBoardSheet_();
}

/**
 * One-time manual cleanup: run this from the Apps Script editor (select the
 * function, click Run) if you suspect duplicate/blank rows for an email are
 * shadowing a manually-configured scriptUrl. For each email, keeps the row
 * with a scriptUrl if one exists (else the most recently created row) and
 * deletes the rest. Logs what it removed so you can double-check before
 * trusting it — review the execution log after running.
 */
function cleanupDuplicateUserOnBoardRows_() {
  const sheet = getUserOnBoardSheet_();
  const columns = getUserOnBoardColumns_();
  const values = sheet.getDataRange().getValues();
  const bestRowForEmail = {}; // email -> {rowIndex (1-based), hasScriptUrl}
  const rowsToDelete = [];

  for (let row = 1; row < values.length; row++) {
    const email = String(values[row][columns.email] || '').trim().toLowerCase();
    if (!email) continue;
    const scriptUrl = String(values[row][columns.scripturl] || '').trim();
    const sheetRow = row + 1;
    const existing = bestRowForEmail[email];
    if (!existing) {
      bestRowForEmail[email] = { sheetRow: sheetRow, hasScriptUrl: !!scriptUrl };
    } else if (!existing.hasScriptUrl && scriptUrl) {
      // Newly found row is better (has scriptUrl, previous kept one didn't) — drop the old one.
      rowsToDelete.push(existing.sheetRow);
      bestRowForEmail[email] = { sheetRow: sheetRow, hasScriptUrl: true };
    } else {
      // Either this row is worse, or both/neither have scriptUrl — drop this duplicate.
      rowsToDelete.push(sheetRow);
    }
  }

  rowsToDelete.sort(function(a, b) { return b - a; }); // delete bottom-up so row numbers stay valid
  rowsToDelete.forEach(function(sheetRow) { sheet.deleteRow(sheetRow); });
  Logger.log('Removed ' + rowsToDelete.length + ' duplicate row(s): ' + JSON.stringify(rowsToDelete));
  return rowsToDelete.length;
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action === 'verifyOtp') return json_(handleVerifyOtpAction_(payload));
    if (payload.action === 'resendOtp') return json_(handleResendOtpAction_(payload));
    if (payload.action !== 'onboard') throw new Error('Unsupported action');
    return json_(handleOnboardAction_(payload));
  } catch (err) {
    return json_({ok: false, error: err.message || 'Request failed'});
  }
}

function doGet() {
  setupUserOnBoardSheet_();
  return json_({ok: true, service: 'userOnBoard'});
}