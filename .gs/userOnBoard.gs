/**
 * Our Khata onboarding and ownership helpers.
 * Merge these functions into the Apps Script deployment used by the app.
 * Set GOOGLE_CLIENT_ID to the same web client ID used in index.html.
 */
const USER_ONBOARD_SPREADSHEET_ID = '1KqL4d9KmoTNWAoJE5kaJBryGaSjIfgpr6ZRRjQmr50Y';
const USER_ONBOARD_SHEET = 'userOnBoard';
const MAX_BOOK_NAME_LENGTH = 100;
// NOTE: GmailApp.sendEmail always sends from whichever Google account owns
// and is authorized to run this deployment (Deploy > Manage deployments >
// Execute as: Me). This constant does NOT control the sender — it exists
// only as a reminder of which account that must be. If OTP emails aren't
// arriving, first confirm this project is deployed/authorized under this
// exact account, not just that this constant is set correctly.
const OTP_SENDER_EMAIL = 'ganesh.tekne101@gmail.com';
const OTP_SENDER_NAME = 'Our Khata';
const OTP_RESEND_COOLDOWN_SECONDS = 60;
const SESSION_TOKEN_LENGTH = 32;

function getUserOnBoardSheet_() {
  const spreadsheet = SpreadsheetApp.openById(USER_ONBOARD_SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(USER_ONBOARD_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(USER_ONBOARD_SHEET);
    sheet.getRange(1, 1, 1, 15).setValues([[
      'email', 'name', 'lastName', 'mobile', 'workbookName', 'scriptUrl', 'SHEET_ID', 'createdAt',
      'created_OTP', 'userEntered_OTP', 'otpExpiresAt', 'onboardingVerified', 'otpAttempts', 'sessionToken', 'sessionTokenCreatedAt'
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
  const managedHeaders = ['lastName', 'mobile', 'created_OTP', 'userEntered_OTP', 'otpExpiresAt', 'onboardingVerified', 'otpAttempts', 'sessionToken', 'sessionTokenCreatedAt'];
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
  ['email', 'name', 'workbookname', 'scripturl', 'createdat', 'createdotp', 'userenteredotp', 'otpexpiresat', 'onboardingverified', 'otpattempts', 'sessiontoken'].forEach(function(required) {
    if (columns[required] === undefined) throw new Error('Missing userOnBoard column: ' + required);
  });
  // workbookurl remains optional; SHEET_ID can hold the raw spreadsheet ID
  // to pin each user's data to a specific spreadsheet.
  // Until one is added, onboarding still works exactly as before.
  return columns;
}

function validateRegistration_(payload) {
  const email = String(payload.email || '').trim().toLowerCase();
  const firstName = String(payload.name || '').trim();
  const lastName = String(payload.lastName || '').trim();
  const mobile = String(payload.mobile || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address');
  if (!firstName) throw new Error('First name is required');
  if (!lastName) throw new Error('Last name is required');
  if (mobile && !/^\+?[0-9 ()-]{7,20}$/.test(mobile)) throw new Error('Enter a valid mobile number');
  return { email: email, name: firstName, lastName: lastName, mobile: mobile };
}

function createSessionToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').substring(0, SESSION_TOKEN_LENGTH);
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
      lastName: String(values[row][columns.lastname] || '').trim(),
      mobile: String(values[row][columns.mobile] || '').trim(),
      workbookName: String(values[row][columns.workbookname] || '').trim(),
      scriptUrl: String(values[row][columns.scripturl] || '').trim(),
      workbookUrl: rawWorkbookUrl,
      spreadsheetId: parseSpreadsheetId_(rawSheetId) || parseSpreadsheetId_(rawWorkbookUrl),
      createdOtp: String(values[row][columns.createdotp] || '').trim(),
      userEnteredOtp: String(values[row][columns.userenteredotp] || '').trim(),
      otpExpiresAt: values[row][columns.otpexpiresat],
      onboardingVerified: String(values[row][columns.onboardingverified] || '').toLowerCase() === 'true',
      otpAttempts: parseInt(values[row][columns.otpattempts], 10) || 0
      ,sessionToken: String(values[row][columns.sessiontoken] || '').trim()
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

function findUserBySessionToken_(sessionToken) {
  const token = String(sessionToken || '').trim();
  if (!token) return null;
  const sheet = getUserOnBoardSheet_();
  const columns = getUserOnBoardColumns_();
  const values = sheet.getDataRange().getValues();
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][columns.sessiontoken] || '').trim() !== token) continue;
    const user = findUserOnBoard_(String(values[row][columns.email] || '').trim());
    if (user && user.onboardingVerified) return user;
  }
  return null;
}

function onboardUser_(payload) {
  const profile = validateRegistration_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // avoid two near-simultaneous first-logins creating duplicate rows
  try {
    const sheet = getUserOnBoardSheet_();
    const columns = getUserOnBoardColumns_();
    let user = findUserOnBoard_(profile.email);
    if (!user) {
      const row = new Array(sheet.getLastColumn()).fill('');
      row[columns.email] = profile.email;
      row[columns.name] = profile.name;
      row[columns.lastname] = profile.lastName;
      row[columns.mobile] = profile.mobile;
      row[columns.workbookname] = workbookNameForEmail_(profile.email);
      row[columns.createdat] = new Date();
      sheet.appendRow(row);
      user = findUserOnBoard_(profile.email);
    } else {
      sheet.getRange(user.row, columns.name + 1).setValue(profile.name);
      sheet.getRange(user.row, columns.lastname + 1).setValue(profile.lastName);
      sheet.getRange(user.row, columns.mobile + 1).setValue(profile.mobile);
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

function requireUserBook_(sessionToken, requestedBook) {
  const user = findUserBySessionToken_(sessionToken);
  if (!user) throw new Error('Complete onboarding first');
  if (!user.onboardingVerified) throw new Error('Verify your email before accessing the app');
  if (requestedBook && requestedBook !== user.workbookName) {
    throw new Error('You do not have access to this book');
  }
  return { profile: { email: user.email, name: user.name }, user: user };
}

function renameUserBook_(sessionToken, newWorkbookName) {
  const access = requireUserBook_(sessionToken, accessWorkbookNameForUser_(sessionToken));
  const cleanName = String(newWorkbookName || '').trim().substring(0, MAX_BOOK_NAME_LENGTH);
  if (!cleanName) throw new Error('Book name is required');
  const spreadsheet = SpreadsheetApp.openById(USER_ONBOARD_SPREADSHEET_ID);
  if (spreadsheet.getSheetByName(cleanName)) throw new Error('Book name is already in use');
  spreadsheet.getSheetByName(access.user.workbookName).setName(cleanName);
  const columns = getUserOnBoardColumns_();
  getUserOnBoardSheet_().getRange(access.user.row, columns.workbookname + 1).setValue(cleanName);
  return cleanName;
}

function accessWorkbookNameForUser_(sessionToken) {
  const user = findUserBySessionToken_(sessionToken);
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
  if (force && currentExpiry && currentExpiry.getTime() > now.getTime()) {
    throw new Error('You can request a new OTP after the current OTP expires.');
  }
  const cooldownKey = 'otp-sent-' + user.email;
  const cache = CacheService.getScriptCache();
  if (force && cache.get(cooldownKey)) {
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
  // Gmail sends from the account that owns the web-app deployment. Do not
  // call Session.getEffectiveUser(): it requires the userinfo.email scope.
  // Deploy this web app to execute as OTP_SENDER_EMAIL.
  const mailOptions = {
    name: OTP_SENDER_NAME,
    htmlBody: htmlBody
  };
  try {
    GmailApp.sendEmail(user.email, subject, body, mailOptions);
  } catch (err) {
    // Surface the real reason in the Apps Script Executions log (Extensions
    // > Apps Script > Executions) instead of a generic silent failure, and
    // fail the request loudly so the client shows an error instead of
    // pretending an OTP was sent.
    Logger.log('sendOtp_ failed for ' + user.email + ': ' + err.message);
    throw new Error('Could not send the OTP email. Ask the app owner to check the Apps Script Executions log and Gmail authorization for this deployment. (' + err.message + ')');
  }
  if (force) cache.put(cooldownKey, '1', OTP_RESEND_COOLDOWN_SECONDS);
}

function verifyOtp_(email, submittedOtp) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getUserOnBoardSheet_();
    const columns = getUserOnBoardColumns_();
    const user = findUserOnBoard_(String(email || '').trim().toLowerCase());
    if (!user) throw new Error('Complete onboarding first');
    const otp = String(submittedOtp || '').trim();
    sheet.getRange(user.row, columns.userenteredotp + 1).setNumberFormat('@').setValue(otp);
    if (user.onboardingVerified) {
      if (!user.sessionToken) {
        const token = createSessionToken_();
        sheet.getRange(user.row, columns.sessiontoken + 1).setNumberFormat('@').setValue(token);
        sheet.getRange(user.row, columns.sessiontokencreatedat + 1).setValue(new Date());
        user.sessionToken = token;
      }
      return { profile: { email: user.email, name: user.name }, user: user };
    }
    if (user.otpAttempts >= 5) throw new Error('Too many incorrect attempts. Request a new OTP.');
    const expiry = user.otpExpiresAt instanceof Date ? user.otpExpiresAt : null;
    if (!expiry || expiry.getTime() < Date.now()) throw new Error('This OTP has expired. Request a new OTP.');
    sheet.getRange(user.row, columns.otpattempts + 1).setValue(user.otpAttempts + 1);
    if (!/^\d{4}$/.test(otp) || otp !== user.createdOtp) throw new Error('Incorrect OTP');
    sheet.getRange(user.row, columns.onboardingverified + 1).setValue(true);
    const sessionToken = createSessionToken_();
    sheet.getRange(user.row, columns.sessiontoken + 1).setNumberFormat('@').setValue(sessionToken);
    sheet.getRange(user.row, columns.sessiontokencreatedat + 1).setValue(new Date());
    user.onboardingVerified = true;
    user.sessionToken = sessionToken;
    return { profile: { email: user.email, name: user.name }, user: user };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Call this at the beginning of the existing doGet(e):
 *   const access = requireUserBook_(e.parameter.sessionToken, e.parameter.book);
 * Then use access.user.workbookName instead of trusting e.parameter.book.
 */
function handleOnboardAction_(payload) {
  const result = onboardUser_(payload);
  if (!result.user.onboardingVerified) {
    sendOtp_(result.user, false);
    return { ok: true, email: result.profile.email, requiresOtp: true };
  }
  if (!result.user.sessionToken) {
    const columns = getUserOnBoardColumns_();
    const token = createSessionToken_();
    getUserOnBoardSheet_().getRange(result.user.row, columns.sessiontoken + 1).setNumberFormat('@').setValue(token);
    result.user.sessionToken = token;
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
    sessionToken: result.user.sessionToken || '',
    workbookUrl: result.user.workbookUrl || '',
    books: configured ? [result.user.workbookName] : []
  };
}

function handleVerifyOtpAction_(payload) {
  const result = verifyOtp_(payload.email, payload.otp);
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
    sessionToken: result.user.sessionToken || '',
    workbookUrl: result.user.workbookUrl || '',
    books: configured ? [result.user.workbookName] : []
  };
}

function handleResendOtpAction_(payload) {
  const user = findUserOnBoard_(String(payload.email || '').trim().toLowerCase());
  if (!user) throw new Error('Complete onboarding first');
  if (user.onboardingVerified) return { ok: true, verified: true, sessionToken: user.sessionToken || '' };
  sendOtp_(user, true);
  return { ok: true, requiresOtp: true };
}

/**
 * Call this at the beginning of the existing doPost(e):
 *   const payload = JSON.parse(e.postData.contents);
 *   if (payload.action === 'onboard') return json_(handleOnboardAction_(payload));
 *   const access = requireUserBook_(payload.sessionToken, payload.book);
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