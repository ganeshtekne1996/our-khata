/**
 * Our Khata onboarding and ownership helpers.
 * Merge these functions into the Apps Script deployment used by the app.
 * Set GOOGLE_CLIENT_ID to the same web client ID used in index.html.
 */
const GOOGLE_CLIENT_ID = '938761019761-ur365n4lqqav25vs2eb38clqajvhb0f3.apps.googleusercontent.com';
const USER_ONBOARD_SPREADSHEET_ID = '1DjmVesr0grY9xDJwTRqVHNVvX4Fizn2EwLHPx2Pj0RI';
const USER_ONBOARD_SHEET = 'userOnBoard';
const MAX_BOOK_NAME_LENGTH = 100;

function getUserOnBoardSheet_() {
  const spreadsheet = SpreadsheetApp.openById(USER_ONBOARD_SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(USER_ONBOARD_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(USER_ONBOARD_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([[
      'email', 'name', 'workbookName', 'scriptUrl', 'createdAt'
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
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(header) {
      return String(header).trim().toLowerCase().replace(/[\s_-]+/g, '');
    });
  const columns = {};
  headers.forEach(function(header, index) { columns[header] = index; });
  ['email', 'name', 'workbookname', 'scripturl', 'createdat'].forEach(function(required) {
    if (columns[required] === undefined) throw new Error('Missing userOnBoard column: ' + required);
  });
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

function findUserOnBoard_(email) {
  const sheet = getUserOnBoardSheet_();
  const columns = getUserOnBoardColumns_();
  const values = sheet.getDataRange().getValues();
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][columns.email]).toLowerCase() === email) {
      return {
        row: row + 1,
        email: email,
        name: String(values[row][columns.name] || '').trim(),
        workbookName: String(values[row][columns.workbookname] || '').trim(),
        scriptUrl: String(values[row][columns.scripturl] || '').trim()
      };
    }
  }
  return null;
}

function onboardUser_(credential, displayName) {
  const profile = verifyGoogleCredential_(credential);
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
}

function requireUserBook_(credential, requestedBook) {
  const profile = verifyGoogleCredential_(credential);
  const user = findUserOnBoard_(profile.email);
  if (!user) throw new Error('Complete onboarding first');
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
  getUserOnBoardSheet_().getRange(access.user.row, 3).setValue(cleanName);
  return cleanName;
}

function accessWorkbookNameForUser_(credential) {
  const profile = verifyGoogleCredential_(credential);
  const user = findUserOnBoard_(profile.email);
  if (!user) throw new Error('Complete onboarding first');
  return user.workbookName;
}

/**
 * Call this at the beginning of the existing doGet(e):
 *   const access = requireUserBook_(e.parameter.credential, e.parameter.book);
 * Then use access.user.workbookName instead of trusting e.parameter.book.
 */
function handleOnboardAction_(payload) {
  const result = onboardUser_(payload.credential, payload.name);
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
    books: configured ? [result.user.workbookName] : []
  };
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

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
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
