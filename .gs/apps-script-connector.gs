// Legacy default — kept only as a fallback for requests that don't send an
// explicit spreadsheetId. New requests should always pass spreadsheetId so
// this script writes to the correct per-user workbook regardless of what
// this constant happens to be set to in any given copy of this script.
const SHEET_ID = '1IbFAET-Cixodx6sLjEPzftjVfKSK8Upr0bfFrkJTJpc';
const ENTRY_HEADERS = ['id', 'date', 'type', 'amount', 'note', 'paymentMode', 'addedBy', 'createdAt'];

// Set once per request (from doGet/doPost) to the spreadsheetId the client
// sent, if any. getSS() prefers this over the hardcoded SHEET_ID.
let ACTIVE_SPREADSHEET_ID_ = null;

function setActiveSpreadsheetId_(id) {
  const clean = String(id || '').trim();
  ACTIVE_SPREADSHEET_ID_ = clean || null;
}

function getSS() {
  const id = ACTIVE_SPREADSHEET_ID_ || SHEET_ID;
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    // Surface a clear error instead of a cryptic Apps Script permission
    // stack trace — most likely cause is either a bad/mistyped
    // spreadsheetId, or this script's executing account doesn't have
    // edit access to that spreadsheet yet.
    throw new Error('Could not open spreadsheet ' + id + ': ' + err.message);
  }
}

function getTZ() {
  return getSS().getSpreadsheetTimeZone();
}

function getSettingsSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName('Settings');
  if (!sh) {
    sh = ss.insertSheet('Settings');
    sh.appendRow(['you', 'friend']);
    sh.appendRow(['You', 'Friend']);
  }
  return sh;
}

function sanitizeBookName(name) {
  let n = (name || '').toString().trim();
  n = n.replace(/[\[\]\*\?:\/\\]/g, '-');
  if (n.length > 95) n = n.substring(0, 95);
  if (!n) n = 'Khatabook';
  return n;
}

function isBookSheet(sh) {
  if (sh.getName() === 'Settings') return false;
  if (sh.getLastRow() === 0) return false;
  const a1 = sh.getRange(1, 1).getValue();
  return a1 === 'id';
}

function migrateBookSheetIfNeeded(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return;
  const header = sh.getRange(1, 1, 1, Math.min(lastCol, 8)).getValues()[0];
  if (header[0] !== 'id') return;

  let matches = header.length === ENTRY_HEADERS.length;
  if (matches) {
    for (let i = 0; i < ENTRY_HEADERS.length; i++) {
      if (header[i] !== ENTRY_HEADERS[i]) { matches = false; break; }
    }
  }
  if (matches) return;

  const isLegacySeven = header[0] === 'id' && header[1] === 'date' && header[2] === 'type' &&
    header[3] === 'amount' && header[4] === 'note' && header[5] === 'addedBy' && header[6] === 'createdAt';
  if (isLegacySeven) {
    sh.insertColumnBefore(6);
    sh.getRange(1, 1, 1, ENTRY_HEADERS.length).setValues([ENTRY_HEADERS]);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('@');
  }
}

function listBookNames() {
  const ss = getSS();
  return ss.getSheets()
    .filter(function (sh) { migrateBookSheetIfNeeded(sh); return isBookSheet(sh) && !sh.isSheetHidden(); })
    .map(function (sh) { return sh.getName(); });
}

function listArchivedBookNames() {
  const ss = getSS();
  return ss.getSheets()
    .filter(function (sh) { return isBookSheet(sh) && sh.isSheetHidden(); })
    .map(function (sh) { return sh.getName(); });
}

function getOrCreateBookSheet(name) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(ENTRY_HEADERS);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('@');
  } else {
    migrateBookSheetIfNeeded(sh);
    if (sh.isSheetHidden()) sh.showSheet();
  }
  return sh;
}

function ensureAtLeastOneBook() {
  const books = listBookNames();
  if (books.length === 0) {
    getOrCreateBookSheet('General');
    return 'General';
  }
  return books[0];
}

// Writes a single text cell in a way that reliably survives Google Sheets'
// auto date/number detection: format is forced to Plain Text on the exact
// cell immediately before the value is written.
function writeTextCell(sh, row, col, textValue) {
  const range = sh.getRange(row, col);
  range.setNumberFormat('@');
  range.setValue(textValue);
}

// Defensive normalizer: if a legacy row's date/createdAt cell was still
// auto-converted to a real Date type before this fix, convert it back to
// the DD-MM-YYYY (or plain timestamp) string the app expects.
function normalizeEntry(obj, tz) {
  if (Object.prototype.toString.call(obj.date) === '[object Date]') {
    obj.date = Utilities.formatDate(obj.date, tz, 'dd-MM-yyyy');
  }
  if (Object.prototype.toString.call(obj.createdAt) === '[object Date]') {
    obj.createdAt = String(obj.createdAt.getTime());
  } else {
    obj.createdAt = String(obj.createdAt);
  }
  return obj;
}

function readBookEntries(sh) {
  migrateBookSheetIfNeeded(sh);
  const tz = getTZ();
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  return data
    .filter(function (r) { return r[0] !== ''; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return normalizeEntry(obj, tz);
    });
}

function entriesDoGet_(e) {
  setActiveSpreadsheetId_(e && e.parameter && e.parameter.spreadsheetId);
  try {
    const access = requireUserBook_(e && e.parameter && e.parameter.sessionToken, e && e.parameter && e.parameter.book);
    setActiveSpreadsheetId_(access.user.spreadsheetId);
    const settingsSh = getSettingsSheet();
    const sData = settingsSh.getDataRange().getValues();
    const settings = { you: sData[1][0] || 'You', friend: sData[1][1] || 'Friend' };

    let books = listBookNames();
    if (books.length === 0) {
      ensureAtLeastOneBook();
      books = listBookNames();
    }
    const archivedBooks = listArchivedBookNames();

    let activeBook = (e && e.parameter && e.parameter.book) ? e.parameter.book : books[0];
    if (books.indexOf(activeBook) === -1) activeBook = books[0];

    const sh = getOrCreateBookSheet(activeBook);
    const entries = readBookEntries(sh);

    return jsonOut({
      books: books,
      archivedBooks: archivedBooks,
      activeBook: activeBook,
      entries: entries,
      settings: settings
    });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function entriesDoPost_(e) {
  const body = JSON.parse(e.postData.contents);
  setActiveSpreadsheetId_(body.spreadsheetId);
  try {
    const access = requireUserBook_(body.sessionToken, body.book);
    setActiveSpreadsheetId_(access.user.spreadsheetId);
    if (body.action === 'add') {
      const sh = getOrCreateBookSheet(body.book);
      const row = sh.getLastRow() + 1;
      writeTextCell(sh, row, 1, body.id);
      writeTextCell(sh, row, 2, body.date);          // DD-MM-YYYY text, forced
      sh.getRange(row, 3).setValue(body.type);
      sh.getRange(row, 4).setValue(body.amount);
      sh.getRange(row, 5).setValue(body.note || '');
      sh.getRange(row, 6).setValue(body.paymentMode || '');
      sh.getRange(row, 7).setValue(body.addedBy || '');
      writeTextCell(sh, row, 8, String(body.createdAt)); // forced text
    } else if (body.action === 'update') {
      const sh = getOrCreateBookSheet(body.book);
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          const row = i + 1;
          writeTextCell(sh, row, 2, body.date);        // DD-MM-YYYY text, forced
          sh.getRange(row, 3).setValue(body.type);
          sh.getRange(row, 4).setValue(body.amount);
          sh.getRange(row, 5).setValue(body.note || '');
          sh.getRange(row, 6).setValue(body.paymentMode || '');
          sh.getRange(row, 7).setValue(body.addedBy || '');
          break;
        }
      }
    } else if (body.action === 'delete') {
      const sh = getOrCreateBookSheet(body.book);
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) { sh.deleteRow(i + 1); break; }
      }
    } else if (body.action === 'settings') {
      const sh = getSettingsSheet();
      sh.getRange(2, 1, 1, 2).setValues([[body.you, body.friend]]);
    } else if (body.action === 'createBook') {
      const name = sanitizeBookName(body.name);
      const ss = getSS();
      if (ss.getSheetByName(name)) throw new Error('A khatabook named "' + name + '" already exists');
      getOrCreateBookSheet(name);
      return jsonOut({ ok: true, book: name });
    } else if (body.action === 'renameBook') {
      const ss = getSS();
      const sh = ss.getSheetByName(body.oldName);
      if (!sh) throw new Error('Khatabook not found');
      const newName = sanitizeBookName(body.newName);
      if (newName !== body.oldName && ss.getSheetByName(newName)) {
        throw new Error('A khatabook named "' + newName + '" already exists');
      }
      sh.setName(newName);
      return jsonOut({ ok: true, book: newName });
    } else if (body.action === 'deleteBook') {
      const ss = getSS();
      const sh = ss.getSheetByName(body.name);
      if (!sh) throw new Error('Khatabook not found');
      const visibleCount = ss.getSheets().filter(function (s) { return isBookSheet(s) && !s.isSheetHidden(); }).length;
      if (visibleCount <= 1) throw new Error('Create another khatabook before deleting the last one');
      sh.hideSheet();
      return jsonOut({ ok: true });
    } else if (body.action === 'restoreBook') {
      const ss = getSS();
      const sh = ss.getSheetByName(body.name);
      if (!sh) throw new Error('Khatabook not found');
      sh.showSheet();
      return jsonOut({ ok: true, book: body.name });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
  return jsonOut({ ok: true });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}