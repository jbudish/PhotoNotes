/**
 * PhotoNotes / Unit Inspector — Google Apps Script web app.
 *
 * This is the server the phone page posts to. It is NOT deployed from GitHub.
 * Paste it into the script bound to the walk-through Google Sheet.
 *
 * ── Update the live app and KEEP THE SAME URL ──────────────────────────────
 * 1. Open the Google Sheet → Extensions → Apps Script.
 * 2. Select all of the existing code and replace it with this file.
 * 3. Read CONFIG below. Sheet tab names and the Drive folder name in this
 *    repo are best guesses — the previous script was not in the repo. Change
 *    them so rows and photos keep landing in the same places as before.
 * 4. Deploy → Manage deployments → pencil icon on the EXISTING deployment
 *    → Version: New version → Deploy.
 *    Do not choose "New deployment". A new deployment changes the URL and
 *    phones still pointing at the old one will not see this script.
 * 5. Leave access as: Execute as me, Who has access: Anyone.
 *
 * ── First-time install ─────────────────────────────────────────────────────
 * Deploy → New deployment → type Web app → Execute as me → Anyone.
 * Copy the URL into PhotoNotes → Settings.
 * The page calls doGet and expects: {"status":"Unit Inspector script is running"}
 *
 * ── What the page sends ────────────────────────────────────────────────────
 *   getDropdowns      → { success, categories: string[] }
 *   submitRows        → { success, rowsAdded }
 *   uploadPhoto       → { success, fileId, viewUrl, thumbUrl }
 *       unit     = property name
 *       subunit  = unit number
 *   submitFollowUp    → { success }   body.row = { property, location, category, question, date }
 *   submitPhotoRow    → one photo + its row, idempotent on body.id
 *       { id, filename, mimeType, base64, row: { property, unit, inspector,
 *         category, note, photoUrl, location, date } }
 *       Success means the photo (if any) AND the row are saved, or this id
 *       was already saved. viewUrl is the existing link when already saved.
 *
 * Unknown actions return { success:false, error:"Unknown action" } so an
 * older copy of the phone page can fall back to uploadPhoto + submitRows.
 */
var CONFIG = {
  // Walk-through rows. Leave '' to use the first tab that is not one of the
  // helper tabs below (safest when this script is bound to the existing sheet).
  // Set a name only if the rows live on a specific tab, e.g. 'Sheet1'.
  ROWS_SHEET_NAME: '',

  // Created if missing. Point this at the existing follow-up tab if it has
  // another name (Follow Up, Follow-ups, …).
  FOLLOWUP_SHEET_NAME: 'FollowUps',

  // Optional. Column A is the category list for getDropdowns. When the tab
  // is missing, categories are the unique values already in the rows tab.
  CATEGORIES_SHEET_NAME: 'Categories',

  // Hidden tab of client ids so a retry does not create a second file or row.
  IDS_SHEET_NAME: 'ProcessedIds',

  // Property folders are created inside this Drive folder, with a subfolder
  // per unit. Set to '' to create property folders directly in My Drive,
  // which matches a script that never used a parent folder.
  // If your current photos live in a differently named folder, change this
  // BEFORE deploying or new photos will split into a second tree.
  DRIVE_ROOT_FOLDER_NAME: 'PhotoNotes',

  // Row columns, in the same order the phone's submitRows object is built.
  // An existing header row is matched by name (any order). A sheet that
  // already has data and no header is appended in this order and is not
  // rewritten. An empty sheet gets this header row.
  ROW_HEADERS: ['property', 'unit', 'inspector', 'category', 'note', 'photoUrl', 'location', 'date'],
  FOLLOWUP_HEADERS: ['property', 'location', 'category', 'question', 'date'],
};

// ─── Web app entry ───────────────────────────────────────────────────────────

function doGet() {
  return json_({ status: 'Unit Inspector script is running' });
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action;
    if (action === 'getDropdowns') return json_(getDropdowns_());
    if (action === 'submitRows') return json_(submitRows_(body));
    if (action === 'uploadPhoto') return json_(uploadPhoto_(body));
    if (action === 'submitFollowUp') return json_(submitFollowUp_(body));
    if (action === 'submitPhotoRow') return json_(submitPhotoRow_(body));
    return json_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ success: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Actions ─────────────────────────────────────────────────────────────────

function getDropdowns_() {
  var categories = readCategoryList_();
  return { success: true, categories: categories };
}

function submitRows_(body) {
  // Preserve the previous script's error when `rows` is missing.
  var rows = body.rows;
  rows.forEach(function () {});
  return withLock_(function () {
    var id = clientId_(body);
    if (id && rowAlreadySaved_(id)) {
      return { success: true, alreadySaved: true, rowsAdded: 0 };
    }
    var sheet = getRowsSheet_();
    var added = 0;
    rows.forEach(function (row) {
      appendMapped_(sheet, CONFIG.ROW_HEADERS, rowValues_(row));
      added++;
    });
    if (id) {
      var photoUrl = rows.length && rows[0] && rows[0].photoUrl ? String(rows[0].photoUrl) : '';
      rememberId_(id, { viewUrl: photoUrl, fileId: '', rowSaved: true, kind: 'rows' });
    }
    return { success: true, rowsAdded: added };
  });
}

function uploadPhoto_(body) {
  return withLock_(function () {
    var id = clientId_(body);
    if (id) {
      var existing = getIdRecord_(id);
      if (existing && existing.viewUrl) {
        return photoResult_(existing.fileId, existing.viewUrl, true);
      }
    }
    var saved = savePhotoFile_(body);
    if (id) {
      rememberId_(id, {
        viewUrl: saved.viewUrl,
        fileId: saved.fileId,
        rowSaved: false,
        kind: 'photo',
      });
    }
    return photoResult_(saved.fileId, saved.viewUrl, false);
  });
}

function submitFollowUp_(body) {
  // Preserve the previous script's error when `row` is missing.
  var row = body.row;
  var property = row.property;
  return withLock_(function () {
    var id = clientId_(body);
    if (id && rowAlreadySaved_(id)) {
      return { success: true, alreadySaved: true };
    }
    appendMapped_(getFollowupSheet_(), CONFIG.FOLLOWUP_HEADERS, {
      property: property || '',
      location: row.location || '',
      category: row.category || '',
      question: row.question || '',
      date: row.date || '',
    });
    if (id) rememberId_(id, { viewUrl: '', fileId: '', rowSaved: true, kind: 'followUp' });
    return { success: true };
  });
}

/**
 * Save one walk-through item. The phone retries this with the same id until
 * success, including after a timeout where the write actually finished.
 * LockService plus the ProcessedIds tab make that safe.
 */
function submitPhotoRow_(body) {
  return withLock_(function () {
    var id = clientId_(body);
    var existing = id ? getIdRecord_(id) : null;
    if (existing && existing.rowSaved) {
      var done = photoResult_(existing.fileId, existing.viewUrl, true);
      done.rowsAdded = 0;
      return done;
    }

    var row = body.row || {};
    var viewUrl = (existing && existing.viewUrl) || row.photoUrl || '';
    var fileId = (existing && existing.fileId) || '';

    if (body.base64 && !viewUrl) {
      var saved = savePhotoFile_({
        unit: row.property || '',
        subunit: row.unit || '',
        category: row.category || '',
        filename: body.filename,
        mimeType: body.mimeType || 'image/jpeg',
        base64: body.base64,
      });
      viewUrl = saved.viewUrl;
      fileId = saved.fileId;
      if (id) {
        rememberId_(id, { viewUrl: viewUrl, fileId: fileId, rowSaved: false, kind: 'photoRow' });
      }
    }

    var values = rowValues_(row);
    values.photoUrl = viewUrl || '';
    appendMapped_(getRowsSheet_(), CONFIG.ROW_HEADERS, values);
    if (id) {
      rememberId_(id, { viewUrl: viewUrl, fileId: fileId, rowSaved: true, kind: 'photoRow' });
    }

    var result = photoResult_(fileId, viewUrl, false);
    result.rowsAdded = 1;
    return result;
  });
}

// ─── Sheets ──────────────────────────────────────────────────────────────────

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Bind this script to the walk-through Google Sheet (Extensions → Apps Script). A standalone script has no spreadsheet.');
  }
  return ss;
}

function getRowsSheet_() {
  var ss = getSpreadsheet_();
  if (CONFIG.ROWS_SHEET_NAME) {
    var named = ss.getSheetByName(CONFIG.ROWS_SHEET_NAME);
    if (named) return named;
  }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name === CONFIG.IDS_SHEET_NAME) continue;
    if (name === CONFIG.FOLLOWUP_SHEET_NAME) continue;
    if (name === CONFIG.CATEGORIES_SHEET_NAME) continue;
    return sheets[i];
  }
  var created = ss.insertSheet(CONFIG.ROWS_SHEET_NAME || 'Sheet1');
  created.appendRow(CONFIG.ROW_HEADERS);
  return created;
}

function getFollowupSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.FOLLOWUP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.FOLLOWUP_SHEET_NAME);
    sheet.appendRow(CONFIG.FOLLOWUP_HEADERS);
  }
  return sheet;
}

function getIdsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.IDS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.IDS_SHEET_NAME);
    sheet.appendRow(['id', 'viewUrl', 'fileId', 'rowSaved', 'kind', 'savedAt']);
  }
  // Hide the bookkeeping tab. Harmless if the spreadsheet refuses (only tab).
  try { sheet.hideSheet(); } catch (err) {}
  return sheet;
}

function readCategoryList_() {
  var ss = getSpreadsheet_();
  var categories = [];
  var seen = {};
  function add(value) {
    var text = String(value || '').trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    categories.push(text);
  }

  var named = ss.getSheetByName(CONFIG.CATEGORIES_SHEET_NAME);
  if (named && named.getLastRow() > 0) {
    var fromList = named.getRange(1, 1, named.getLastRow(), 1).getValues();
    for (var i = 0; i < fromList.length; i++) {
      var cell = String(fromList[i][0] || '').trim();
      if (i === 0 && cell.toLowerCase() === 'category') continue;
      add(cell);
    }
    if (categories.length) return categories;
  }

  var rowsSheet = getRowsSheet_();
  if (rowsSheet.getLastRow() < 1) return categories;
  var header = headerCells_(rowsSheet);
  var col = headerIndex_(header, 'category');
  if (col < 0) col = 3; // COLUMN D when there is no header row
  var last = rowsSheet.getLastRow();
  var start = headerHas_(header, 'category') ? 2 : 1;
  if (last < start) return categories;
  var values = rowsSheet.getRange(start, col + 1, last - start + 1, 1).getValues();
  for (var r = 0; r < values.length; r++) add(values[r][0]);
  return categories;
}

function rowValues_(row) {
  row = row || {};
  return {
    property: row.property || '',
    unit: row.unit || '',
    inspector: row.inspector || '',
    category: row.category || '',
    note: row.note || '',
    photoUrl: row.photoUrl || '',
    location: row.location || '',
    date: row.date || '',
  };
}

/**
 * Match an existing header row by name so column order on the live sheet is
 * left alone. With no header, append positionally and do not insert a header
 * above walk-through data that is already there.
 */
function appendMapped_(sheet, headers, valuesByHeader) {
  var width = Math.max(sheet.getLastColumn(), headers.length);
  var header = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, width).getValues()[0]
    : [];
  if (!headerHas_(header, 'property') && !headerHas_(header, 'category')) {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      header = headers.slice();
    } else {
      var positional = [];
      for (var i = 0; i < headers.length; i++) positional.push(valuesByHeader[headers[i]] || '');
      sheet.appendRow(positional);
      return;
    }
  }
  var line = [];
  for (var c = 0; c < header.length; c++) {
    var key = canonicalHeader_(header[c]);
    line.push(key && Object.prototype.hasOwnProperty.call(valuesByHeader, key) ? valuesByHeader[key] : '');
  }
  sheet.appendRow(line);
}

function headerCells_(sheet) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function headerHas_(header, name) {
  return headerIndex_(header, name) >= 0;
}

function headerIndex_(header, name) {
  var want = String(name || '').toLowerCase();
  for (var i = 0; i < header.length; i++) {
    if (String(header[i] || '').trim().toLowerCase() === want) return i;
  }
  return -1;
}

function canonicalHeader_(cell) {
  var text = String(cell || '').trim().toLowerCase();
  var headers = CONFIG.ROW_HEADERS.concat(CONFIG.FOLLOWUP_HEADERS);
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase() === text) return headers[i];
  }
  return '';
}

// ─── Idempotency ─────────────────────────────────────────────────────────────

function clientId_(body) {
  if (!body) return '';
  return String(body.id || body.clientId || '').trim();
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function rowAlreadySaved_(id) {
  var existing = getIdRecord_(id);
  return !!(existing && existing.rowSaved);
}

function getIdRecord_(id) {
  if (!id) return null;
  var sheet = getIdsSheet_();
  var found = sheet.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!found || found.getRow() === 1) return null;
  var vals = sheet.getRange(found.getRow(), 1, 1, 6).getValues()[0];
  return {
    rowIndex: found.getRow(),
    id: String(vals[0]),
    viewUrl: vals[1] ? String(vals[1]) : '',
    fileId: vals[2] ? String(vals[2]) : '',
    rowSaved: vals[3] === true || vals[3] === 'true' || vals[3] === 1 || vals[3] === '1',
    kind: vals[4] ? String(vals[4]) : '',
  };
}

function rememberId_(id, info) {
  if (!id) return;
  var sheet = getIdsSheet_();
  var existing = getIdRecord_(id);
  var viewUrl = info.viewUrl || (existing && existing.viewUrl) || '';
  var fileId = info.fileId || (existing && existing.fileId) || '';
  var rowSaved = !!info.rowSaved || !!(existing && existing.rowSaved);
  var kind = info.kind || (existing && existing.kind) || '';
  var record = [id, viewUrl, fileId, rowSaved, kind, new Date()];
  if (existing) {
    sheet.getRange(existing.rowIndex, 1, 1, 6).setValues([record]);
  } else {
    sheet.appendRow(record);
  }
}

// ─── Drive ───────────────────────────────────────────────────────────────────

/**
 * Folder per property, then per unit:
 *   <DRIVE_ROOT_FOLDER_NAME> / <property> / <unit> / file.jpg
 * uploadPhoto calls the property `unit` and the unit number `subunit`,
 * because that is what the phone has always sent.
 */
function savePhotoFile_(body) {
  var base64 = String(body.base64 || '');
  var comma = base64.indexOf(',');
  if (base64.indexOf('data:') === 0 && comma !== -1) base64 = base64.substring(comma + 1);
  if (!base64) throw new Error('Argument cannot be null: data');

  var propertyFolder = getOrCreateChild_(getRootFolder_(), body.unit || 'Unspecified');
  var folder = body.subunit ? getOrCreateChild_(propertyFolder, body.subunit) : propertyFolder;
  var filename = sanitizeName_(body.filename || ('photo_' + new Date().getTime() + '.jpg'), 180);
  if (!/\.jpe?g$/i.test(filename)) filename += '.jpg';

  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, body.mimeType || 'image/jpeg', filename);
  var file = folder.createFile(blob);
  // Anyone with the link can open viewUrl. The file stays saved even if the
  // domain blocks public sharing; the row still gets the Drive URL.
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {}

  return {
    fileId: file.getId(),
    viewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
  };
}

function photoResult_(fileId, viewUrl, alreadySaved) {
  fileId = fileId || '';
  viewUrl = viewUrl || '';
  return {
    success: true,
    alreadySaved: !!alreadySaved,
    fileId: fileId,
    viewUrl: viewUrl,
    thumbUrl: fileId ? ('https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400') : '',
  };
}

function getRootFolder_() {
  var name = CONFIG.DRIVE_ROOT_FOLDER_NAME;
  if (!name) return DriveApp.getRootFolder();
  var found = DriveApp.getFoldersByName(name);
  if (found.hasNext()) return found.next();
  return DriveApp.createFolder(name);
}

function getOrCreateChild_(parent, name) {
  var safe = sanitizeName_(name || 'Unspecified', 80);
  var found = parent.getFoldersByName(safe);
  if (found.hasNext()) return found.next();
  return parent.createFolder(safe);
}

function sanitizeName_(name, maxLen) {
  var cleaned = String(name || '').replace(/[\\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
  if (!cleaned) cleaned = 'Unspecified';
  if (cleaned.length > maxLen) cleaned = cleaned.substring(0, maxLen);
  return cleaned;
}
