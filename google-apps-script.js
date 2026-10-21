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
 *   endSession        → file THIS phone's photo ids into the matching deal
 *       { id, property, location, inspector, ids: [clientId, ...] }
 *       { success, matched, message }
 *       Only those ids are moved. A retry of the same id returns the same
 *       message and does not copy rows again.
 *
 * Unknown actions return { success:false, error:"Unknown action" } so an
 * older copy of the phone page can fall back to uploadPhoto + submitRows.
 *
 * OAuth: SpreadsheetApp.openById and DriveApp moves use the same scopes
 * photo upload already needs — https://www.googleapis.com/auth/spreadsheets
 * and https://www.googleapis.com/auth/drive. No Gmail or external-request scope.
 */
var CONFIG = {
  // Walk-through rows. Leave '' to use the first tab that is not one of the
  // helper tabs below (safest when this script is bound to the existing sheet).
  // Set a name only if the rows live on a specific tab, e.g. 'Sheet1'.
  // Jeff's sheet: walk-through rows live on the 'PhotoNotes' tab (first tab).
  // Pinned so the old 'Inspections' tab can never be picked by mistake.
  ROWS_SHEET_NAME: 'PhotoNotes',

  // Created if missing. Point this at the existing follow-up tab if it has
  // another name (Follow Up, Follow-ups, …).
  // Jeff's sheet: the existing tab is 'FollowUp' (singular).
  FOLLOWUP_SHEET_NAME: 'FollowUp',

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
  // 'time' added: Jeff's tab has a Time column the old script filled in.
  ROW_HEADERS: ['property', 'unit', 'inspector', 'category', 'note', 'photoUrl', 'location', 'date', 'time'],
  FOLLOWUP_HEADERS: ['property', 'location', 'category', 'question', 'date'],

  // Live header names on Jeff's 'PhotoNotes' tab that differ from the keys
  // above (lower-case header -> key). Without these, User, Notes and
  // Photo link would be left blank on every new row.
  // Live header row: Property | Unit | Location | User | Category | Notes | Photo link | Date | Time
  HEADER_ALIASES: {
    'user': 'inspector',
    'notes': 'note',
    'photo link': 'photoUrl',
  },

  // Loose Google Sheets (and some property subfolders) live here.
  // Inspectors type a short property name; endSession matches a workbook.
  CAMPAIGN_FOLDER_ID: '15nJZofaS0kvAWpoawBACUc9XV5seHiYQ',
  DEAL_TAB_NAME: 'Photonotes',
  DEAL_HEADERS: ['Property', 'Unit', 'Location', 'User', 'Category', 'Notes', 'Photo link', 'Date', 'Time'],
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
    if (action === 'endSession') return json_(endSession_(body));
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
    var lastValues = null;
    rows.forEach(function (row) {
      lastValues = rowValues_(row);
      appendMapped_(sheet, CONFIG.ROW_HEADERS, lastValues);
      added++;
    });
    if (added === 1) tagRowId_(sheet, id);
    if (id) {
      var photoUrl = lastValues && lastValues.photoUrl ? String(lastValues.photoUrl) : '';
      rememberId_(id, {
        viewUrl: photoUrl,
        fileId: '',
        rowSaved: true,
        kind: 'rows',
        extra: snapshotExtra_(lastValues || {}),
      });
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
    var rowsSheet = getRowsSheet_();
    appendMapped_(rowsSheet, CONFIG.ROW_HEADERS, values);
    tagRowId_(rowsSheet, id);
    if (id) {
      rememberId_(id, {
        viewUrl: viewUrl,
        fileId: fileId,
        rowSaved: true,
        kind: 'photoRow',
        extra: snapshotExtra_(values),
      });
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
    sheet.appendRow(['id', 'viewUrl', 'fileId', 'rowSaved', 'kind', 'savedAt', 'extra']);
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
    // The phone sends only a date; stamp the save time like the old script.
    time: row.time || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'h:mm:ss a'),
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
  var aliases = CONFIG.HEADER_ALIASES || {};
  if (Object.prototype.hasOwnProperty.call(aliases, text)) return aliases[text];
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
  lock.waitLock(60000);
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
  var width = Math.max(7, sheet.getLastColumn());
  var vals = sheet.getRange(found.getRow(), 1, 1, width).getValues()[0];
  return {
    rowIndex: found.getRow(),
    id: String(vals[0]),
    viewUrl: vals[1] ? String(vals[1]) : '',
    fileId: vals[2] ? String(vals[2]) : '',
    rowSaved: vals[3] === true || vals[3] === 'true' || vals[3] === 1 || vals[3] === '1',
    kind: vals[4] ? String(vals[4]) : '',
    extra: parseExtra_(vals[6]),
  };
}

function rememberId_(id, info) {
  if (!id) return;
  info = info || {};
  var sheet = getIdsSheet_();
  var existing = getIdRecord_(id);
  var viewUrl = info.viewUrl || (existing && existing.viewUrl) || '';
  var fileId = info.fileId || (existing && existing.fileId) || '';
  var rowSaved = !!info.rowSaved || !!(existing && existing.rowSaved);
  var kind = info.kind || (existing && existing.kind) || '';
  var extra = existing && existing.extra ? existing.extra : {};
  var incoming = info.extra || {};
  for (var key in incoming) {
    if (Object.prototype.hasOwnProperty.call(incoming, key)) extra[key] = incoming[key];
  }
  var record = [id, viewUrl, fileId, rowSaved, kind, new Date(), JSON.stringify(extra)];
  if (existing) {
    sheet.getRange(existing.rowIndex, 1, 1, 7).setValues([record]);
  } else {
    sheet.appendRow(record);
  }
}

function parseExtra_(value) {
  if (!value) return {};
  if (typeof value === 'object') return value;
  try {
    var parsed = JSON.parse(String(value));
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (err) {
    return {};
  }
}

function snapshotExtra_(values) {
  values = values || {};
  return {
    property: values.property || '',
    unit: values.unit || '',
    inspector: values.inspector || '',
    category: values.category || '',
    note: values.note || '',
    photoUrl: values.photoUrl || '',
    location: values.location || '',
    date: values.date || '',
    time: values.time || '',
    filed: false,
    fileMoved: false,
  };
}

function tagRowId_(sheet, clientId) {
  if (!sheet || !clientId) return;
  var row = sheet.getLastRow();
  if (row < 1) return;
  try {
    sheet.getRange(row, 1).addDeveloperMetadata('photonotesId', String(clientId));
  } catch (err) {}
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

// ─── End session: file this phone's rows into the deal workbook ─────────────
// Matching is deliberate and refuses ties. Workbook titles look like
// "Millie Beneke Manor - MB 9/26". Compared names ignore case, punctuation,
// and that " - MB m/yy" suffix, and a leading "The ". "Millie" can match
// "Millie Beneke Manor" when it is the only workbook that fits. Location
// (the city after the comma, or the last word) breaks a tie against folder
// names such as Millie_Beneke_Manor_Glencoe. Anything still ambiguous is
// left on the master PhotoNotes tab.

function endSession_(body) {
  return withLock_(function () {
    var id = clientId_(body);
    if (id) {
      var prior = getIdRecord_(id);
      if (prior && prior.extra && prior.extra.resultMessage) {
        return {
          success: true,
          alreadySaved: true,
          matched: !!prior.extra.matched,
          message: String(prior.extra.resultMessage),
        };
      }
    }

    var ids = normalizeIdList_(body && body.ids);
    var property = String((body && body.property) || '').trim();
    var location = String((body && body.location) || '').trim();
    if (!ids.length) {
      return finishEndSession_(id, 'No photos in this session.', false);
    }

    var listed = listCampaignMatches_(property);
    var picked = pickPropertyMatch_(property, location, listed.workbooks, listed.folders);
    if (!picked.confident) {
      return finishEndSession_(id, noMatchMessage_(property, picked.candidates), false);
    }

    var filed = fileSessionRows_(ids, picked.workbook);
    var photos = moveSessionPhotos_(filed.fileIds, picked.folder);
    var message = filingResultMessage_({
      count: filed.count,
      missing: filed.missing,
      workbookName: picked.workbook.name,
      folderName: photos.folderName,
      photosLeft: !picked.folder,
    });
    return finishEndSession_(id, message, true);
  });
}

function finishEndSession_(id, message, matched) {
  if (id) {
    rememberId_(id, {
      rowSaved: true,
      kind: 'endSession',
      extra: { resultMessage: message, matched: !!matched },
    });
  }
  return { success: true, matched: !!matched, message: message };
}

function normalizeIdList_(ids) {
  var out = [];
  var seen = {};
  var list = Object.prototype.toString.call(ids) === '[object Array]' ? ids : [];
  for (var i = 0; i < list.length; i++) {
    var id = String(list[i] || '').trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    out.push(id);
  }
  return out;
}

function noMatchMessage_(property, candidates) {
  var label = String(property || '').trim() || 'this property';
  if (candidates && candidates.length) {
    return 'No confident match for "' + label + '" — candidates: ' + candidates.join('; ') + '. Photos left in PhotoNotes';
  }
  return 'No confident match for "' + label + '" — photos left in PhotoNotes';
}

function filingResultMessage_(info) {
  if (!info.count) {
    return 'No photos from this session were on the PhotoNotes tab. Nothing was moved.';
  }
  var noun = info.count === 1 ? 'photo' : 'photos';
  var msg = 'Moved ' + info.count + ' ' + noun + ' to ' + info.workbookName + ' (Photonotes tab)';
  if (info.folderName) {
    msg += ' and Drive folder ' + info.folderName;
  } else if (info.photosLeft) {
    msg += '. No Drive folder in Property Campaigns; photos left in PhotoNotes';
  }
  if (info.missing) {
    msg += '. ' + info.missing + (info.missing === 1 ? ' was' : ' were') + ' not on the PhotoNotes tab';
  }
  return msg;
}

function listCampaignMatches_(property) {
  var root = DriveApp.getFolderById(CONFIG.CAMPAIGN_FOLDER_ID);
  var query = normalizePropertyName_(property);
  var tokens = query ? query.split(' ') : [];
  var token = '';
  for (var i = 0; i < tokens.length; i++) {
    if (tokens[i].length > token.length) token = tokens[i];
  }
  if (token.length < 3) return { workbooks: [], folders: [] };
  var escaped = token.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  var workbooks = [];
  var seenFiles = {};
  var files = root.searchFiles(
    "mimeType = '" + MimeType.GOOGLE_SHEETS + "' and title contains '" + escaped + "'"
  );
  var fileCount = 0;
  while (files.hasNext() && fileCount < 40) {
    var file = files.next();
    fileCount++;
    if (seenFiles[file.getId()]) continue;
    if (!isDirectChild_(file, root.getId())) continue;
    seenFiles[file.getId()] = true;
    workbooks.push({ id: file.getId(), name: file.getName() });
  }
  var folders = [];
  var seenFolders = {};
  var foundFolders = root.searchFolders("title contains '" + escaped + "'");
  var folderCount = 0;
  while (foundFolders.hasNext() && folderCount < 40) {
    var folder = foundFolders.next();
    folderCount++;
    if (seenFolders[folder.getId()]) continue;
    if (!isDirectChild_(folder, root.getId())) continue;
    seenFolders[folder.getId()] = true;
    folders.push({ id: folder.getId(), name: folder.getName() });
  }
  return { workbooks: workbooks, folders: folders };
}

function isDirectChild_(item, parentId) {
  var parents = item.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === parentId) return true;
  }
  return false;
}

function stripCampaignSuffix_(text) {
  return String(text || '')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\s*-\s*mb\s+\d{1,2}\s*\/\s*\d{2,4}\s*$/i, '')
    .trim();
}

function normalizePropertyName_(value) {
  var text = stripCampaignSuffix_(String(value || '').toLowerCase());
  text = text.replace(/[_/]+/g, ' ');
  text = text.replace(/[^a-z0-9\s]/g, ' ');
  text = text.replace(/\s+/g, ' ').trim();
  text = text.replace(/^the\s+/, '');
  return text;
}

function cityFromLocation_(location) {
  var text = String(location || '').trim();
  if (!text) return '';
  var parts = text.split(',');
  for (var i = parts.length - 1; i >= 0; i--) {
    var chunk = parts[i].replace(/\b\d{5}(?:-\d{4})?\b/g, ' ');
    var tokens = normalizePropertyName_(chunk).split(' ').filter(function (token) {
      return token && token.length > 2 && !/^\d+$/.test(token);
    });
    if (tokens.length) return tokens[tokens.length - 1];
  }
  return '';
}

function tokensInOrder_(queryTokens, candidateTokens) {
  var qi = 0;
  for (var i = 0; i < candidateTokens.length && qi < queryTokens.length; i++) {
    if (candidateTokens[i] === queryTokens[qi]) qi++;
  }
  return qi === queryTokens.length;
}

function scorePropertyName_(query, candidate) {
  if (!query || !candidate) return 0;
  if (query === candidate) return 100;
  var q = query.split(' ');
  var c = candidate.split(' ');
  if (!tokensInOrder_(q, c)) return 0;
  if (q.length === 1) {
    if (c.length === 1) return 100;
    if (c[0] === q[0]) return 88;
    return q[0].length >= 4 ? 80 : 68;
  }
  if (candidate.indexOf(query + ' ') === 0) return 92;
  return 84;
}

function folderMatchesWorkbook_(folderNorm, workbookNorm) {
  if (!folderNorm || !workbookNorm) return false;
  if (folderNorm === workbookNorm) return true;
  if (folderNorm.indexOf(workbookNorm + ' ') === 0) return true;
  return tokensInOrder_(workbookNorm.split(' '), folderNorm.split(' '));
}

function pickPropertyMatch_(property, location, workbooks, folders) {
  var query = normalizePropertyName_(property);
  var qTokens = query ? query.split(' ').filter(Boolean) : [];
  var longEnough = false;
  for (var t = 0; t < qTokens.length; t++) {
    if (qTokens[t].length >= 3) longEnough = true;
  }
  if (!query || !longEnough) {
    return { confident: false, candidates: [], workbook: null, folder: null };
  }
  var city = cityFromLocation_(location);
  var scored = [];
  for (var i = 0; i < workbooks.length; i++) {
    var wb = workbooks[i];
    var nameNorm = normalizePropertyName_(wb.name);
    var base = scorePropertyName_(query, nameNorm);
    if (base < 70) continue;
    var linked = [];
    for (var f = 0; f < folders.length; f++) {
      if (folderMatchesWorkbook_(normalizePropertyName_(folders[f].name), nameNorm)) {
        linked.push(folders[f]);
      }
    }
    var cityHit = false;
    if (city) {
      if (nameNorm.split(' ').indexOf(city) !== -1) cityHit = true;
      for (var c = 0; c < linked.length; c++) {
        if (normalizePropertyName_(linked[c].name).split(' ').indexOf(city) !== -1) cityHit = true;
      }
    }
    scored.push({
      wb: wb,
      base: base,
      score: base + (cityHit ? 20 : 0),
      linked: linked,
    });
  }
  scored.sort(function (a, b) { return b.score - a.score || b.base - a.base; });
  var candidates = [];
  for (var s = 0; s < scored.length && s < 5; s++) candidates.push(scored[s].wb.name);
  if (!scored.length) {
    return { confident: false, candidates: [], workbook: null, folder: null };
  }
  var best = scored[0];
  var second = scored.length > 1 ? scored[1] : null;
  var gap = second ? best.score - second.score : 100;
  var confident = best.score >= 80 && (!second || gap >= 12);
  if (best.base === 100 && (!second || second.base < 100)) confident = true;
  return {
    confident: confident,
    workbook: confident ? best.wb : null,
    folder: confident ? chooseCampaignFolder_(best, city) : null,
    candidates: confident ? [] : candidates,
  };
}

function chooseCampaignFolder_(best, city) {
  var linked = best.linked || [];
  if (!linked.length) return null;
  if (linked.length === 1) return linked[0];
  if (!city) return null;
  var hits = [];
  for (var i = 0; i < linked.length; i++) {
    if (normalizePropertyName_(linked[i].name).split(' ').indexOf(city) !== -1) hits.push(linked[i]);
  }
  return hits.length === 1 ? hits[0] : null;
}

function fileSessionRows_(ids, workbook) {
  var dealSs = SpreadsheetApp.openById(workbook.id);
  var dealSheet = getDealPhotoSheet_(dealSs);
  ensureDealHeaders_(dealSheet);
  var masterSs = getSpreadsheet_();
  var master = readMasterRows_();
  var used = {};
  var fileIds = [];
  var count = 0;
  var missing = 0;
  var deleteRows = [];

  for (var i = 0; i < ids.length; i++) {
    var clientId = ids[i];
    var record = getIdRecord_(clientId);
    var alreadyFiled = metadataRow_(dealSs, clientId) || (record && record.extra && record.extra.filed);
    var located = locateMasterRow_(masterSs, master, record, clientId, used);
    if (alreadyFiled) {
      count++;
      if (located && !used[located.rowIndex]) {
        used[located.rowIndex] = true;
        deleteRows.push(located.rowIndex);
      }
      if (record && record.fileId) fileIds.push(record.fileId);
      continue;
    }
    if (!located || !located.values) {
      if (record && record.viewUrl && dealHasPhoto_(dealSheet, record.viewUrl)) {
        count++;
        rememberId_(clientId, { extra: { filed: true } });
        if (record.fileId) fileIds.push(record.fileId);
      } else {
        missing++;
      }
      continue;
    }
    used[located.rowIndex] = true;
    appendMapped_(dealSheet, CONFIG.ROW_HEADERS, located.values);
    tagRowId_(dealSheet, clientId);
    deleteRows.push(located.rowIndex);
    count++;
    if (record && record.fileId) fileIds.push(record.fileId);
    rememberId_(clientId, { extra: { filed: true } });
  }

  deleteRows.sort(function (a, b) { return b - a; });
  for (var d = 0; d < deleteRows.length; d++) {
    master.sheet.deleteRow(deleteRows[d]);
  }
  return { count: count, missing: missing, fileIds: fileIds };
}

function getDealPhotoSheet_(ss) {
  var sheets = ss.getSheets();
  var want = String(CONFIG.DEAL_TAB_NAME || 'Photonotes').toLowerCase();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).toLowerCase() === want) return sheets[i];
  }
  var created = ss.insertSheet(CONFIG.DEAL_TAB_NAME || 'Photonotes');
  created.appendRow(CONFIG.DEAL_HEADERS);
  return created;
}

function ensureDealHeaders_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow(CONFIG.DEAL_HEADERS);
}

function metadataRow_(ss, clientId) {
  if (!ss || !clientId) return null;
  try {
    var found = ss.createDeveloperMetadataFinder()
      .withKey('photonotesId')
      .withValue(String(clientId))
      .find();
    if (!found || !found.length) return null;
    var range = found[0].getLocation().getRange();
    if (!range) return null;
    return { sheet: range.getSheet(), rowIndex: range.getRow() };
  } catch (err) {
    return null;
  }
}

function readMasterRows_() {
  var sheet = getRowsSheet_();
  var last = sheet.getLastRow();
  var width = Math.max(sheet.getLastColumn(), CONFIG.ROW_HEADERS.length);
  if (last < 1 || width < 1) return { sheet: sheet, rows: [] };
  var grid = sheet.getRange(1, 1, last, width).getValues();
  var header = grid[0];
  var hasHeader = headerHas_(header, 'property') || headerHas_(header, 'category');
  var rows = [];
  for (var r = hasHeader ? 1 : 0; r < grid.length; r++) {
    var values = {};
    if (hasHeader) {
      for (var c = 0; c < header.length; c++) {
        var key = canonicalHeader_(header[c]);
        if (key) values[key] = grid[r][c];
      }
    } else {
      for (var h = 0; h < CONFIG.ROW_HEADERS.length && h < grid[r].length; h++) {
        values[CONFIG.ROW_HEADERS[h]] = grid[r][h];
      }
    }
    rows.push({ rowIndex: r + 1, values: values });
  }
  return { sheet: sheet, rows: rows };
}

function cellText_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value || '').trim().toLowerCase();
}

function sameCell_(a, b) {
  return cellText_(a) === cellText_(b);
}

function locateMasterRow_(masterSs, master, record, clientId, used) {
  var tagged = metadataRow_(masterSs, clientId);
  var taggedIndex = masterRowIndex_(master, tagged);
  if (taggedIndex && !used[taggedIndex]) {
    var taggedValues = valuesAt_(master, taggedIndex);
    if (taggedValues) return { rowIndex: taggedIndex, values: taggedValues };
  }
  return findMasterRow_(master, record, used);
}

function masterRowIndex_(master, located) {
  if (!located || !located.rowIndex) return 0;
  if (located.sheet && located.sheet.getSheetId() !== master.sheet.getSheetId()) return 0;
  return located.rowIndex;
}

function valuesAt_(master, rowIndex) {
  for (var i = 0; i < master.rows.length; i++) {
    if (master.rows[i].rowIndex === rowIndex) return master.rows[i].values;
  }
  return null;
}

function findMasterRow_(master, record, used) {
  if (!record) return null;
  var snap = record.extra || {};
  var url = String(record.viewUrl || snap.photoUrl || '').trim();
  if (url) {
    for (var i = 0; i < master.rows.length; i++) {
      var row = master.rows[i];
      if (used[row.rowIndex]) continue;
      if (String(row.values.photoUrl || '').trim() === url) return row;
    }
    return null;
  }
  var hits = [];
  for (var n = 0; n < master.rows.length; n++) {
    var candidate = master.rows[n];
    if (used[candidate.rowIndex]) continue;
    var values = candidate.values;
    if (sameCell_(values.property, snap.property) &&
        sameCell_(values.unit, snap.unit) &&
        sameCell_(values.inspector, snap.inspector) &&
        sameCell_(values.category, snap.category) &&
        sameCell_(values.note, snap.note)) {
      hits.push(candidate);
    }
  }
  return hits.length === 1 ? hits[0] : null;
}

function dealHasPhoto_(sheet, photoUrl) {
  if (!photoUrl || sheet.getLastRow() < 1) return false;
  var header = headerCells_(sheet);
  var col = -1;
  for (var i = 0; i < header.length; i++) {
    if (canonicalHeader_(header[i]) === 'photoUrl') { col = i + 1; break; }
  }
  if (col < 1) return false;
  var found = sheet.getRange(1, col, sheet.getLastRow(), 1).createTextFinder(String(photoUrl)).matchEntireCell(true).findNext();
  return !!(found && found.getRow() > 1);
}

function moveSessionPhotos_(fileIds, folderInfo) {
  if (!folderInfo || !folderInfo.id) return { folderName: '' };
  var campaign = DriveApp.getFolderById(folderInfo.id);
  var idSet = {};
  var files = [];
  for (var i = 0; i < fileIds.length; i++) {
    var fileId = fileIds[i];
    if (!fileId || idSet[fileId]) continue;
    idSet[fileId] = true;
    try {
      var file = DriveApp.getFileById(fileId);
      if (!file.isTrashed()) files.push(file);
    } catch (err) {}
  }
  var groups = {};
  for (var f = 0; f < files.length; f++) {
    var one = files[f];
    if (isUnderFolder_(one, campaign.getId())) continue;
    var parent = firstParent_(one);
    var key = parent ? parent.getId() : '';
    if (!groups[key]) groups[key] = { folder: parent, files: [] };
    groups[key].files.push(one);
  }
  for (var key in groups) {
    if (!Object.prototype.hasOwnProperty.call(groups, key)) continue;
    var group = groups[key];
    movePhotoGroup_(group, idSet, campaign);
  }
  return { folderName: campaign.getName() };
}

function firstParent_(item) {
  var parents = item.getParents();
  return parents.hasNext() ? parents.next() : null;
}

function isUnderFolder_(item, ancestorId) {
  var frontier = [];
  var parents = item.getParents();
  while (parents.hasNext()) frontier.push(parents.next());
  var seen = {};
  var guard = 0;
  while (frontier.length && guard < 25) {
    var folder = frontier.shift();
    var id = folder.getId();
    if (id === ancestorId) return true;
    if (seen[id]) { guard++; continue; }
    seen[id] = true;
    var up = folder.getParents();
    while (up.hasNext()) frontier.push(up.next());
    guard++;
  }
  return false;
}

function onlyTheseFiles_(folder, idSet) {
  if (folder.getFolders().hasNext()) return false;
  var any = false;
  var files = folder.getFiles();
  while (files.hasNext()) {
    any = true;
    if (!idSet[files.next().getId()]) return false;
  }
  return any;
}

function folderIsEmpty_(folder) {
  return !folder.getFiles().hasNext() && !folder.getFolders().hasNext();
}

function mergeFolderInto_(sourceFolder, destParent) {
  var existing = destParent.getFoldersByName(sourceFolder.getName());
  if (!existing.hasNext()) {
    sourceFolder.moveTo(destParent);
    return;
  }
  var dest = existing.next();
  var files = sourceFolder.getFiles();
  while (files.hasNext()) files.next().moveTo(dest);
  var subs = sourceFolder.getFolders();
  while (subs.hasNext()) mergeFolderInto_(subs.next(), dest);
  if (folderIsEmpty_(sourceFolder)) {
    try { sourceFolder.setTrashed(true); } catch (err) {}
  }
}

function movePhotoGroup_(group, idSet, campaign) {
  if (!group.folder) {
    for (var i = 0; i < group.files.length; i++) group.files[i].moveTo(campaign);
    return;
  }
  var unitFolder = group.folder;
  var propertyFolder = firstParent_(unitFolder);
  var destParent = campaign;
  if (propertyFolder && propertyFolder.getName() !== CONFIG.DRIVE_ROOT_FOLDER_NAME) {
    destParent = getOrCreateChild_(campaign, propertyFolder.getName());
  }
  if (onlyTheseFiles_(unitFolder, idSet)) {
    mergeFolderInto_(unitFolder, destParent);
    if (propertyFolder &&
        propertyFolder.getName() !== CONFIG.DRIVE_ROOT_FOLDER_NAME &&
        folderIsEmpty_(propertyFolder)) {
      try { propertyFolder.setTrashed(true); } catch (err) {}
    }
    return;
  }
  var destUnit = getOrCreateChild_(destParent, unitFolder.getName());
  for (var f = 0; f < group.files.length; f++) {
    if (!isUnderFolder_(group.files[f], campaign.getId())) group.files[f].moveTo(destUnit);
  }
}
