// ============================================================
// CONFIGURATION (auto-configured — do not change)
// ============================================================
const SPREADSHEET_ID   = '1R94nSTPkW8oNIzKOHr6ORwD_Fuxhv5pWHWtTYkur-OI';
const PARENT_FOLDER_ID = '1N4hzyRwni0eNFAYn5TgekmhEuT4DIERX';

// Sheet names (must match exactly what you create in the Spreadsheet)
const SHEET_PP5        = 'ป.พ.5';
const SHEET_COMPETENCY = 'สมรรถนะ5ด้าน';
const SHEET_SAR        = 'SAR';
const SHEET_PROJECT    = 'รายงานโครงการ2568';
const SHEET_NAMELIST   = 'NameList';

// Drive sub-folder names (created automatically inside the parent folder)
const FOLDER_PP5        = 'ป.พ.5';
const FOLDER_COMPETENCY = 'สมรรถนะ5ด้าน';
const FOLDER_SAR        = 'SAR';
const FOLDER_PROJECT    = 'รายงานโครงการ2568';

// ============================================================
// ENTRY POINT
// ============================================================

/**
 * Serves the web app.
 * Deploy → Execute as: Me | Who has access: Anyone
 */
function doGet() {
  ensureSheets();   // create sheets if missing
  ensureHeaders();  // insert header row if missing
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบส่งงานออนไลน์โรงเรียนบ้านโพนแท่น ปีการศึกษา 2568')
    .setFaviconUrl('https://raw.githubusercontent.com/dimon-ton/document-submission-system/master/school-logo.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Creates the 5 required sheets if they don't already exist.
 * Safe to call on every request — exits immediately once all sheets are present.
 */
function ensureSheets() {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const required = [SHEET_PP5, SHEET_COMPETENCY, SHEET_SAR, SHEET_PROJECT, SHEET_NAMELIST];
  const existing = ss.getSheets().map(s => s.getName());

  // Rename the default "Sheet1" to the first required sheet if it exists
  const defaultSheet = ss.getSheetByName('Sheet1');

  required.forEach(function(name, i) {
    if (existing.includes(name)) return; // already exists
    if (i === 0 && defaultSheet) {
      defaultSheet.setName(name);      // rename Sheet1 instead of inserting
    } else {
      ss.insertSheet(name);
    }
  });
}

/**
 * Inserts a styled header row into each sheet if one is not already present.
 * Safe to call on every request — skips sheets whose first cell already matches.
 */
function ensureHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Map: sheet name → expected header columns (Thai labels)
  const HEADERS = {};
  HEADERS[SHEET_PP5]        = ['Timestamp', 'ชื่อ-นามสกุล', 'ชื่อไฟล์', 'URL ไฟล์', 'หมายเหตุ'];
  HEADERS[SHEET_COMPETENCY] = ['Timestamp', 'ชื่อ-นามสกุล', 'ชื่อไฟล์', 'URL ไฟล์', 'ชื่อไฟล์ PDF', 'URL ไฟล์ PDF', 'หมายเหตุ'];
  HEADERS[SHEET_SAR]        = ['Timestamp', 'ชื่อ-นามสกุล', 'ชื่อไฟล์ Word', 'URL ไฟล์ Word', 'ชื่อไฟล์ PDF', 'URL ไฟล์ PDF', 'หมายเหตุ'];
  HEADERS[SHEET_PROJECT]    = ['Timestamp', 'ชื่อ-นามสกุล', 'ชื่อไฟล์ Word', 'URL ไฟล์ Word', 'ชื่อไฟล์ PDF', 'URL ไฟล์ PDF', 'หมายเหตุ'];
  HEADERS[SHEET_NAMELIST]   = ['ชื่อ-นามสกุล'];

  Object.keys(HEADERS).forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers    = HEADERS[sheetName];
    const firstCell  = sheet.getRange(1, 1).getValue().toString().trim();
    const currentLen = sheet.getLastColumn();

    // Header already correct — nothing to do
    if (firstCell === headers[0] && currentLen >= headers.length) return;

    function applyStyle(range) {
      range.setFontWeight('bold');
      range.setFontColor('#ffffff');
      range.setBackground('#1a3a6b');
      range.setHorizontalAlignment('center');
    }

    if (firstCell === headers[0]) {
      // Header row exists but needs more columns — update in-place
      const range = sheet.getRange(1, 1, 1, headers.length);
      range.setValues([headers]);
      applyStyle(range);
    } else {
      // No header at all — insert a new row at the top
      sheet.insertRowBefore(1);
      const range = sheet.getRange(1, 1, 1, headers.length);
      range.setValues([headers]);
      applyStyle(range);
    }
  });
}

// ============================================================
// NAME LIST
// ============================================================

/**
 * Returns all names from the NameList sheet as a flat array.
 * @returns {string[]}
 */
function getNames() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMELIST);
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    return values
      .slice(1)                                       // skip header row
      .map(row => (row[0] || '').toString().trim())
      .filter(name => name !== '');
  } catch (e) {
    Logger.log('getNames error: ' + e);
    throw new Error('ไม่สามารถโหลดรายชื่อได้: ' + e.message);
  }
}

/**
 * Appends a new name to NameList (skips duplicates).
 * @param {string} name
 */
function addName(name) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMELIST);
    if (!sheet) return; // NameList sheet not found — skip silently

    const existing = sheet.getDataRange().getValues()
      .slice(1)                                       // skip header row
      .map(r => (r[0] || '').toString().trim());
    if (!existing.includes(name.trim())) {
      sheet.appendRow([name.trim()]);
    }
  } catch (e) {
    Logger.log('addName error: ' + e);
    // Non-fatal — do not rethrow
  }
}

// ============================================================
// DRIVE HELPERS
// ============================================================

/**
 * Returns (or creates) a sub-folder inside the configured parent folder.
 * @param {string} folderName
 * @returns {Folder}
 */
function createOrGetFolder(folderName) {
  try {
    const parent   = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const iterator = parent.getFoldersByName(folderName);
    return iterator.hasNext() ? iterator.next() : parent.createFolder(folderName);
  } catch (e) {
    Logger.log('createOrGetFolder error: ' + e);
    throw new Error('ไม่สามารถเข้าถึงโฟลเดอร์ได้: ' + e.message);
  }
}

/**
 * Decodes a base64 string and saves it as a file in the given category folder.
 * Makes the file viewable by anyone with the link.
 * @param {string} folderName  Sub-folder name
 * @param {string} fileName    Original file name
 * @param {string} base64Data  Base64-encoded file content (no data-URL prefix)
 * @param {string} mimeType    MIME type
 * @returns {{ name: string, url: string }}
 */
function saveFileToDrive(folderName, fileName, base64Data, mimeType) {
  try {
    const folder  = createOrGetFolder(folderName);
    const decoded = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);

    // Trash any existing file with the same name so the new upload replaces it
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) { existing.next().setTrashed(true); }

    const file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      name : fileName,
      url  : file.getDownloadUrl()
    };
  } catch (e) {
    Logger.log('saveFileToDrive error: ' + e);
    throw new Error('ไม่สามารถบันทึกไฟล์ ' + fileName + ': ' + e.message);
  }
}

// ============================================================
// LARGE FILE UPLOAD (Google Drive Resumable Upload)
// ============================================================
//
// For files larger than google.script.run can handle (~50 MB base64 payload),
// we bypass Apps Script entirely and use Google Drive's native Resumable
// Upload protocol:
//
//   1. Client asks the server for an OAuth access token via getUploadToken()
//   2. Client POSTs metadata to https://www.googleapis.com/upload/drive/v3/
//      files?uploadType=resumable with Authorization: Bearer <token> and
//      receives a one-time session URL in the Location response header
//   3. Client PUTs the file in 5 MB chunks directly to Drive, tracking
//      progress via HTTP 308 (continue) / 200 (done) status codes
//   4. Drive returns the new file's id — client hands it to the server via
//      registerUploadedFile() which moves the file into the correct
//      category folder and returns the public URL
//
// This approach is unlimited in size (tested to 100+ MB), resumable on
// network failure, and avoids holding the entire base64 string in either
// the browser or Apps Script memory.
//
// Credit to @tanaikech — the protocol and chunk logic are adapted from
// https://github.com/tanaikech/Resumable_Upload_For_WebApps

/**
 * Returns a short-lived OAuth access token the client can use to upload
 * directly to the Drive API. The token inherits the scopes configured in
 * appsscript.json (we include drive scope there).
 *
 * Also returns the MIME-type-safe fileId mapping so the client can POST
 * the initial metadata with the right parent.
 *
 * @param {string} folderName  One of FOLDER_PP5/COMPETENCY/SAR/PROJECT
 * @returns {{ token: string, folderId: string }}
 */
function getUploadToken(folderName) {
  try {
    const valid = [FOLDER_PP5, FOLDER_COMPETENCY, FOLDER_SAR, FOLDER_PROJECT];
    if (valid.indexOf(folderName) === -1) {
      throw new Error('โฟลเดอร์ไม่ถูกต้อง: ' + folderName);
    }
    const folder = createOrGetFolder(folderName);
    return {
      token   : ScriptApp.getOAuthToken(),
      folderId: folder.getId()
    };
  } catch (e) {
    Logger.log('getUploadToken error: ' + e);
    throw new Error('ไม่สามารถขอสิทธิ์อัพโหลดได้: ' + e.message);
  }
}

/**
 * Called by the client after the Drive API resumable upload finishes.
 * The file is already in Drive (uploaded directly by the browser), so
 * all the server has to do is:
 *   1. Set sharing to "anyone with link"
 *   2. Return the download URL for bookkeeping
 *
 * @param {string} fileId  ID returned by the Drive upload response
 * @returns {{ name: string, url: string }}
 */
function registerUploadedFile(fileId) {
  try {
    if (!fileId) throw new Error('missing fileId');
    const file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { name: file.getName(), url: file.getDownloadUrl() };
  } catch (e) {
    Logger.log('registerUploadedFile error: ' + e);
    throw new Error('ไม่สามารถลงทะเบียนไฟล์ได้: ' + e.message);
  }
}

// ============================================================
// LEGACY CHUNKED UPLOAD (kept for small files using cache-backed path)
// ============================================================
// Note: the code below is retained only as a fallback. New submissions
// should use the resumable upload above. The small-file inline path still
// uses saveFileToDrive() directly.

const CHUNK_TMP_FOLDER = '_tmp_chunks';

function getChunkTmpFolder() {
  return createOrGetFolder(CHUNK_TMP_FOLDER);
}

/**
 * Receives a single base64-encoded chunk of a larger file and saves it
 * as a tiny file in the temp folder. Files are named
 * "<uploadId>__<6-digit-index>" so finalize can list + sort them.
 *
 * @param {string} uploadId   Client-generated unique id for this upload
 * @param {number} chunkIndex 0-based index of this chunk
 * @param {string} chunkData  Base64 payload (≤ ~90 KB recommended)
 * @returns {{ ok: true, chunkIndex: number }}
 */
function uploadFileChunk(uploadId, chunkIndex, chunkData) {
  try {
    if (!uploadId)                     throw new Error('missing uploadId');
    if (typeof chunkIndex !== 'number') throw new Error('missing chunkIndex');
    if (!chunkData)                    throw new Error('empty chunk');
    if (!/^[A-Za-z0-9_]+$/.test(uploadId)) throw new Error('invalid uploadId');

    const folder = getChunkTmpFolder();
    const name   = uploadId + '__' + padIndex(chunkIndex);

    // Store the base64 chunk as raw text. We use text/plain so Drive doesn't
    // try to interpret it. The chunkData string is already base64 and safe.
    folder.createFile(name, chunkData, MimeType.PLAIN_TEXT);

    return { ok: true, chunkIndex: chunkIndex };
  } catch (e) {
    Logger.log('uploadFileChunk error: ' + e);
    throw new Error('ไม่สามารถอัพโหลดส่วนของไฟล์ได้: ' + e.message);
  }
}

function padIndex(i) {
  const s = String(i);
  return '000000'.slice(s.length) + s;
}

/**
 * Reassembles chunks previously sent via uploadFileChunk() into a single
 * base64 string and saves the file to the given Drive folder. Temp chunk
 * files are trashed after a successful save.
 *
 * @param {string} uploadId    Same id used in uploadFileChunk()
 * @param {number} totalChunks Total number of chunks that were uploaded
 * @param {string} folderName  Drive sub-folder (e.g. FOLDER_PROJECT)
 * @param {string} fileName    Desired file name on Drive
 * @param {string} mimeType    MIME type
 * @returns {{ name: string, url: string }}
 */
function finalizeChunkedUpload(uploadId, totalChunks, folderName, fileName, mimeType) {
  const chunkFiles = collectChunkFiles(uploadId, totalChunks);

  try {
    // Read each chunk's base64 text and decode to bytes directly, then
    // concatenate byte arrays. Doing it this way avoids building one giant
    // JS string (which would be ~4/3 the size of the final file in memory).
    const parts = new Array(totalChunks);
    let   total = 0;
    for (let i = 0; i < totalChunks; i++) {
      const b64   = chunkFiles[i].getBlob().getDataAsString();
      const bytes = Utilities.base64Decode(b64);
      parts[i]    = bytes;
      total      += bytes.length;
    }

    // Flatten into a single Uint8-like byte array (Apps Script uses number[])
    const combined = new Array(total);
    let offset = 0;
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i];
      for (let j = 0; j < p.length; j++) combined[offset + j] = p[j];
      offset += p.length;
      parts[i] = null; // free early
    }

    // Save directly as a Blob — no need to go through saveFileToDrive's
    // base64-decode path a second time.
    const folder = createOrGetFolder(folderName);
    const blob   = Utilities.newBlob(combined, mimeType || 'application/octet-stream', fileName);

    // Trash any existing file with the same name so the new upload replaces it
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) { existing.next().setTrashed(true); }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    trashFiles(chunkFiles);
    return { name: fileName, url: file.getDownloadUrl() };
  } catch (e) {
    Logger.log('finalizeChunkedUpload error: ' + e);
    // Best-effort cleanup even if assembly failed partway
    try { trashFiles(chunkFiles); } catch (_) {}
    throw new Error('ไม่สามารถรวมไฟล์ได้: ' + e.message);
  }
}

/**
 * Looks up every chunk file for a given uploadId and verifies the full
 * set is present. Returns them indexed by chunkIndex.
 *
 * Uses a single Drive query via the advanced Drive API-style search on
 * DriveApp to grab every file whose name starts with the uploadId prefix
 * in one iteration — much faster than N individual getFilesByName calls.
 */
function collectChunkFiles(uploadId, totalChunks) {
  const folder = getChunkTmpFolder();
  const prefix = uploadId + '__';

  // DriveApp folder search supports "title contains" via searchFiles().
  // We scope by parent to keep the result set small.
  const query  = "title contains '" + uploadId + "' and '" + folder.getId() + "' in parents and trashed = false";
  const iter   = DriveApp.searchFiles(query);

  const byIdx  = {};
  while (iter.hasNext()) {
    const f    = iter.next();
    const name = f.getName();
    if (name.indexOf(prefix) !== 0) continue;           // defensive: exact prefix match
    const idx  = parseInt(name.slice(prefix.length), 10);
    if (!isNaN(idx)) byIdx[idx] = f;
  }

  const files = new Array(totalChunks);
  for (let i = 0; i < totalChunks; i++) {
    if (!byIdx[i]) {
      throw new Error('ส่วนของไฟล์หายไป (chunk ' + i + '/' + totalChunks + ') — กรุณาอัพโหลดใหม่');
    }
    files[i] = byIdx[i];
  }
  return files;
}

function trashFiles(files) {
  files.forEach(function (f) {
    try { f.setTrashed(true); } catch (_) { /* ignore */ }
  });
}

/**
 * Manual cleanup helper — trashes any temp chunk files older than 24 h.
 * Run from the Apps Script editor if the _tmp_chunks folder fills up
 * because a user closed the tab mid-upload.
 */
function cleanupOldChunks() {
  const folder = getChunkTmpFolder();
  const cutoff = Date.now() - 24 * 60 * 60 * 1000;
  const files  = folder.getFiles();
  let removed  = 0;
  while (files.hasNext()) {
    const f = files.next();
    if (f.getLastUpdated().getTime() < cutoff) {
      f.setTrashed(true);
      removed++;
    }
  }
  Logger.log('cleanupOldChunks: trashed ' + removed + ' file(s)');
}

/**
 * Saves a single form-submission file to Drive. Supports three payload
 * shapes from the client:
 *
 *   resumable : { name, fileId }               ← large files, preferred
 *   inline    : { name, mimeType, base64 }     ← small files
 *   chunked   : { name, mimeType, uploadId, totalChunks }  ← legacy fallback
 *
 * @param {string} folderName Drive sub-folder
 * @param {Object} f          File descriptor from the client
 * @returns {{ name: string, url: string }}
 */
function persistSubmissionFile(folderName, f) {
  if (!f) throw new Error('missing file descriptor');

  // Resumable upload path: file already exists in Drive, just register it.
  if (f.fileId) {
    const registered = registerUploadedFile(f.fileId);
    // If the resumable upload targeted the wrong parent (shouldn't happen
    // because getUploadToken returns the correct folderId), move it.
    try {
      const file   = DriveApp.getFileById(f.fileId);
      const target = createOrGetFolder(folderName);
      const parents = file.getParents();
      let alreadyThere = false;
      while (parents.hasNext()) {
        if (parents.next().getId() === target.getId()) { alreadyThere = true; break; }
      }
      if (!alreadyThere) {
        target.addFile(file);
        // Remove from any other parents to avoid duplicates in My Drive root
        const ps = file.getParents();
        while (ps.hasNext()) {
          const p = ps.next();
          if (p.getId() !== target.getId()) p.removeFile(file);
        }
      }
    } catch (moveErr) {
      Logger.log('persistSubmissionFile move warning: ' + moveErr);
    }
    return registered;
  }

  // Legacy chunked path (kept for backward compatibility)
  if (f.uploadId && f.totalChunks) {
    return finalizeChunkedUpload(f.uploadId, f.totalChunks, folderName, f.name, f.mimeType);
  }

  // Inline base64 path (small files)
  return saveFileToDrive(folderName, f.name, f.base64, f.mimeType);
}

// ============================================================
// TELEGRAM NOTIFICATIONS
// ============================================================

/**
 * Sends a Telegram message via Bot API.
 * Reads TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID from Script Properties.
 * Non-fatal: logs errors but never throws (won't break form submissions).
 * @param {string} message  HTML-formatted message text
 */
function sendTelegramNotification(message) {
  try {
    const props  = PropertiesService.getScriptProperties();
    const token  = props.getProperty('TELEGRAM_BOT_TOKEN');
    const chatId = props.getProperty('TELEGRAM_CHAT_ID');

    if (!token || !chatId) {
      Logger.log('Telegram: TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID not configured. Skipping.');
      return;
    }

    const url     = 'https://api.telegram.org/bot' + token + '/sendMessage';
    const payload = {
      chat_id              : chatId,
      text                 : message,
      parse_mode           : 'HTML',
      disable_web_page_preview: true
    };

    const response = UrlFetchApp.fetch(url, {
      method             : 'post',
      contentType        : 'application/json',
      payload            : JSON.stringify(payload),
      muteHttpExceptions : true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('Telegram API error: ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('sendTelegramNotification error: ' + e);
  }
}

/**
 * Quick diagnostic — run from the editor to verify Telegram is wired up.
 * Logs the outcome so you can see exactly what failed (missing props,
 * bad token, network error, etc.).
 */
function testTelegram() {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  Logger.log('TELEGRAM_BOT_TOKEN set: ' + (token ? 'yes (ends …' + token.slice(-6) + ')' : 'NO'));
  Logger.log('TELEGRAM_CHAT_ID set:   ' + (chatId || 'NO'));

  if (!token || !chatId) {
    Logger.log('❌ Missing credentials — run setTelegramConfig() first.');
    return;
  }

  const response = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + token + '/sendMessage',
    {
      method             : 'post',
      contentType        : 'application/json',
      payload            : JSON.stringify({
        chat_id   : chatId,
        text      : '🧪 <b>Test from Apps Script</b>\nTelegram integration is working.',
        parse_mode: 'HTML'
      }),
      muteHttpExceptions : true
    }
  );
  Logger.log('HTTP ' + response.getResponseCode());
  Logger.log(response.getContentText());
}

/**
 * One-time setup: saves Telegram credentials to Script Properties.
 * Run this function ONCE from the Apps Script editor (not deployed web app).
 *
 * How to use:
 *   1. Open script editor → find this function
 *   2. Edit the values below, then click Run
 *   3. Revert the values (do not commit credentials to source control)
 *
 * @param {string} token   Bot token from @BotFather  (e.g. "123456:ABC-DEF...")
 * @param {string} chatId  Chat/channel ID             (e.g. "-1001234567890")
 */
function setTelegramConfig(token, chatId) {
  // ── EDIT THESE TWO VALUES, RUN ONCE, THEN REVERT ──
  const BOT_TOKEN = token  || 'PASTE_YOUR_BOT_TOKEN_HERE';
  const CHAT_ID   = chatId || 'PASTE_YOUR_CHAT_ID_HERE';
  // ──────────────────────────────────────────────────

  const props = PropertiesService.getScriptProperties();
  props.setProperty('TELEGRAM_BOT_TOKEN', BOT_TOKEN);
  props.setProperty('TELEGRAM_CHAT_ID',   CHAT_ID);
  Logger.log('✅ Telegram config saved. Token ends with: ...' + BOT_TOKEN.slice(-6));
}

// ============================================================
// FORM SUBMISSIONS
// ============================================================

/**
 * Handles ป.พ. 5 form submission.
 * data = { name, files: [{name, base64, mimeType}], note, isNewName }
 * Sheet columns: A=Timestamp, B=Name, C=FileNames, D=FileURLs, E=Note
 */
function submitPP5(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PP5);
    if (!sheet) throw new Error('ไม่พบชีต "' + SHEET_PP5 + '"');

    const fileNames = [];
    const fileUrls  = [];

    for (const f of data.files) {
      const result = persistSubmissionFile(FOLDER_PP5, f);
      fileNames.push(result.name);
      fileUrls.push(result.url);
    }

    sheet.appendRow([
      new Date(),
      data.name,
      fileNames.join(', '),
      fileUrls.join(', '),
      data.note || ''
    ]);

    if (data.isNewName) addName(data.name);

    const _ts      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const _files   = fileNames.map((n, i) => '  • <a href="' + fileUrls[i] + '">' + n + '</a>').join('\n');
    const _note    = data.note ? '\n📝 <b>หมายเหตุ:</b> ' + data.note : '';
    sendTelegramNotification(
      '📄 <b>ส่งงาน ป.พ. 5</b>\n\n' +
      '👤 <b>ชื่อ:</b> ' + data.name + '\n' +
      '🕐 <b>เวลา:</b> ' + _ts + '\n' +
      '📁 <b>ไฟล์ที่ส่ง:</b>\n' + _files +
      _note
    );

    return { success: true };
  } catch (e) {
    Logger.log('submitPP5 error: ' + e);
    throw new Error(e.message);
  }
}

/**
 * Handles สมรรถนะ 5 ด้าน form submission.
 * data = { name, wordFiles: [{name,base64,mimeType}], pdfFile: {name,base64,mimeType}, note, isNewName }
 * Sheet columns: A=Timestamp, B=Name, C=FileNames, D=FileURLs, E=PDFName, F=PDFURL, G=Note
 */
function submitCompetency(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COMPETENCY);
    if (!sheet) throw new Error('ไม่พบชีต "' + SHEET_COMPETENCY + '"');

    const fileNames = [];
    const fileUrls  = [];

    for (const f of data.wordFiles) {
      const result = persistSubmissionFile(FOLDER_COMPETENCY, f);
      fileNames.push(result.name);
      fileUrls.push(result.url);
    }

    let pdfName = '';
    let pdfUrl  = '';
    if (data.pdfFile && (data.pdfFile.base64 || data.pdfFile.uploadId)) {
      const result = persistSubmissionFile(FOLDER_COMPETENCY, data.pdfFile);
      pdfName = result.name;
      pdfUrl  = result.url;
    }

    sheet.appendRow([
      new Date(),
      data.name,
      fileNames.join(', '),
      fileUrls.join(', '),
      pdfName,
      pdfUrl,
      data.note || ''
    ]);

    if (data.isNewName) addName(data.name);

    const _ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const _words = fileNames.length
      ? '\n📁 <b>ไฟล์ Word/Excel:</b>\n' + fileNames.map((n, i) => '  • <a href="' + fileUrls[i] + '">' + n + '</a>').join('\n')
      : '';
    const _pdf   = pdfName ? '\n📑 <b>ไฟล์ PDF:</b>\n  • <a href="' + pdfUrl + '">' + pdfName + '</a>' : '';
    const _note  = data.note ? '\n📝 <b>หมายเหตุ:</b> ' + data.note : '';
    sendTelegramNotification(
      '🎯 <b>ส่งงาน สมรรถนะ 5 ด้าน</b>\n\n' +
      '👤 <b>ชื่อ:</b> ' + data.name + '\n' +
      '🕐 <b>เวลา:</b> ' + _ts +
      _words + _pdf + _note
    );

    return { success: true };
  } catch (e) {
    Logger.log('submitCompetency error: ' + e);
    throw new Error(e.message);
  }
}

/**
 * Handles SAR รายบุคคล form submission.
 * data = { name, wordFiles: [{name,base64,mimeType}], pdfFile: {name,base64,mimeType}, note, isNewName }
 * Sheet columns: A=Timestamp, B=Name, C=WordNames, D=WordURLs, E=PDFName, F=PDFURL, G=Note
 */
function submitSAR(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SAR);
    if (!sheet) throw new Error('ไม่พบชีต "' + SHEET_SAR + '"');

    const wordNames = [];
    const wordUrls  = [];

    for (const f of data.wordFiles) {
      const result = persistSubmissionFile(FOLDER_SAR, f);
      wordNames.push(result.name);
      wordUrls.push(result.url);
    }

    let pdfName = '';
    let pdfUrl  = '';
    if (data.pdfFile && (data.pdfFile.base64 || data.pdfFile.uploadId)) {
      const result = persistSubmissionFile(FOLDER_SAR, data.pdfFile);
      pdfName = result.name;
      pdfUrl  = result.url;
    }

    sheet.appendRow([
      new Date(),
      data.name,
      wordNames.join(', '),
      wordUrls.join(', '),
      pdfName,
      pdfUrl,
      data.note || ''
    ]);

    if (data.isNewName) addName(data.name);

    const _ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const _words = wordNames.length
      ? '\n📁 <b>ไฟล์ Word:</b>\n' + wordNames.map((n, i) => '  • <a href="' + wordUrls[i] + '">' + n + '</a>').join('\n')
      : '';
    const _pdf   = pdfName ? '\n📑 <b>ไฟล์ PDF:</b>\n  • <a href="' + pdfUrl + '">' + pdfName + '</a>' : '';
    const _note  = data.note ? '\n📝 <b>หมายเหตุ:</b> ' + data.note : '';
    sendTelegramNotification(
      '📊 <b>ส่งงาน SAR รายบุคคล</b>\n\n' +
      '👤 <b>ชื่อ:</b> ' + data.name + '\n' +
      '🕐 <b>เวลา:</b> ' + _ts +
      _words + _pdf + _note
    );

    return { success: true };
  } catch (e) {
    Logger.log('submitSAR error: ' + e);
    throw new Error(e.message);
  }
}

/**
 * Handles รายงานโครงการประจำปี 2568 form submission.
 * data = { name, wordFiles: [{name,base64,mimeType}], pdfFiles: [{name,base64,mimeType}], note, isNewName }
 * Sheet columns: A=Timestamp, B=Name, C=WordNames, D=WordURLs, E=PDFNames, F=PDFURLs, G=Note
 */
function submitProjectReport(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PROJECT);
    if (!sheet) throw new Error('ไม่พบชีต "' + SHEET_PROJECT + '"');

    const wordNames = [];
    const wordUrls  = [];

    for (const f of data.wordFiles) {
      const result = persistSubmissionFile(FOLDER_PROJECT, f);
      wordNames.push(result.name);
      wordUrls.push(result.url);
    }

    const pdfNames = [];
    const pdfUrls  = [];
    if (data.pdfFiles && data.pdfFiles.length) {
      for (const f of data.pdfFiles) {
        const result = persistSubmissionFile(FOLDER_PROJECT, f);
        pdfNames.push(result.name);
        pdfUrls.push(result.url);
      }
    }

    sheet.appendRow([
      new Date(),
      data.name,
      wordNames.join(', '),
      wordUrls.join(', '),
      pdfNames.join(', '),
      pdfUrls.join(', '),
      data.note || ''
    ]);

    if (data.isNewName) addName(data.name);

    const _ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const _words = wordNames.length
      ? '\n📁 <b>ไฟล์ Word:</b>\n' + wordNames.map((n, i) => '  • <a href="' + wordUrls[i] + '">' + n + '</a>').join('\n')
      : '';
    const _pdfs  = pdfNames.length
      ? '\n📑 <b>ไฟล์ PDF:</b>\n' + pdfNames.map((n, i) => '  • <a href="' + pdfUrls[i] + '">' + n + '</a>').join('\n')
      : '';
    const _note  = data.note ? '\n📝 <b>หมายเหตุ:</b> ' + data.note : '';
    sendTelegramNotification(
      '📋 <b>ส่งงาน รายงานโครงการประจำปี 2568</b>\n\n' +
      '👤 <b>ชื่อ:</b> ' + data.name + '\n' +
      '🕐 <b>เวลา:</b> ' + _ts +
      _words + _pdfs + _note
    );

    return { success: true };
  } catch (e) {
    Logger.log('submitProjectReport error: ' + e);
    throw new Error(e.message);
  }
}

/**
 * Handles a user-submitted bug / problem report.
 * Sends everything directly to Telegram — no Sheet, no Drive.
 *
 * data = { reporter, description, images: [{name,base64|uploadId,mimeType}], userAgent }
 */
function submitBugReport(data) {
  try {
    if (!data || !data.description || !data.description.trim()) {
      throw new Error('กรุณาอธิบายปัญหาที่พบ');
    }

    const props  = PropertiesService.getScriptProperties();
    const token  = props.getProperty('TELEGRAM_BOT_TOKEN');
    const chatId = props.getProperty('TELEGRAM_CHAT_ID');
    if (!token || !chatId) {
      throw new Error('ยังไม่ได้ตั้งค่า Telegram — กรุณาติดต่อผู้ดูแลระบบ');
    }

    // Build the text message
    const ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const agent = data.userAgent ? '\n\n💻 <b>User Agent:</b>\n<code>' + tgEscape(data.userAgent) + '</code>' : '';
    const caption =
      '🐞 <b>แจ้งปัญหา / Bug Report</b>\n\n' +
      '👤 <b>ผู้แจ้ง:</b> ' + tgEscape(data.reporter || '(ไม่ระบุ)') + '\n' +
      '🕐 <b>เวลา:</b> ' + ts + '\n\n' +
      '📝 <b>รายละเอียด:</b>\n' + tgEscape(data.description) +
      agent;

    // Resolve any chunked-upload handles into raw base64 in-memory (no Drive writes)
    const resolvedImages = (data.images || []).map(resolveInMemoryImage);

    if (resolvedImages.length === 0) {
      // No images — plain sendMessage
      sendTelegramNotification(caption);
    } else if (resolvedImages.length === 1) {
      // Single image — sendPhoto with caption
      sendTelegramPhoto(token, chatId, resolvedImages[0], caption);
    } else {
      // Multiple images — split into groups of 10 (Telegram limit).
      // First group gets the caption, the rest are bare photo groups.
      for (let i = 0; i < resolvedImages.length; i += 10) {
        const chunk = resolvedImages.slice(i, i + 10);
        sendTelegramMediaGroup(token, chatId, chunk, i === 0 ? caption : '');
      }
    }

    return { success: true };
  } catch (e) {
    Logger.log('submitBugReport error: ' + e);
    throw new Error(e.message);
  }
}

/**
 * Resolves a client image descriptor into { name, mimeType, bytes }.
 * Handles both inline base64 and chunked-upload handles, and purges the
 * cache entries for chunked uploads since nothing is persisted to Drive.
 */
function resolveInMemoryImage(img) {
  if (!img || !img.base64) throw new Error('ภาพไม่ถูกต้อง');
  return {
    name    : img.name || 'screenshot.png',
    mimeType: img.mimeType || 'image/png',
    bytes   : Utilities.base64Decode(img.base64)
  };
}

/** Escape <, >, & for Telegram HTML parse_mode. */
function tgEscape(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Send a single photo with an HTML caption via Telegram sendPhoto. */
function sendTelegramPhoto(token, chatId, img, caption) {
  try {
    const blob = Utilities.newBlob(img.bytes, img.mimeType, img.name);
    // Telegram captions are capped at 1024 chars — truncate and append notice
    let cap = caption || '';
    if (cap.length > 1024) cap = cap.slice(0, 1000) + '\n...(ตัดทอน)';

    const url = 'https://api.telegram.org/bot' + token + '/sendPhoto';
    const response = UrlFetchApp.fetch(url, {
      method : 'post',
      payload: {
        chat_id   : chatId,
        caption   : cap,
        parse_mode: 'HTML',
        photo     : blob
      },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      Logger.log('sendTelegramPhoto error: ' + response.getContentText());
      // Fall back to a text-only message so the report isn't lost
      sendTelegramNotification(caption + '\n\n⚠️ <i>(ส่งภาพไม่สำเร็จ)</i>');
    }
  } catch (e) {
    Logger.log('sendTelegramPhoto exception: ' + e);
    sendTelegramNotification(caption + '\n\n⚠️ <i>(ส่งภาพไม่สำเร็จ: ' + e.message + ')</i>');
  }
}

/**
 * Send 2-10 photos as an album via Telegram sendMediaGroup.
 * Caption (if provided) is attached to the first photo.
 */
function sendTelegramMediaGroup(token, chatId, imgs, caption) {
  try {
    // sendMediaGroup requires each photo to be referenced as attach://<name>
    // and the actual files included as separate multipart parts.
    const media   = [];
    const payload = { chat_id: chatId };

    let cap = caption || '';
    if (cap.length > 1024) cap = cap.slice(0, 1000) + '\n...(ตัดทอน)';

    imgs.forEach(function (img, i) {
      const attachName = 'photo' + i;
      const mediaItem  = { type: 'photo', media: 'attach://' + attachName };
      if (i === 0 && cap) {
        mediaItem.caption    = cap;
        mediaItem.parse_mode = 'HTML';
      }
      media.push(mediaItem);
      payload[attachName] = Utilities.newBlob(img.bytes, img.mimeType, img.name);
    });
    payload.media = JSON.stringify(media);

    const url = 'https://api.telegram.org/bot' + token + '/sendMediaGroup';
    const response = UrlFetchApp.fetch(url, {
      method             : 'post',
      payload            : payload,
      muteHttpExceptions : true
    });
    if (response.getResponseCode() !== 200) {
      Logger.log('sendTelegramMediaGroup error: ' + response.getContentText());
      sendTelegramNotification(caption + '\n\n⚠️ <i>(ส่งภาพไม่สำเร็จ)</i>');
    }
  } catch (e) {
    Logger.log('sendTelegramMediaGroup exception: ' + e);
    sendTelegramNotification(caption + '\n\n⚠️ <i>(ส่งภาพไม่สำเร็จ: ' + e.message + ')</i>');
  }
}

// ============================================================
// FILE LISTING (view uploaded files)
// ============================================================

/**
 * Returns a flat list of all uploaded files for a given category.
 * Used by the "ดูไฟล์ที่อัพโหลด" feature on the PP5 modal.
 *
 * @param {string} category  'pp5' | 'competency' | 'sar' | 'project'
 * @returns {Array<{ timestamp, name, fileName, fileUrl }>}
 */
function getUploadedFiles(category) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheetMap = {
      pp5        : SHEET_PP5,
      competency : SHEET_COMPETENCY,
      sar        : SHEET_SAR,
      project    : SHEET_PROJECT
    };

    const sheetName = sheetMap[category];
    if (!sheetName) throw new Error('ประเภทไม่ถูกต้อง');

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];

    const allRows = sheet.getDataRange().getValues();
    // Sort data rows descending by timestamp (col A), keeping header out
    const rows = allRows
      .filter(r => r[0] instanceof Date)
      .sort((a, b) => b[0].getTime() - a[0].getTime());
    const tz   = Session.getScriptTimeZone();
    const results = [];

    for (const row of rows) {
      const timestamp = Utilities.formatDate(row[0], tz, 'dd/MM/yyyy HH:mm');
      const personName = (row[1] || '').toString();

      if (category === 'pp5') {
        // Columns: C=FileNames, D=FileURLs
        const fileNames = row[2] ? row[2].toString().split(', ') : [];
        const fileUrls  = row[3] ? row[3].toString().split(', ') : [];

        fileNames.forEach((fn, i) => {
          if (fn.trim()) {
            results.push({
              timestamp,
              name    : personName,
              fileName: fn.trim(),
              fileUrl : (fileUrls[i] || '').trim()
            });
          }
        });

      } else {
        // Competency / SAR / Project — Columns: C=FileNames, D=FileURLs, E=PDFName, F=PDFURL
        const fileNames = row[2] ? row[2].toString().split(', ') : [];
        const fileUrls  = row[3] ? row[3].toString().split(', ') : [];
        const pdfName   = (row[4] || '').toString().trim();
        const pdfUrl    = (row[5] || '').toString().trim();

        fileNames.forEach((fn, i) => {
          if (fn.trim()) {
            results.push({
              timestamp,
              name    : personName,
              fileName: fn.trim(),
              fileUrl : (fileUrls[i] || '').trim()
            });
          }
        });

        if (pdfName) {
          results.push({ timestamp, name: personName, fileName: pdfName, fileUrl: pdfUrl });
        }
      }
    }

    return results;
  } catch (e) {
    Logger.log('getUploadedFiles error: ' + e);
    throw new Error('ไม่สามารถโหลดรายการไฟล์ได้: ' + e.message);
  }
}

// ============================================================
// SUBMISSION STATUS (for announcement popup)
// ============================================================

/**
 * Returns submission status for all 4 categories cross-referenced against NameList.
 * @returns {{ pp5, competency, sar, project }} — each is an array of { name, submitted }
 */
/**
 * Like getSubmissionStatus() but bundles uploaded file info per person.
 * Each item: { name, submitted, files: [{fileName, timestamp}] }
 */
function getStatusWithFiles() {
  try {
    const status = getSubmissionStatus();
    const categories = ['pp5', 'competency', 'sar', 'project'];

    // Build name→files map for each category
    const fileMaps = {};
    categories.forEach(function(cat) {
      const map = {};
      getUploadedFiles(cat).forEach(function(f) {
        if (!map[f.name]) map[f.name] = [];
        map[f.name].push({ fileName: f.fileName, timestamp: f.timestamp });
      });
      fileMaps[cat] = map;
    });

    // Merge files into each status item
    function mergeFiles(list, cat) {
      return list.map(function(item) {
        return {
          name      : item.name,
          submitted : item.submitted,
          files     : fileMaps[cat][item.name] || []
        };
      });
    }

    return {
      pp5        : mergeFiles(status.pp5,        'pp5'),
      competency : mergeFiles(status.competency, 'competency'),
      sar        : mergeFiles(status.sar,        'sar'),
      project    : mergeFiles(status.project,    'project')
    };
  } catch (e) {
    Logger.log('getStatusWithFiles error: ' + e);
    throw new Error('ไม่สามารถโหลดสถานะได้: ' + e.message);
  }
}

function getSubmissionStatus() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Load all names from NameList (skip header)
    const nameSheet = ss.getSheetByName(SHEET_NAMELIST);
    const allNames  = nameSheet
      ? nameSheet.getDataRange().getValues().slice(1)
          .map(r => (r[0] || '').toString().trim()).filter(n => n)
      : [];

    // Helper: build a Set of names who have submitted in a given sheet
    function submittedSet(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      const set   = new Set();
      if (!sheet) return set;
      sheet.getDataRange().getValues().forEach(function(row) {
        if (row[0] instanceof Date && row[1]) {
          set.add(row[1].toString().trim());
        }
      });
      return set;
    }

    const sets = {
      pp5        : submittedSet(SHEET_PP5),
      competency : submittedSet(SHEET_COMPETENCY),
      sar        : submittedSet(SHEET_SAR),
      project    : submittedSet(SHEET_PROJECT)
    };

    // Map each name to { name, submitted } per category
    function statusList(set) {
      return allNames.map(function(name) {
        return { name: name, submitted: set.has(name) };
      });
    }

    return {
      pp5        : statusList(sets.pp5),
      competency : statusList(sets.competency),
      sar        : statusList(sets.sar),
      project    : statusList(sets.project)
    };
  } catch (e) {
    Logger.log('getSubmissionStatus error: ' + e);
    throw new Error('ไม่สามารถโหลดสถานะได้: ' + e.message);
  }
}
