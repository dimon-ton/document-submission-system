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
    const file    = folder.createFile(blob);

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
      const result = saveFileToDrive(FOLDER_PP5, f.name, f.base64, f.mimeType);
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
      const result = saveFileToDrive(FOLDER_COMPETENCY, f.name, f.base64, f.mimeType);
      fileNames.push(result.name);
      fileUrls.push(result.url);
    }

    let pdfName = '';
    let pdfUrl  = '';
    if (data.pdfFile && data.pdfFile.base64) {
      const result = saveFileToDrive(FOLDER_COMPETENCY, data.pdfFile.name, data.pdfFile.base64, data.pdfFile.mimeType);
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
      const result = saveFileToDrive(FOLDER_SAR, f.name, f.base64, f.mimeType);
      wordNames.push(result.name);
      wordUrls.push(result.url);
    }

    let pdfName = '';
    let pdfUrl  = '';
    if (data.pdfFile && data.pdfFile.base64) {
      const result = saveFileToDrive(FOLDER_SAR, data.pdfFile.name, data.pdfFile.base64, data.pdfFile.mimeType);
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
      const result = saveFileToDrive(FOLDER_PROJECT, f.name, f.base64, f.mimeType);
      wordNames.push(result.name);
      wordUrls.push(result.url);
    }

    const pdfNames = [];
    const pdfUrls  = [];
    if (data.pdfFiles && data.pdfFiles.length) {
      for (const f of data.pdfFiles) {
        const result = saveFileToDrive(FOLDER_PROJECT, f.name, f.base64, f.mimeType);
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
    return { success: true };
  } catch (e) {
    Logger.log('submitProjectReport error: ' + e);
    throw new Error(e.message);
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

    const rows = sheet.getDataRange().getValues();
    const tz   = Session.getScriptTimeZone();
    const results = [];

    for (const row of rows) {
      // Skip header row and empty rows — data rows always have a Date in column A
      if (!(row[0] instanceof Date)) continue;

      const timestamp = Utilities.formatDate(row[0], tz, 'dd/MM/yyyy HH:mm');
      const personName = (row[1] || '').toString();

      if (category === 'pp5' || category === 'competency') {
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
        // SAR / Project — Columns: C=WordNames, D=WordURLs, E=PDFName, F=PDFURL
        const wordNames = row[2] ? row[2].toString().split(', ') : [];
        const wordUrls  = row[3] ? row[3].toString().split(', ') : [];
        const pdfName   = (row[4] || '').toString().trim();
        const pdfUrl    = (row[5] || '').toString().trim();

        wordNames.forEach((fn, i) => {
          if (fn.trim()) {
            results.push({
              timestamp,
              name    : personName,
              fileName: fn.trim(),
              fileUrl : (wordUrls[i] || '').trim()
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
