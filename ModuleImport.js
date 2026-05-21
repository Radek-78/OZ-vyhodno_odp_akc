/**
 * ══════════════════════════════════════════════
 * ModuleImport.gs — Modul: Import souborů
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * KOMPLETNÍ REWORK PRO MAXIMÁLNÍ STABILITU
 * ══════════════════════════════════════════════
 */

// Regex kompilovaný jednou pro celý modul — ne uvnitř map() kde by se re-kompiloval per buňce
const _ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;

const ModuleImport = {

  /**
   * Import datové dávky do listu
   * @param {string} jsonPayload - JSON řetězec s 2D polem dat
   * @param {string} sheetName - název cílového listu
   * @param {Object} options - { isFirstChunk: boolean, overwrite: boolean }
   */
  importChunk(jsonPayload, sheetName, options) {
    let data;
    try {
      data = JSON.parse(jsonPayload);

      // Detekce a převod ISO dat zpět na objekty Date
      data = data.map(row => row.map(val => {
        if (typeof val === 'string' && _ISO_DATE_RE.test(val)) {
          const d = new Date(val);
          return isNaN(d.getTime()) ? val : d;
        }
        return val;
      }));
    } catch (e) {
      return { success: false, error: 'Chyba při čtení datového balíčku: ' + e.message };
    }

    if (!data || data.length === 0) return { success: true, count: 0 };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    // 1. Příprava listu (pouze u první dávky souboru)
    if (options.isFirstChunk) {
      if (sheet) {
        if (options.overwrite) {
          AppLogger.info('Mazání listu: ' + sheetName);
          sheet.clear();
        } else {
          AppLogger.info('Přidávám k existujícím datům: ' + sheetName);
        }
      } else {
        AppLogger.info('Vytvářím nový list: ' + sheetName);
        sheet = ss.insertSheet(sheetName);
      }
    }

    if (!sheet) throw new Error('Nepodařilo se získat cílový list: ' + sheetName);

    // 2. Zjištění startovní pozice a rozměrů
    let startRow = (options.startRow !== undefined) ? options.startRow : (sheet.getLastRow() + 1);
    const numRows = data.length;
    const maxCols = data.reduce((max, row) => Math.max(max, row.length), 0);
    if (maxCols === 0) return { success: true, count: 0, actualStartRow: startRow };

    // AUTOMATICKÉ ROZŠÍŘENÍ LISTU
    const maxRows = sheet.getMaxRows();
    const neededRows = startRow + numRows - 1;
    const totalToPrepare = (options.isFirstChunk && options.totalRows) ? (options.totalRows) : neededRows;

    if (totalToPrepare > maxRows) {
      const toAdd = totalToPrepare - maxRows;
      sheet.insertRowsAfter(maxRows, toAdd);
      AppLogger.info('Příprava prostoru: +' + toAdd + ' řádků.');
    }

    // 3. PŘÍPRAVA DAT PRO ADVANCED SHEETS API (v4)
    const normalizedData = data.map(row => {
      const r = row.slice(0, maxCols);
      while (r.length < maxCols) r.push('');
      return r;
    });

    try {
      const spreadsheetId = ss.getId();
      const range = sheetName + '!' + Utils.columnToLetter(1) + startRow + ':' + Utils.columnToLetter(maxCols) + (startRow + numRows - 1);

      const valueRange = Sheets.newValueRange();
      valueRange.values = normalizedData;

      Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, range, { valueInputOption: 'RAW' });

      // Formátování záhlaví (podmíněno parametrem formatHeader)
      if (startRow === 1 && options.formatHeader !== false) {
        const header = sheet.getRange(1, 1, 1, maxCols);
        header.setFontWeight('bold');
        header.setBackground('#eeeeee');
        sheet.setFrozenRows(1);
      }

      return { success: true, count: numRows, actualStartRow: startRow };
    } catch (e) {
      AppLogger.error('Chyba V4 (fallback): ' + e.message);
      try {
        sheet.getRange(startRow, 1, numRows, maxCols).setValues(normalizedData);
        return { success: true, count: numRows, actualStartRow: startRow };
      } catch (err) {
        return { success: false, error: 'Kritické selhání zápisu: ' + err.message };
      }
    }
  },

  /**
   * Import souborů z Google Drive
   */
  importToDrive(fileIds, options) {
    const results = [];
    const targetFolderId = options.targetFolderId;
    const prefix = options.prefix || '';

    let targetFolder = targetFolderId ? DriveApp.getFolderById(targetFolderId) : DriveApp.getRootFolder();

    fileIds.forEach(id => {
      try {
        const file = DriveApp.getFileById(id);
        const name = prefix + file.getName();
        file.makeCopy(name, targetFolder);
        results.push({ name: name, success: true });
        AppLogger.ok('Kopírováno: ' + name);
      } catch (e) {
        results.push({ name: id, success: false, error: e.message });
      }
    });
    return { success: true, results: results };
  },

  /**
   * Nahrání souboru na Drive
   */
  uploadToDrive(fileObj, options) {
    try {
      const blob = ModuleImport.base64ToBlob_(fileObj.data, fileObj.name);
      const folder = options.targetFolderId ? DriveApp.getFolderById(options.targetFolderId) : DriveApp.getRootFolder();
      const file = folder.createFile(blob);
      if (options.prefix) file.setName(options.prefix + file.getName());
      return { success: true, name: file.getName() };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  /**
   * Rychlý import velkého XLSX/CSV souboru:
   * - soubor se nahraje jako dočasný Google Sheet
   * - hodnoty z vybraného listu se zapíšou do stávajícího cílového listu
   * - cílový list se nemaže, aby zůstaly funkční vzorce navázané na jeho sheetId
   */
  importBySheetCopy(fileObj, targetSheetName, options) {
    const tempName = '_tmp_import_' + Date.now() + '_' + (fileObj.name || 'soubor');
    let tempFileId = '';

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const blob = ModuleImport.base64ToBlob_(fileObj.data, fileObj.name);
      const tempFileResource = {
        name: tempName,
        mimeType: MimeType.GOOGLE_SHEETS
      };
      const parentFolderId = ModuleImport.getActiveSpreadsheetParentFolderId_(ss.getId());
      if (parentFolderId) tempFileResource.parents = [parentFolderId];

      const created = Drive.Files.create(tempFileResource, blob, {
        fields: 'id,name'
      });

      tempFileId = created.id;
      const sourceSs = ModuleImport.openSpreadsheetWithRetry_(tempFileId);
      const sourceSheets = sourceSs.getSheets();
      if (sourceSheets.length === 0) throw new Error('Dočasný soubor neobsahuje žádný list.');

      const requestedIndex = Math.max(0, Number(options && options.sourceSheetIndex) || 0);
      const sourceSheet = sourceSheets[requestedIndex] || sourceSheets[0];
      const importedSheet = ModuleImport.getOrCreateTargetSheet_(ss, targetSheetName);
      const result = ModuleImport.replaceSheetValuesFromSource_(sourceSheet, importedSheet, options);

      return {
        success: true,
        targetSheet: targetSheetName,
        sourceSheet: sourceSheet.getName(),
        rows: result.rows,
        columns: result.columns,
        replacedDashCount: result.replacedDashCount
      };
    } catch (e) {
      return { success: false, error: 'Rychlý import selhal: ' + e.message };
    } finally {
      if (tempFileId) {
        try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (err) { }
      }
    }
  },

  getOrCreateTargetSheet_(ss, sheetName) {
    return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  },

  replaceSheetValuesFromSource_(sourceSheet, targetSheet, options) {
    const sourceRange = sourceSheet.getDataRange();
    let values = sourceRange.getValues();
    const numRows = values.length;
    const numCols = values.reduce(function (max, row) {
      return Math.max(max, row.length);
    }, 0);

    let replacedDashCount = 0;
    if (options && options.replaceDashWithZero) {
      values = values.map(function (row) {
        return row.map(function (value) {
          if (typeof value === 'string' && value.trim() === '--') {
            replacedDashCount++;
            return 0;
          }
          return value;
        });
      });
    }

    targetSheet.clear();
    if (numRows === 0 || numCols === 0) {
      SpreadsheetApp.flush();
      return { rows: 0, columns: 0, replacedDashCount: replacedDashCount };
    }

    ModuleImport.ensureSheetSize_(targetSheet, numRows, numCols);

    const batchSize = 10000;
    for (let start = 0; start < numRows; start += batchSize) {
      const batch = values.slice(start, start + batchSize).map(function (row) {
        const normalized = row.slice(0, numCols);
        while (normalized.length < numCols) normalized.push('');
        return normalized;
      });
      targetSheet.getRange(start + 1, 1, batch.length, numCols).setValues(batch);
    }

    SpreadsheetApp.flush();
    return { rows: numRows, columns: numCols, replacedDashCount: replacedDashCount };
  },

  ensureSheetSize_(sheet, numRows, numCols) {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    if (numRows > maxRows) sheet.insertRowsAfter(maxRows, numRows - maxRows);
    if (numCols > maxCols) sheet.insertColumnsAfter(maxCols, numCols - maxCols);
  },

  openSpreadsheetWithRetry_(spreadsheetId) {
    let lastError = null;
    for (let i = 0; i < 8; i++) {
      try {
        return SpreadsheetApp.openById(spreadsheetId);
      } catch (e) {
        lastError = e;
        Utilities.sleep(750);
      }
    }
    throw lastError || new Error('Dočasnou tabulku se nepodařilo otevřít.');
  },

  getActiveSpreadsheetParentFolderId_(spreadsheetId) {
    try {
      const parents = DriveApp.getFileById(spreadsheetId).getParents();
      if (parents.hasNext()) return parents.next().getId();
    } catch (e) {
      AppLogger.warn('Nepodařilo se zjistit složku aktuální tabulky: ' + e.message);
    }
    return '';
  },

  base64ToBlob_(dataUrl, fileName) {
    const parts = dataUrl.split(',');
    const mime = parts[0].match(/:(.*?);/)[1];
    const decoded = Utilities.base64Decode(parts[1]);
    return Utilities.newBlob(decoded, mime, fileName);
  }
};

/** API Wrappery */
function moduleImport_chunk(data, sheetName, options) {
  return ModuleImport.importChunk(data, sheetName, options);
}

function moduleImport_uploadToDrive(fileObj, options) {
  return ModuleImport.uploadToDrive(fileObj, options);
}

function moduleImport_fastSheetCopy(fileObj, sheetName, options) {
  return ModuleImport.importBySheetCopy(fileObj, sheetName, options);
}

function moduleImport_runDrive(fileIds, options) {
  return ModuleImport.importToDrive(fileIds, options);
}

