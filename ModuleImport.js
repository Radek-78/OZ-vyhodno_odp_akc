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

function moduleImport_runDrive(fileIds, options) {
  return ModuleImport.importToDrive(fileIds, options);
}

