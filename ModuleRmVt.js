/**
 * ══════════════════════════════════════════════
 * ModuleRmVt.gs — Aktualizace RM + VT
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 * ══════════════════════════════════════════════
 */

const ModuleRmVt = {
  CONFIG_FOLDER_KEY: 'rmVtFolderId',
  LOCAL_SHEET_NAME: 'Filiálky',
  SOURCE_SHEET_NAME: 'Filiálky',
  SUPPORTED_MIME_TYPES: [
    MimeType.GOOGLE_SHEETS,
    MimeType.MICROSOFT_EXCEL,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  ],

  /**
   * @returns {{folderId: string}}
   */
  getConfig() {
    return {
      folderId: String(AppConfig.get(ModuleRmVt.CONFIG_FOLDER_KEY) || '')
    };
  },

  /**
   * @param {string} rawFolderId
   * @returns {{success: boolean, folderId: string, folderName: string}}
   */
  saveFolder(rawFolderId) {
    const folderId = ModuleRmVt.normalizeFolderId_(rawFolderId);
    if (!folderId) throw new Error('Zadejte ID nebo URL složky.');

    const folder = DriveApp.getFolderById(folderId);
    AppConfig.set(ModuleRmVt.CONFIG_FOLDER_KEY, folderId);
    return { success: true, folderId: folderId, folderName: folder.getName() };
  },

  /**
   * Najde zdroj, porovná data, aktualizuje RM/VT/LC v lokálním listu a spustí obarvení skupin RM.
   * @param {string} rawFolderId
   * @returns {Object}
   */
  run(rawFolderId) {
    const folderId = ModuleRmVt.normalizeFolderId_(rawFolderId || AppConfig.get(ModuleRmVt.CONFIG_FOLDER_KEY));
    if (!folderId) throw new Error('Nejdřív zadejte ID složky se zdrojovým souborem.');

    ModuleRmVt.saveFolder(folderId);

    const sourceFile = ModuleRmVt.findSourceFile_(folderId);
    const sourceData = ModuleRmVt.readSourceData_(sourceFile);
    const localData = ModuleRmVt.readLocalData_();
    const comparison = ModuleRmVt.compareAndUpdate_(localData, sourceData);
    const coloring = ModuleManagerColoring.run();

    return {
      success: true,
      sourceFile: sourceFile,
      summary: comparison.summary,
      changes: comparison.changes,
      lcCounts: comparison.lcCounts,
      coloring: coloring.results
    };
  },

  /**
   * @param {string} folderId
   * @returns {{id: string, name: string, mimeType: string, modifiedTime: string}}
   */
  findSourceFile_(folderId) {
    const escapedFolderId = String(folderId).replace(/'/g, "\\'");
    const q = "'" + escapedFolderId + "' in parents and trashed = false and (" +
      "mimeType = '" + MimeType.GOOGLE_SHEETS + "' or " +
      "mimeType = '" + MimeType.MICROSOFT_EXCEL + "' or " +
      "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')";

    const res = Drive.Files.list({
      q: q,
      orderBy: 'modifiedTime desc',
      pageSize: 1,
      fields: 'files(id,name,mimeType,modifiedTime)'
    });

    if (!res.files || res.files.length === 0) {
      throw new Error('Ve složce nebyl nalezen žádný podporovaný zdrojový soubor Google Sheets/XLSX.');
    }
    return res.files[0];
  },

  /**
   * @param {{id: string, name: string, mimeType: string}} file
   * @returns {{rowsByBranch: Object, lcCounts: Object, rowCount: number}}
   */
  readSourceData_(file) {
    let spreadsheetId = file.id;
    let tempFileId = null;

    try {
      if (file.mimeType !== MimeType.GOOGLE_SHEETS) {
        const converted = Drive.Files.copy({
          name: '_tmp_RMVT_' + file.name,
          mimeType: MimeType.GOOGLE_SHEETS
        }, file.id);
        spreadsheetId = converted.id;
        tempFileId = converted.id;
      }

      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheet = ss.getSheetByName(ModuleRmVt.SOURCE_SHEET_NAME) || ss.getSheets()[0];
      const values = sheet.getDataRange().getDisplayValues();
      return ModuleRmVt.buildSourceIndex_(values);
    } finally {
      if (tempFileId) {
        try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e) { }
      }
    }
  },

  /**
   * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, rowsByBranch: Object, lcCounts: Object, rowCount: number}}
   */
  readLocalData_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ModuleRmVt.LOCAL_SHEET_NAME);
    if (!sheet) throw new Error('Lokální list "' + ModuleRmVt.LOCAL_SHEET_NAME + '" nebyl nalezen.');

    const values = sheet.getDataRange().getDisplayValues();
    const rowsByBranch = {};
    const lcCounts = {};

    for (let r = 1; r < values.length; r++) {
      const branchId = ModuleRmVt.normalizeBranchId_(values[r][0]);
      if (!branchId) continue;

      const lc = ModuleRmVt.clean_(values[r][4]);
      rowsByBranch[branchId] = {
        row: r + 1,
        branchId: branchId,
        lc: lc,
        rm: ModuleRmVt.clean_(values[r][12]),
        vt: ModuleRmVt.clean_(values[r][13])
      };
      ModuleRmVt.increment_(lcCounts, lc);
    }

    return { sheet: sheet, rowsByBranch: rowsByBranch, lcCounts: lcCounts, rowCount: Object.keys(rowsByBranch).length };
  },

  /**
   * @param {Array<Array<string>>} values
   * @returns {{rowsByBranch: Object, lcCounts: Object, rowCount: number}}
   */
  buildSourceIndex_(values) {
    const rowsByBranch = {};
    const lcCounts = {};

    for (let r = 1; r < values.length; r++) {
      const branchId = ModuleRmVt.normalizeBranchId_(values[r][0]);
      if (!branchId) continue;

      const lc = ModuleRmVt.clean_(values[r][2]);
      rowsByBranch[branchId] = {
        branchId: branchId,
        lc: lc,
        rm: ModuleRmVt.clean_(values[r][5]),
        vt: ModuleRmVt.clean_(values[r][4])
      };
      ModuleRmVt.increment_(lcCounts, lc);
    }

    return { rowsByBranch: rowsByBranch, lcCounts: lcCounts, rowCount: Object.keys(rowsByBranch).length };
  },

  /**
   * @param {Object} localData
   * @param {Object} sourceData
   * @returns {{summary: Object, changes: Array, lcCounts: Array}}
   */
  compareAndUpdate_(localData, sourceData) {
    const changes = [];
    const updates = [];
    let lcChanges = 0;
    let rmChanges = 0;
    let vtChanges = 0;
    let missingInSource = 0;
    let missingInLocal = 0;

    Object.keys(localData.rowsByBranch).forEach(branchId => {
      const local = localData.rowsByBranch[branchId];
      const source = sourceData.rowsByBranch[branchId];
      if (!source) {
        missingInSource++;
        changes.push({ type: 'missingSource', branchId: branchId, message: 'Filiálka chybí ve zdroji.' });
        return;
      }

      const rowUpdate = { row: local.row, lc: null, rm: null, vt: null };
      if (local.lc !== source.lc) {
        lcChanges++;
        rowUpdate.lc = source.lc;
        changes.push({ type: 'lc', branchId: branchId, from: local.lc, to: source.lc });
      }
      if (local.rm !== source.rm) {
        rmChanges++;
        rowUpdate.rm = source.rm;
        changes.push({ type: 'rm', branchId: branchId, from: local.rm, to: source.rm });
      }
      if (local.vt !== source.vt) {
        vtChanges++;
        rowUpdate.vt = source.vt;
        changes.push({ type: 'vt', branchId: branchId, from: local.vt, to: source.vt });
      }

      if (rowUpdate.lc !== null || rowUpdate.rm !== null || rowUpdate.vt !== null) {
        updates.push(rowUpdate);
      }
    });

    Object.keys(sourceData.rowsByBranch).forEach(branchId => {
      if (!localData.rowsByBranch[branchId]) {
        missingInLocal++;
        changes.push({ type: 'missingLocal', branchId: branchId, message: 'Filiálka chybí v lokálním listu.' });
      }
    });

    updates.forEach(update => {
      if (update.lc !== null) localData.sheet.getRange(update.row, 5).setValue(update.lc);
      if (update.rm !== null) localData.sheet.getRange(update.row, 13).setValue(update.rm);
      if (update.vt !== null) localData.sheet.getRange(update.row, 14).setValue(update.vt);
    });

    const lcCounts = ModuleRmVt.compareLcCounts_(localData.lcCounts, sourceData.lcCounts);
    return {
      summary: {
        localBranches: localData.rowCount,
        sourceBranches: sourceData.rowCount,
        updatedRows: updates.length,
        lcChanges: lcChanges,
        rmChanges: rmChanges,
        vtChanges: vtChanges,
        missingInSource: missingInSource,
        missingInLocal: missingInLocal
      },
      changes: changes,
      lcCounts: lcCounts
    };
  },

  compareLcCounts_(localCounts, sourceCounts) {
    const keys = {};
    Object.keys(localCounts).forEach(k => keys[k] = true);
    Object.keys(sourceCounts).forEach(k => keys[k] = true);

    return Object.keys(keys).sort().map(lc => ({
      lc: lc || '(bez LC)',
      local: localCounts[lc] || 0,
      source: sourceCounts[lc] || 0,
      delta: (sourceCounts[lc] || 0) - (localCounts[lc] || 0)
    })).filter(row => row.delta !== 0);
  },

  normalizeFolderId_(value) {
    return Utils.extractFileIdFromUrl(String(value || '').trim()) || String(value || '').trim();
  },

  normalizeBranchId_(value) {
    const s = String(value || '').trim();
    if (!s) return '';
    return s.replace(/\.0$/, '');
  },

  clean_(value) {
    return String(value || '').trim();
  },

  increment_(obj, key) {
    obj[key] = (obj[key] || 0) + 1;
  }
};

function showRmVtModal() {
  const html = HtmlService
    .createTemplateFromFile('Modal_RmVt')
    .evaluate()
    .setWidth(920)
    .setHeight(680);

  SpreadsheetApp.getUi().showModelessDialog(html, ' ');
}

function moduleRmVt_getConfig() {
  return ModuleRmVt.getConfig();
}

function moduleRmVt_saveFolder(folderId) {
  return ModuleRmVt.saveFolder(folderId);
}

function moduleRmVt_run(folderId) {
  return ModuleRmVt.run(folderId);
}
