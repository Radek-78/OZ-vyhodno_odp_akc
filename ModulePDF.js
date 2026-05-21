/**
 * ══════════════════════════════════════════════
 * ModulePDF.gs — Modul: Generování PDF
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Export vybraných listů do PDF.
 * Nastavení layoutu, uložení na Google Drive.
 * ══════════════════════════════════════════════
 */

const ModulePDF = {

  /**
   * Vygeneruje samostatné PDF pro každý vybraný list.
   * Listy exportuje po jednom, dočasně ponechá viditelný jen právě generovaný list.
   *
   * @param {string[]} sheetNames
   * @param {Object} options
   * @param {string} options.rangeId
   * @param {string} options.size
   * @param {boolean} options.portrait
   * @param {boolean} options.gridlines
   * @param {boolean} options.pageNumbers
   * @returns {{success: boolean, count?: number, files?: Array, error?: string}}
   */
  exportSheets(sheetNames, options) {
    AppLogger.info('Spouštím generování PDF pro ' + sheetNames.length + ' listů.');

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const folderResult = ModulePDF.resolveCurrentFileFolder_();
      if (!folderResult.success) return { success: false, error: folderResult.error };
      const folder = folderResult.folder;

      const allSheets = ss.getSheets();
      const originalHidden = {};
      const originalFrozen = {};
      allSheets.forEach(sheet => {
        originalHidden[sheet.getSheetId()] = sheet.isSheetHidden();
        originalFrozen[sheet.getSheetId()] = {
          rows: sheet.getFrozenRows(),
          columns: sheet.getFrozenColumns()
        };
      });

      const originalActive = ss.getActiveSheet();
      const selectedSheets = sheetNames
        .map(name => ss.getSheetByName(name))
        .filter(sheet => sheet);

      if (selectedSheets.length === 0) {
        return { success: false, error: 'Nebyl nalezen žádný vybraný list.' };
      }

      const token = ScriptApp.getOAuthToken();
      const ssId = ss.getId();
      const rangeId = (options.rangeId || 'B2:M10').trim();
      const files = [];

      try {
        for (let i = 0; i < selectedSheets.length; i++) {
          const sheet = selectedSheets[i];
          AppLogger.info('[' + (i + 1) + '/' + selectedSheets.length + '] Generuji list ' + sheet.getName());

          sheet.showSheet();
          ss.setActiveSheet(sheet);
          sheet.setFrozenRows(0);
          sheet.setFrozenColumns(0);
          allSheets.forEach(other => {
            if (other.getSheetId() !== sheet.getSheetId() && !other.isSheetHidden()) {
              other.hideSheet();
            }
          });
          SpreadsheetApp.flush();

          const range = sheet.getRange(rangeId);
          const r1 = range.getRow() - 1;
          const c1 = range.getColumn() - 1;
          const r2 = range.getLastRow();
          const c2 = range.getLastColumn();

          const urlOptions =
            'exportFormat=pdf&format=pdf' +
            '&size=' + (options.size === 'A3' ? 'A3' : 'A4') +
            '&portrait=' + (options.portrait ? 'true' : 'false') +
            '&fitw=true' +
            '&horizontal_alignment=CENTER' +
            '&vertical_alignment=TOP' +
            '&gridlines=' + (options.gridlines ? 'true' : 'false') +
            '&printtitle=false' +
            '&sheetnames=false' +
            '&pagenumbers=' + (options.pageNumbers ? 'true' : 'false') +
            '&attachment=true' +
            '&gid=' + sheet.getSheetId() +
            '&r1=' + r1 + '&c1=' + c1 + '&r2=' + r2 + '&c2=' + c2;

          const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' + urlOptions;
          const response = UrlFetchApp.fetch(url, {
            headers: { Authorization: 'Bearer ' + token },
            muteHttpExceptions: true
          });

          if (response.getResponseCode() !== 200) {
            throw new Error('Generování listu "' + sheet.getName() + '" selhalo: ' + response.getContentText().substring(0, 200));
          }

          const fileName = ModulePDF.buildSheetPdfName_(ss.getName(), sheet.getName());
          const file = folder.createFile(response.getBlob().setName(fileName));
          files.push({ name: fileName, id: file.getId() });

          Utilities.sleep(200);
        }
      } finally {
        ModulePDF.restoreSheetState_(ss, allSheets, originalHidden, originalFrozen, originalActive);
      }

      AppLogger.ok('PDF generování dokončeno: ' + files.length + ' souborů.');
      return { success: true, count: files.length, files: files };
    } catch (e) {
      AppLogger.error('Chyba PDF generování: ' + e.message);
      return { success: false, error: e.message };
    }
  },

  /**
   * Exportuje zadané listy do jednoho PDF a uloží jej na disk Google.
   *
   * @param {string[]} sheetNames - Pole názvů listů, které se mají exportovat.
   * @param {Object} options - Konfigurační objekt.
   * @param {string} options.size - Velikost papíru (A4, A3).
   * @param {boolean} options.portrait - Orientace: true = na výšku, false = na šířku.
   * @param {boolean} options.gridlines - Zda zobrazit mřížku.
   * @param {boolean} options.pageNumbers - Zda zobrazit čísla stránek.
   * @param {string} [options.fileName] - Volitelný vlastní název souboru.
   * @returns {{success: boolean, fileId?: string, error?: string}}
   */
  export(sheetNames, options) {
    AppLogger.info('Spouštím PDF export pro ' + sheetNames.length + ' listů.');

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Nalezení nebo určení cílové složky (sdílená logika)
      const folderResult = ModulePDF.resolveTargetFolder_();
      if (!folderResult.success) return { success: false, error: folderResult.error };
      const folder = folderResult.folder;

      // Nastavení názvu výstupního souboru
      let fileName = options.fileName;
      if (!fileName) {
        let ts = Utils.formatDate(new Date(), 'yyyy-MM-dd_HH-mm-ss');
        let prefix = ss.getName().replace(/[\/\\]/g, '-');
        fileName = prefix + '_' + ts + '.pdf';
      }
      if (!fileName.toLowerCase().endsWith('.pdf')) {
        fileName += '.pdf';
      }

      // Určení formátu URL pro export
      const urlBase = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';
      let exportOptions =
        'exportFormat=pdf&format=pdf' +
        '&size=' + (options.size === 'A3' ? 'A3' : 'A4') +
        '&portrait=' + (options.portrait ? 'true' : 'false') +
        '&gridlines=' + (options.gridlines ? 'true' : 'false') +
        '&pagenumbers=' + (options.pageNumbers ? 'true' : 'false') +
        '&fitw=true' + // vždy fit to width
        '&sheetnames=false&printtitle=false' +
        '&attachment=true';

      // Identifikace GIDs požadovaných listů
      let gids = [];
      const allSheets = ss.getSheets();
      for (let s of allSheets) {
        if (sheetNames.includes(s.getName())) {
          gids.push(s.getSheetId());
        }
      }

      if (gids.length === 0) {
        return { success: false, error: 'Žádný ze zadaných listů nebyl v dokumentu nalezen.' };
      }

      AppLogger.info('Vytvářím dočasnou kopii pro seřazení listů a tisk...');
      const tempFile = DriveApp.getFileById(ss.getId()).makeCopy('tmp_export_' + Date.now());

      // try/finally zajistí smazání dočasného souboru i při chybě
      try {
        const tempSs = SpreadsheetApp.openById(tempFile.getId());

        let listToDelete = [];
        for (let s of tempSs.getSheets()) {
          const sName = s.getName();
          if (!sheetNames.includes(sName)) {
            listToDelete.push(s);
          }
        }

        if (listToDelete.length === tempSs.getSheets().length) {
          return { success: false, error: 'Chyba přípravy PDF. Nebyl by zachován ani jeden list.' };
        }

        for (let s of listToDelete) {
          tempSs.deleteSheet(s);
        }

        SpreadsheetApp.flush();

        AppLogger.info('Stahuji PDF data z Google serverů...');
        const url = 'https://docs.google.com/spreadsheets/d/' + tempSs.getId() + '/export?' + exportOptions;

        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(url, {
          headers: {
            'Authorization': 'Bearer ' + token
          },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
          AppLogger.error('API vrátilo stav ' + response.getResponseCode());
          return { success: false, error: 'Generování selhalo na straně serveru: ' + response.getContentText().substring(0, 100) };
        }

        const blob = response.getBlob().setName(fileName);
        const newFile = folder.createFile(blob);

        AppLogger.ok('PDF ' + fileName + ' vygenerováno.');
        return { success: true, fileId: newFile.getId() };
      } finally {
        // Vždy smazat dočasný soubor — i při výjimce výše
        try { tempFile.setTrashed(true); } catch (e) { /* ignore */ }
      }

    } catch (e) {
      AppLogger.error('Chyba těla PDF exportu: ' + e.message);
      return { success: false, error: e.message };
    }
  },

  /**
   * Generuje PDF pro každou vybranou filiálku do zvoleného adresáře.
   * @param {Array<{id: string, name: string}>} branches - Seznam cílových filiálek
   * @param {Object} options - Nastavení tisku (sheetName, cellId, rangeId, prefix, base, suffix, portrait)
   * @returns {{success: boolean, count?: number, error?: string}}
   */
  exportBatch(branches, options) {
    AppLogger.info('Začínám dávkový export PDF pro ' + branches.length + ' filiálek.');

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(options.sheetName);

      if (!sheet) {
        return { success: false, error: 'Šablonový list "' + options.sheetName + '" nebyl nalezen.' };
      }

      // 1. Zjistíme cílovou složku (sdílená logika)
      const folderResult = ModulePDF.resolveTargetFolder_();
      if (!folderResult.success) return { success: false, error: folderResult.error };
      const folder = folderResult.folder;

      // Příprava parametrů pro URL tabulky
      const ssId = ss.getId();
      const sheetId = sheet.getSheetId();

      // Token potřebný pro UrlFetchApp (aby skript mohl číst tabulku)
      const token = ScriptApp.getOAuthToken();

      let successCount = 0;
      const total = branches.length;

      // HLAVNÍ CYKLUS
      for (let i = 0; i < total; i++) {
        const branch = branches[i];
        AppLogger.info('[' + (i + 1) + '/' + total + '] Filiálka ' + branch.id + ' — ' + branch.name);

        // A) Změnit ID v buňce
        sheet.getRange(options.cellId).setValue(branch.id);

        // B) Vynutit přepočet všech vzorců
        SpreadsheetApp.flush();

        // Získat souřadnice rozsahu z textu "A1:G50" (Google export používá 0-based r1,c1,r2,c2)
        const range = sheet.getRange(options.rangeId);
        const r1 = range.getRow() - 1;
        const c1 = range.getColumn() - 1;
        const r2 = range.getLastRow();
        const c2 = range.getLastColumn();

        // C) Připravit Export URL pro jedinou stránku
        // format=pdf&gid=...&portrait=...&size=A4&gridlines=false... atd
        let urlOptions =
          'exportFormat=pdf&format=pdf' +
          '&size=A4' +
          '&portrait=' + (options.portrait ? 'true' : 'false') +
          '&fitw=true' +       // fit to width
          '&horizontal_alignment=CENTER' +
          '&vertical_alignment=TOP' +
          '&gridlines=false' +
          '&printtitle=false' +
          '&sheetnames=false' +
          '&pagenumbers=false' +
          '&attachment=true' +
          '&gid=' + sheetId +
          '&r1=' + r1 + '&c1=' + c1 + '&r2=' + r2 + '&c2=' + c2;

        const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' + urlOptions;

        // D) Stáhnout Blob
        const response = UrlFetchApp.fetch(url, {
          headers: {
            'Authorization': 'Bearer ' + token
          },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
          throw new Error('Chyba při stahování PDF pro LC: ' + branch.id + ' (' + response.getContentText() + ')');
        }

        const blob = response.getBlob();

        // E) Stanovit jméno souboru
        // Nahrazení tagů v base
        let baseName = options.base || 'Filiálka_{id}';
        baseName = baseName.replace(/{id}/g, branch.id).replace(/{name}/g, branch.name);

        const finalName = (options.prefix || '') + baseName + (options.suffix || '') + '.pdf';
        blob.setName(finalName);

        // F) Uložit soubor
        folder.createFile(blob);
        successCount++;

        // Malá pauza pro stabilitu API
        Utilities.sleep(200);
      }

      AppLogger.ok('Dávkový export PDF dokončen: ' + successCount + ' / ' + branches.length);
      return { success: true, count: successCount };

    } catch (e) {
      AppLogger.error('Chyba PDF exportu: ' + e.message);
      return { success: false, error: e.message };
    }
  },

  /**
   * Sdílené řešení cílové složky pro export() i exportBatch().
   * Vrátí { success: true, folder } nebo { success: false, error }.
   * @private
   * @returns {{ success: boolean, folder?: GoogleAppsScript.Drive.Folder, error?: string }}
   */
  resolveTargetFolder_() {
    let folderId = AppConfig.get('pdfFolderId');

    if (!folderId) {
      folderId = Utils.getBaseFolderId();
      if (!folderId) {
        return { success: false, error: 'Nelze určit kořenovou složku. Nastavte složku v Nastavení.' };
      }
    }

    try {
      return { success: true, folder: DriveApp.getFolderById(folderId) };
    } catch (e) {
      return { success: false, error: 'Cílová složka (ID: ' + folderId + ') nebyla nalezena. Zkontrolujte Nastavení.' };
    }
  },

  resolveCurrentFileFolder_() {
    const folderId = Utils.getBaseFolderId();
    if (!folderId) {
      return { success: false, error: 'Nelze určit složku aktuální tabulky.' };
    }

    try {
      return { success: true, folder: DriveApp.getFolderById(folderId) };
    } catch (e) {
      return { success: false, error: 'Složka aktuální tabulky nebyla nalezena.' };
    }
  },

  restoreSheetState_(ss, allSheets, originalHidden, originalFrozen, originalActive) {
    try {
      allSheets.forEach(sheet => sheet.showSheet());
      SpreadsheetApp.flush();

      allSheets.forEach(sheet => {
        const frozen = originalFrozen[sheet.getSheetId()];
        if (frozen) {
          sheet.setFrozenRows(frozen.rows || 0);
          sheet.setFrozenColumns(frozen.columns || 0);
        }
      });
      SpreadsheetApp.flush();

      allSheets.forEach(sheet => {
        if (originalHidden[sheet.getSheetId()] && ss.getSheets().filter(s => !s.isSheetHidden()).length > 1) {
          sheet.hideSheet();
        }
      });

      if (originalActive && !originalActive.isSheetHidden()) {
        ss.setActiveSheet(originalActive);
      }
      SpreadsheetApp.flush();
    } catch (e) {
      AppLogger.warn('Nepodařilo se plně obnovit viditelnost listů: ' + e.message);
    }
  },

  buildSheetPdfName_(spreadsheetName, sheetName) {
    const base = ModulePDF.sanitizeFileName_(spreadsheetName || 'Soubor');
    const suffix = ModulePDF.sanitizeFileName_(sheetName || 'List');
    return base + '_' + suffix + '.pdf';
  },

  sanitizeFileName_(name) {
    return String(name)
      .replace(/[\/\\:*?"<>|#%\{\}~&]/g, '-')
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 180);
  },

  /**
   * Dostupné listy pro export (bez skrytých a _prefixed)
   * @returns {Array<{name: string, rowCount: number, colCount: number}>}
   */
  getExportableSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheets()
      .filter(s => !s.getName().startsWith('_') && !s.isSheetHidden())
      .map(s => ({
        name: s.getName(),
        rowCount: s.getLastRow(),
        colCount: s.getLastColumn()
      }));
  },

  /**
   * Načte počáteční data pro zobrazení formuláře:
   * 1. Seznam listů
   * 2. Seznam logistických center (z nastavení)
   * 3. Párování LC + Filiálek (z listu 'Filiálky')
   */
  getInitData() {
    AppLogger.info('Načítám inicializační data PDF...');

    // 1. Listy
    const sheets = this.getExportableSheets();

    // 2. LC z konfigurace
    let lcs = [];
    try {
      const raw = AppConfig.get('lcList');
      if (raw) lcs = JSON.parse(raw);
    } catch (e) {
      AppLogger.warn('Chyba při parsování LC');
    }

    // 3. Filiálky
    const branches = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const brSheet = ss.getSheetByName('Filiálky');

    if (brSheet) {
      const lastRow = brSheet.getLastRow();
      if (lastRow > 1) {
        // Zkusíme najít data (A: LC, B: ID, C: Jméno, J: Zavřeno, K: Teststore)
        // Chtějí se sloupce A až K (11 sloupců)
        const vals = brSheet.getRange(2, 1, lastRow - 1, 11).getValues();
        const today = new Date();
        today.setHours(0, 0, 0, 0); // Půlnoc pro porovnání dat

        for (let i = 0; i < vals.length; i++) {
          const row = vals[i];
          const lcId = row[0] ? String(row[0]).trim() : '';
          let bId = row[1] ? String(row[1]).trim() : '';

          if (bId) {
            // Kontrola Teststore (Sloupec K, index 10)
            const teststoreVal = row[10] ? String(row[10]).trim().toLowerCase() : '';
            if (teststoreVal === 'ano') {
              continue; // Ignorovat Teststore
            }

            // Kontrola Zavřeno (Sloupec J, index 9)
            const closedDateVal = row[9];
            if (closedDateVal && closedDateVal instanceof Date) {
              if (closedDateVal < today) {
                continue; // Ignorovat zavřené pobočky (datum v minulosti)
              }
            }

            // Odstranit .0 pokud je cislo jako string
            if (bId.endsWith('.0')) bId = bId.replace('.0', '');

            // Ignorovat filiálky s ID > 900
            const numId = parseInt(bId, 10);
            if (!isNaN(numId) && numId > 900) {
              continue;
            }

            // Jméno
            let bName = row[2] ? String(row[2]).trim() : '';

            branches.push({
              lcId: lcId,
              id: bId,
              name: bName
            });
          }
        }
      }
    } else {
      AppLogger.warn('List "Filiálky" nebyl nalezen.');
    }

    let folderId = Utils.getBaseFolderId() || '';
    let folderName = '';
    if (folderId) {
      try {
        folderName = DriveApp.getFolderById(folderId).getName();
      } catch (e) { }
    }

    return {
      sheets: sheets.map(s => s.name),
      lcs: lcs,
      branches: branches,
      targetFolderId: folderId,
      targetFolderName: folderName,
      settings: ModulePDF.getSettings_(),
      activeLcId: AppConfig.get('activeLcId') || ''
    };
  },

  getSettings_() {
    return {
      rangeId: AppConfig.get('pdfRangeId') || 'B2:M10',
      size: AppConfig.get('pdfSize') || 'A4',
      orientation: AppConfig.get('pdfOrientation') || 'portrait',
      gridlines: String(AppConfig.get('pdfGridlines') || '') === 'true',
      pageNumbers: String(AppConfig.get('pdfPageNumbers') || '') === 'true'
    };
  },

  saveSettings(options) {
    options = options || {};
    AppConfig.setMultiple({
      pdfRangeId: options.rangeId || 'B2:M10',
      pdfSize: options.size === 'A3' ? 'A3' : 'A4',
      pdfOrientation: options.orientation === 'landscape' ? 'landscape' : 'portrait',
      pdfGridlines: options.gridlines ? 'true' : 'false',
      pdfPageNumbers: options.pageNumbers ? 'true' : 'false'
    });
    return { success: true, settings: ModulePDF.getSettings_() };
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function modulePDF_export(sheetNames, options) {
  return ModulePDF.export(sheetNames, options);
}

function modulePDF_exportSheets(sheetNames, options) {
  return ModulePDF.exportSheets(sheetNames, options);
}

function modulePDF_exportBatch(branches, options) {
  return ModulePDF.exportBatch(branches, options);
}

function modulePDF_getExportableSheets() {
  return ModulePDF.getExportableSheets();
}

function modulePDF_getInitData() {
  return ModulePDF.getInitData();
}

function modulePDF_saveSettings(options) {
  return ModulePDF.saveSettings(options);
}
