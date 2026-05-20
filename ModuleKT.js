/**
 * ══════════════════════════════════════════════
 * ModuleKT.gs — Modul: Nový KT soubor
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Vytváří nový Google Sheets soubor
 * zkopírováním šablony (mustru) a pojmenuje
 * jej podle data akce.
 * ══════════════════════════════════════════════
 */

const ModuleKT = {

  /**
   * Vytvoření nového KT souboru
   * Šablona = aktuální tabulka (mustr).
   * @param {Object} params
   * @param {string} params.actionDate - datum akce ve formátu yyyy-MM-dd
   * @param {string} [params.targetFolderId] - URL nebo ID cílové složky (prázdné = auto)
   * @returns {{ success: boolean, fileId?: string, fileName?: string, url?: string, error?: string }}
   */
  create(params) {
    params = params || {};
    AppLogger.info('═══ Nový KT soubor ═══');

    try {
      var dateInfo = ModuleKT.getDateInfo_(params.actionDate);
      var week = dateInfo.week;
      var year = dateInfo.year;
      var useSubfolders = true;
      var fileName = year + '_KT' + String(week).padStart(2, '0') + '_' + dateInfo.dayName;

      AppLogger.info('Týden: KT ' + String(week).padStart(2, '0') + ' / ' + year);
      AppLogger.info('Datum akce: ' + dateInfo.displayDate + ' (' + dateInfo.dayName + ')');
      AppLogger.info('Název souboru: ' + fileName);
      AppLogger.dim('Podsložky: ' + (useSubfolders ? 'ano' : 'ne'));

      // Šablona = aktuální tabulka
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var templateFile = DriveApp.getFileById(ss.getId());
      AppLogger.ok('Šablona: ' + templateFile.getName() + ' ✓');

      // Uložení preferencí pro příště
      ModuleKT.savePreferences_(params);

      // Cílová složka
      var baseFolder = ModuleKT.resolveBaseFolder_(params.targetFolderId, templateFile);
      AppLogger.dim('Výchozí složka: ' + baseFolder.getName());

      // Podsložky: rok/KTxx
      var targetFolder = baseFolder;
      if (useSubfolders) {
        var yearFolder = ModuleKT.getOrCreateSubfolder_(baseFolder, String(year));
        var ktFolder = ModuleKT.getOrCreateSubfolder_(yearFolder, 'KT' + String(week).padStart(2, '0'));
        targetFolder = ktFolder;
        AppLogger.dim('Cesta: ' + baseFolder.getName() + '/' + String(year) + '/KT' + String(week).padStart(2, '0'));
      }

      // Kopírování šablony
      AppLogger.info('Kopíruji šablonu...');
      var newFile = templateFile.makeCopy(fileName, targetFolder);
      AppLogger.ok('Soubor vytvořen: ' + fileName + ' ✓');

      // Otevření a nastavení parametrů
      AppLogger.info('Nastavuji parametry...');
      var newSS = SpreadsheetApp.open(newFile);
      ModuleKT.writeActionDate_(newSS, dateInfo.displayDate);

      // Pokud nový soubor má _Config list, nastavíme week/year
      var configSheet = newSS.getSheetByName('_Config');
      if (configSheet) {
        var cfgLastRow = configSheet.getLastRow();
        var data = cfgLastRow > 0 ? configSheet.getRange(1, 1, cfgLastRow, 2).getValues() : [];
        var weekSet = false, yearSet = false, dateSet = false;

        for (var i = 0; i < data.length; i++) {
          if (data[i][0] === 'ktWeek') {
            configSheet.getRange(i + 1, 2).setValue(week);
            weekSet = true;
          } else if (data[i][0] === 'ktYear') {
            configSheet.getRange(i + 1, 2).setValue(year);
            yearSet = true;
          } else if (data[i][0] === 'ktActionDate') {
            configSheet.getRange(i + 1, 2).setValue(dateInfo.isoDate);
            dateSet = true;
          }
        }

        if (!weekSet) {
          var lr = configSheet.getLastRow() + 1;
          configSheet.getRange(lr, 1, 1, 2).setValues([['ktWeek', week]]);
        }
        if (!yearSet) {
          var lr2 = configSheet.getLastRow() + 1;
          configSheet.getRange(lr2, 1, 1, 2).setValues([['ktYear', year]]);
        }
        if (!dateSet) {
          var lr3 = configSheet.getLastRow() + 1;
          configSheet.getRange(lr3, 1, 1, 2).setValues([['ktActionDate', dateInfo.isoDate]]);
        }

        AppLogger.ok('Parametry nastaveny (KT' + String(week).padStart(2, '0') + '/' + year + ') ✓');
      } else {
        AppLogger.dim('_Config list v novém souboru nenalezen — přeskakuji parametry');
      }

      var url = newFile.getUrl();
      AppLogger.ok('═══ KT soubor připraven ═══');
      AppLogger.info('URL: ' + url);

      return {
        success: true,
        fileId: newFile.getId(),
        fileName: fileName,
        url: url
      };

    } catch (e) {
      AppLogger.error('Chyba: ' + e.message);
      return {
        success: false,
        error: e.message
      };
    }
  },

  /**
   * @private
   * @param {string} actionDate
   * @returns {{date: Date, isoDate: string, displayDate: string, week: number, year: number, dayName: string}}
   */
  getDateInfo_(actionDate) {
    if (!actionDate) throw new Error('Zadejte datum akce.');
    var parts = String(actionDate).split('-').map(Number);
    if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) {
      throw new Error('Datum akce musí být ve formátu yyyy-MM-dd.');
    }

    var date = new Date(parts[0], parts[1] - 1, parts[2]);
    if (isNaN(date.getTime())) throw new Error('Neplatné datum akce.');

    var dayNames = ['NEDĚLE', 'PONDĚLÍ', 'ÚTERÝ', 'STŘEDA', 'ČTVRTEK', 'PÁTEK', 'SOBOTA'];
    return {
      date: date,
      isoDate: Utilities.formatDate(date, 'Europe/Prague', 'yyyy-MM-dd'),
      displayDate: Utilities.formatDate(date, 'Europe/Prague', 'dd.MM.yyyy'),
      week: Utils.getWeekNumber(date),
      year: date.getFullYear(),
      dayName: dayNames[date.getDay()]
    };
  },

  /**
   * Zapíše datum akce do výstupních listů nového souboru.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
   * @param {string} displayDate
   */
  writeActionDate_(ss, displayDate) {
    ['BNL', 'OLO', 'CER', 'BUS', 'BRV'].forEach(function (name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) sheet.getRange('B8').setValue('Akce od ' + displayDate);
    });
    AppLogger.ok('Datum akce zapsáno do listů BNL/OLO/CER/BUS/BRV ✓');
  },

  /**
   * Resolve cílové složky — uživatelský přepis nebo složka mustru
   * @private
   * @param {string} targetFolderInput - URL nebo ID nebo prázdné
   * @param {GoogleAppsScript.Drive.File} templateFile
   * @returns {GoogleAppsScript.Drive.Folder}
   */
  resolveBaseFolder_(targetFolderInput, templateFile) {
    // 1) Přímý vstup z formuláře (URL nebo ID)
    if (targetFolderInput && String(targetFolderInput).trim()) {
      var trimmed = String(targetFolderInput).trim();
      var folderId = trimmed;

      if (trimmed.startsWith('http')) {
        var extracted = Utils.extractFileIdFromUrl(trimmed);
        if (extracted) folderId = extracted;
      }

      try {
        return DriveApp.getFolderById(folderId);
      } catch (e) {
        AppLogger.warn('Zadaná složka nenalezena, zkouším nastavení...');
      }
    }

    // 2) Uloženou override z nastavení
    var ktFolderOverride = AppConfig.get('ktFolderId');
    if (ktFolderOverride) {
      try {
        return DriveApp.getFolderById(ktFolderOverride);
      } catch (e) {
        AppLogger.warn('Nastavená složka nenalezena, používám výchozí');
      }
    }

    // 3) Výchozí = složka tabulky
    var autoId = Utils.getBaseFolderId();
    if (autoId) {
      try {
        return DriveApp.getFolderById(autoId);
      } catch (e) {
        // Fallback níže
      }
    }

    return DriveApp.getRootFolder();
  },

  /**
   * Najde nebo vytvoří podsložku v dané složce
   * @private
   * @param {GoogleAppsScript.Drive.Folder} parent
   * @param {string} name
   * @returns {GoogleAppsScript.Drive.Folder}
   */
  getOrCreateSubfolder_(parent, name) {
    var folders = parent.getFoldersByName(name);
    if (folders.hasNext()) {
      return folders.next();
    }
    AppLogger.dim('Vytvářím podsložku: ' + name);
    return parent.createFolder(name);
  },

  /**
   * Uloží uživatelské preference do _Config
   * @private
   * @param {Object} params
   */
  savePreferences_(params) {
    try {
      // Dávkový zápis — 1 čtení + 1 zápis místo 4 × (čtení + zápis)
      const updates = {};
      if (params.targetFolderId !== undefined) updates['ktTargetFolderId'] = params.targetFolderId || '';
      if (params.actionDate !== undefined) updates['ktActionDate'] = params.actionDate || '';
      if (Object.keys(updates).length > 0) AppConfig.setMultiple(updates);
    } catch (e) {
      // Preferences save is non-critical
      AppLogger.dim('Preference neuloženy: ' + e.message);
    }
  },

  /**
   * Načtení uložených preferencí
   * @returns {Object}
   */
  getPreferences() {
    return {
      actionDate: AppConfig.get('ktActionDate') || '',
      useSubfolders: true,
      targetFolderId: AppConfig.get('ktTargetFolderId') || ''
    };
  },

  /**
   * Získání informací o šabloně (= aktuální tabulka) + auto složka
   * @returns {{ configured: boolean, name?: string, id?: string, autoFolderName?: string, autoFolderId?: string }}
   */
  getTemplateInfo() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var autoFolderId = Utils.getBaseFolderId() || '';
      var autoFolderName = 'Kořen Drive';
      if (autoFolderId) {
        try {
          autoFolderName = DriveApp.getFolderById(autoFolderId).getName();
        } catch (e) {
          autoFolderName = 'Složka nenalezena';
        }
      }

      return {
        configured: true,
        name: ss.getName(),
        id: ss.getId(),
        autoFolderName: autoFolderName,
        autoFolderId: autoFolderId
      };
    } catch (e) {
      return {
        configured: false,
        error: 'Nelze získat info o aktuální tabulce'
      };
    }
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function moduleKT_create(params) {
  return ModuleKT.create(params);
}

function moduleKT_getTemplateInfo() {
  return ModuleKT.getTemplateInfo();
}

function moduleKT_getPreferences() {
  return ModuleKT.getPreferences();
}

