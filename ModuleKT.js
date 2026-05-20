/**
 * ══════════════════════════════════════════════
 * ModuleKT.gs — Modul: Nový KT soubor
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Vytváří nový Google Sheets soubor
 * zkopírováním šablony (mustru) a pojmenuje
 * jej podle kalendářního týdne.
 * ══════════════════════════════════════════════
 */

const ModuleKT = {

  /**
   * Vytvoření nového KT souboru
   * Šablona = aktuální tabulka (mustr).
   * @param {Object} params
   * @param {number} params.week - číslo kalendářního týdne
   * @param {number} params.year - rok
   * @param {string} params.prefix - prefix názvu (výchozí 'KT')
   * @param {string} params.name - střed názvu (výchozí = číslo týdne)
   * @param {string} params.suffix - suffix názvu (výchozí = '_rok')
   * @param {string} [params.targetFolderId] - URL nebo ID cílové složky (prázdné = auto)
   * @param {boolean} [params.useSubfolders] - true = vytvořit rok/KTxx podsložky
   * @returns {{ success: boolean, fileId?: string, fileName?: string, url?: string, error?: string }}
   */
  create(params) {
    params = params || {};
    AppLogger.info('═══ Nový KT soubor ═══');

    try {
      var week = params.week;
      var year = params.year;
      var prefix = (params.prefix !== undefined && params.prefix !== null) ? params.prefix : 'KT';
      var name = (params.name !== undefined && params.name !== null) ? params.name : String(week).padStart(2, '0');
      var suffix = (params.suffix !== undefined && params.suffix !== null) ? params.suffix : ('_' + year);
      var useSubfolders = (params.useSubfolders === true || params.useSubfolders === 'true');

      var fileName = prefix + name + suffix;

      AppLogger.info('Týden: KT ' + String(week).padStart(2, '0') + ' / ' + year);
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

      // Pokud nový soubor má _Config list, nastavíme week/year
      var configSheet = newSS.getSheetByName('_Config');
      if (configSheet) {
        var cfgLastRow = configSheet.getLastRow();
        var data = cfgLastRow > 0 ? configSheet.getRange(1, 1, cfgLastRow, 2).getValues() : [];
        var weekSet = false, yearSet = false;

        for (var i = 0; i < data.length; i++) {
          if (data[i][0] === 'ktWeek') {
            configSheet.getRange(i + 1, 2).setValue(week);
            weekSet = true;
          } else if (data[i][0] === 'ktYear') {
            configSheet.getRange(i + 1, 2).setValue(year);
            yearSet = true;
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
      if (params.prefix !== undefined) updates['ktPrefix'] = params.prefix;
      if (params.suffix !== undefined) updates['ktSuffix'] = params.suffix;
      if (params.useSubfolders !== undefined) updates['ktUseSubfolders'] = params.useSubfolders ? '1' : '0';
      if (params.targetFolderId !== undefined) updates['ktTargetFolderId'] = params.targetFolderId || '';
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
    var subVal = AppConfig.get('ktUseSubfolders');
    // Default = true (zaškrtnuto) — pokud klíč neexistuje (null), default je true
    // Sheets může konvertovat '1' na číslo 1, proto porovnáváme přes String()
    var useSub;
    if (subVal === null || subVal === undefined || subVal === '') {
      useSub = true; // default
    } else {
      useSub = (String(subVal) === '1' || subVal === true);
    }
    return {
      prefix: AppConfig.get('ktPrefix') || 'KT',
      suffix: AppConfig.get('ktSuffix') || '',
      useSubfolders: useSub,
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

