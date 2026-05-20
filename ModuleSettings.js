/**
 * ══════════════════════════════════════════════
 * ModuleSettings.gs — Modul: Nastavení
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Rozhraní pro správu konfigurace aplikace.
 * Složky jsou detekované automaticky, ale uživatel
 * si je může změnit (vloží URL nebo ID).
 * ══════════════════════════════════════════════
 */

const ModuleSettings = {

  /**
   * Získání dat pro stránku nastavení
   * Vrátí automaticky detekované složky + případné uživatelské přepisy
   * @returns {Object}
   */
  getData() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Automatická složka (kde leží tabulka)
      let autoFolderId = Utils.getBaseFolderId();
      let autoFolderName = 'Kořen Drive';
      if (autoFolderId) {
        try {
          autoFolderName = DriveApp.getFolderById(autoFolderId).getName();
        } catch (e) {
          autoFolderName = 'Složka nenalezena';
        }
      } else {
        autoFolderId = '';
      }

      // Uživatelské přepisy (pokud existují v _Config)
      const ktFolderOverride = AppConfig.get('ktFolderId') || '';
      const pdfFolderOverride = AppConfig.get('pdfFolderId') || '';
      const tempFolderOverride = AppConfig.get('tempFolderId') || '';

      // Zjistíme jména přepsaných složek
      let ktFolderName = '';
      if (ktFolderOverride) {
        try {
          ktFolderName = DriveApp.getFolderById(ktFolderOverride).getName();
        } catch (e) {
          ktFolderName = '⚠ Složka nenalezena';
        }
      }

      let pdfFolderName = '';
      if (pdfFolderOverride) {
        try {
          pdfFolderName = DriveApp.getFolderById(pdfFolderOverride).getName();
        } catch (e) {
          pdfFolderName = '⚠ Složka nenalezena';
        }
      }

      let tempFolderName = '';
      if (tempFolderOverride) {
        try {
          tempFolderName = DriveApp.getFolderById(tempFolderOverride).getName();
        } catch (e) {
          tempFolderName = '⚠ Složka nenalezena';
        }
      }

      return {
        spreadsheetName: ss.getName(),
        appName: AppConfig.get('appName') || 'Vyhodnocení odpisů akčních artiklů',
        autoFolderId: autoFolderId,
        autoFolderName: autoFolderName,
        ktFolderId: ktFolderOverride,
        ktFolderName: ktFolderName,
        pdfFolderId: pdfFolderOverride,
        pdfFolderName: pdfFolderName,
        tempFolderId: tempFolderOverride,
        tempFolderName: tempFolderName,
        activeLcId: AppConfig.get('activeLcId') || '',
        showLcInHeader: String(AppConfig.get('showLcInHeader') ?? '') !== 'false'
      };
    } catch (e) {
      return { error: e.message };
    }
  },

  /**
   * Uložení názvu aplikace
   * @param {string} name
   * @returns {{ success: boolean }}
   */
  saveName(name) {
    try {
      name = name || 'Vyhodnocení odpisů akčních artiklů';
      AppConfig.set('appName', name);

      // Přejmenování tabulky
      try {
        SpreadsheetApp.getActiveSpreadsheet().rename(name);
      } catch (e) {
        AppLogger.warn('Přejmenování tabulky selhalo: ' + e.message);
      }

      // Přejmenování Apps Script projektu přes Drive API v3
      // (script projekt je Drive soubor — rename jde přes files.patch)
      try {
        const scriptId = ScriptApp.getScriptId();
        const token = ScriptApp.getOAuthToken();
        const resp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + scriptId, {
          method: 'patch',
          contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + token },
          payload: JSON.stringify({ name: name }),
          muteHttpExceptions: true
        });
        const code = resp.getResponseCode();
        if (code !== 200) {
          AppLogger.warn('Přejmenování Apps Script selhalo (' + code + '): ' + resp.getContentText().substring(0, 200));
        }
      } catch (e) {
        AppLogger.warn('Přejmenování Apps Script selhalo: ' + e.message);
      }

      // Menu se nedá aktualizovat za běhu (GAS omezení) — projeví se při příštím onOpen

      AppLogger.ok('Název aplikace uložen: ' + name);
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  /**
   * Uložení nastavení složky (přijímá URL nebo ID)
   * @param {string} key - 'ktFolderId' nebo 'pdfFolderId'
   * @param {string} value - Google Drive URL nebo ID (prázdný = reset na auto)
   * @returns {{ success: boolean, folderId?: string, folderName?: string, error?: string }}
   */
  saveFolder(key, value) {
    try {
      if (!value || value.trim() === '') {
        // Reset na automatickou detekci
        AppConfig.set(key, '');
        AppLogger.ok('Složka resetována na automatickou detekci (' + key + ')');
        return { success: true, folderId: '', folderName: '' };
      }

      // Extrakce ID z URL nebo použití přímo jako ID
      const trimmed = value.trim();
      let folderId = trimmed;

      // Pokud je to URL, extrahujeme ID
      if (trimmed.startsWith('http')) {
        const extracted = Utils.extractFileIdFromUrl(trimmed);
        if (!extracted) {
          return { success: false, error: 'Nepodařilo se extrahovat ID z URL' };
        }
        folderId = extracted;
      }

      // Ověření, že je to platná složka
      let folder;
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (e) {
        return { success: false, error: 'Složka s tímto ID nebyla nalezena' };
      }

      AppConfig.set(key, folderId);
      AppLogger.ok('Složka nastavena: ' + folder.getName() + ' (' + key + ')');

      return {
        success: true,
        folderId: folderId,
        folderName: folder.getName()
      };
    } catch (e) {
      AppLogger.error('Chyba nastavení: ' + e.message);
      return { success: false, error: e.message };
    }
  },

  /**
   * Získá uložená Logistická centra z _Config listu jako pole objektů.
   * @returns {Array<{id: string, name: string, abbr: string}>}
   */
  getLCList() {
    try {
      const raw = AppConfig.get('lcList');
      if (!raw) return [];
      return JSON.parse(raw);
    } catch (e) {
      AppLogger.error('Chyba při čtení LC dat: ' + e.message);
      return [];
    }
  },

  /**
   * Přidá nebo upraví Logistické centrum.
   * @param {string} id - číslo LC (např. '01')
   * @param {string} name - jméno (např. 'Brandýs')
   * @param {string} abbr - zkratka (např. 'BDY')
   * @returns {{success: boolean, lcs?: Array, error?: string}}
   */
  saveLC(id, name, abbr) {
    try {
      if (!id || !name) return { success: false, error: 'Chybí ID nebo Název LC' };
      const list = ModuleSettings.getLCList();
      const existingIdx = list.findIndex(l => l.id === id);

      if (existingIdx >= 0) {
        list[existingIdx] = { id: id, name: name, abbr: abbr || '' };
        AppLogger.ok('LC úprava: ' + id + ' (' + name + ')');
      } else {
        list.push({ id: id, name: name, abbr: abbr || '' });
        // Setřídit podle ID
        list.sort((a, b) => {
          const numA = parseInt(a.id, 10);
          const numB = parseInt(b.id, 10);
          if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
          return a.id.localeCompare(b.id);
        });
        AppLogger.ok('Nové LC přidáno: ' + id + ' (' + name + ')');
      }

      AppConfig.set('lcList', JSON.stringify(list));
      return { success: true, lcs: list };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  /**
   * Smaže Logistické centrum.
   * @param {string} id
   * @returns {{success: boolean, lcs?: Array, error?: string}}
   */
  deleteLC(id) {
    try {
      const list = ModuleSettings.getLCList();
      const filtered = list.filter(l => l.id !== id);

      if (list.length === filtered.length) {
        return { success: false, error: 'LC s ID ' + id + ' nebylo nalezeno' };
      }

      AppConfig.set('lcList', JSON.stringify(filtered));
      AppLogger.ok('LC odstraněno: ' + id);
      return { success: true, lcs: filtered };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  /**
   * Nastaví a zapíše zvolené aktivní LC.
   * Uloží do klíče a natvrdo zapíše hodnoty na list _Config do I, J, K.
   * @param {string} id
   * @returns {{success: boolean, error?: string}}
   */
  setActiveLC(id) {
    try {
      const list = ModuleSettings.getLCList();
      const lc = list.find(l => l.id === id);
      if (!lc) return { success: false, error: 'LC nebylo nalezeno' };

      AppConfig.set('activeLcId', id);
      AppConfig.writeActiveLC(lc.id, lc.name, lc.abbr);
      AppLogger.ok('Aktivní LC nastaveno: ' + lc.id + ' (' + lc.name + ')');

      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  /**
   * Zruší výběr aktivního LC (vyčistí klíč i list _Config).
   * @returns {{success: boolean, error?: string}}
   */
  clearActiveLC() {
    try {
      AppConfig.set('activeLcId', '');
      AppConfig.clearActiveLC();
      AppLogger.ok('Aktivní LC zrušeno');
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  /**
   * Nastaví zobrazení LC u loga v hlavičce.
   * @param {boolean} show
   * @returns {{success: boolean, error?: string}}
   */
  saveShowLc(show) {
    try {
      AppConfig.set('showLcInHeader', show ? 'true' : 'false');
      AppLogger.ok('Nastavení zobrazení LC uloženo: ' + (show ? 'Zobrazit' : 'Skrýt'));
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function moduleSettings_getData() {
  return ModuleSettings.getData();
}

function moduleSettings_saveFolder(key, value) {
  return ModuleSettings.saveFolder(key, value);
}

function moduleSettings_saveName(name) {
  return ModuleSettings.saveName(name);
}

function moduleSettings_getLCList() {
  return ModuleSettings.getLCList();
}

function moduleSettings_saveLC(id, name, abbr) {
  return ModuleSettings.saveLC(id, name, abbr);
}

function moduleSettings_deleteLC(id) {
  return ModuleSettings.deleteLC(id);
}

function moduleSettings_setActiveLC(id) {
  return ModuleSettings.setActiveLC(id);
}

function moduleSettings_clearActiveLC() {
  return ModuleSettings.clearActiveLC();
}

function moduleSettings_saveShowLc(show) {
  return ModuleSettings.saveShowLc(show);
}

