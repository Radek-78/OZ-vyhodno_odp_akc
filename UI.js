/**
 * ══════════════════════════════════════════════
 * UI.gs — Server-side funkce pro UI
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Poskytuje data pro Modal.html při inicializaci
 * a průběžných operacích.
 * ══════════════════════════════════════════════
 */

const UI = {

  /**
   * Inicializační data pro modal UI
   * Voláno jednou při otevření modalu.
   * @returns {{
   *   version: string,
   *   userEmail: string,
   *   config: Object,
   *   sheets: Array,
   *   currentWeek: { week: number, year: number }
   * }}
   */
  getInitialData() {
    AppLogger.clear();

    const email = Session.getActiveUser().getEmail() || 'neznámý uživatel';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = AppConfig.getAll();
    const version = AppConfig.getVersion();
    const sheets = Utils.getVisibleSheets();

    // Zpráva "připraveno" se loguje až po načtení loga na klientovi

    // KT preference — čte pouze z AppConfig (bez Drive API)
    // POZN: ModuleKT.getTemplateInfo() (2× Drive API) se načítá lazy v klientu.
    var ktPref = { prefix: 'KT', suffix: '', useSubfolders: true, targetFolderId: '' };
    try { ktPref = ModuleKT.getPreferences(); } catch (e) { /* fallback defaults */ }

    // Zjistíme aktuální počet log zpráv pro synchronizaci s klientem
    const logTotal = AppLogger.getMessages().length;

    // LC info pro hlavičku
    const activeLcId = config['activeLcId'] || '';
    let activeLcName = '';
    if (activeLcId) {
      try {
        const lcList = JSON.parse(config['lcList'] || '[]');
        const lc = lcList.find(l => l.id === activeLcId);
        if (lc) activeLcName = lc.name;
      } catch (e) { }
    }
    const showLcInHeader = String(config['showLcInHeader'] ?? '') !== 'false';

    // Logo se NEZAHRNUJE do inicializačních dat — klient ho načte asynchronně
    // přes ui_getLogoChunks() aby se neblokovala inicializace Drive API voláním.
    return {
      version: version,
      userEmail: email,
      config: config,
      sheets: sheets,
      currentWeek: Utils.getCurrentWeek(),
      logTotal: logTotal,
      ktPref: ktPref,
      activeLcId: activeLcId,
      activeLcName: activeLcName,
      showLcInHeader: showLcInHeader
    };
  },

  /**
   * Načte logo z Drive a rozdělí ho na bloky (pro obejití limitu scriptletu)
   * @returns {Array<string>}
   */
  getLogoChunks() {
    try {
      const cache = CacheService.getScriptCache();
      const cachedLogo = cache.get('app_logo_chunks');

      if (cachedLogo) {
        return JSON.parse(cachedLogo);
      }

      // Výchozí ID loga z Mustru (Lidl Logo)
      const defaultId = '1OBtCu-6ezu0NuPrLs9Bkrw8QMuM2htaD';

      // Zkusíme najít vlastní ID v konfiguraci
      const customId = AppConfig.get('appLogoId');
      const fileId = customId || defaultId;

      const blob = DriveApp.getFileById(fileId).getBlob();
      const base64 = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;

      // Rozdělit na 30KB bloky (pro větší jistotu)
      const chunkSize = 30000;
      const chunks = [];
      for (let i = 0; i < base64.length; i += chunkSize) {
        chunks.push(base64.slice(i, i + chunkSize)); // slice je rychlejší než substring
      }

      // Uložení do Cache na 6 hodin (21600 sekund)
      cache.put('app_logo_chunks', JSON.stringify(chunks), 21600);

      return chunks;
    } catch (e) {
      console.error('UI.getLogoChunks error: ' + e.message);
      return [];
    }
  },

  /**
   * Informace o aktuální tabulce
   * @returns {{ id: string, name: string, url: string, sheetCount: number }}
   */
  getSpreadsheetInfo() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      id: ss.getId(),
      name: ss.getName(),
      url: ss.getUrl(),
      sheetCount: ss.getSheets().length
    };
  },

  /**
   * Výběr souboru z Google Drive (pro import)
   * Vrátí URL pro Google Picker
   * @returns {{ token: string, appId: string }}
   */
  getPickerData() {
    return {
      token: ScriptApp.getOAuthToken(),
      appId: ScriptApp.getProjectKey() || ''
    };
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function ui_getInitialData() {
  return UI.getInitialData();
}

function ui_getLogoChunks() {
  return UI.getLogoChunks();
}

function ui_getSpreadsheetInfo() {
  return UI.getSpreadsheetInfo();
}

function ui_getPickerData() {
  return UI.getPickerData();
}

