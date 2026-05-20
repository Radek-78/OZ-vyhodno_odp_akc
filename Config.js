/**
 * ══════════════════════════════════════════════
 * Config.gs — Správa konfigurace a verzí
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Data uložena v skrytém listu _Config:
 *   A:B — klíč/hodnota (nastavení)
 *   D:G — changelog (Datum, Typ, Popis, Autor)
 * ══════════════════════════════════════════════
 */

const AppConfig = {

  /** Název konfiguračního listu */
  SHEET_NAME: '_Config',

  /** Výchozí verze */
  DEFAULT_VERSION: '1.0.0',

  /** Cache pro instanci listu (přežije v rámci jednoho requestu) */
  _sheetInstance: null,

  /** Cache pro konfigurační data (přežije v rámci jednoho requestu) */
  _configCache: null,

  // ─── Inicializace ───────────────────────────

  /**
   * Inicializace _Config listu (vytvoří pokud neexistuje)
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  init() {
    if (AppConfig._sheetInstance) return AppConfig._sheetInstance;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(AppConfig.SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(AppConfig.SHEET_NAME);
      try {
        sheet.hideSheet();
      } catch (e) {
        // Ignorujeme pokud nelze schovat (např. jediný list)
      }

      // Hlavička — Nastavení (A:B)
      sheet.getRange('A1:B1').setValues([['Klíč', 'Hodnota']]);
      sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#e8eaf6');

      // Výchozí hodnoty
      const defaults = [
        ['version', AppConfig.DEFAULT_VERSION],
        ['appName', 'Vyhodnocení odpisů akčních artiklů']
      ];
      sheet.getRange(2, 1, defaults.length, 2).setValues(defaults);

      // Formátování sloupce B na Plain Text pro prevenci auto-dat
      sheet.getRange('B:B').setNumberFormat('@');

      // Hlavička — Changelog (D:G)
      sheet.getRange('D1:G1').setValues([['Datum', 'Typ', 'Popis', 'Autor']]);
      sheet.getRange('D1:G1').setFontWeight('bold').setBackground('#e8eaf6');

      // Hlavička — Aktivní LC (I:K)
      sheet.getRange('I1:K1').setValues([['Aktivní LC - ID', 'Název', 'Zkratka']]);
      sheet.getRange('I1:K1').setFontWeight('bold').setBackground('#e8eaf6');

      // Šířky sloupců
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 300);
      sheet.setColumnWidth(3, 20); // mezera
      sheet.setColumnWidth(4, 120);
      sheet.setColumnWidth(5, 80);
      sheet.setColumnWidth(6, 300);
      sheet.setColumnWidth(7, 200);
      sheet.setColumnWidth(8, 20); // mezera
      sheet.setColumnWidth(9, 100);
      sheet.setColumnWidth(10, 150);
      sheet.setColumnWidth(11, 100);

      // První changelog záznam
      const now = new Date();
      sheet.getRange(2, 4, 1, 4).setValues([[
        now, 'patch', 'Inicializace projektu', Session.getActiveUser().getEmail() || 'system'
      ]]);

      AppLogger.ok('Konfigurační list _Config vytvořen');
    } else {
      // List existuje — zkontrolujeme jestli má changelog hlavičku (sloupec D)
      if (sheet.getLastColumn() < 4 || sheet.getRange(1, 4).getValue() === '') {
        AppLogger.info('Doplňování changelog struktury do _Config...');
        sheet.getRange('D1:G1').setValues([['Datum', 'Typ', 'Popis', 'Autor']]);
        sheet.getRange('D1:G1').setFontWeight('bold').setBackground('#e8eaf6');

        // Zkusíme jestli tam jsou nějaká data, pokud ne, přidáme první záznam
        if (sheet.getRange(2, 4).getValue() === '') {
          const now = new Date();
          sheet.getRange(2, 4, 1, 4).setValues([[
            now, 'patch', 'Inicializace changelogu', Session.getActiveUser().getEmail() || 'system'
          ]]);
        }
      }

      // Zkontrolujeme existenci LC hlaviček v I1
      if (sheet.getRange('I1').getValue() === '') {
        AppLogger.info('Doplňování Aktivní LC struktury do _Config...');
        sheet.getRange('I1:K1').setValues([['Aktivní LC - ID', 'Název', 'Zkratka']]);
        sheet.getRange('I1:K1').setFontWeight('bold').setBackground('#e8eaf6');
        sheet.setColumnWidth(9, 100);
        sheet.setColumnWidth(10, 150);
        sheet.setColumnWidth(11, 100);
      }
    }

    AppConfig._sheetInstance = sheet;
    return sheet;
  },

  // ─── Čtení / Zápis nastavení ────────────────

  /**
   * Získání hodnoty nastavení (využívá cache přes getAll)
   * @param {string} key
   * @returns {*}
   */
  get(key) {
    const all = AppConfig.getAll();
    return (key in all) ? all[key] : null;
  },

  /**
   * Nastavení hodnoty
   * @param {string} key
   * @param {*} value
   */
  set(key, value) {
    AppConfig.setMultiple({ [key]: value });
  },

  /**
   * Dávkové nastavení více hodnot najednou.
   * Provede jediné čtení a skupinové zápisy — výrazně rychlejší než N volání set().
   * @param {Object} kvMap - objekt { klíč: hodnota, ... }
   */
  setMultiple(kvMap) {
    const sheet = AppConfig.init();
    const lastRow = sheet.getLastRow();
    const data = lastRow >= 1 ? sheet.getRange(1, 1, lastRow, 2).getValues() : [];

    // Projdeme existující řádky a najdeme shody
    const toUpdate = {}; // rowIndex (0-based v data) → { key, value }
    const remaining = Object.assign({}, kvMap); // klíče které ještě nebyly nalezeny

    for (let i = 1; i < data.length; i++) {
      const rowKey = data[i][0];
      if (rowKey && rowKey in remaining) {
        toUpdate[i] = { key: rowKey, value: remaining[rowKey] };
        delete remaining[rowKey];
      }
    }

    // Zapsat updates (existující řádky) — každý zvlášť (různé řádky)
    for (const [idx, { key, value }] of Object.entries(toUpdate)) {
      const cell = sheet.getRange(Number(idx) + 1, 2);
      if (key === 'version') cell.setNumberFormat('@');
      cell.setValue(value);
    }

    // Zapsat nové klíče na konec
    let nextRow = sheet.getLastRow() + 1;
    for (const [key, value] of Object.entries(remaining)) {
      sheet.getRange(nextRow, 1).setValue(key);
      const cell = sheet.getRange(nextRow, 2);
      if (key === 'version') cell.setNumberFormat('@');
      cell.setValue(value);
      nextRow++;
    }

    AppConfig._configCache = null; // invalidace cache
  },

  /**
   * Fyzicky zapíše hodnoty zvoleného Aktivního LC do listu _Config do sloupců I, J, K na řádek 2
   * @param {string} id
   * @param {string} name
   * @param {string} abbr
   */
  writeActiveLC(id, name, abbr) {
    const sheet = AppConfig.init();

    // Zapíšeme na I2, J2, K2
    sheet.getRange('I2:K2').setValues([[id, name, abbr]]);
  },

  /**
   * Vymaže hodnoty Aktivního LC z listu _Config ze sloupců I, J, K na řádku 2
   */
  clearActiveLC() {
    const sheet = AppConfig.init();
    sheet.getRange('I2:K2').clearContent();
  },

  /**
   * Promaže cache pro načtení čerstvých dat
   */
  clearCache() {
    AppConfig._configCache = null;
  },

  /**
   * Všechna nastavení jako objekt (s cache pro aktuální request)
   * @returns {Object}
   */
  getAll() {
    if (AppConfig._configCache) return AppConfig._configCache;

    const sheet = AppConfig.init();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { version: AppConfig.DEFAULT_VERSION };

    const data = sheet.getRange(1, 1, lastRow, 2).getValues();
    const config = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        config[data[i][0]] = data[i][1];
      }
    }

    AppConfig._configCache = config;
    return config;
  },

  // ─── Verzování ──────────────────────────────

  /**
   * Aktuální verze
   * @returns {string}
   */
  getVersion() {
    const v = AppConfig.get('version');
    if (!v) return AppConfig.DEFAULT_VERSION;

    // Pokud Sheets převedlo verzi na Date objekt (1.1.0 → 1.1.2024)
    // nebo je to string podobný datu (1.1.2024), vynutíme přepočet
    const needsRecalc = (v instanceof Date) ||
      (typeof v === 'string' && /^\d+\.\d+\.\d{4}$/.test(v));

    if (needsRecalc) {
      const recalculated = AppConfig.calculateVersion();
      AppConfig.set('version', recalculated);
      return recalculated;
    }

    return String(v);
  },

  /**
   * Přepočet verze z changelog záznamů
   * Pravidla:
   *   'major' → major++, minor=0, patch=0
   *   'minor' → minor++, patch=0
   *   'patch' → patch++
   * @returns {string} - formát MAJOR.MINOR.PATCH
   */
  calculateVersion() {
    const changelog = AppConfig.getChangelogRaw_();
    // Pokud nejsou žádné záznamy nebo jen jeden, verze je 1.0.0
    if (changelog.length <= 1) return AppConfig.DEFAULT_VERSION;

    let major = 1, minor = 0, patch = 0;

    // Počítáme až od druhého záznamu (první je inicializační/vstupní)
    changelog.slice(1).forEach(entry => {
      const type = String(entry.type).toLowerCase().trim();
      if (type === 'major') {
        major++;
        minor = 0;
        patch = 0;
      } else if (type === 'minor') {
        minor++;
        patch = 0;
      } else {
        patch++;
      }
    });

    return major + '.' + minor + '.' + patch;
  },

  // ─── Changelog ──────────────────────────────

  /**
   * Interní: zjistí poslední obsazený řádek v changelog bloku (sloupec D)
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @returns {number} - číslo řádku (1-based), nebo 1 pokud changelog prázdný
   */
  _getLastChangelogRow(sheet) {
    const lastRowTotal = sheet.getLastRow();
    if (lastRowTotal < 2) return 1;
    const dValues = sheet.getRange(1, 4, lastRowTotal, 1).getValues();
    for (let i = dValues.length - 1; i >= 0; i--) {
      if (dValues[i][0] !== '') return i + 1;
    }
    return 1;
  },

  /**
   * Interní: surová data changelogu (chronologicky)
   * Čte D:G najednou a filtruje v jednom průchodu.
   * @private
   * @returns {Array<{date: Date, type: string, description: string, author: string}>}
   */
  getChangelogRaw_() {
    const sheet = AppConfig.init();
    const lastClRow = AppConfig._getLastChangelogRow(sheet);
    if (lastClRow < 2) return [];

    const data = sheet.getRange(2, 4, lastClRow - 1, 4).getValues();
    return data
      .filter(row => row[0] !== '')
      .map(row => ({
        date: row[0],
        type: row[1] || 'patch',
        description: row[2] || '',
        author: row[3] || ''
      }));
  },

  /**
   * Changelog pro zobrazení (nejnovější první, formátovaná data)
   * @returns {Array<{date: string, type: string, description: string, author: string}>}
   */
  getChangelog() {
    try {
      const raw = AppConfig.getChangelogRaw_();
      return raw
        .map(e => ({
          date: e.date instanceof Date ? Utils.formatDate(e.date, 'dd.MM.yyyy') : String(e.date || ''),
          type: e.type || 'patch',
          description: e.description || '',
          author: e.author || ''
        }))
        .reverse();
    } catch (e) {
      return [];
    }
  },

  /**
   * Přidání záznamu do changelogu
   * Automaticky přepočítá verzi.
   * @param {string} type - 'major', 'minor', 'patch'
   * @param {string} description
   * @param {string} [author] - email (default: aktuální uživatel)
   * @returns {{ success: boolean, newVersion: string }}
   */
  addChangelogEntry(type, description, author) {
    try {
      const sheet = AppConfig.init();
      const lastClRow = AppConfig._getLastChangelogRow(sheet);

      author = author || Session.getActiveUser().getEmail() || 'unknown';

      sheet.getRange(lastClRow + 1, 4, 1, 4).setValues([[
        new Date(),
        type || 'patch',
        description || '',
        author
      ]]);

      // Přepočet verze
      const newVersion = AppConfig.calculateVersion();
      AppConfig.set('version', newVersion);

      AppLogger.ok('Changelog: ' + type + ' — ' + description);
      AppLogger.info('Nová verze: ' + newVersion);

      return { success: true, newVersion: newVersion };
    } catch (e) {
      AppLogger.error('Chyba při přidávání do changelogu: ' + e.message);
      return { success: false, error: e.message };
    }
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function config_getAll() {
  return AppConfig.getAll();
}

function config_get(key) {
  return AppConfig.get(key);
}

function config_set(key, value) {
  AppConfig.set(key, value);
  return { success: true };
}

function config_setMultiple(kvMap) {
  AppConfig.setMultiple(kvMap);
  return { success: true };
}

function config_getVersion() {
  return AppConfig.getVersion();
}

function config_getChangelog() {
  return AppConfig.getChangelog();
}

function config_addChangelogEntry(type, description) {
  return AppConfig.addChangelogEntry(type, description);
}

