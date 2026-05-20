/**
 * ══════════════════════════════════════════════
 * Utils.gs — Pomocné funkce
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 * ══════════════════════════════════════════════
 */

const Utils = {

  /**
   * Formátování data pro zobrazení
   * @param {Date} date
   * @param {string} format - formátovací řetězec (default: dd.MM.yyyy HH:mm)
   * @returns {string}
   */
  formatDate(date, format) {
    format = format || 'dd.MM.yyyy HH:mm';
    return Utilities.formatDate(
      date instanceof Date ? date : new Date(date),
      'Europe/Prague',
      format
    );
  },

  /**
   * Formátování času (pouze hodiny:minuty:sekundy)
   * @param {Date} date
   * @returns {string}
   */
  formatTime(date) {
    return Utilities.formatDate(
      date instanceof Date ? date : new Date(date),
      'Europe/Prague',
      'HH:mm:ss'
    );
  },

  /**
   * Výpočet čísla kalendářního týdne (ISO 8601)
   * @param {Date} date
   * @returns {number}
   */
  getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  },

  /**
   * Aktuální kalendářní týden a rok
   * @returns {{ week: number, year: number }}
   */
  getCurrentWeek() {
    const now = new Date();
    return {
      week: Utils.getWeekNumber(now),
      year: now.getFullYear()
    };
  },

  /**
   * Rozdělení pole na bloky (pro dávkové zpracování)
   * @param {Array} array
   * @param {number} size - velikost bloku
   * @returns {Array<Array>}
   */
  chunkArray(array, size) {
    const chunks = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  },

  /**
   * Pauza (prevence timeout při dlouhých operacích)
   * @param {number} ms - milisekundy
   */
  sleep(ms) {
    Utilities.sleep(ms);
  },

  /**
   * Bezpečné získání ID z URL Google Drive
   * @param {string} url
   * @returns {string|null}
   */
  extractFileIdFromUrl(url) {
    if (!url) return null;
    // Specifické Google Drive URL patterny (přesnější než generický fallback)
    const patterns = [
      /\/file\/d\/([a-zA-Z0-9_-]{25,})/,
      /\/folders\/([a-zA-Z0-9_-]{25,})/,
      /\/spreadsheets\/d\/([a-zA-Z0-9_-]{25,})/,
      /\/document\/d\/([a-zA-Z0-9_-]{25,})/,
      /[?&]id=([a-zA-Z0-9_-]{25,})/
    ];
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) return match[1];
    }
    // Fallback: 25+ znakový identifikátor (pro přímá ID)
    const fallback = url.match(/[a-zA-Z0-9_-]{25,}/);
    return fallback ? fallback[0] : null;
  },

  /**
   * Formátování čísla s mezerami jako oddělovačem tisíců
   * @param {number} num
   * @returns {string}
   */
  formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  },

  /**
   * Získání seznamu listů v tabulce (bez skrytých _prefixed)
   * @returns {Array<{name: string, index: number, id: number}>}
   */
  getVisibleSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheets()
      .filter(s => !s.getName().startsWith('_') && !s.isSheetHidden())
      .map((s, i) => ({
        name: s.getName(),
        index: i,
        id: s.getSheetId()
      }));
  },

  /**
   * Převede index sloupce (1-based) na písmeno (A, B, ...)
   * @param {number} col
   * @returns {string}
   */
  columnToLetter(col) {
    let letter = '';
    while (col > 0) {
      let temp = (col - 1) % 26;
      letter = String.fromCharCode(65 + temp) + letter;
      col = (col - temp - 1) / 26;
    }
    return letter;
  },

  /**
   * Získání ID domovské složky aktuální tabulky (kde tabulka leží)
   * Centralizovaná funkce pro všechny moduly (Settings, PDF, KT)
   * @returns {string|null} ID kořenové složky nebo null
   */
  getBaseFolderId() {
    try {
      const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
      const parents = DriveApp.getFileById(ssId).getParents();
      return parents.hasNext() ? parents.next().getId() : null;
    } catch (e) {
      return null;
    }
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function utils_getCurrentWeek() {
  return Utils.getCurrentWeek();
}

function utils_getVisibleSheets() {
  return Utils.getVisibleSheets();
}

