/**
 * ══════════════════════════════════════════════
 * ModuleManagerColoring.gs — Obarvení řádků podle manažerů
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 * ══════════════════════════════════════════════
 */

const ModuleManagerColoring = {
  SHEET_NAMES: ['BNL', 'OLO', 'CER', 'BUS', 'BRV'],
  START_ROW: 10,
  MANAGER_COL: 4,
  LAST_ROW_MARKER_COL: 2,
  COLOR_START_COL: 2,
  COLOR_END_COL: 13,
  STRIPE_COLOR: '#D9E1F2',
  CLEAR_COLOR: '#ffffff',

  /**
   * Obarví řádky na definovaných listech podle po sobě jdoucích skupin manažerů.
   * První skupina manažera zůstane bez barvy, další skupina dostane STRIPE_COLOR.
   * @returns {{success: boolean, results: Array<{sheet: string, rows: number, lastRow?: number, message?: string}>}}
   */
  run() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];

    ModuleManagerColoring.SHEET_NAMES.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) {
        results.push({ sheet: name, rows: 0, message: 'list nenalezen' });
        return;
      }

      const lastRow = ModuleManagerColoring.findLastNumericRow_(sheet);
      if (lastRow < ModuleManagerColoring.START_ROW) {
        results.push({ sheet: name, rows: 0, message: 'bez dat od řádku 10' });
        return;
      }

      const rowCount = lastRow - ModuleManagerColoring.START_ROW + 1;
      const colorColCount = ModuleManagerColoring.COLOR_END_COL - ModuleManagerColoring.COLOR_START_COL + 1;
      const managers = sheet
        .getRange(ModuleManagerColoring.START_ROW, ModuleManagerColoring.MANAGER_COL, rowCount, 1)
        .getDisplayValues()
        .map(row => String(row[0] || '').trim());

      const backgrounds = [];
      let previousManager = '';
      let useStripe = false;

      managers.forEach(manager => {
        if (manager && previousManager && manager !== previousManager) {
          useStripe = !useStripe;
        }
        if (manager) previousManager = manager;

        const color = useStripe ? ModuleManagerColoring.STRIPE_COLOR : ModuleManagerColoring.CLEAR_COLOR;
        backgrounds.push(Array(colorColCount).fill(color));
      });

      sheet
        .getRange(ModuleManagerColoring.START_ROW, ModuleManagerColoring.COLOR_START_COL, rowCount, colorColCount)
        .setBackgrounds(backgrounds);

      results.push({ sheet: name, rows: rowCount, lastRow: lastRow, message: 'hotovo' });
    });

    return { success: true, results: results };
  },

  /**
   * Najde poslední řádek, který má ve sloupci B číselnou hodnotu nebo text obsahující číslici.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @returns {number}
   */
  findLastNumericRow_(sheet) {
    const maxRow = sheet.getLastRow();
    if (maxRow < ModuleManagerColoring.START_ROW) return 0;

    const rowCount = maxRow - ModuleManagerColoring.START_ROW + 1;
    const values = sheet
      .getRange(ModuleManagerColoring.START_ROW, ModuleManagerColoring.LAST_ROW_MARKER_COL, rowCount, 1)
      .getValues();

    for (let i = values.length - 1; i >= 0; i--) {
      if (ModuleManagerColoring.hasNumber_(values[i][0])) {
        return ModuleManagerColoring.START_ROW + i;
      }
    }
    return 0;
  },

  /**
   * @param {*} value
   * @returns {boolean}
   */
  hasNumber_(value) {
    if (typeof value === 'number') return isFinite(value);
    if (value === null || value === undefined) return false;
    return /\d/.test(String(value));
  }
};

function moduleManagerColoring_run() {
  return ModuleManagerColoring.run();
}

function runManagerRowColoring() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = ModuleManagerColoring.run();
    const lines = result.results.map(r => {
      if (r.rows > 0) return r.sheet + ': ' + r.rows + ' řádků (do řádku ' + r.lastRow + ')';
      return r.sheet + ': ' + r.message;
    });

    SpreadsheetApp.getActiveSpreadsheet().toast('Obarvení skupin RM dokončeno', 'Hotovo', 5);
    ui.alert('Obarvení skupin RM', lines.join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Chyba obarvení', e.message, ui.ButtonSet.OK);
  }
}
