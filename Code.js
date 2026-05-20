/**
 * ══════════════════════════════════════════════
 * Code.gs — Hlavní vstupní bod
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Trigger onOpen, menu, spouštění modal dialogu.
 * ══════════════════════════════════════════════
 */


// ─── Konstanty aplikace ───────────────────────

const APP = {
  NAME: 'Vyhodnocení odpisů akčních artiklů',
  MENU_TITLE: '⚙️ Vyhodnocení odpisů akčních artiklů',
  MODAL_WIDTH: 1200,
  MODAL_HEIGHT: 800
};


// ─── Trigger: onOpen ──────────────────────────

/**
 * Automaticky se spustí při otevření tabulky.
 * Vytvoří vlastní menu v Google Sheets.
 */
function onOpen(e) {
  let name = APP.NAME;
  try { name = AppConfig.get('appName') || APP.NAME; } catch (ex) { }
  buildMenu_(name);
}

/**
 * Sestaví (nebo přestaví) vlastní menu s daným názvem aplikace.
 * Voláno z onOpen i po přejmenování aplikace v Nastavení.
 * @param {string} name - Název aplikace (bez emoji prefixu)
 */
function buildMenu_(name) {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ ' + (name || APP.NAME))
    .addItem('📂 Otevřít panel', 'showMainModal')
    .addSeparator()
    .addItem('📅 Nový KT soubor', 'showModuleKT')
    .addItem('📥 Import souborů', 'showModuleImport')
    .addItem('📄 Generování PDF', 'showModulePDF')
    .addSeparator()
    .addItem('⚙️ Nastavení', 'showModuleSettings')
    .addSeparator()
    .addItem('📖 Nápověda', 'showModuleHelp')
    .addItem('📋 Changelog', 'showModuleChangelog')
    .addItem('ℹ️ O aplikaci', 'showAbout')
    .addToUi();
}


// ─── Hlavní modal dialog ──────────────────────

/**
 * Otevře hlavní modal dialog s domovskou obrazovkou
 */
function showMainModal() {
  showModalWithModule_('home');
}

/**
 * Otevře modal s konkrétním modulem předvybraným
 * @param {string} moduleName - identifikátor modulu
 */
function showModalWithModule_(moduleName) {
  // Vyčistit logy pro novou session hned při otevření
  AppLogger.clear();

  const template = HtmlService.createTemplateFromFile('Modal');
  template.initialModule = moduleName || 'home';

  const html = template
    .evaluate()
    .setWidth(APP.MODAL_WIDTH)
    .setHeight(APP.MODAL_HEIGHT);

  SpreadsheetApp.getUi().showModelessDialog(html, ' ');
}


// ─── Zkratky pro menu ─────────────────────────

function showModuleKT() { showModalWithModule_('kt'); }
function showModuleImport() { showModalWithModule_('import'); }
function showModulePDF() { showModalWithModule_('pdf'); }
function showModuleSettings() { showModalWithModule_('settings'); }
function showModuleHelp() { showModalWithModule_('help'); }
function showModuleChangelog() { showModalWithModule_('changelog'); }


/**
 * Pomocná funkce pro vkládání HTML souborů do šablon.
 * @param {string} filename - Název souboru (bez .html)
 * @return {string} Obsah souboru
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ─── O aplikaci ───────────────────────────────

/**
 * Zobrazí informace o aplikaci
 */
function showAbout() {
  const version = AppConfig.getVersion();
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    'O aplikaci',
    APP.NAME + '\n' +
    'Verze: ' + version + '\n' +
    'Runtime: V8\n' +
    'Časová zóna: Europe/Prague\n\n' +
    'Lidl Česká republika — interní nástroj',
    ui.ButtonSet.OK
  );
}


// ─── Zpětná kompatibilita pro admin funkce ───
// Implementace přesunuta do Admin.js s admin_ prefixem
function initializeApp() { admin_initializeApp(); }
function updateHelpContent() { admin_updateHelpContent(); }
function addKtChangelogEntries() { admin_addKtChangelogEntries(); }

