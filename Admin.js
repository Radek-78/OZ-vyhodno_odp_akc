/**
 * ══════════════════════════════════════════════
 * Admin.gs — Jednorázové a administrátorské utility
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Funkce určené k ručnímu spuštění z editoru.
 * Nejsou volány z UI — pouze pro údržbu.
 * ══════════════════════════════════════════════
 */


// ─── Inicializace (první spuštění) ────────────

/**
 * Ruční inicializace — vytvoří _Config a _Help listy
 * Spusťte jednou při prvním nasazení.
 */
function admin_initializeApp() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Inicializace',
    'Chcete inicializovat aplikaci ' + APP.NAME + '?\n\n' +
    'Budou vytvořeny konfigurační listy (_Config, _Help).',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  AppConfig.init();
  ModuleHelp.init();

  ui.alert('Hotovo', 'Aplikace byla úspěšně inicializována.\nVerze: ' + AppConfig.getVersion(), ui.ButtonSet.OK);
}


// ─── Aktualizace obsahu nápovědy ─────────────

/**
 * Přepíše obsah _Help listu aktuálním obsahem.
 * Spusťte po úpravě textů nápovědy.
 */
function admin_updateHelpContent() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('_Help');
  if (!sheet) { ModuleHelp.init(); sheet = ss.getSheetByName('_Help'); }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  var helpContent = [
    ['Úvod',
      'Vyhodnocení odpisů akčních artiklů je interní nástroj pro správu KT souborů, import dat a generování PDF.\n\n' +
      'Aplikaci ovládáte přes centrální panel, který otevřete z menu „⚙️ Vyhodnocení odpisů akčních artiklů → Otevřít panel".\n' +
      'Panel obsahuje přehled dostupných modulů, systémový log a stavový řádek.'],
    ['Nový KT soubor',
      'Modul pro vytvoření nového souboru pro kalendářní týden (KT).\n' +
      'Slouží k archivaci dat — pro každý týden si můžete vytvořit vlastní kopii mustru.\n\n' +
      'TÝDEN A ROK\n' +
      '• Číslo týdne a rok se načtou automaticky podle ISO 8601\n' +
      '• Obě hodnoty jsou editovatelné\n\n' +
      'NÁZEV SOUBORU\n' +
      '• Název se skládá z: Prefix + Název + Suffix\n' +
      '• Výchozí formát: KT06_2026 (prefix „KT", název „06", suffix „_2026")\n' +
      '• Všechny části lze libovolně upravit\n' +
      '• Živý náhled výsledného názvu se zobrazuje vpravo\n\n' +
      'CÍLOVÁ SLOŽKA\n' +
      '• Výchozí = složka kde je uložen mustr\n' +
      '• Vlastní složka = zadejte URL nebo ID složky z Google Drive\n\n' +
      'PODSLOŽKY\n' +
      '• Zaškrtněte „Vytvořit podsložky rok / KTxx"\n' +
      '• Soubor se uloží do struktury: složka/2026/KT06/KT06_2026\n' +
      '• Podsložky se vytvoří automaticky pokud neexistují\n\n' +
      'ULOŽENÍ PREFERENCÍ\n' +
      '• Prefix, suffix, volba podsložek a vlastní složka se ukládají\n' +
      '• Při příštím otevření se formulář předvyplní posledním nastavením\n\n' +
      'PO VYTVOŘENÍ\n' +
      '• Zobrazí se potvrzení s názvem a ID nového souboru\n' +
      '• Automaticky se spustí 5s odpočet k otevření souboru\n' +
      '• Soubor můžete otevřít ihned nebo odpočet zrušit'],
    ['Import souborů',
      'Modul pro import XLSX a CSV souborů do aktivní tabulky.\n\n' +
      '1. Klikněte na kartu „Import souborů"\n' +
      '2. Vyberte soubor z Google Drive\n' +
      '3. Zvolte cílový list (nebo vytvořte nový)\n' +
      '4. Nastavte volby importu\n' +
      '5. Klikněte „Importovat"\n\n' +
      'Import probíhá po dávkách s průběžným logováním.\n' +
      'Podporované formáty: .xlsx, .xls, .csv\n' +
      'Maximální doporučená velikost: 50 000 řádků.'],
    ['Generování PDF',
      'Modul pro export vybraných listů jako PDF soubor.\n\n' +
      '1. Klikněte na kartu „Generování PDF"\n' +
      '2. Vyberte listy k exportu (zaškrtněte)\n' +
      '3. Nastavte formát (A4/A3, orientace, mřížka)\n' +
      '4. Volitelně zadejte název souboru\n' +
      '5. Klikněte „Generovat PDF"\n\n' +
      'PDF bude uloženo na Google Drive do stejné složky jako tabulka.'],
    ['Jak to funguje',
      'Aplikace pracuje automaticky:\n\n' +
      '• Logo je vloženo přímo v aplikaci\n' +
      '• Šablona pro nový KT = aktuální tabulka (mustr)\n' +
      '• Cílová složka se detekuje automaticky (složka mustru)\n' +
      '• Volitelně lze nastavit vlastní složku (URL nebo ID)\n' +
      '• PDF se ukládá do stejné složky jako tabulka\n' +
      '• Nový KT soubor lze organizovat do podsložek rok/KTxx\n\n' +
      'V Nastavení (⚙️) můžete přepsat výchozí složky pro KT i PDF.\n' +
      'Preference modulu KT se automaticky ukládají pro příští použití.'],
    ['Changelog',
      'Historie změn aplikace.\n\n' +
      'Každý záznam má typ (major/minor/patch), popis a autora.\n' +
      'Verze se automaticky přepočítává z changelog záznamů:\n' +
      '• major = zásadní změna (verze x+1.0.0)\n' +
      '• minor = nová funkce (verze x.y+1.0)\n' +
      '• patch = oprava (verze x.y.z+1)'],
    ['Řešení problémů',
      'Časté problémy:\n\n' +
      '• „Šablona nenalezena" → Ověřte, že spouštíte z mustru\n' +
      '• Import trvá dlouho → Normální u velkých souborů, sledujte log\n' +
      '• PDF se nevytváří → Zkontrolujte oprávnění k tabulce\n\n' +
      'Pro další pomoc kontaktujte správce aplikace.']
  ];

  sheet.getRange(2, 1, helpContent.length, 2).setValues(helpContent);
  sheet.getRange(2, 2, helpContent.length, 1).setWrap(true);

  SpreadsheetApp.getUi().alert('Nápověda aktualizována (' + helpContent.length + ' sekcí).');
}


// ─── Hromadné přidání changelog záznamů ──────

/**
 * Přidá changelog záznamy pro KT modul.
 * Spusťte jednou po nasazení KT funkcí.
 */
function admin_addKtChangelogEntries() {
  var entries = [
    ['minor', 'Logo Lidl vloženo jako base64 přímo do aplikace — bez závislosti na Google Drive'],
    ['minor', 'Automatická detekce složky mustru a cílových složek — odstranění manuálního nastavení ID'],
    ['minor', 'Nastavení: volitelné přepsání složek pro KT a PDF (URL nebo ID)'],
    ['patch', 'Oprava duplicitních záznamů v systémovém logu — atomické zápisy + odložený polling'],
    ['patch', 'Hlavička: dvouřádkový layout (název + verze / email), tlačítko Zpět v hlavičce'],
    ['minor', 'Redesign modulu KT: konfigurovatelný název souboru (prefix/název/suffix)'],
    ['minor', 'Modul KT: volba cílové složky (výchozí auto nebo vlastní URL/ID)'],
    ['minor', 'Modul KT: volitelné podsložky rok/KTxx s živým náhledem cesty'],
    ['patch', 'Modul KT: uložení preferencí (prefix, suffix, podsložky, složka) pro příští otevření'],
    ['patch', 'Modul KT: vylepšená zpráva po vytvoření souboru s 5s odpočtem a automatickým otevřením'],
    ['patch', 'Oprava scrollování systémového logu (min-height:0 na grid kontejneru)'],
    ['patch', 'Nastavení: kompaktní 3-sloupcový grid s ikonami']
  ];

  entries.forEach(function (e) {
    AppConfig.addChangelogEntry(e[0], e[1]);
  });

  SpreadsheetApp.getUi().alert('Přidáno ' + entries.length + ' changelog záznamů.\nNová verze: ' + AppConfig.getVersion());
}


// ─── Konsolidace changelogu ──────────────────

/**
 * Vyčistí changelog a nahradí jedním souhrnným záznamem (v1.0.0).
 * DESTRUKTIVNÍ OPERACE — smaže celou historii!
 */
function admin_consolidateChangelog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠ DESTRUKTIVNÍ OPERACE',
    'Tato akce NEVRATNĚ smaže celou historii changelogu a nahradí ji jedním souhrnným záznamem.\n\n' +
    'Verze se resetuje na 1.0.0.\n\n' +
    'Opravdu chcete pokračovat?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_Config');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange('D2:G' + lastRow).clearContent();
  }

  const summary =
    'Kompletní inicializace systému Vyhodnocení odpisů akčních artiklů:\n' +
    '• Core: Optimalizace výkonu při práci s konfigurací a verzováním.\n' +
    '• UI: Moderní Lidl design, kompaktní grid v nastavení, vylepšená hlavička s auto-synchronizací verze.\n' +
    '• Modul KT: Flexibilní generování názvů, support podsložek rok/KT a ukládání preferencí.\n' +
    '• Import: Stabilní import XLSX/CSV s dávkovým zpracováním a logováním.\n' +
    '• PDF: Export listů s volbou formátu (A4/A3) a orientace.\n' +
    '• Opravy: Vyřešeny TransportError kolize a problémy s auto-převodem verze na datum.';

  sheet.getRange(2, 4, 1, 4).setValues([[
    new Date(), 'patch', summary, Session.getActiveUser().getEmail() || 'admin'
  ]]);

  AppConfig.set('version', AppConfig.calculateVersion());

  SpreadsheetApp.getUi().alert('Changelog byl úspěšně konsolidován na verzi 1.0.0.');
}

