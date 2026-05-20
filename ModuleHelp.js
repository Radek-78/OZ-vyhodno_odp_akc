/**
 * ══════════════════════════════════════════════
 * ModuleHelp.gs — Modul: Nápověda
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Obsah nápovědy uložen v skrytém listu _Help.
 * Každá sekce má název a obsah (HTML podporován).
 * ══════════════════════════════════════════════
 */

const ModuleHelp = {

  /** Název listu s nápovědou */
  SHEET_NAME: '_Help',

  /**
   * Inicializace _Help listu s výchozím obsahem
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  init() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(ModuleHelp.SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(ModuleHelp.SHEET_NAME);
      sheet.hideSheet();

      // Struktura
      sheet.getRange('A1:B1').setValues([['Sekce', 'Obsah']]);
      sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#e8eaf6');

      // Výchozí obsah nápovědy
      const helpContent = [
        [
          'Úvod',
          'Vyhodnocení odpisů akčních artiklů je interní nástroj pro správu KT souborů, import dat a generování PDF.\n\n' +
          'Aplikaci ovládáte přes centrální panel, který otevřete z menu „⚙️ Vyhodnocení odpisů akčních artiklů → Otevřít panel".\n' +
          'Panel obsahuje přehled dostupných modulů, systémový log a stavový řádek.'
        ],
        [
          'Nový KT soubor',
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
          '• Soubor můžete otevřít ihned nebo odpočet zrušit'
        ],
        [
          'Import souborů',
          'Modul pro import XLSX a CSV souborů do aktivní tabulky.\n\n' +
          '1. Klikněte na kartu „Import souborů"\n' +
          '2. Vyberte soubor z Google Drive\n' +
          '3. Zvolte cílový list (nebo vytvořte nový)\n' +
          '4. Nastavte volby importu\n' +
          '5. Klikněte „Importovat"\n\n' +
          'Import probíhá po dávkách s průběžným logováním.\n' +
          'Podporované formáty: .xlsx, .xls, .csv\n' +
          'Maximální doporučená velikost: 50 000 řádků.'
        ],
        [
          'Generování PDF',
          'Modul pro export vybraných listů jako PDF soubor.\n\n' +
          '1. Klikněte na kartu „Generování PDF"\n' +
          '2. Vyberte listy k exportu (zaškrtněte)\n' +
          '3. Nastavte formát (A4/A3, orientace, mřížka)\n' +
          '4. Volitelně zadejte název souboru\n' +
          '5. Klikněte „Generovat PDF"\n\n' +
          'PDF bude uloženo na Google Drive do stejné složky jako tabulka.'
        ],
        [
          'Jak to funguje',
          'Aplikace pracuje automaticky:\n\n' +
          '• Logo je vloženo přímo v aplikaci\n' +
          '• Šablona pro nový KT = aktuální tabulka (mustr)\n' +
          '• Cílová složka se detekuje automaticky (složka mustru)\n' +
          '• Volitelně lze nastavit vlastní složku (URL nebo ID)\n' +
          '• PDF se ukládá do stejné složky jako tabulka\n' +
          '• Nový KT soubor lze organizovat do podsložek rok/KTxx\n\n' +
          'V Nastavení (⚙️) můžete přepsat výchozí složky pro KT i PDF.\n' +
          'Preference modulu KT se automaticky ukládají pro příští použití.'
        ],
        [
          'Changelog',
          'Historie změn aplikace.\n\n' +
          'Každý záznam má typ (major/minor/patch), popis a autora.\n' +
          'Verze se automaticky přepočítává z changelog záznamů:\n' +
          '• major = zásadní změna (verze x+1.0.0)\n' +
          '• minor = nová funkce (verze x.y+1.0)\n' +
          '• patch = oprava (verze x.y.z+1)'
        ],
        [
          'Řešení problémů',
          'Časté problémy:\n\n' +
          '• „Šablona nenalezena" → Ověřte, že spouštíte z mustru\n' +
          '• Import trvá dlouho → Normální u velkých souborů, sledujte log\n' +
          '• PDF se nevytváří → Zkontrolujte oprávnění k tabulce\n\n' +
          'Pro další pomoc kontaktujte správce aplikace.'
        ]
      ];

      sheet.getRange(2, 1, helpContent.length, 2).setValues(helpContent);

      // Šířky
      sheet.setColumnWidth(1, 180);
      sheet.setColumnWidth(2, 600);

      // Zalamování textu
      sheet.getRange(2, 2, helpContent.length, 1).setWrap(true);

      AppLogger.ok('List _Help vytvořen s výchozím obsahem');
    }

    return sheet;
  },

  /**
   * Získání obsahu nápovědy
   * @returns {Array<{title: string, content: string}>}
   */
  getContent() {
    const sheet = ModuleHelp.init();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const sections = [];

    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) {
        sections.push({
          title: data[i][0],
          content: data[i][1] || ''
        });
      }
    }

    return sections;
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function moduleHelp_getContent() {
  return ModuleHelp.getContent();
}

