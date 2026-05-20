/**
 * ══════════════════════════════════════════════
 * ModuleChangelog.gs — Modul: Changelog
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Zobrazení a správa changelogu.
 * Data uložena v _Config listu (sloupce D:G).
 * ══════════════════════════════════════════════
 */

const ModuleChangelog = {

  /**
   * Získání changelogu pro zobrazení
   * @returns {{ entries: Array, version: string }}
   */
  getData() {
    return {
      entries: AppConfig.getChangelog(),
      version: AppConfig.getVersion()
    };
  },

  /**
   * Přidání nového záznamu
   * @param {string} type - 'major', 'minor', 'patch'
   * @param {string} description - popis změny
   * @returns {{ success: boolean, newVersion?: string, error?: string }}
   */
  addEntry(type, description) {
    if (!description || description.trim() === '') {
      return { success: false, error: 'Popis změny je povinný' };
    }

    if (!['major', 'minor', 'patch'].includes(type)) {
      return { success: false, error: 'Neplatný typ změny (povoleno: major, minor, patch)' };
    }

    return AppConfig.addChangelogEntry(type, description.trim());
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function moduleChangelog_getData() {
  try {
    return ModuleChangelog.getData();
  } catch (e) {
    return { entries: [], version: AppConfig.getVersion() || '?', error: e.message };
  }
}

function moduleChangelog_addEntry(type, description) {
  return ModuleChangelog.addEntry(type, description);
}

