/**
 * ══════════════════════════════════════════════
 * AppLogger.gs — Logovací systém
 * Vyhodnocení odpisů akčních artiklů | Lidl interní nástroj
 *
 * Používá CacheService pro sdílení logů
 * mezi server-side kódem a klientem (Modal.html).
 * Klient polluje každou sekundu přes appLog_getMessages().
 * ══════════════════════════════════════════════
 */

const AppLogger = {

  /** Klíč v CacheService */
  CACHE_KEY: 'VYCHOZI2_LOG',

  /** Maximální počet uchovávaných zpráv */
  MAX_ENTRIES: 150,

  /** Expirace cache v sekundách (10 minut) */
  CACHE_EXPIRY: 600,

  /**
   * Interní buffer pro dávkové logování.
   * Pokud není null, zprávy se ukládají sem a nezapisují do cache ihned.
   * @private
   */
  _buffer: null,

  /**
   * Zahájí dávkové logování — zprávy se akumulují v paměti.
   * Ukončit voláním endBatch(). Redukuje počet cache read-write cyklů
   * z N na 1 read + 1 write bez ohledu na počet zpráv.
   */
  beginBatch() {
    AppLogger._buffer = [];
  },

  /**
   * Ukončí dávkové logování a zapíše všechny buffered zprávy najednou.
   */
  endBatch() {
    if (!AppLogger._buffer || AppLogger._buffer.length === 0) {
      AppLogger._buffer = null;
      return;
    }
    const buffered = AppLogger._buffer;
    AppLogger._buffer = null;

    const cache = CacheService.getUserCache();
    let logs;
    try {
      const raw = cache.get(AppLogger.CACHE_KEY);
      logs = raw ? JSON.parse(raw) : [];
    } catch (e) {
      logs = [];
    }

    logs = logs.concat(buffered);
    if (logs.length > AppLogger.MAX_ENTRIES) {
      logs = logs.slice(-AppLogger.MAX_ENTRIES);
    }

    try {
      cache.put(AppLogger.CACHE_KEY, JSON.stringify(logs), AppLogger.CACHE_EXPIRY);
    } catch (e) {
      logs = logs.slice(-50);
      cache.put(AppLogger.CACHE_KEY, JSON.stringify(logs), AppLogger.CACHE_EXPIRY);
    }
  },

  /**
   * Přidání zprávy do logu.
   * Pokud je aktivní batch režim (beginBatch), zpráva se bufferuje v paměti
   * a neprovede se žádná cache operace — viz endBatch().
   * @param {string} message - text zprávy
   * @param {string} level - úroveň: info, ok, warn, error, dim
   */
  log(message, level) {
    level = level || 'info';
    const entry = {
      time: Utils.formatTime(new Date()),
      message: String(message),
      level: level
    };

    // Současně i do console.log (pro Stackdriver)
    console.log('[' + entry.time + '] [' + level.toUpperCase() + '] ' + message);

    // Batch režim: pouze přidat do bufferu, žádná cache operace
    if (AppLogger._buffer !== null) {
      AppLogger._buffer.push(entry);
      return;
    }

    // Normální režim: read-modify-write do cache
    const cache = CacheService.getUserCache();
    let logs;
    try {
      const raw = cache.get(AppLogger.CACHE_KEY);
      logs = raw ? JSON.parse(raw) : [];
    } catch (e) {
      logs = [];
    }

    logs.push(entry);

    if (logs.length > AppLogger.MAX_ENTRIES) {
      logs = logs.slice(-AppLogger.MAX_ENTRIES);
    }

    try {
      cache.put(AppLogger.CACHE_KEY, JSON.stringify(logs), AppLogger.CACHE_EXPIRY);
    } catch (e) {
      // Cache plná — zkusíme zkrátit
      logs = logs.slice(-50);
      cache.put(AppLogger.CACHE_KEY, JSON.stringify(logs), AppLogger.CACHE_EXPIRY);
    }
  },

  /** Informační zpráva */
  info(msg) { AppLogger.log(msg, 'info'); },

  /** Úspěch */
  ok(msg) { AppLogger.log(msg, 'ok'); },

  /** Varování */
  warn(msg) { AppLogger.log(msg, 'warn'); },

  /** Chyba */
  error(msg) { AppLogger.log(msg, 'error'); },

  /** Tlumená zpráva (systémová) */
  dim(msg) { AppLogger.log(msg, 'dim'); },

  /**
   * Získání všech zpráv (voláno z klienta)
   * @returns {Array<{time: string, message: string, level: string}>}
   */
  getMessages() {
    const cache = CacheService.getUserCache();
    try {
      const raw = cache.get(AppLogger.CACHE_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      return [];
    }
  },

  /**
   * Získání zpráv od určitého indexu (pro inkrementální polling)
   * @param {number} fromIndex - index od kterého začít
   * @returns {{ messages: Array, total: number }}
   */
  getMessagesSince(fromIndex) {
    const all = AppLogger.getMessages();
    // Pokud je fromIndex větší než počet zpráv v cache, cache byla
    // mezitím smazána a naplněna znovu od nuly (appLog_clear + nové zprávy).
    // V takovém případě vrátíme vše od začátku aby klient viděl nové zprávy.
    const from = (fromIndex > 0 && fromIndex > all.length) ? 0 : (fromIndex || 0);
    return {
      messages: all.slice(from),
      total: all.length
    };
  },

  /**
   * Vymazání logu
   */
  clear() {
    const cache = CacheService.getUserCache();
    cache.put(AppLogger.CACHE_KEY, JSON.stringify([]), AppLogger.CACHE_EXPIRY);
  },

  /**
   * Inicializace logu s uvítací sekvencí
   * Vymaže cache a zapíše úvodní zprávy atomicky
   */
  initSequence() {
    const cache = CacheService.getUserCache();
    const now = Utils.formatTime(new Date());

    // Vytvoříme celou úvodní sekvenci najednou (atomicky)
    const initLogs = [
      { time: now, message: '────────────────────────────', level: 'dim' },
      { time: now, message: 'Vyhodnocení odpisů akčních artiklů — Start', level: 'info' },
      { time: now, message: 'Verze: ' + AppConfig.getVersion() + ' | Runtime: V8', level: 'dim' }
    ];

    // Zapíšeme vše najednou — předejde race condition s pollingem
    cache.put(AppLogger.CACHE_KEY, JSON.stringify(initLogs), AppLogger.CACHE_EXPIRY);
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

/**
 * Získání log zpráv (polling z klienta)
 */
function appLog_getMessages() {
  return AppLogger.getMessages();
}

/**
 * Inkrementální polling — vrátí jen nové zprávy
 * @param {number} fromIndex
 */
function appLog_getMessagesSince(fromIndex) {
  return AppLogger.getMessagesSince(fromIndex);
}

/**
 * Vymazání logu
 */
function appLog_clear() {
  AppLogger.clear();
}

/**
 * Zápis logu z klienta
 * @param {string} msg
 * @param {string} level
 */
function appLog_log(msg, level) {
  AppLogger.log(msg, level);
}

