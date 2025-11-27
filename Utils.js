/**
 * Utils.js
 * Ce fichier contient les fonctions utilitaires partagées par tous les scripts du projet.
 * Cela évite la duplication de code et facilite la maintenance.
 */

/**
 * Affiche un message "toast" dans l'interface Google Sheets.
 */
function Utils_toast(msg, title, seconds) {
    try {
        SpreadsheetApp.getActive().toast(msg, title || 'Info', seconds || 5);
    } catch (e) {
        // Ignore error if UI is not available (e.g. time-driven trigger)
    }
}

/**
 * Exécute une fonction avec des réessais automatiques en cas d'erreur (backoff exponentiel).
 */
function Utils_executeWithRetry(fn, label, maxRetries) {
    label = label || 'task';
    maxRetries = maxRetries || 3;
    let lastErr;
    for (let i = 1; i <= maxRetries; i++) {
        try {
            const t0 = Date.now();
            const out = fn();
            Logger.log(`[OK ] ${label} (${((Date.now() - t0) / 1000).toFixed(1)}s)`);
            return out;
        } catch (e) {
            lastErr = e;
            const msg = String(e && e.message || e);
            Logger.log(`[ERR] ${label} try ${i}/${maxRetries}: ${msg}`);
            // Retry only on specific errors or always?
            // The original scripts had regex checks, we'll keep a generic check or the one from Site Internet.js
            if (i === maxRetries || !/(^|\s)(429|5\d\d|RATE_LIMIT|RESOURCE_EXHAUSTED|DEADLINE_EXCEEDED)/i.test(msg)) break;
            Utilities.sleep(Math.pow(2, i) * 1000);
        }
    }
    throw lastErr;
}

/**
 * Normalise une chaîne pour la comparaison d'en-têtes (minuscule, sans accents, sans espaces superflus).
 */
function Utils_normHeader(s) {
    return String(s || '').toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/[’'`]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * Trouve l'index d'une colonne (1-based) en cherchant parmi une liste d'alias.
 */
function Utils_findColByHeaderAliases(sheet, aliases, headerRow) {
    const row = headerRow || 3;
    const headers = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
    const HN = headers.map(Utils_normHeader);
    const wanted = aliases.map(Utils_normHeader);
    for (let i = 0; i < HN.length; i++) {
        const h = HN[i];
        if (!h) continue;
        if (wanted.some(w => h === w || h.includes(w))) return i + 1;
    }
    return 0;
}

/**
 * S'assure qu'une colonne existe avec l'en-tête donné, sinon la crée.
 */
function Utils_ensureColByHeader(sheet, wantedHeader, headerRow) {
    const row = headerRow || 3;
    const col = Utils_findColByHeaderAliases(sheet, [wantedHeader], row);
    if (col) return col;
    const lastCol = sheet.getLastColumn();
    sheet.insertColumnAfter(lastCol);
    const newCol = lastCol + 1;
    sheet.getRange(row, newCol).setValue(wantedHeader);
    return newCol;
}

/**
 * Détermine si une ligne est une ligne de séparation d'année (contient juste "2024", "2025"...).
 */
function Utils_isYearSeparatorRow(cell) {
    if (cell == null) return false;
    if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) return false;
    const s = String(cell).trim();
    if (/^\d{4}$/.test(s)) {
        const y = +s;
        return y >= 1900 && y <= 2100;
    }
    if (typeof cell === 'number' && isFinite(cell)) {
        const y = Math.round(cell);
        return y >= 1900 && y <= 2100;
    }
    return false;
}

/**
 * Convertit une valeur de cellule (Date, string, etc.) en clé "YYYY-MM".
 * Gère les formats français (ex: "novembre 2025").
 */
function Utils_sheetCellToYYYYMM(cell) {
    // 1) Date object
    if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
        const y = cell.getFullYear();
        const m = cell.getMonth() + 1;
        return `${y}-${String(m).padStart(2, '0')}`;
    }

    const raw = String(cell || '').trim();
    if (!raw) return null;

    const lower = raw.toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ').trim();

    let m;
    // 2) Numeric formats
    if ((m = lower.match(/^(\d{4})-(\d{1,2})$/))) return `${m[1]}-${String(+m[2]).padStart(2, '0')}`;
    if ((m = lower.match(/^(\d{4})\/(\d{1,2})$/))) return `${m[1]}-${String(+m[2]).padStart(2, '0')}`;
    if ((m = lower.match(/^(\d{4})(\d{2})$/))) return `${m[1]}-${m[2]}`;
    if ((m = lower.match(/^(\d{4})-(\d{2})-\d{2}$/))) return `${m[1]}-${m[2]}`;

    // 3) French text formats
    const frMap = {
        'janvier': 1, 'janv': 1, 'jan': 1,
        'fevrier': 2, 'fevr': 2, 'fev': 2,
        'mars': 3,
        'avril': 4, 'avr': 4,
        'mai': 5,
        'juin': 6,
        'juillet': 7, 'juil': 7,
        'aout': 8, 'août': 8, 'aou': 8,
        'septembre': 9, 'sept': 9,
        'octobre': 10, 'oct': 10,
        'novembre': 11, 'nov': 11,
        'decembre': 12, 'dec': 12
    };

    if ((m = lower.match(/^([a-z\.]+)\s+(\d{4})$/))) {
        const name = m[1].replace(/\./g, '');
        const year = m[2];
        const monthNum = frMap[name];
        if (monthNum) return `${year}-${String(monthNum).padStart(2, '0')}`;
    }

    // 4) JS Date parsing fallback
    const d = new Date(raw);
    if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        const mm = d.getMonth() + 1;
        return `${y}-${String(mm).padStart(2, '0')}`;
    }

    return null;
}

/**
 * Convertit "YYYY-MM" en "mois année" (français).
 */
function Utils_monthKeyToFr(ym) {
    const FR = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
    const y = ym.slice(0, 4);
    const m = Math.max(0, Math.min(11, parseInt(ym.slice(5, 7), 10) - 1));
    return `${FR[m]} ${y}`;
}

/**
 * Écrit une valeur dans une cellule UNIQUEMENT si elle ne contient pas de formule.
 */
function Utils_setPreserveFormula(sh, row, col, value, numberFormat) {
    if (!col || !sh) return;
    const cell = sh.getRange(row, col);
    const formula = cell.getFormula();
    if (formula) {
        if (formula.trim()) {
            Logger.log(`[SKIP] Préserve formule en ${cell.getA1Notation()} -> ${formula}`);
        }
        return;
    }
    // Logger.log(`[WRITE] ${sh.getName()} | ${cell.getA1Notation()} | Val: "${value}"`);
    cell.setValue(value);
    if (numberFormat) cell.setNumberFormat(numberFormat);
}

/**
 * Convertit des secondes en jours pour l'affichage "Durée" de Sheets ([h]:mm:ss).
 */
function Utils_setSecondsAsDuration(sh, row, col, secs) {
    const days = (typeof secs === 'number' && !isNaN(secs)) ? secs / 86400 : 0;
    Utils_setPreserveFormula(sh, row, col, days, '[h]:mm:ss');
}

/**
 * Normalise un nom de domaine (supprime https, www, path).
 */
function Utils_normalizeDomain(d) {
    return String(d || '').toLowerCase()
        .replace(/^https?:\/\//, '')
        .replace(/^www\./, '')
        .replace(/\/.*$/, '')
        .trim();
}

/**
 * Récupère une valeur imbriquée dans un objet via un chemin (ex: "a.b.c").
 */
function Utils_valAt(obj, path) {
    return path.split('.').reduce((o, k) => (o && o[k] != null ? o[k] : undefined), obj);
}

/**
 * Vérifie si un texte correspond à l'un des éléments d'un tableau (comparaison normalisée).
 */
function Utils_textEqualsAny(txt, arr) {
    const t = Utils_normHeader(txt || '');
    return (arr || []).some(v => t === Utils_normHeader(v));
}

/**
 * Vérifie si un texte contient l'un des éléments d'un tableau (comparaison normalisée).
 */
function Utils_textIncludesAny(txt, arr) {
    const t = Utils_normHeader(txt || '');
    return (arr || []).some(v => t.includes(Utils_normHeader(v)));
}

/**
 * Vérifie si deux textes sont égaux (comparaison normalisée).
 */
function Utils_textEquals(a, b) {
    return Utils_normHeader(a) === Utils_normHeader(b);
}

/**
 * Récupère une propriété de script avec support de fallback et gestion d'erreurs.
 * @param {string} key - Nom de la propriété principale
 * @param {string[]} fallbackKeys - Noms alternatifs à essayer si la principale n'existe pas
 * @param {boolean} required - Si true, lance une erreur si aucune propriété n'est trouvée
 * @returns {string|null} Valeur de la propriété ou null
 */
function Utils_getProperty(key, fallbackKeys = [], required = true) {
    const props = PropertiesService.getScriptProperties();
    const value = props.getProperty(key);

    if (value) return value;

    // Essayer les fallbacks
    for (const fallback of fallbackKeys) {
        const val = props.getProperty(fallback);
        if (val) {
            Logger.log(`[WARN] Using fallback property '${fallback}' for '${key}'`);
            return val;
        }
    }

    if (required) {
        const tried = [key, ...fallbackKeys].join(', ');
        throw new Error(`Property '${key}' not found (tried: ${tried})`);
    }

    return null;
}

/**
 * Récupère le token Matomo.
 */
function Utils_getMatomoToken() {
    return Utils_getProperty('MATOMO_TOKEN');
}

/**
 * Récupère l'ID du segment Matomo (optionnel).
 */
function Utils_getMatomoSegmentId() {
    return Utils_getProperty('MATOMO_SEGMENT_ID', [], false);
}

/**
 * Récupère l'API key Magnetis.
 */
function Utils_getMagnetisApiKey() {
    return Utils_getProperty('MAGNETIS_API_KEY');
}

/**
 * Récupère le token Paperform.
 * Support de la typo historique PAPARFORM_TOKEN.
 */
function Utils_getPaperformToken() {
    return Utils_getProperty('PAPERFORM_TOKEN', ['PAPARFORM_TOKEN']);
}

/**
 * Récupère le token Monday.
 */
function Utils_getMondayToken() {
    return Utils_getProperty('MONDAY_TOKEN');
}

/**
 * Récupère le token Meta.
 */
function Utils_getMetaToken() {
    return Utils_getProperty('META_TOKEN');
}

/**
 * Exécute une requête GraphQL vers l'API Monday.com.
 * @param {string} query - Requête GraphQL
 * @param {object} variables - Variables de la requête
 * @returns {object} Données de réponse
 */
function Utils_mondayGraphQL(query, variables) {
    const token = Utils_getMondayToken();
    const res = UrlFetchApp.fetch('https://api.monday.com/v2', {
        method: 'post',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': token,
            'API-Version': '2023-10'
        },
        payload: JSON.stringify({ query, variables: variables || {} }),
        muteHttpExceptions: true
    });

    if (res.getResponseCode() >= 300) {
        throw new Error(`Monday HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    }

    const body = JSON.parse(res.getContentText());
    if (body.errors) {
        throw new Error('Monday GraphQL error: ' + JSON.stringify(body.errors));
    }

    return body.data;
}

