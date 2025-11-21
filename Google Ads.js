/**************************** ORCHESTRATION ****************************/
// Fenêtre Full History forcée : depuis 1er janvier 2024 inclus
const FULLHIST_START_LOCAL = new Date(2024, 0, 1);                    // pour APIs en local time
const FULLHIST_START_UTC   = new Date(Date.UTC(2024, 0, 1, 0,0,0));   // pour APIs en UTC

/**************************** FLAGS ****************************/
let ADS_ENABLE_MATOMO = true;  // mettre false pour couper Matomo dans les runners globaux
let ADS_ENABLE_LEADS  = true;  // mettre false pour couper Monday Leads (appels/formulaires lead)

// Petit toast dans l'UI de Sheets (garde une seule définition si déjà présente)
function toast_(msg, title, seconds) {
  try { SpreadsheetApp.getActive().toast(msg, title || 'Sync', seconds || 5); } catch(e) {}
}

// Retry avec backoff (2s, 4s, 8s)
function executeWithRetry_(fn, label, maxRetries) {
  label = label || 'task';
  maxRetries = maxRetries || 3;
  let attempt = 0, lastErr;
  while (attempt < maxRetries) {
    try {
      attempt++;
      Logger.log(`[RUN] ${label} (try ${attempt}/${maxRetries})`);
      const t0 = Date.now();
      const out = fn();
      Logger.log(`[OK ] ${label} in ${((Date.now()-t0)/1000).toFixed(1)}s`);
      return out;
    } catch (err) {
      lastErr = err;
      const msg = String(err && err.message || err);
      Logger.log(`[ERR] ${label}: ${msg}`);
      const isRetryable = /(^|\s)(429|5\d\d|RATE_LIMIT|RESOURCE_EXHAUSTED)/i.test(msg);
      if (attempt >= maxRetries || !isRetryable) break;
      Utilities.sleep(Math.pow(2, attempt) * 1000); // 2s, 4s, 8s
    }
  }
  throw lastErr;
}

// FULL HISTORY : Google Ads + (Forms Paperform+Monday) + Magnetis + Matomo
function run_All_FullHistory() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);
  try {
    toast_('Full history : démarrage…', 'Sync ALL', 5);

    // 1) Google Ads (réécrit l’onglet/mois, insère les séparateurs d’années si besoin)
    executeWithRetry_(() => run_GAds_FullHistory(), 'Google Ads FullHistory');
    Utilities.sleep(400);

    // 2) Formulaires unifiés (Paperform + Monday)
    executeWithRetry_(() => run_FormsToAds_SyncForSheet(), 'Forms (Paperform+Monday) Sync');
    Utilities.sleep(250);

    // 3) Magnetis (appels)
    executeWithRetry_(() => run_MagnetisToAds_SyncForSheet(), 'Magnetis Sync');
    Utilities.sleep(250);

    // 3bis) Monday Leads (appels lead + formulaires lead)
    if (ADS_ENABLE_LEADS) {
      executeWithRetry_(() => run_MondayLeadsToAds_SyncForSheet(), 'Monday Leads (Ads) Sync');
      Utilities.sleep(200);

      // N-1 Monday Leads (utile si tu veux forcer la dernière ligne)
      executeWithRetry_(() => run_MondayLeadsToAds_AddLastMonth(), 'Monday Leads (Ads) N-1');
      Utilities.sleep(150);
    }

    // 4) Recalcule uniquement CTR/Taux de conversion (après clics/appels/formulaires)
    executeWithRetry_(() => run_Ads_Derived_RecomputeAll(), 'Ads Derived (CTR/CR)');

    // 5) Matomo (durée moyenne)
    if (ADS_ENABLE_MATOMO) {
      executeWithRetry_(() => run_MatomoToAds_SyncForSheet(), 'Matomo Sync');
    }

    toast_('Full history : terminé ✅', 'Sync ALL', 5);
  } finally {
    lock.releaseLock();
  }
}

// LAST MONTH : Google Ads + Forms unifiés + Magnetis + (recalc) + Matomo
function Ads_run_All_AddLastMonth() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);
  try {
    toast_('Last month : démarrage…', 'Sync ALL', 5);

    // 1) Google Ads N-1
    executeWithRetry_(() => run_GAds_AddLastMonth(), 'Google Ads AddLastMonth');
    Utilities.sleep(300);

    // 2) Formulaires (Paperform + Monday) N-1
    executeWithRetry_(() => run_FormsToAds_AddLastMonth(), 'Forms AddLastMonth');
    Utilities.sleep(200);

    // 3) Magnetis N-1
    executeWithRetry_(() => run_MagnetisToAds_AddLastMonth(), 'Magnetis AddLastMonth');
    Utilities.sleep(200);

        // 3bis) Monday Leads N-1 (appels lead + formulaires lead)
    if (ADS_ENABLE_LEADS) {
      executeWithRetry_(() => run_MondayLeadsToAds_AddLastMonth(), 'Monday Leads (Ads) N-1');
      Utilities.sleep(150);
    }

    // 4) Recalcule CTR/CR pour N-1 (après formulaires + appels)
    executeWithRetry_(() => run_Ads_Derived_RecomputeLastMonth(), 'Ads Derived N-1');

    // 5) Matomo N-1
    if (ADS_ENABLE_MATOMO) {
      executeWithRetry_(() => run_MatomoToAds_AddLastMonth(), 'Matomo AddLastMonth');
    }


    toast_('Last month : terminé ✅', 'Sync ALL', 5);
  } finally {
    lock.releaseLock();
  }
}
function Ads_run_All_CurrentMonth() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);
  try {
    toast_('Mois en cours : démarrage…', 'Sync ALL', 5);

    // 1) Google Ads mois en cours
    executeWithRetry_(() => run_GAds_CurrentMonth(), 'Google Ads CurrentMonth');
    Utilities.sleep(300);

    // 2) Formulaires (Paperform + Monday) mois en cours
    executeWithRetry_(() => run_FormsToAds_CurrentMonth(), 'Forms CurrentMonth');
    Utilities.sleep(200);

    // 3) Magnetis mois en cours
    executeWithRetry_(() => run_MagnetisToAds_CurrentMonth(), 'Magnetis CurrentMonth');
    Utilities.sleep(200);

    // 3bis) Monday Leads mois en cours (appels lead + formulaires lead)
    if (ADS_ENABLE_LEADS) {
      executeWithRetry_(() => run_MondayLeadsToAds_CurrentMonth(), 'Monday Leads (Ads) CurrentMonth');
      Utilities.sleep(150);
    }

    // 4) Recalcule CTR/CR pour le mois en cours
    executeWithRetry_(() => run_Ads_Derived_RecomputeCurrentMonth(), 'Ads Derived CurrentMonth');

    // 5) Matomo mois en cours
    if (ADS_ENABLE_MATOMO) {
      executeWithRetry_(() => run_MatomoToAds_CurrentMonth(), 'Matomo CurrentMonth');
    }

    toast_('Mois en cours : terminé ✅', 'Sync ALL', 5);
  } finally {
    lock.releaseLock();
  }
}

/**************************** MENU ****************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Update Google Ads')
    .addItem('▶ Tout (Full history)', 'run_All_FullHistory')
    .addItem('▶ Dernier mois',        'Ads_run_All_AddLastMonth')
    .addToUi();
}

/**************************** HELPERS GÉNÉRAUX ****************************/

// Détecte une ligne "année" (ex: 2024 / 2025)
function isYearSeparatorRow_(cell) {
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

// Convertit la valeur de la cellule (A: "mois") -> "YYYY-MM" (gère "août 2025", etc.)
function sheetCellToYYYYMM_(cell) {
  // Date Sheets ?
  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    const y = cell.getFullYear(), m = cell.getMonth()+1;
    return `${y}-${String(m).padStart(2,'0')}`;
  }
  const raw = String(cell || '').trim();
  if (!raw) return null;

  // Formats numériques
  let m;
  if ((m = raw.match(/^(\d{4})-(\d{1,2})$/)))   return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  if ((m = raw.match(/^(\d{4})\/(\d{1,2})$/)))  return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  if ((m = raw.match(/^(\d{4})(\d{2})$/)))      return `${m[1]}-${m[2]}`;
  if ((m = raw.match(/^(\d{4})-(\d{2})-\d{2}$/))) return `${m[1]}-${m[2]}`;

  // Mois FR
  const norm = raw.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const FR = {
    'janvier':1,'fevrier':2,'mars':3,'avril':4,'mai':5,'juin':6,
    'juillet':7,'aout':8,'septembre':9,'octobre':10,'novembre':11,'decembre':12
  };
  const mm = norm.match(/(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre)\s+(\d{4})/);
  if (mm) {
    const mo = FR[mm[1]], y = +mm[2];
    return `${y}-${String(mo).padStart(2,'0')}`;
  }
  return null;
}


// Affiche correctement des "secondes" en durée [h]:mm:ss dans Sheets
function setSecondsAsDuration_(sh, row, col, secs) {
  var cell = sh.getRange(row, col);
  if (cell.getFormula()) {
    Logger.log(
      `[SKIP] Préserve formule durée en ${cell.getA1Notation()} (${cell.getFormula()})`
    );
    return;
  }
  var days = (typeof secs === 'number' && !isNaN(secs)) ? secs / 86400 : 0;
  cell.setValue(days);
  cell.setNumberFormat('[h]:mm:ss');
}

// Écrit une valeur uniquement si la cellule NE contient PAS de formule
function setPreserveFormula_(sh, row, col, value, numberFormat) {
  const cell = sh.getRange(row, col);
  const formula = cell.getFormula();
  if (formula) {
    Logger.log(
      `[SKIP] Préserve formule en ${cell.getA1Notation()} -> ${formula}`
    );
    return;
  }
  cell.setValue(value);
  if (numberFormat) {
    cell.setNumberFormat(numberFormat);
  }
}

// Normalisation & repérage d’en-têtes
const HEADERS_ROW_ADS = 3;     // ligne d'en-têtes pour l’onglet "Google Ads"
function _normHeader_(s){
  return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[’'`]/g,'').replace(/\s+/g,' ').trim();
}
function findColByHeaderAliases_(sheet, aliases, headerRow){
  headerRow = headerRow || HEADERS_ROW_ADS || 3;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0]||[];
  const HN = headers.map(_normHeader_);
  const wanted = aliases.map(_normHeader_);
  for (let i=0;i<HN.length;i++){
    const h = HN[i]; if (!h) continue;
    if (wanted.some(w=>h===w || h.includes(w))) return i+1;
  }
  return 0;
}

// ➜ Crée la colonne si manquante (en-tête ligne HEADERS_ROW_ADS)
function ensureColByHeader_(sheet, wantedHeader) {
  const col = findColByHeaderAliases_(sheet, [wantedHeader], HEADERS_ROW_ADS);
  if (col) return col;
  const lastCol = sheet.getLastColumn();
  sheet.insertColumnAfter(lastCol);
  const newCol = lastCol + 1;
  sheet.getRange(HEADERS_ROW_ADS, newCol).setValue(wantedHeader);
  return newCol;
}

// Propriétés de script
function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**************************** GOOGLE ADS ****************************/

const SHEET_NAME_ADS = 'Google Ads';
const START_ROW_ADS = 6;

// ⚠️ Sans tirets
const CUSTOMER_ID = '3794858751';
const DEVELOPER_TOKEN = 'FRKxkVUQzqZSenRB3MC6EA';
const LOGIN_CUSTOMER_ID = '7343744510';

// Historique Google Ads (en mois)
const GADS_MONTHS_BACK = 24;

// OAuth client (même que GBM si tu réutilises le projet)
const CLIENT_ID = scriptProperties.getProperty('CLIENT_ID');
const CLIENT_SECRET = scriptProperties.getProperty('CLIENT_SECRET');

// AUTH
function getGoogleAdsAccessToken_() {
  const tokenUrl = 'https://oauth2.googleapis.com/token';
  const payload = {
    client_id: OAUTH_CLIENT_ID,
    client_secret: OAUTH_CLIENT_SECRET,
    refresh_token: getProp_('REFRESH_TOKEN'),
    grant_type: 'refresh_token'
  };
  const res = UrlFetchApp.fetch(tokenUrl, { method: 'post', payload, muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error(`OAuth HTTP ${code}: ${res.getContentText()}`);
  return JSON.parse(res.getContentText()).access_token;
}

// DATA
function fetchGoogleAdsMonthly(startDate, endDate) {
  // API v21 (sept 2025)
  const url = `https://googleads.googleapis.com/v21/customers/${CUSTOMER_ID}/googleAds:searchStream`;

  const query = `
    SELECT
      segments.month,
      metrics.impressions,
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros
    FROM customer
    WHERE segments.date BETWEEN '${startDate}' AND '${endDate}'
    ORDER BY segments.month
  `.trim();

  const accessToken = getGoogleAdsAccessToken_();
  const headers = {
    Authorization: 'Bearer ' + accessToken,
    'developer-token': DEVELOPER_TOKEN,
    'Content-Type': 'application/json'
  };
  if (LOGIN_CUSTOMER_ID) headers['login-customer-id'] = LOGIN_CUSTOMER_ID;

  const res = UrlFetchApp.fetch(url, { method: 'post', headers, payload: JSON.stringify({ query }), muteHttpExceptions: true });
  if (res.getResponseCode() >= 300) throw new Error(`HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
  return JSON.parse(res.getContentText()); // chunks
}

// Agrège les chunks searchStream par "YYYY-MM"
function aggregateAdsByMonth_(data) {
  const byMonth = {};
  (data || []).forEach((chunk) => {
    (chunk.results || []).forEach((r) => {
      let mk = r.segments && (r.segments.month || r.segments.date);
      if (!mk) return;
      const m = String(mk).match(/(\d{4})-(\d{2})/);
      if (!m) return;
      const key = `${m[1]}-${m[2]}`;

      const met = r.metrics || {};
      const costMicros = (met.costMicros != null ? met.costMicros :
                         (met.cost_micros != null ? met.cost_micros : 0));
      const impressions = Number(met.impressions || 0);
      const clicks      = Number(met.clicks || 0);

      if (!byMonth[key]) byMonth[key] = { impressions: 0, clicks: 0, cost: 0, calls: 0 };
      byMonth[key].impressions += impressions;
      byMonth[key].clicks      += clicks;
      byMonth[key].cost        += Number(costMicros) / 1e6; // micros -> devise
      const calls = Number(met.phoneCalls || met.calls || 0);
      if (calls) byMonth[key].calls += calls;
    });
  });

  Object.values(byMonth).forEach(v => {
    v.ctr = v.impressions ? v.clicks / v.impressions : 0;
    v.cpc = v.clicks ? v.cost / v.clicks : 0;
  });
  return byMonth;
}

// "YYYY-MM" -> "mois yyyy" FR
function monthKeyToFr_(ym) {
  const FR = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  const y = ym.slice(0,4);
  const m = Math.max(0, Math.min(11, parseInt(ym.slice(5,7),10)-1));
  return `${FR[m]} ${y}`;
}

// Style pour une ligne "année"
function _adsStyleYearRow_(sh, row, moisCol) {
  sh.getRange(row, 1, 1, sh.getLastColumn()).setBackground('#e6e1f5');
  sh.getRange(row, moisCol).setFontWeight('bold');
}

// Trouve la ligne existante d’un mois "YYYY-MM" (sans insertion)
function _adsFindExistingMonthRow_(sh, moisCol, targetYM) {
  const last = sh.getLastRow();
  if (last < START_ROW_ADS) return 0;
  for (let r = START_ROW_ADS; r <= last; r++) {
    const v = sh.getRange(r, moisCol).getValue();
    if (isYearSeparatorRow_(v)) continue;
    const ym = sheetCellToYYYYMM_(v);
    if (ym === targetYM) return r;
  }
  return 0;
}

/**
 * Garantit qu’une ligne existe pour le mois targetYM
 * - Crée la ligne année si nécessaire (2024, 2025…)
 * - Ajoute le mois à la bonne place dans l’année
 * - Ne touche pas aux autres colonnes (les formules restent)
 */
function _adsFindOrInsertMonthRow_(sh, moisCol, targetYM) {
  const existing = _adsFindExistingMonthRow_(sh, moisCol, targetYM);
  if (existing) {
    Logger.log(`[Ads/ensureRow] Ligne déjà présente pour ${targetYM} → ${existing}`);
    return existing;
  }

  const targetYear = parseInt(targetYM.slice(0, 4), 10);
  let lastRow = Math.max(sh.getLastRow(), START_ROW_ADS - 1);
  let yearRow = null;

  // Cherche si une ligne "année" existe déjà
  for (let r = START_ROW_ADS; r <= lastRow; r++) {
    const v = sh.getRange(r, moisCol).getValue();
    if (isYearSeparatorRow_(v)) {
      const y = parseInt(String(v).trim(), 10);
      if (y === targetYear) {
        yearRow = r;
        break;
      }
    }
  }

  // Si pas de ligne année → on l’ajoute (au bon endroit parmi les autres années)
  if (!yearRow) {
    let insertAt = lastRow + 1;
    for (let r = START_ROW_ADS; r <= lastRow; r++) {
      const v = sh.getRange(r, moisCol).getValue();
      if (isYearSeparatorRow_(v)) {
        const y = parseInt(String(v).trim(), 10);
        if (y > targetYear) {
          insertAt = r;
          break;
        }
      }
    }
    sh.insertRowBefore(insertAt);
    yearRow = insertAt;
    sh.getRange(yearRow, moisCol).setValue(String(targetYear));
    _adsStyleYearRow_(sh, yearRow, moisCol);
    lastRow++; // on a ajouté une ligne
  }

  // Cherche la position où insérer le mois dans cette année
  let insertAt = yearRow + 1;
  for (let r = yearRow + 1; r <= lastRow; r++) {
    const v = sh.getRange(r, moisCol).getValue();
    if (isYearSeparatorRow_(v)) break; // année suivante
    const ym = sheetCellToYYYYMM_(v);
    if (!ym) continue;
    if (ym > targetYM) {
      insertAt = r;
      break;
    }
    insertAt = r + 1; // après le dernier mois rencontré
  }

  sh.insertRowBefore(insertAt);
  Logger.log(`[Ads/ensureRow] Insertion nouvelle ligne pour ${targetYM} → ${insertAt}`);
  return insertAt;
}


// ÉCRITURE (in-place, préserve D/E/F/G — CTR sera recalculé ensuite)
function writeAdsMonthlyToSheetFlexible(data) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const byMonth = aggregateAdsByMonth_(data);
  const monthKeys = Object.keys(byMonth).sort();
  if (!monthKeys.length) throw new Error('[Ads] Aucune donnée agrégée (0 mois).');

  const moisCol   = findColByHeaderAliases_(sh, ['mois'], HEADERS_ROW_ADS);
  const budgetCol = findColByHeaderAliases_(sh, ['budget investi','depenses','depense','cout','coût','spend','cost'], HEADERS_ROW_ADS);
  const impCol    = findColByHeaderAliases_(sh, ['impressions','nb impressions','nombre dimpressions','nombre d impressions'], HEADERS_ROW_ADS);
  const clkCol    = findColByHeaderAliases_(sh, ['clics','clicks','nb clics','nombre de clics'], HEADERS_ROW_ADS);
  const cpcCol    = findColByHeaderAliases_(sh, ['cpc','cout par clic','coût par clic','cost per click'], HEADERS_ROW_ADS);
  const callsCol  = findColByHeaderAliases_(sh, ['appels google ads','appels','calls','nombre dappels'], HEADERS_ROW_ADS);

  const missing = [];
  if (!moisCol)   missing.push('Mois');
  if (!budgetCol) missing.push('Budget investi');
  if (!impCol)    missing.push('Impressions');
  if (!clkCol)    missing.push('Clics');
  if (!cpcCol)    missing.push('CPC');
  if (missing.length) throw new Error('[Ads] Colonnes manquantes: ' + missing.join(', '));

  let wrote = 0;
  for (const k of monthKeys) {
    const v = byMonth[k];
    const row = _adsFindOrInsertMonthRow_(sh, moisCol, k);

    // Mois = texte => en général pas une formule, mais on log si jamais
    const moisCell = sh.getRange(row, moisCol);
    if (!moisCell.getFormula()) {
      moisCell.setValue(monthKeyToFr_(k));
    } else {
      Logger.log(`[SKIP] Préserve formule en ${moisCell.getA1Notation()}`);
    }

    setPreserveFormula_(sh, row, budgetCol, v.cost, '0.00 €');
    setPreserveFormula_(sh, row, impCol,    v.impressions, null);
    setPreserveFormula_(sh, row, clkCol,    v.clicks, null);
    setPreserveFormula_(sh, row, cpcCol,    v.cpc, '0.00 €');
    if (callsCol) {
      setPreserveFormula_(sh, row, callsCol, v.calls || 0, null);
    }
    wrote++;
  }
  Logger.log(`[Ads in-place] Mois écrits/MAJ: ${wrote}`);
}


// JOB Ads — Full history
function run_GAds_FullHistory() {
  const today = new Date();
  const start = FULLHIST_START_LOCAL;                                // ← forcer Jan 2024
  const end   = new Date(today.getFullYear(), today.getMonth(), 0);
  const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const endStr   = Utilities.formatDate(end,   Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const data = fetchGoogleAdsMonthly(startStr, endStr);
  writeAdsMonthlyToSheetFlexible(data);
}

// JOB Ads — N-1 ciblé
function run_GAds_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const moisCol   = findColByHeaderAliases_(sh, ['mois'], HEADERS_ROW_ADS);
  const budgetCol = findColByHeaderAliases_(sh, ['budget investi','depenses','depense','cout','coût','spend','cost'], HEADERS_ROW_ADS);
  const impCol    = findColByHeaderAliases_(sh, ['impressions','nb impressions','nombre dimpressions','nombre d impressions'], HEADERS_ROW_ADS);
  const clkCol    = findColByHeaderAliases_(sh, ['clics','clicks','nb clics','nombre de clics'], HEADERS_ROW_ADS);
  const ctrCol    = findColByHeaderAliases_(sh, ['ctr','taux de clics','click through rate'], HEADERS_ROW_ADS);
  const cpcCol    = findColByHeaderAliases_(sh, ['cpc','cout par clic','coût par clic','cost per click'], HEADERS_ROW_ADS);
  const callsCol  = findColByHeaderAliases_(sh, ['appels google ads','appels','calls','nombre dappels'], HEADERS_ROW_ADS);

  const missing = [];
  if (!moisCol) missing.push('Mois');
  if (!budgetCol) missing.push('Budget investi');
  if (!impCol) missing.push('Impressions');
  if (!clkCol) missing.push('Clics');
  if (!ctrCol) missing.push('CTR');
  if (!cpcCol) missing.push('CPC');
  if (missing.length) throw new Error('[GAds AddLastMonth] Colonnes manquantes: ' + missing.join(', '));

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end   = new Date(today.getFullYear(), today.getMonth(), 0);
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');

  const data = fetchGoogleAdsMonthly(
    Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    Utilities.formatDate(end,   Session.getScriptTimeZone(), 'yyyy-MM-dd')
  );
  const byMonth = aggregateAdsByMonth_(data);
  const v = byMonth[ymKey];
  if (!v) { Logger.log('[GAds AddLastMonth] Aucune donnée renvoyée pour ' + ymKey); return; }

  const row = _adsFindOrInsertMonthRow_(sh, moisCol, ymKey);

  const moisCell = sh.getRange(row, moisCol);
  if (!moisCell.getFormula()) {
    moisCell.setValue(monthKeyToFr_(ymKey));
  } else {
    Logger.log(`[SKIP] Préserve formule en ${moisCell.getA1Notation()}`);
  }

  setPreserveFormula_(sh, row, budgetCol, v.cost, '0.00');
  setPreserveFormula_(sh, row, impCol,    v.impressions, null);
  setPreserveFormula_(sh, row, clkCol,    v.clicks, null);
  setPreserveFormula_(sh, row, ctrCol,    v.ctr, '0.00%');
  setPreserveFormula_(sh, row, cpcCol,    v.cpc, '0.00');
  if (callsCol) setPreserveFormula_(sh, row, callsCol, v.calls || 0, null);

  Logger.log(`[GAds AddLastMonth] ${ymKey} -> cost=${v.cost}, imp=${v.impressions}, clk=${v.clicks}, ctr=${v.ctr}, cpc=${v.cpc}`);
}


/******************** Google Ads — Recalcul : seulement CTR & Taux de conversion ********************/

function _adsFindDerivedCols_(sh) {
  return {
    // lecture uniquement
    clicks:  findColByHeaderAliases_(sh, ['clics','clicks','nb clics','nombre de clics'], HEADERS_ROW_ADS),
    impr:    findColByHeaderAliases_(sh, ['impressions','nb impressions','nombre dimpressions','nombre d impressions'], HEADERS_ROW_ADS),
    appels:  findColByHeaderAliases_(sh, ["nombre d'appels","nombre dappels","appels","appels google ads"], HEADERS_ROW_ADS),
    forms:   findColByHeaderAliases_(sh, ['nombre de formulaire','formulaires','forms','paperform','monday forms','submissions','soumissions'], HEADERS_ROW_ADS),

    // écriture : seulement Taux de conversion & CTR (plus de CPL, ne touche pas "Nombre de contacts")
    tauxConv:findColByHeaderAliases_(sh, ['taux de conversion','conversion rate','cr'], HEADERS_ROW_ADS),
    ctrCol:  findColByHeaderAliases_(sh, ['ctr','taux de clics','click through rate'], HEADERS_ROW_ADS)
  };
}

function _adsRecomputeDerivedForRow_(sh, r, cols) {
  const moisVal = sh.getRange(r, 1).getValue();
  if (isYearSeparatorRow_(moisVal)) return;
  if (!sheetCellToYYYYMM_(moisVal)) return;

  const getN = (c) => (c ? Number(sh.getRange(r, c).getValue() || 0) : 0);

  const clicks = getN(cols.clicks);
  const impr   = getN(cols.impr);
  const appels = getN(cols.appels);
  const forms  = getN(cols.forms);

  const contactsLocal = appels + forms;
  const tauxConv = clicks > 0 ? (contactsLocal / clicks) : 0;
  const ctr      = impr > 0   ? (clicks / impr)          : 0;

  if (cols.tauxConv) {
    setPreserveFormula_(sh, r, cols.tauxConv, tauxConv, '0.00%');
  }
  if (cols.ctrCol) {
    setPreserveFormula_(sh, r, cols.ctrCol, ctr, '0.00%');
  }
}


function run_Ads_Derived_RecomputeAll() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);
  const last = sh.getLastRow();
  if (last < START_ROW_ADS) return;

  const cols = _adsFindDerivedCols_(sh);
  for (let r = START_ROW_ADS; r <= last; r++) _adsRecomputeDerivedForRow_(sh, r, cols);
  Logger.log('[Derived] Recalcul CTR/CR terminé.');
}

function run_Ads_Derived_RecomputeLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const cols = _adsFindDerivedCols_(sh);
  const last = sh.getLastRow();
  if (last < START_ROW_ADS) return;

  const today = new Date();
  const targetYM = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth()-1, 1),
                                        Session.getScriptTimeZone(), 'yyyy-MM');

  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r, 1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    if (sheetCellToYYYYMM_(val) === targetYM) {
      _adsRecomputeDerivedForRow_(sh, r, cols);
      Logger.log(`[Derived] N-1 recalculé (${targetYM})`);
      break;
    }
  }
}
function run_Ads_Derived_RecomputeCurrentMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const cols = _adsFindDerivedCols_(sh);
  const last = sh.getLastRow();
  if (last < START_ROW_ADS) return;

  const today = new Date();
  const targetYM = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), 1),
                                        Session.getScriptTimeZone(), 'yyyy-MM');

  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r, 1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    if (sheetCellToYYYYMM_(val) === targetYM) {
      _adsRecomputeDerivedForRow_(sh, r, cols);
      Logger.log(`[Derived] Mois en cours recalculé (${targetYM})`);
      break;
    }
  }
}

/**************************** MAGNETIS → "Nombre d’appels" ****************************/

const MAG_TIMEZONE = 'Europe/Paris';
const MAG_NUMBER_IDS_ADS = []; // si tu veux limiter à certains numéros
const MAG_ADS_CHANNEL_NAMES = ['google ads','adwords','google/cpc','paid search','sea'];
const MAG_CHANNEL_MATCH = 'equals'; // 'equals' | 'includes' | 'regex'
const MAG_COUNT_ANSWERED_ONLY = false;

function getMagApiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('MAGNETIS_API_KEY');
  if (!k) throw new Error("Définis la propriété de script 'MAGNETIS_API_KEY'.");
  return k;
}
function valAt_(obj, path) {
  return path.split('.').reduce((o, k) => (o && o[k] != null ? o[k] : undefined), obj);
}
function getChannelName_(c) {
  const candidates = ['channel_name','channel','analysis.channel','analytics.channel','utm.channel','session.channel'];
  for (const p of candidates) {
    const v = valAt_(c, p);
    if (v != null && String(v).trim() !== '') return String(v).trim().toLowerCase();
  }
  return '';
}
function channelMatchesAds_(name) {
  if (!name) return false;
  const list = MAG_ADS_CHANNEL_NAMES.map(x => String(x).toLowerCase());
  switch (MAG_CHANNEL_MATCH) {
    case 'includes': return list.some(a => name.includes(a));
    case 'regex':    return list.some(rx => new RegExp(rx, 'i').test(name));
    default:         return list.some(a => name === a);
  }
}
function isGoogleAdsCall_(c) {
  if (MAG_NUMBER_IDS_ADS && MAG_NUMBER_IDS_ADS.length) {
    const numId = String(c.number_id || c.numberId || c.tracked_number_id || c.tracking_number_id || '');
    if (!MAG_NUMBER_IDS_ADS.includes(numId)) return false;
  }
  const ch = getChannelName_(c);
  return channelMatchesAds_(ch);
}

function magnetisFetchCalls_(fromDate, toDate, extraParams = {}) {
  const apiKey = getMagApiKey_();
  const base = 'https://api.magnetis.io/calls';
  const fmtUTC = (d) => Utilities.formatDate(d, 'UTC', 'yyyyMMddHHmmss');

  let start = new Date(fromDate), end = new Date(toDate);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) throw new Error('Dates invalides (Magnetis)');
  if (start > end) { const t=start; start=end; end=t; }

  const fromStr = fmtUTC(start), toStr = fmtUTC(end);

  let page = 1, all = [];
  while (true) {
    const params = { from: fromStr, to: toStr, limit: 250, page, analysis: 1, ...extraParams };
    const qs = Object.keys(params).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`).join('&');
    const res = UrlFetchApp.fetch(`${base}?${qs}`, { method: 'get', headers: { 'x-api-key': apiKey, 'Accept': 'application/json' }, muteHttpExceptions: true });
    if (res.getResponseCode() >= 300) throw new Error(`Magnetis HTTP ${res.getResponseCode()}: ${res.getContentText()}`);

    const json = JSON.parse(res.getContentText());
    const data = json.data || json || [];
    if (!data.length) break;
    all = all.concat(data);

    page++;
    if (json.links && !json.links.next) break;
    if (!json.links && data.length < 250) break;
  }
  return all;
}
function magnetisAggregateMonthlyCounts_(calls) {
  const bucket = {};
  calls.forEach(c => {
    if (!isGoogleAdsCall_(c)) return;
    const rawDate = (c.start_at || c.created_at || c.date || c.started_at);
    if (!rawDate) return;
    const d = new Date(rawDate);
    if (isNaN(d.getTime())) return;
    const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;

    if (MAG_COUNT_ANSWERED_ONLY) {
      const duration = Number(c.duration || (c.analysis && c.analysis.duration) || 0);
      if (!(duration > 0)) return;
    }
    if (!bucket[key]) bucket[key] = 0;
    bucket[key] += 1;
  });
  return bucket;
}

function run_MagnetisToAds_SyncForSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const CALLS_COL = findColByHeaderAliases_(sh, ["nombre d'appels","nombre dappels","appels","appels google ads"], HEADERS_ROW_ADS);
  if (!CALLS_COL) throw new Error(`Colonne 'Nombre d’appels' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  const today = new Date();
  const from = FULLHIST_START_LOCAL;                                   // ← Jan 2024
  const to   = new Date(today.getFullYear(), today.getMonth(), 0);

  const callsByMonth = magnetisAggregateMonthlyCounts_(magnetisFetchCalls_(from, to));
  const last = sh.getLastRow();

  for (let r = START_ROW_ADS; r <= last; r++) {
    const cell = sh.getRange(r, 1).getValue();
    if (isYearSeparatorRow_(cell)) continue;
    const ym = sheetCellToYYYYMM_(cell);
    if (!ym) continue;
    sh.getRange(r, CALLS_COL).setValue(callsByMonth[ym] || 0);
  }
}

function run_MagnetisToAds_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const CALLS_COL = findColByHeaderAliases_(sh, ["nombre d'appels","nombre dappels","appels","appels google ads"], HEADERS_ROW_ADS);
  if (!CALLS_COL) throw new Error(`Colonne 'Nombre d’appels' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  const today = new Date();
  const monthKey = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth()-1, 1), Session.getScriptTimeZone(), 'yyyy-MM');
  const start = new Date(today.getFullYear(), today.getMonth()-1, 1);
  const end   = new Date(today.getFullYear(), today.getMonth(), 0);

  const byMonth = magnetisAggregateMonthlyCounts_(magnetisFetchCalls_(start, end));
  const n = byMonth[monthKey] || 0;

  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === monthKey) {
      sh.getRange(r, CALLS_COL).setValue(n);
      break;
    }
  }
}
function run_MagnetisToAds_CurrentMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const CALLS_COL = findColByHeaderAliases_(sh, ["nombre d'appels","nombre dappels","appels","appels google ads"], HEADERS_ROW_ADS);
  if (!CALLS_COL) throw new Error(`Colonne 'Nombre d’appels' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const end   = today;
  const monthKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');

  const byMonth = magnetisAggregateMonthlyCounts_(magnetisFetchCalls_(start, end));
  const n = byMonth[monthKey] || 0;

  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === monthKey) {
      sh.getRange(r, CALLS_COL).setValue(n);
      break;
    }
  }
}

/**************************** PAPERFORM → helper (compte par mois) ****************************/

const PAPERFORM_FORM_ID = 'zkzwbmgp';

function getPaperformToken_() {
  const props = PropertiesService.getScriptProperties();
  const t = props.getProperty('PAPARFORM_TOKEN') || props.getProperty('PAPERFORM_TOKEN');
  if (!t) throw new Error("Définis la propriété 'PAPARFORM_TOKEN' (ou 'PAPERFORM_TOKEN').");
  return t;
}

function extractPaperformArray_(json) {
  if (!json) return [];
  if (Array.isArray(json)) return json;
  if (json.results) {
    if (Array.isArray(json.results)) return json.results;
    if (Array.isArray(json.results.submissions)) return json.results.submissions;
    if (Array.isArray(json.results.data)) return json.results.data;
    for (const k in json.results) {
      if (Array.isArray(json.results[k])) return json.results[k];
    }
  }
  if (Array.isArray(json.data)) return json.data;
  for (const k in json) {
    if (Array.isArray(json[k])) return json[k];
  }
  return [];
}

// "Jun 27th 25, 6:04am" → Date UTC
function parsePaperformHumanDateToUTC_(str) {
  if (!str || typeof str !== 'string') return new Date(NaN);
  const s = str.trim().replace(/\s+/g, ' ');
  const MM = { jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,oct:9,october:9,nov:10,november:10,dec:11,december:11 };
  const rx = /^([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\d{2,4})[, ]\s*(\d{1,2}):(\d{2})\s*(am|pm)$/i;
  const m = s.match(rx);
  if (!m) return new Date(NaN);
  const monName = m[1].toLowerCase();
  const day = parseInt(m[2], 10);
  let year = parseInt(m[3], 10);
  let hour = parseInt(m[4], 10);
  const min = parseInt(m[5], 10);
  const ap = m[6].toLowerCase();
  const mon = MM[monName];
  if (mon == null || isNaN(day) || isNaN(year) || isNaN(hour) || isNaN(min)) return new Date(NaN);
  if (year < 100) year = 2000 + year;
  hour = hour % 12;
  if (ap === 'pm') hour += 12;
  return new Date(Date.UTC(year, mon, day, hour, min, 0));
}

// Détermine une date pour une soumission Paperform (plusieurs champs possibles)
function getSubmissionDate_(s) {
  const candidates = [
    s.submitted_at, s.submittedAt, s.submitted_at_utc, s.submittedAtUTC,
    s.created_at,   s.createdAt,   s.created_at_utc,   s.createdAtUTC,
    s.date,         s.timestamp,   s.time_submitted,   s.timeSubmitted,
    s.created,      s.submitted,   (s.meta && (s.meta.submitted_at || s.meta.created_at))
  ];
  for (let v of candidates) {
    if (v == null) continue;
    if (typeof v === 'number') {
      const ms = v < 1e12 ? v * 1000 : v;
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d;
      continue;
    }
    if (typeof v === 'string') {
      const dStd = new Date(v);
      if (!isNaN(dStd.getTime())) return dStd;
      const dHuman = parsePaperformHumanDateToUTC_(v);
      if (dHuman && !isNaN(dHuman.getTime())) return dHuman;
      const m = v.match(/^\d+$/);
      if (m) {
        const n = Number(v), ms = n < 1e12 ? n * 1000 : n;
        const dNum = new Date(ms);
        if (!isNaN(dNum.getTime())) return dNum;
      }
    }
  }
  return new Date(NaN);
}

function paperformFetchSubmissions_(slugOrId, fromDate, toDate) {
  const token = getPaperformToken_();
  const base = `https://api.paperform.co/v1/forms/${encodeURIComponent(slugOrId)}/submissions`;

  const limit = 100;
  let all = [];
  const fromTs = fromDate.getTime(), toTs = toDate.getTime();
  let mode = 'skip', skip = 0, page = 1;

  while (true) {
    let url = (mode === 'skip') ? `${base}?limit=${limit}&skip=${skip}` : `${base}?limit=${limit}&page=${page}`;
    let res = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' }, muteHttpExceptions: true });

    if (res.getResponseCode() === 422 && mode === 'skip') { mode = 'page'; page = 1; continue; }
    if (res.getResponseCode() >= 300) throw new Error(`Paperform HTTP ${res.getResponseCode()}: ${res.getContentText()}`);

    const json = JSON.parse(res.getContentText());
    const batch = extractPaperformArray_(json);
    if (!batch.length) break;

    for (const s of batch) {
      const d = getSubmissionDate_(s);
      if (isNaN(d.getTime())) continue;
      const t = d.getTime();
      if (t >= fromTs && t <= toTs) all.push(s);
    }

    if (mode === 'skip') skip += batch.length; else page += 1;
    if (batch.length < limit) break;
    Utilities.sleep(120);
  }
  return all;
}

function paperformAggregateMonthlyCounts_(subs) {
  const bucket = {};
  subs.forEach(s => {
    const d = getSubmissionDate_(s);
    if (isNaN(d.getTime())) return;
    const key = d.getUTCFullYear() + '-' + String(d.getUTCMonth()+1).padStart(2, '0');
    if (!bucket[key]) bucket[key] = 0;
    bucket[key] += 1;
  });
  return bucket;
}

/***** MONDAY (board mixte) → helpers comptage *****/

// ====== CONFIG ======
const MONDAY_BOARD_ID = 9950271520; // ← ID du board Monday

// Titre (ou ID direct) de la colonne dédiée au formulaire Google Ads
const MONDAY_FORM_COLUMN_TITLE_GOOGLE_ADS = 'Google Ads';
const MONDAY_FORM_COLUMN_ID_GOOGLE_ADS = '';

// Mode de filtrage : 'nonempty' (non vide), 'equals' (texte EXACT), 'includes' (contient)
const MONDAY_MATCH_MODE = 'nonempty';
const MONDAY_MATCH_VALUES = []; // si equals/includes

/***** MONDAY LEADS (Google Ads) — config *****/
const MONDAY_LEADS_BOARD_ID = MONDAY_BOARD_ID; // même board
const MONDAY_LEADS_COL_SOURCE = 'Source du Lead';
const MONDAY_LEADS_COL_TYPE   = 'FORMULAIRE / APPELS';
const MONDAY_LEADS_COL_STATUS = 'Nature du contact';

// égalité stricte (insensible casse/accents)
const MONDAY_LEADS_SOURCE_EQUALS = 'LP (via Google Ads)';
// status = Lead (peut accepter variations ex: "lead")
const MONDAY_LEADS_STATUS_MATCH = ['lead'];

// reconnaissance du type
const MONDAY_LEADS_TYPE_CALL_MATCH = ['Appel','call','téléphone'];
const MONDAY_LEADS_TYPE_FORM_MATCH = ['Formulaire','form','paperform'];

function _textEquals_(txt, want){
  const a = _norm_(txt||'');
  const b = _norm_(want||'');
  return a === b;
}
function _textIncludesAny_(txt, arr){
  const t = _norm_(txt||'');
  return (arr||[]).some(v => t.includes(_norm_(v)));
}

function getMondayToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('MONDAY_TOKEN');
  if (!t) throw new Error("Définis la propriété de script 'MONDAY_TOKEN' (token API Monday).");
  return t;
}
function mondayGraphQL_(query, variables) {
  const res = UrlFetchApp.fetch('https://api.monday.com/v2', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': getMondayToken_(),
      'API-Version': '2023-10'
    },
    payload: JSON.stringify({ query, variables: variables || {} }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error(`Monday HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
  }
  const body = JSON.parse(res.getContentText());
  if (body.errors) throw new Error('Monday GraphQL error: ' + JSON.stringify(body.errors));
  return body.data;
}
function _norm_(s) {
  return String(s||'').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[’'`]/g,'').replace(/\s+/g,' ').trim();
}
function mondayResolveColumnIdByTitleOrId_(boardId, titleOrId) {
  if (!titleOrId) return '';
  const looksLikeId = /^[a-z0-9_]+$/i.test(titleOrId) && !_norm_(titleOrId).includes(' ');
  if (looksLikeId) return titleOrId;
  const want = _norm_(titleOrId);

  const q = `
    query($bid:[ID!]) {
      boards(ids:$bid){
        columns { id title type }
      }
    }`;
  const data = mondayGraphQL_(q, { bid: [Number(boardId)] });
  const cols = (data && data.boards && data.boards[0] && data.boards[0].columns) || [];
  let best = '';
  for (const c of cols) {
    if (_norm_(c.title) === want) { best = c.id; break; }
  }
  if (!best) {
    for (const c of cols) {
      if (_norm_(c.title).includes(want)) { best = c.id; break; }
    }
  }
  return best;
}
function mondayValueMatches_(colValueText) {
  const txt = _norm_(colValueText||'');
  switch (MONDAY_MATCH_MODE) {
    case 'equals':
      return MONDAY_MATCH_VALUES.some(v => txt === _norm_(v));
    case 'includes':
      return MONDAY_MATCH_VALUES.some(v => txt.includes(_norm_(v)));
    default: // 'nonempty'
      return txt.length > 0;
  }
}
function mondayFetchFormCountsByMonth_(boardId, columnIdOrTitle, fromDateUTC, toDateUTC) {
  const colId = MONDAY_FORM_COLUMN_ID_GOOGLE_ADS
    ? MONDAY_FORM_COLUMN_ID_GOOGLE_ADS
    : mondayResolveColumnIdByTitleOrId_(boardId, columnIdOrTitle);

  if (!colId) throw new Error("Impossible de résoudre l'ID de la colonne du formulaire Google Ads.");

  const counts = {};
  let cursor = null;

  do {
    const query = `
      query($bid:[ID!], $cursor:String, $colId:[String!]) {
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id
              state
              created_at
              column_values(ids:$colId){ id text }
            }
          }
        }
      }`;
    const vars = { bid: [Number(boardId)], cursor, colId: [colId] };

    const data = mondayGraphQL_(query, vars);
    const page = data && data.boards && data.boards[0] && data.boards[0].items_page;
    if (!page) break;

    (page.items || []).forEach(it => {
      if (String(it.state||'').toLowerCase() === 'archived') return;
      const d = new Date(it.created_at);
      if (isNaN(d.getTime())) return;
      if (d < fromDateUTC || d > toDateUTC) return;

      const cv = (it.column_values || [])[0];
      const ok = mondayValueMatches_(cv && cv.text);
      if (!ok) return;

      const key = d.getUTCFullYear() + '-' + String(d.getUTCMonth()+1).padStart(2,'0');
      counts[key] = (counts[key] || 0) + 1;
    });

    cursor = page.cursor;
  } while (cursor);

  return counts;
}

// { 'YYYY-MM': n } — APPELS lead (source == "LP (via Google Ads)" && status ∈ ["lead"] && type ~ appel)
function mondayLeadCallsCountsByMonth_Ads_(boardId, colSource, colType, colStatus, fromUTC, toUTC){
  const sourceId = mondayResolveColumnIdByTitleOrId_(boardId, colSource);
  const typeId   = mondayResolveColumnIdByTitleOrId_(boardId, colType);
  const statusId = mondayResolveColumnIdByTitleOrId_(boardId, colStatus);
  if (!sourceId || !typeId || !statusId) throw new Error('[Monday leads calls] Colonnes introuvables (source/type/status).');

  const counts = {}; let cursor = null;
  do{
    const q = `
      query($bid:[ID!], $cursor:String, $cols:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id state created_at
              column_values(ids:$cols){ id text }
            }
          }
        }
      }`;
    const d = mondayGraphQL_(q, { bid:[Number(boardId)], cursor, cols:[sourceId,typeId,statusId] });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page; if (!page) break;

    (page.items||[]).forEach(it=>{
      if (String(it.state||'').toLowerCase()==='archived') return;
      const when = new Date(it.created_at); if (isNaN(when.getTime())) return;
      if (when < fromUTC || when > toUTC) return;

      const cv = it.column_values || [];
      const src = (cv.find(c=>c.id===sourceId)||{}).text || '';
      const typ = (cv.find(c=>c.id===typeId)  ||{}).text || '';
      const sts = (cv.find(c=>c.id===statusId)||{}).text || '';

      if (!_textEquals_(src, MONDAY_LEADS_SOURCE_EQUALS)) return;
      if (!_textIncludesAny_(sts, MONDAY_LEADS_STATUS_MATCH)) return;
      if (!_textIncludesAny_(typ, MONDAY_LEADS_TYPE_CALL_MATCH)) return;

      const key = when.getUTCFullYear()+'-'+String(when.getUTCMonth()+1).padStart(2,'0');
      counts[key] = (counts[key]||0) + 1;
    });
    cursor = page.cursor;
  } while (cursor);

  return counts;
}

// { 'YYYY-MM': n } — FORMULAIRES lead (source == "LP (via Google Ads)" && status ∈ ["lead"] && type ~ formulaire)
function mondayLeadFormsCountsByMonth_Ads_(boardId, colSource, colType, colStatus, fromUTC, toUTC){
  const sourceId = mondayResolveColumnIdByTitleOrId_(boardId, colSource);
  const typeId   = mondayResolveColumnIdByTitleOrId_(boardId, colType);
  const statusId = mondayResolveColumnIdByTitleOrId_(boardId, colStatus);
  if (!sourceId || !typeId || !statusId) throw new Error('[Monday leads forms] Colonnes introuvables (source/type/status).');

  const counts = {}; let cursor = null;
  do{
    const q = `
      query($bid:[ID!], $cursor:String, $cols:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id state created_at
              column_values(ids:$cols){ id text }
            }
          }
        }
      }`;
    const d = mondayGraphQL_(q, { bid:[Number(boardId)], cursor, cols:[sourceId,typeId,statusId] });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page; if (!page) break;

    (page.items||[]).forEach(it=>{
      if (String(it.state||'').toLowerCase()==='archived') return;
      const when = new Date(it.created_at); if (isNaN(when.getTime())) return;
      if (when < fromUTC || when > toUTC) return;

      const cv = it.column_values || [];
      const src = (cv.find(c=>c.id===sourceId)||{}).text || '';
      const typ = (cv.find(c=>c.id===typeId)  ||{}).text || '';
      const sts = (cv.find(c=>c.id===statusId)||{}).text || '';

      if (!_textEquals_(src, MONDAY_LEADS_SOURCE_EQUALS)) return;
      if (!_textIncludesAny_(sts, MONDAY_LEADS_STATUS_MATCH)) return;
      if (!_textIncludesAny_(typ, MONDAY_LEADS_TYPE_FORM_MATCH)) return;

      const key = when.getUTCFullYear()+'-'+String(when.getUTCMonth()+1).padStart(2,'0');
      counts[key] = (counts[key]||0) + 1;
    });
    cursor = page.cursor;
  } while (cursor);

  return counts;
}

/************** Monday Leads → écrire colonnes lead **************/

function run_MondayLeadsToAds_SyncForSheet(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  // ➜ crée / trouve les colonnes demandées
  const CALLS_LEAD_COL = ensureColByHeader_(sh, "Nombre d'appels lead");
  const FORMS_LEAD_COL = ensureColByHeader_(sh, "Nombre de formulaire lead");

  // Fenêtre forcée : Jan 2024 → fin du mois précédent (UTC)
  const from = FULLHIST_START_UTC;
  const now  = new Date();
  const to   = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 0, 23,59,59));

  let callsByMonth = {}, formsByMonth = {};
  try {
    callsByMonth = mondayLeadCallsCountsByMonth_Ads_(MONDAY_LEADS_BOARD_ID, MONDAY_LEADS_COL_SOURCE, MONDAY_LEADS_COL_TYPE, MONDAY_LEADS_COL_STATUS, from, to);
  } catch(e){ Logger.log('[Leads Ads] Calls KO: ' + e); }
  try {
    formsByMonth = mondayLeadFormsCountsByMonth_Ads_(MONDAY_LEADS_BOARD_ID, MONDAY_LEADS_COL_SOURCE, MONDAY_LEADS_COL_TYPE, MONDAY_LEADS_COL_STATUS, from, to);
  } catch(e){ Logger.log('[Leads Ads] Forms KO: ' + e); }

  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++){
    const cell = sh.getRange(r, 1).getValue();
    if (isYearSeparatorRow_(cell)) continue;
    const ym = sheetCellToYYYYMM_(cell);
    if (!ym) continue; // ← ne rien écrire hors lignes mois
    sh.getRange(r, CALLS_LEAD_COL).setValue(callsByMonth[ym] || 0);
    sh.getRange(r, FORMS_LEAD_COL).setValue(formsByMonth[ym] || 0);
  }
}

function run_MondayLeadsToAds_AddLastMonth(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  // ➜ crée / trouve les colonnes demandées
  const CALLS_LEAD_COL = ensureColByHeader_(sh, "Nombre d'appels lead");
  const FORMS_LEAD_COL = ensureColByHeader_(sh, "Nombre de formulaire lead");

  const now = new Date();
  const y = now.getUTCFullYear(), m = now.getUTCMonth();
  const start = new Date(Date.UTC(y, m-1, 1, 0,0,0));
  const end   = new Date(Date.UTC(y, m,   0, 23,59,59));
  const ymKey = start.getUTCFullYear() + '-' + String(start.getUTCMonth()+1).padStart(2,'0');

  let nCalls = 0, nForms = 0;
  try {
    const c = mondayLeadCallsCountsByMonth_Ads_(MONDAY_LEADS_BOARD_ID, MONDAY_LEADS_COL_SOURCE, MONDAY_LEADS_COL_TYPE, MONDAY_LEADS_COL_STATUS, start, end);
    nCalls = c[ymKey] || 0;
  } catch(e){ Logger.log('[Leads Ads N-1] Calls KO: ' + e); }
  try {
    const f = mondayLeadFormsCountsByMonth_Ads_(MONDAY_LEADS_BOARD_ID, MONDAY_LEADS_COL_SOURCE, MONDAY_LEADS_COL_TYPE, MONDAY_LEADS_COL_STATUS, start, end);
    nForms = f[ymKey] || 0;
  } catch(e){ Logger.log('[Leads Ads N-1] Forms KO: ' + e); }

  const last = sh.getLastRow();
  for (let r=START_ROW_ADS; r<=last; r++){
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === ymKey){
      sh.getRange(r, CALLS_LEAD_COL).setValue(nCalls);
      sh.getRange(r, FORMS_LEAD_COL).setValue(nForms);
      break;
    }
  }
}
function run_MondayLeadsToAds_CurrentMonth(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const CALLS_LEAD_COL = ensureColByHeader_(sh, "Nombre d'appels lead");
  const FORMS_LEAD_COL = ensureColByHeader_(sh, "Nombre de formulaire lead");

  const now = new Date();
  const y = now.getUTCFullYear(), m = now.getUTCMonth();
  const start = new Date(Date.UTC(y, m, 1, 0,0,0));
  const end   = new Date(Date.UTC(y, m, now.getUTCDate(), 23,59,59));
  const ymKey = start.getUTCFullYear() + '-' + String(start.getUTCMonth()+1).padStart(2,'0');

  let nCalls = 0, nForms = 0;
  try {
    const c = mondayLeadCallsCountsByMonth_Ads_(MONDAY_LEADS_BOARD_ID, MONDAY_LEADS_COL_SOURCE, MONDAY_LEADS_COL_TYPE, MONDAY_LEADS_COL_STATUS, start, end);
    nCalls = c[ymKey] || 0;
  } catch(e){ Logger.log('[Leads Ads Current] Calls KO: ' + e); }
  try {
    const f = mondayLeadFormsCountsByMonth_Ads_(MONDAY_LEADS_BOARD_ID, MONDAY_LEADS_COL_SOURCE, MONDAY_LEADS_COL_TYPE, MONDAY_LEADS_COL_STATUS, start, end);
    nForms = f[ymKey] || 0;
  } catch(e){ Logger.log('[Leads Ads Current] Forms KO: ' + e); }

  const last = sh.getLastRow();
  for (let r=START_ROW_ADS; r<=last; r++){
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === ymKey){
      sh.getRange(r, CALLS_LEAD_COL).setValue(nCalls);
      sh.getRange(r, FORMS_LEAD_COL).setValue(nForms);
      break;
    }
  }
}

/************** FORMULAIRES UNIFIÉS (Paperform + Monday) → "Nombre de formulaire" **************/

// Full history (depuis Jan 2024)
function run_FormsToAds_SyncForSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const FORMS_COL = findColByHeaderAliases_(sh, [
    "nombre de formulaire","formulaires","forms","monday forms","paperform","soumissions","submissions"
  ], HEADERS_ROW_ADS);
  if (!FORMS_COL) throw new Error(`Colonne 'Nombre de formulaire' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  // Fenêtre forcée : Jan 2024 → fin du mois précédent (UTC côté agrégats)
  const from = FULLHIST_START_UTC;
  const now  = new Date();
  const to   = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 0, 23,59,59));

  // Récup Paperform (tolérant)
  let countsPaper = {};
  try {
    const subs = paperformFetchSubmissions_(PAPERFORM_FORM_ID, from, to);
    countsPaper = paperformAggregateMonthlyCounts_(subs);
  } catch (e) {
    Logger.log('[Forms] Paperform KO: ' + e);
  }

  // Récup Monday (tolérant)
  let countsMonday = {};
  try {
    countsMonday = mondayFetchFormCountsByMonth_(MONDAY_BOARD_ID, MONDAY_FORM_COLUMN_TITLE_GOOGLE_ADS, from, to);
  } catch (e) {
    Logger.log('[Forms] Monday KO: ' + e);
  }

  // Fusion = somme par mois
  const allKeys = new Set([...Object.keys(countsPaper), ...Object.keys(countsMonday)]);
  const totalCounts = {};
  allKeys.forEach(k => totalCounts[k] = (countsPaper[k]||0) + (countsMonday[k]||0));

  // Écriture
  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++) {
    const cell = sh.getRange(r, 1).getValue();
    if (isYearSeparatorRow_(cell)) continue;
    const ym = sheetCellToYYYYMM_(cell);
    if (!ym) continue;
    sh.getRange(r, FORMS_COL).setValue(totalCounts[ym] || 0);
  }
}

// Dernier mois uniquement
function run_FormsToAds_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const FORMS_COL = findColByHeaderAliases_(sh, [
    "nombre de formulaire","formulaires","forms","monday forms","paperform","soumissions","submissions"
  ], HEADERS_ROW_ADS);
  if (!FORMS_COL) throw new Error(`Colonne 'Nombre de formulaire' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  const now = new Date();
  const y = now.getUTCFullYear(), m = now.getUTCMonth();
  const start = new Date(Date.UTC(y, m-1, 1, 0, 0, 0));
  const end   = new Date(Date.UTC(y, m,   0, 23, 59, 59));
  const monthKey = start.getUTCFullYear() + '-' + String(start.getUTCMonth()+1).padStart(2,'0');

  // Compte Paperform N-1
  let nPaper = 0;
  try {
    const c = paperformAggregateMonthlyCounts_(paperformFetchSubmissions_(PAPERFORM_FORM_ID, start, end));
    nPaper = c[monthKey] || 0;
  } catch(e){ Logger.log('[Forms N-1] Paperform KO: ' + e); }

  // Compte Monday N-1
  let nMon = 0;
  try {
    const c = mondayFetchFormCountsByMonth_(MONDAY_BOARD_ID, MONDAY_FORM_COLUMN_TITLE_GOOGLE_ADS, start, end);
    nMon = c[monthKey] || 0;
  } catch(e){ Logger.log('[Forms N-1] Monday KO: ' + e); }

  const n = nPaper + nMon;

  // Écrire la ligne N-1
  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === monthKey) { sh.getRange(r, FORMS_COL).setValue(n); break; }
  }
}
// Mois en cours uniquement
function run_FormsToAds_CurrentMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const FORMS_COL = findColByHeaderAliases_(sh, [
    "nombre de formulaire","formulaires","forms","monday forms","paperform","soumissions","submissions"
  ], HEADERS_ROW_ADS);
  if (!FORMS_COL) throw new Error(`Colonne 'Nombre de formulaire' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  const now = new Date();
  const startLocal = new Date(now.getFullYear(), now.getMonth(), 1);
  const endLocal   = now;
  const monthKey   = Utilities.formatDate(startLocal, Session.getScriptTimeZone(), 'yyyy-MM');

  const fromUTC = new Date(Date.UTC(startLocal.getFullYear(), startLocal.getMonth(), 1, 0,0,0));
  const toUTC   = new Date(Date.UTC(endLocal.getFullYear(),   endLocal.getMonth(),   endLocal.getDate(), 23,59,59));

  let nPaper = 0, nMon = 0;
  try {
    const c = paperformAggregateMonthlyCounts_(paperformFetchSubmissions_(PAPERFORM_FORM_ID, fromUTC, toUTC));
    nPaper = c[monthKey] || 0;
  } catch(e){ Logger.log('[Forms Current] Paperform KO: ' + e); }

  try {
    const c = mondayFetchFormCountsByMonth_(MONDAY_BOARD_ID, MONDAY_FORM_COLUMN_TITLE_GOOGLE_ADS, fromUTC, toUTC);
    nMon = c[monthKey] || 0;
  } catch(e){ Logger.log('[Forms Current] Monday KO: ' + e); }

  const n = nPaper + nMon;

  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === monthKey) {
      sh.getRange(r, FORMS_COL).setValue(n);
      break;
    }
  }
}


/**************************** MATOMO → "Durée moyenne visite" ****************************/

const MATOMO_BASE_URL = 'https://matomo.aleo.agency';
const MATOMO_SITE_ID  = 1492;
const MATOMO_PAGE_PATTERN = 'lp-ga-';   // motif "contient"

function getMatomoToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('MATOMO_TOKEN');
  if (!t) throw new Error("Définis la propriété 'MATOMO_TOKEN'.");
  return t;
}

function monthCellToYYYYMM_(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    const y = cell.getFullYear(), m = cell.getMonth() + 1;
    return `${y}-${String(m).padStart(2,'0')}`;
  }
  const s = String(cell || '').trim();
  if (!s) return null;
  let m;
  if (m = s.match(/^(\d{4})-(\d{1,2})$/))   return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  if (m = s.match(/^(\d{4})\/(\d{1,2})$/))  return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  if (m = s.match(/^(\d{4})(\d{2})$/))      return `${m[1]}-${m[2]}`;
  if (m = s.match(/^(\d{4})-(\d{2})-\d{2}$/)) return `${m[1]}-${m[2]}`;
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear(), mm = d.getMonth() + 1;
    return `${y}-${String(mm).padStart(2,'0')}`;
  }
  return null;
}

function matomoFetch_(params) {
  const token = getMatomoToken_();
  const full = { module: 'API', format: 'JSON', token_auth: token, idSite: String(MATOMO_SITE_ID), ...params };
  const qs = Object.keys(full).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(full[k])).join('&');
  const url = `${MATOMO_BASE_URL}/index.php?${qs}`;
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error(`Matomo HTTP ${code}: ${res.getContentText()}`);
  try { return JSON.parse(res.getContentText()); }
  catch(e){ throw new Error('Matomo JSON parse error: ' + e + ' body=' + res.getContentText().slice(0, 400)); }
}

function matomoFetchVisitAvgSecondsWithSegment_(monthYYYYMM) {
  const seg = `pageUrl=@${MATOMO_PAGE_PATTERN}`;
  const js = matomoFetch_({ method: 'VisitsSummary.get', period: 'month', date: monthYYYYMM, hideMetricTranslations: 1, segment: seg });

  if (js && (js.avg_time_on_site != null || (js.sum_visit_length != null && js.nb_visits != null))) {
    const direct = Number(js.avg_time_on_site);
    if (!isNaN(direct)) return { secs: direct, via: 'VisitsSummary+segment' };
    const sum = Number(js.sum_visit_length), nb = Number(js.nb_visits);
    if (!isNaN(sum) && nb > 0) return { secs: sum/nb, via: 'VisitsSummary+segment' };
  }
    if (js && typeof js === 'object' && !Array.isArray(js)) {

    const node = js[`${monthYYYYMM}-01`] || js[monthYYYYMM];
    if (node) {
      if (node.avg_time_on_site != null) {
        const v = Number(node.avg_time_on_site); if (!isNaN(v)) return { secs: v, via: 'VisitsSummary+segment' };
      }
      if (node.sum_visit_length != null && node.nb_visits != null) {
        const sum = Number(node.sum_visit_length), nb = Number(node.nb_visits);
        if (!isNaN(sum) && nb > 0) return { secs: sum/nb, via: 'VisitsSummary+segment' };
      }
    }
  }
  if (Array.isArray(js) && js.length) {
    for (const row of js) {
      if (!row) continue;
      if (row.avg_time_on_site != null) {
        const v = Number(row.avg_time_on_site); if (!isNaN(v)) return { secs: v, via: 'VisitsSummary+segment' };
      }
      if (row.sum_visit_length != null && row.nb_visits != null) {
        const sum = Number(row.sum_visit_length), nb = Number(row.nb_visits);
        if (!isNaN(sum) && nb > 0) return { secs: sum/nb, via: 'VisitsSummary+segment' };
      }
    }
  }
  return { secs: 0, via: 'VisitsSummary+segment' };
}

function matomoFetchPageAvgSecondsByPattern_(monthYYYYMM) {
  const js = matomoFetch_({
    method: 'Actions.getPageUrls',
    period: 'month',
    date: monthYYYYMM,
    flat: 1,
    filter_limit: -1,
    filter_column: 'label',
    filter_pattern: MATOMO_PAGE_PATTERN,
    hideMetricTranslations: 1
  });
  let sum = 0, hits = 0;
  const arr = Array.isArray(js) ? js : (js && js.data ? js.data : []);
  (arr || []).forEach(row => {
    const h = Number(row && row.nb_hits);
    const s = Number(row && row.sum_time_spent);
    if (!isNaN(h) && h > 0 && !isNaN(s)) { sum += s; hits += h; }
  });
  if (hits > 0) return { secs: sum / hits, via: 'Actions.getPageUrls' };
  return { secs: 0, via: 'Actions.getPageUrls' };
}

function matomoResolveAvgSeconds_(monthYYYYMM) {
  const a = matomoFetchVisitAvgSecondsWithSegment_(monthYYYYMM);
  if (a.secs > 0) return a;
  const b = matomoFetchPageAvgSecondsByPattern_(monthYYYYMM);
  return b;
}

// Sync Matomo (une seule colonne "Durée moyenne visite") — ignore les lignes année
function run_MatomoToAds_SyncForSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);
  const last = sh.getLastRow();
  if (last < START_ROW_ADS) return;

  // Trouve la colonne par en-tête
  const MATOMO_COL = findColByHeaderAliases_(sh, ['duree moyenne visite','durée moyenne visite'], HEADERS_ROW_ADS);
  if (!MATOMO_COL) throw new Error(`Colonne 'Durée moyenne visite' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  for (let r = START_ROW_ADS; r <= last; r++) {
    const cell = sh.getRange(r, 1).getValue();
    if (isYearSeparatorRow_(cell)) continue;
    const ym = sheetCellToYYYYMM_(cell);
    if (!ym) continue;

    const { secs, via } = matomoResolveAvgSeconds_(ym);
    setSecondsAsDuration_(sh, r, MATOMO_COL, secs);
    Logger.log(`[Matomo] ${ym}: ${secs.toFixed(2)} sec via ${via}`);
    Utilities.sleep(120);
  }
}

// MàJ du mois précédent uniquement
function run_MatomoToAds_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const MATOMO_COL = findColByHeaderAliases_(sh, ['duree moyenne visite','durée moyenne visite'], HEADERS_ROW_ADS);
  if (!MATOMO_COL) throw new Error(`Colonne 'Durée moyenne visite' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  const today = new Date();
  const monthKey = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth()-1, 1), Session.getScriptTimeZone(), 'yyyy-MM');

  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === monthKey) {
      const { secs, via } = matomoResolveAvgSeconds_(monthKey);
      setSecondsAsDuration_(sh, r, MATOMO_COL, secs);
      Logger.log(`[Matomo] ${monthKey}: ${secs.toFixed(2)} sec via ${via}`);
      break;
    }
  }
}
function run_MatomoToAds_CurrentMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const MATOMO_COL = findColByHeaderAliases_(sh, ['duree moyenne visite','durée moyenne visite'], HEADERS_ROW_ADS);
  if (!MATOMO_COL) throw new Error(`Colonne 'Durée moyenne visite' introuvable (ligne ${HEADERS_ROW_ADS}).`);

  const today = new Date();
  const monthKey = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), 1), Session.getScriptTimeZone(), 'yyyy-MM');

  const last = sh.getLastRow();
  for (let r = START_ROW_ADS; r <= last; r++) {
    const val = sh.getRange(r,1).getValue();
    if (isYearSeparatorRow_(val)) continue;
    const ym = sheetCellToYYYYMM_(val);
    if (ym === monthKey) {
      const { secs, via } = matomoResolveAvgSeconds_(monthKey);
      setSecondsAsDuration_(sh, r, MATOMO_COL, secs);
      Logger.log(`[Matomo Current] ${monthKey}: ${secs.toFixed(2)} sec via ${via}`);
      break;
    }
  }
}

// JOB Ads — Mois en cours
function run_GAds_CurrentMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_ADS);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME_ADS} introuvable`);

  const moisCol   = findColByHeaderAliases_(sh, ['mois'], HEADERS_ROW_ADS);
  const budgetCol = findColByHeaderAliases_(sh, ['budget investi','depenses','depense','cout','coût','spend','cost'], HEADERS_ROW_ADS);
  const impCol    = findColByHeaderAliases_(sh, ['impressions','nb impressions','nombre dimpressions','nombre d impressions'], HEADERS_ROW_ADS);
  const clkCol    = findColByHeaderAliases_(sh, ['clics','clicks','nb clics','nombre de clics'], HEADERS_ROW_ADS);
  const ctrCol    = findColByHeaderAliases_(sh, ['ctr','taux de clics','click through rate'], HEADERS_ROW_ADS);
  const cpcCol    = findColByHeaderAliases_(sh, ['cpc','cout par clic','coût par clic','cost per click'], HEADERS_ROW_ADS);
  const callsCol  = findColByHeaderAliases_(sh, ['appels google ads','appels','calls','nombre dappels'], HEADERS_ROW_ADS);

  const missing = [];
  if (!moisCol)   missing.push('Mois');
  if (!budgetCol) missing.push('Budget investi');
  if (!impCol)    missing.push('Impressions');
  if (!clkCol)    missing.push('Clics');
  if (!ctrCol)    missing.push('CTR');
  if (!cpcCol)    missing.push('CPC');
  if (missing.length) throw new Error('[GAds CurrentMonth] Colonnes manquantes: ' + missing.join(', '));

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const end   = today; // jusqu’à aujourd’hui
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');

  const data = fetchGoogleAdsMonthly(
    Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    Utilities.formatDate(end,   Session.getScriptTimeZone(), 'yyyy-MM-dd')
  );
  const byMonth = aggregateAdsByMonth_(data);
  const v = byMonth[ymKey];
  if (!v) {
    Logger.log('[GAds CurrentMonth] Aucune donnée renvoyée pour ' + ymKey);
    return;
  }

  const row = _adsFindOrInsertMonthRow_(sh, moisCol, ymKey);

  const moisCell = sh.getRange(row, moisCol);
  if (!moisCell.getFormula()) {
    moisCell.setValue(monthKeyToFr_(ymKey));
  } else {
    Logger.log(`[SKIP] Préserve formule en ${moisCell.getA1Notation()}`);
  }

  setPreserveFormula_(sh, row, budgetCol, v.cost, '0.00');
  setPreserveFormula_(sh, row, impCol,    v.impressions, null);
  setPreserveFormula_(sh, row, clkCol,    v.clicks, null);
  if (ctrCol) setPreserveFormula_(sh, row, ctrCol, v.ctr, '0.00%');
  if (cpcCol) setPreserveFormula_(sh, row, cpcCol, v.cpc, '0.00');
  if (callsCol) setPreserveFormula_(sh, row, callsCol, v.calls || 0, null);

  Logger.log(`[GAds CurrentMonth] ${ymKey} -> cost=${v.cost}, imp=${v.impressions}, clk=${v.clicks}, ctr=${v.ctr}, cpc=${v.cpc}`);
}


