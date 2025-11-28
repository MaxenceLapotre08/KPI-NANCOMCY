/**************************** GBM (Business Profile) ****************************/

const SHEET_NAME = 'GMB';                 // nom de l’onglet GBM
const HEADERS_ROW = 3;                     // ligne des entêtes
const START_ROW = 6;                     // première ligne de données
const LOCATION_NAME = 'locations/17344379108514631991'; // ID de fiche GBM

// Metrics dispo via l’API Business Profile Performance
const METRICS = {
  'BUSINESS_IMPRESSIONS_DESKTOP_SEARCH': 'Vues recherche Google Ordinateur',
  'BUSINESS_IMPRESSIONS_MOBILE_SEARCH': 'Vues recherche Google Mobile',
  'BUSINESS_IMPRESSIONS_DESKTOP_MAPS': 'Vues Google Maps Ordinateur',
  'BUSINESS_IMPRESSIONS_MOBILE_MAPS': 'Vues Google Maps Mobile',
  'WEBSITE_CLICKS': 'Clics site web',
  'CALL_CLICKS': 'Appels',
  'BUSINESS_DIRECTION_REQUESTS': 'Demande d\'itinéraire'
};

// ************************** HELPERS GÉNÉRAUX (Utilise Utils.js) ****************************/
// Les fonctions suivantes sont maintenant dans Utils.js :
// Utils_normHeader -> Utils_normHeader
// Utils_findColByHeaderAliases -> Utils_findColByHeaderAliases
// Utils_isYearSeparatorRow -> Utils_isYearSeparatorRow
// Utils_sheetCellToYYYYMM -> Utils_sheetCellToYYYYMM
// Utils_executeWithRetry -> Utils_executeWithRetry

// ✅ Fonction maintenant dans SheetHelpers.js
// _gbmStyleYearRow_ -> SheetHelpers.styleYearRow


// VRAI si la colonne est listée comme ayant une formule dans Sheet_Structures.md
function GMB_isProtectedHeader_(sh, col) {
  if (!col) return false;
  const headerVal = sh.getRange(HEADERS_ROW, col).getValue();
  const h = Utils_normHeader(headerVal);

  // Liste EXACTE des colonnes avec formules pour "GMB"
  const protectedHeaders = [
    'nombre de signature', 'total montant signe', 'montant moyen signature'
  ].map(Utils_normHeader);

  // La protection est active si le nom de colonne est dans la liste.
  return protectedHeaders.some(p => h.includes(p));
}

/* ---------- OAuth ---------- */
function getService() {
  return OAuth2.createService('GBP')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/business.manage')
    .setCallbackFunction('authCallback');
}
function authorize() {
  const service = getService();
  if (!service.hasAccess()) Logger.log('Open this URL: ' + service.getAuthorizationUrl());
  else Logger.log('Already authorized');
}
function authCallback(request) {
  const service = getService();
  const ok = service.handleCallback(request);
  return HtmlService.createHtmlOutput(ok ? 'Success!' : 'Denied.');
}

// Utils_executeWithRetry est maintenant dans Utils.js

/* ---------- API BPP (performance) ---------- */
function fetchDailyMetrics_(locationName, startDate, endDate) {
  const service = getService();
  if (!service.hasAccess()) throw new Error('No OAuth access. Run authorize().');

  const base = 'https://businessprofileperformance.googleapis.com/v1';
  const metrics = Object.keys(METRICS);
  const qs = metrics.map(m => `dailyMetrics=${encodeURIComponent(m)}`).join('&')
    + `&dailyRange.startDate.year=${startDate.getFullYear()}`
    + `&dailyRange.startDate.month=${startDate.getMonth() + 1}`
    + `&dailyRange.startDate.day=${startDate.getDate()}`
    + `&dailyRange.endDate.year=${endDate.getFullYear()}`
    + `&dailyRange.endDate.month=${endDate.getMonth() + 1}`
    + `&dailyRange.endDate.day=${endDate.getDate()}`;

  const url = `${base}/${locationName}:fetchMultiDailyMetricsTimeSeries?${qs}`;
  const headers = { Authorization: `Bearer ${service.getAccessToken()}` };

  const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
  const code = res.getResponseCode();
  const body = res.getContentText();
  Logger.log(`[GBM] GET ${url}\nHTTP ${code}`);
  if (code >= 300) {
    Logger.log('Response body:\n' + body);
    throw new Error(`GBM HTTP ${code}: ${body}`);
  }
  return JSON.parse(body);
}

function aggregateByMonth_(responseJson) {
  const out = {};
  const groups = responseJson && responseJson.multiDailyMetricTimeSeries || [];
  groups.forEach(g => {
    const list = g && g.dailyMetricTimeSeries || [];
    list.forEach(ts => {
      const metric = ts.dailyMetric;
      const pts = (ts.timeSeries && (ts.timeSeries.points || ts.timeSeries.datedValues)) || [];
      pts.forEach(p => {
        // deux schémas possibles: {date:{y,m,d}, value} ou {timeDimension:{timePoint:{date:{...}}}, value}
        const d = p.date || (p.timeDimension && p.timeDimension.timePoint && p.timeDimension.timePoint.date);
        const y = d && d.year, m = d && d.month, day = d && d.day;
        if (!y || !m || !day) return;

        const key = `${y}-${String(m).padStart(2, '0')}`;
        const val = Number(p.value || 0);
        if (!out[key]) out[key] = {};
        out[key][metric] = (out[key][metric] || 0) + val;
      });
    });
  });
  return out;
}


/* ---------- Reviews (My Business v4) ---------- */

// On mémorise l'accountId pour éviter de le redemander à chaque fois
function _getCachedAccountId_() {
  const props = PropertiesService.getUserProperties();
  let acc = props.getProperty('GBP_ACCOUNT_ID');
  if (acc) return acc;

  const service = getService();
  if (!service.hasAccess()) throw new Error('No OAuth access. Run authorize().');

  // Account Management API v1 (recommandé)
  const url = 'https://mybusinessaccountmanagement.googleapis.com/v1/accounts';
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${service.getAccessToken()}` },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error(`AccountManagement HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
  }
  const json = JSON.parse(res.getContentText());
  const accounts = json.accounts || [];
  if (!accounts.length) throw new Error('Aucun compte GBP accessible avec ce token.');
  // on prend le premier
  const name = accounts[0].name; // ex "accounts/1234567890"
  acc = String(name).split('/')[1];
  props.setProperty('GBP_ACCOUNT_ID', acc);
  return acc;
}

function _parseStarRating_(rev) {
  // essaie nombre, chaîne numérique, ou enum 'ONE'..'FIVE'
  const candidates = [
    rev.starRating, rev.rating,
    rev.reviewRating && rev.reviewRating.starRating
  ];
  for (let v of candidates) {
    if (v == null) continue;
    if (typeof v === 'number') return v;
    if (typeof v === 'string') {
      const t = v.trim();
      if (/^\d+(\.\d+)?$/.test(t)) return parseFloat(t);
      const map = { ONE: 1, TWO: 2, THREE: 3, FOUR: 4, FIVE: 5 };
      const up = t.toUpperCase();
      if (map[up]) return map[up];
    }
  }
  return NaN;
}

/** Récupère les avis (paginés) et renvoie une map { 'YYYY-MM': {count, avg} } dans la fenêtre [start,end] */
function fetchReviewsMonthly_(startDate, endDate) {
  const accountId = _getCachedAccountId_();
  const locationId = String(LOCATION_NAME).split('/')[1];

  const service = getService();
  if (!service.hasAccess()) throw new Error('No OAuth access. Run authorize().');

  const base = `https://mybusiness.googleapis.com/v4/accounts/${encodeURIComponent(accountId)}/locations/${encodeURIComponent(locationId)}/reviews`;
  let pageToken = '', done = false;

  const fromTs = startDate.getTime();
  const toTs = endDate.getTime();

  const bucket = {}; // {ym:{sum, count}}

  while (!done) {
    const url = base + `?pageSize=50&orderBy=updateTime%20desc` + (pageToken ? `&pageToken=${encodeURIComponent(pageToken)}` : '');
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${service.getAccessToken()}` },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 404) {
      // API non activée ou pas d'accès à reviews
      Logger.log(`Reviews 404 — vérifie que l'API "Google My Business API" est bien activée pour ce projet GCP et que le compte a l’accès GBP.`);
      break;
    }
    if (res.getResponseCode() >= 300) {
      throw new Error(`Reviews HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    }

    const json = JSON.parse(res.getContentText());
    const reviews = json.reviews || [];
    for (const r of reviews) {
      const t = r.createTime || r.updateTime;
      if (!t) continue;
      const d = new Date(t);
      const ts = d.getTime();
      // comme on trie desc, si on dépasse la borne inférieure on peut s'arrêter plus tard, mais on continue simple
      if (ts < fromTs || ts > toTs) continue;

      const ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      const star = _parseStarRating_(r);
      if (!bucket[ym]) bucket[ym] = { sum: 0, count: 0 };
      if (!isNaN(star)) bucket[ym].sum += star;
      bucket[ym].count += 1;
    }

    pageToken = json.nextPageToken || '';
    if (!pageToken) done = true;
  }

  // transforme en {ym: {count, avg}}
  const out = {};
  Object.keys(bucket).forEach(ym => {
    const b = bucket[ym];
    out[ym] = { count: b.count, avg: b.count ? (b.sum / b.count) : 0 };
  });
  return out;
}


// ✅ Fonction maintenant dans SheetHelpers.js (56 lignes supprimées)
// _gbmFindOrInsertMonthRow_ -> SheetHelpers.ensureMonthRow

/* ---------- Écriture d’une ligne mois (détection par entêtes) ---------- */
function writeGBMRowAt_(sh, rowIndex, monthKey, metrics, reviewsAgg) {
  const moisCol = Utils_findColByHeaderAliases(sh, ['mois']);
  const vuesCol = Utils_findColByHeaderAliases(sh, ['vues']);
  const clicsCol = Utils_findColByHeaderAliases(sh, ['clics site web', 'clics']);
  const itinCol = Utils_findColByHeaderAliases(sh, ['demande ditineraire', 'demande d itineraire', 'itineraire']);
  const appelsCol = Utils_findColByHeaderAliases(sh, ['appels']);
  const tauxIntCol = Utils_findColByHeaderAliases(sh, ['taux dinteraction', 'taux interaction']);
  const tauxAppelCol = Utils_findColByHeaderAliases(sh, ['taux dappel', 'taux appel']);
  const avisNbCol = Utils_findColByHeaderAliases(sh, ['nombre davis', 'nombre d avis', 'nb avis']);
  const avisScoreCol = Utils_findColByHeaderAliases(sh, ['score avis']);
  const mapsMobCol = Utils_findColByHeaderAliases(sh, ['vues google maps mobile', 'maps mobile']);
  const mapsDeskCol = Utils_findColByHeaderAliases(sh, ['vues google maps ordinateur', 'maps ordinateur', 'maps desktop']);
  const srchMobCol = Utils_findColByHeaderAliases(sh, ['vues recherche google mobile', 'recherche mobile', 'search mobile']);
  const srchDeskCol = Utils_findColByHeaderAliases(sh, ['vues recherche google ordinateur', 'recherche ordinateur', 'search desktop']);

  const vues = (metrics['BUSINESS_IMPRESSIONS_DESKTOP_SEARCH'] || 0)
    + (metrics['BUSINESS_IMPRESSIONS_MOBILE_SEARCH'] || 0)
    + (metrics['BUSINESS_IMPRESSIONS_DESKTOP_MAPS'] || 0)
    + (metrics['BUSINESS_IMPRESSIONS_MOBILE_MAPS'] || 0);
  const clics = metrics['WEBSITE_CLICKS'] || 0;
  const appels = metrics['CALL_CLICKS'] || 0;
  const itin = metrics['BUSINESS_DIRECTION_REQUESTS'] || 0;

  const tauxInteraction = vues ? ((clics + appels + itin) / vues) : 0;
  const tauxAppel = vues ? (appels / vues) : 0;

  if (!GMB_isProtectedHeader_(sh, moisCol)) Utils_setPreserveFormula(sh, rowIndex, moisCol, monthKey);
  if (!GMB_isProtectedHeader_(sh, vuesCol)) Utils_setPreserveFormula(sh, rowIndex, vuesCol, vues);
  if (!GMB_isProtectedHeader_(sh, clicsCol)) Utils_setPreserveFormula(sh, rowIndex, clicsCol, clics);
  if (!GMB_isProtectedHeader_(sh, itinCol)) Utils_setPreserveFormula(sh, rowIndex, itinCol, itin);
  if (!GMB_isProtectedHeader_(sh, appelsCol)) Utils_setPreserveFormula(sh, rowIndex, appelsCol, appels);

  // Avis (par mois) si reviewsAgg fourni
  if (reviewsAgg && (avisNbCol || avisScoreCol)) {
    const r = reviewsAgg[monthKey];
    if (r) {
      if (!GMB_isProtectedHeader_(sh, avisNbCol)) Utils_setPreserveFormula(sh, rowIndex, avisNbCol, r.count);
      if (!GMB_isProtectedHeader_(sh, avisScoreCol)) Utils_setPreserveFormula(sh, rowIndex, avisScoreCol, r.avg, '0.00');
    }
  }

  if (!GMB_isProtectedHeader_(sh, mapsMobCol)) Utils_setPreserveFormula(sh, rowIndex, mapsMobCol, metrics['BUSINESS_IMPRESSIONS_MOBILE_MAPS'] || 0);
  if (!GMB_isProtectedHeader_(sh, mapsDeskCol)) Utils_setPreserveFormula(sh, rowIndex, mapsDeskCol, metrics['BUSINESS_IMPRESSIONS_DESKTOP_MAPS'] || 0);
  if (!GMB_isProtectedHeader_(sh, srchMobCol)) Utils_setPreserveFormula(sh, rowIndex, srchMobCol, metrics['BUSINESS_IMPRESSIONS_MOBILE_SEARCH'] || 0);
  if (!GMB_isProtectedHeader_(sh, srchDeskCol)) Utils_setPreserveFormula(sh, rowIndex, srchDeskCol, metrics['BUSINESS_IMPRESSIONS_DESKTOP_SEARCH'] || 0);
  if (!GMB_isProtectedHeader_(sh, tauxIntCol)) Utils_setPreserveFormula(sh, rowIndex, tauxIntCol, tauxInteraction, '0.00%');
  if (!GMB_isProtectedHeader_(sh, tauxAppelCol)) Utils_setPreserveFormula(sh, rowIndex, tauxAppelCol, tauxAppel, '0.00%');
}

/* ---------- JOBS ---------- */

function run_GBM_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME} introuvable`);

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);

  const raw = Utils_executeWithRetry(() => fetchDailyMetrics_(LOCATION_NAME, start, end), 'GBM fetch N-1');
  const monthly = aggregateByMonth_(raw);

  // Avis du même mois
  const reviewsAgg = Utils_executeWithRetry(() => fetchReviewsMonthly_(start, end), 'GBM reviews N-1');

  Object.keys(monthly).forEach(ym => {
    const moisCol = Utils_findColByHeaderAliases(sh, ['mois']);
    const row = _gbmFindOrInsertMonthRow_(sh, moisCol, ym);
    writeGBMRowAt_(sh, row, ym, monthly[ym], reviewsAgg);
  });
}

function run_GBM_FullHistory() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME} introuvable`);

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 24, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);

  const raw = Utils_executeWithRetry(() => fetchDailyMetrics_(LOCATION_NAME, start, end), 'GBM fetch history');
  const monthly = aggregateByMonth_(raw);

  // Avis sur la même fenêtre
  const reviewsAgg = Utils_executeWithRetry(() => fetchReviewsMonthly_(start, end), 'GBM reviews history');

  // Écriture cellule par cellule
  const allKeys = new Set([...Object.keys(monthly), ...Object.keys(reviewsAgg)]);
  Array.from(allKeys).sort().forEach(ym => {
    const moisCol = Utils_findColByHeaderAliases(sh, ['mois']);
    if (!moisCol) throw new Error("Colonne 'Mois' introuvable dans GMB.");
    const row = _gbmFindOrInsertMonthRow_(sh, moisCol, ym);
    writeGBMRowAt_(sh, row, ym, monthly[ym] || {}, reviewsAgg);
  });
}
