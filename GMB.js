/**************************** GBM (Business Profile) ****************************/

const SHEET_NAME   = 'GMB';                 // nom de l’onglet GBM
const HEADERS_ROW  = 3;                     // ligne des entêtes
const START_ROW    = 6;                     // première ligne de données
const LOCATION_NAME = 'locations/17344379108514631991'; // ID de fiche GBM

// Metrics dispo via l’API Business Profile Performance
const METRICS = {
  'BUSINESS_IMPRESSIONS_DESKTOP_SEARCH': 'Vues recherche Google Ordinateur',
  'BUSINESS_IMPRESSIONS_MOBILE_SEARCH':  'Vues recherche Google Mobile',
  'BUSINESS_IMPRESSIONS_DESKTOP_MAPS':   'Vues Google Maps Ordinateur',
  'BUSINESS_IMPRESSIONS_MOBILE_MAPS':    'Vues Google Maps Mobile',
  'WEBSITE_CLICKS':                      'Clics site web',
  'CALL_CLICKS':                         'Appels',
  'BUSINESS_DIRECTION_REQUESTS':         'Demande d\'itinéraire'
};

const CLIENT_ID = scriptProperties.getProperty('CLIENT_ID');
const CLIENT_SECRET = scriptProperties.getProperty('CLIENT_SECRET');

/* ---------- Helpers entêtes/format ---------- */
function _normHeader_(s){
  return String(s||'').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[’'`]/g,'')
    .replace(/\s+/g,' ')
    .trim();
}
function findColByHeaderAliases_(sheet, aliases, headerRow){
  const row = headerRow || HEADERS_ROW;
  const headers = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]||[];
  const HN = headers.map(_normHeader_);
  const wanted = aliases.map(_normHeader_);
  for (let i=0;i<HN.length;i++){
    const h = HN[i]; if (!h) continue;
    if (wanted.some(w=>h===w || h.includes(w))) return i+1;
  }
  return 0;
}
function isYearSeparatorRow_(cell) {
  if (cell == null) return false;
  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) return false;
  const s = String(cell).trim();
  if (/^\d{4}$/.test(s)) {
    const y = +s; return y>=1900 && y<=2100;
  }
  if (typeof cell === 'number' && isFinite(cell)) {
    const y = Math.round(cell); return y>=1900 && y<=2100;
  }
  return false;
}
function sheetCellToYYYYMM_(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    const y = cell.getFullYear(), m = cell.getMonth()+1;
    return `${y}-${String(m).padStart(2,'0')}`;
  }
  const raw = String(cell || '').trim();
  if (!raw) return null;
  let m;
  if ((m = raw.match(/^(\d{4})-(\d{1,2})$/)))    return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  if ((m = raw.match(/^(\d{4})\/(\d{1,2})$/)))   return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  if ((m = raw.match(/^(\d{4})(\d{2})$/)))       return `${m[1]}-${m[2]}`;
  if ((m = raw.match(/^(\d{4})-(\d{2})-\d{2}$/)))return `${m[1]}-${m[2]}`;
  const d = new Date(raw);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear(), mo = d.getMonth()+1;
    return `${y}-${String(mo).padStart(2,'0')}`;
  }
  return null;
}
function _gbmStyleYearRow_(sh, row, moisCol) {
  sh.getRange(row, 1, 1, sh.getLastColumn()).setBackground('#e6e1f5');
  sh.getRange(row, moisCol).setFontWeight('bold');
}

function setPreserveFormula_(sh, row, col, value, numberFormat) {
  if (!col) return;
  const cell = sh.getRange(row, col);
  const formula = cell.getFormula();
  if (formula) {
    Logger.log(`[GMB SKIP] Préserve formule en ${cell.getA1Notation()} -> ${formula}`);
    return;
  }
  cell.setValue(value);
  if (numberFormat) cell.setNumberFormat(numberFormat);
}

/* ---------- OAuth ---------- */
function getService() {
  return OAuth2.createService('GBP')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(GBP_OAUTH_CLIENT_ID)
    .setClientSecret(GBP_OAUTH_CLIENT_SECRET)
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

/* ---------- Retry/backoff générique ---------- */
function executeWithRetry_(fn, label, maxRetries) {
  label = label || 'task'; maxRetries = maxRetries || 3;
  let attempt = 0, lastErr;
  while (attempt < maxRetries) {
    try {
      attempt++;
      const out = fn();
      return out;
    } catch (err) {
      lastErr = err;
      const msg = String(err && err.message || err);
      const retryable = /(^|\s)(429|5\d\d|RATE_LIMIT|RESOURCE_EXHAUSTED)/i.test(msg);
      if (attempt >= maxRetries || !retryable) break;
      Utilities.sleep(Math.pow(2, attempt) * 1000); // 2s, 4s, 8s
    }
  }
  throw lastErr;
}

/* ---------- API BPP (performance) ---------- */
function fetchDailyMetrics_(locationName, startDate, endDate) {
  const service = getService();
  if (!service.hasAccess()) throw new Error('No OAuth access. Run authorize().');

  const base = 'https://businessprofileperformance.googleapis.com/v1';
  const metrics = Object.keys(METRICS);
  const qs = metrics.map(m => `dailyMetrics=${encodeURIComponent(m)}`).join('&')
    + `&dailyRange.startDate.year=${startDate.getFullYear()}`
    + `&dailyRange.startDate.month=${startDate.getMonth()+1}`
    + `&dailyRange.startDate.day=${startDate.getDate()}`
    + `&dailyRange.endDate.year=${endDate.getFullYear()}`
    + `&dailyRange.endDate.month=${endDate.getMonth()+1}`
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

        const key = `${y}-${String(m).padStart(2,'0')}`;
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
      const map = { ONE:1, TWO:2, THREE:3, FOUR:4, FIVE:5 };
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
  const toTs   = endDate.getTime();

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

      const ym = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
      const star = _parseStarRating_(r);
      if (!bucket[ym]) bucket[ym] = { sum:0, count:0 };
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

/* ---------- Placement intelligent (lignes “année” + mois) ---------- */
function _gbmFindOrInsertMonthRow_(sh, moisCol, targetYM) {
  const targetYear = parseInt(targetYM.slice(0,4), 10);
  const lastRow = Math.max(sh.getLastRow(), START_ROW-1);

  if (lastRow < START_ROW) {
    sh.insertRowsBefore(START_ROW, 2);
    sh.getRange(START_ROW, moisCol).setValue(String(targetYear));
    _gbmStyleYearRow_(sh, START_ROW, moisCol);
    return START_ROW + 1;
  }

  let yearRowForTarget = null;
  let nextYearRowAfterTarget = null;
  let existingMonthRow = null;

  for (let r = START_ROW; r <= sh.getLastRow(); r++) {
    const v = sh.getRange(r, moisCol).getValue();
    if (isYearSeparatorRow_(v)) {
      const y = parseInt(String(v).trim(), 10);
      if (yearRowForTarget === null && y === targetYear) yearRowForTarget = r;
      if (nextYearRowAfterTarget === null && y > targetYear && yearRowForTarget !== null) {
        nextYearRowAfterTarget = r; break;
      }
      continue;
    }
    const ym = sheetCellToYYYYMM_(v);
    if (ym === targetYM) { existingMonthRow = r; break; }
  }
  if (existingMonthRow) return existingMonthRow;

  if (yearRowForTarget !== null) {
    const blockStart = yearRowForTarget + 1;
    const blockEnd = (nextYearRowAfterTarget ? nextYearRowAfterTarget : (sh.getLastRow()+1)) - 1;
    for (let r = blockStart; r <= blockEnd; r++) {
      const ym = sheetCellToYYYYMM_(sh.getRange(r, moisCol).getValue());
      if (ym && ym > targetYM) { sh.insertRowsBefore(r, 1); return r; }
    }
    const insertAt = blockEnd + 1;
    sh.insertRowsBefore(insertAt, 1);
    return insertAt;
  }

  let firstGreaterYearRow = null;
  for (let r = START_ROW; r <= sh.getLastRow(); r++) {
    const v = sh.getRange(r, moisCol).getValue();
    if (isYearSeparatorRow_(v)) {
      const y = parseInt(String(v).trim(), 10);
      if (y > targetYear) { firstGreaterYearRow = r; break; }
    }
  }
  const yearRow = firstGreaterYearRow ? firstGreaterYearRow : (sh.getLastRow()+1);
  sh.insertRowsBefore(yearRow, 2);
  sh.getRange(yearRow, moisCol).setValue(String(targetYear));
  _gbmStyleYearRow_(sh, yearRow, moisCol);
  return yearRow + 1;
}

/* ---------- Écriture d’une ligne mois (détection par entêtes) ---------- */
function writeGBMRowAt_(sh, rowIndex, monthKey, metrics, reviewsAgg) {
  const moisCol = findColByHeaderAliases_(sh, ['mois']);
  const vuesCol = findColByHeaderAliases_(sh, ['vues']);
  const clicsCol = findColByHeaderAliases_(sh, ['clics site web','clics']);
  const itinCol = findColByHeaderAliases_(sh, ['demande ditineraire','demande d itineraire','itineraire']);
  const appelsCol = findColByHeaderAliases_(sh, ['appels']);
  const tauxIntCol = findColByHeaderAliases_(sh, ['taux dinteraction','taux interaction']);
  const tauxAppelCol = findColByHeaderAliases_(sh, ['taux dappel','taux appel']);
  const avisNbCol = findColByHeaderAliases_(sh, ['nombre davis','nombre d avis','nb avis']);
  const avisScoreCol = findColByHeaderAliases_(sh, ['score avis']);
  const mapsMobCol = findColByHeaderAliases_(sh, ['vues google maps mobile','maps mobile']);
  const mapsDeskCol= findColByHeaderAliases_(sh, ['vues google maps ordinateur','maps ordinateur','maps desktop']);
  const srchMobCol = findColByHeaderAliases_(sh, ['vues recherche google mobile','recherche mobile','search mobile']);
  const srchDeskCol= findColByHeaderAliases_(sh, ['vues recherche google ordinateur','recherche ordinateur','search desktop']);

  const vues = (metrics['BUSINESS_IMPRESSIONS_DESKTOP_SEARCH']||0)
             + (metrics['BUSINESS_IMPRESSIONS_MOBILE_SEARCH']||0)
             + (metrics['BUSINESS_IMPRESSIONS_DESKTOP_MAPS']||0)
             + (metrics['BUSINESS_IMPRESSIONS_MOBILE_MAPS']||0);
  const clics = metrics['WEBSITE_CLICKS']||0;
  const appels= metrics['CALL_CLICKS']||0;
  const itin  = metrics['BUSINESS_DIRECTION_REQUESTS']||0;

  const tauxInteraction = vues ? ( (clics + appels + itin) / vues ) : 0;
  const tauxAppel       = vues ? ( appels / vues ) : 0;

  setPreserveFormula_(sh, rowIndex, moisCol, monthKey);
  setPreserveFormula_(sh, rowIndex, vuesCol, vues);
  setPreserveFormula_(sh, rowIndex, clicsCol, clics);
  setPreserveFormula_(sh, rowIndex, itinCol, itin);
  setPreserveFormula_(sh, rowIndex, appelsCol, appels);

  // Avis (par mois) si reviewsAgg fourni
  if (reviewsAgg && (avisNbCol || avisScoreCol)) {
    const r = reviewsAgg[monthKey];
    if (r) {
      setPreserveFormula_(sh, rowIndex, avisNbCol, r.count);
      setPreserveFormula_(sh, rowIndex, avisScoreCol, r.avg, '0.00');
    }
  }

  setPreserveFormula_(sh, rowIndex, mapsMobCol, metrics['BUSINESS_IMPRESSIONS_MOBILE_MAPS']||0);
  setPreserveFormula_(sh, rowIndex, mapsDeskCol, metrics['BUSINESS_IMPRESSIONS_DESKTOP_MAPS']||0);
  setPreserveFormula_(sh, rowIndex, srchMobCol, metrics['BUSINESS_IMPRESSIONS_MOBILE_SEARCH']||0);
  setPreserveFormula_(sh, rowIndex, srchDeskCol, metrics['BUSINESS_IMPRESSIONS_DESKTOP_SEARCH']||0);
  setPreserveFormula_(sh, rowIndex, tauxIntCol, tauxInteraction, '0.00%');
  setPreserveFormula_(sh, rowIndex, tauxAppelCol, tauxAppel, '0.00%');
}

/* ---------- JOBS ---------- */

function run_GBM_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME} introuvable`);

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth()-1, 1);
  const end   = new Date(today.getFullYear(), today.getMonth(), 0);

  const raw = executeWithRetry_(() => fetchDailyMetrics_(LOCATION_NAME, start, end), 'GBM fetch N-1');
  const monthly = aggregateByMonth_(raw);

  // Avis du même mois
  const reviewsAgg = executeWithRetry_(() => fetchReviewsMonthly_(start, end), 'GBM reviews N-1');

  Object.keys(monthly).forEach(ym => {
    const moisCol = findColByHeaderAliases_(sh, ['mois']);
    const row = _gbmFindOrInsertMonthRow_(sh, moisCol, ym);
    writeGBMRowAt_(sh, row, ym, monthly[ym], reviewsAgg);
  });
}

function run_GBM_FullHistory() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Onglet ${SHEET_NAME} introuvable`);

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 24, 1);
  const end   = new Date(today.getFullYear(), today.getMonth(), 0);

  const raw = executeWithRetry_(() => fetchDailyMetrics_(LOCATION_NAME, start, end), 'GBM fetch history');
  const monthly = aggregateByMonth_(raw);

  // Avis sur la même fenêtre
  const reviewsAgg = executeWithRetry_(() => fetchReviewsMonthly_(start, end), 'GBM reviews history');

  // Reset zone et réécriture avec séparateurs d’années
  const lastRow = sh.getLastRow();
  if (lastRow >= START_ROW) {
    sh.getRange(START_ROW, 1, lastRow - START_ROW + 1, sh.getLastColumn()).clearContent().clearFormat();
  }

  const moisCol = findColByHeaderAliases_(sh, ['mois']);
  let row = START_ROW;
  let currentYear = null;

  Object.keys(monthly).sort().forEach(ym => {
    const y = ym.slice(0,4);
    if (y !== currentYear) {
      sh.getRange(row, moisCol).setValue(y);
      _gbmStyleYearRow_(sh, row, moisCol);
      currentYear = y;
      row++;
    }
    writeGBMRowAt_(sh, row, ym, monthly[ym], reviewsAgg);
    row++;
  });
}
