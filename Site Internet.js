/****************************** CONFIG FEUILLE ******************************/
const SHEET_NAME_SITE = 'Site Internet';
const HEADERS_ROW_SITE = 3;
const START_ROW_SITE = 6;

/****************************** CONFIG GA4 / GSC ****************************/
const GA4_PROPERTY = 'properties/333690282';
const GSC_SITE_URL = 'https://www.nancomcy.fr/';

/****************************** CONFIG MATOMO *******************************/
// Tokens maintenant gérés par Utils.js via Utils_getMatomoToken() et Utils_getMatomoSegmentId()

/****************************** CONFIG MAGNETIS *****************************/
// Token maintenant géré par Utils.js via Utils_getMagnetisApiKey()
const SITE_MAG_FILTER_CHANNELS = ['Tout le SEO'];
const SITE_MAG_MATCH_MODE = 'equals';
const SITE_MAG_COUNT_ANSWERED_ONLY = false;

/****************************** CONFIG PAPERFORM ****************************/
const SITE_PAPERFORM_FORM_IDS = ['xuefmtzg'];
// Token maintenant géré par Utils.js via Utils_getPaperformToken()

/****************************** CONFIG MONDAY (forms) ************************/
const SITE_MONDAY_BOARD_ID = 9950271520;
const SITE_MONDAY_FORM_COLUMN_TITLE = 'Site Internet';
const SITE_MONDAY_FORM_COLUMN_ID = '';
const SITE_MONDAY_MATCH_MODE = 'nonempty';
const SITE_MONDAY_MATCH_VALUES = [];

// Token Monday maintenant géré par Utils.js via Utils_getMondayToken()

/************* CONFIG MONDAY (leads: appels + formulaires) ******************/
const SITE_MONDAY_LEADS_BOARD_ID = SITE_MONDAY_BOARD_ID;
const SITE_MONDAY_LEADS_COL_SOURCE = 'Source du Lead';
const SITE_MONDAY_LEADS_COL_TYPE = 'FORMULAIRE / APPELS';
// ★ nouvelle colonne pour le statut “Lead”
const SITE_MONDAY_LEADS_COL_STATUS = 'Nature du contact'; // ← mets le libellé exact de ta colonne statut
const SITE_MONDAY_LEADS_COL_DATE = '🗓️ Date de création'; // ✅ CORRIGÉ: utilise la date de création
// correspondances (insensibles casse/accents)
const SITE_MONDAY_LEADS_SOURCE_MATCH = ['SITE INTERNET']; // ✅ Valeur exacte dans Monday
const SITE_MONDAY_LEADS_TYPE_MATCH = ['Appel'];
// ★ correspondances pour FORMULAIRE
const SITE_MONDAY_LEADS_FORM_MATCH = ['Formulaire'];
// ★ correspondances pour le statut Lead
const SITE_MONDAY_LEADS_STATUS_MATCH = ['Lead'];

/****************************** HELPERS COMMUNS *****************************/
/****************************** HELPERS COMMUNS (Utilise Utils.js) *****************************/
// Les fonctions suivantes sont maintenant dans Utils.js :
// Utils_toast -> Utils_toast
// Utils_executeWithRetry -> Utils_executeWithRetry
// SITE__normHeader_ -> Utils_normHeader
// Utils_findColByHeaderAliases -> Utils_findColByHeaderAliases
// Utils_isYearSeparatorRow -> Utils_isYearSeparatorRow
// Utils_sheetCellToYYYYMM -> Utils_sheetCellToYYYYMM
// Utils_monthKeyToFr -> Utils_monthKeyToFr
// Utils_setSecondsAsDuration -> Utils_setSecondsAsDuration
// SITE_normalizeDomain_ -> Utils_normalizeDomain


// ✅ Fonctions maintenant dans SheetHelpers.js (élimine 52 lignes de duplication)
// SITE__styleYearRow_ -> SheetHelpers.styleYearRow
// SITE__findOrInsertMonthRow_ -> SheetHelpers.ensureMonthRow


/******************************** GA4 ***************************************/
function ga4FetchMonthly_(property, startDate, endDate) {
  if (typeof AnalyticsData === 'undefined' || !AnalyticsData.Properties) {
    throw new Error('Active le service avancé "AnalyticsData" pour GA4.');
  }
  const request = {
    dateRanges: [{ startDate, endDate }],
    dimensions: [{ name: 'year' }, { name: 'month' }],
    metrics: [{ name: 'sessions' }, { name: 'averageSessionDuration' }, { name: 'bounceRate' }],
  };
  const resp = AnalyticsData.Properties.runReport(request, property);
  const rows = resp.rows || [];
  const out = {};
  rows.forEach(r => {
    const year = r.dimensionValues[0].value;
    const month = r.dimensionValues[1].value.padStart(2, '0');
    const key = `${year}-${month}`;
    out[key] = {
      sessions: Number(r.metricValues[0].value || 0),
      avgSec: Number(r.metricValues[1].value || 0),
      bouncePct: Number(r.metricValues[2].value || 0)
    };
  });
  return out;
}

/******************************** MATOMO (pour visites post-2025) ***********/
function SITE_matomoFetch_(params) {
  const token = Utils_getMatomoToken();
  const MATOMO_BASE_URL = 'https://matomo.aleo.agency';
  const MATOMO_SITE_ID = 1492;
  const full = { module: 'API', format: 'JSON', token_auth: token, idSite: String(MATOMO_SITE_ID), ...params };
  const qs = Object.keys(full).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(full[k])).join('&');
  const url = `${MATOMO_BASE_URL}/index.php?${qs}`;

  Logger.log(`[Matomo Fetch] Appel de l'URL : ${url}`);

  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const responseCode = res.getResponseCode();
  const responseBody = res.getContentText();

  Logger.log(`[Matomo Fetch] Code de réponse : ${responseCode}`);
  Logger.log(`[Matomo Fetch] Réponse brute de l'API : ${responseBody}`);

  if (responseCode >= 300) throw new Error(`Matomo HTTP ${responseCode}: ${responseBody}`);
  return JSON.parse(res.getContentText());
}

function SITE_matomoGetSegmentDefinition_(segmentId) {
  const params = {
    method: 'SegmentEditor.get',
    idSegment: segmentId
  };
  try {
    const res = SITE_matomoFetch_(params);
    if (res && res.definition) {
      Logger.log(`[Matomo Segment] ID ${segmentId} resolved to definition: ${res.definition}`);
      return res.definition;
    }
    Logger.log(`[Matomo Segment] ID ${segmentId} did not return a definition. Response: ${JSON.stringify(res)}`);
    return null;
  } catch (e) {
    Logger.log(`[Matomo Segment] Error fetching definition for ID ${segmentId}: ${e.message}`);
    return null;
  }
}

function SITE_matomoFetchVisitsMonthly_(startDate, endDate) {
  const dateRange = `${startDate},${endDate}`;
  const segmentId = Utils_getMatomoSegmentId();

  Logger.log(`[Matomo Visits] Période: ${dateRange}, Segment ID récupéré: ${segmentId}`);

  // ✅ OPTIMISATION: Cache pour éviter les appels répétitifs
  const cacheKey = `matomo_visits_${startDate}_${endDate}_${segmentId || 'noseg'}`;

  return CacheManager.cachedFetch(
    cacheKey,
    () => {
      const params = {
        method: 'VisitsSummary.get',
        period: 'month',
        date: dateRange,
        format: 'JSON'
      };

      if (segmentId) {
        const def = SITE_matomoGetSegmentDefinition_(segmentId);
        if (def) {
          params.segment = def;
          Logger.log(`[Matomo Visits] Ajout du paramètre segment (résolu): ${params.segment}`);
        } else {
          Logger.log(`[Matomo Visits] Impossible de résoudre le segment ID ${segmentId}. Requête sans segment.`);
        }
      } else {
        Logger.log(`[Matomo Visits] Aucun segment appliqué.`);
      }

      // ✅ OPTIMISATION: Rate limiting pour Matomo
      return RateLimiter.throttle('matomo', () => SITE_matomoFetch_(params));
    },
    CacheManager.CACHE_DURATION.MEDIUM // 30min
  );
}

/**
 * Transforme la réponse Matomo en format compatible avec le reste du code.
 * Matomo retourne: { "2025-11": { nb_visits: 123, avg_time_on_site: 174, bounce_rate: "22 %", ... } }
 * On transforme en: { "2025-11": { sessions: 123, avgSec: 174, bouncePct: 0.22 } }
 * @param {object} matomoResp - Réponse brute de l'API Matomo
 * @returns {object} Objet avec format { "YYYY-MM": { sessions, avgSec, bouncePct } }
 */
function SITE_parseMatomoVisitsResponse_(matomoResp) {
  if (!matomoResp || typeof matomoResp !== 'object') {
    Logger.log('[Matomo Parse] Réponse vide ou invalide');
    return {};
  }

  const result = {};

  for (const monthKey in matomoResp) {
    const data = matomoResp[monthKey];

    // Si c'est un message d'erreur Matomo (string)
    if (data && typeof data === 'string') {
      Logger.log(`[Matomo Parse] Message pour ${monthKey}: ${data}`);
      continue;
    }

    // Extraire les métriques Matomo
    if (data && typeof data === 'object') {
      const sessions = Number(data.nb_visits || 0);
      const avgSec = Number(data.avg_time_on_site || 0);

      // bounce_rate est une string "22 %" → convertir en 0.22
      let bouncePct = 0;
      if (data.bounce_rate) {
        const bounceStr = String(data.bounce_rate).replace(/[^0-9.]/g, ''); // Enlever "%" et espaces
        bouncePct = Number(bounceStr) / 100;
      }

      result[monthKey] = { sessions, avgSec, bouncePct };
      Logger.log(`[Matomo Parse] ${monthKey}: ${sessions} visites, ${avgSec}s moyenne, ${(bouncePct * 100).toFixed(1)}% rebond`);
    }
  }

  return result;
}

/******************************** Search Console ****************************/
function gscQuery_(siteUrl, req) {
  if (typeof SearchConsole !== 'undefined' &&
    SearchConsole.Searchanalytics &&
    SearchConsole.Searchanalytics.query) {
    try { return SearchConsole.Searchanalytics.query(siteUrl, req); }
    catch (e) { return SearchConsole.Searchanalytics.query(req, siteUrl); }
  }
  if (typeof Webmasters !== 'undefined' &&
    Webmasters.Searchanalytics &&
    Webmasters.Searchanalytics.query) {
    return Webmasters.Searchanalytics.query(siteUrl, req);
  }
  var token = ScriptApp.getOAuthToken();
  var endpoint = 'https://www.googleapis.com/webmasters/v3/sites/' +
    encodeURIComponent(siteUrl) + '/searchAnalytics/query';
  var res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(req),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('GSC REST HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}
function gscFetchMonthly_(siteUrl, startDate, endDate) {
  const req = { startDate, endDate, dimensions: ['date'], rowLimit: 25000 };
  const resp = gscQuery_(siteUrl, req);
  const rows = (resp && resp.rows) || [];
  const buckets = {};
  rows.forEach(row => {
    const d = row.keys && row.keys[0];
    if (!d) return;
    const ym = d.slice(0, 7);
    const imp = Number(row.impressions || 0);
    const clk = Number(row.clicks || 0);
    if (!buckets[ym]) buckets[ym] = { imp: 0, clk: 0 };
    buckets[ym].imp += imp;
    buckets[ym].clk += clk;
  });
  const out = {};
  Object.keys(buckets).forEach(ym => {
    const b = buckets[ym];
    out[ym] = { impressions: b.imp, clicks: b.clk };
  });
  return out;
}

/******************************** Magnetis **********************************/

function SITE_getChannelName_(c) {
  const candidates = ['channel_name', 'channel', 'analysis.channel', 'analytics.channel', 'utm.channel', 'session.channel'];
  for (const p of candidates) { const v = Utils_valAt(c, p); if (v != null && String(v).trim() !== '') return String(v).trim().toLowerCase(); }
  return '';
}
function SITE_magChannelOk_(name) {
  if (!SITE_MAG_FILTER_CHANNELS || SITE_MAG_FILTER_CHANNELS.length === 0) return true;
  const list = SITE_MAG_FILTER_CHANNELS.map(x => String(x).toLowerCase());
  switch (SITE_MAG_MATCH_MODE) {
    case 'includes': return list.some(a => name.includes(a));
    case 'regex': return list.some(rx => new RegExp(rx, 'i').test(name));
    default: return list.some(a => name === a);
  }
}
function SITE_magnetisFetchCalls_(fromDate, toDate) {
  const apiKey = Utils_getMagnetisApiKey();
  const base = 'https://api.magnetis.io/calls';
  const fmtUTC = d => Utilities.formatDate(d, 'UTC', 'yyyyMMddHHmmss');
  let start = new Date(fromDate), end = new Date(toDate);
  if (start > end) { const t = start; start = end; end = t; }
  const fromStr = fmtUTC(start), toStr = fmtUTC(end);
  let page = 1, all = [];
  while (true) {
    const params = { from: fromStr, to: toStr, limit: 250, page, analysis: 1 };
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
function SITE_magnetisAggregateMonthlyCounts_(calls) {
  const bucket = {};
  calls.forEach(c => {
    const ch = SITE_getChannelName_(c); if (!SITE_magChannelOk_(ch)) return;
    const raw = c.start_at || c.created_at || c.date || c.started_at;
    if (!raw) return;
    const d = new Date(raw); if (isNaN(d.getTime())) return;
    if (SITE_MAG_COUNT_ANSWERED_ONLY) {
      const dur = Number(c.duration || (c.analysis && c.analysis.duration) || 0);
      if (!(dur > 0)) return;
    }
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    bucket[key] = (bucket[key] || 0) + 1;
  });
  return bucket;
}

/******************************** Paperform *********************************/
function SITE_extractPaperformArray_(json) {
  if (!json) return [];
  if (Array.isArray(json)) return json;
  if (json.results) {
    if (Array.isArray(json.results)) return json.results;
    if (Array.isArray(json.results.submissions)) return json.results.submissions;
    if (Array.isArray(json.results.data)) return json.results.data;
    for (const k in json.results) { if (Array.isArray(json.results[k])) return json.results[k]; }
  }
  if (Array.isArray(json.data)) return json.data;
  for (const k in json) { if (Array.isArray(json[k])) return json[k]; }
  return [];
}
function SITE_parsePaperformHumanDateToUTC_(str) {
  if (!str || typeof str !== 'string') return new Date(NaN);
  const s = str.trim().replace(/\s+/g, ' ');
  const MM = { jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3, may: 4, jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7, sep: 8, sept: 8, september: 8, oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11 };
  const m = s.match(/^([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\d{2,4})[, ]\s*(\d{1,2}):(\d{2})\s*(am|pm)$/i);
  if (!m) return new Date(NaN);
  const mon = MM[m[1].toLowerCase()]; let year = +m[3]; let h = +m[4]; const min = +m[5]; const ap = m[6].toLowerCase();
  if (isNaN(mon) || isNaN(year) || isNaN(h) || isNaN(min)) return new Date(NaN);
  if (year < 100) year = 2000 + year; h = h % 12; if (ap === 'pm') h += 12;
  return new Date(Date.UTC(year, mon, +m[2], h, min, 0));
}
function SITE_getSubmissionDate_(s) {
  const cand = [s.submitted_at, s.submittedAt, s.submitted_at_utc, s.submittedAtUTC, s.created_at, s.createdAt, s.created_at_utc, s.createdAtUTC, s.date, s.timestamp, s.time_submitted, s.timeSubmitted, s.created, s.submitted, (s.meta && (s.meta.submitted_at || s.meta.created_at))];
  for (let v of cand) {
    if (v == null) continue;
    if (typeof v === 'number') { const ms = v < 1e12 ? v * 1000 : v; const d = new Date(ms); if (!isNaN(d.getTime())) return d; continue; }
    if (typeof v === 'string') {
      const dStd = new Date(v); if (!isNaN(dStd.getTime())) return dStd;
      const dHuman = SITE_parsePaperformHumanDateToUTC_(v); if (dHuman && !isNaN(dHuman.getTime())) return dHuman;
      if (/^\d+$/.test(v)) { const n = +v, ms = n < 1e12 ? n * 1000 : n; const dNum = new Date(ms); if (!isNaN(dNum.getTime())) return dNum; }
    }
  }
  return new Date(NaN);
}
function SITE_paperformFetchSubmissionsInRange_(slugOrId, fromDate, toDate) {
  const token = Utils_getPaperformToken();
  const base = `https://api.paperform.co/v1/forms/${encodeURIComponent(slugOrId)}/submissions`;
  const limit = 100; let all = []; let mode = 'skip', skip = 0, page = 1;
  const fromTs = fromDate.getTime(), toTs = toDate.getTime();
  while (true) {
    const url = (mode === 'skip') ? `${base}?limit=${limit}&skip=${skip}` : `${base}?limit=${limit}&page=${page}`;
    const res = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' }, muteHttpExceptions: true });
    if (res.getResponseCode() === 422 && mode === 'skip') { mode = 'page'; page = 1; continue; }
    if (res.getResponseCode() >= 300) throw new Error(`Paperform HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    const batch = SITE_extractPaperformArray_(JSON.parse(res.getContentText())); if (!batch.length) break;
    for (const s of batch) { const d = SITE_getSubmissionDate_(s); if (!isNaN(d.getTime())) { const t = d.getTime(); if (t >= fromTs && t <= toTs) all.push(s); } }
    if (mode === 'skip') skip += batch.length; else page += 1;
    if (batch.length < limit) break;
    Utilities.sleep(120);
  }
  return all;
}
function SITE_paperformCountsByMonth_(fromDate, toDate) {
  const bucket = {};
  SITE_PAPERFORM_FORM_IDS.forEach(fid => {
    const subs = SITE_paperformFetchSubmissionsInRange_(fid, fromDate, toDate);
    subs.forEach(s => {
      const d = SITE_getSubmissionDate_(s); if (isNaN(d.getTime())) return;
      const key = d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
      bucket[key] = (bucket[key] || 0) + 1;
    });
  });
  return bucket;
}

/******************************** Monday (général) **************************/

// Fonction mondayGraphQL maintenant centralisée dans Utils.js via Utils_mondayGraphQL()
function SITE_mondayResolveColumnId_(boardId, titleOrId) {
  if (!titleOrId) return '';
  const looksId = /^[a-z0-9_]+$/i.test(titleOrId) && !Utils_normHeader(titleOrId).includes(' ');
  if (looksId) return titleOrId;

  const want = Utils_normHeader(titleOrId);

  // ✅ OPTIMISATION: Utilise le cache Monday au lieu d'appeler l'API à chaque fois
  const cols = CacheManager.getMondayBoardColumns(boardId);

  let best = '';
  for (const c of cols) { if (Utils_normHeader(c.title) === want) { best = c.id; break; } }
  if (!best) { for (const c of cols) { if (Utils_normHeader(c.title).includes(want)) { best = c.id; break; } } }
  return best;
}
function SITE_mondayValueMatches_(txt) {
  const t = Utils_normHeader(txt || '');
  switch (SITE_MONDAY_MATCH_MODE) {
    case 'equals': return SITE_MONDAY_MATCH_VALUES.some(v => t === Utils_normHeader(v));
    case 'includes': return SITE_MONDAY_MATCH_VALUES.some(v => t.includes(Utils_normHeader(v)));
    default: return t.length > 0;
  }
}
function SITE_mondayFormCountsByMonth_(boardId, colTitleOrId, fromUTC, toUTC) {
  const colId = SITE_MONDAY_FORM_COLUMN_ID ? SITE_MONDAY_FORM_COLUMN_ID : SITE_mondayResolveColumnId_(boardId, colTitleOrId);
  if (!colId) throw new Error("Colonne Monday 'Site Internet' introuvable.");
  const counts = {}; let cursor = null;
  do {
    const q = `
      query($bid:[ID!], $cursor:String, $col:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{ id state created_at column_values(ids:$col){ id text } }
          }
        }
      }`;
    // ✅ OPTIMISATION: Rate limiting Monday
    const d = RateLimiter.throttle('monday', () =>
      Utils_mondayGraphQL(q, { bid: [Number(boardId)], cursor, col: [colId] })
    );
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;
    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;
      const dt = new Date(it.created_at); if (isNaN(dt.getTime())) return;
      if (dt < fromUTC || dt > toUTC) return;
      const txt = (it.column_values && it.column_values[0] && it.column_values[0].text) || '';
      if (!SITE_mondayValueMatches_(txt)) return;
      const key = dt.getUTCFullYear() + '-' + String(dt.getUTCMonth() + 1).padStart(2, '0');
      counts[key] = (counts[key] || 0) + 1;
    });
    cursor = page.cursor;
  } while (cursor);
  return counts;
}



/************ Monday — Lead calls & lead forms ******************************/
function SITE_resolveTwoCols_(boardId, titleOrId1, titleOrId2) {
  const id1 = SITE_mondayResolveColumnId_(boardId, titleOrId1);
  const id2 = SITE_mondayResolveColumnId_(boardId, titleOrId2);
  if (!id1 || !id2) throw new Error('Colonnes Monday (source/type) introuvables.');
  return { id1, id2 };
}

// Appels lead
// { 'YYYY-MM': n } — APPELS lead = Source match + Type "appel" + Statut "Lead"
function SITE_mondayLeadCallsCountsByMonth_(boardId, colSource, colType, colStatus, fromUTC, toUTC) {
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId = SITE_mondayResolveColumnId_(boardId, colType);
  const statusId = colStatus ? SITE_mondayResolveColumnId_(boardId, colStatus) : null;
  const dateId = (typeof SITE_MONDAY_LEADS_COL_DATE !== 'undefined' && SITE_MONDAY_LEADS_COL_DATE)
    ? SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_DATE)
    : null;
  if (!sourceId || !typeId) throw new Error('Colonnes Monday (source/type) introuvables.');

  const counts = {};
  let cursor = null;

  // ✅ OPTIMISATION: Convertir textes → index Monday pour les filtres
  const sourceIndexes = MONDAY_textToIndexes_(boardId, sourceId, SITE_MONDAY_LEADS_SOURCE_MATCH);
  const typeIndexes = MONDAY_textToIndexes_(boardId, typeId, SITE_MONDAY_LEADS_TYPE_MATCH);
  const statusIndexes = statusId ? MONDAY_textToIndexes_(boardId, statusId, SITE_MONDAY_LEADS_STATUS_MATCH) : [];

  do {
    // ✅ OPTIMISATION MAJEURE: Filtres Monday côté serveur (réduction 99%)
    const rules = [
      { column_id: sourceId, compare_value: sourceIndexes, operator: "any_of" },
      { column_id: typeId, compare_value: typeIndexes, operator: "any_of" }
    ];

    if (statusId && statusIndexes.length > 0) {
      rules.push({ column_id: statusId, compare_value: statusIndexes, operator: "any_of" });
    }

    const hasFilters = !cursor && sourceIndexes.length > 0 && typeIndexes.length > 0;
    const queryParams = hasFilters ? { rules, operator: "and" } : null;
    const colIds = [sourceId, typeId].concat(statusId ? [statusId] : []).concat(dateId ? [dateId] : []);

    // ✅ Construire la query dynamiquement pour éviter variables non utilisées
    const querySignature = cursor
      ? '$bid:[ID!], $cursor:String, $cols:[String!]'
      : (hasFilters ? '$bid:[ID!], $cols:[String!], $qp:ItemsQuery' : '$bid:[ID!], $cols:[String!]');

    const queryParams_str = cursor
      ? 'limit:500, cursor:$cursor'
      : (hasFilters ? 'limit:500, query_params:$qp' : 'limit:500');

    const q = `
      query(${querySignature}){
        boards(ids:$bid){
          items_page(${queryParams_str}){
            cursor
            items{
              id state created_at
              column_values(ids:$cols){ id text value }
            }
          }
        }
      }`;

    // ✅ OPTIMISATION: Rate limiting Monday
    const variables = { bid: [Number(boardId)], cols: colIds };
    if (cursor) variables.cursor = cursor;
    else if (queryParams) variables.qp = queryParams;

    const d = RateLimiter.throttle('monday', () => Utils_mondayGraphQL(q, variables));
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page; if (!page) break;

    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;

      const createdAt = new Date(it.created_at);
      const cv = it.column_values || [];
      const srcText = (cv.find(c => c.id === sourceId) || {}).text || '';
      const typText = (cv.find(c => c.id === typeId) || {}).text || '';
      const sts = statusId ? ((cv.find(c => c.id === statusId) || {}).text || '') : '';
      const dcv = dateId ? ((cv.find(c => c.id === dateId) || {}).value || '') : '';
      const when = SITE_parseMondayDateValue_(dcv) || createdAt;

      if (!(when >= fromUTC && when <= toUTC)) return;

      // Source : ÉGALITÉ stricte à une des valeurs autorisées
      const sourceMatch = Utils_textEqualsAny(srcText, SITE_MONDAY_LEADS_SOURCE_MATCH);
      // Type : doit contenir "appel"
      const typeMatch = Utils_textIncludesAny(typText, SITE_MONDAY_LEADS_TYPE_MATCH);
      // ★ NOUVEAU : Statut doit être "Lead" si statusId fourni
      if (statusId && !Utils_textIncludesAny(sts, SITE_MONDAY_LEADS_STATUS_MATCH)) return;

      if (sourceMatch && typeMatch) {
        const key = when.getUTCFullYear() + '-' + String(when.getUTCMonth() + 1).padStart(2, '0');
        counts[key] = (counts[key] || 0) + 1;
      }
    });

    cursor = page.cursor;
  } while (cursor);

  return counts;
}

// ★ Formulaires lead (Source=Site Internet, Type=Formulaire, Statut=Lead)
// { 'YYYY-MM': n } — FORMULAIRES lead = Source match + Type "formulaire" + Statut "lead"
function SITE_mondayLeadFormsCountsByMonth_(boardId, colSource, colType, colStatus, fromUTC, toUTC) {
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId = SITE_mondayResolveColumnId_(boardId, colType);
  const statusId = colStatus ? SITE_mondayResolveColumnId_(boardId, colStatus) : null;
  const dateId = SITE_MONDAY_LEADS_COL_DATE ? SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_DATE) : null;
  if (!sourceId || !typeId) throw new Error('Colonnes Monday (source/type) introuvables.');

  const counts = {};
  let cursor = null;

  // ✅ OPTIMISATION: Convertir textes → index Monday pour les filtres
  const sourceIndexes = MONDAY_textToIndexes_(boardId, sourceId, SITE_MONDAY_LEADS_SOURCE_MATCH);
  const typeIndexes = MONDAY_textToIndexes_(boardId, typeId, SITE_MONDAY_LEADS_FORM_MATCH);
  const statusIndexes = statusId ? MONDAY_textToIndexes_(boardId, statusId, SITE_MONDAY_LEADS_STATUS_MATCH) : [];

  do {
    // ✅ OPTIMISATION MAJEURE: Filtres Monday côté serveur (réduction 99%)
    const rules = [
      { column_id: sourceId, compare_value: sourceIndexes, operator: "any_of" },
      { column_id: typeId, compare_value: typeIndexes, operator: "any_of" }
    ];

    if (statusId && statusIndexes.length > 0) {
      rules.push({ column_id: statusId, compare_value: statusIndexes, operator: "any_of" });
    }

    const hasFilters = !cursor && sourceIndexes.length > 0 && typeIndexes.length > 0;
    const queryParams = hasFilters ? { rules, operator: "and" } : null;
    const colIds = [sourceId, typeId].concat(statusId ? [statusId] : []).concat(dateId ? [dateId] : []);

    const querySignature = cursor
      ? '$bid:[ID!], $cursor:String, $cols:[String!]'
      : (hasFilters ? '$bid:[ID!], $cols:[String!], $qp:ItemsQuery' : '$bid:[ID!], $cols:[String!]');

    const queryParams_str = cursor
      ? 'limit:500, cursor:$cursor'
      : (hasFilters ? 'limit:500, query_params:$qp' : 'limit:500');

    const q = `
      query(${querySignature}){
        boards(ids:$bid){
          items_page(${queryParams_str}){
            cursor
            items{
              id state created_at
              column_values(ids:$cols){ id text value }
            }
          }
        }
      }`;

    const variables = { bid: [Number(boardId)], cols: colIds };
    if (cursor) variables.cursor = cursor;
    else if (queryParams) variables.qp = queryParams;

    const d = RateLimiter.throttle('monday', () => Utils_mondayGraphQL(q, variables));
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page; if (!page) break;

    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;
      const createdAt = new Date(it.created_at);
      const cv = it.column_values || [];
      const src = (cv.find(c => c.id === sourceId) || {}).text || '';
      const typ = (cv.find(c => c.id === typeId) || {}).text || '';
      const sts = statusId ? ((cv.find(c => c.id === statusId) || {}).text || '') : '';
      const dcv = dateId ? ((cv.find(c => c.id === dateId) || {}).value || '') : '';
      const when = SITE_parseMondayDateValue_(dcv) || createdAt;

      if (!(when >= fromUTC && when <= toUTC)) return;
      if (!Utils_textIncludesAny(src, SITE_MONDAY_LEADS_SOURCE_MATCH)) return;
      if (!Utils_textIncludesAny(typ, SITE_MONDAY_LEADS_FORM_MATCH)) return;
      if (statusId && !Utils_textIncludesAny(sts, SITE_MONDAY_LEADS_STATUS_MATCH)) return;

      const key = when.getUTCFullYear() + '-' + String(when.getUTCMonth() + 1).padStart(2, '0');
      counts[key] = (counts[key] || 0) + 1;
    });

    cursor = page.cursor;
  } while (cursor);

  return counts;
}


/********************* FUSION “Nombre de formulaire” ************************/
function SITE_formsCountsByMonth_(fromDateUTC, toDateUTC) {
  const paper = SITE_paperformCountsByMonth_(fromDateUTC, toDateUTC);
  const monday = SITE_mondayFormCountsByMonth_(SITE_MONDAY_BOARD_ID, SITE_MONDAY_FORM_COLUMN_TITLE, fromDateUTC, toDateUTC);
  const keys = Array.from(new Set(Object.keys(paper).concat(Object.keys(monday))));
  const out = {}; keys.forEach(k => out[k] = (paper[k] || 0) + (monday[k] || 0));
  return out;
}



/**************************** HELPERS SPÉCIFIQUES FEUILLE *************************/

// Colonne "Budget investi" SÉCURISÉE pour le site
function SITE_findBudgetColSafe_(sheet) {
  const headerRow = HEADERS_ROW_SITE || 3;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const HN = headers.map(Utils_normHeader);

  // Colonnes à NE JAMAIS prendre pour le budget
  const forbiddenRe = /(contact|contacts|lead|leads|signature|signatures|roas|roi|rebond|bounce|ctr|taux de clics|cpl)/;

  // 1) matches EXACTS
  const exact = ['budget investi', 'budget', 'depenses', 'dépenses'];
  for (let i = 0; i < HN.length; i++) {
    const h = HN[i];
    if (!h || forbiddenRe.test(h)) continue;
    if (exact.includes(h)) return i + 1;
  }

  // 2) alias élargis
  const aliases = ['budget investi', 'budget', 'depenses', 'dépenses', 'spend', 'cost'];
  const wanted = aliases.map(Utils_normHeader);
  for (let i = 0; i < HN.length; i++) {
    const h = HN[i];
    if (!h || forbiddenRe.test(h)) continue;
    if (wanted.some(w => h === w || h.includes(w))) return i + 1;
  }

  Logger.log('[Site/findBudgetColSafe] Aucune colonne budget trouvée.');
  return 0;
}

// VRAI si la colonne est listée comme ayant une formule dans Sheet_Structures.md
// VRAI si la colonne est listée comme ayant une formule dans Sheet_Structures.md
function SITE_isProtectedHeader_(sh, col) {
  if (!col) return false;
  const headerVal = sh.getRange(HEADERS_ROW_SITE, col).getValue();
  const h = Utils_normHeader(headerVal);

  // Liste EXACTE des colonnes avec formules pour "Site Internet"
  // tirée de Sheet_Structures.md
  const protectedHeaders = [
    'nombre de lead', 'nombre de contacts', 'ctr',
    'montant moyen signature', 'total montant signe', 'nombre de signature'
  ].map(Utils_normHeader);

  // La colonne est protégée si son nom normalisé est dans la liste.
  return protectedHeaders.includes(h);
}



/**************************** ÉCRITURE DANS LA FEUILLE ****************************/
function SITE_findCols_(sh) {
  return {
    mois: Utils_findColByHeaderAliases(sh, ['mois'], HEADERS_ROW_SITE),

    // budget sécurisé (pas utilisé en écriture pour l’instant, mais prêt)
    budget: SITE_findBudgetColSafe_(sh),

    // CPL = souvent une formule → pas écrite par le script
    cpl: Utils_findColByHeaderAliases(sh, ['cout par contact', 'coût par contact', 'cpl'], HEADERS_ROW_SITE),

    impr: Utils_findColByHeaderAliases(sh, ['impressions', 'nb impressions', 'nombre dimpressions'], HEADERS_ROW_SITE),
    visits: Utils_findColByHeaderAliases(sh, ['nombre de visites', 'visites', 'sessions', 'sessions (ga4)'], HEADERS_ROW_SITE),

    // CTR (calculé par le script)
    ctr: Utils_findColByHeaderAliases(sh, ['ctr', 'taux de clics', 'taux de clic'], HEADERS_ROW_SITE),

    dur: Utils_findColByHeaderAliases(sh, ['duree moyenne des visites', 'durée moyenne des visites'], HEADERS_ROW_SITE),
    bounce: Utils_findColByHeaderAliases(sh, ['taux de rebond', 'bounce'], HEADERS_ROW_SITE),

    appels: Utils_findColByHeaderAliases(sh, ["nombre d'appels", "nombre dappels", "appels"], HEADERS_ROW_SITE),
    appelsLead: Utils_findColByHeaderAliases(sh, ["nombre d'appels lead", "appels lead", "lead appels"], HEADERS_ROW_SITE),

    forms: Utils_findColByHeaderAliases(sh, ['nombre de formulaire', 'formulaires', 'forms', 'soumissions'], HEADERS_ROW_SITE),
    formsLead: Utils_findColByHeaderAliases(sh, ['nombre de formulaires lead', 'formulaires lead', 'lead formulaires'], HEADERS_ROW_SITE),

    // Colonne "Nombre de contacts" (contact = appels + formulaires)
    contacts: Utils_findColByHeaderAliases(sh, ['nombre de contacts', 'contacts', 'leads'], HEADERS_ROW_SITE),

    // Ancien taux de conversion générique (si tu en as encore un)
    conv: Utils_findColByHeaderAliases(sh, ['taux de conversion', 'conversion rate', 'cr'], HEADERS_ROW_SITE),
  };
}


function SITE_writeMonthRow_(sh, cols, ymKey, values) {
  const row = SheetHelpers.ensureMonthRow(sh, cols.mois, ymKey, START_ROW_SITE, 'Site');
  Logger.log(`[Site/WRITE] Écriture des KPIs pour ${ymKey} sur la ligne ${row}`);

  // ✅ OPTIMISATION BATCH: Préparer toutes les updates
  const updates = [];
  if (values.impressions != null) updates.push({ col: cols.impr, value: values.impressions, format: null });
  if (values.sessions != null) updates.push({ col: cols.visits, value: values.sessions, format: null });
  if (values.avgSec != null) updates.push({ col: cols.dur, value: values.avgSec / 86400, format: '[h]:mm:ss' });
  if (values.calls != null) updates.push({ col: cols.appels, value: values.calls, format: null });
  if (values.forms != null) updates.push({ col: cols.forms, value: values.forms, format: null });
  if (values.leadCalls != null) updates.push({ col: cols.appelsLead, value: values.leadCalls, format: null });
  if (values.leadForms != null) updates.push({ col: cols.formsLead, value: values.leadForms, format: null });
  if (values.bounce != null) updates.push({ col: cols.bounce, value: values.bounce, format: '0.00%' });

  // CTR calculé
  const impressions = values.impressions ?? 0;
  const sessions = values.sessions ?? 0;
  if (impressions > 0 && cols.ctr) {
    updates.push({ col: cols.ctr, value: sessions / impressions, format: '0.00%' });
  }

  if (updates.length === 0) return;

  // 🚀 BATCH: Récupérer toutes les formules EN UNE FOIS
  const cols_list = updates.map(u => u.col);
  const minCol = Math.min(...cols_list);
  const maxCol = Math.max(...cols_list);
  const numCols = maxCol - minCol + 1;
  const formulas = sh.getRange(row, minCol, 1, numCols).getFormulas()[0];

  // Filtrer les cellules avec formules + protégées
  const toWrite = updates.filter(u => {
    if (SITE_isProtectedHeader_(sh, u.col)) {
      Logger.log(`[PROTECTED] Colonne ${sh.getRange(HEADERS_ROW_SITE, u.col).getValue()} ignorée`);
      return false;
    }
    const idx = u.col - minCol;
    if (formulas[idx] && formulas[idx].trim()) {
      return false; // Skip formules
    }
    return true;
  });

  // Ã‰crire chaque valeur individuellement en prÃ©servant les formules
  toWrite.forEach(u => {
    Utils_setPreserveFormula(sh, row, u.col, u.value, u.format);
  });

  // Toujours Ã©crire le mois
  sh.getRange(row, cols.mois).setValue(Utils_monthKeyToFr(ymKey));
}


/**************************** RUNNERS ***************************************/
const SITE_MONTHS_BACK = 21;

function run_Site_FullHistory() {
  PerformanceLogger.start('Site_FullHistory');

  return ErrorHandler.wrapFunction(() => {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SITE);
    Validators.validateSheetExists(SpreadsheetApp.getActive(), SHEET_NAME_SITE, 'run_Site_FullHistory');

    const cols = SITE_findCols_(sh);
    if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d'entêtes).");

    const today = new Date();
    // Début au 1er janvier 2024
    const start = new Date(2024, 0, 1);
    const end = new Date(today.getFullYear(), today.getMonth(), 0);

    Validators.validateDateRange(start, end, 'run_Site_FullHistory');

    const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const matomoCutoffDate = new Date(2025, 10, 1); // 1er Novembre 2025

    Utils_toast('Site: collecte…', 'Site Internet', 4);

    const ga = ErrorHandler.executeWithRetry(() => ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr), 'Site - GA4 monthly', 3, 1000);
    const gsc = Utils_executeWithRetry(() => gscFetchMonthly_(GSC_SITE_URL, startStr, endStr), 'GSC monthly');
    const callsByMonth = Utils_executeWithRetry(() => SITE_magnetisAggregateMonthlyCounts_(SITE_magnetisFetchCalls_(start, end)), 'Magnetis calls');
    const formsByMonth = Utils_executeWithRetry(() => SITE_formsCountsByMonth_(new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)), new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))), 'Forms (Paper+Monday)');

    let matomoVisits = {};
    // On ne lance l'appel Matomo que si la période de fin est après la date de bascule
    if (end >= matomoCutoffDate) {
      const matomoStartDate = start > matomoCutoffDate ? start : matomoCutoffDate;
      const matomoStartStr = Utilities.formatDate(matomoStartDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const rawMatomo = Utils_executeWithRetry(
        () => SITE_matomoFetchVisitsMonthly_(matomoStartStr, endStr), 'Matomo Visits'
      );
      matomoVisits = SITE_parseMatomoVisitsResponse_(rawMatomo);
    }

    const leadCallsByMonth = Utils_executeWithRetry(
      () => SITE_mondayLeadCallsCountsByMonth_(
        SITE_MONDAY_LEADS_BOARD_ID,
        SITE_MONDAY_LEADS_COL_SOURCE,
        SITE_MONDAY_LEADS_COL_TYPE,
        SITE_MONDAY_LEADS_COL_STATUS,
        new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
        new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
      ),
      'Monday lead calls'
    );

    const leadFormsByMonth = Utils_executeWithRetry(
      () => SITE_mondayLeadFormsCountsByMonth_(
        SITE_MONDAY_LEADS_BOARD_ID,
        SITE_MONDAY_LEADS_COL_SOURCE,
        SITE_MONDAY_LEADS_COL_TYPE,
        SITE_MONDAY_LEADS_COL_STATUS,
        new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
        new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
      ),
      'Monday lead forms'
    );

    const keys = Array.from(new Set([]
      .concat(Object.keys(ga), Object.keys(gsc), Object.keys(callsByMonth), Object.keys(formsByMonth),
        Object.keys(leadCallsByMonth), Object.keys(leadFormsByMonth))
    )).sort();



    keys.forEach(ym => {
      const monthDate = new Date(parseInt(ym.slice(0, 4)), parseInt(ym.slice(5, 7)) - 1, 1);
      const useMatomo = monthDate >= matomoCutoffDate;

      const t1 = new Date();
      Logger.log(`[PERF] Début construction values pour ${ym}: ${t1.toISOString()}`);

      const values = {
        // Utilise Matomo pour les sessions si la date est >= Nov 2025, sinon GA4
        sessions: useMatomo ? (matomoVisits[ym]?.sessions ?? null) : (ga[ym]?.sessions ?? null),
        // Si Matomo (>= nov 2025), utiliser Matomo pour avgSec et bounce, sinon GA4
        avgSec: useMatomo ? (matomoVisits[ym]?.avgSec ?? null) : (ga[ym]?.avgSec ?? null),
        impressions: gsc[ym]?.impressions ?? null,
        calls: callsByMonth[ym] ?? null,
        forms: formsByMonth[ym] ?? null,
        leadCalls: leadCallsByMonth[ym] ?? null,
        leadForms: leadFormsByMonth[ym] ?? null,
        bounce: useMatomo ? (matomoVisits[ym]?.bouncePct ?? null) : (ga[ym]?.bouncePct ?? null)
      };

      const t2 = new Date();
      Logger.log(`[PERF] Fin construction values (${t2 - t1}ms). Début SITE_writeMonthRow_: ${t2.toISOString()}`);

      SITE_writeMonthRow_(sh, cols, ym, values);

      const t3 = new Date();
      Logger.log(`[PERF] Fin SITE_writeMonthRow_ (${t3 - t2}ms)`);
    });

  }, 'run_Site_FullHistory', { rethrow: true });

  Utils_toast('Site: Full history terminé ✅', 'Site Internet', 5);
  PerformanceLogger.end('Site_FullHistory');
}

function run_Site_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SITE);
  if (!sh) throw new Error(`Onglet '${SHEET_NAME_SITE}' introuvable`);
  const cols = SITE_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');
  const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const matomoCutoffDate = new Date(2025, 10, 1); // 1er Novembre 2025
  const useMatomo = start >= matomoCutoffDate;

  let sessionsData;
  if (useMatomo) {
    const rawMatomo = Utils_executeWithRetry(() => SITE_matomoFetchVisitsMonthly_(startStr, endStr), 'Matomo N-1');
    sessionsData = SITE_parseMatomoVisitsResponse_(rawMatomo);
  } else {
    sessionsData = Utils_executeWithRetry(() => ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr), 'GA4 N-1');
  }

  const gsc = Utils_executeWithRetry(() => gscFetchMonthly_(GSC_SITE_URL, startStr, endStr), 'GSC N-1');
  const callsByMonth = Utils_executeWithRetry(() => SITE_magnetisAggregateMonthlyCounts_(SITE_magnetisFetchCalls_(start, end)), 'Magnetis N-1');
  const formsByMonth = Utils_executeWithRetry(() => SITE_formsCountsByMonth_(new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)), new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))), 'Forms N-1');

  // Pour la durée moyenne, on continue d'utiliser GA4 pour l'instant
  // Si vous souhaitez aussi la basculer sur Matomo, il faudra une logique similaire
  const gaForDuration = Utils_executeWithRetry(() => ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr), 'GA4 N-1 (duration)');

  const leadCallsByMonth = Utils_executeWithRetry(
    () => SITE_mondayLeadCallsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      SITE_MONDAY_LEADS_COL_STATUS,
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
    ),
    'Monday lead calls N-1'
  );

  const leadFormsByMonth = Utils_executeWithRetry(
    () => SITE_mondayLeadFormsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      SITE_MONDAY_LEADS_COL_STATUS,
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
    ),
    'Monday lead forms N-1'
  );



  const t1 = new Date();
  Logger.log(`[PERF] Début construction values: ${t1.toISOString()}`);

  const values = {
    sessions: sessionsData[ymKey]?.sessions ?? null,
    // Si Matomo (>= nov 2025), utiliser Matomo pour avgSec et bounce, sinon GA4
    avgSec: useMatomo ? (sessionsData[ymKey]?.avgSec ?? null) : (gaForDuration[ymKey]?.avgSec ?? null),
    impressions: gsc[ymKey]?.impressions ?? null,
    calls: callsByMonth[ymKey] ?? null,
    forms: formsByMonth[ymKey] ?? null,
    leadCalls: leadCallsByMonth[ymKey] ?? null,
    leadForms: leadFormsByMonth[ymKey] ?? null,
    bounce: useMatomo ? (sessionsData[ymKey]?.bouncePct ?? null) : (gaForDuration[ymKey]?.bouncePct ?? null)
  };

  const t2 = new Date();
  Logger.log(`[PERF] Fin construction values (${t2 - t1}ms). Début SITE_writeMonthRow_: ${t2.toISOString()}`);

  SITE_writeMonthRow_(sh, cols, ymKey, values);

  const t3 = new Date();
  Logger.log(`[PERF] Fin SITE_writeMonthRow_ (${t3 - t2}ms)`);

  Utils_toast('Site: N-1 mis à jour ✅', 'Site Internet', 5);
}

/*********************** TEST — Leads (appels & formulaires) ************************/

/** Liste détaillée des ITEMS Monday correspondant aux APPELS lead
 *  (Source ∈ SITE_MONDAY_LEADS_SOURCE_MATCH, Type ∈ SITE_MONDAY_LEADS_TYPE_MATCH)
 */
function SITE_mondayFetchLeadCallsItems_(boardId, colSource, colType, fromUTC, toUTC) {
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId = SITE_mondayResolveColumnId_(boardId, colType);
  if (!sourceId || !typeId) throw new Error('Colonnes Monday (source/type) introuvables.');

  const out = []; let cursor = null;
  do {
    const q = `
      query($bid:[ID!], $cursor:String, $cols:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id name state created_at
              column_values(ids:$cols){ id text }
            }
          }
        }
      }`;
    const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)], cursor, cols: [sourceId, typeId] });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;

    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;
      const dt = new Date(it.created_at); if (isNaN(dt.getTime())) return;
      if (dt < fromUTC || dt > toUTC) return;

      const cv = it.column_values || [];
      const srcText = (cv.find(c => c.id === sourceId) || {}).text || '';
      const typText = (cv.find(c => c.id === typeId) || {}).text || '';
      if (!Utils_textIncludesAny(srcText, SITE_MONDAY_LEADS_SOURCE_MATCH)) return;
      if (!Utils_textIncludesAny(typText, SITE_MONDAY_LEADS_TYPE_MATCH)) return;

      out.push({ id: it.id, name: it.name, created_at: dt.toISOString(), source: srcText, type: typText });
    });

    cursor = page.cursor;
  } while (cursor);

  return out;
}

/** Liste détaillée des ITEMS Monday correspondant aux FORMULAIRES lead
 *  (Source ∈ SITE_MONDAY_LEADS_SOURCE_MATCH, Type ∈ SITE_MONDAY_LEADS_FORM_MATCH, Statut ∈ SITE_MONDAY_LEADS_STATUS_MATCH)
 */
function SITE_mondayFetchLeadFormsItems_(boardId, colSource, colType, colStatus, fromUTC, toUTC) {
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId = SITE_mondayResolveColumnId_(boardId, colType);
  const statusId = SITE_mondayResolveColumnId_(boardId, colStatus);
  if (!sourceId || !typeId || !statusId) throw new Error('Colonnes Monday (source/type/statut) introuvables.');

  const out = []; let cursor = null;
  do {
    const q = `
      query($bid:[ID!], $cursor:String, $cols:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id name state created_at
              column_values(ids:$cols){ id text }
            }
          }
        }
      }`;
    const d = Utils_mondayGraphQL(q, { bid: [Number(boardId)], cursor, cols: [sourceId, typeId, statusId] });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;

    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;
      const dt = new Date(it.created_at); if (isNaN(dt.getTime())) return;
      if (dt < fromUTC || dt > toUTC) return;

      const cv = it.column_values || [];
      const srcText = (cv.find(c => c.id === sourceId) || {}).text || '';
      const typText = (cv.find(c => c.id === typeId) || {}).text || '';
      const stText = (cv.find(c => c.id === statusId) || {}).text || '';
      if (!Utils_textIncludesAny(srcText, SITE_MONDAY_LEADS_SOURCE_MATCH)) return;
      if (!Utils_textIncludesAny(typText, SITE_MONDAY_LEADS_FORM_MATCH)) return;
      if (!Utils_textIncludesAny(stText, SITE_MONDAY_LEADS_STATUS_MATCH)) return;

      out.push({ id: it.id, name: it.name, created_at: dt.toISOString(), source: srcText, type: typText, status: stText });
    });

    cursor = page.cursor;
  } while (cursor);

  return out;
}

/** Runner de TEST : logs pour les 3 derniers mois complets */
function run_Site_TestLeads() {
  const today = new Date();
  const endLocal = new Date(today.getFullYear(), today.getMonth(), 0); // dernier jour du mois précédent
  const startLocal = new Date(endLocal.getFullYear(), endLocal.getMonth() - 2, 1); // 1er jour il y a 2 mois
  const fromUTC = new Date(Date.UTC(startLocal.getFullYear(), startLocal.getMonth(), 1, 0, 0, 0));
  const toUTC = new Date(Date.UTC(endLocal.getFullYear(), endLocal.getMonth(), endLocal.getDate(), 23, 59, 59));

  Logger.log('=== TEST LEADS — Fenêtre UTC ===');
  Logger.log('From: %s  To: %s', fromUTC.toISOString(), toUTC.toISOString());

  const callsByMonth = SITE_mondayLeadCallsCountsByMonth_(
    SITE_MONDAY_LEADS_BOARD_ID,
    SITE_MONDAY_LEADS_COL_SOURCE,
    SITE_MONDAY_LEADS_COL_TYPE,
    SITE_MONDAY_LEADS_COL_STATUS,
    fromUTC, toUTC
  );
  const formsByMonth = SITE_mondayLeadFormsCountsByMonth_(
    SITE_MONDAY_LEADS_BOARD_ID,
    SITE_MONDAY_LEADS_COL_SOURCE,
    SITE_MONDAY_LEADS_COL_TYPE,
    SITE_MONDAY_LEADS_COL_STATUS,
    fromUTC, toUTC
  );

  Logger.log('--- Appels lead (par mois) ---');
  Object.keys(callsByMonth).sort().forEach(k => Logger.log('%s : %s', k, callsByMonth[k]));
  Logger.log('--- Formulaires lead (par mois) ---');
  Object.keys(formsByMonth).sort().forEach(k => Logger.log('%s : %s', k, formsByMonth[k]));

  const callItems = SITE_mondayFetchLeadCallsItems_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_SOURCE, SITE_MONDAY_LEADS_COL_TYPE, fromUTC, toUTC);
  const formItems = SITE_mondayFetchLeadFormsItems_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_SOURCE, SITE_MONDAY_LEADS_COL_TYPE, SITE_MONDAY_LEADS_COL_STATUS, fromUTC, toUTC);

  Logger.log('=== Détails — APPELS lead (%s items) ===', callItems.length);
  callItems.forEach(it => {
    Logger.log('[Call] %s | %s | %s | src="%s" | type="%s"', it.id, it.created_at, it.name, it.source, it.type);
  });

  Logger.log('=== Détails — FORMULAIRES lead (%s items) ===', formItems.length);
  formItems.forEach(it => {
    Logger.log('[Form] %s | %s | %s | src="%s" | type="%s" | status="%s"', it.id, it.created_at, it.name, it.source, it.type, it.status);
  });

  Logger.log('=== FIN TEST LEADS ===');
}

// Parse JSON de colonne Date Monday: {"date":"YYYY-MM-DD", ...} -> Date UTC 00:00
function SITE_parseMondayDateValue_(val) {
  if (!val) return null;
  try {
    const j = JSON.parse(val);
    if (j && j.date) {
      const [Y, M, D] = String(j.date).split('-').map(Number);
      const d = new Date(Date.UTC(Y, (M || 1) - 1, D || 1, 0, 0, 0));
      return isNaN(d.getTime()) ? null : d;
    }
  } catch (e) { }
  return null;
}

// Cherche text & value d'une colonne par id dans column_values
function SITE_cv_(cv, id) {
  const c = (cv || []).find(x => x.id === id);
  return {
    text: (c && c.text) || '',
    value: (c && c.value) || ''
  };
}

// bool includes any (case/accents insensitive)
function SITE_includesAny_(txt, arr) { return Utils_textIncludesAny(txt, arr); }

// DEBUG: inspecte items et explique pourquoi inclus/exclus
function run_Site_TestLeads_Debug() {
  const today = new Date();
  const endLocal = new Date(today.getFullYear(), today.getMonth(), 0);
  const startLocal = new Date(endLocal.getFullYear(), endLocal.getMonth() - 2, 1);
  const fromUTC = new Date(Date.UTC(startLocal.getFullYear(), startLocal.getMonth(), 1, 0, 0, 0));
  const toUTC = new Date(Date.UTC(endLocal.getFullYear(), endLocal.getMonth(), endLocal.getDate(), 23, 59, 59));

  Logger.log('=== DEBUG LEADS — Fenêtre ===');
  Logger.log('From: %s  To: %s', fromUTC.toISOString(), toUTC.toISOString());

  const sourceId = SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_SOURCE);
  const typeId = SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_TYPE);
  const statusId = SITE_MONDAY_LEADS_COL_STATUS ? SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_STATUS) : null;
  const dateId = SITE_MONDAY_LEADS_COL_DATE ? SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_DATE) : null;

  if (!sourceId || !typeId) throw new Error('Colonnes source/type introuvables (vérifie les intitulés).');

  let cursor = null, items = [];
  do {
    const q = `
      query($bid:[ID!], $cursor:String, $cols:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id name state created_at
              column_values(ids:$cols){ id text value }
            }
          }
        }
      }`;
    const colIds = [sourceId, typeId].concat(statusId ? [statusId] : []).concat(dateId ? [dateId] : []);
    const d = Utils_mondayGraphQL(q, { bid: [Number(SITE_MONDAY_LEADS_BOARD_ID)], cursor, cols: colIds });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page; if (!page) break;

    (page.items || []).forEach(it => {
      const createdAt = new Date(it.created_at);
      const cv = it.column_values || [];
      const src = SITE_cv_(cv, sourceId);
      const typ = SITE_cv_(cv, typeId);
      const sts = statusId ? SITE_cv_(cv, statusId) : { text: '', value: '' };
      const dcv = dateId ? SITE_cv_(cv, dateId) : { text: '', value: '' };

      const dateFromCol = SITE_parseMondayDateValue_(dcv.value);
      const when = dateFromCol || createdAt;

      const inWindow = (when >= fromUTC && when <= toUTC);
      const srcOK = SITE_includesAny_(src.text, SITE_MONDAY_LEADS_SOURCE_MATCH);
      const typeIsCall = SITE_includesAny_(typ.text, SITE_MONDAY_LEADS_TYPE_MATCH);
      const typeIsForm = SITE_includesAny_(typ.text, SITE_MONDAY_LEADS_FORM_MATCH || []);
      const statusOK = !statusId || SITE_includesAny_(sts.text, SITE_MONDAY_LEADS_STATUS_MATCH || []);

      const isLeadCall = inWindow && srcOK && typeIsCall;
      const isLeadForm = inWindow && srcOK && typeIsForm && statusOK;

      const reasons = [];
      if (!inWindow) reasons.push('hors fenêtre');
      if (!srcOK) reasons.push('source !match');
      if (!(typeIsCall || typeIsForm)) reasons.push('type !match');
      if (typeIsForm && !statusOK) reasons.push('statut !match');

      items.push({
        id: it.id,
        name: it.name,
        created_at: isNaN(createdAt.getTime()) ? '' : createdAt.toISOString(),
        date_col: dateFromCol ? dateFromCol.toISOString() : '',
        when_used: when.toISOString(),
        src: src.text, typ: typ.text, sts: sts.text,
        isLeadCall, isLeadForm,
        exclude: reasons.join(', ')
      });
    });

    cursor = page.cursor;
  } while (cursor);

  const aggCalls = {}, aggForms = {};
  items.forEach(x => {
    const ym = x.when_used.slice(0, 7);
    if (x.isLeadCall) aggCalls[ym] = (aggCalls[ym] || 0) + 1;
    if (x.isLeadForm) aggForms[ym] = (aggForms[ym] || 0) + 1;
  });

  Logger.log('--- Totaux APPELS lead par mois ---');
  Object.keys(aggCalls).sort().forEach(k => Logger.log('%s : %s', k, aggCalls[k]));
  Logger.log('--- Totaux FORMULAIRES lead par mois ---');
  Object.keys(aggForms).sort().forEach(k => Logger.log('%s : %s', k, aggForms[k]));

  Logger.log('=== Détails (dans la fenêtre) — candidats APPELS ===');
  items.filter(x => x.when_used >= fromUTC.toISOString() && x.when_used <= toUTC.toISOString() && (x.typ || '').length)
    .filter(x => SITE_includesAny_(x.typ, SITE_MONDAY_LEADS_TYPE_MATCH))
    .forEach(x => Logger.log('[CALL?] %s | %s | when=%s | src="%s" type="%s" status="%s" | lead=%s | excl=%s',
      x.id, x.name, x.when_used, x.src, x.typ, x.sts, x.isLeadCall, x.exclude));

  Logger.log('=== Détails (dans la fenêtre) — candidats FORMULAIRES ===');
  items.filter(x => x.when_used >= fromUTC.toISOString() && x.when_used <= toUTC.toISOString() && (x.typ || '').length)
    .filter(x => SITE_includesAny_(x.typ, SITE_MONDAY_LEADS_FORM_MATCH || []))
    .forEach(x => Logger.log('[FORM?] %s | %s | when=%s | src="%s" type="%s" status="%s" | lead=%s | excl=%s',
      x.id, x.name, x.when_used, x.src, x.typ, x.sts, x.isLeadForm, x.exclude));

  Logger.log('=== FIN DEBUG LEADS ===');
}

function run_Site_CurrentMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SITE);
  if (!sh) throw new Error(`Onglet '${SHEET_NAME_SITE}' introuvable`);
  const cols = SITE_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);              // 1er jour du mois en cours
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate()); // aujourd'hui
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');
  const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const matomoCutoffDate = new Date(2025, 10, 1); // 1er Novembre 2025
  const useMatomo = start >= matomoCutoffDate;

  Utils_toast(`Site: mise à jour du mois en cours ${Utils_monthKeyToFr(ymKey)}…`, 'Site Internet', 4);

  let sessionsData;
  if (useMatomo) {
    const rawMatomo = Utils_executeWithRetry(() => SITE_matomoFetchVisitsMonthly_(startStr, endStr), 'Matomo mois en cours');
    sessionsData = SITE_parseMatomoVisitsResponse_(rawMatomo);
  } else {
    sessionsData = Utils_executeWithRetry(() => ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr), 'GA4 mois en cours');
  }

  // Pour la durée moyenne et le taux de rebond, on continue d'utiliser GA4 pour l'instant
  const gaForOtherMetrics = Utils_executeWithRetry(() => ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr), 'GA4 mois en cours (metrics)');


  const gsc = Utils_executeWithRetry(() => gscFetchMonthly_(GSC_SITE_URL, startStr, endStr), 'GSC mois en cours');

  // Appels Magnetis (filtrés "Tout le SEO")
  const callsByMonth = Utils_executeWithRetry(
    () => SITE_magnetisAggregateMonthlyCounts_(SITE_magnetisFetchCalls_(start, end)),
    'Magnetis mois en cours'
  );

  // Formulaires (Paperform + Monday)
  const formsByMonth = Utils_executeWithRetry(
    () => SITE_formsCountsByMonth_(
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
    ),
    'Forms mois en cours'
  );

  // Leads Monday : appels + formulaires
  const leadCallsByMonth = Utils_executeWithRetry(
    () => SITE_mondayLeadCallsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      SITE_MONDAY_LEADS_COL_STATUS,  // ⚠️ FIX: Paramètre manquant !
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
    ),
    'Monday lead calls mois en cours'
  );

  const leadFormsByMonth = Utils_executeWithRetry(
    () => SITE_mondayLeadFormsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      SITE_MONDAY_LEADS_COL_STATUS,
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
    ),
    'Monday lead forms mois en cours'
  );

  const t1 = new Date();
  Logger.log(`[PERF] Début construction values: ${t1.toISOString()}`);

  const values = {
    sessions: sessionsData[ymKey]?.sessions ?? null,
    // Si Matomo (>= nov 2025), utiliser Matomo pour avgSec et bounce, sinon GA4
    avgSec: useMatomo ? (sessionsData[ymKey]?.avgSec ?? null) : (gaForOtherMetrics[ymKey]?.avgSec ?? null),
    impressions: gsc[ymKey]?.impressions ?? null,
    calls: callsByMonth[ymKey] ?? null,
    forms: formsByMonth[ymKey] ?? null,
    leadCalls: leadCallsByMonth[ymKey] ?? null,
    leadForms: leadFormsByMonth[ymKey] ?? null,
    bounce: useMatomo ? (sessionsData[ymKey]?.bouncePct ?? null) : (gaForOtherMetrics[ymKey]?.bouncePct ?? null)
  };

  const t2 = new Date();
  Logger.log(`[PERF] Fin construction values (${t2 - t1}ms). Début SITE_writeMonthRow_: ${t2.toISOString()}`);

  // ⚠️ SITE_writeMonthRow_ gère déjà :
  //  - création / insertion de la ligne du mois si elle n'existe pas
  //  - écriture UNIQUEMENT dans les colonnes autorisées (pas les formules)
  SITE_writeMonthRow_(sh, cols, ymKey, values);

  const t3 = new Date();
  Logger.log(`[PERF] Fin SITE_writeMonthRow_ (${t3 - t2}ms)`);

  Utils_toast('Site: mois en cours mis à jour ✅', 'Site Internet', 5);
}

