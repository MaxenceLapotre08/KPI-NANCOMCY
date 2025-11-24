/****************************** CONFIG FENETRE (Full History borné) *********/
const META_FULLHIST_START_LOCAL = new Date(2024, 0, 1);                  // 1 janv 2024 (local)
const META_FULLHIST_START_UTC = new Date(Date.UTC(2024, 0, 1, 0, 0, 0));// 1 janv 2024 (UTC)

/****************************** PERFORMANCE / BATCH **************************/
const META_FULLHIST_BATCH_MONTHS = 4;   // traite 4 mois par exécution
let META_ENABLE_MATOMO = true;        // mettre false pour couper Matomo
let META_ENABLE_LEADS = true;        // mettre false pour couper Monday Leads (appels/formulaires lead)

/****************************** CONFIG FEUILLE ******************************/
const META_SHEET_NAME = 'Meta Ads';
const META_HEADERS_ROW = 3;   // ligne des entêtes
const META_START_ROW = 6;   // première ligne de données

/****************************** CONFIG META MARKETING API *******************/
// ID de compte publicitaire (numérique, sans "act_")
const META_AD_ACCOUNT_ID = '152460256751291';

// Token long-lived avec accès à l’API Marketing (scope ads_read)
function META_getMetaToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('META_TOKEN');
  if (!t) throw new Error("Propriété 'META_TOKEN' manquante.");
  return t;
}
// Version Graph API
const META_GRAPH_VERSION = 'v20.0';

/****************************** CONFIG MAGNETIS (appels) ********************/
function META_getMagApiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('MAGNETIS_API_KEY');
  if (!k) throw new Error("Propriété 'MAGNETIS_API_KEY' manquante.");
  return k;
}
// Canaux qui identifient Meta/Facebook/Instagram (à adapter à ton mapping)
const META_MAG_CHANNELS = [
  'facebook', 'facebook / cpc', 'facebook/cpc',
  'instagram', 'instagram ads',
  'meta', 'meta ads', 'social', 'social paid'
];
const META_MAG_MATCH = 'includes'; // 'equals' | 'includes' | 'regex'
const META_MAG_ANSWERED_ONLY = false; // true => compte uniquement décroches (>0s)

/****************************** CONFIG PAPERFORM (formulaires) *************/
function META_getPaperformToken_() {
  const props = PropertiesService.getScriptProperties();
  const t = props.getProperty('PAPERFORM_TOKEN')
    || props.getProperty('PAPARFORM_TOKEN')
    || null;
  if (!t) throw new Error("Propriété 'PAPERFORM_TOKEN' manquante.");
  return t;
}
// Si tu utilises un ou plusieurs Paperform dédiés à Meta, liste-les ici (sinon, laisse vide [])
const META_PAPERFORM_FORM_IDS = []; // ex: ['abcd1234']

/****************************** CONFIG MONDAY (formulaires) *****************/
// Board mixte + colonne dédiée aux formulaires “META ads”
const META_MONDAY_BOARD_ID = 9950271520;
const META_MONDAY_FORM_COLUMN_TITLE = 'Meta ads'; // le titre dans Monday
const META_MONDAY_FORM_COLUMN_ID = '';            // si tu connais l’ID direct
const META_MONDAY_MATCH_MODE = 'nonempty';        // 'nonempty' | 'equals' | 'includes'
const META_MONDAY_MATCH_VALUES = [];              // utile si equals/includes
function META_getMondayToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('MONDAY_TOKEN');
  if (!t) throw new Error("Propriété 'MONDAY_TOKEN' manquante.");
  return t;
}

/************* CONFIG MONDAY (Leads Meta : appels/formulaires) *************/
const META_LEADS_BOARD_ID = META_MONDAY_BOARD_ID;  // même board
const META_LEADS_COL_SOURCE = 'Source du Lead';
const META_LEADS_COL_TYPE = 'FORMULAIRE / APPELS';
const META_LEADS_COL_STATUS = 'Nature du contact';

// Valeur de source attendue (insensible casse/accents)
const META_LEADS_SOURCE_EQUALS = 'LP (via Meta Ads)';
// Statut Lead
const META_LEADS_STATUS_MATCH = ['lead'];
// Reconnaissance type
const META_LEADS_TYPE_CALL_MATCH = ['Appel', 'call', 'téléphone'];
const META_LEADS_TYPE_FORM_MATCH = ['Formulaire', 'form', 'paperform'];

/****************************** CONFIG MATOMO (durée visite) ****************/
const META_MATOMO_BASE_URL = 'https://matomo.aleo.agency';
const META_MATOMO_SITE_ID = 1492;
const META_MATOMO_PAGE_PATTERN = 'lp-ma-'; // motif URL “contient”
function META_getMatomoToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('MATOMO_TOKEN');
  if (!t) throw new Error("Propriété 'MATOMO_TOKEN' manquante.");
  return t;
}

/****************************** HELPERS GÉNÉRAUX (Utilise Utils.js) *****************/
// Les fonctions suivantes sont maintenant dans Utils.js :
// Utils_toast -> Utils_toast
// Utils_executeWithRetry -> Utils_executeWithRetry
// Utils_normHeader -> Utils_normHeader
// Utils_findColByHeaderAliases -> Utils_findColByHeaderAliases
// Utils_isYearSeparatorRow -> Utils_isYearSeparatorRow
// Utils_sheetCellToYYYYMM -> Utils_sheetCellToYYYYMM
// Utils_monthKeyToFr -> Utils_monthKeyToFr
// Utils_setSecondsAsDuration -> Utils_setSecondsAsDuration

/************ Lignes présentes / recherche de ligne (SANS insertion) ********/
function META_findExistingMonthRow_(sh, moisCol, targetYM) {
  const last = sh.getLastRow();
  if (last < META_START_ROW) return 0;
  for (let r = META_START_ROW; r <= last; r++) {
    const v = sh.getRange(r, moisCol).getValue();
    if (Utils_isYearSeparatorRow(v)) continue;
    const ym = Utils_sheetCellToYYYYMM(v);
    if (ym === targetYM) return r;
  }
  return 0;
}

/************ Insertion propre d’une ligne année + mois *********************/
function META__styleYearRow_(sh, row, moisCol) {
  sh.getRange(row, 1, 1, sh.getLastColumn()).setBackground('#e6e1f5');
  sh.getRange(row, moisCol).setFontWeight('bold');
}

/**
 * Garantit qu’une ligne existe pour le mois targetYM
 * - Crée la ligne année si nécessaire (2024, 2025…)
 * - Ajoute le mois à la bonne place dans l’année
 * - Ne touche pas aux autres colonnes (les formules existantes restent)
 */
function META_ensureMonthRow_(sh, moisCol, targetYM) {
  const lastRow = sh.getLastRow();

  // 1. Chercher si la ligne du mois existe déjà
  for (let r = META_START_ROW; r <= lastRow; r++) {
    const v = sh.getRange(r, moisCol).getValue();
    if (Utils_isYearSeparatorRow(v)) continue;
    const ym = Utils_sheetCellToYYYYMM(v);
    if (ym === targetYM) {
      Logger.log(`[Meta/INFO] Ligne pour ${targetYM} trouvée en ${r}. Utilisation de la ligne existante.`);
      return r;
    }
  }

  // 2. Si elle n'existe pas, trouver la dernière ligne de données en se basant sur la colonne Mois
  let lastDataRow = META_START_ROW - 1;
  for (let r = lastRow; r >= META_START_ROW; r--) {
    const v = sh.getRange(r, moisCol).getValue();
    if (v !== '') {
      lastDataRow = r;
      break;
    }
  }

  const targetRow = lastDataRow + 1;
  const lastDataValue = lastDataRow >= META_START_ROW ? sh.getRange(lastDataRow, moisCol).getValue() : '';
  const lastDataYM = Utils_sheetCellToYYYYMM(lastDataValue) || '1999-12';

  // 3. Gérer le cas particulier du changement d'année
  const lastYear = parseInt(lastDataYM.slice(0, 4), 10);
  const targetYear = parseInt(targetYM.slice(0, 4), 10);

  if (targetYear > lastYear) {
    Logger.log(`[Meta/INFO] Changement d'année détecté (${lastYear} -> ${targetYear}). Ajout d'un séparateur.`);
    sh.getRange(targetRow, moisCol).setValue(String(targetYear));
    META__styleYearRow_(sh, targetRow, moisCol);
    Logger.log(`[Meta/WRITE] Écriture des KPIs pour ${targetYM} sur la nouvelle ligne ${targetRow + 1}`);
    return targetRow + 1;
  }

  // 4. Cas normal : on écrit sur la ligne juste après la dernière
  Logger.log(`[Meta/WRITE] Écriture des KPIs pour ${targetYM} sur la nouvelle ligne ${targetRow}`);
  return targetRow;
}

/****************************** RANGE DE MOIS (bornage strict) **************/
function META_monthRangeKeys_(fromLocal, toLocal) {
  const keys = [];
  const d = new Date(fromLocal.getFullYear(), fromLocal.getMonth(), 1);
  const end = new Date(toLocal.getFullYear(), toLocal.getMonth(), 1);
  while (d <= end) {
    const ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    keys.push(ym);
    d.setMonth(d.getMonth() + 1);
  }
  return keys;
}

/******************************** META INSIGHTS ********************************/
// Retourne { 'YYYY-MM': { spend, impressions, clicks, interactions } }
function META_fetchInsightsMonthly_(fromDate, toDate) {
  if (!META_AD_ACCOUNT_ID || !/^\d+$/.test(META_AD_ACCOUNT_ID)) {
    throw new Error("META_AD_ACCOUNT_ID invalide (mets l'ID numérique sans 'act_').");
  }

  const token = META_getMetaToken_();
  const base = `https://graph.facebook.com/${META_GRAPH_VERSION}/act_${META_AD_ACCOUNT_ID}/insights`;

  const since = Utilities.formatDate(fromDate, 'UTC', 'yyyy-MM-dd');
  const until = Utilities.formatDate(toDate, 'UTC', 'yyyy-MM-dd');

  const fields = [
    'date_start', 'date_stop',
    'spend', 'impressions', 'clicks', 'inline_link_clicks',
    'actions'
  ].join(',');

  const timeRange = { since, until };
  let url = `${base}?level=account&time_increment=monthly&fields=${encodeURIComponent(fields)}&time_range=${encodeURIComponent(JSON.stringify(timeRange))}&limit=100`;

  const out = {};
  Logger.log(`[Meta/insights] Appel API ${since} → ${until}`);

  while (url) {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + token }
    });

    if (res.getResponseCode() >= 300) {
      throw new Error(`Meta insights HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    }

    const body = JSON.parse(res.getContentText());

    (body.data || []).forEach(row => {
      const ym = String(row.date_start || '').slice(0, 7);
      if (!ym) return;

      const spend = Number(row.spend || 0);
      const impressions = Number(row.impressions || 0);
      const clicks = Number(row.inline_link_clicks != null ? row.inline_link_clicks : (row.clicks || 0));
      let interactions = 0;

      const wanted = new Set([
        'post_engagement', 'page_engagement', 'like',
        'post_reaction', 'comment', 'post_comment', 'link_click'
      ]);

      (row.actions || []).forEach(a => {
        const t = String(a.action_type || '').toLowerCase();
        if (wanted.has(t)) interactions += Number(a.value || 0);
      });

      if (!out[ym]) out[ym] = { spend: 0, impressions: 0, clicks: 0, interactions: 0 };

      out[ym].spend += spend;
      out[ym].impressions += impressions;
      out[ym].clicks += clicks;
      out[ym].interactions += interactions;
    });

    url = (body.paging && body.paging.next) ? body.paging.next : null;
  }

  Logger.log(`[Meta/insights] Résumé = ${JSON.stringify(out)}`);
  return out;
}

/******************************** MAGNETIS (appels) **************************/


function META_getChannelName_(c) {
  const cand = ['channel_name', 'channel', 'analysis.channel', 'analytics.channel', 'utm.channel', 'session.channel'];
  for (const p of cand) {
    const v = Utils_valAt(c, p);
    if (v != null && String(v).trim() !== '') return String(v).trim().toLowerCase();
  }
  return '';
}

function META_channelMatch_(name) {
  if (!name) return false;
  switch (META_MAG_MATCH) {
    case 'equals': return META_MAG_CHANNELS.some(a => name === a);
    case 'regex': return META_MAG_CHANNELS.some(rx => new RegExp(rx, 'i').test(name));
    default: return META_MAG_CHANNELS.some(a => name.includes(a));
  }
}

function META_magnetisFetchCalls_(fromDate, toDate) {
  const apiKey = META_getMagApiKey_();
  const base = 'https://api.magnetis.io/calls';
  const fmtUTC = d => Utilities.formatDate(d, 'UTC', 'yyyyMMddHHmmss');
  const fromStr = fmtUTC(fromDate), toStr = fmtUTC(toDate);
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
  Logger.log(`[Magnetis] nb calls bruts=${all.length}`);
  return all;
}

// { 'YYYY-MM': n }
function META_magnetisMonthlyCalls_(fromDate, toDate) {
  const bucket = {};
  META_magnetisFetchCalls_(fromDate, toDate).forEach(c => {
    const ch = META_getChannelName_(c);
    if (!META_channelMatch_(ch)) return;
    const raw = c.start_at || c.created_at || c.date || c.started_at;
    if (!raw) return;
    const d = new Date(raw); if (isNaN(d.getTime())) return;
    if (META_MAG_ANSWERED_ONLY) {
      const dur = Number(c.duration || Utils_valAt(c, 'analysis.duration') || 0);
      if (!(dur > 0)) return;
    }
    const ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    bucket[ym] = (bucket[ym] || 0) + 1;
  });
  Logger.log(`[Magnetis/monthly] ${JSON.stringify(bucket)}`);
  return bucket;
}

/******************************** Paperform (form Meta) *********************/
function META_extractPaperformArray_(json) {
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

function META_getSubmissionDate_(s) {
  const cand = [
    s.submitted_at, s.submittedAt, s.submitted_at_utc, s.submittedAtUTC,
    s.created_at, s.createdAt, s.created_at_utc, s.createdAtUTC,
    s.date, s.timestamp, s.time_submitted, s.timeSubmitted,
    s.created, s.submitted,
    (s.meta && (s.meta.submitted_at || s.meta.created_at))
  ];
  for (let v of cand) {
    if (v == null) continue;
    if (typeof v === 'number') {
      const ms = v < 1e12 ? v * 1000 : v;
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d;
      continue;
    }
    if (typeof v === 'string') {
      const d1 = new Date(v);
      if (!isNaN(d1.getTime())) return d1;
      if (/^\d+$/.test(v)) {
        const n = +v, ms = n < 1e12 ? n * 1000 : n;
        const d2 = new Date(ms);
        if (!isNaN(d2.getTime())) return d2;
      }
    }
  }
  return new Date(NaN);
}

function META_paperformFetchInRange_(slugOrId, fromDate, toDate) {
  const token = META_getPaperformToken_();
  const base = `https://api.paperform.co/v1/forms/${encodeURIComponent(slugOrId)}/submissions`;
  const limit = 100; let all = []; let mode = 'skip', skip = 0, page = 1;
  const fromTs = fromDate.getTime(), toTs = toDate.getTime();
  while (true) {
    const url = (mode === 'skip') ? `${base}?limit=${limit}&skip=${skip}` : `${base}?limit=${limit}&page=${page}`;
    const res = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' }, muteHttpExceptions: true });
    if (res.getResponseCode() === 422 && mode === 'skip') { mode = 'page'; page = 1; continue; }
    if (res.getResponseCode() >= 300) throw new Error(`Paperform HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    const batch = META_extractPaperformArray_(JSON.parse(res.getContentText()));
    if (!batch.length) break;
    for (const s of batch) {
      const d = META_getSubmissionDate_(s);
      if (!isNaN(d.getTime())) {
        const t = d.getTime();
        if (t >= fromTs && t <= toTs) all.push(s);
      }
    }
    if (mode === 'skip') skip += batch.length; else page += 1;
    if (batch.length < limit) break;
    Utilities.sleep(120);
  }
  return all;
}

// { 'YYYY-MM': n } (somme de tous les IDs configurés)
function META_paperformCountsByMonth_(fromDateUTC, toDateUTC) {
  const bucket = {};
  META_PAPERFORM_FORM_IDS.forEach(fid => {
    META_paperformFetchInRange_(fid, fromDateUTC, toDateUTC).forEach(s => {
      const d = META_getSubmissionDate_(s); if (isNaN(d.getTime())) return;
      const ym = d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
      bucket[ym] = (bucket[ym] || 0) + 1;
    });
  });
  Logger.log(`[Paperform/monthly] ${JSON.stringify(bucket)}`);
  return bucket;
}

/******************************** Monday (form Meta) ************************/
function META__norm_(s) {
  return String(s || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’'`]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function META_mondayGraphQL_(query, variables) {
  const res = UrlFetchApp.fetch('https://api.monday.com/v2', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': META_getMondayToken_(),
      'API-Version': '2023-10'
    },
    payload: JSON.stringify({ query, variables: variables || {} }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error(`Monday HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
  const body = JSON.parse(res.getContentText());
  if (body.errors) throw new Error('Monday GraphQL error: ' + JSON.stringify(body.errors));
  return body.data;
}

function META_mondayResolveColId_(boardId, titleOrId) {
  if (!titleOrId) return '';
  const looksId = /^[a-z0-9_]+$/i.test(titleOrId) && !META__norm_(titleOrId).includes(' ');
  if (looksId) return titleOrId;
  const want = META__norm_(titleOrId);
  const q = `query($bid:[ID!]){ boards(ids:$bid){ columns{ id title type } } }`;
  const d = META_mondayGraphQL_(q, { bid: [Number(boardId)] });
  const cols = (d && d.boards && d.boards[0] && d.boards[0].columns) || [];
  let best = '';
  for (const c of cols) { if (META__norm_(c.title) === want) { best = c.id; break; } }
  if (!best) {
    for (const c of cols) { if (META__norm_(c.title).includes(want)) { best = c.id; break; } }
  }
  return best;
}

function META_mondayValueOk_(txt) {
  const t = META__norm_(txt || '');
  switch (META_MONDAY_MATCH_MODE) {
    case 'equals': return META_MONDAY_MATCH_VALUES.some(v => t === META__norm_(v));
    case 'includes': return META_MONDAY_MATCH_VALUES.some(v => t.includes(META__norm_(v)));
    default: return t.length > 0; // nonempty
  }
}

// { 'YYYY-MM': n }
function META_mondayFormCountsByMonth_(fromUTC, toUTC) {
  const colId = META_MONDAY_FORM_COLUMN_ID ? META_MONDAY_FORM_COLUMN_ID : META_mondayResolveColId_(META_MONDAY_BOARD_ID, META_MONDAY_FORM_COLUMN_TITLE);
  if (!colId) throw new Error("Colonne Monday 'META ads' introuvable.");

  const counts = {}; let cursor = null;
  do {
    const q = `query($bid:[ID!], $cursor:String, $cols:[String!]){ 
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
    const d = META_mondayGraphQL_(q, {
      bid: [Number(META_MONDAY_BOARD_ID)],
      cursor,
      cols: [colId]
    });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;
    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;
      const dt = new Date(it.created_at); if (isNaN(dt.getTime())) return;
      if (dt < fromUTC || dt > toUTC) return;
      const txt = (it.column_values && it.column_values[0] && it.column_values[0].text) || '';
      if (!META_mondayValueOk_(txt)) return;
      const ym = dt.getUTCFullYear() + '-' + String(dt.getUTCMonth() + 1).padStart(2, '0');
      counts[ym] = (counts[ym] || 0) + 1;
    });
    cursor = page.cursor;
  } while (cursor);
  Logger.log(`[Monday forms/monthly] ${JSON.stringify(counts)}`);
  return counts;
}

/**************** Monday Leads (Meta) → appels & formulaires lead ***********/



// { 'YYYY-MM': n } — APPELS lead
function META_mondayLeadCallsByMonth_(boardId, colSource, colType, colStatus, fromUTC, toUTC) {
  const sourceId = META_mondayResolveColId_(boardId, colSource);
  const typeId = META_mondayResolveColId_(boardId, colType);
  const statusId = META_mondayResolveColId_(boardId, colStatus);
  if (!sourceId || !typeId || !statusId) throw new Error('[Monday leads calls] Colonnes introuvables (source/type/status).');

  const counts = {}; let cursor = null;
  do {
    const q = `query($bid:[ID!], $cursor:String, $cols:[String!]){ boards(ids:$bid){ items_page(limit:500,cursor:$cursor){ cursor items{ id state created_at column_values(ids:$cols){ id text } } } } }`;
    const d = META_mondayGraphQL_(q, { bid: [Number(boardId)], cursor, cols: [sourceId, typeId, statusId] });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;
    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;
      const when = new Date(it.created_at); if (isNaN(when.getTime())) return;
      if (when < fromUTC || when > toUTC) return;
      const cv = it.column_values || [];
      const src = (cv.find(c => c.id === sourceId) || {}).text || '';
      const typ = (cv.find(c => c.id === typeId) || {}).text || '';
      const sts = (cv.find(c => c.id === statusId) || {}).text || '';
      if (!Utils_textEquals(src, META_LEADS_SOURCE_EQUALS)) return;
      if (!Utils_textIncludesAny(sts, META_LEADS_STATUS_MATCH)) return;
      if (!Utils_textIncludesAny(typ, META_LEADS_TYPE_CALL_MATCH)) return;
      const key = when.getUTCFullYear() + '-' + String(when.getUTCMonth() + 1).padStart(2, '0');
      counts[key] = (counts[key] || 0) + 1;
    });
    cursor = page.cursor;
  } while (cursor);
  Logger.log(`[Monday leads calls/monthly] ${JSON.stringify(counts)}`);
  return counts;
}

// { 'YYYY-MM': n } — FORMULAIRES lead
function META_mondayLeadFormsByMonth_(boardId, colSource, colType, colStatus, fromUTC, toUTC) {
  const sourceId = META_mondayResolveColId_(boardId, colSource);
  const typeId = META_mondayResolveColId_(boardId, colType);
  const statusId = META_mondayResolveColId_(boardId, colStatus);
  if (!sourceId || !typeId || !statusId) throw new Error('[Monday leads forms] Colonnes introuvables (source/type/status).');

  const counts = {}; let cursor = null;
  do {
    const q = `query($bid:[ID!], $cursor:String, $cols:[String!]){ boards(ids:$bid){ items_page(limit:500,cursor:$cursor){ cursor items{ id state created_at column_values(ids:$cols){ id text } } } } }`;
    const d = META_mondayGraphQL_(q, { bid: [Number(boardId)], cursor, cols: [sourceId, typeId, statusId] });
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;
    (page.items || []).forEach(it => {
      if (String(it.state || '').toLowerCase() === 'archived') return;
      const when = new Date(it.created_at); if (isNaN(when.getTime())) return;
      if (when < fromUTC || when > toUTC) return;
      const cv = it.column_values || [];
      const src = (cv.find(c => c.id === sourceId) || {}).text || '';
      const typ = (cv.find(c => c.id === typeId) || {}).text || '';
      const sts = (cv.find(c => c.id === statusId) || {}).text || '';
      if (!Utils_textEquals(src, META_LEADS_SOURCE_EQUALS)) return;
      if (!Utils_textIncludesAny(sts, META_LEADS_STATUS_MATCH)) return;
      if (!Utils_textIncludesAny(typ, META_LEADS_TYPE_FORM_MATCH)) return;
      const key = when.getUTCFullYear() + '-' + String(when.getUTCMonth() + 1).padStart(2, '0');
      counts[key] = (counts[key] || 0) + 1;
    });
    cursor = page.cursor;
  } while (cursor);
  Logger.log(`[Monday leads forms/monthly] ${JSON.stringify(counts)}`);
  return counts;
}

/********************* FUSION FORMULAIRES (Paperform+Monday) ****************/
function META_formsCountsByMonth_(fromUTC, toUTC) {
  const a = META_paperformCountsByMonth_(fromUTC, toUTC);
  const b = META_mondayFormCountsByMonth_(fromUTC, toUTC);
  const keys = Array.from(new Set(Object.keys(a).concat(Object.keys(b))));
  const out = {};
  keys.forEach(k => out[k] = (a[k] || 0) + (b[k] || 0));
  Logger.log(`[Forms (Paper+Monday)/monthly] ${JSON.stringify(out)}`);
  return out;
}

/******************************** MATOMO (durée visite) *********************/
function META_matomoFetch_(params) {
  const token = META_getMatomoToken_();
  const full = { module: 'API', format: 'JSON', token_auth: token, idSite: String(META_MATOMO_SITE_ID), ...params };
  const qs = Object.keys(full).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(full[k])).join('&');
  const url = `${META_MATOMO_BASE_URL}/index.php?${qs}`;
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  if (res.getResponseCode() >= 300) throw new Error(`Matomo HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
  return JSON.parse(res.getContentText());
}

function META_matomoVisitAvgSecs_(ym) {
  const seg = `pageUrl=@${META_MATOMO_PAGE_PATTERN}`;
  const js = META_matomoFetch_({ method: 'VisitsSummary.get', period: 'month', date: ym, hideMetricTranslations: 1, segment: seg });
  if (js && (js.avg_time_on_site != null || (js.sum_visit_length != null && js.nb_visits != null))) {
    const direct = Number(js.avg_time_on_site);
    if (!isNaN(direct)) {
      Logger.log(`[Matomo] ${ym} avgSec=${direct}`);
      return direct;
    }
    const sum = Number(js.sum_visit_length), nb = Number(js.nb_visits);
    if (!isNaN(sum) && nb > 0) {
      const v = sum / nb;
      Logger.log(`[Matomo] ${ym} avgSec=${v}`);
      return v;
    }
  }
  const pages = META_matomoFetch_({
    method: 'Actions.getPageUrls', period: 'month', date: ym, flat: 1, filter_limit: -1,
    filter_column: 'label', filter_pattern: META_MATOMO_PAGE_PATTERN, hideMetricTranslations: 1
  });
  let S = 0, H = 0; (Array.isArray(pages) ? pages : []).forEach(r => {
    const h = Number(r.nb_hits), s = Number(r.sum_time_spent);
    if (!isNaN(h) && h > 0 && !isNaN(s)) { S += s; H += h; }
  });
  const out = H > 0 ? (S / H) : 0;
  Logger.log(`[Matomo] ${ym} avgSec=${out}`);
  return out;
}
// Colonne "Budget investi" SÉCURISÉE pour META
function META_findBudgetColSafe_(sheet) {
  const headerRow = META_HEADERS_ROW || 3;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const HN = headers.map(Utils_normHeader);

  // Colonnes qui ne doivent jamais être prises pour le budget
  const forbiddenRe = /(contact|contacts|lead|leads|signature|signatures|roas|roi)/;

  // 1) Priorité aux matches EXACTS
  const exact = ['budget investi', 'budget', 'depenses', 'dépenses'];
  for (let i = 0; i < HN.length; i++) {
    const h = HN[i];
    if (!h || forbiddenRe.test(h)) continue;
    if (exact.includes(h)) return i + 1;
  }

  // 2) Fallback sur anciens alias, toujours en évitant les colonnes sensibles
  const aliases = ['budget investi', 'budget', 'depenses', 'dépenses', 'spend', 'cost'];
  const wanted = aliases.map(Utils_normHeader);
  for (let i = 0; i < HN.length; i++) {
    const h = HN[i];
    if (!h || forbiddenRe.test(h)) continue;
    if (wanted.some(w => h === w || h.includes(w))) return i + 1;
  }

  Logger.log('[Meta/findBudgetColSafe] Aucune colonne budget trouvée.');
  return 0;
}
// Vrai si la colonne a un header "protégé" (contacts, leads, signatures, ROAS, ROI)
function META_isProtectedHeader_(sh, col) {
  if (!col) return false;
  const headerVal = sh.getRange(META_HEADERS_ROW, col).getValue();
  const h = Utils_normHeader(headerVal);

  // Liste EXACTE des colonnes avec formules pour "Meta Ads"
  // tirée de Sheet_Structures.md
  const protectedHeaders = [
    'cout par contact', 'cout par lead', 'nombre de signature',
    'total montant signe', 'cout par signature', 'montant moyen signature',
    'roas', 'roi', 'nombre de contacts', 'nombre de lead', 'cpm'
  ].map(Utils_normHeader);

  return protectedHeaders.includes(h);
}

/**************************** MAPPING COLONNES ******************************/
function META_findCols_(sh) {
  const cols = {
    // clés principales
    mois: Utils_findColByHeaderAliases(sh, ['mois'], META_HEADERS_ROW),
    budget: META_findBudgetColSafe_(sh),   // ⬅️ ICI, on utilise le helper sécurisé

    // ⚠️ Colonnes existantes à NE PAS TOUCHER (jamais écrites)
    contacts: Utils_findColByHeaderAliases(sh, ['nombre de contacts', 'contacts', 'leads'], META_HEADERS_ROW),
    leadsCount: Utils_findColByHeaderAliases(sh, ['nombre de lead', 'nb lead', 'leads (count)'], META_HEADERS_ROW),
    cpcContact: Utils_findColByHeaderAliases(sh, ['cout par contact', 'coût par contact', 'cpc contact', 'cpl'], META_HEADERS_ROW),
    cplLead: Utils_findColByHeaderAliases(sh, ['cout par lead', 'coût par lead', 'cpl lead'], META_HEADERS_ROW),
    sigNb: Utils_findColByHeaderAliases(sh, ['nombre de signature'], META_HEADERS_ROW),
    sigTot: Utils_findColByHeaderAliases(sh, ['total montant signe', 'total montant signé'], META_HEADERS_ROW),
    sigAvg: Utils_findColByHeaderAliases(sh, ['montant moyen signature'], META_HEADERS_ROW),
    roas: Utils_findColByHeaderAliases(sh, ['roas'], META_HEADERS_ROW),
    roi: Utils_findColByHeaderAliases(sh, ['roi'], META_HEADERS_ROW),

    // métriques sources/agrégées (écriture autorisée)
    impr: Utils_findColByHeaderAliases(sh, ['nombre dimpressions', 'nombre d impressions', 'impressions'], META_HEADERS_ROW),
    clicks: Utils_findColByHeaderAliases(sh, ['nombre de clics', 'clics', 'clicks'], META_HEADERS_ROW),
    ctr: Utils_findColByHeaderAliases(sh, ['ctr', 'taux de clics'], META_HEADERS_ROW),
    conv: Utils_findColByHeaderAliases(sh, ['taux de conversion', 'conversion rate', 'cr'], META_HEADERS_ROW),
    interN: Utils_findColByHeaderAliases(sh, ["nombre d'interaction", "nombre dinteractions", "interactions", "engagements"], META_HEADERS_ROW),
    interR: Utils_findColByHeaderAliases(sh, ['taux dinteraction', "taux d'interaction"], META_HEADERS_ROW),
    appels: Utils_findColByHeaderAliases(sh, ["nombre d'appels", "nombre dappels", "appels"], META_HEADERS_ROW),
    forms: Utils_findColByHeaderAliases(sh, ['nombre de formulaire', 'formulaires', 'forms', 'soumissions'], META_HEADERS_ROW),

    // colonnes leads (Monday) — écriture OK si présentes
    callsLead: Utils_findColByHeaderAliases(sh, ["nombre d'appels lead", "appels lead", "lead appels"], META_HEADERS_ROW),
    formsLead: Utils_findColByHeaderAliases(sh, ["nombre de formulaires lead", "formulaires lead", "lead formulaires"], META_HEADERS_ROW),

    // Matomo
    dur: Utils_findColByHeaderAliases(sh, ['duree moyenne visite', 'durée moyenne visite'], META_HEADERS_ROW),
  };

  Logger.log('[Meta/findCols] Mapping colonnes = ' + JSON.stringify(cols));
  return cols;
}


/**************************** ÉCRITURE D’UNE LIGNE **************************/
function META_writeMonth_(sh, cols, ymKey, v) {
  const row = META_ensureMonthRow_(sh, cols.mois, ymKey);

  // Mois : on n’écrase pas une éventuelle formule
  Utils_setPreserveFormula(sh, row, cols.mois, Utils_monthKeyToFr(ymKey));

  Logger.log(
    `[Meta/write] ${ymKey} → ligne ${row} | ` +
    `spend=${v.spend ?? '∅'}, impr=${v.impressions ?? '∅'}, clicks=${v.clicks ?? '∅'}, inter=${v.interactions ?? '∅'}, ` +
    `calls=${v.calls ?? '∅'}, forms=${v.forms ?? '∅'}, callsLead=${v.callsLead ?? '∅'}, formsLead=${v.formsLead ?? '∅'}, avgSec=${v.avgSec ?? '∅'}`
  );

  // 1. Préparer le plan d'écriture
  const writePlan = {};
  if (v.spend != null) writePlan[cols.budget] = { value: v.spend, format: '0.00 €' };
  if (v.impressions != null) writePlan[cols.impr] = { value: v.impressions, format: null };
  if (v.clicks != null) writePlan[cols.clicks] = { value: v.clicks, format: null };
  if (v.interactions != null) writePlan[cols.interN] = { value: v.interactions, format: null };
  if (v.calls != null) writePlan[cols.appels] = { value: v.calls, format: null };
  if (v.forms != null) writePlan[cols.forms] = { value: v.forms, format: null };
  if (v.callsLead != null) writePlan[cols.callsLead] = { value: v.callsLead, format: null };
  if (v.formsLead != null) writePlan[cols.formsLead] = { value: v.formsLead, format: null };
  if (v.avgSec != null) writePlan[cols.dur] = { value: v.avgSec / 86400, format: '[h]:mm:ss' };

  // 2. Exécuter le plan, cellule par cellule
  for (const col in writePlan) {
    if (!col || col === '0') continue;
    if (META_isProtectedHeader_(sh, col)) {
      Logger.log(`[PROTECTED] Colonne ${sh.getRange(META_HEADERS_ROW, col).getValue()} ignorée car protégée.`);
      continue;
    }
    Utils_setPreserveFormula(sh, row, col, writePlan[col].value, writePlan[col].format);
  }


  // 🔹 Dérivés autorisés : CTR / Taux d’interaction / Taux de conversion
  //    (eux aussi protégés au cas où tu mettrais une formule perso)
  const clicks = Number(sh.getRange(row, cols.clicks).getValue() || 0);
  const impr = Number(sh.getRange(row, cols.impr).getValue() || 0);
  const inters = Number(sh.getRange(row, cols.interN).getValue() || 0);

  if (cols.ctr && !META_isProtectedHeader_(sh, cols.ctr)) {
    const ctr = impr > 0 ? (clicks / impr) : 0;
    Utils_setPreserveFormula(sh, row, cols.ctr, ctr, '0.00%');
  }

  if (cols.interR && !META_isProtectedHeader_(sh, cols.interR)) {
    const ir = impr > 0 ? (inters / impr) : 0;
    Utils_setPreserveFormula(sh, row, cols.interR, ir, '0.00%');
  }

  if (cols.conv && !META_isProtectedHeader_(sh, cols.conv)) {
    const calls = Number(sh.getRange(row, cols.appels).getValue() || 0);
    const forms = Number(sh.getRange(row, cols.forms).getValue() || 0);
    const cr = clicks > 0 ? ((calls + forms) / clicks) : 0;
    Utils_setPreserveFormula(sh, row, cols.conv, cr, '0.00%');
  }

  // ❌ Toujours : on NE TOUCHE PAS à contacts / CPL / signatures / ROAS / ROI…
}


/**************************** RUNNERS (BATCHE) ******************************/
function run_Meta_FullHistory() {
  run_Meta_FullHistory_Batched();
}

// Traite la fenêtre Jan 2024 → fin du mois dernier, en sous-tranches de META_FULLHIST_BATCH_MONTHS
function run_Meta_FullHistory_Batched() {
  const sh = SpreadsheetApp.getActive().getSheetByName(META_SHEET_NAME);
  if (!sh) throw new Error(`Onglet '${META_SHEET_NAME}' introuvable`);
  const cols = META_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today = new Date();
  const start = META_FULLHIST_START_LOCAL;
  const end = new Date(today.getFullYear(), today.getMonth(), 0); // fin du mois dernier

  const windowKeys = META_monthRangeKeys_(start, end);
  Logger.log(`[Meta/batch] Mois à traiter = ${JSON.stringify(windowKeys)}`);

  for (let i = 0; i < windowKeys.length; i += META_FULLHIST_BATCH_MONTHS) {
    const chunk = windowKeys.slice(i, i + META_FULLHIST_BATCH_MONTHS);

    const fromYM = chunk[0];
    const toYM = chunk[chunk.length - 1];
    const y1 = +fromYM.slice(0, 4), m1 = +fromYM.slice(5, 7) - 1;
    const y2 = +toYM.slice(0, 4), m2 = +toYM.slice(5, 7) - 1;
    const fromLocal = new Date(y1, m1, 1);
    const toLocal = new Date(y2, m2 + 1, 0);

    Utils_toast(`Meta: collecte ${fromYM} → ${toYM}…`, 'Meta Ads', 3);

    const insights = Utils_executeWithRetry(() => META_fetchInsightsMonthly_(fromLocal, toLocal), 'Meta insights');
    const callsByM = Utils_executeWithRetry(() => META_magnetisMonthlyCalls_(fromLocal, toLocal), 'Magnetis calls');
    const formsByM = Utils_executeWithRetry(() => META_formsCountsByMonth_(
      new Date(Date.UTC(fromLocal.getFullYear(), fromLocal.getMonth(), 1)),
      new Date(Date.UTC(toLocal.getFullYear(), toLocal.getMonth(), toLocal.getDate(), 23, 59, 59))
    ), 'Forms (Paper+Monday)');

    let callsLeadByM = {}, formsLeadByM = {};
    if (META_ENABLE_LEADS) {
      try {
        callsLeadByM = Utils_executeWithRetry(() => META_mondayLeadCallsByMonth_(
          META_LEADS_BOARD_ID, META_LEADS_COL_SOURCE, META_LEADS_COL_TYPE, META_LEADS_COL_STATUS,
          new Date(Date.UTC(fromLocal.getFullYear(), fromLocal.getMonth(), 1)),
          new Date(Date.UTC(toLocal.getFullYear(), toLocal.getMonth(), toLocal.getDate(), 23, 59, 59))
        ), 'Monday Leads (calls)');
      } catch (e) { Logger.log('[Meta Leads] Calls KO: ' + e); }
      try {
        formsLeadByM = Utils_executeWithRetry(() => META_mondayLeadFormsByMonth_(
          META_LEADS_BOARD_ID, META_LEADS_COL_SOURCE, META_LEADS_COL_TYPE, META_LEADS_COL_STATUS,
          new Date(Date.UTC(fromLocal.getFullYear(), fromLocal.getMonth(), 1)),
          new Date(Date.UTC(toLocal.getFullYear(), toLocal.getMonth(), toLocal.getDate(), 23, 59, 59))
        ), 'Monday Leads (forms)');
      } catch (e) { Logger.log('[Meta Leads] Forms KO: ' + e); }
    }

    chunk.forEach(ym => {
      const v = {
        spend: (insights[ym]?.spend ?? null),
        impressions: (insights[ym]?.impressions ?? null),
        clicks: (insights[ym]?.clicks ?? null),
        interactions: (insights[ym]?.interactions ?? null),
        calls: (callsByM[ym] ?? null),
        forms: (formsByM[ym] ?? null),
        callsLead: (callsLeadByM[ym] ?? null),
        formsLead: (formsLeadByM[ym] ?? null),
        avgSec: null
      };

      Logger.log(`[Meta/debug] Traitement mois ${ym} | insights=${JSON.stringify(insights[ym] || {})} | calls=${callsByM[ym] || 0} | forms=${formsByM[ym] || 0}`);

      if (META_ENABLE_MATOMO) {
        try { v.avgSec = Utils_executeWithRetry(() => META_matomoVisitAvgSecs_(ym), 'Matomo ' + ym, 2) ?? null; }
        catch (e) { Logger.log('[Matomo] ' + ym + ' KO: ' + e); }
      }

      META_writeMonth_(sh, cols, ym, v);
    });

    Utilities.sleep(250);
  }

  Utils_toast('Meta: Full history (batched) terminé ✅', 'Meta Ads', 5);
}

/**************************** RUNNER N-1 (création si besoin) ***************/
function run_Meta_AddLastMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(META_SHEET_NAME);
  if (!sh) throw new Error(`Onglet '${META_SHEET_NAME}' introuvable`);
  const cols = META_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');

  Utils_toast(`Meta: mise à jour du mois ${Utils_monthKeyToFr(ymKey)}…`, 'Meta Ads', 4);

  const insights = Utils_executeWithRetry(() => META_fetchInsightsMonthly_(start, end), 'Meta N-1');
  const callsByM = Utils_executeWithRetry(() => META_magnetisMonthlyCalls_(start, end), 'Magnetis N-1');
  const formsByM = Utils_executeWithRetry(() => META_formsCountsByMonth_(
    new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
    new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
  ), 'Forms N-1');

  let callsLeadByM = {}, formsLeadByM = {};
  if (META_ENABLE_LEADS) {
    try {
      callsLeadByM = Utils_executeWithRetry(() => META_mondayLeadCallsByMonth_(
        META_LEADS_BOARD_ID, META_LEADS_COL_SOURCE, META_LEADS_COL_TYPE, META_LEADS_COL_STATUS,
        new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
        new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
      ), 'Leads calls N-1');
    } catch (e) { Logger.log('[Meta Leads N-1] Calls KO: ' + e); }
    try {
      formsLeadByM = Utils_executeWithRetry(() => META_mondayLeadFormsByMonth_(
        META_LEADS_BOARD_ID, META_LEADS_COL_SOURCE, META_LEADS_COL_TYPE, META_LEADS_COL_STATUS,
        new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
        new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
      ), 'Leads forms N-1');
    } catch (e) { Logger.log('[Meta Leads N-1] Forms KO: ' + e); }
  }

  const avgSec = META_ENABLE_MATOMO ? Utils_executeWithRetry(() => META_matomoVisitAvgSecs_(ymKey), 'Matomo N-1', 2) : null;

  const v = {
    spend: (insights[ymKey]?.spend ?? null),
    impressions: (insights[ymKey]?.impressions ?? null),
    clicks: (insights[ymKey]?.clicks ?? null),
    interactions: (insights[ymKey]?.interactions ?? null),
    calls: (callsByM[ymKey] ?? null),
    forms: (formsByM[ymKey] ?? null),
    callsLead: (callsLeadByM[ymKey] ?? null),
    formsLead: (formsLeadByM[ymKey] ?? null),
    avgSec: avgSec ?? null
  };

  META_writeMonth_(sh, cols, ymKey, v);

  Utils_toast('Meta: N-1 mis à jour ✅', 'Meta Ads', 5);
}

/*********************** RUNNER MOIS EN COURS (création si besoin) **********/
function run_Meta_CurrentMonth() {
  const sh = SpreadsheetApp.getActive().getSheetByName(META_SHEET_NAME);
  if (!sh) throw new Error(`Onglet '${META_SHEET_NAME}' introuvable`);
  const cols = META_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate()); // jusqu’à aujourd’hui
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');

  Utils_toast(`Meta: mise à jour du mois en cours ${Utils_monthKeyToFr(ymKey)}…`, 'Meta Ads', 4);

  const insights = Utils_executeWithRetry(() => META_fetchInsightsMonthly_(start, end), 'Meta mois en cours');
  const callsByM = Utils_executeWithRetry(() => META_magnetisMonthlyCalls_(start, end), 'Magnetis mois en cours');
  const formsByM = Utils_executeWithRetry(() => META_formsCountsByMonth_(
    new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
    new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
  ), 'Forms mois en cours');

  let callsLeadByM = {}, formsLeadByM = {};
  if (META_ENABLE_LEADS) {
    try {
      callsLeadByM = Utils_executeWithRetry(() => META_mondayLeadCallsByMonth_(
        META_LEADS_BOARD_ID, META_LEADS_COL_SOURCE, META_LEADS_COL_TYPE, META_LEADS_COL_STATUS,
        new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
        new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
      ), 'Leads calls mois en cours');
    } catch (e) { Logger.log('[Meta Leads current] Calls KO: ' + e); }
    try {
      formsLeadByM = Utils_executeWithRetry(() => META_mondayLeadFormsByMonth_(
        META_LEADS_BOARD_ID, META_LEADS_COL_SOURCE, META_LEADS_COL_TYPE, META_LEADS_COL_STATUS,
        new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
        new Date(Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59))
      ), 'Leads forms mois en cours');
    } catch (e) { Logger.log('[Meta Leads current] Forms KO: ' + e); }
  }

  const avgSec = META_ENABLE_MATOMO ? Utils_executeWithRetry(() => META_matomoVisitAvgSecs_(ymKey), 'Matomo mois en cours', 2) : null;

  const v = {
    spend: (insights[ymKey]?.spend ?? null),
    impressions: (insights[ymKey]?.impressions ?? null),
    clicks: (insights[ymKey]?.clicks ?? null),
    interactions: (insights[ymKey]?.interactions ?? null),
    calls: (callsByM[ymKey] ?? null),
    forms: (formsByM[ymKey] ?? null),
    callsLead: (callsLeadByM[ymKey] ?? null),
    formsLead: (formsLeadByM[ymKey] ?? null),
    avgSec: avgSec ?? null
  };

  META_writeMonth_(sh, cols, ymKey, v);

  Utils_toast('Meta: mois en cours mis à jour ✅', 'Meta Ads', 5);
}

/**************************** RUNNER RAPIDE SANS MATOMO/LEADS ***************/
function run_Meta_FullHistory_FastCore() {
  const prevMatomo = META_ENABLE_MATOMO;
  const prevLeads = META_ENABLE_LEADS;
  META_ENABLE_MATOMO = false;
  META_ENABLE_LEADS = false;
  try { run_Meta_FullHistory_Batched(); }
  finally {
    META_ENABLE_MATOMO = prevMatomo;
    META_ENABLE_LEADS = prevLeads;
  }
}

/**************************** MENU ****************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Update Meta Ads')
    .addItem('▶ Full history (Jan 2024 → N-1)', 'run_Meta_FullHistory')
    .addItem('▶ Mois précédent (N-1)', 'run_Meta_AddLastMonth')
    .addItem('▶ Mois en cours', 'run_Meta_CurrentMonth')
    .addItem('▶ Full history FAST (sans Matomo/Leads)', 'run_Meta_FullHistory_FastCore')
    .addToUi();
}
