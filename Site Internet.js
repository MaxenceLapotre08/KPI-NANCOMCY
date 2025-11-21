/****************************** CONFIG FEUILLE ******************************/
const SHEET_NAME_SITE   = 'Site Internet';
const HEADERS_ROW_SITE  = 3;
const START_ROW_SITE    = 6;

/****************************** CONFIG GA4 / GSC ****************************/
const GA4_PROPERTY = 'properties/333690282';
const GSC_SITE_URL = 'https://www.nancomcy.fr/';

/****************************** CONFIG CHECK POSITION ************************/
const SITE_CP_DOMAIN = 'nancomcy.fr';
function SITE_getCheckPositionToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('CHECK_POSITION_TOKEN');
  if (!t) throw new Error("Propriété 'CHECK_POSITION_TOKEN' manquante.");
  return t;
}

/****************************** CONFIG MAGNETIS *****************************/
function SITE_getMagApiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('MAGNETIS_API_KEY');
  if (!k) throw new Error("Propriété 'MAGNETIS_API_KEY' manquante.");
  return k;
}
const SITE_MAG_FILTER_CHANNELS = ['Tout le SEO'];
const SITE_MAG_MATCH_MODE = 'equals';
const SITE_MAG_COUNT_ANSWERED_ONLY = false;

/****************************** CONFIG PAPERFORM ****************************/
const SITE_PAPERFORM_FORM_IDS = ['xuefmtzg'];
function SITE_getPaperformToken_() {
  const props = PropertiesService.getScriptProperties();
  const t = props.getProperty('PAPARFORM_TOKEN') || props.getProperty('PAPERFORM_TOKEN');
  if (!t) throw new Error("Propriété 'PAPARFORM_TOKEN' (ou 'PAPERFORM_TOKEN') manquante.");
  return t;
}

/****************************** CONFIG MONDAY (forms) ************************/
const SITE_MONDAY_BOARD_ID = 9950271520;
const SITE_MONDAY_FORM_COLUMN_TITLE = 'Site Internet';
const SITE_MONDAY_FORM_COLUMN_ID = '';
const SITE_MONDAY_MATCH_MODE = 'nonempty';
const SITE_MONDAY_MATCH_VALUES = [];

function SITE_getMondayToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('MONDAY_TOKEN');
  if (!t) throw new Error("Propriété 'MONDAY_TOKEN' manquante.");
  return t;
}

/************* CONFIG MONDAY (leads: appels + formulaires) ******************/
const SITE_MONDAY_LEADS_BOARD_ID = SITE_MONDAY_BOARD_ID;
const SITE_MONDAY_LEADS_COL_SOURCE = 'Source du Lead';
const SITE_MONDAY_LEADS_COL_TYPE   = 'FORMULAIRE / APPELS';
// ★ nouvelle colonne pour le statut “Lead”
const SITE_MONDAY_LEADS_COL_STATUS = 'Nature du contact'; // ← mets le libellé exact de ta colonne statut
const SITE_MONDAY_LEADS_COL_DATE   = 'Date d\'entrée'; // colonne Date à utiliser si dispo
// correspondances (insensibles casse/accents)
const SITE_MONDAY_LEADS_SOURCE_MATCH = ['site internet','site-internet','site','SITE INTERNET'];
const SITE_MONDAY_LEADS_TYPE_MATCH   = ['appel','call','téléphone','Appel'];
// ★ correspondances pour FORMULAIRE
const SITE_MONDAY_LEADS_FORM_MATCH  = ['formulaire','form','paperform','Formulaire'];
// ★ correspondances pour le statut Lead
const SITE_MONDAY_LEADS_STATUS_MATCH = ['lead','prospect','Lead'];

/****************************** HELPERS COMMUNS *****************************/
function SITE_toast_(msg, title, seconds){ try{ SpreadsheetApp.getActive().toast(msg, title||'Site', seconds||5); }catch(e){} }
function SITE_executeWithRetry_(fn, label, maxRetries){
  label = label||'task'; maxRetries = maxRetries||3; let last;
  for (let i=1;i<=maxRetries;i++){
    try{
      const t0=Date.now(); const out = fn(); Logger.log(`[OK ] ${label} (${((Date.now()-t0)/1000).toFixed(1)}s)`); return out;
    }catch(e){
      last=e; const msg=String(e&&e.message||e); Logger.log(`[ERR] ${label} try ${i}/${maxRetries}: ${msg}`);
      if (i===maxRetries || !/(^|\s)(429|5\d\d|RATE_LIMIT|RESOURCE_EXHAUSTED)/i.test(msg)) break;
      Utilities.sleep(Math.pow(2,i)*1000);
    }
  }
  throw last;
}
function SITE__normHeader_(s){
  return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[’'`]/g,'').replace(/\s+/g,' ').trim();
}
function SITE_findColByHeaderAliases_(sheet, aliases, headerRow){
  const row = headerRow || HEADERS_ROW_SITE;
  const headers = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]||[];
  const HN = headers.map(SITE__normHeader_);
  const wanted = aliases.map(SITE__normHeader_);
  for (let i=0;i<HN.length;i++){
    const h = HN[i]; if(!h) continue;
    if (wanted.some(w=>h===w || h.includes(w))) return i+1;
  }
  return 0;
}
function SITE_isYearSeparatorRow_(cell){
  if (cell==null) return false;
  if (Object.prototype.toString.call(cell)==='[object Date]' && !isNaN(cell.getTime())) return false;
  const s=String(cell).trim();
  if (/^\d{4}$/.test(s)){ const y=+s; return y>=1900 && y<=2100; }
  if (typeof cell==='number' && isFinite(cell)){ const y=Math.round(cell); return y>=1900 && y<=2100; }
  return false;
}
function SITE_sheetCellToYYYYMM_(cell){
  // 1) Si c'est déjà un objet Date → on formate direct
  if (Object.prototype.toString.call(cell)==='[object Date]' && !isNaN(cell.getTime())){
    const y = cell.getFullYear();
    const m = cell.getMonth()+1;
    return `${y}-${String(m).padStart(2,'0')}`;
  }

  const raw = String(cell||'').trim();
  if (!raw) return null;

  // Normalisation pour gérer accents, majuscules, etc.
  const lower = raw
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'') // enlève les accents
    .replace(/\s+/g,' ')
    .trim();

  let m;

  // 2) Formats numériques déjà gérés (on garde ta logique existante)
  if ((m = lower.match(/^(\d{4})-(\d{1,2})$/))) {
    return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  }
  if ((m = lower.match(/^(\d{4})\/(\d{1,2})$/))) {
    return `${m[1]}-${String(+m[2]).padStart(2,'0')}`;
  }
  if ((m = lower.match(/^(\d{4})(\d{2})$/))) {
    return `${m[1]}-${m[2]}`;
  }
  if ((m = lower.match(/^(\d{4})-(\d{2})-\d{2}$/))) {
    return `${m[1]}-${m[2]}`;
  }

  // 3) Formats texte FR : "novembre 2025", "fevr 2025", "sept. 2024"…
  const frMap = {
    'janvier':1, 'janv':1, 'jan':1,
    'fevrier':2, 'fevr':2, 'fev':2,
    'mars':3,
    'avril':4, 'avr':4,
    'mai':5,
    'juin':6,
    'juillet':7, 'juil':7,
    'aout':8, 'août':8, 'aou':8,
    'septembre':9, 'sept':9,
    'octobre':10, 'oct':10,
    'novembre':11, 'nov':11,
    'decembre':12, 'dec':12
  };

  // ex : "novembre 2025" / "nov. 2025"
  if ((m = lower.match(/^([a-z\.]+)\s+(\d{4})$/))) {
    const name = m[1].replace(/\./g,'');
    const year = m[2];
    const monthNum = frMap[name];
    if (monthNum) {
      return `${year}-${String(monthNum).padStart(2,'0')}`;
    }
  }

  // 4) Dernière chance : laisser le parseur JS tenter sa chance
  const d = new Date(raw);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const mm = d.getMonth()+1;
    return `${y}-${String(mm).padStart(2,'0')}`;
  }

  return null;
}

function SITE_monthKeyToFr_(ym){
  const FR=['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  const y=ym.slice(0,4), m=Math.max(0,Math.min(11, parseInt(ym.slice(5,7),10)-1));
  return `${FR[m]} ${y}`;
}
function SITE_setSecondsAsDuration_(sh, row, col, secs){
  const days = (typeof secs==='number' && !isNaN(secs)) ? secs/86400 : 0;
  const rng = sh.getRange(row,col); rng.setValue(days); rng.setNumberFormat('[h]:mm:ss');
}
function SITE_normalizeDomain_(d){
  return String(d||'').toLowerCase().replace(/^https?:\/\//,'').replace(/^www\./,'').replace(/\/.*$/,'').trim();
}

/************ Insertion ligne “année” et mois ************/
function SITE__styleYearRow_(sh,row,moisCol){
  sh.getRange(row,1,1,sh.getLastColumn()).setBackground('#e6e1f5');
  sh.getRange(row,moisCol).setFontWeight('bold');
}
function SITE__findOrInsertMonthRow_(sh, moisCol, targetYM){
  const targetYear = parseInt(targetYM.slice(0,4),10);
  const lastRow = sh.getLastRow();

  // Si la feuille n'a encore aucune donnée sous les entêtes
  if (lastRow < START_ROW_SITE){
    sh.insertRowsBefore(START_ROW_SITE, 2);
    sh.getRange(START_ROW_SITE, moisCol).setValue(String(targetYear));
    SITE__styleYearRow_(sh, START_ROW_SITE, moisCol);
    return START_ROW_SITE + 1;
  }

  const nbRows = lastRow - START_ROW_SITE + 1;
  const values = sh.getRange(START_ROW_SITE, moisCol, nbRows, 1).getValues();

  let existingRow = null;
  let yearRowTarget = null;
  let firstGreaterYearRow = null;
  let lastMonthRowInTargetYear = null;

  for (let i = 0; i < values.length; i++){
    const rowIndex = START_ROW_SITE + i;
    const v = values[i][0];

    // Ligne "année" (ex: 2024)
    if (SITE_isYearSeparatorRow_(v)){
      const y = parseInt(String(v).trim(),10);
      if (y === targetYear && yearRowTarget === null) {
        yearRowTarget = rowIndex;
      }
      if (y > targetYear && firstGreaterYearRow === null) {
        firstGreaterYearRow = rowIndex;
      }
      continue;
    }

    // Ligne "mois" (ex: "novembre 2024", "2024-11"...)
    const ym = SITE_sheetCellToYYYYMM_(v);
    if (!ym) continue;

    if (ym === targetYM){
      existingRow = rowIndex;
      break;
    }

    const y = parseInt(ym.slice(0,4),10);
    if (y === targetYear){
      lastMonthRowInTargetYear = rowIndex;
    }
  }

  // 1) Si le mois existe déjà → on réutilise la ligne
  if (existingRow) return existingRow;

  // 2) S'il y a déjà des mois pour cette année → on insère juste après le dernier mois
  if (lastMonthRowInTargetYear){
    sh.insertRowsAfter(lastMonthRowInTargetYear, 1);
    return lastMonthRowInTargetYear + 1;
  }

  // 3) S'il y a déjà la ligne "année" mais aucun mois → on ajoute juste en dessous
  if (yearRowTarget){
    sh.insertRowsAfter(yearRowTarget, 1);
    return yearRowTarget + 1;
  }

  // 4) Sinon, on insère l'année avant la première année plus grande
  if (firstGreaterYearRow){
    sh.insertRowsBefore(firstGreaterYearRow, 2);
    sh.getRange(firstGreaterYearRow, moisCol).setValue(String(targetYear));
    SITE__styleYearRow_(sh, firstGreaterYearRow, moisCol);
    return firstGreaterYearRow + 1;
  }

  // 5) Dernier cas : aucune année supérieure → on ajoute à la fin
  const yearRowIns = lastRow + 1;
  sh.insertRowsBefore(yearRowIns, 2);
  sh.getRange(yearRowIns, moisCol).setValue(String(targetYear));
  SITE__styleYearRow_(sh, yearRowIns, moisCol);
  return yearRowIns + 1;
}

/******************************** GA4 ***************************************/
function ga4FetchMonthly_(property, startDate, endDate){
  if (typeof AnalyticsData==='undefined' || !AnalyticsData.Properties) {
    throw new Error('Active le service avancé "AnalyticsData" pour GA4.');
  }
  const request = {
    dateRanges: [{ startDate, endDate }],
    dimensions: [{name:'year'},{name:'month'}],
    metrics: [{name:'sessions'},{name:'averageSessionDuration'},{name:'bounceRate'}],
  };
  const resp = AnalyticsData.Properties.runReport(request, property);
  const rows = resp.rows || [];
  const out = {};
  rows.forEach(r=>{
    const year = r.dimensionValues[0].value;
    const month= r.dimensionValues[1].value.padStart(2,'0');
    const key = `${year}-${month}`;
    out[key] = {
      sessions: Number(r.metricValues[0].value || 0),
      avgSec:   Number(r.metricValues[1].value || 0),
      bouncePct:Number(r.metricValues[2].value || 0)
    };
  });
  return out;
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
function gscFetchMonthly_(siteUrl, startDate, endDate){
  const req = { startDate, endDate, dimensions:['date'], rowLimit: 25000 };
  const resp = gscQuery_(siteUrl, req);
  const rows = (resp && resp.rows) || [];
  const buckets = {};
  rows.forEach(row=>{
    const d = row.keys && row.keys[0];
    if (!d) return;
    const ym = d.slice(0,7);
    const imp = Number(row.impressions||0);
    const clk = Number(row.clicks||0);
    if (!buckets[ym]) buckets[ym] = { imp:0, clk:0 };
    buckets[ym].imp += imp;
    buckets[ym].clk += clk;
  });
  const out = {};
  Object.keys(buckets).forEach(ym=>{
    const b = buckets[ym];
    out[ym] = { impressions: b.imp, clicks: b.clk };
  });
  return out;
}

/******************************** Magnetis **********************************/
function SITE_valAt_(obj,path){ return path.split('.').reduce((o,k)=>(o&&o[k]!=null?o[k]:undefined), obj); }
function SITE_getChannelName_(c){
  const candidates=['channel_name','channel','analysis.channel','analytics.channel','utm.channel','session.channel'];
  for (const p of candidates){ const v=SITE_valAt_(c,p); if (v!=null && String(v).trim()!=='') return String(v).trim().toLowerCase(); }
  return '';
}
function SITE_magChannelOk_(name){
  if (!SITE_MAG_FILTER_CHANNELS || SITE_MAG_FILTER_CHANNELS.length===0) return true;
  const list = SITE_MAG_FILTER_CHANNELS.map(x=>String(x).toLowerCase());
  switch(SITE_MAG_MATCH_MODE){
    case 'includes': return list.some(a=>name.includes(a));
    case 'regex':    return list.some(rx=>new RegExp(rx,'i').test(name));
    default:         return list.some(a=>name===a);
  }
}
function SITE_magnetisFetchCalls_(fromDate, toDate){
  const apiKey = SITE_getMagApiKey_();
  const base = 'https://api.magnetis.io/calls';
  const fmtUTC = d => Utilities.formatDate(d,'UTC','yyyyMMddHHmmss');
  let start=new Date(fromDate), end=new Date(toDate);
  if (start>end){ const t=start; start=end; end=t; }
  const fromStr=fmtUTC(start), toStr=fmtUTC(end);
  let page=1, all=[];
  while(true){
    const params={from:fromStr,to:toStr,limit:250,page,analysis:1};
    const qs = Object.keys(params).map(k=>`${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`).join('&');
    const res = UrlFetchApp.fetch(`${base}?${qs}`, { method:'get', headers:{'x-api-key':apiKey,'Accept':'application/json'}, muteHttpExceptions:true });
    if (res.getResponseCode()>=300) throw new Error(`Magnetis HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    const json = JSON.parse(res.getContentText());
    const data = json.data || json || [];
    if (!data.length) break;
    all = all.concat(data);
    page++;
    if (json.links && !json.links.next) break;
    if (!json.links && data.length<250) break;
  }
  return all;
}
function SITE_magnetisAggregateMonthlyCounts_(calls){
  const bucket={};
  calls.forEach(c=>{
    const ch = SITE_getChannelName_(c); if (!SITE_magChannelOk_(ch)) return;
    const raw = c.start_at || c.created_at || c.date || c.started_at;
    if (!raw) return;
    const d=new Date(raw); if (isNaN(d.getTime())) return;
    if (SITE_MAG_COUNT_ANSWERED_ONLY){
      const dur = Number(c.duration || (c.analysis && c.analysis.duration) || 0);
      if (!(dur>0)) return;
    }
    const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    bucket[key]=(bucket[key]||0)+1;
  });
  return bucket;
}

/******************************** Paperform *********************************/
function SITE_extractPaperformArray_(json){
  if (!json) return [];
  if (Array.isArray(json)) return json;
  if (json.results){
    if (Array.isArray(json.results)) return json.results;
    if (Array.isArray(json.results.submissions)) return json.results.submissions;
    if (Array.isArray(json.results.data)) return json.results.data;
    for (const k in json.results){ if (Array.isArray(json.results[k])) return json.results[k]; }
  }
  if (Array.isArray(json.data)) return json.data;
  for (const k in json){ if (Array.isArray(json[k])) return json[k]; }
  return [];
}
function SITE_parsePaperformHumanDateToUTC_(str){
  if (!str||typeof str!=='string') return new Date(NaN);
  const s=str.trim().replace(/\s+/g,' ');
  const MM={jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,oct:9,october:9,nov:10,november:10,dec:11,december:11};
  const m=s.match(/^([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\d{2,4})[, ]\s*(\d{1,2}):(\d{2})\s*(am|pm)$/i);
  if (!m) return new Date(NaN);
  const mon=MM[m[1].toLowerCase()]; let year=+m[3]; let h=+m[4]; const min=+m[5]; const ap=m[6].toLowerCase();
  if (isNaN(mon)||isNaN(year)||isNaN(h)||isNaN(min)) return new Date(NaN);
  if (year<100) year=2000+year; h=h%12; if(ap==='pm') h+=12;
  return new Date(Date.UTC(year,mon,+m[2],h,min,0));
}
function SITE_getSubmissionDate_(s){
  const cand=[s.submitted_at,s.submittedAt,s.submitted_at_utc,s.submittedAtUTC,s.created_at,s.createdAt,s.created_at_utc,s.createdAtUTC,s.date,s.timestamp,s.time_submitted,s.timeSubmitted,s.created,s.submitted,(s.meta&&(s.meta.submitted_at||s.meta.created_at))];
  for (let v of cand){
    if (v==null) continue;
    if (typeof v==='number'){ const ms = v<1e12 ? v*1000 : v; const d=new Date(ms); if(!isNaN(d.getTime())) return d; continue; }
    if (typeof v==='string'){
      const dStd=new Date(v); if(!isNaN(dStd.getTime())) return dStd;
      const dHuman=SITE_parsePaperformHumanDateToUTC_(v); if(dHuman&&!isNaN(dHuman.getTime())) return dHuman;
      if (/^\d+$/.test(v)){ const n=+v, ms=n<1e12?n*1000:n; const dNum=new Date(ms); if(!isNaN(dNum.getTime())) return dNum; }
    }
  }
  return new Date(NaN);
}
function SITE_paperformFetchSubmissionsInRange_(slugOrId, fromDate, toDate){
  const token = SITE_getPaperformToken_();
  const base = `https://api.paperform.co/v1/forms/${encodeURIComponent(slugOrId)}/submissions`;
  const limit=100; let all=[]; let mode='skip', skip=0, page=1;
  const fromTs=fromDate.getTime(), toTs=toDate.getTime();
  while(true){
    const url = (mode==='skip') ? `${base}?limit=${limit}&skip=${skip}` : `${base}?limit=${limit}&page=${page}`;
    const res = UrlFetchApp.fetch(url, { method:'get', headers:{'Authorization':'Bearer '+token,'Accept':'application/json'}, muteHttpExceptions:true });
    if (res.getResponseCode()===422 && mode==='skip'){ mode='page'; page=1; continue; }
    if (res.getResponseCode()>=300) throw new Error(`Paperform HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
    const batch = SITE_extractPaperformArray_(JSON.parse(res.getContentText())); if (!batch.length) break;
    for (const s of batch){ const d=SITE_getSubmissionDate_(s); if (!isNaN(d.getTime())){ const t=d.getTime(); if (t>=fromTs && t<=toTs) all.push(s); } }
    if (mode==='skip') skip+=batch.length; else page+=1;
    if (batch.length<limit) break;
    Utilities.sleep(120);
  }
  return all;
}
function SITE_paperformCountsByMonth_(fromDate, toDate){
  const bucket={};
  SITE_PAPERFORM_FORM_IDS.forEach(fid=>{
    const subs = SITE_paperformFetchSubmissionsInRange_(fid, fromDate, toDate);
    subs.forEach(s=>{
      const d=SITE_getSubmissionDate_(s); if (isNaN(d.getTime())) return;
      const key = d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0');
      bucket[key]=(bucket[key]||0)+1;
    });
  });
  return bucket;
}

/******************************** Monday (général) **************************/
function SITE__norm_(s){ return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[’'`]/g,'').replace(/\s+/g,' ').trim(); }
function SITE_mondayGraphQL_(query,variables){
  const res = UrlFetchApp.fetch('https://api.monday.com/v2', {
    method:'post',
    headers:{'Content-Type':'application/json','Authorization':SITE_getMondayToken_(),'API-Version':'2023-10'},
    payload: JSON.stringify({query, variables:variables||{}}),
    muteHttpExceptions:true
  });
  if (res.getResponseCode()>=300) throw new Error(`Monday HTTP ${res.getResponseCode()}: ${res.getContentText()}`);
  const body = JSON.parse(res.getContentText());
  if (body.errors) throw new Error('Monday GraphQL error: '+JSON.stringify(body.errors));
  return body.data;
}
function SITE_mondayResolveColumnId_(boardId, titleOrId){
  if (!titleOrId) return '';
  const looksId = /^[a-z0-9_]+$/i.test(titleOrId) && !SITE__norm_(titleOrId).includes(' ');
  if (looksId) return titleOrId;
  const want=SITE__norm_(titleOrId);
  const q=`query($bid:[ID!]){ boards(ids:$bid){ columns{ id title type } } }`;
  const d=SITE_mondayGraphQL_(q,{bid:[Number(boardId)]});
  const cols=(d&&d.boards&&d.boards[0]&&d.boards[0].columns)||[];
  let best='';
  for (const c of cols){ if (SITE__norm_(c.title)===want){ best=c.id; break; } }
  if (!best){ for (const c of cols){ if (SITE__norm_(c.title).includes(want)){ best=c.id; break; } } }
  return best;
}
function SITE_mondayValueMatches_(txt){
  const t=SITE__norm_(txt||'');
  switch(SITE_MONDAY_MATCH_MODE){
    case 'equals':   return SITE_MONDAY_MATCH_VALUES.some(v=>t===SITE__norm_(v));
    case 'includes': return SITE_MONDAY_MATCH_VALUES.some(v=>t.includes(SITE__norm_(v)));
    default:         return t.length>0;
  }
}
function SITE_mondayFormCountsByMonth_(boardId, colTitleOrId, fromUTC, toUTC){
  const colId = SITE_MONDAY_FORM_COLUMN_ID ? SITE_MONDAY_FORM_COLUMN_ID : SITE_mondayResolveColumnId_(boardId, colTitleOrId);
  if (!colId) throw new Error("Colonne Monday 'Site Internet' introuvable.");
  const counts={}; let cursor=null;
  do{
    const q=`
      query($bid:[ID!], $cursor:String, $col:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{ id state created_at column_values(ids:$col){ id text } }
          }
        }
      }`;
    const d=SITE_mondayGraphQL_(q,{bid:[Number(boardId)],cursor, col:[colId]});
    const page=d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;
    (page.items||[]).forEach(it=>{
      if (String(it.state||'').toLowerCase()==='archived') return;
      const dt=new Date(it.created_at); if (isNaN(dt.getTime())) return;
      if (dt<fromUTC || dt>toUTC) return;
      const txt=(it.column_values&&it.column_values[0]&&it.column_values[0].text)||'';
      if (!SITE_mondayValueMatches_(txt)) return;
      const key = dt.getUTCFullYear()+'-'+String(dt.getUTCMonth()+1).padStart(2,'0');
      counts[key]=(counts[key]||0)+1;
    });
    cursor=page.cursor;
  } while(cursor);
  return counts;
}

function SITE_textEqualsAny_(txt, arr){
  const t = SITE__norm_(txt||'');
  return (arr||[]).some(v => t === SITE__norm_(v));
}

/************ Monday — Lead calls & lead forms ******************************/
function SITE_resolveTwoCols_(boardId, titleOrId1, titleOrId2){
  const id1 = SITE_mondayResolveColumnId_(boardId, titleOrId1);
  const id2 = SITE_mondayResolveColumnId_(boardId, titleOrId2);
  if (!id1 || !id2) throw new Error('Colonnes Monday (source/type) introuvables.');
  return {id1, id2};
}
function SITE_textIncludesAny_(txt, arr){
  const t = SITE__norm_(txt||'');
  return (arr||[]).some(v => t.includes(SITE__norm_(v)));
}
// Appels lead
// { 'YYYY-MM': n } — APPELS lead = Source match + Type "appel" (date=colonne Date si dispo sinon created_at)
function SITE_mondayLeadCallsCountsByMonth_(boardId, colSource, colType, fromUTC, toUTC){
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId   = SITE_mondayResolveColumnId_(boardId, colType);
  const dateId   = (typeof SITE_MONDAY_LEADS_COL_DATE !== 'undefined' && SITE_MONDAY_LEADS_COL_DATE)
                   ? SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_DATE)
                   : null;
  if (!sourceId || !typeId) throw new Error('Colonnes Monday (source/type) introuvables.');

  const counts = {};
  let cursor = null;

  do{
    const q = `
      query($bid:[ID!], $cursor:String, $cols:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id state created_at
              column_values(ids:$cols){ id text value }
            }
          }
        }
      }`;
    const colIds = [sourceId, typeId].concat(dateId ? [dateId] : []);
    const d = SITE_mondayGraphQL_(q,{bid:[Number(boardId)], cursor, cols: colIds});
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page; if (!page) break;

    (page.items||[]).forEach(it=>{
      if (String(it.state||'').toLowerCase()==='archived') return;

      const createdAt = new Date(it.created_at);
      const cv = it.column_values || [];
      const srcText = (cv.find(c=>c.id===sourceId)||{}).text || '';
      const typText = (cv.find(c=>c.id===typeId)  ||{}).text || '';
      const dcv     = dateId ? ((cv.find(c=>c.id===dateId)||{}).value || '') : '';
      const when    = SITE_parseMondayDateValue_(dcv) || createdAt;

      if (!(when >= fromUTC && when <= toUTC)) return;

      // Source : ÉGALITÉ stricte à une des valeurs autorisées (pas de "includes")
      if (!SITE_textEqualsAny_(srcText, SITE_MONDAY_LEADS_SOURCE_MATCH)) return;

      // Type : doit contenir "appel" (on garde la souplesse "includes")
      if (!SITE_textIncludesAny_(typText, SITE_MONDAY_LEADS_TYPE_MATCH)) return;
    });

    cursor = page.cursor;
  } while(cursor);

  return counts;
}

// ★ Formulaires lead (Source=Site Internet, Type=Formulaire, Statut=Lead)
// { 'YYYY-MM': n } — FORMULAIRES lead = Source match + Type "formulaire" + Statut "lead"
function SITE_mondayLeadFormsCountsByMonth_(boardId, colSource, colType, colStatus, fromUTC, toUTC){
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId   = SITE_mondayResolveColumnId_(boardId, colType);
  const statusId = colStatus ? SITE_mondayResolveColumnId_(boardId, colStatus) : null;
  const dateId   = SITE_MONDAY_LEADS_COL_DATE ? SITE_mondayResolveColumnId_(boardId, SITE_MONDAY_LEADS_COL_DATE) : null;
  if (!sourceId || !typeId) throw new Error('Colonnes Monday (source/type) introuvables.');

  const counts = {};
  let cursor = null;

  do{
    const q = `
      query($bid:[ID!], $cursor:String, $cols:[String!]){
        boards(ids:$bid){
          items_page(limit:500, cursor:$cursor){
            cursor
            items{
              id state created_at
              column_values(ids:$cols){ id text value }
            }
          }
        }
      }`;
    const colIds = [sourceId, typeId].concat(statusId?[statusId]:[]).concat(dateId?[dateId]:[]);
    const d = SITE_mondayGraphQL_(q,{bid:[Number(boardId)], cursor, cols: colIds});
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page; if (!page) break;

    (page.items||[]).forEach(it=>{
      if (String(it.state||'').toLowerCase()==='archived') return;
      const createdAt = new Date(it.created_at);
      const cv = it.column_values || [];
      const src = (cv.find(c=>c.id===sourceId)||{}).text || '';
      const typ = (cv.find(c=>c.id===typeId)  ||{}).text || '';
      const sts = statusId ? ((cv.find(c=>c.id===statusId)||{}).text || '') : '';
      const dcv = dateId ? ((cv.find(c=>c.id===dateId)||{}).value || '') : '';
      const when = SITE_parseMondayDateValue_(dcv) || createdAt;

      if (!(when >= fromUTC && when <= toUTC)) return;
      if (!SITE_textIncludesAny_(src, SITE_MONDAY_LEADS_SOURCE_MATCH)) return;
      if (!SITE_textIncludesAny_(typ, SITE_MONDAY_LEADS_FORM_MATCH))   return;
      if (statusId && !SITE_textIncludesAny_(sts, SITE_MONDAY_LEADS_STATUS_MATCH)) return;

      const key = when.getUTCFullYear()+'-'+String(when.getUTCMonth()+1).padStart(2,'0');
      counts[key] = (counts[key]||0) + 1;
    });

    cursor = page.cursor;
  } while(cursor);

  return counts;
}


/********************* FUSION “Nombre de formulaire” ************************/
function SITE_formsCountsByMonth_(fromDateUTC, toDateUTC){
  const paper = SITE_paperformCountsByMonth_(fromDateUTC, toDateUTC);
  const monday= SITE_mondayFormCountsByMonth_(SITE_MONDAY_BOARD_ID, SITE_MONDAY_FORM_COLUMN_TITLE, fromDateUTC, toDateUTC);
  const keys = Array.from(new Set(Object.keys(paper).concat(Object.keys(monday))));
  const out={}; keys.forEach(k=> out[k]=(paper[k]||0)+(monday[k]||0) );
  return out;
}

/******************************** CHECK POSITION ****************************/
function SITE_cpGetAccountHid_() {
  const token = SITE_getCheckPositionToken_();
  const res = UrlFetchApp.fetch('https://api.check-position.com/account', {method:'get', headers:{ 'Authorization':'Bearer '+token }, muteHttpExceptions:true});
  if (res.getResponseCode()>=300) throw new Error('CheckPosition /account HTTP '+res.getResponseCode()+': '+res.getContentText());
  const data = JSON.parse(res.getContentText()).data;
  return data && (data.hid || data.id || null);
}
function SITE_cpGetMonitors_() {
  const token = SITE_getCheckPositionToken_();
  const q = {'with[]': ['website','keyword','tags']};
  let queryUrl = Object.keys(q).map(k=> q[k].map(v=>k+'='+encodeURIComponent(v)).join('&')).join('&');
  const url = 'https://api.check-position.com/monitors?'+queryUrl;
  const res = UrlFetchApp.fetch(url, {method:'get', headers:{'Authorization':'Bearer '+token}, muteHttpExceptions:true});
  if (res.getResponseCode()>=300) throw new Error('CheckPosition /monitors HTTP '+res.getResponseCode()+': '+res.getContentText());
  return JSON.parse(res.getContentText()).data || [];
}
function SITE_cpFetchMonthReport_(accountHid, ymKey){
  const y = +ymKey.slice(0,4), m = +ymKey.slice(5,7)-1;
  for (let back=0; back<8; back++){
    const d = new Date(Date.UTC(y, m+1, 0-back, 0,0,0));
    const yyyy = d.getUTCFullYear(), mm = String(d.getUTCMonth()+1).padStart(2,'0'), dd = String(d.getUTCDate()).padStart(2,'0');
    const url = `https://storage.check-position.com/monitor/report/${yyyy}/${mm}/${dd}/${accountHid}.json`;
    try{
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions:true });
      if (res.getResponseCode()===200) return JSON.parse(res.getContentText());
    }catch(e){}
  }
  return null;
}
function SITE_cpMonthlyAveragePosition_(domain, ymKey, monitorsCache){
  const accountHid = SITE_cpGetAccountHid_();
  const report = SITE_cpFetchMonthReport_(accountHid, ymKey);
  if (!report || !report.monitors) return null;
  const domainNorm = SITE_normalizeDomain_(domain);
  const monitors = monitorsCache || SITE_cpGetMonitors_();
  const wanted = {};
  (monitors||[]).forEach(m=>{
    if (m && m.website && SITE_normalizeDomain_(m.website.domain) === domainNorm) wanted[String(m.id)] = true;
  });
  let sum=0, n=0;
  Object.keys(report.monitors).forEach(mid=>{
    if (!wanted[mid]) return;
    const node = report.monitors[mid];
    let pos = null;
    if (Array.isArray(node)) pos = Number(node[0]);
    else if (node && typeof node==='object' && node.position!=null) pos = Number(node.position);
    else if (typeof node === 'number') pos = Number(node);
    if (pos==null || isNaN(pos)) return;
    if (pos===0) pos = 100;
    sum += pos; n += 1;
  });
  if (n===0) return null;
  return sum / n;
}
function SITE_cpMonthlyAvgBatch_(domain, ymKeys){
  const out={}, accountHid = SITE_cpGetAccountHid_(), monitors = SITE_cpGetMonitors_();
  ymKeys.forEach(ym=>{
    const report = SITE_cpFetchMonthReport_(accountHid, ym);
    if (!report || !report.monitors){ out[ym]=null; return; }
    const domainNorm = SITE_normalizeDomain_(domain);
    let sum=0, n=0;
    (monitors||[]).forEach(m=>{
      if (!m || !m.website) return;
      if (SITE_normalizeDomain_(m.website.domain)!==domainNorm) return;
      const node = report.monitors[String(m.id)];
      if (!node) return;
      let pos = null;
      if (Array.isArray(node)) pos = Number(node[0]);
      else if (node && typeof node==='object' && node.position!=null) pos = Number(node.position);
      else if (typeof node === 'number') pos = Number(node);
      if (pos==null || isNaN(pos)) return;
      if (pos===0) pos = 100;
      sum += pos; n += 1;
    });
    out[ym] = n>0 ? (sum/n) : null;
  });
  return out;
}

/**************************** HELPERS SPÉCIFIQUES FEUILLE *************************/

// Colonne "Budget investi" SÉCURISÉE pour le site
function SITE_findBudgetColSafe_(sheet) {
  const headerRow = HEADERS_ROW_SITE || 3;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const HN = headers.map(SITE__normHeader_);

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
  const wanted = aliases.map(SITE__normHeader_);
  for (let i = 0; i < HN.length; i++) {
    const h = HN[i];
    if (!h || forbiddenRe.test(h)) continue;
    if (wanted.some(w => h === w || h.includes(w))) return i + 1;
  }

  Logger.log('[Site/findBudgetColSafe] Aucune colonne budget trouvée.');
  return 0;
}

function SITE_isProtectedHeader_(sh, col) {
  if (!col) return false;
  const headerVal = sh.getRange(HEADERS_ROW_SITE, col).getValue();
  const h = SITE__normHeader_(headerVal);

  // 🟢 Colonnes autorisées même si elles contiennent "contact", "lead", etc.
  // -> on VEUT que le script les remplisse
  if (/(taux de conversion contact|taux de conversion lead|ctr|taux de clics|taux de clic)/.test(h)) {
    return false;
  }

  // 🟥 Colonnes vraiment protégées (formules ou KPI manuels)
  return /(contact|contacts|lead|leads|signature|signatures|roas|roi|cpl|cout par contact|coût par contact|montant signe|montant signé)/.test(h);
}



/**************************** ÉCRITURE DANS LA FEUILLE ****************************/
function SITE_findCols_(sh){
  return {
    mois:       SITE_findColByHeaderAliases_(sh, ['mois'], HEADERS_ROW_SITE),

    // budget sécurisé (pas utilisé en écriture pour l’instant, mais prêt)
    budget:     SITE_findBudgetColSafe_(sh),

    // CPL = souvent une formule → pas écrite par le script
    cpl:        SITE_findColByHeaderAliases_(sh, ['cout par contact','coût par contact','cpl'], HEADERS_ROW_SITE),

    impr:       SITE_findColByHeaderAliases_(sh, ['impressions','nb impressions','nombre dimpressions'], HEADERS_ROW_SITE),
    visits:     SITE_findColByHeaderAliases_(sh, ['nombre de visites','visites','sessions','sessions (ga4)'], HEADERS_ROW_SITE),

    // CTR (calculé par le script)
    ctr:        SITE_findColByHeaderAliases_(sh, ['ctr','taux de clics','taux de clic'], HEADERS_ROW_SITE),

    dur:        SITE_findColByHeaderAliases_(sh, ['duree moyenne des visites','durée moyenne des visites'], HEADERS_ROW_SITE),
    bounce:     SITE_findColByHeaderAliases_(sh, ['taux de rebond','bounce'], HEADERS_ROW_SITE),

    appels:     SITE_findColByHeaderAliases_(sh, ["nombre d'appels","nombre dappels","appels"], HEADERS_ROW_SITE),
    appelsLead: SITE_findColByHeaderAliases_(sh, ["nombre d'appels lead","appels lead","lead appels"], HEADERS_ROW_SITE),

    forms:      SITE_findColByHeaderAliases_(sh, ['nombre de formulaire','formulaires','forms','soumissions'], HEADERS_ROW_SITE),
    formsLead:  SITE_findColByHeaderAliases_(sh, ['nombre de formulaires lead','formulaires lead','lead formulaires'], HEADERS_ROW_SITE),

    // Colonne "Nombre de contacts" (contact = appels + formulaires)
    contacts:   SITE_findColByHeaderAliases_(sh, ['nombre de contacts','contacts','leads'], HEADERS_ROW_SITE),

    // Ancien taux de conversion générique (si tu en as encore un)
    conv:       SITE_findColByHeaderAliases_(sh, ['taux de conversion','conversion rate','cr'], HEADERS_ROW_SITE),

    // 🆕 Nouveaux taux de conversion dédiés
    convContact: SITE_findColByHeaderAliases_(sh, ['taux de conversion contact','taux de conversion (contact)'], HEADERS_ROW_SITE),
    convLead:    SITE_findColByHeaderAliases_(sh, ['taux de conversion lead','taux de conversion (lead)'], HEADERS_ROW_SITE),

    pos:        SITE_findColByHeaderAliases_(sh, ['position moyenne','avg position','average position'], HEADERS_ROW_SITE),
  };
}


function SITE_writeMonthRow_(sh, cols, ymKey, values){
  const row = SITE__findOrInsertMonthRow_(sh, cols.mois, ymKey);
  sh.getRange(row, cols.mois).setValue(SITE_monthKeyToFr_(ymKey));

  // Écritures autorisées (avec garde-fou sur les headers)
  if (cols.impr && values.impressions != null && !SITE_isProtectedHeader_(sh, cols.impr)) {
    sh.getRange(row, cols.impr).setValue(values.impressions);
  }

  if (cols.visits && values.sessions != null && !SITE_isProtectedHeader_(sh, cols.visits)) {
    sh.getRange(row, cols.visits).setValue(values.sessions);
  }

  if (cols.dur && values.avgSec != null && !SITE_isProtectedHeader_(sh, cols.dur)) {
    SITE_setSecondsAsDuration_(sh, row, cols.dur, values.avgSec);
  }

  if (cols.appels && values.calls != null && !SITE_isProtectedHeader_(sh, cols.appels)) {
    sh.getRange(row, cols.appels).setValue(values.calls);
  }

  if (cols.forms && values.forms != null && !SITE_isProtectedHeader_(sh, cols.forms)) {
    sh.getRange(row, cols.forms).setValue(values.forms);
  }

  if (cols.appelsLead && values.leadCalls != null && !SITE_isProtectedHeader_(sh, cols.appelsLead)) {
    sh.getRange(row, cols.appelsLead).setValue(values.leadCalls);
  }

  if (cols.formsLead && values.leadForms != null && !SITE_isProtectedHeader_(sh, cols.formsLead)) {
    sh.getRange(row, cols.formsLead).setValue(values.leadForms);
  }

  if (cols.bounce && values.bounce != null && !SITE_isProtectedHeader_(sh, cols.bounce)) {
    sh.getRange(row, cols.bounce).setValue(values.bounce).setNumberFormat('0.00%');
  }

  // Position moyenne (OK, donnée brute)
  if (cols.pos && !SITE_isProtectedHeader_(sh, cols.pos)) {
    if (values.avgPos == null || isNaN(values.avgPos)) {
      sh.getRange(row, cols.pos).setValue('Pas de données');
    } else {
      sh.getRange(row, cols.pos).setValue(values.avgPos).setNumberFormat('0.00');
    }
  }

  // ===================== MÉTRIQUES DÉRIVÉES =====================

  // Sessions & impressions "fiables" (on combine ce qu'on vient d'écrire + ce qui existait déjà)
  const sessions = (values.sessions ??
    Number(cols.visits ? sh.getRange(row, cols.visits).getValue() : 0)) || 0;

  const impressions = (values.impressions ??
    Number(cols.impr ? sh.getRange(row, cols.impr).getValue() : 0)) || 0;

  // 1) CTR = sessions / impressions
  if (cols.ctr && !SITE_isProtectedHeader_(sh, cols.ctr)) {
    const ctr = impressions > 0 ? (sessions / impressions) : 0;
    sh.getRange(row, cols.ctr).setValue(ctr).setNumberFormat('0.00%');
  }

  // 2) Taux de conversion CONTACT = Nombre de contacts / sessions
  if (cols.convContact && !SITE_isProtectedHeader_(sh, cols.convContact)) {
    const contactsExisting = Number(
      cols.contacts ? sh.getRange(row, cols.contacts).getValue() : 0
    ) || 0;
    const convContact = sessions > 0 ? (contactsExisting / sessions) : 0;
    sh.getRange(row, cols.convContact).setValue(convContact).setNumberFormat('0.00%');
  }

  // 3) Taux de conversion LEAD = (Appels lead + Formulaires lead) / sessions
  if (cols.convLead && !SITE_isProtectedHeader_(sh, cols.convLead)) {
    const leadCalls = (values.leadCalls != null)
      ? values.leadCalls
      : Number(cols.appelsLead ? sh.getRange(row, cols.appelsLead).getValue() : 0) || 0;

    const leadForms = (values.leadForms != null)
      ? values.leadForms
      : Number(cols.formsLead ? sh.getRange(row, cols.formsLead).getValue() : 0) || 0;

    const leadTotal = leadCalls + leadForms;
    const convLead = sessions > 0 ? (leadTotal / sessions) : 0;
    sh.getRange(row, cols.convLead).setValue(convLead).setNumberFormat('0.00%');
  }

  // 4) Rétrocompat : si tu as encore une colonne "Taux de conversion" générique (conv)
  if (cols.conv && !SITE_isProtectedHeader_(sh, cols.conv)
      && !cols.convContact && !cols.convLead) {
    const contactsExisting = Number(
      cols.contacts ? sh.getRange(row, cols.contacts).getValue() : 0
    ) || 0;
    const conv = sessions > 0 ? (contactsExisting / sessions) : 0;
    sh.getRange(row, cols.conv).setValue(conv).setNumberFormat('0.00%');
  }
}

/**************************** RUNNERS ***************************************/
const SITE_MONTHS_BACK = 21;

function run_Site_FullHistory(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SITE);
  if (!sh) throw new Error(`Onglet '${SHEET_NAME_SITE}' introuvable`);
  const cols = SITE_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today=new Date();
  const start = new Date(today.getFullYear(), today.getMonth()-SITE_MONTHS_BACK, 1);
  const end   = new Date(today.getFullYear(), today.getMonth(), 0);
  const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const endStr   = Utilities.formatDate(end,   Session.getScriptTimeZone(), 'yyyy-MM-dd');

  SITE_toast_('Site: collecte…','Site Internet',4);

  const ga  = SITE_executeWithRetry_(()=>ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr),'GA4 monthly');
  const gsc = SITE_executeWithRetry_(()=>gscFetchMonthly_(GSC_SITE_URL, startStr, endStr),'GSC monthly');
  const callsByMonth = SITE_executeWithRetry_(()=>SITE_magnetisAggregateMonthlyCounts_(SITE_magnetisFetchCalls_(start, end)),'Magnetis calls');
  const formsByMonth = SITE_executeWithRetry_(()=>SITE_formsCountsByMonth_(new Date(Date.UTC(start.getFullYear(),start.getMonth(),1)), new Date(Date.UTC(end.getFullYear(),end.getMonth(),end.getDate(),23,59,59))),'Forms (Paper+Monday)');

  const leadCallsByMonth = SITE_executeWithRetry_(
    ()=>SITE_mondayLeadCallsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      new Date(Date.UTC(start.getFullYear(),start.getMonth(),1)),
      new Date(Date.UTC(end.getFullYear(),end.getMonth(),end.getDate(),23,59,59))
    ),
    'Monday lead calls'
  );

  const leadFormsByMonth = SITE_executeWithRetry_(
    ()=>SITE_mondayLeadFormsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      SITE_MONDAY_LEADS_COL_STATUS,
      new Date(Date.UTC(start.getFullYear(),start.getMonth(),1)),
      new Date(Date.UTC(end.getFullYear(),end.getMonth(),end.getDate(),23,59,59))
    ),
    'Monday lead forms'
  );

  const keys = Array.from(new Set([]
    .concat(Object.keys(ga), Object.keys(gsc), Object.keys(callsByMonth), Object.keys(formsByMonth),
            Object.keys(leadCallsByMonth), Object.keys(leadFormsByMonth))
  )).sort();

  const cpAvg = SITE_executeWithRetry_(()=>SITE_cpMonthlyAvgBatch_(SITE_CP_DOMAIN, keys),'CheckPosition monthly avg');

  keys.forEach(ym=>{
    const values = {
      sessions:    ga[ym]?.sessions ?? null,
      avgSec:      ga[ym]?.avgSec ?? null,
      impressions: gsc[ym]?.impressions ?? null,
      calls:       callsByMonth[ym] ?? null,
      forms:       formsByMonth[ym] ?? null,
      leadCalls:   leadCallsByMonth[ym] ?? null,
      leadForms:   leadFormsByMonth[ym] ?? null,
      avgPos:      cpAvg[ym] ?? null,
      bounce:      ga[ym]?.bouncePct ?? null
    };
    SITE_writeMonthRow_(sh, cols, ym, values);
  });

  SITE_toast_('Site: Full history terminé ✅','Site Internet',5);
}

function run_Site_AddLastMonth(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SITE);
  if (!sh) throw new Error(`Onglet '${SHEET_NAME_SITE}' introuvable`);
  const cols = SITE_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today=new Date();
  const start = new Date(today.getFullYear(), today.getMonth()-1, 1);
  const end   = new Date(today.getFullYear(), today.getMonth(), 0);
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');
  const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const endStr   = Utilities.formatDate(end,   Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const ga  = SITE_executeWithRetry_(()=>ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr),'GA4 N-1');
  const gsc = SITE_executeWithRetry_(()=>gscFetchMonthly_(GSC_SITE_URL, startStr, endStr),'GSC N-1');
  const callsByMonth = SITE_executeWithRetry_(()=>SITE_magnetisAggregateMonthlyCounts_(SITE_magnetisFetchCalls_(start, end)),'Magnetis N-1');
  const formsByMonth = SITE_executeWithRetry_(()=>SITE_formsCountsByMonth_(new Date(Date.UTC(start.getFullYear(),start.getMonth(),1)), new Date(Date.UTC(end.getFullYear(),end.getMonth(),end.getDate(),23,59,59))),'Forms N-1');

  const leadCallsByMonth = SITE_executeWithRetry_(
    ()=>SITE_mondayLeadCallsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      new Date(Date.UTC(start.getFullYear(),start.getMonth(),1)),
      new Date(Date.UTC(end.getFullYear(),end.getMonth(),end.getDate(),23,59,59))
    ),
    'Monday lead calls N-1'
  );

  const leadFormsByMonth = SITE_executeWithRetry_(
    ()=>SITE_mondayLeadFormsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      SITE_MONDAY_LEADS_COL_STATUS,
      new Date(Date.UTC(start.getFullYear(),start.getMonth(),1)),
      new Date(Date.UTC(end.getFullYear(),end.getMonth(),end.getDate(),23,59,59))
    ),
    'Monday lead forms N-1'
  );

  const avgPos = SITE_executeWithRetry_(()=>SITE_cpMonthlyAveragePosition_(SITE_CP_DOMAIN, ymKey, null),'CheckPosition N-1');

  const values = {
    sessions:    ga[ymKey]?.sessions ?? null,
    avgSec:      ga[ymKey]?.avgSec ?? null,
    impressions: gsc[ymKey]?.impressions ?? null,
    calls:       callsByMonth[ymKey] ?? null,
    forms:       formsByMonth[ymKey] ?? null,
    leadCalls:   leadCallsByMonth[ymKey] ?? null,
    leadForms:   leadFormsByMonth[ymKey] ?? null,
    avgPos:      avgPos ?? null,
    bounce:      ga[ymKey]?.bouncePct ?? null
  };
  SITE_writeMonthRow_(sh, cols, ymKey, values);

  SITE_toast_('Site: N-1 mis à jour ✅','Site Internet',5);
}

/*********************** TEST — Leads (appels & formulaires) ************************/

/** Liste détaillée des ITEMS Monday correspondant aux APPELS lead
 *  (Source ∈ SITE_MONDAY_LEADS_SOURCE_MATCH, Type ∈ SITE_MONDAY_LEADS_TYPE_MATCH)
 */
function SITE_mondayFetchLeadCallsItems_(boardId, colSource, colType, fromUTC, toUTC){
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId   = SITE_mondayResolveColumnId_(boardId, colType);
  if (!sourceId || !typeId) throw new Error('Colonnes Monday (source/type) introuvables.');

  const out=[]; let cursor=null;
  do{
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
    const d = SITE_mondayGraphQL_(q,{bid:[Number(boardId)],cursor, cols:[sourceId,typeId]});
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;

    (page.items||[]).forEach(it=>{
      if (String(it.state||'').toLowerCase()==='archived') return;
      const dt=new Date(it.created_at); if (isNaN(dt.getTime())) return;
      if (dt<fromUTC || dt>toUTC) return;

      const cv = it.column_values || [];
      const srcText = (cv.find(c=>c.id===sourceId)||{}).text || '';
      const typText = (cv.find(c=>c.id===typeId)  ||{}).text || '';
      if (!SITE_textIncludesAny_(srcText, SITE_MONDAY_LEADS_SOURCE_MATCH)) return;
      if (!SITE_textIncludesAny_(typText, SITE_MONDAY_LEADS_TYPE_MATCH))   return;

      out.push({id:it.id, name:it.name, created_at:dt.toISOString(), source:srcText, type:typText});
    });

    cursor = page.cursor;
  } while (cursor);

  return out;
}

/** Liste détaillée des ITEMS Monday correspondant aux FORMULAIRES lead
 *  (Source ∈ SITE_MONDAY_LEADS_SOURCE_MATCH, Type ∈ SITE_MONDAY_LEADS_FORM_MATCH, Statut ∈ SITE_MONDAY_LEADS_STATUS_MATCH)
 */
function SITE_mondayFetchLeadFormsItems_(boardId, colSource, colType, colStatus, fromUTC, toUTC){
  const sourceId = SITE_mondayResolveColumnId_(boardId, colSource);
  const typeId   = SITE_mondayResolveColumnId_(boardId, colType);
  const statusId = SITE_mondayResolveColumnId_(boardId, colStatus);
  if (!sourceId || !typeId || !statusId) throw new Error('Colonnes Monday (source/type/statut) introuvables.');

  const out=[]; let cursor=null;
  do{
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
    const d = SITE_mondayGraphQL_(q,{bid:[Number(boardId)],cursor, cols:[sourceId,typeId,statusId]});
    const page = d && d.boards && d.boards[0] && d.boards[0].items_page;
    if (!page) break;

    (page.items||[]).forEach(it=>{
      if (String(it.state||'').toLowerCase()==='archived') return;
      const dt=new Date(it.created_at); if (isNaN(dt.getTime())) return;
      if (dt<fromUTC || dt>toUTC) return;

      const cv = it.column_values || [];
      const srcText = (cv.find(c=>c.id===sourceId)||{}).text || '';
      const typText = (cv.find(c=>c.id===typeId)  ||{}).text || '';
      const stText  = (cv.find(c=>c.id===statusId)||{}).text || '';
      if (!SITE_textIncludesAny_(srcText,  SITE_MONDAY_LEADS_SOURCE_MATCH)) return;
      if (!SITE_textIncludesAny_(typText,  SITE_MONDAY_LEADS_FORM_MATCH))   return;
      if (!SITE_textIncludesAny_(stText,   SITE_MONDAY_LEADS_STATUS_MATCH)) return;

      out.push({id:it.id, name:it.name, created_at:dt.toISOString(), source:srcText, type:typText, status:stText});
    });

    cursor = page.cursor;
  } while (cursor);

  return out;
}

/** Runner de TEST : logs pour les 3 derniers mois complets */
function run_Site_TestLeads(){
  const today = new Date();
  const endLocal = new Date(today.getFullYear(), today.getMonth(), 0); // dernier jour du mois précédent
  const startLocal = new Date(endLocal.getFullYear(), endLocal.getMonth()-2, 1); // 1er jour il y a 2 mois
  const fromUTC = new Date(Date.UTC(startLocal.getFullYear(), startLocal.getMonth(), 1, 0,0,0));
  const toUTC   = new Date(Date.UTC(endLocal.getFullYear(), endLocal.getMonth(), endLocal.getDate(), 23,59,59));

  Logger.log('=== TEST LEADS — Fenêtre UTC ===');
  Logger.log('From: %s  To: %s', fromUTC.toISOString(), toUTC.toISOString());

  const callsByMonth = SITE_mondayLeadCallsCountsByMonth_(
    SITE_MONDAY_LEADS_BOARD_ID,
    SITE_MONDAY_LEADS_COL_SOURCE,
    SITE_MONDAY_LEADS_COL_TYPE,
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
  Object.keys(callsByMonth).sort().forEach(k=> Logger.log('%s : %s', k, callsByMonth[k]));
  Logger.log('--- Formulaires lead (par mois) ---');
  Object.keys(formsByMonth).sort().forEach(k=> Logger.log('%s : %s', k, formsByMonth[k]));

  const callItems  = SITE_mondayFetchLeadCallsItems_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_SOURCE, SITE_MONDAY_LEADS_COL_TYPE, fromUTC, toUTC);
  const formItems  = SITE_mondayFetchLeadFormsItems_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_SOURCE, SITE_MONDAY_LEADS_COL_TYPE, SITE_MONDAY_LEADS_COL_STATUS, fromUTC, toUTC);

  Logger.log('=== Détails — APPELS lead (%s items) ===', callItems.length);
  callItems.forEach(it=>{
    Logger.log('[Call] %s | %s | %s | src="%s" | type="%s"', it.id, it.created_at, it.name, it.source, it.type);
  });

  Logger.log('=== Détails — FORMULAIRES lead (%s items) ===', formItems.length);
  formItems.forEach(it=>{
    Logger.log('[Form] %s | %s | %s | src="%s" | type="%s" | status="%s"', it.id, it.created_at, it.name, it.source, it.type, it.status);
  });

  Logger.log('=== FIN TEST LEADS ===');
}

// Parse JSON de colonne Date Monday: {"date":"YYYY-MM-DD", ...} -> Date UTC 00:00
function SITE_parseMondayDateValue_(val){
  if (!val) return null;
  try{
    const j = JSON.parse(val);
    if (j && j.date){
      const [Y,M,D] = String(j.date).split('-').map(Number);
      const d = new Date(Date.UTC(Y, (M||1)-1, D||1, 0,0,0));
      return isNaN(d.getTime()) ? null : d;
    }
  }catch(e){}
  return null;
}

// Cherche text & value d'une colonne par id dans column_values
function SITE_cv_(cv, id){
  const c = (cv||[]).find(x=>x.id===id);
  return {
    text:  (c && c.text)  || '',
    value: (c && c.value) || ''
  };
}

// bool includes any (case/accents insensitive)
function SITE_includesAny_(txt, arr){ return SITE_textIncludesAny_(txt, arr); }

// DEBUG: inspecte items et explique pourquoi inclus/exclus
function run_Site_TestLeads_Debug(){
  const today = new Date();
  const endLocal = new Date(today.getFullYear(), today.getMonth(), 0);
  const startLocal = new Date(endLocal.getFullYear(), endLocal.getMonth()-2, 1);
  const fromUTC = new Date(Date.UTC(startLocal.getFullYear(), startLocal.getMonth(), 1, 0,0,0));
  const toUTC   = new Date(Date.UTC(endLocal.getFullYear(), endLocal.getMonth(), endLocal.getDate(), 23,59,59));

  Logger.log('=== DEBUG LEADS — Fenêtre ===');
  Logger.log('From: %s  To: %s', fromUTC.toISOString(), toUTC.toISOString());

  const sourceId = SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_SOURCE);
  const typeId   = SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_TYPE);
  const statusId = SITE_MONDAY_LEADS_COL_STATUS ? SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_STATUS) : null;
  const dateId   = SITE_MONDAY_LEADS_COL_DATE   ? SITE_mondayResolveColumnId_(SITE_MONDAY_LEADS_BOARD_ID, SITE_MONDAY_LEADS_COL_DATE)   : null;

  if (!sourceId || !typeId) throw new Error('Colonnes source/type introuvables (vérifie les intitulés).');

  let cursor=null, items=[];
  do{
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
    const colIds = [sourceId, typeId].concat(statusId? [statusId]: []).concat(dateId? [dateId]: []);
    const d=SITE_mondayGraphQL_(q,{bid:[Number(SITE_MONDAY_LEADS_BOARD_ID)],cursor, cols:colIds});
    const page=d && d.boards && d.boards[0] && d.boards[0].items_page; if(!page) break;

    (page.items||[]).forEach(it=>{
      const createdAt = new Date(it.created_at);
      const cv = it.column_values||[];
      const src = SITE_cv_(cv, sourceId);
      const typ = SITE_cv_(cv,   typeId);
      const sts = statusId ? SITE_cv_(cv, statusId) : {text:'', value:''};
      const dcv = dateId   ? SITE_cv_(cv, dateId)   : {text:'', value:''};

      const dateFromCol = SITE_parseMondayDateValue_(dcv.value);
      const when = dateFromCol || createdAt;

      const inWindow = (when>=fromUTC && when<=toUTC);
      const srcOK = SITE_includesAny_(src.text, SITE_MONDAY_LEADS_SOURCE_MATCH);
      const typeIsCall = SITE_includesAny_(typ.text, SITE_MONDAY_LEADS_TYPE_MATCH);
      const typeIsForm = SITE_includesAny_(typ.text, SITE_MONDAY_LEADS_FORM_MATCH||[]);
      const statusOK   = !statusId || SITE_includesAny_(sts.text, SITE_MONDAY_LEADS_STATUS_MATCH||[]);

      const isLeadCall = inWindow && srcOK && typeIsCall;
      const isLeadForm = inWindow && srcOK && typeIsForm && statusOK;

      const reasons=[];
      if (!inWindow) reasons.push('hors fenêtre');
      if (!srcOK)    reasons.push('source !match');
      if (!(typeIsCall||typeIsForm)) reasons.push('type !match');
      if (typeIsForm && !statusOK) reasons.push('statut !match');

      items.push({
        id: it.id,
        name: it.name,
        created_at: isNaN(createdAt.getTime())?'':createdAt.toISOString(),
        date_col: dateFromCol? dateFromCol.toISOString() : '',
        when_used: when.toISOString(),
        src: src.text, typ: typ.text, sts: sts.text,
        isLeadCall, isLeadForm,
        exclude: reasons.join(', ')
      });
    });

    cursor=page.cursor;
  } while(cursor);

  const aggCalls={}, aggForms={};
  items.forEach(x=>{
    const ym = x.when_used.slice(0,7);
    if (x.isLeadCall) aggCalls[ym]=(aggCalls[ym]||0)+1;
    if (x.isLeadForm) aggForms[ym]=(aggForms[ym]||0)+1;
  });

  Logger.log('--- Totaux APPELS lead par mois ---');
  Object.keys(aggCalls).sort().forEach(k=> Logger.log('%s : %s', k, aggCalls[k]));
  Logger.log('--- Totaux FORMULAIRES lead par mois ---');
  Object.keys(aggForms).sort().forEach(k=> Logger.log('%s : %s', k, aggForms[k]));

  Logger.log('=== Détails (dans la fenêtre) — candidats APPELS ===');
  items.filter(x=> x.when_used>=fromUTC.toISOString() && x.when_used<=toUTC.toISOString() && (x.typ||'').length)
       .filter(x=> SITE_includesAny_(x.typ, SITE_MONDAY_LEADS_TYPE_MATCH))
       .forEach(x=> Logger.log('[CALL?] %s | %s | when=%s | src="%s" type="%s" status="%s" | lead=%s | excl=%s',
                               x.id, x.name, x.when_used, x.src, x.typ, x.sts, x.isLeadCall, x.exclude));

  Logger.log('=== Détails (dans la fenêtre) — candidats FORMULAIRES ===');
  items.filter(x=> x.when_used>=fromUTC.toISOString() && x.when_used<=toUTC.toISOString() && (x.typ||'').length)
       .filter(x=> SITE_includesAny_(x.typ, SITE_MONDAY_LEADS_FORM_MATCH||[]))
       .forEach(x=> Logger.log('[FORM?] %s | %s | when=%s | src="%s" type="%s" status="%s" | lead=%s | excl=%s',
                               x.id, x.name, x.when_used, x.src, x.typ, x.sts, x.isLeadForm, x.exclude));

  Logger.log('=== FIN DEBUG LEADS ===');
}

function run_Site_CurrentMonth(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_SITE);
  if (!sh) throw new Error(`Onglet '${SHEET_NAME_SITE}' introuvable`);
  const cols = SITE_findCols_(sh);
  if (!cols.mois) throw new Error("Colonne 'Mois' introuvable (ligne d’entêtes).");

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);              // 1er jour du mois en cours
  const end   = new Date(today.getFullYear(), today.getMonth(), today.getDate()); // aujourd'hui
  const ymKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM');
  const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const endStr   = Utilities.formatDate(end,   Session.getScriptTimeZone(), 'yyyy-MM-dd');

  SITE_toast_(`Site: mise à jour du mois en cours ${SITE_monthKeyToFr_(ymKey)}…`,'Site Internet',4);

  // GA4 & GSC sur la plage [1er du mois -> aujourd'hui]
  const ga  = SITE_executeWithRetry_(()=>ga4FetchMonthly_(GA4_PROPERTY, startStr, endStr),'GA4 mois en cours');
  const gsc = SITE_executeWithRetry_(()=>gscFetchMonthly_(GSC_SITE_URL, startStr, endStr),'GSC mois en cours');

  // Appels Magnetis (filtrés "Tout le SEO")
  const callsByMonth = SITE_executeWithRetry_(
    () => SITE_magnetisAggregateMonthlyCounts_(SITE_magnetisFetchCalls_(start, end)),
    'Magnetis mois en cours'
  );

  // Formulaires (Paperform + Monday)
  const formsByMonth = SITE_executeWithRetry_(
    () => SITE_formsCountsByMonth_(
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(),   end.getMonth(),   end.getDate(), 23,59,59))
    ),
    'Forms mois en cours'
  );

  // Leads Monday : appels + formulaires
  const leadCallsByMonth = SITE_executeWithRetry_(
    () => SITE_mondayLeadCallsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(),   end.getMonth(),   end.getDate(), 23,59,59))
    ),
    'Monday lead calls mois en cours'
  );

  const leadFormsByMonth = SITE_executeWithRetry_(
    () => SITE_mondayLeadFormsCountsByMonth_(
      SITE_MONDAY_LEADS_BOARD_ID,
      SITE_MONDAY_LEADS_COL_SOURCE,
      SITE_MONDAY_LEADS_COL_TYPE,
      SITE_MONDAY_LEADS_COL_STATUS,
      new Date(Date.UTC(start.getFullYear(), start.getMonth(), 1)),
      new Date(Date.UTC(end.getFullYear(),   end.getMonth(),   end.getDate(), 23,59,59))
    ),
    'Monday lead forms mois en cours'
  );

  // Position moyenne Check Position sur le mois (rapport mensuel)
  const avgPos = SITE_executeWithRetry_(
    () => SITE_cpMonthlyAveragePosition_(SITE_CP_DOMAIN, ymKey, null),
    'CheckPosition mois en cours'
  );

  const values = {
    sessions:    ga[ymKey]?.sessions ?? null,
    avgSec:      ga[ymKey]?.avgSec ?? null,
    impressions: gsc[ymKey]?.impressions ?? null,
    calls:       callsByMonth[ymKey] ?? null,
    forms:       formsByMonth[ymKey] ?? null,
    leadCalls:   leadCallsByMonth[ymKey] ?? null,
    leadForms:   leadFormsByMonth[ymKey] ?? null,
    avgPos:      avgPos ?? null,
    bounce:      ga[ymKey]?.bouncePct ?? null
  };

  // ⚠️ SITE_writeMonthRow_ gère déjà :
  //  - création / insertion de la ligne du mois si elle n'existe pas
  //  - écriture UNIQUEMENT dans les colonnes autorisées (pas les formules)
  SITE_writeMonthRow_(sh, cols, ymKey, values);

  SITE_toast_('Site: mois en cours mis à jour ✅','Site Internet',5);
}
