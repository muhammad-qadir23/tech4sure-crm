/**
 * Tech4Sure — Team database + Prospect Hunter backend
 * Runs inside a Google Sheet (Extensions → Apps Script). Free.
 *
 * SETUP (see SETUP-PROSPECTING-AND-SYNC.md):
 * 1. Script Properties: PLACES_KEY = your Google Places API key
 *                       TEAM_KEY   = any password you invent (members enter it in the app)
 * 2. Deploy → New deployment → Web app → Execute as: Me → Access: Anyone
 * 3. Paste the /exec URL into the app's Settings → Team Sync.
 */

var LEAD_COLS = ['id','updatedAt','rep','name','type','address','phone','contact','status',
  'reviewUrl','placeId','nfc','hasWebsite','trialStart','followUp','reviewsStart','reviewsNow',
  'plan','intro','customPrice','ghl','notes','visitedAt','createdAt','lat','lng'];
var PROS_COLS = ['placeId','name','address','phone','website','rating','reviews','type',
  'lat','lng','score','reasons','area','foundAt'];

function doGet() {
  return json_({ ok: true, msg: 'Tech4Sure backend is running. Use the app to talk to me.' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty('TEAM_KEY') || req.key !== props.getProperty('TEAM_KEY'))
      return json_({ ok: false, error: 'Wrong team key' });
    ensure_();
    if (req.biz && req.biz.length > 5000) return json_({ ok: false, error: 'Payload too large' });
    if (req.action === 'push')      return json_({ ok: true, biz: push_(req.biz || [], String(req.rep || '').slice(0, 60)) });
    if (req.action === 'pull')      return json_({ ok: true, biz: push_([], '') });
    if (req.action === 'hunt')      return json_(hunt_(req.q, props.getProperty('PLACES_KEY')));
    if (req.action === 'prospects') return json_({ ok: true, prospects: readAll_('Prospects') });
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- helpers ---------- */
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function sheet_(n) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(n) || ss.insertSheet(n);
}
function ensure_() {
  [['Leads', LEAD_COLS], ['Prospects', PROS_COLS]].forEach(function (t) {
    var sh = sheet_(t[0]);
    if (sh.getLastRow() === 0) sh.appendRow(t[1]);
  });
}
function readAll_(name) {
  var sh = sheet_(name);
  if (sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues();
  var head = v[0];
  return v.slice(1).map(function (r) {
    var o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  }).filter(function (o) { return o[head[0]]; });
}

/* ---------- leads: merge-by-newest sync ---------- */
function push_(incoming, rep) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var map = {};
    readAll_('Leads').forEach(function (b) { map[b.id] = b; });
    incoming.forEach(function (b) {
      if (!b.id) return;
      b.rep = b.rep || rep;
      var ex = map[b.id];
      if (!ex || String(b.updatedAt || '') > String(ex.updatedAt || '')) map[b.id] = b;
    });
    var merged = Object.keys(map).map(function (k) {
      var b = map[k];
      b.intro = (b.intro === true || b.intro === 'true' || b.intro === 'TRUE' || b.intro == 1) ? 1 : 0;
      return b;
    });
    var sh = sheet_('Leads');
    sh.clearContents();
    sh.appendRow(LEAD_COLS);
    if (merged.length) {
      var rng = sh.getRange(2, 1, merged.length, LEAD_COLS.length);
      rng.setNumberFormat('@'); // keep dates/ids as plain text
      rng.setValues(merged.map(function (b) {
        return LEAD_COLS.map(function (c) { return b[c] !== undefined && b[c] !== null ? String(b[c]) : ''; });
      }));
    }
    return merged;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- prospect hunter: Places API + scoring ---------- */
var MAX_HUNTS_PER_DAY = 80; // protects your Google bill even if the team key leaks

function hunt_(q, key) {
  if (!key) return { ok: false, error: 'PLACES_KEY not set in Script Properties' };
  if (!q)   return { ok: false, error: 'Empty search' };
  q = String(q).slice(0, 200);
  // daily rate cap
  var props = PropertiesService.getScriptProperties();
  var todayStr = new Date().toISOString().slice(0, 10);
  if (props.getProperty('HUNT_DATE') !== todayStr) {
    props.setProperty('HUNT_DATE', todayStr);
    props.setProperty('HUNT_COUNT', '0');
  }
  var count = Number(props.getProperty('HUNT_COUNT') || 0);
  if (count >= MAX_HUNTS_PER_DAY)
    return { ok: false, error: 'Daily search limit reached (bill protection). Resets at midnight UTC.' };
  props.setProperty('HUNT_COUNT', String(count + 1));
  var fm = 'places.id,places.displayName,places.formattedAddress,places.location,' +
           'places.rating,places.userRatingCount,places.websiteUri,places.nationalPhoneNumber,' +
           'places.primaryType,places.businessStatus,nextPageToken';
  var out = [], token = '', pages = 0;
  do {
    var body = { textQuery: q, pageSize: 20 };
    if (token) body.pageToken = token;
    var res = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:searchText', {
      method: 'post', contentType: 'application/json',
      headers: { 'X-Goog-Api-Key': key, 'X-Goog-FieldMask': fm },
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    var d = JSON.parse(res.getContentText());
    if (d.error) return { ok: false, error: d.error.message };
    (d.places || []).forEach(function (p) { out.push(p); });
    token = d.nextPageToken || '';
    pages++;
    if (token) Utilities.sleep(1200);
  } while (token && pages < 3); // up to 60 businesses per hunt

  var leadPids = {};
  readAll_('Leads').forEach(function (b) { if (b.placeId) leadPids[b.placeId] = 1; });

  var scored = out.filter(function (p) {
    return p.businessStatus !== 'CLOSED_PERMANENTLY' && !leadPids[p.id];
  }).map(function (p) { return score_(p, q); });

  upsertPros_(scored);
  scored.sort(function (a, b) { return b.score - a.score; });
  return { ok: true, prospects: scored };
}

function score_(p, area) {
  var rc = p.userRatingCount || 0, rating = p.rating || 0, site = p.websiteUri || '';
  var s = 0, why = [];
  if      (rc < 10)  { s += 40; why.push('only ' + rc + ' reviews'); }
  else if (rc < 30)  { s += 32; why.push(rc + ' reviews'); }
  else if (rc < 75)  { s += 24; why.push(rc + ' reviews'); }
  else if (rc < 150) { s += 15; why.push(rc + ' reviews'); }
  else if (rc < 300) { s += 6;  why.push(rc + ' reviews'); }
  else               {          why.push(rc + ' reviews — established'); }
  if (!site) { s += 25; why.push('NO WEBSITE'); }
  if (!rating)            { s += 10; why.push('unrated'); }
  else if (rating < 3)    { s += 5;  why.push(rating + '★ struggling'); }
  else if (rating <= 4.4) { s += 15; why.push(rating + '★ needs a boost'); }
  else                    { s += 3;  why.push(rating + '★'); }
  var t = p.primaryType || '';
  if (/restaurant|cafe|coffee|barber|beauty|hair|gym|dental|dentist|auto|car_repair|salon|nail|spa|bakery|repair|laundry|pet/.test(t)) s += 10;
  return {
    placeId: p.id,
    name: p.displayName ? p.displayName.text : '',
    address: p.formattedAddress || '',
    phone: p.nationalPhoneNumber || '',
    website: site,
    rating: rating,
    reviews: rc,
    type: t.replace(/_/g, ' '),
    lat: p.location ? p.location.latitude : '',
    lng: p.location ? p.location.longitude : '',
    score: s,
    reasons: why.join(', '),
    area: area,
    foundAt: new Date().toISOString()
  };
}

function upsertPros_(list) {
  var sh = sheet_('Prospects');
  var existing = {};
  readAll_('Prospects').forEach(function (p) { existing[p.placeId] = 1; });
  var rows = list.filter(function (p) { return !existing[p.placeId]; })
    .map(function (p) { return PROS_COLS.map(function (c) { return p[c] !== undefined ? String(p[c]) : ''; }); });
  if (rows.length) {
    var rng = sh.getRange(sh.getLastRow() + 1, 1, rows.length, PROS_COLS.length);
    rng.setNumberFormat('@');
    rng.setValues(rows);
  }
}
