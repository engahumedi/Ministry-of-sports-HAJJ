
// MON_PASS: fallback to default so monitors work immediately
// ADMIN_PASS: must be set in Script Properties (Project Settings → Script Properties)
var _p         = PropertiesService.getScriptProperties();
var MON_PASS    = _p.getProperty('MON_PASS')   || 'Aa@12345678';
var ADMIN_PASS  = _p.getProperty('ADMIN_PASS');
var ADMIN_EMAIL = (_p.getProperty('ADMIN_EMAIL') || '').toLowerCase();
var SHEET_KEY   = 'TAFWEEJ_SHEET_ID';
var COLS       = ['id','ts','monitorEmail','zone','camp','violation','notes','shift','status','lat','lng','acc','photoNote'];

function getSheet() {
  var props = PropertiesService.getScriptProperties();
  var sid = props.getProperty(SHEET_KEY);
  if (!sid) {
    var ss = SpreadsheetApp.create('نظام التفويج 1447 — الوردية الثانية');
    sid = ss.getId();
    props.setProperty(SHEET_KEY, sid);
    ss.getActiveSheet().appendRow(COLS);
    ss.getActiveSheet().setFrozenRows(1);
  }
  return SpreadsheetApp.openById(sid).getActiveSheet();
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action !== 'add') return out({error:'unknown action'});

    var sheet = getSheet();
    var now   = new Date();
    // Use frontend-provided id/ts so local and server ids always match
    var id    = String(data.id || now.getTime());
    var ts    = data.ts || now.toISOString();

    // Prevent duplicate rows (same id already in sheet)
    var vals = sheet.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === id) return out({ok: true, id: id, dup: true});
    }

    sheet.appendRow([
      id,
      ts,
      data.monitorEmail || '',
      data.zone         || '',
      data.camp         || '',
      data.violation    || '',
      data.notes        || '',
      data.shift        || 'الوردية الثانية',
      'مفتوح',
      data.lat          || '',
      data.lng          || '',
      data.acc          || '',
      data.photo === '[image]' ? 'يوجد صورة' : ''
    ]);

    return out({ok: true, id: id});
  } catch(err) {
    return out({error: err.message});
  }
}

function doGet(e) {
  var action = (e.parameter.action || 'list').trim();
  var pass   = (e.parameter.pass   || '').trim();

  // ── validateAdmin (admin login — credentials verified here, never in frontend) ──
  if (action === 'validateAdmin') {
    var em = (e.parameter.email || '').toLowerCase();
    var pw = e.parameter.pass  || '';
    if (!ADMIN_PASS || !ADMIN_EMAIL) return out({error:'properties not configured'});
    if (em === ADMIN_EMAIL && pw === ADMIN_PASS) return out({ok: true});
    return out({error: 'unauthorized'});
  }

  // ── add via GET (avoids POST/CORS/redirect issues on mobile) ──
  if (action === 'add') {
    if (pass !== MON_PASS && pass !== ADMIN_PASS) return out({error:'unauthorized'});
    var sheet = getSheet();
    var now   = new Date();
    var id    = String(e.parameter.id || now.getTime());
    var ts    = e.parameter.ts || now.toISOString();

    // Prevent duplicate rows
    var vals = sheet.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === id) return out({ok: true, id: id, dup: true});
    }

    sheet.appendRow([
      id, ts,
      e.parameter.monitorEmail || '',
      e.parameter.zone         || '',
      e.parameter.camp         || '',
      e.parameter.violation    || '',
      e.parameter.notes        || '',
      e.parameter.shift        || 'الوردية الثانية',
      'مفتوح',
      e.parameter.lat          || '',
      e.parameter.lng          || '',
      e.parameter.acc          || '',
      e.parameter.photo === '[image]' ? 'يوجد صورة' : ''
    ]);
    return out({ok: true, id: id});
  }

  // ── list (monitor sees own, admin sees all) ──
  if (action === 'list') {
    if (pass !== MON_PASS && pass !== ADMIN_PASS) return out({error:'unauthorized'});
    var email = (e.parameter.email || '').toLowerCase();
    return out(getRows(function(r) {
      return pass === ADMIN_PASS ? true : r.monitorEmail.toLowerCase() === email;
    }));
  }

  // ── listByZones (supervisor) ──
  if (action === 'listByZones') {
    if (pass !== MON_PASS && pass !== ADMIN_PASS) return out({error:'unauthorized'});
    var zones = (e.parameter.zones || '').split(',').map(function(z){ return z.trim(); });
    return out(getRows(function(r){ return zones.indexOf(r.zone) !== -1; }));
  }

  // ── updateStatus (admin) ──
  if (action === 'updateStatus') {
    if (pass !== ADMIN_PASS) return out({error:'unauthorized'});
    var id     = e.parameter.id;
    var status = e.parameter.status;
    var sheet  = getSheet();
    var vals   = sheet.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(id)) {
        sheet.getRange(i + 1, 9).setValue(status);
        break;
      }
    }
    return out({ok: true});
  }

  // ── deleteReport (supervisor or admin) ──
  if (action === 'deleteReport') {
    if (pass !== MON_PASS && pass !== ADMIN_PASS) return out({error:'unauthorized'});
    var id    = e.parameter.id;
    var sheet = getSheet();
    var vals  = sheet.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return out({ok: true});
  }

  // ── stats (admin) ──
  if (action === 'stats') {
    if (pass !== ADMIN_PASS) return out({error:'unauthorized'});
    var rows = getRows(function(){ return true; });
    var byV  = {}, byM = {}, open = 0;
    rows.forEach(function(r) {
      byV[r.violation] = (byV[r.violation] || 0) + 1;
      byM[r.monitorEmail] = (byM[r.monitorEmail] || 0) + 1;
      if (r.status === 'مفتوح') open++;
    });
    return out({total: rows.length, open: open, byViolation: byV, byMonitor: byM});
  }

  return out({error:'unknown action'});
}

// ── helpers ──
function getRows(filterFn) {
  var sheet = getSheet();
  var vals  = sheet.getDataRange().getValues();
  var result = [];
  vals.slice(1).forEach(function(row) {
    var rec = {};
    COLS.forEach(function(col, i){ rec[col] = row[i] !== undefined ? row[i] : ''; });
    // Always return id as string to avoid float notation for large timestamps
    if (rec.id) rec.id = String(rec.id);
    if (rec.id && filterFn(rec)) result.push(rec);
  });
  return result;
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
