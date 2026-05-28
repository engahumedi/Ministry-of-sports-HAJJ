// ════════════════════════════════════════════════════════════
// نظام مراقبة التفويج — Google Apps Script
// الإعداد المطلوب في Script Properties:
//   MON_PASS   → كلمة مرور المراقبين والمتابعين
//   ADMIN_PASS → كلمة مرور مشرف النظام
//   ADMIN_EMAIL→ بريد مشرف النظام (بالأحرف الصغيرة)
//   SHEET_ID   → (اختياري) يُنشأ تلقائياً عند أول تشغيل
// ════════════════════════════════════════════════════════════

function getCfg() {
  var p = PropertiesService.getScriptProperties();
  return {
    MON_PASS:    p.getProperty('MON_PASS')    || '',
    ADMIN_PASS:  p.getProperty('ADMIN_PASS')  || '',
    ADMIN_EMAIL: (p.getProperty('ADMIN_EMAIL') || '').toLowerCase().trim(),
    SHEET_ID:    p.getProperty('SHEET_ID')    || ''
  };
}

function getSheet() {
  var p   = PropertiesService.getScriptProperties();
  var sid = p.getProperty('SHEET_ID');
  if (sid) return SpreadsheetApp.openById(sid).getSheets()[0];
  var ss = SpreadsheetApp.create('بلاغات التفويج 1447');
  sid = ss.getId();
  p.setProperty('SHEET_ID', sid);
  ss.getSheets()[0].appendRow([
    'id','ts','monitorEmail','zone','camp',
    'violation','notes','shift','status','lat','lng','photo'
  ]);
  return ss.getSheets()[0];
}

function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRows(sh) {
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var hdr = vals[0];
  return vals.slice(1).map(function(row) {
    var o = {};
    hdr.forEach(function(h, i) { o[h] = row[i]; });
    o.id = String(o.id); // prevent float notation for large timestamps
    return o;
  });
}

// ════════════════════════════════════════
// MAIN HANDLER
// ════════════════════════════════════════
function doGet(e) {
  var cfg    = getCfg();
  var d      = e.parameter;
  var action = d.action || 'list';
  var pass   = d.pass   || '';

  // ── تحقق من هوية المستخدم (مراقب/متابع/أدمن) ────────────
  if (action === 'validateUser') {
    var email = (d.email || '').toLowerCase().trim();
    // أدمن
    if (cfg.ADMIN_PASS && email === cfg.ADMIN_EMAIL && pass === cfg.ADMIN_PASS) {
      return resp({ok: true, role: 'admin', token: pass});
    }
    // مراقب أو متابع
    if (cfg.MON_PASS && pass === cfg.MON_PASS) {
      return resp({ok: true, role: 'user', token: pass});
    }
    return resp({ok: false});
  }

  // ── إضافة بلاغ ────────────────────────────────────────────
  if (action === 'add') {
    if (!cfg.MON_PASS || pass !== cfg.MON_PASS) return resp({ok:false, err:'unauthorized'});
    var sh   = getSheet();
    var rows = getRows(sh);
    var rid  = String(d.id || new Date().getTime());
    // منع التكرار
    if (rows.some(function(r){ return r.id === rid; })) {
      return resp({ok: true, dup: true});
    }
    sh.appendRow([
      rid,
      d.ts           || new Date().toISOString(),
      d.monitorEmail || '',
      d.zone         || '',
      d.camp         || '',
      d.violation    || '',
      d.notes        || '',
      d.shift        || '',
      d.status       || 'مفتوح',
      d.lat          || '',
      d.lng          || '',
      d.photo        || ''
    ]);
    return resp({ok: true, id: rid});
  }

  // ── قائمة البلاغات (مراقب أو أدمن) ───────────────────────
  if (action === 'list') {
    var isAdmin = cfg.ADMIN_PASS && pass === cfg.ADMIN_PASS;
    var isMon   = cfg.MON_PASS   && pass === cfg.MON_PASS;
    if (!isAdmin && !isMon) return resp({ok:false, err:'unauthorized'});
    var rows = getRows(getSheet());
    // المراقب يرى بلاغاته فقط
    if (isMon && d.email) {
      rows = rows.filter(function(r){ return r.monitorEmail === d.email; });
    }
    return resp({ok: true, rows: rows});
  }

  // ── قائمة البلاغات حسب المناطق (المتابع) ─────────────────
  if (action === 'listByZones') {
    if (!cfg.MON_PASS || pass !== cfg.MON_PASS) return resp({ok:false, err:'unauthorized'});
    var zones = (d.zones || '').split(',').map(function(z){ return z.trim(); });
    var rows  = getRows(getSheet()).filter(function(r){ return zones.indexOf(r.zone) !== -1; });
    return resp({ok: true, rows: rows});
  }

  // ── تحديث حالة البلاغ ─────────────────────────────────────
  if (action === 'updateStatus') {
    var isAdmin = cfg.ADMIN_PASS && pass === cfg.ADMIN_PASS;
    var isMon   = cfg.MON_PASS   && pass === cfg.MON_PASS;
    if (!isAdmin && !isMon) return resp({ok:false, err:'unauthorized'});
    var sh  = getSheet();
    var all = sh.getDataRange().getValues();
    var hdr = all[0];
    var ci  = hdr.indexOf('id');
    var cs  = hdr.indexOf('status');
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][ci]) === String(d.id)) {
        sh.getRange(i + 1, cs + 1).setValue(d.status || 'مفتوح');
        return resp({ok: true});
      }
    }
    return resp({ok: false, err: 'not found'});
  }

  // ── حذف بلاغ ──────────────────────────────────────────────
  if (action === 'deleteReport') {
    var isAdmin = cfg.ADMIN_PASS && pass === cfg.ADMIN_PASS;
    var isMon   = cfg.MON_PASS   && pass === cfg.MON_PASS;
    if (!isAdmin && !isMon) return resp({ok:false, err:'unauthorized'});
    var sh  = getSheet();
    var all = sh.getDataRange().getValues();
    var ci  = all[0].indexOf('id');
    for (var i = all.length - 1; i >= 1; i--) {
      if (String(all[i][ci]) === String(d.id)) {
        sh.deleteRow(i + 1);
        return resp({ok: true});
      }
    }
    return resp({ok: false, err: 'not found'});
  }

  return resp({ok: false, err: 'unknown action'});
}
