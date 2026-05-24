var ADMIN_PASS = 'hajj1447';

function getSheet() {
  var props = PropertiesService.getScriptProperties();
  var sid = props.getProperty('SHEET_ID');
  if (!sid) {
    var ss = SpreadsheetApp.create('حضور معسكر المعيصم 1447');
    sid = ss.getId();
    props.setProperty('SHEET_ID', sid);
    SpreadsheetApp.openById(sid).getActiveSheet().appendRow(
      ['_id','الاسم','الهوية','التاريخ','الوقت','خط العرض','خط الطول','الدقة']
    );
  }
  return SpreadsheetApp.openById(sid).getActiveSheet();
}

function doPost(e) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('REG_OPEN') === 'false') {
      return ContentService.createTextOutput('closed');
    }
    var sheet = getSheet();
    var rec = JSON.parse(e.postData.contents);
    var row = [
      new Date().getTime(),
      rec['name'] || '',
      rec['id'] || '',
      rec['date'] || '',
      rec['time'] || '',
      rec['lat'] || '',
      rec['lng'] || '',
      rec['accuracy'] || ''
    ];
    sheet.appendRow(row);
    return ContentService.createTextOutput('ok');
  } catch(err) {
    return ContentService.createTextOutput('error:' + err.message);
  }
}

function doGet(e) {
  var action = e.parameter.action || 'list';
  var props  = PropertiesService.getScriptProperties();

  if (action === 'status') {
    var isOpen = props.getProperty('REG_OPEN') !== 'false';
    return out({open: isOpen});
  }

  var pass = e.parameter.pass || '';
  if (pass !== ADMIN_PASS) {
    return out({error: 'unauthorized'});
  }

  var sheet = getSheet();

  if (action === 'list') {
    var rows = sheet.getDataRange().getValues();
    var recs = [];
    rows.slice(1).forEach(function(r) {
      recs.push({
        _id: r[0], name: r[1], id: r[2],
        date: r[3], time: r[4],
        lat: r[5], lng: r[6], accuracy: r[7]
      });
    });
    return out(recs);
  }

  if (action === 'setStatus') {
    var val = e.parameter.value;
    props.setProperty('REG_OPEN', val === 'open' ? 'true' : 'false');
    return out({status: 'ok', open: val === 'open'});
  }

  if (action === 'delete') {
    var tid = e.parameter.id;
    var all = sheet.getDataRange().getValues();
    all.slice(1).forEach(function(r, idx) {
      if (String(r[0]) === String(tid)) {
        sheet.deleteRow(idx + 2);
      }
    });
    return out({status: 'ok'});
  }

  if (action === 'clear') {
    var last = sheet.getLastRow();
    if (last !== 1) {
      sheet.deleteRows(2, last - 1);
    }
    return out({status: 'ok'});
  }

  return out({error: 'unknown'});
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
