const CONFIG = {
  spreadsheetId: '', // ضع هنا معرّف Google Sheet
  timezone: 'Asia/Riyadh',
  sheets: {
    employees: 'Employees',
    absences: 'Absences',
    lateness: 'Lateness',
    signatures: 'Signatures',
    settings: 'Settings'
  }
};

function doGet(e) {
  return json_({ ok:true, service:'IbnHisham Attendance API', time:new Date().toISOString() });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (!CONFIG.spreadsheetId) return json_({ok:false,error:'لم يتم ضبط spreadsheetId في Code.gs'});
    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const action = body.action || 'save';
    const map = {
      employee: CONFIG.sheets.employees,
      absence: CONFIG.sheets.absences,
      lateness: CONFIG.sheets.lateness,
      signature: CONFIG.sheets.signatures,
      settings: CONFIG.sheets.settings
    };
    const sheetName = map[action];
    if (!sheetName) throw new Error('Action غير معروف: '+action);
    const sh = getOrCreate_(ss, sheetName);
    const record = body.data || {};
    const headers = Object.keys(record);
    ensureHeaders_(sh, headers);
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0];
    sh.appendRow(current.map(h => record[h] === undefined ? '' : record[h]));
    return json_({ok:true,saved:true,action});
  } catch(err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function getOrCreate_(ss,name) {
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sh, headers) {
  if (!headers.length) return;
  if (sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else if (sh.getLastColumn()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupSheets() {
  if (!CONFIG.spreadsheetId) throw new Error('ضع spreadsheetId أولاً');
  const ss=SpreadsheetApp.openById(CONFIG.spreadsheetId);
  Object.keys(CONFIG.sheets).forEach(k=>getOrCreate_(ss,CONFIG.sheets[k]));
}
