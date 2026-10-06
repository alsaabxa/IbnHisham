const CONFIG = {
  spreadsheetId: '1W8yxCurpAbKqfSO6q3GNUilFxPCGQQ6GXU9ixD6gl3I',
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
  try {
    const action=(e&&e.parameter&&e.parameter.action)||'ping';
    if(!CONFIG.spreadsheetId) return json_({ok:false,error:'لم يتم ضبط spreadsheetId في Code.gs'});
    if(action==='ping') return json_({ok:true,service:'IbnHisham Attendance API',message:'الاتصال يعمل بنجاح',time:new Date().toISOString()});
    const ss=SpreadsheetApp.openById(CONFIG.spreadsheetId);
    if(action==='all') return json_({ok:true,data:getAllData_(ss)});
    if(action==='employees') return json_({ok:true,data:readSheet_(ss,CONFIG.sheets.employees)});
    if(action==='absences') return json_({ok:true,data:readSheet_(ss,CONFIG.sheets.absences)});
    if(action==='lateness'||action==='lates') return json_({ok:true,data:readSheet_(ss,CONFIG.sheets.lateness)});
    if(action==='settings') return json_({ok:true,data:readSettings_(ss)});
    throw new Error('Action غير معروف: '+action);
  } catch(err){ return json_({ok:false,error:String(err.message||err)}); }
}

function doPost(e){
  try{
    const body=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    if(!CONFIG.spreadsheetId) return json_({ok:false,error:'لم يتم ضبط spreadsheetId في Code.gs'});
    const ss=SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const action=body.action||'save';
    if(action==='sync'){
      const d=body.data||{};
      return json_({ok:true,synced:true,result:{
        employees:upsertMany_(ss,CONFIG.sheets.employees,d.employees||[]),
        absences:upsertMany_(ss,CONFIG.sheets.absences,d.absences||[]),
        lateness:upsertMany_(ss,CONFIG.sheets.lateness,d.lates||d.lateness||[]),
        signatures:upsertSignatures_(ss,d.signatures||{}),
        settings:upsertSettings_(ss,d.settings||{})
      }});
    }
    const map={employee:CONFIG.sheets.employees,absence:CONFIG.sheets.absences,lateness:CONFIG.sheets.lateness,late:CONFIG.sheets.lateness,signature:CONFIG.sheets.signatures};
    if(action==='settings'){upsertSettings_(ss,body.data||{});return json_({ok:true,saved:true,action:action});}
    if(action==='delete'){
      const d=body.data||{};
      const type=String(d.type||'');
      const id=String(d.id||'');
      const sheetName=type==='absence'?CONFIG.sheets.absences:(type==='late'||type==='lateness'?CONFIG.sheets.lateness:'');
      if(!sheetName||!id) throw new Error('بيانات الحذف غير مكتملة');
      const deleted=deleteById_(ss,sheetName,id);
      return json_({ok:deleted,deleted:deleted,action:'delete',type:type,id:id,error:deleted?'':('السجل غير موجود: '+id)});
    }
    const sheetName=map[action];
    if(!sheetName) throw new Error('Action غير معروف: '+action);
    upsertOne_(ss,sheetName,body.data||{});
    return json_({ok:true,saved:true,action:action});
  }catch(err){return json_({ok:false,error:String(err.message||err)});}
}

function setupSheets(){
  if(!CONFIG.spreadsheetId) throw new Error('ضع spreadsheetId أولاً');
  const ss=SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const headers={
    Employees:['id','name','job','createdAt'],
    Absences:['id','empId','name','job','type','month','from','to','days','ref','reportDate','place','note','created'],
    Lateness:['id','empId','name','job','date','arrival','start','minutes','reason','note','created'],
    Signatures:['id','empId','name','type','data','at','createdAt'],
    Settings:['key','value','updatedAt']
  };
  Object.keys(headers).forEach(k=>{const sh=getOrCreate_(ss,k);ensureHeaders_(sh,headers[k]);});
  return 'تم تهيئة جداول منصة ابن هشام';
}

function getAllData_(ss){
  return {
    employees:readSheet_(ss,CONFIG.sheets.employees),
    absences:readSheet_(ss,CONFIG.sheets.absences),
    lates:readSheet_(ss,CONFIG.sheets.lateness),
    signatures:readSignatures_(ss),
    settings:readSettings_(ss)
  };
}

function readSheet_(ss,name){
  const sh=getOrCreate_(ss,name),values=sh.getDataRange().getValues();
  if(!values.length)return[];
  const headers=values[0].map(String);
  return values.slice(1).filter(r=>r.some(v=>v!=='')).map(r=>{const o={};headers.forEach((h,i)=>o[h]=normalize_(r[i]));return o;});
}
function readSignatures_(ss){
  const rows=readSheet_(ss,CONFIG.sheets.signatures),out={};
  rows.forEach(r=>{if(r.id)out[r.empId||r.id]={data:r.data,type:r.type,at:r.at};});
  return out;
}
function readSettings_(ss){
  const rows=readSheet_(ss,CONFIG.sheets.settings),out={};
  rows.forEach(r=>{if(r.key)out[r.key]=r.value;});
  return out;
}
function upsertMany_(ss,name,records){
  let n=0;(records||[]).forEach(r=>{if(r&&(r.id||r.key)){upsertOne_(ss,name,r);n++;}});return n;
}
function upsertSignatures_(ss,obj){
  let n=0;Object.keys(obj||{}).forEach(id=>{const s=obj[id]||{};upsertOne_(ss,CONFIG.sheets.signatures,{id:id,empId:id,name:s.name||'',type:s.type||'',data:s.data||'',at:s.at||'',createdAt:s.createdAt||s.at||''});n++;});return n;
}
function deleteById_(ss,name,id){
  const sh=getOrCreate_(ss,name);
  const headers=getHeaders_(sh);
  const keyField=name==='Settings'?'key':'id';
  const row=findRow_(sh,keyField,String(id),headers);
  if(!row)return false;
  sh.deleteRow(row);
  return true;
}
function upsertOne_(ss,name,record){
  if(!record||(!record.id&&!record.key))return;
  const sh=getOrCreate_(ss,name);
  const preferred={
    Employees:['id','name','job','createdAt'],
    Absences:['id','empId','name','job','type','month','from','to','days','ref','reportDate','place','note','created'],
    Lateness:['id','empId','name','job','date','arrival','start','minutes','reason','note','created'],
    Signatures:['id','empId','name','type','data','at','createdAt'],
    Settings:['key','value','updatedAt']
  }[name]||Object.keys(record);
  ensureHeaders_(sh,preferred);
  const headers=getHeaders_(sh),keyField=name==='Settings'?'key':'id',key=String(record[keyField]||'');
  if(!key)return;
  const row=findRow_(sh,keyField,key,headers),values=headers.map(h=>record[h]===undefined?'':normalize_(record[h]));
  if(row)sh.getRange(row,1,1,headers.length).setValues([values]);
  else sh.getRange(sh.getLastRow()+1,1,1,headers.length).setValues([values]);
}
function upsertSettings_(ss,settings){
  Object.keys(settings||{}).forEach(k=>upsertOne_(ss,CONFIG.sheets.settings,{key:k,value:settings[k],updatedAt:new Date().toISOString()}));
  return Object.keys(settings||{}).length;
}
function findRow_(sh,keyField,key,headers){
  const col=headers.indexOf(keyField)+1;if(col<1||sh.getLastRow()<2)return 0;
  const vals=sh.getRange(2,col,sh.getLastRow()-1,1).getValues();
  for(let i=0;i<vals.length;i++)if(String(vals[i][0])===key)return i+2;
  return 0;
}
function getHeaders_(sh){return sh.getLastColumn()<1?[]:sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);}
function ensureHeaders_(sh,headers){
  if(!headers.length)return;
  if(sh.getLastRow()===0||sh.getLastColumn()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);
}
function getOrCreate_(ss,name){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);return sh;}
function normalize_(v){return v instanceof Date?Utilities.formatDate(v,CONFIG.timezone,"yyyy-MM-dd'T'HH:mm:ss"):v;}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function onOpen(){
  SpreadsheetApp.getUi().createMenu('منصة ابن هشام').addItem('تهيئة الجداول','setupSheets').addItem('اختبار الاتصال','testConnection').addToUi();
}
function testConnection(){SpreadsheetApp.getUi().alert('الاتصال يعمل — منصة ابن هشام');}
