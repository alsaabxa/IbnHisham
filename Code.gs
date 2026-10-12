const CONFIG = {
  spreadsheetId: '1W8yxCurpAbKqfSO6q3GNUilFxPCGQQ6GXU9ixD6gl3I',
  timezone: 'Asia/Riyadh',
  sheets: {
    employees: 'الموظفات',
    absences: 'الغياب',
    lateness: 'التأخير',
    signatures: 'Signatures',
    settings: 'الإعدادات',
    actions: 'الإجراءات'
  }
};

function doGet(e) {
  try {
    const action=(e&&e.parameter&&e.parameter.action)||'ping';
    if(!CONFIG.spreadsheetId) return json_({ok:false,error:'لم يتم ضبط spreadsheetId في Code.gs'});
    if(action==='ping') return json_({ok:true,service:'IbnHisham Attendance API',message:'الاتصال يعمل بنجاح',time:new Date().toISOString()});
    const ss=SpreadsheetApp.openById(CONFIG.spreadsheetId);
    if(action==='all') return json_({ok:true,data:getAllData_(ss)});
    if(action==='employeeReport'){
      if(String((e.parameter&&e.parameter.shared)||'')==='1'&&!sharedSigningEnabled_())return json_({ok:false,error:'تم إيقاف رابط التوقيع الموحّد مؤقتًا من الإدارة. يرجى المحاولة لاحقًا.'});
      return json_(employeeReport_(ss, String((e.parameter&&e.parameter.token)||'')));
    }
    if(action==='portalSignatures') return json_({ok:true,data:listPortalSignatures_(ss)});
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
    if(action==='generateEmployeeAccessCodes'){
      const d=body.data||{};
      if(!portalAdminAuthorized_(d.adminPin))return json_({ok:false,error:'الرقم الإداري غير مضبوط أو غير صحيح. اضبطي PORTAL_ADMIN_PIN في خصائص المشروع.'});
      const codes=ensureEmployeeAccessCodes_(ss,String(d.portalBaseUrl||''));
      return json_({ok:true,codes:codes,sharedUrl:portalLinkBase_(d.portalBaseUrl)+'?sign=1',message:'تمت إضافة الرموز الناقصة مع الحفاظ على الرموز الحالية.'});
    }
    if(action==='listEmployeeAccessCodes'){
      const d=body.data||{};
      if(!portalAdminAuthorized_(d.adminPin))return json_({ok:false,error:'الرقم الإداري غير مضبوط أو غير صحيح. اضبطي PORTAL_ADMIN_PIN في خصائص المشروع.'});
      return json_({ok:true,codes:listEmployeeAccessCodes_(ss)});
    }
    if(action==='updateEmployeeAccessCode'){
      const d=body.data||{};
      if(!portalAdminAuthorized_(d.adminPin))return json_({ok:false,error:'الرقم الإداري غير مضبوط أو غير صحيح. اضبطي PORTAL_ADMIN_PIN في خصائص المشروع.'});
      return json_(updateEmployeeAccessCode_(ss,d));
    }
    if(action==='getSharedSigningStatus'){
      const d=body.data||{};
      if(!portalAdminAuthorized_(d.adminPin))return json_({ok:false,error:'الرقم الإداري غير مضبوط أو غير صحيح. اضبطي PORTAL_ADMIN_PIN في خصائص المشروع.'});
      return json_({ok:true,enabled:sharedSigningEnabled_()});
    }
    if(action==='setSharedSigningEnabled'){
      const d=body.data||{};
      if(!portalAdminAuthorized_(d.adminPin))return json_({ok:false,error:'الرقم الإداري غير مضبوط أو غير صحيح. اضبطي PORTAL_ADMIN_PIN في خصائص المشروع.'});
      const enabled=String(d.enabled)==='true';
      PropertiesService.getScriptProperties().setProperty('SHARED_SIGNING_ENABLED',enabled?'true':'false');
      return json_({ok:true,enabled:enabled,message:enabled?'تم تفعيل الرابط الموحّد':'تم إيقاف الرابط الموحّد مؤقتًا'});
    }
    if(action==='redeemEmployeeAccessCode'){
      const d=body.data||{};
      return json_(redeemEmployeeAccessCode_(ss,d.code));
    }
    if(action==='createEmployeeLinks'){
      const d=body.data||{};
      return json_(createEmployeeLinks_(ss,String(d.employeeId||''),String(d.reportType||'all'),String(d.portalBaseUrl||'')));
    }
    if(action==='saveEmployeeSignature'){
      const d=body.data||{};
      return json_(saveEmployeeSignature_(ss,d));
    }
    if(action==='deletePortalSignature'){
      const d=body.data||{};
      return json_(deletePortalSignature_(ss,d));
    }
    if(action==='deleteAllPortalSignatures'){
      const d=body.data||{};
      return json_(deleteAllPortalSignatures_(ss,d));
    }
    if(action==='sync'){
      const lock=LockService.getScriptLock();
      if(!lock.tryLock(10000)) return json_({ok:false,error:'المزامنة مشغولة بطلب آخر؛ أعد المحاولة بعد قليل'});
      try {
        const d=body.data||{};
        const delSuccess=[],delFailed=[];
        (d.deletes||[]).forEach(x=>{
          const type=String(x.type||''),id=String(x.id||'');
          const sheetName=type==='employee'?CONFIG.sheets.employees:(type==='absence'?CONFIG.sheets.absences:(type==='late'||type==='lateness'?CONFIG.sheets.lateness:(type==='action'?CONFIG.sheets.actions:'')));
          if(!sheetName||!id){delFailed.push(id||'');return;}
          try{
            // سجل الحذف مركزيًا قبل حذف الصف حتى لا تعيده أجهزة أخرى من بياناتها القديمة.
            markDeletedRecord_(ss,type,id);
            deleteById_(ss,sheetName,id);
            delSuccess.push(id);
          }catch(err){delFailed.push(id);}
        });
        const result={
          employees:upsertMany_(ss,CONFIG.sheets.employees,d.employees||[]),
          absences:upsertMany_(ss,CONFIG.sheets.absences,d.absences||[]),
          lateness:upsertMany_(ss,CONFIG.sheets.lateness,d.lates||d.lateness||[]),
          signatures:upsertSignatures_(ss,d.signatures||{}),
          settings:upsertSettings_(ss,d.settings||{}),
          actions:upsertMany_(ss,CONFIG.sheets.actions,d.actions||[]),
          deletes:{successIds:delSuccess,failedIds:delFailed}
        };
        return json_({ok:true,synced:true,result:result});
      } finally {
        lock.releaseLock();
      }
    }
    const map={employee:CONFIG.sheets.employees,absence:CONFIG.sheets.absences,lateness:CONFIG.sheets.lateness,late:CONFIG.sheets.lateness,signature:CONFIG.sheets.signatures};
    if(action==='settings'){upsertSettings_(ss,body.data||{});return json_({ok:true,saved:true,action:action});}
    if(action==='delete'){
      const d=body.data||{};
      const type=String(d.type||'');
      const id=String(d.id||'');
      const sheetName=type==='employee'?CONFIG.sheets.employees:(type==='absence'?CONFIG.sheets.absences:(type==='late'||type==='lateness'?CONFIG.sheets.lateness:(type==='action'?CONFIG.sheets.actions:'')));
      if(!sheetName||!id) throw new Error('بيانات الحذف غير مكتملة');
      markDeletedRecord_(ss,type,id);
      const deleted=deleteById_(ss,sheetName,id);
      return json_({ok:true,deleted:deleted,action:'delete',type:type,id:id,error:''});
    }
    if(action==='deleteAll'){
      const d=body.data||{},type=String(d.type||'');
      const sheetName=type==='absence'?CONFIG.sheets.absences:(type==='late'||type==='lateness'?CONFIG.sheets.lateness:(type==='action'?CONFIG.sheets.actions:''));
      if(!sheetName) throw new Error('نوع الحذف غير معروف');
      markAllRecordsDeleted_(ss,type,sheetName);
      const deleted=deleteAllRows_(ss,sheetName);
      return json_({ok:true,deleted:true,action:'deleteAll',type:type,count:deleted});
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
  const headers={};
  headers[CONFIG.sheets.employees]=['id','name','job','createdAt'];
  headers[CONFIG.sheets.absences]=['id','empId','name','job','type','month','from','to','days','ref','reportDate','place','note','created'];
  headers[CONFIG.sheets.lateness]=['id','empId','name','job','date','arrival','start','minutes','reason','note','created'];
  headers[CONFIG.sheets.signatures]=['id','empId','name','type','data','at','createdAt'];
  headers[CONFIG.sheets.settings]=['key','value','updatedAt'];
  headers[CONFIG.sheets.actions]=['id','empId','name','job','date','type','text','note','created'];
  Object.keys(headers).forEach(k=>{const sh=getOrCreate_(ss,k);ensureHeaders_(sh,headers[k]);});
  return 'تم تهيئة جداول منصة ابن هشام';
}

function getAllData_(ss){
  return {
    employees:readSheet_(ss,CONFIG.sheets.employees),
    absences:readSheet_(ss,CONFIG.sheets.absences),
    lates:readSheet_(ss,CONFIG.sheets.lateness),
    signatures:readSignatures_(ss),
    settings:readSettings_(ss),
    actions:readSheet_(ss,CONFIG.sheets.actions)
  };
}

function readSheet_(ss,name){
  const sh=getOrCreate_(ss,name);
  // Read only the used rectangle. getDataRange() can include large formatted-only
  // areas, making the sync endpoint slow even when the sheet has few real records.
  const lastRow=sh.getLastRow(),lastCol=sh.getLastColumn();
  if(lastRow<1||lastCol<1)return[];
  const values=sh.getRange(1,1,lastRow,lastCol).getValues();
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
  rows.forEach(r=>{
    if(!r.key)return;
    let value=r.value;
    // Settings that contain maps/lists (such as employeeEdits) must survive a Sheets round-trip.
    if(typeof value==='string'){
      const trimmed=value.trim();
      if((trimmed.startsWith('{')&&trimmed.endsWith('}'))||(trimmed.startsWith('[')&&trimmed.endsWith(']'))){
        try{value=JSON.parse(trimmed);}catch(_){/* Keep legacy plain-text values unchanged. */}
      }
    }
    out[r.key]=value;
  });
  return out;
}
function upsertMany_(ss,name,records){
  const type=deletedTypeForSheet_(name);
  const deleted=type?getDeletedRecordIds_(ss,type):new Set();
  const input=(records||[]).filter(r=>r&&(r.id||r.key)&&!(type&&deleted.has(String(r.id||''))));
  if(!input.length)return 0;
  const sh=getOrCreate_(ss,name);
  const preferred={
    Employees:['id','name','job','createdAt'],
    Absences:['id','empId','name','job','type','month','from','to','days','ref','reportDate','place','note','created'],
    Lateness:['id','empId','name','job','date','arrival','start','minutes','reason','note','created'],
    Signatures:['id','empId','name','type','data','at','createdAt'],
    Settings:['key','value','updatedAt'],
    Actions:['id','empId','name','job','date','type','text','note','created']
  }[name]||Object.keys(input[0]);
  ensureHeaders_(sh,preferred);
  const headers=getHeaders_(sh),keyField=isSettingsSheet_(name)?'key':'id',keyCol=headers.indexOf(keyField);
  if(!headers.length||keyCol<0){
    input.forEach(r=>upsertOne_(ss,name,r));
    return input.length;
  }

  // Keep the previous last-record-wins behavior for duplicate keys.
  const latestByKey=new Map();
  input.forEach(r=>{const k=String(r[keyField]||'');if(k)latestByKey.set(k,r);});
  if(!latestByKey.size)return 0;

  // Read existing rows once, but write only rows whose values actually changed.
  const lastRow=sh.getLastRow();
  const rows=lastRow>1?sh.getRange(2,1,lastRow-1,headers.length).getValues():[];
  const rowByKey=new Map();
  rows.forEach((row,i)=>{const k=String(row[keyCol]===null||row[keyCol]===undefined?'':row[keyCol]).trim();if(k)rowByKey.set(k,i);});
  const updates=new Map(),newRows=[];
  const sameValue=(a,b)=>{
    const av=normalize_(a),bv=normalize_(b);
    return String(av===null||av===undefined?'':av)===String(bv===null||bv===undefined?'':bv);
  };
  latestByKey.forEach((record,key)=>{
    const values=headers.map(h=>record[h]===undefined?'':normalize_(record[h]));
    const index=rowByKey.get(key);
    if(index!==undefined){
      const old=rows[index];
      if(values.some((v,j)=>!sameValue(old[j],v)))updates.set(index,values);
    }else{
      newRows.push(values);
    }
  });

  // Batch contiguous changed rows to reduce Spreadsheet service calls.
  const indices=[...updates.keys()].sort((a,b)=>a-b);
  let p=0;
  while(p<indices.length){
    let q=p+1;
    while(q<indices.length&&indices[q]===indices[q-1]+1)q++;
    const block=indices.slice(p,q);
    sh.getRange(block[0]+2,1,block.length,headers.length).setValues(block.map(i=>updates.get(i)));
    p=q;
  }
  // Append new records in one write.
  if(newRows.length){
    const appendAt=Math.max(2,lastRow+1);
    sh.getRange(appendAt,1,newRows.length,headers.length).setValues(newRows);
  }
  return latestByKey.size;
}
function upsertSignatures_(ss,obj){
  const records=Object.keys(obj||{}).map(id=>{const s=obj[id]||{};return {id:id,empId:id,name:s.name||'',type:s.type||'',data:s.data||'',at:s.at||'',createdAt:s.createdAt||s.at||''};});
  return upsertMany_(ss,CONFIG.sheets.signatures,records);
}

/**
 * سجل مركزي للحذف يمنع الأجهزة الأخرى من إعادة إنشاء سجل محذوف عبر المزامنة.
 * لا تُحذف العلامات من هذا السجل؛ معرّفات السجلات الجديدة يجب أن تكون فريدة.
 */
const DELETED_RECORDS_SHEET_ = '_DeletedRecords';
function deletedTypeForSheet_(name){
  if(name===CONFIG.sheets.employees)return 'employee';
  if(name===CONFIG.sheets.absences)return 'absence';
  if(name===CONFIG.sheets.lateness)return 'late';
  if(name===CONFIG.sheets.actions)return 'action';
  return '';
}
function deletedRecordsSheet_(ss){
  let sh=ss.getSheetByName(DELETED_RECORDS_SHEET_);
  if(!sh)sh=ss.insertSheet(DELETED_RECORDS_SHEET_);
  if(sh.getLastRow()===0)sh.getRange(1,1,1,3).setValues([['type','id','deletedAt']]);
  return sh;
}
function getDeletedRecordIds_(ss,type){
  type=normalizeDeletedType_(type);
  const sh=ss.getSheetByName(DELETED_RECORDS_SHEET_);
  const ids=new Set();
  if(!sh||sh.getLastRow()<2)return ids;
  sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues().forEach(row=>{
    if(String(row[0]||'')===String(type))ids.add(String(row[1]||''));
  });
  return ids;
}
function normalizeDeletedType_(type){
  type=String(type||'').trim().toLowerCase();
  return type==='lateness'?'late':type;
}
function markDeletedRecord_(ss,type,id){
  id=String(id||'').trim();type=normalizeDeletedType_(type);
  if(!id||!type)return;
  const sh=deletedRecordsSheet_(ss),last=sh.getLastRow();
  if(last>1){
    const rows=sh.getRange(2,1,last-1,2).getDisplayValues();
    for(let i=0;i<rows.length;i++){
      if(String(rows[i][0]||'')===type&&String(rows[i][1]||'')===id)return;
    }
  }
  sh.getRange(sh.getLastRow()+1,1,1,3).setValues([[type,id,new Date().toISOString()]]);
}
function markAllRecordsDeleted_(ss,type,sheetName){
  type=normalizeDeletedType_(type);
  const source=ss.getSheetByName(sheetName);
  if(!source||source.getLastRow()<2)return;
  const headers=getHeaders_(source);
  const idCol=headers.findIndex(h=>String(h||'').trim().toLowerCase()==='id');
  if(idCol<0)return;
  const ids=source.getRange(2,idCol+1,source.getLastRow()-1,1).getDisplayValues()
    .map(row=>String(row[0]||'').trim()).filter(Boolean);
  if(!ids.length)return;
  const target=deletedRecordsSheet_(ss),known=new Set();
  if(target.getLastRow()>1){
    target.getRange(2,1,target.getLastRow()-1,2).getDisplayValues().forEach(row=>{
      if(String(row[0]||'')===String(type))known.add(String(row[1]||''));
    });
  }
  const now=new Date().toISOString(),newRows=[];
  ids.forEach(id=>{if(!known.has(id)){known.add(id);newRows.push([type,id,now]);}});
  if(newRows.length)target.getRange(target.getLastRow()+1,1,newRows.length,3).setValues(newRows);
}

function deleteAllRows_(ss,name){
  const sh=getOrCreate_(ss,name);
  const last=sh.getLastRow();
  if(last<2)return 0;
  const count=last-1;
  sh.deleteRows(2,count);
  SpreadsheetApp.flush();
  return count;
}
function deleteById_(ss,name,id){
  const sh=getOrCreate_(ss,name);
  const lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if(lastRow<2 || lastCol<1)return false;
  const wanted=String(id===null||id===undefined?'':id).trim();
  if(!wanted)return false;
  const header=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(x=>String(x||'').trim().toLowerCase());
  const keyField=isSettingsSheet_(name)?'key':'id';
  let col=header.indexOf(keyField)+1;
  // دعم اختلاف كتابة رأس العمود مثل ID أو id أو وجود مسافات.
  if(col<1)col=header.findIndex(x=>x.replace(/[^a-z]/g,'')===keyField)+1;
  if(col>0){
    const vals=sh.getRange(2,col,lastRow-1,1).getDisplayValues();
    for(let i=0;i<vals.length;i++){
      if(String(vals[i][0]||'').trim()===wanted){
        sh.deleteRow(i+2);
        SpreadsheetApp.flush();
        return true;
      }
    }
  }
  // احتياط: ابحث عن الـID في كامل الصف إذا كان ملف الشيت القديم لا يملك رأس id مطابقًا.
  const rows=sh.getRange(2,1,lastRow-1,lastCol).getDisplayValues();
  for(let i=0;i<rows.length;i++){
    if(rows[i].some(v=>String(v||'').trim()===wanted)){
      sh.deleteRow(i+2);
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}
function upsertOne_(ss,name,record){
  if(!record||(!record.id&&!record.key))return;
  const deletedType=deletedTypeForSheet_(name);
  if(deletedType&&getDeletedRecordIds_(ss,deletedType).has(String(record.id||'')))return;
  const sh=getOrCreate_(ss,name);
  const preferred={
    Employees:['id','name','job','createdAt'],
    Absences:['id','empId','name','job','type','month','from','to','days','ref','reportDate','place','note','created'],
    Lateness:['id','empId','name','job','date','arrival','start','minutes','reason','note','created'],
    Signatures:['id','empId','name','type','data','at','createdAt'],
    Settings:['key','value','updatedAt']
  }[name]||Object.keys(record);
  ensureHeaders_(sh,preferred);
  const headers=getHeaders_(sh),keyField=isSettingsSheet_(name)?'key':'id',key=String(record[keyField]||'');
  if(!key)return;
  const row=findRow_(sh,keyField,key,headers),values=headers.map(h=>record[h]===undefined?'':normalize_(record[h]));
  if(row)sh.getRange(row,1,1,headers.length).setValues([values]);
  else sh.getRange(sh.getLastRow()+1,1,1,headers.length).setValues([values]);
}
function isSettingsSheet_(name){return name==='Settings'||name===CONFIG.sheets.settings;}
function upsertSettings_(ss,settings){
  const entries=Object.keys(settings||{});
  if(!entries.length)return 0;
  const sh=getOrCreate_(ss,CONFIG.sheets.settings);
  const headers=['key','value','updatedAt'];
  ensureHeaders_(sh,headers);
  const actualHeaders=getHeaders_(sh);
  const keyCol=actualHeaders.indexOf('key');
  if(keyCol<0){
    entries.forEach(k=>{
      const raw=settings[k];
      const value=raw!==null&&typeof raw==='object'?JSON.stringify(raw):raw;
      upsertOne_(ss,CONFIG.sheets.settings,{key:k,value:value,updatedAt:new Date().toISOString()});
    });
    return entries.length;
  }
  const last=sh.getLastRow();
  const rows=last>1?sh.getRange(2,1,last-1,actualHeaders.length).getValues():[];
  const rowByKey=new Map();
  rows.forEach((row,i)=>{const key=String(row[keyCol]??'').trim();if(key)rowByKey.set(key,i);});
  const updates=new Map(),appends=[];
  const now=new Date().toISOString();
  entries.forEach(k=>{
    const raw=settings[k];
    const value=raw!==null&&typeof raw==='object'?JSON.stringify(raw):raw;
    const record={key:k,value:value,updatedAt:now};
    const values=actualHeaders.map(h=>record[h]===undefined?'':normalize_(record[h]));
    const idx=rowByKey.get(String(k));
    if(idx===undefined)appends.push(values);
    else updates.set(idx,values);
  });
  const indices=[...updates.keys()].sort((a,b)=>a-b);
  let p=0;
  while(p<indices.length){
    let q=p+1;
    while(q<indices.length&&indices[q]===indices[q-1]+1)q++;
    const block=indices.slice(p,q);
    sh.getRange(block[0]+2,1,block.length,actualHeaders.length).setValues(block.map(i=>updates.get(i)));
    p=q;
  }
  if(appends.length)sh.getRange(sh.getLastRow()+1,1,appends.length,actualHeaders.length).setValues(appends);
  return entries.length;
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


/* ===== روابط تقارير الموظفات: قراءة فقط، دون مفتاح إداري =====
   يتم حفظ الرموز في ورقة منفصلة لتفادي حدود حجم Script Properties.
*/
const PORTAL_SHEET_='EmployeePortalTokens';
const PORTAL_SIGNATURES_SHEET_='EmployeeReportSignatures';
function portalTokenSheet_(ss){
  const sh=getOrCreate_(ss,PORTAL_SHEET_);
  const headers=['token','id','name','job','createdAt','reportType'];
  if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}
function portalSignaturesSheet_(ss){
  const sh=getOrCreate_(ss,PORTAL_SIGNATURES_SHEET_);
  const headers=['id','token','empId','name','type','data','at','createdAt'];
  if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}
function portalRandomToken_(){
  return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
}
function createEmployeeLinks_(ss,employeeId,reportType,portalBaseUrl){
  const validTypes=['all','absence','late'];
  reportType=validTypes.includes(reportType)?reportType:'all';
  const employees=readSheet_(ss,CONFIG.sheets.employees).filter(e=>e&&e.id&&e.name&&(!employeeId||String(e.id)===String(employeeId)));
  if(!employees.length) return {ok:true,links:[],message:employeeId?'لم يتم العثور على الموظفة المحددة':'لا توجد موظفات مسجلة'};
  const sh=portalTokenSheet_(ss);
  const existing=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,6).getDisplayValues():[];
  const replacing=new Set(employees.map(e=>String(e.id)+'|'+reportType));
  const kept=existing.filter(r=>!replacing.has(String(r[1]||'')+'|'+String(r[5]||'all')));
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,6).clearContent();
  const now=new Date(),rows=[],links=[];
  kept.forEach(r=>rows.push(r));
  employees.forEach(e=>{
    const token=portalRandomToken_();
    rows.push([token,String(e.id),String(e.name||'').trim(),String(e.job||'').trim(),now,reportType]);
    const safeBase=portalLinkBase_(portalBaseUrl);
    links.push({id:String(e.id),name:String(e.name),reportType:reportType,token:token,url:safeBase+'?portal='+encodeURIComponent(token)});
  });
  if(rows.length) sh.getRange(2,1,rows.length,6).setValues(rows);
  SpreadsheetApp.flush();
  return {ok:true,links:links,message:'تم إنشاء الرابط المحدد؛ تم استبدال الرابط السابق لنفس الموظفة ونوع التقرير فقط'};
}

function portalLinkBase_(candidate){
  const fallback='https://alsaabxa.github.io/IbnHisham/';
  try{
    const u=new URL(String(candidate||''));
    const host=String(u.hostname||'').toLowerCase();
    const allowed=host==='alsaabxa.github.io'||host==='ibnhisham.pages.dev'||/^[a-z0-9-]+\.ibnhisham\.pages\.dev$/.test(host);
    if(!allowed||u.protocol!=='https:')return fallback;
    return u.origin+u.pathname.replace(/[^/]*$/,'');
  }catch(e){return fallback;}
}


const PORTAL_ACCESS_CODES_SHEET_='EmployeePortalAccessCodes';
function portalAdminAuthorized_(provided){
  const expected=String(PropertiesService.getScriptProperties().getProperty('PORTAL_ADMIN_PIN')||'');
  return !!expected && String(provided||'')===expected;
}
function portalAccessCodeSheet_(ss){
  const sh=getOrCreate_(ss,PORTAL_ACCESS_CODES_SHEET_);
  const headers=['code','token','empId','name','reportType','active','updatedAt'];
  if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}
function portalNewAccessCode_(ss){
  const sh=portalAccessCodeSheet_(ss);
  const used=new Set(sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().flat().map(r=>String(r).toUpperCase()):[]);
  const chars='ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code='';
  do{code='';for(let i=0;i<10;i++)code+=chars.charAt(Math.floor(Math.random()*chars.length));}while(used.has(code));
  return code;
}
function ensureEmployeeAccessCodes_(ss,portalBaseUrl){
  const employees=readSheet_(ss,CONFIG.sheets.employees).filter(e=>e&&e.id&&String(e.name||'').trim());
  const sh=portalAccessCodeSheet_(ss);
  const existing=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,7).getDisplayValues():[];
  const rowsByEmp=new Map();
  existing.forEach((r,i)=>{
    const id=String(r[2]||'');
    if(id){
      if(!rowsByEmp.has(id))rowsByEmp.set(id,[]);
      rowsByEmp.get(id).push(i+2);
    }
  });
  const tokenSheet=portalTokenSheet_(ss);
  const tokenRows=tokenSheet.getDataRange().getDisplayValues();
  employees.forEach(emp=>{
    const id=String(emp.id),name=String(emp.name||'').trim();
    const existingRows=rowsByEmp.get(id)||[];
    if(existingRows.length){
      // Preserve each employee's access code, but repair a missing/stale token
      // for this employee only. This avoids invalidating other employees' links.
      const repairedTokens={};
      existingRows.forEach(row=>{
        const currentName=String(sh.getRange(row,4).getDisplayValue()||'').trim();
        if(currentName!==name)sh.getRange(row,4).setValue(name);
        const reportType=String(sh.getRange(row,5).getDisplayValue()||'all').trim()||'all';
        const currentToken=String(sh.getRange(row,2).getDisplayValue()||'').trim();
        let tokenInfo=currentToken?findPortalToken_(ss,currentToken):null;
        const tokenIsValid=tokenInfo&&String(tokenInfo.id)===id&&String(tokenInfo.reportType||'all')===reportType;
        if(tokenIsValid)return;
        const key=id+'|'+reportType;
        let replacementToken=repairedTokens[key]||'';
        if(!replacementToken){
          // Reuse a valid token for the same employee/report type if one exists.
          const freshTokenRows=portalTokenSheet_(ss).getDataRange().getDisplayValues();
          for(let i=1;i<freshTokenRows.length;i++){
            const tr=freshTokenRows[i];
            if(String(tr[1]||'').trim()!==id||String(tr[5]||'all').trim()!==reportType)continue;
            const candidate=String(tr[0]||'').trim();
            const candidateInfo=candidate?findPortalToken_(ss,candidate):null;
            if(candidateInfo&&String(candidateInfo.id)===id&&String(candidateInfo.reportType||'all')===reportType){
              replacementToken=candidate;break;
            }
          }
          // If no valid token remains, issue a new token for this employee/type only.
          if(!replacementToken){
            const created=createEmployeeLinks_(ss,id,reportType,String(portalBaseUrl||''));
            if(created&&created.ok&&created.links&&created.links.length){
              replacementToken=String(created.links[0].token||'');
            }
          }
          if(replacementToken)repairedTokens[key]=replacementToken;
        }
        if(replacementToken){
          sh.getRange(row,2).setValue(replacementToken);
          sh.getRange(row,7).setValue(new Date());
        }
      });
      return;
    }
    let token='';
    for(let i=1;i<tokenRows.length;i++){
      if(String(tokenRows[i][1]||'')===id&&String(tokenRows[i][5]||'all')==='all'){
        const candidate=String(tokenRows[i][0]||'');
        if(candidate&&findPortalToken_(ss,candidate)){token=candidate;break;}
      }
    }
    if(!token){
      const created=createEmployeeLinks_(ss,id,'all',String(portalBaseUrl||''));
      if(created&&created.ok&&created.links&&created.links.length)token=String(created.links[0].token||'');
    }
    if(!token)return;
    const code=portalNewAccessCode_(ss);
    const row=[code,token,id,name,'all','TRUE',new Date()];
    const newRow=sh.getLastRow()+1;
    sh.getRange(newRow,1,1,7).setValues([row]);
    rowsByEmp.set(id,[newRow]);
  });
  SpreadsheetApp.flush();
  const last=sh.getLastRow();
  if(last<2)return [];
  return sh.getRange(2,1,last-1,7).getDisplayValues().filter(r=>r[0]&&r[1])
    .map(r=>({code:String(r[0]),empId:String(r[2]),name:String(r[3]),reportType:String(r[4]||'all'),active:String(r[5]).toUpperCase()!=='FALSE'}));
}
function listEmployeeAccessCodes_(ss){
  // عرض الرموز عملية قراءة فقط: لا تفحص جميع الموظفات ولا تصلح/تنشئ الروابط عند فتح الإعدادات.
  // إنشاء الرموز الناقصة أو إصلاحها يبقى ضمن زر الإنشاء الصريح فقط.
  const sh=portalAccessCodeSheet_(ss);
  const last=sh.getLastRow();
  if(last<2)return [];
  return sh.getRange(2,1,last-1,7).getDisplayValues()
    .filter(r=>r[0]&&r[1])
    .map(r=>({code:String(r[0]),empId:String(r[2]),name:String(r[3]),reportType:String(r[4]||'all'),active:String(r[5]).toUpperCase()!=='FALSE'}));
}
function sharedSigningEnabled_(){
  return String(PropertiesService.getScriptProperties().getProperty('SHARED_SIGNING_ENABLED')||'true').toLowerCase()!=='false';
}
function redeemEmployeeAccessCode_(ss,provided){
  if(!sharedSigningEnabled_())return {ok:false,error:'تم إيقاف رابط التوقيع الموحّد مؤقتًا من الإدارة. يرجى المحاولة لاحقًا.'};
  const code=String(provided||'').trim().toUpperCase();
  if(!/^[A-HJ-NP-Z2-9]{6,12}$/.test(code))return {ok:false,error:'تحققي من الرمز ثم حاولي مرة أخرى'};
  const sh=portalAccessCodeSheet_(ss);
  if(sh.getLastRow()<2)return {ok:false,error:'لم يتم تفعيل رموز الدخول بعد'};
  const values=sh.getRange(2,1,sh.getLastRow()-1,7).getDisplayValues();
  for(let i=0;i<values.length;i++){
    const r=values[i];
    if(String(r[0]).toUpperCase()===code && String(r[5]).toUpperCase()!=='FALSE' && r[1]){
      if(!findPortalToken_(ss,String(r[1])))return {ok:false,error:'الرمز غير صالح حاليًا. تواصلي مع الإدارة.'};
      return {ok:true,token:String(r[1])};
    }
  }
  return {ok:false,error:'الرمز غير صحيح أو تم تغييره. تواصلي مع الإدارة.'};
}
function updateEmployeeAccessCode_(ss,data){
  const empId=String(data.empId||'').trim(),code=String(data.code||'').trim().toUpperCase();
  if(!empId)return {ok:false,error:'معرّف الموظفة غير محدد'};
  if(!/^[A-HJ-NP-Z2-9]{6,12}$/.test(code))return {ok:false,error:'استخدمي رمزًا من 6 إلى 12 حرفًا/رقمًا دون مسافات'};
  const sh=portalAccessCodeSheet_(ss),last=sh.getLastRow();
  if(last<2)return {ok:false,error:'لا توجد رموز. أنشئي الرموز أولاً.'};
  const values=sh.getRange(2,1,last-1,7).getDisplayValues();let rowIndex=-1;
  for(let i=0;i<values.length;i++){
    if(String(values[i][2])===empId)rowIndex=i+2;
    else if(String(values[i][0]).toUpperCase()===code)return {ok:false,error:'هذا الرمز مستخدم لموظفة أخرى'};
  }
  if(rowIndex<0)return {ok:false,error:'لم يتم العثور على رمز لهذه الموظفة'};
  sh.getRange(rowIndex,1).setValue(code);sh.getRange(rowIndex,7).setValue(new Date());SpreadsheetApp.flush();
  return {ok:true,code:code,message:'تم تعديل الرمز'};
}

function findPortalToken_(ss,token){
  if(!token)return null;
  const values=portalTokenSheet_(ss).getDataRange().getDisplayValues();
  for(let i=1;i<values.length;i++){
    if(String(values[i][0]||'')===String(token)){
      const id=String(values[i][1]||'').trim();
      let name=String(values[i][2]||'').trim();
      let job=String(values[i][3]||'').trim();
      // Always use the current employee name/job from the master Employees sheet.
      // Portal tokens can outlive name corrections, so their stored copy may be stale.
      const current=readSheet_(ss,CONFIG.sheets.employees).find(e=>String(e.id||'').trim()===id);
      if(current){
        name=String(current.name||name).trim();
        job=String(current.job||job).trim();
      }
      return {token:String(values[i][0]),id:id,name:name,job:job,reportType:String(values[i][5]||'all')||'all'};
    }
  }
  return null;
}
function saveEmployeeSignature_(ss,data){
  if(data.shared===true&&!sharedSigningEnabled_())return {ok:false,error:'تم إيقاف رابط التوقيع الموحّد مؤقتًا. لم يتم حفظ التوقيع.'};
  const token=String(data.token||'');
  const emp=findPortalToken_(ss,token);
  if(!emp)return {ok:false,error:'الرابط غير صالح أو تم إلغاؤه. اطلبي رابطًا جديدًا من الإدارة.'};
  const reportType=String(data.reportType||'all');
  if(reportType!==emp.reportType)return {ok:false,error:'نوع التقرير لا يطابق الرابط المخصص.'};
  const image=String(data.data||'');
  if(!/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(image))return {ok:false,error:'بيانات التوقيع غير صالحة.'};
  if(image.length>1500000)return {ok:false,error:'حجم التوقيع كبير؛ امسحي مساحة التوقيع وأعيدي التوقيع بحجم أصغر.'};
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(10000))return {ok:false,error:'الخدمة مشغولة بحفظ عملية أخرى. انتظري قليلًا ثم حاولي مرة أخرى.'};
  try{
    const sh=portalSignaturesSheet_(ss);
    const last=sh.getLastRow();
    const nowDate=new Date();
    const today=Utilities.formatDate(nowDate,CONFIG.timezone,'yyyy-MM-dd');
    const existing=last>1?sh.getRange(2,1,last-1,8).getValues():[];
    const alreadySignedToday=existing.some(r=>{
      const sameEmployee=String(r[2]||'')===String(emp.id) || String(r[1]||r[0]||'')===token;
      if(!sameEmployee)return false;
      const raw=r[6];
      if(!raw)return false;
      const signedAt=raw instanceof Date?raw:new Date(raw);
      if(isNaN(signedAt.getTime()))return String(raw).slice(0,10)===today;
      return Utilities.formatDate(signedAt,CONFIG.timezone,'yyyy-MM-dd')===today;
    });
    if(alreadySignedToday)return {ok:false,error:'تم اعتماد توقيعك لهذا اليوم بالفعل. يمكنك التوقيع مرة أخرى غدًا.'};
    const now=nowDate.toISOString();
    const row=[token,token,String(emp.id),String(emp.name||''),reportType,image,now,now];
    sh.getRange(sh.getLastRow()+1,1,1,row.length).setValues([row]);
    SpreadsheetApp.flush();
    return {ok:true,saved:true,at:now,message:'تم حفظ التوقيع بنجاح'};
  }finally{
    lock.releaseLock();
  }
}
function listPortalSignatures_(ss){
  return readSheet_(ss,PORTAL_SIGNATURES_SHEET_).map(r=>({id:String(r.id||r.token||''),name:r.name||'',empId:r.empId||'',type:r.type||'all',at:r.at||''})).sort((a,b)=>String(b.at).localeCompare(String(a.at)));
}
function deletePortalSignature_(ss,data){
  const id=String(data.id||'').trim();
  const password=String(data.password||'');
  if(password!=='1234') return {ok:false,error:'الرقم السري غير صحيح'};
  if(!/^[a-f0-9]{32,64}$/i.test(id)) return {ok:false,error:'معرّف التوقيع غير صالح'};
  const sh=portalSignaturesSheet_(ss);
  const last=sh.getLastRow();
  if(last<2) return {ok:false,error:'التوقيع غير موجود أو سبق حذفه'};
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
  for(let i=ids.length-1;i>=0;i--){
    if(String(ids[i][0]||'')===id){
      sh.deleteRow(i+2);
      SpreadsheetApp.flush();
      return {ok:true,deleted:true,id:id,message:'تم حذف التوقيع المحدد'};
    }
  }
  return {ok:false,error:'التوقيع غير موجود أو سبق حذفه'};
}
function deletePortalSignature_(ss,data){
  const id=String(data.id||'').trim();
  const password=String(data.password||'');
  if(password!=='1234') return {ok:false,error:'الرقم السري غير صحيح'};
  if(!/^[a-f0-9]{32,64}$/i.test(id)) return {ok:false,error:'معرّف التوقيع غير صالح'};
  const sh=portalSignaturesSheet_(ss);
  const last=sh.getLastRow();
  if(last<2) return {ok:false,error:'التوقيع غير موجود أو سبق حذفه'};
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
  for(let i=ids.length-1;i>=0;i--){
    if(String(ids[i][0]||'')===id){
      sh.deleteRow(i+2);
      SpreadsheetApp.flush();
      return {ok:true,deleted:true,id:id,message:'تم حذف التوقيع المحدد'};
    }
  }
  return {ok:false,error:'التوقيع غير موجود أو سبق حذفه'};
}
function deleteAllPortalSignatures_(ss,data){
  const password=String((data&&data.password)||'');
  if(password!=='1234') return {ok:false,error:'الرقم السري غير صحيح'};
  const sh=portalSignaturesSheet_(ss);
  const last=sh.getLastRow();
  if(last<2) return {ok:true,deleted:true,count:0,message:'لا توجد توقيعات لحذفها'};
  const count=last-1;
  sh.deleteRows(2,count);
  SpreadsheetApp.flush();
  return {ok:true,deleted:true,count:count,message:'تم حذف جميع التوقيعات الإلكترونية'};
}

function employeeReport_(ss,token){
  const emp=findPortalToken_(ss,token);
  if(!emp)return {ok:false,error:token?'الرابط غير صالح أو تم إلغاؤه. أنشئي رابطًا جديدًا من الإدارة.':'الرابط غير مكتمل'};
  const belongs=x=>{
    const id=String(x.empId??x.employeeId??x.employeeID??x.employee_id??'').trim();
    if(id && emp.id && id===emp.id) return true;
    return String(x.name??x.employeeName??x.employee_name??'').trim()===emp.name;
  };
  const absences=emp.reportType==='late'?[]:readSheet_(ss,CONFIG.sheets.absences).filter(belongs);
  const lates=emp.reportType==='absence'?[]:readSheet_(ss,CONFIG.sheets.lateness).filter(belongs);
  const actions=readSheet_(ss,CONFIG.sheets.actions).filter(belongs);
  const dateVal=x=>String(x.date||x.from||x.reportDate||x.created||'');
  const desc=(a,b)=>dateVal(b).localeCompare(dateVal(a));
  absences.sort(desc); lates.sort(desc); actions.sort(desc);
  const signatures=readSheet_(ss,PORTAL_SIGNATURES_SHEET_);
  const today=Utilities.formatDate(new Date(),CONFIG.timezone,'yyyy-MM-dd');
  const signature=signatures.find(x=>{
    if(String(x.token||x.id||'')!==String(token))return false;
    const raw=x.at||'';
    const signedAt=raw instanceof Date?raw:new Date(raw);
    return !isNaN(signedAt.getTime())&&Utilities.formatDate(signedAt,CONFIG.timezone,'yyyy-MM-dd')===today;
  })||null;
  return {ok:true,reportType:emp.reportType,employee:{id:emp.id,name:emp.name,job:emp.job},data:{absences:absences,lates:lates,actions:actions,signature:signature?{data:signature.data,at:signature.at,type:signature.type}:null}};
}
