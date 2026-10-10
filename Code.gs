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
      const d=body.data||{};
      const delSuccess=[],delFailed=[];
      (d.deletes||[]).forEach(x=>{
        const type=String(x.type||''),id=String(x.id||'');
        const sheetName=type==='absence'?CONFIG.sheets.absences:(type==='late'||type==='lateness'?CONFIG.sheets.lateness:(type==='action'?CONFIG.sheets.actions:''));
        if(!sheetName||!id){delFailed.push(id||'');return;}
        try{if(deleteById_(ss,sheetName,id))delSuccess.push(id);else delFailed.push(id);}catch(err){delFailed.push(id);}
      });
      return json_({ok:true,synced:true,result:{
        employees:upsertMany_(ss,CONFIG.sheets.employees,d.employees||[]),
        absences:upsertMany_(ss,CONFIG.sheets.absences,d.absences||[]),
        lateness:upsertMany_(ss,CONFIG.sheets.lateness,d.lates||d.lateness||[]),
        signatures:upsertSignatures_(ss,d.signatures||{}),
        settings:upsertSettings_(ss,d.settings||{}),
        actions:upsertMany_(ss,CONFIG.sheets.actions,d.actions||[]),
        deletes:{successIds:delSuccess,failedIds:delFailed}
      }});
    }
    const map={employee:CONFIG.sheets.employees,absence:CONFIG.sheets.absences,lateness:CONFIG.sheets.lateness,late:CONFIG.sheets.lateness,signature:CONFIG.sheets.signatures};
    if(action==='settings'){upsertSettings_(ss,body.data||{});return json_({ok:true,saved:true,action:action});}
    if(action==='delete'){
      const d=body.data||{};
      const type=String(d.type||'');
      const id=String(d.id||'');
      const sheetName=type==='absence'?CONFIG.sheets.absences:(type==='late'||type==='lateness'?CONFIG.sheets.lateness:(type==='action'?CONFIG.sheets.actions:''));
      if(!sheetName||!id) throw new Error('بيانات الحذف غير مكتملة');
      const deleted=deleteById_(ss,sheetName,id);
      return json_({ok:deleted,deleted:deleted,action:'delete',type:type,id:id,error:deleted?'':('السجل غير موجود: '+id)});
    }
    if(action==='deleteAll'){
      const d=body.data||{},type=String(d.type||'');
      const sheetName=type==='absence'?CONFIG.sheets.absences:(type==='late'||type==='lateness'?CONFIG.sheets.lateness:(type==='action'?CONFIG.sheets.actions:''));
      if(!sheetName) throw new Error('نوع الحذف غير معروف');
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
  const input=(records||[]).filter(r=>r&&(r.id||r.key));
  if(!input.length)return 0;
  const sh=getOrCreate_(ss,name);
  const preferred={
    Employees:['id','name','job','createdAt'],
    Absences:['id','empId','name','job','type','month','from','to','days','ref','reportDate','place','note','created'],
    Lateness:['id','empId','name','job','date','arrival','start','minutes','reason','note','created'],
    Signatures:['id','empId','name','type','data','at','createdAt'],
    Settings:['key','value','updatedAt']
  }[name]||Object.keys(input[0]);
  ensureHeaders_(sh,preferred);
  const headers=getHeaders_(sh),keyField=name==='Settings'?'key':'id',keyCol=headers.indexOf(keyField)+1;
  if(!headers.length||keyCol<1){
    input.forEach(r=>upsertOne_(ss,name,r));
    return input.length;
  }
  const last=sh.getLastRow(),rowByKey=new Map();
  if(last>1){
    const keys=sh.getRange(2,keyCol,last-1,1).getDisplayValues();
    keys.forEach((r,i)=>{const k=String(r[0]||'');if(k)rowByKey.set(k,i+2);});
  }
  // Collapse duplicate IDs in the same sync; the last record wins, as in the old sequential upsert.
  const latestByKey=new Map();
  input.forEach(r=>{const k=String(r[keyField]||'');if(k)latestByKey.set(k,r);});
  const append=[];
  latestByKey.forEach((record,key)=>{
    const values=headers.map(h=>record[h]===undefined?'':normalize_(record[h]));
    const row=rowByKey.get(key);
    if(row)sh.getRange(row,1,1,headers.length).setValues([values]);
    else append.push(values);
  });
  if(append.length)sh.getRange(sh.getLastRow()+1,1,append.length,headers.length).setValues(append);
  return latestByKey.size;
}
function upsertSignatures_(ss,obj){
  const records=Object.keys(obj||{}).map(id=>{const s=obj[id]||{};return {id:id,empId:id,name:s.name||'',type:s.type||'',data:s.data||'',at:s.at||'',createdAt:s.createdAt||s.at||''};});
  return upsertMany_(ss,CONFIG.sheets.signatures,records);
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
  const keyField=name==='Settings'?'key':'id';
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
      // Update the display name on every existing row for this employee.
      // Keep all existing access codes and tokens unchanged.
      existingRows.forEach(row=>{
        const currentName=String(sh.getRange(row,4).getDisplayValue()||'').trim();
        if(currentName!==name)sh.getRange(row,4).setValue(name);
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
  return ensureEmployeeAccessCodes_(ss,'');
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
    const existing=last>1?sh.getRange(2,1,last-1,3).getDisplayValues():[];
    if(existing.some(r=>String(r[1]||r[0]||'')===token||String(r[2]||'')===String(emp.id)))return {ok:false,error:'تم اعتماد توقيع هذه الموظفة مسبقًا. لا يمكن اعتماد توقيع آخر إلا بعد حذف التوقيع الحالي من الإدارة.'};
    const now=new Date().toISOString();
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
  const signature=signatures.find(x=>String(x.token||x.id||'')===String(token))||null;
  return {ok:true,reportType:emp.reportType,employee:{id:emp.id,name:emp.name,job:emp.job},data:{absences:absences,lates:lates,actions:actions,signature:signature?{data:signature.data,at:signature.at,type:signature.type}:null}};
}
