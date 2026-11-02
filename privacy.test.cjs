const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const source=fs.readFileSync(__dirname+'/Code.gs','utf8');
let opens=0;
const ctx=vm.createContext({console,URL,PropertiesService:{getScriptProperties:()=>({getProperty:k=>k==='PORTAL_ADMIN_PIN'?'test-private-pin':null})},SpreadsheetApp:{openById:()=>{opens++;return {}; }},ContentService:{MimeType:{JSON:'json'},createTextOutput:text=>({setMimeType:()=>JSON.parse(text)})},Utilities:{formatDate:()=> '2026-11-02'}});
vm.runInContext(source,ctx);
for(const action of ['all','employees','absences','lates','lateness','settings','portalSignatures'])assert.equal(ctx.doGet({parameter:{action}}).ok,false,action);
for(const action of ['adminRead','sync','createEmployeeLinks','delete','deleteAll','employee','absence','late','settings','setStartTime','deletePortalSignature','deleteAllPortalSignatures','generateEmployeeAccessCodes']){
 for(const adminPin of [undefined,'wrong'])assert.equal(ctx.doPost({postData:{contents:JSON.stringify({action,adminPin,data:{}})}}).ok,false,action);
}
assert.equal(opens,0,'Unauthorized requests must not open the spreadsheet');
assert.equal(ctx.doPost({postData:{contents:JSON.stringify({action:'adminRead',adminPin:'test-private-pin',data:{readAction:'ping'}})}}).ok,true);
vm.runInContext(`
findPortalToken_=(ss,token)=>token==='valid'?{id:'one',name:'Same Name',job:'Teacher',reportType:'all'}:null;
readSheet_=(ss,name)=>name===CONFIG.sheets.employees?[]:name===PORTAL_SIGNATURES_SHEET_?[]:[
 {empId:'one',name:'Renamed',marker:'own'},
 {empId:'two',name:'Same Name',marker:'other'},
 {name:'Same Name',marker:'missing'},
 {employeeId:'one',marker:'alias'}];
`,ctx);
for(const token of ['', 'invalid'])assert.equal(ctx.employeeReport_({},token).ok,false);
const report=ctx.employeeReport_({},'valid');
for(const rows of [report.data.absences,report.data.lates,report.data.actions])assert.deepEqual(Array.from(rows,x=>x.marker),['own','alias']);
// Deleted employees invalidate old bearer links.
const original=vm.createContext({});vm.runInContext(source,original);
original.portalTokenSheet_=()=>({getDataRange:()=>({getDisplayValues:()=>[['token','id','name','job','createdAt','reportType'],['old','deleted','Name','Job','','all']]})});
original.readSheet_=()=>[];
assert.equal(original.findPortalToken_({},'old'),null);
const html=fs.readFileSync(__dirname+'/index.html','utf8');
for(const script of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g))new vm.Script(script[1]);
new vm.Script(fs.readFileSync(__dirname+'/sw.js','utf8'));
assert.ok(html.includes("const portalMode=portalParams.get('admin')!=='1'"));
assert.ok(html.includes("if(!token){"));
assert.ok(html.includes("adminPin:adminCredential"));
console.log('PASS: unauthorized read/write matrix, authenticated admin check, per-employee report isolation, invalid tokens, deleted employees, and script syntax.');
