const vm=require('vm'),fs=require('fs'),assert=require('assert/strict');
const html=fs.readFileSync(__dirname+'/index.html','utf8');
const code=html.slice(html.indexOf('let sharedCodesLoading=false;'),html.indexOf('function clearPortalLinks(){'));
const run=async()=>{
 let calls=0,rendered=0,visible=false,fail=false,release;
 const elements={sharedCodesStatus:{textContent:''},sharedCodesResult:{querySelector:()=>true,textContent:'old rows'}};
 const ctx=vm.createContext({_g:id=>elements[id],setSharedCodesVisibility:hidden=>{visible=!hidden},readCachedSharedSigningCodes:()=>null,sharedAdminPin:()=> 'pin',activeSharedSigningCodes:x=>x,renderSharedSigningCodes:()=>{rendered++},window:{location:{origin:'https://example.test',pathname:'/'}},apiPost:async()=>{calls++;await new Promise(r=>release=r);return fail?{ok:false,error:'Request failed'}:{ok:true,codes:[]}}});
 vm.runInContext(code,ctx);
 const pending=ctx.loadSharedSigningCodes(true);await ctx.loadSharedSigningCodes(true);assert.equal(calls,1);assert.equal(visible,true);release();await pending;assert.equal(rendered,1);
 fail=true;const retry=ctx.loadSharedSigningCodes(true);release();await retry;assert.equal(elements.sharedCodesResult.textContent,'old rows');assert.equal(elements.sharedCodesStatus.textContent,'Request failed');
 console.log('PASS: forced refresh reveals list, concurrent requests are prevented, successful result renders, failed request preserves existing list, retry unlocks.');
};run().catch(e=>{console.error(e);process.exit(1)});
