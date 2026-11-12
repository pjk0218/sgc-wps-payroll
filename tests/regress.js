// 회귀 검증: 10월 HR 원본 + 9월 마스터 + 양식 → 생성 결과를 회계팀 10월 최종본과 셀 단위 대조
const X=require('../app/xlsbin.js'),W=require('../app/wps.js'),fs=require('fs'),path=require('path'),cr=require('crypto');
const U='/mnt/user-data/uploads/', HR='/home/claude/x/hr/HR/', ACC='/home/claude/x/acc/Accoutning Team_WPS File/';
const rd=(f)=>new Uint8Array(fs.readFileSync(f));
(async()=>{
  const hrFiles=fs.readdirSync(HR).map(f=>HR+f).concat(fs.readdirSync(U).filter(f=>/SIPCHEM_EVA/.test(f)).map(f=>U+f));
  const master=W.readMasterSheet(X.readXls(rd(U+fs.readdirSync(U).find(f=>/Master/.test(f))),W.SHEET));
  const inputs=[]; for(const f of hrFiles){ const sh=await X.readAny(rd(f),W.SHEET); const {employees,header}=W.readEmployees(sh); inputs.push({fileName:path.basename(f),output:W.detectOutput(path.basename(f)),employees,header}); }
  console.log('입력 분류:'); inputs.forEach(i=>console.log('  ',String(i.output).padEnd(12),i.employees.length,'명 ',i.fileName));
  const vd=W.valueDateFor(new Date(2026,9,20)); console.log('Value Date(10/20 화 작성):',vd, '| 목요일 11/12 →',W.valueDateFor(new Date(2026,10,12)));
  const res=W.build(inputs,{year:2026,month:10,master,special:JSON.parse(fs.readFileSync('/home/claude/secret-special.json','utf8'))});
  const tmpl={}; for(const f of fs.readdirSync(ACC)){ let k=W.detectTemplate(f); if(/^05\.1/.test(f)) k='pdh-trainee'; tmpl[k]=ACC+f; }
  const nonWb=(u)=>X.cfbRead(u).entries.filter(e=>e.type===2&&e.path!=='Workbook').map(e=>e.path+':'+cr.createHash('md5').update(e.data).digest('hex')).sort().join('|');
  fs.mkdirSync('/home/claude/out',{recursive:true}); let allDiff=0;
  for(const o of res.outputs){
    const t=rd(tmpl[o.key]); const r=W.fill(t,o,vd); fs.writeFileSync('/home/claude/out/'+o.fileName,r.bytes);
    const a=X.readXls(t,W.SHEET), b=X.readXls(r.bytes,W.SHEET); const keys=new Set([...a.cells.keys(),...b.cells.keys()]); const diffs=[];
    for(const k of keys){ const x=a.cells.get(k), y=b.cells.get(k); if(x===y) continue; if(typeof x==='number'&&typeof y==='number'&&Math.abs(x-y)<1e-9) continue; const [rr,cc]=k.split(',').map(Number); diffs.push(String.fromCharCode(65+cc)+(rr+1)+': 최종본='+JSON.stringify(x)+' / 생성='+JSON.stringify(y)); }
    allDiff+=diffs.length;
    console.log(`\n[${o.key}] ${o.count}명 합계 ${o.total} | 수식 유지 ${r.info.formulaKept}·복제 ${r.info.formulaCloned}·실패 ${r.info.formulaFail} | 최종본과 다른 셀 ${diffs.length}개 | 매크로·버튼 스트림 동일 ${nonWb(t)===nonWb(r.bytes)} | 매크로 표준 위치 ${r.macroStandard}`);
    diffs.slice(0,8).forEach(d=>console.log('    ',d));
    try{ const pa=W.macroPreview(t), pb=W.macroPreview(r.bytes); console.log('     매크로 출력(txt) 미리보기: 최종본과 동일',pa===pb, '| 줄 수',pb.split('\r\n').length-1);}catch(e){console.log('     매크로 출력 미리보기 실패:',e.message);}
  }
  const by={}; res.notices.forEach(n=>{by[n.type]=(by[n.type]||0)+1;}); console.log('\n알림 집계:',by,'| 전체 차이 셀',allDiff);
  res.notices.filter(n=>n.type!=='new').slice(0,20).forEach(n=>console.log('  ',n.level,n.output,n.id,n.name,'—',n.msg));
})();
