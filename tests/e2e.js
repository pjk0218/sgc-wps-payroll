// 화면 E2E: 실제 크롬에서 양식·마스터 등록 → HR 파일 투입 → 파일 만들기 → 내려받은 결과를 최종본과 대조
const { chromium } = require('/opt/npm-tools/node_modules/playwright'); const http=require('http'),fs=require('fs'),path=require('path');
const X=require('../app/xlsbin.js'),W=require('../app/wps.js');
const srv=http.createServer((q,s)=>{ const f=path.join('/home/claude/app',q.url==='/'?'index.html':decodeURIComponent(q.url.split('?')[0])); fs.readFile(f,(e,d)=>{ if(e){s.writeHead(404);s.end();} else {s.writeHead(200,{'Content-Type':f.endsWith('.html')?'text/html':'text/javascript'});s.end(d);} }); }).listen(8765);
(async()=>{
  const b=await chromium.launch({executablePath:'/usr/bin/google-chrome',args:['--no-sandbox']}); const ctx=await b.newContext({acceptDownloads:true}); const p=await ctx.newPage();
  const errs=[]; p.on('pageerror',e=>errs.push(e.message)); p.on('console',m=>{ if(m.type()==='error') errs.push(m.text()); });
  await p.goto('http://localhost:8765/'); await p.waitForSelector('nav button');
  await p.click('nav button[data-tab="tpl"]'); await p.setInputFiles('input[data-file="tpl"]','/tmp/e2e/acc.zip'); await p.waitForFunction(()=>Object.keys(window.__wps.templates).length===8);
  console.log('양식 등록:',await p.evaluate(()=>Object.entries(window.__wps.templates).map(([k,v])=>k+':'+v.rows+(v.std?'':'(비표준)')).join(', ')));
  await p.click('nav button[data-tab="master"]'); await p.setInputFiles('input[data-file="master"]','/tmp/e2e/master.xls'); await p.waitForFunction(()=>Object.keys(window.__wps.master).length>300);
  console.log('마스터:',await p.evaluate(()=>Object.keys(window.__wps.master).length),'명');
  const sp=JSON.parse(fs.readFileSync('/home/claude/secret-special.json','utf8'));
  for(const [id,k] of Object.entries(sp)){ await p.fill('#sp_id',id); await p.selectOption('#sp_out',k); await p.click('button[data-act="spAdd"]'); await p.waitForTimeout(300); }
  await p.click('nav button[data-tab="tpl"]'); await p.fill('#exp_path','Z:\\{month}\\Accounting'); await p.click('button[data-act="expSave"]'); await p.waitForTimeout(300);
  console.log('Export 폴더 설정:',await p.evaluate(()=>window.__wps.exportPath));
  console.log('별도 파일 대상자:',await p.evaluate(()=>Object.keys(window.__wps.special).length),'명');
  await p.click('nav button[data-tab="run"]');
  await p.fill('input[data-act="month"]','2026-10'); await p.dispatchEvent('input[data-act="month"]','change');
  await p.fill('input[data-act="runDate"]','2026-10-20'); await p.dispatchEvent('input[data-act="runDate"]','change');
  await p.setInputFiles('input[data-file="hr"]',fs.readdirSync('/tmp/e2e/in').map(f=>'/tmp/e2e/in/'+f)); await p.waitForFunction(()=>window.__wps.inputs.length===9);
  console.log('Value Date:',await p.inputValue('input[data-act="valueDate"]'),'| 분류:',await p.evaluate(()=>window.__wps.inputs.map(i=>i.output).join(',')));
  console.log('알림:',await p.evaluate(()=>{const c={};window.__wps.res.notices.forEach(n=>c[n.type]=(c[n.type]||0)+1);return JSON.stringify(c);}));
  console.log('만들기 버튼 비활성(오류 확인 전):',await p.isDisabled('button[data-act="create"]'));
  await p.screenshot({path:'/tmp/e2e/run.png',fullPage:true});
  if (await p.$('input[data-act="ack"]')) await p.check('input[data-act="ack"]'); await p.click('button[data-act="create"]'); await p.waitForFunction(()=>window.__wps.created,{timeout:60000});
  const [dl]=await Promise.all([p.waitForEvent('download'),p.click('button[data-act="dlAll"]')]); await dl.saveAs('/tmp/e2e/out.zip');
  console.log('마스터(신규 반영 후):',await p.evaluate(()=>Object.keys(window.__wps.master).length),'명');
  await p.screenshot({path:'/tmp/e2e/done.png',fullPage:true});
  // 다음 달 시나리오: 같은 파일을 11월로 다시 → 신규 0명, 전월 대비 0
  await p.fill('input[data-act="month"]','2026-11'); await p.dispatchEvent('input[data-act="month"]','change'); await p.waitForTimeout(300);
  console.log('11월 재실행 알림:',await p.evaluate(()=>{const c={};window.__wps.res.notices.forEach(n=>c[n.type]=(c[n.type]||0)+1);return JSON.stringify(c);}),'| 전월 자료 인식:',await p.evaluate(()=>!!window.__wps.prev));
  await p.click('nav button[data-tab="hist"]'); await p.waitForSelector('#hist details'); console.log('월별 기록:',(await p.textContent('#hist summary')).slice(0,90));
  for(const tb of ['master','tpl','help']){ await p.click(`nav button[data-tab="${tb}"]`); } await p.click('nav button[data-tab="run"]');
  console.log('페이지 오류:',errs.length?errs:'없음');
  await b.close(); srv.close();
  // 내려받은 zip 검증
  const files=await X.unzip(new Uint8Array(fs.readFileSync('/tmp/e2e/out.zip'))); const ACC='/home/claude/x/acc/Accoutning Team_WPS File/'; let bad=0;
  for(const [n,d] of Object.entries(files)){ const k=W.detectTemplate(n); const ref=ACC+fs.readdirSync(ACC).find(f=>{let kk=W.detectTemplate(f); if(/^05\.1/.test(f)) kk='pdh-trainee'; return kk===k;});
    const a=W.readEmployees(X.readXls(new Uint8Array(fs.readFileSync(ref)),W.SHEET)), c=W.readEmployees(X.readXls(d,W.SHEET));
    const same=a.employees.length===c.employees.length&&a.employees.every((e,i)=>e.id===c.employees[i].id&&e.iban===c.employees[i].iban&&Math.abs(e.salary-c.employees[i].salary)<1e-9&&Math.abs(e.basic-c.employees[i].basic)<1e-9&&Math.abs(e.deduction-c.employees[i].deduction)<1e-9)&&a.header.total===c.header.total&&String(a.header.valueDate)===String(c.header.valueDate)&&a.header.count===c.header.count;
    if(!same) bad++; const f18=X.readXls(d,W.SHEET).cells.get('17,5'); console.log(' ',same?'일치':'불일치',n.slice(0,40),c.employees.length,'명',c.header.total,'F18=',f18); }
  console.log('브라우저 산출물 vs 회계팀 최종본(사번·IBAN·금액·합계·인원·Value Date):',bad?bad+'개 불일치':'8개 모두 일치');
})().catch(e=>{console.error('E2E 실패',e);process.exit(1);});
