// 교차 검증: 인원이 양식보다 많을 때(수식 복제)와 적을 때(남는 행 정리)
const X=require('../app/xlsbin.js'),W=require('../app/wps.js'),fs=require('fs');
const ACC='/home/claude/x/acc/Accoutning Team_WPS File/'; const rd=(f)=>new Uint8Array(fs.readFileSync(f));
const find=(p)=>ACC+fs.readdirSync(ACC).find(f=>f.startsWith(p));
const emp=(f)=>W.readEmployees(X.readXls(rd(f),W.SHEET)).employees;
const mk=(src)=>{const rows=emp(src).map(e=>({id:e.id,name:e.name,bank:e.bank,iban:e.iban,salary:e.salary,basic:e.basic,housing:e.housing,other:e.other,deduction:e.deduction,desc:'Salary for November 2026',addr1:e.addr1,addr2:e.addr2,addr3:e.addr3})); return {key:'x',rows,count:rows.length,total:W.r2(rows.reduce((a,r)=>a+r.salary,0))};};
const cases=[['PDH 116명 → trainee 양식(1명)',find('05 '),find('05.1')],['trainee 1명 → PDH 양식(116명)',find('05.1'),find('05 ')],['LNIC 33명 → Khobar 양식(15열)',find('06'),find('01')],['APOC 71명 → Rakan 양식',find('03'),find('04.1')]];
cases.forEach(([nm,src,t],i)=>{ const o=mk(src); const r=W.fill(rd(t),o,'20261125'); fs.writeFileSync(`/tmp/cross${i}.xls`,r.bytes); console.log(nm,'→',o.count,'명',o.total,'| 수식 유지',r.info.formulaKept,'복제',r.info.formulaCloned,'실패',r.info.formulaFail,'| 자체검증',JSON.stringify(r.check)); });
