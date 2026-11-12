// 엔진 단위 검증: (1) 무수정 재기록 (2) 값 수정·행 추가 후 재읽기 (3) 매크로 스트림 바이트 동일성
const X=require('../app/xlsbin.js'),fs=require('fs'),cr=require('crypto');
const f=process.argv[2]; const src=new Uint8Array(fs.readFileSync(f));
const md=(u)=>cr.createHash('md5').update(u).digest('hex').slice(0,8);
const nonWb=(u)=>X.cfbRead(u).entries.filter(e=>e.type===2&&e.path!=='Workbook').map(e=>e.path+':'+md(e.data)).sort().join('|');
const a=X.readXls(src,'Payroll Information');
// 1) 수정 없이 통과
const o1=X.patchXls(src,'Payroll Information',[]); const b=X.readXls(o1,'Payroll Information');
let same=a.cells.size===b.cells.size; for(const [k,v] of a.cells) if(b.cells.get(k)!==v) same=false;
console.log('무수정 재기록: 셀 동일',same,'| 매크로 등 스트림 동일',nonWb(src)===nonWb(o1));
// 2) 수정
const edits=[{r:24,c:1,v:'테스트 이름 ΩΩ'},{r:24,c:4,v:1234.56},{r:25,c:1,v:null},{r:300,c:0,v:9999999999},{r:300,c:1,v:'New Person Added'},{r:17,c:2,v:77}];
const o2=X.patchXls(src,'Payroll Information',edits,()=>80); const c=X.readXls(o2,'Payroll Information');
let ok=true; for(const e of edits){const v=c.cells.get(e.r+','+e.c); if((e.v===null?undefined:e.v)!==v){ok=false;console.log('불일치',e,v);}}
let others=0; for(const [k,v] of a.cells){ if(edits.some(e=>e.r+','+e.c===k)) continue; if(c.cells.get(k)!==v) others++; }
const s2=X.readXls(o2,'Bank Code'), s0=X.readXls(src,'Bank Code'); let bank=s2.cells.size===s0.cells.size; for(const [k,v] of s0.cells) if(s2.cells.get(k)!==v) bank=false;
console.log('수정 반영',ok,'| 다른 셀 변동',others,'| 둘째 시트 동일',bank,'| 매크로 스트림 동일',nonWb(src)===nonWb(o2),'| 크기',src.length,'→',o2.length);
fs.writeFileSync(process.argv[3]||'/tmp/t1.xls',o2);
