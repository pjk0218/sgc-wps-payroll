/* wps.js — 급여 WPS 가공 규칙 (HR 원본 → 회계팀 은행 양식)
   규칙: 금액은 HR, 이름·은행·계좌는 마스터. 마스터에 없으면 HR 값으로 만들고 알림. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./xlsbin.js'));
  else root.Wps = factory(root.XlsBin);
})(typeof self !== 'undefined' ? self : this, function (X) {
  'use strict';
  const SHEET = 'Payroll Information';
  const OUTPUTS = [
    { key: 'khobar', no: '01', tag: 'Khobar', tail: '', label: 'Khobar (Office Staff)' },
    { key: 'sepc', no: '02', tag: 'SEPC', tail: '', label: 'SEPC' },
    { key: 'apoc', no: '03', tag: 'APOC', tail: '', label: 'APOC' },
    { key: 'eva', no: '04', tag: 'EVA', tail: '', label: 'EVA' },
    { key: 'eva-rakan', no: '04.1', tag: 'EVA', tail: ' -Rakan', label: 'EVA – D360 bank' },
    { key: 'pdh', no: '05', tag: 'PDH', tail: '', label: 'PDH' },
    { key: 'pdh-trainee', no: '05.1', tag: 'PDH', tail: ' -Trainee', label: 'PDH – Trainee' },
    { key: 'lnic', no: '06', tag: 'LNIC', tail: '', label: 'LNIC' },
  ];
  // 항상 별도 파일로 빼는 사람(사번 → 출력)은 코드에 두지 않고 공유 폴더 settings.json에 둔다 (개인 신분번호 공개 방지)
  const DEFAULT_SPECIAL = {};
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const FIELDS = ['id', 'name', 'bank', 'iban', 'salary', 'basic', 'housing', 'other', 'deduction', 'desc', 'addr1', 'addr2', 'addr3'];
  const HEAD = { id: /^emp\.?\s*id/i, name: /^emp\.?\s*name/i, bank: /^emp\.?\s*bank/i, iban: /^emp\.?\s*acc/i, salary: /^salary\s*amount/i, basic: /^basic/i, housing: /^housing/i, other: /^other/i, deduction: /^deduction/i, desc: /^payment\s*desc/i, addr1: /adre+s+\s*1/i, addr2: /adre+s+\s*2/i, addr3: /adre+s+\s*3/i };

  function detectOutput(fileName) {
    const n = fileName.toLowerCase();
    if (/rakan/.test(n)) return 'eva-rakan';
    if (/trainee/.test(n)) return 'pdh-trainee';
    if (/office|khobar/.test(n)) return 'khobar';
    if (/lnic/.test(n)) return 'lnic';
    if (/sepc/.test(n)) return 'sepc';
    if (/apoc/.test(n)) return 'apoc';
    if (/\beva\b|\(eva\)|_eva|eva[ _-]/.test(n)) return 'eva';
    if (/pdh/.test(n)) return 'pdh';
    return null;
  }
  // 회계팀 최종본(양식) 파일명 → 출력 종류
  function detectTemplate(fileName) {
    const n = fileName.toLowerCase(); const m = /^\s*(\d\d(?:\.\d)?)\b/.exec(n);
    if (/rakan/.test(n)) return 'eva-rakan'; if (/trainee/.test(n)) return 'pdh-trainee';
    if (m) { const o = OUTPUTS.find((x) => x.no === m[1]); if (o && n.includes(o.tag.toLowerCase())) return o.key; }
    return detectOutput(fileName);
  }
  const normId = (v) => { if (typeof v === 'number') return String(Math.round(v)); return String(v === undefined || v === null ? '' : v).replace(/\s+/g, '').replace(/\.0$/, ''); };
  const clean = (v) => String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim();
  const cleanIban = (v) => String(v === undefined || v === null ? '' : v).replace(/\s+/g, '').toUpperCase();
  const r2 = (x) => Math.round((x + Number.EPSILON) * 100) / 100;

  // 시트 구조 파악: 헤더 행, 열 위치, 상단 정보 칸
  function layout(sheet) {
    const get = (r, c) => sheet.cells.get(r + ',' + c);
    let hr = -1; for (let r = 0; r <= Math.min(sheet.maxR, 80) && hr < 0; r++) if (typeof get(r, 0) === 'string' && HEAD.id.test(get(r, 0).trim())) hr = r;
    if (hr < 0) throw new Error("Header row 'Emp. ID. No.' not found — is this a BSF-WPS payroll file?");
    const cols = {}; for (let c = 0; c <= sheet.maxC; c++) { const h = get(hr, c); if (typeof h !== 'string') continue; for (const k of FIELDS) if (cols[k] === undefined && HEAD[k].test(h.trim())) { cols[k] = c; break; } }
    for (const k of FIELDS) if (cols[k] === undefined) throw new Error('Header column missing: ' + k);
    const meta = {};
    for (let r = 0; r < hr; r++) for (let c = 0; c <= sheet.maxC; c++) { const v = get(r, c); if (typeof v !== 'string') continue; if (/value\s*date/i.test(v)) meta.valueDate = [r, c + 1]; else if (/number\s*of/i.test(v)) meta.count = [r, c + 1]; else if (/total\s*salary/i.test(v)) meta.total = [r, c + 1]; else if (/start/i.test(v) && /\(/.test(v)) meta.start = [r, c + 1]; }
    const c0 = Math.min(...FIELDS.map((k) => cols[k])), c1 = Math.max(...FIELDS.map((k) => cols[k]));
    // 매크로는 고정 위치(C16·C18·C20·F16, A~M열)를 읽는다 → 양식이 그 위치와 맞는지
    const std = FIELDS.every((k, i) => cols[k] === i) && hr === 23 && meta.valueDate && meta.valueDate[0] === 15 && meta.valueDate[1] === 5 && meta.count && meta.count[0] === 17 && meta.count[1] === 2 && meta.total && meta.total[0] === 19 && meta.total[1] === 2;
    return { hr, cols, meta, c0, c1, macroStandard: !!std };
  }
  function readEmployees(sheet, lay) {
    lay = lay || layout(sheet); const get = (r, c) => sheet.cells.get(r + ',' + c), out = [];
    for (let r = lay.hr + 1; r <= sheet.maxR; r++) {
      const id = normId(get(r, lay.cols.id)), name = clean(get(r, lay.cols.name)); if (!id && !name) continue;
      const e = { row: r + 1, id, name, bank: clean(get(r, lay.cols.bank)), iban: cleanIban(get(r, lay.cols.iban)), rawBank: get(r, lay.cols.bank), desc: clean(get(r, lay.cols.desc)), addr1: clean(get(r, lay.cols.addr1)), addr2: clean(get(r, lay.cols.addr2)), addr3: clean(get(r, lay.cols.addr3)), bad: [] };
      for (const k of ['salary', 'basic', 'housing', 'other', 'deduction']) { const v = get(r, lay.cols[k]); if (v === undefined || v === '' || (typeof v === 'string' && !v.trim())) e[k] = 0; else if (typeof v === 'number') e[k] = v; else { const n = Number(String(v).replace(/,/g, '')); if (isFinite(n)) e[k] = n; else { e[k] = 0; e.bad.push(k + '="' + v + '"'); } } }
      out.push(e);
    }
    const v = (p) => (p ? get(p[0], p[1]) : undefined);
    return { employees: out, header: { valueDate: v(lay.meta.valueDate), count: v(lay.meta.count), total: v(lay.meta.total) } };
  }
  function readMasterSheet(sheet) {
    const lay = layout(sheet), get = (r, c) => sheet.cells.get(r + ',' + c), m = {};
    for (let r = lay.hr + 1; r <= sheet.maxR; r++) { const id = normId(get(r, lay.cols.id)); if (!id) continue; const s = (c) => { const v = get(r, c); return v === undefined ? '' : String(v); }; m[id] = { name: s(lay.cols.name), bank: s(lay.cols.bank), iban: s(lay.cols.iban) }; }
    return m;
  }
  // 집행일(작성일) → Value Date: 다음 날. 목요일이면 다음 주 일요일 (금·토는 휴일)
  function valueDateFor(d) {
    const x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); const dow = x.getDay();
    x.setDate(x.getDate() + (dow === 4 ? 3 : dow === 5 ? 2 : 1));
    return '' + x.getFullYear() + String(x.getMonth() + 1).padStart(2, '0') + String(x.getDate()).padStart(2, '0');
  }
  function shorten(name, max) { if (name.length <= max) return name; const w = name.split(' '); while (w.length > 1 && w.join(' ').length > max) w.pop(); return w.join(' ').slice(0, max); }

  /* inputs: [{fileName, output, employees}]  →  출력별 행 + 알림
     opts: {year, month(1~12), master, special, overrides:{id:{name,bank,iban,useHr}}} */
  function build(inputs, opts) {
    const master = opts.master || {}, special = opts.special || DEFAULT_SPECIAL, ov = opts.overrides || {};
    const desc = 'Salary for ' + MONTHS[opts.month - 1] + ' ' + opts.year, monthName = MONTHS[opts.month - 1].toLowerCase();
    const groups = {}; const notices = []; const seen = new Map();
    const note = (level, type, output, e, msg, extra) => notices.push(Object.assign({ level, type, output, id: e ? e.id : '', name: e ? e.name : '', msg }, extra || {}));
    for (const inp of inputs) {
      if (!inp.output) { note('error', 'file', null, null, `Choose the output file for "${inp.fileName}"`); continue; }
      let wrongMonth = 0;
      for (const e of inp.employees) {
        const out = special[e.id] || inp.output; (groups[out] = groups[out] || []);
        const m = master[e.id], o = ov[e.id] || {}; const row = { id: e.id, src: inp.fileName, srcRow: e.row, inMaster: !!m };
        if (m) {
          const mn = String(m.name).trim(); row.name = mn.length > 35 ? shorten(clean(mn), 35) : mn; row.bank = clean(m.bank); row.iban = cleanIban(m.iban);
          if (String(m.name).trim().length > 35) note('info', 'name-short', out, e, `Name was longer than 35 characters and was shortened to "${row.name}"`);
          const hrI = e.iban, mI = cleanIban(m.iban), hrB = e.bank.toUpperCase(), mB = clean(m.bank).toUpperCase();
          if (hrI && hrI !== mI) { note('error', 'iban-diff', out, e, `IBAN differs from the master list — HR: ${hrI} / master: ${mI}`, { hr: hrI, master: mI }); if (o.useHr) row.iban = hrI; }
          if (hrB && hrB !== mB) { note('warn', 'bank-diff', out, e, `Bank code differs from the master list — HR: ${e.bank} / master: ${clean(m.bank)}`, { hr: e.bank, master: clean(m.bank) }); if (o.useHr) row.bank = e.bank; }
        } else {
          row.name = o.name !== undefined ? clean(o.name) : shorten(e.name, 35); row.bank = o.bank !== undefined ? clean(o.bank) : e.bank; row.iban = o.iban !== undefined ? cleanIban(o.iban) : e.iban;
          note('new', 'new', out, e, 'Not in the master list (written from HR values)', { hrName: e.name, row });
          if (e.name.length > 35 && o.name === undefined) note('info', 'name-short', out, e, `Name was longer than 35 characters and was shortened to "${row.name}"`);
        }
        for (const k of ['salary', 'basic', 'housing', 'other', 'deduction']) row[k] = e[k];
        row.desc = desc; row.addr1 = e.addr1; row.addr2 = e.addr2; row.addr3 = e.addr3;
        // 검증
        if (!/^\d{10}$/.test(e.id)) note('error', 'id', out, e, `ID is not a 10-digit number: "${e.id}"`);
        const ib = cleanIban(row.iban); if (!/^SA[0-9A-Z]{22}$/.test(ib)) note('error', 'iban', out, e, `IBAN is not 24 characters starting with SA: "${row.iban}"`);
        if (!clean(row.bank)) note('error', 'bank', out, e, 'Bank code is empty');
        if (!clean(row.name)) note('error', 'name', out, e, 'Name is empty');
        if (String(row.name).length > 35) note('error', 'name-long', out, e, 'Name is longer than 35 characters (the Export macro fails)');
        for (const k of ['addr1', 'addr2', 'addr3', 'bank']) if (String(row[k]).length > 35) note('error', 'len', out, e, `${k} is longer than 35 characters`);
        if (e.bad.length) note('error', 'amount', out, e, 'An amount cell is not a number: ' + e.bad.join(', '));
        if (!(row.salary > 0)) note('warn', 'zero', out, e, `Salary amount is zero or negative (${row.salary})`);
        const calc = row.basic + row.housing + row.other - row.deduction, tr = Math.trunc(calc * 100 + 1e-7) / 100; if (Math.abs(tr - row.salary) > 0.005 && Math.abs(calc - row.salary) > 0.005) note('error', 'sum', out, e, `Salary amount ${row.salary} ≠ basic + housing + other − deduction ${r2(calc)} — the salary cell in the template is a formula, so Excel will change this value when it recalculates`);
        const tagOf = (k) => (OUTPUTS.find((x) => x.key === k) || {}).tag || k;
        if (seen.has(e.id)) { const pv = seen.get(e.id); if (pv.out === out) note('error', 'dup', out, e, `Same ID appears twice in the same output file (${pv.file} / ${inp.fileName})`); else note('warn', 'dup-move', out, e, `Same person is in two projects — ${tagOf(pv.out)} ${r2(pv.salary)} SAR and ${tagOf(out)} ${r2(row.salary)} SAR. Normal if the person moved project; otherwise check for a double payment`); } else seen.set(e.id, { file: inp.fileName, out, salary: row.salary });
        const dm = MONTHS.find((mm) => e.desc.toLowerCase().includes(mm.toLowerCase())); if (dm && dm.toLowerCase() !== monthName) wrongMonth++;
        groups[out].push(row);
      }
      if (wrongMonth) note('warn', 'month', inp.output, null, `Payment description in "${inp.fileName}" is not ${MONTHS[opts.month - 1]} (${wrongMonth} rows). Check that this is not another month's file`);
      if (inp.header && typeof inp.header.total === 'number') { const s = r2(inp.employees.reduce((a, e) => a + e.salary, 0)); if (Math.abs(s - inp.header.total) > 0.05) note('warn', 'hr-total', inp.output, null, `Total Salary in the header of "${inp.fileName}" (${inp.header.total}) differs from the sum of rows (${s})`); }
    }
    const outputs = OUTPUTS.filter((o) => groups[o.key]).map((o) => ({ key: o.key, label: o.label, rows: groups[o.key], count: groups[o.key].length, total: r2(groups[o.key].reduce((a, r) => a + r.salary, 0)), fileName: `${o.no} ★ BSF-WPS Conversion Tool (${o.tag})-${MONTHS[opts.month - 1]} ${opts.year}${o.tail}.xls` }));
    return { outputs, notices, desc };
  }
  // 양식 + 행 → 완성 .xls (매크로 보존)
  function fill(templateU8, out, valueDate) {
    const sheet = X.readXls(templateU8, SHEET), lay = layout(sheet), first = lay.hr + 1, n = out.rows.length;
    // 열별 기준 서식: 양식 데이터 행에서 가장 많이 쓰인 XF
    const refXf = {}; for (let c = lay.c0; c <= lay.c1; c++) { const cnt = new Map(); for (let r = first; r <= sheet.maxR; r++) { if (sheet.cells.get(r + ',' + lay.cols.id) === undefined) continue; const x = sheet.xfs.get(r + ',' + c); if (x !== undefined) cnt.set(x, (cnt.get(x) || 0) + 1); } let best, bn = 0; for (const [x, k] of cnt) if (k > bn) { best = x; bn = k; } refXf[c] = best === undefined ? (sheet.xfs.get(first + ',' + c) ?? 15) : best; }
    const edits = new Map(); const put = (r, c, v) => edits.set(r + ',' + c, { r, c, v });
    const isF = (r, c) => sheet.formulas.has(r + ',' + c), sc = lay.cols.salary;
    // 이전 달 데이터 비우기 (수식 칸은 수식을 남기고, 남는 행의 결과값만 0으로)
    for (let r = first; r <= sheet.maxR; r++) for (let c = lay.c0; c <= lay.c1; c++) { const v = sheet.cells.get(r + ',' + c); if (isF(r, c)) { if (c === sc && r >= first + n && typeof v === 'number' && v !== 0) edits.set(r + ',' + c, { r, c, v: 0, f: 'keep' }); } else if (v !== undefined) put(r, c, null); }
    const salaryIsFormula = isF(first, sc);
    out.rows.forEach((row, i) => { const r = first + i;
      put(r, lay.cols.id, /^\d+$/.test(row.id) ? Number(row.id) : row.id); put(r, lay.cols.name, row.name); put(r, lay.cols.bank, row.bank); put(r, lay.cols.iban, row.iban);
      for (const k of ['basic', 'housing', 'other', 'deduction']) put(r, lay.cols[k], row[k]);
      edits.set(r + ',' + sc, salaryIsFormula ? { r, c: sc, v: row.salary, f: [first, sc] } : { r, c: sc, v: row.salary });
      put(r, lay.cols.desc, row.desc); put(r, lay.cols.addr1, row.addr1); put(r, lay.cols.addr2, row.addr2); put(r, lay.cols.addr3, row.addr3); });
    if (!lay.meta.count || !lay.meta.total || !lay.meta.valueDate) throw new Error('Template header cells (number of employees / total / value date) not found');
    const putK = (p, v) => edits.set(p[0] + ',' + p[1], { r: p[0], c: p[1], v, f: 'keep' });
    putK(lay.meta.count, n); putK(lay.meta.total, out.total); put(lay.meta.valueDate[0], lay.meta.valueDate[1], String(valueDate));
    if (lay.meta.start) putK(lay.meta.start, first + 1);
    const bytes = X.patchXls(templateU8, SHEET, [...edits.values()], (r, c) => refXf[c] ?? 15);
    // 자체 검증: 만든 파일을 다시 읽어 행·합계·상단 값 확인
    const back = readEmployees(X.readXls(bytes, SHEET));
    const sum = r2(back.employees.reduce((a, e) => a + e.salary, 0));
    const ok = back.employees.length === n && Math.abs(sum - out.total) < 0.005 && back.header.count === n && Math.abs(back.header.total - out.total) < 0.005 && String(back.header.valueDate) === String(valueDate) && back.employees.every((e, i) => e.id === out.rows[i].id && cleanIban(e.iban) === cleanIban(out.rows[i].iban) && Math.abs(e.salary - out.rows[i].salary) < 1e-9);
    if (!ok) throw new Error(`Self-check failed (${out.key}): the created file does not match the calculated values. Nothing was saved`);
    return { bytes, macroStandard: lay.macroStandard, info: bytes.info, salaryIsFormula, check: { count: back.employees.length, total: sum } };
  }
  // 매크로(Export 버튼)가 만들 은행 텍스트 본문을 미리 계산 — 대조·검증용
  function macroPreview(xlsU8) {
    const s = X.readXls(xlsU8, SHEET), g = (r, c) => { const v = s.cells.get((r - 1) + ',' + (c - 1)); return v === undefined ? '' : v; };
    const str = (v) => (typeof v === 'number' ? String(v) : String(v)); const pad = (v, n) => { const t = str(v); if (t.length > n) throw new Error(`"${t}" is longer than ${n} characters; the macro would fail`); return t + ' '.repeat(n - t.length); };
    const amt = (v) => (typeof v === 'number' ? v : Number(v) || 0).toFixed(2).replace('.', ',');
    const first = Number(g(16, 3)), lastRow = first + Number(g(18, 3)) - 1, lines = [];
    lines.push(['BSFR', pad(g(13, 7), 6), pad(g(14, 3), 24), 'SAR', str(g(16, 6)), amt(g(20, 3)), str(g(16, 6)), '<작성시각>', '      ', pad(g(14, 7), 18)].join('\t'));
    for (let r = first; r <= lastRow; r++) lines.push([amt(g(r, 5)), pad(g(r, 4), 34), pad(g(r, 2), 35) + pad(g(r, 11), 35) + pad(g(r, 12), 35) + pad(g(r, 13), 35), pad(g(r, 3), 35), pad(g(r, 10), 35), '      ', amt(g(r, 6)), amt(g(r, 7)), amt(g(r, 8)), amt(g(r, 9)), pad(g(r, 1), 10), ' '.repeat(16), ' '.repeat(8), ' '.repeat(8)].join('\t'));
    lines.push('-'); return lines.join('\r\n') + '\r\n';
  }
  return { SHEET, OUTPUTS, MONTHS, DEFAULT_SPECIAL, detectOutput, detectTemplate, layout, readEmployees, readMasterSheet, valueDateFor, build, fill, macroPreview, normId, clean, cleanIban, r2 };
});
