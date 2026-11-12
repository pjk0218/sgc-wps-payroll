/* app.js — 화면과 저장. 급여 데이터는 이 브라우저와 선택한 공유 폴더 밖으로 나가지 않는다(서버 전송 없음). */
(function () {
  'use strict';
  const X = window.XlsBin, W = window.Wps, VERSION = '1.0.4';
  const $ = (s, el) => (el || document).querySelector(s);
  const esc = (s) => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = (n) => Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const pad2 = (n) => String(n).padStart(2, '0');

  // ───────── 언어 ─────────
  const T = {
    en: { tabRun: 'Monthly run', tabMaster: 'Master list', tabTpl: 'Bank templates', tabHist: 'History', tabHelp: 'How it works',
      folderNone: 'Storage: <b>this browser only</b> (not shared)', folderOn: 'Shared folder: <b>{0}</b>', connect: 'Connect shared folder', reconnect: 'Re-open shared folder', change: 'Change folder',
      s1: 'Month and value date', month: 'Payroll month', runDate: 'Execution date', valueDate: 'Value date (YYYYMMDD)', vdHint: 'Value date = next day after execution. Thursday → next Sunday.',
      s2: 'HR files', dropHr: 'Drop the HR files here (.xls, .xlsx or one .zip) — or click to choose', file: 'File', project: 'Output file', people: 'People', sum: 'Total (SAR)', ignore: '— ignore —', choose: '— choose —', remove: 'Remove', clear: 'Clear all',
      s3: 'Checks', noIssues: 'No issues found.', newTitle: 'Not in the master list ({0}) — file is still created from HR values', newHint: 'Check name (max 35 characters), bank code and IBAN. Ticked people are added to the master list when you create the files.', hrName: 'Name in HR file', name: 'Name', bank: 'Bank code', iban: 'IBAN', addMaster: 'Add to master',
      diffTitle: 'Bank account differs from the master list ({0})', useMaster: 'Keep master', useHr: 'Use HR value and update master', errTitle: 'Errors ({0})', warnTitle: 'Warnings ({0})',
      s4: 'Result', out: 'Output', tpl: 'Template', vsPrev: 'vs. previous month', tplMissing: 'missing', tplOk: 'ready', tplOdd: 'layout differs from macro', grand: 'Grand total', joined: 'joined', left: 'left', noPrev: 'no previous month saved',
      ack: 'I have checked the errors above and want to create the files anyway.', create: 'Create bank files', needTpl: 'Register a bank template for: {0}', needFiles: 'Add HR files first.', needVd: 'Value date must be 8 digits (YYYYMMDD).',
      done: 'Files created and saved under {0}.', dl: 'Download', dlAll: 'Download all (.zip)', selfOk: 'Self-check passed: each file was re-read and matches people count, total and value date.', masterAdded: '{0} people added to the master list.',
      mTitle: 'Master list', mCount: '{0} people', mSearch: 'Search ID, name, IBAN…', mImport: 'Import master .xls', mExport: 'Export CSV', mAdd: 'Add person', id: 'ID', actions: '', edit: 'Edit', save: 'Save', del: 'Delete', cancel: 'Cancel', mEmpty: 'The master list is empty. Import the master .xls file once to start.', mImported: 'Master imported: {0} people ({1} new, {2} updated).', confirmDel: 'Delete {0} from the master list?', idBad: 'ID must be 10 digits.', ibanBad: 'IBAN must be 24 characters starting with SA.',
      tTitle: 'Bank templates', tHint: 'A template is the accounting team’s final bank file (.xls with the Export macro). Only cell values are replaced; the macro, button and formatting stay byte-for-byte the same. Drop last month’s final files (or the .zip) once.', dropTpl: 'Drop the accounting final files here (.xls or .zip)', tPeople: 'rows in template', tUpload: 'Upload…', tStd: 'standard', tOddMsg: 'This template does not use the standard cell positions that the Export macro reads (salary must be in column E, value date in F16). The macro output from this file may be wrong. Upload a standard template (for example a copy of the SEPC file) for this output.',
      hTitle: 'Saved months', hEmpty: 'Nothing saved yet.', created: 'Created', files: 'Files',
      perm: 'Click to allow access to the shared folder again.', unsupported: 'This browser cannot open folders. Use Chrome or Edge on a PC to share with the team; until then data stays in this browser only.',
      tplSaved: '{0} template(s) registered.', tplUnknown: 'Could not tell which output this file is for: {0}', readFail: 'Could not read {0}: {1}', working: 'Working…' },
  };
  const lang = 'en';
  const t = (k, ...a) => { let s = (T[lang][k] !== undefined ? T[lang][k] : T.en[k]) || k; a.forEach((v, i) => { s = s.replace('{' + i + '}', v); }); return s; };
  const noticeText = (n) => esc(n.msg);

  // ───────── 저장소: 공유 폴더(File System Access) 또는 브라우저(IndexedDB) ─────────
  const idb = { db: null,
    async open() { if (this.db) return this.db; this.db = await new Promise((res, rej) => { const r = indexedDB.open('wps-app', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); return this.db; },
    async tx(mode, fn) { const db = await this.open(); return new Promise((res, rej) => { const s = db.transaction('kv', mode).objectStore('kv'); const q = fn(s); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
    get(k) { return this.tx('readonly', (s) => s.get(k)); }, set(k, v) { return this.tx('readwrite', (s) => s.put(v, k)); }, del(k) { return this.tx('readwrite', (s) => s.delete(k)); }, keys() { return this.tx('readonly', (s) => s.getAllKeys()); } };
  const localStore = { kind: 'local', name: '',
    async read(p) { const v = await idb.get('f:' + p); return v ? new Uint8Array(v) : null; },
    async write(p, u8) { await idb.set('f:' + p, u8); },
    async list(dir) { const pre = 'f:' + (dir ? dir + '/' : ''); const set = new Set(); for (const k of await idb.keys()) if (typeof k === 'string' && k.startsWith(pre)) set.add(k.slice(pre.length).split('/')[0]); return [...set].sort(); } };
  function folderStore(root) {
    const dirOf = async (parts, create) => { let d = root; for (const p of parts) d = await d.getDirectoryHandle(p, { create }); return d; };
    return { kind: 'folder', name: root.name,
      async read(p) { try { const parts = p.split('/'), f = parts.pop(); const d = await dirOf(parts, false); const h = await d.getFileHandle(f); return new Uint8Array(await (await h.getFile()).arrayBuffer()); } catch (e) { if (e.name === 'NotFoundError' || e.name === 'TypeMismatchError') return null; throw e; } },
      async write(p, u8) { const parts = p.split('/'), f = parts.pop(); const d = await dirOf(parts, true); const h = await d.getFileHandle(f, { create: true }); const w = await h.createWritable(); await w.write(u8); await w.close(); },
      async list(dir) { try { const d = await dirOf(dir ? dir.split('/') : [], false); const out = []; for await (const [n] of d.entries()) out.push(n); return out.sort(); } catch (e) { return []; } } };
  }
  const te = new TextEncoder(), td = new TextDecoder();
  const readJson = async (p, dflt) => { const u = await S.store.read(p); if (!u) return dflt; try { return JSON.parse(td.decode(u)); } catch (e) { return dflt; } };
  const writeJson = (p, o) => S.store.write(p, te.encode(JSON.stringify(o, null, 1)));

  // ───────── 상태 ─────────
  const now = new Date();
  const S = { special: {}, store: localStore, pendingHandle: null, tab: 'run', master: {}, templates: {}, inputs: [], overrides: {}, addNew: {}, res: null, ack: false, prev: null, created: null,
    month: now.getFullYear() + '-' + pad2(now.getMonth() + 1), runDate: now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate()), valueDate: W.valueDateFor(now), vdTouched: false, mq: '', mEdit: null };
  window.__wps = S;

  function toast(msg) { const el = $('#toast'); el.textContent = msg; el.classList.add('on'); clearTimeout(toast.t); toast.t = setTimeout(() => el.classList.remove('on'), 3500); }
  function download(name, u8, type) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([u8], { type: type || 'application/octet-stream' })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }

  async function loadShared() {
    S.master = await readJson('master.json', {});
    S.special = (await readJson('settings.json', {})).special || {};
    S.templates = {};
    for (const o of W.OUTPUTS) { const u = await S.store.read('templates/' + o.key + '.xls'); if (u) { try { S.templates[o.key] = tplInfo(u); } catch (e) { S.templates[o.key] = { error: e.message }; } } }
    await loadPrev();
  }
  function tplInfo(u8) { const sh = X.readXls(u8, W.SHEET), lay = W.layout(sh); return { bytes: u8, rows: W.readEmployees(sh, lay).employees.length, std: lay.macroStandard }; }
  const prevMonthOf = (ym) => { const [y, m] = ym.split('-').map(Number); return m === 1 ? (y - 1) + '-12' : y + '-' + pad2(m - 1); };
  async function loadPrev() { S.prev = await readJson(prevMonthOf(S.month) + '/run.json', null); }

  async function connectFolder(handle) {
    try {
      if (!handle) handle = await window.showDirectoryPicker({ mode: 'readwrite' });
      if ((await handle.requestPermission({ mode: 'readwrite' })) !== 'granted') return;
      S.store = folderStore(handle); S.pendingHandle = null; await idb.set('dirHandle', handle); await loadShared(); recompute(); render();
    } catch (e) { if (e.name !== 'AbortError') toast(e.message); }
  }

  // ───────── 파일 받기 ─────────
  async function expand(fileList) {
    const out = [];
    for (const f of fileList) {
      const u8 = new Uint8Array(await f.arrayBuffer());
      if (/\.zip$/i.test(f.name)) { const files = await X.unzip(u8); for (const [n, d] of Object.entries(files)) { const base = n.split('/').pop(); if (/\.xlsx?$/i.test(base) && !base.startsWith('~$') && !n.startsWith('__MACOSX')) out.push({ name: base, data: d }); } }
      else if (/\.xlsx?$/i.test(f.name)) out.push({ name: f.name, data: u8 });
    }
    return out;
  }
  async function addHr(fileList) {
    for (const f of await expand(fileList)) {
      try { const sh = await X.readAny(f.data, W.SHEET); const { employees, header } = W.readEmployees(sh); S.inputs = S.inputs.filter((i) => i.fileName !== f.name); S.inputs.push({ fileName: f.name, data: f.data, output: W.detectOutput(f.name), employees, header }); }
      catch (e) { toast(t('readFail', f.name, e.message)); }
    }
    S.created = null; recompute(); render();
  }
  async function addTemplates(fileList, forceKey) {
    let n = 0;
    for (const f of await expand(fileList)) {
      const key = forceKey || W.detectTemplate(f.name); if (!key) { toast(t('tplUnknown', f.name)); continue; }
      try { const info = tplInfo(f.data); await S.store.write('templates/' + key + '.xls', f.data); S.templates[key] = info; n++; } catch (e) { toast(t('readFail', f.name, e.message)); }
    }
    if (n) toast(t('tplSaved', n)); recompute(); render();
  }
  async function importMaster(fileList) {
    for (const f of await expand(fileList)) {
      try { const m = W.readMasterSheet(await X.readAny(f.data, W.SHEET)); const cur = await readJson('master.json', {}); let add = 0, upd = 0;
        for (const [id, v] of Object.entries(m)) { if (!cur[id]) add++; else if (cur[id].name !== v.name || cur[id].bank !== v.bank || cur[id].iban !== v.iban) upd++; cur[id] = Object.assign({}, cur[id], v); }
        await writeJson('master.json', cur); S.master = cur; toast(t('mImported', Object.keys(m).length, add, upd)); }
      catch (e) { toast(t('readFail', f.name, e.message)); }
    }
    recompute(); render();
  }
  // 마스터 변경은 항상 최신 파일을 다시 읽어 그 위에 적용(여러 사람이 동시에 써도 덮어쓰지 않도록)
  async function masterApply(fn) { const cur = await readJson('master.json', {}); fn(cur); await writeJson('master.json', cur); S.master = cur; }

  function recompute() {
    const [y, m] = S.month.split('-').map(Number);
    const act = S.inputs.filter((i) => i.output !== 'ignore');
    S.res = act.length ? W.build(act, { year: y, month: m, master: S.master, special: S.special, overrides: S.overrides }) : null;
  }

  // ───────── 만들기 ─────────
  async function createFiles() {
    const r = S.res, files = []; const dir = S.month;
    try {
      for (const o of r.outputs) { const f = W.fill(S.templates[o.key].bytes, o, S.valueDate); files.push({ name: o.fileName, data: f.bytes, key: o.key, count: f.check.count, total: f.check.total }); }
      for (const i of S.inputs) await S.store.write(dir + '/HR/' + i.fileName, i.data);
      for (const f of files) await S.store.write(dir + '/Accounting/' + f.name, f.data);
      // 마스터 반영: 체크된 신규 인원, HR 계좌로 바꾸기로 한 사람
      const adds = [], upds = [];
      for (const o of r.outputs) for (const row of o.rows) { if (!row.inMaster && S.addNew[row.id] !== false) adds.push(row); else if (row.inMaster && (S.overrides[row.id] || {}).useHr) upds.push(row); }
      if (adds.length || upds.length) await masterApply((cur) => { for (const x of adds) cur[x.id] = { name: x.name, bank: x.bank, iban: x.iban, added: S.month }; for (const x of upds) cur[x.id] = Object.assign({}, cur[x.id], { bank: x.bank, iban: x.iban, updated: S.month }); });
      const log = { version: VERSION, month: S.month, valueDate: S.valueDate, createdAt: new Date().toISOString(), grandTotal: W.r2(r.outputs.reduce((a, o) => a + o.total, 0)),
        inputs: S.inputs.map((i) => ({ fileName: i.fileName, output: i.output, count: i.employees.length })),
        outputs: r.outputs.map((o) => ({ key: o.key, label: o.label, fileName: o.fileName, count: o.count, total: o.total, people: o.rows.map((x) => ({ id: x.id, name: x.name, salary: x.salary })) })),
        notices: r.notices.map((n) => ({ level: n.level, type: n.type, output: n.output, id: n.id, name: n.name, msg: n.msg })) };
      await writeJson(dir + '/run.json', log);
      S.created = { files, added: adds.length, dir }; S.overrides = {}; S.addNew = {}; recompute(); render();
    } catch (e) { toast(e.message); console.error(e); }
  }

  // ───────── 화면 ─────────
  function render() {
    document.documentElement.lang = lang;
    $('#storeInfo').innerHTML = S.store.kind === 'folder' ? t('folderOn', esc(S.store.name)) : t('folderNone');
    $('#btnFolder').textContent = S.pendingHandle ? t('reconnect') : S.store.kind === 'folder' ? t('change') : t('connect');
    $('#btnFolder').classList.toggle('hide', !window.showDirectoryPicker);
    $('#nav').innerHTML = [['run', 'tabRun'], ['master', 'tabMaster'], ['tpl', 'tabTpl'], ['hist', 'tabHist'], ['help', 'tabHelp']].map(([k, l]) => `<button data-tab="${k}" class="${S.tab === k ? 'on' : ''}">${t(l)}</button>`).join('');
    const m = $('#main');
    if (S.tab === 'run') m.innerHTML = viewRun(); else if (S.tab === 'master') m.innerHTML = viewMaster(); else if (S.tab === 'tpl') m.innerHTML = viewTpl(); else if (S.tab === 'hist') { m.innerHTML = `<section class="card"><h2>${t('hTitle')}</h2><div id="hist">…</div></section>`; viewHist(); } else m.innerHTML = viewHelp();
  }
  const outSelect = (i, idx) => `<select data-act="setOut" data-i="${idx}"><option value="">${t('choose')}</option>${W.OUTPUTS.map((o) => `<option value="${o.key}" ${i.output === o.key ? 'selected' : ''}>${esc(o.label)}</option>`).join('')}<option value="ignore" ${i.output === 'ignore' ? 'selected' : ''}>${t('ignore')}</option></select>`;
  const outLabel = (k) => (W.OUTPUTS.find((o) => o.key === k) || { label: k || '' }).label;

  function viewRun() {
    let h = '';
    if (!window.showDirectoryPicker) h += `<div class="box warn">${t('unsupported')}</div>`;
    h += `<section class="card"><h2><span class="step">1</span>${t('s1')}</h2><div class="row">
      <label>${t('month')}<input type="month" data-act="month" value="${S.month}"></label>
      <label>${t('runDate')}<input type="date" data-act="runDate" value="${S.runDate}"></label>
      <label>${t('valueDate')}<input type="text" data-act="valueDate" value="${esc(S.valueDate)}" maxlength="8" size="10" style="font-family:Consolas,monospace"></label>
      </div><p class="hint">${t('vdHint')}</p></section>`;
    h += `<section class="card"><h2><span class="step">2</span>${t('s2')}</h2><div class="drop" data-drop="hr">${t('dropHr')}<input type="file" multiple accept=".xls,.xlsx,.zip" class="hide" data-file="hr"></div>`;
    if (S.inputs.length) { h += `<table style="margin-top:12px"><tr><th>${t('file')}</th><th>${t('project')}</th><th class="n">${t('people')}</th><th class="n">${t('sum')}</th><th></th></tr>` + S.inputs.map((i, idx) => `<tr><td>${esc(i.fileName)}</td><td>${outSelect(i, idx)}</td><td class="n">${i.employees.length}</td><td class="n">${money(i.employees.reduce((a, e) => a + e.salary, 0))}</td><td><button class="link" data-act="rmFile" data-i="${idx}">${t('remove')}</button></td></tr>`).join('') + `</table><p style="margin:10px 0 0"><button class="link" data-act="clearFiles">${t('clear')}</button></p>`; }
    h += `</section>`;
    if (!S.res) return h + createdBox();
    // 점검
    const N = S.res.notices, news = N.filter((n) => n.type === 'new'), diffs = N.filter((n) => n.type === 'iban-diff' || n.type === 'bank-diff'), errs = N.filter((n) => n.level === 'error' && !diffs.includes(n)), warns = N.filter((n) => (n.level === 'warn' || n.level === 'info') && !diffs.includes(n));
    h += `<section class="card"><h2><span class="step">3</span>${t('s3')}</h2>`;
    if (!N.length) h += `<div class="box ok">${t('noIssues')}</div>`;
    if (errs.length) h += `<div class="box err"><h3>${t('errTitle', errs.length)}</h3><ul>${errs.map((n) => `<li><b>${esc(outLabel(n.output))}</b> ${n.id ? esc(n.id) + ' ' + esc(n.name) + ' — ' : ''}${noticeText(n)}</li>`).join('')}</ul></div>`;
    if (diffs.length) { const ids = [...new Set(diffs.map((n) => n.id))]; h += `<div class="box err"><h3>${t('diffTitle', ids.length)}</h3><table>${ids.map((id) => { const ns = diffs.filter((n) => n.id === id), use = (S.overrides[id] || {}).useHr; return `<tr><td class="mono">${esc(id)}</td><td>${esc(ns[0].name)}<div class="small">${ns.map((n) => noticeText(n)).join('<br>')}</div></td><td style="white-space:nowrap"><label><input type="radio" name="d${esc(id)}" data-act="useHr" data-id="${esc(id)}" value="0" ${use ? '' : 'checked'}> ${t('useMaster')}</label><br><label><input type="radio" name="d${esc(id)}" data-act="useHr" data-id="${esc(id)}" value="1" ${use ? 'checked' : ''}> ${t('useHr')}</label></td></tr>`; }).join('')}</table></div>`; }
    if (news.length) h += `<div class="box new"><h3>${t('newTitle', news.length)}</h3><p class="hint">${t('newHint')}</p><table><tr><th>${t('out')}</th><th>${t('id')}</th><th>${t('hrName')}</th><th>${t('name')}</th><th>${t('bank')}</th><th>${t('iban')}</th><th>${t('addMaster')}</th></tr>${news.map((n) => `<tr><td>${esc(outLabel(n.output))}</td><td class="mono">${esc(n.id)}</td><td class="small">${esc(n.hrName)}</td><td><input class="cell" data-act="ov" data-id="${esc(n.id)}" data-k="name" maxlength="35" value="${esc(n.row.name)}"></td><td style="width:90px"><input class="cell" data-act="ov" data-id="${esc(n.id)}" data-k="bank" maxlength="35" value="${esc(n.row.bank)}"></td><td style="width:230px"><input class="cell mono" data-act="ov" data-id="${esc(n.id)}" data-k="iban" maxlength="34" value="${esc(n.row.iban)}"></td><td style="text-align:center"><input type="checkbox" data-act="addNew" data-id="${esc(n.id)}" ${S.addNew[n.id] === false ? '' : 'checked'}></td></tr>`).join('')}</table></div>`;
    if (warns.length) h += `<div class="box warn"><h3>${t('warnTitle', warns.length)}</h3><ul>${warns.map((n) => `<li><b>${esc(outLabel(n.output))}</b> ${n.id ? esc(n.id) + ' ' + esc(n.name) + ' — ' : ''}${noticeText(n)}</li>`).join('')}</ul></div>`;
    h += `</section>`;
    // 결과
    const missing = S.res.outputs.filter((o) => !S.templates[o.key] || !S.templates[o.key].bytes), grand = S.res.outputs.reduce((a, o) => a + o.total, 0);
    h += `<section class="card"><h2><span class="step">4</span>${t('s4')}</h2><table><tr><th>${t('out')}</th><th class="n">${t('people')}</th><th class="n">${t('sum')}</th><th>${t('tpl')}</th><th>${t('vsPrev')}</th></tr>`;
    for (const o of S.res.outputs) {
      const tp = S.templates[o.key]; const tag = !tp || !tp.bytes ? `<span class="tag err">${t('tplMissing')}</span>` : tp.std ? `<span class="tag ok">${t('tplOk')}</span>` : `<span class="tag warn">${t('tplOdd')}</span>`;
      let cmp = `<span class="muted small">${t('noPrev')}</span>`; const p = S.prev && S.prev.outputs.find((x) => x.key === o.key);
      if (p) { const pid = new Set(p.people.map((x) => x.id)), cid = new Set(o.rows.map((x) => x.id)); const joined = o.rows.filter((x) => !pid.has(x.id)), left = p.people.filter((x) => !cid.has(x.id)); const dc = o.count - p.count, dt = o.total - p.total;
        cmp = `${dc >= 0 ? '+' : ''}${dc} · ${dt >= 0 ? '+' : ''}${money(dt)}` + (joined.length || left.length ? ` <details style="display:inline-block;vertical-align:top"><summary>${t('joined')} ${joined.length} / ${t('left')} ${left.length}</summary><div class="small">${joined.map((x) => '+ ' + esc(x.name)).concat(left.map((x) => '− ' + esc(x.name))).join('<br>')}</div></details>` : ''); }
      h += `<tr><td><b>${esc(o.label)}</b><div class="small muted">${esc(o.fileName)}</div></td><td class="n">${o.count}</td><td class="n">${money(o.total)}</td><td>${tag}</td><td>${cmp}</td></tr>`;
    }
    h += `<tr><td><b>${t('grand')}</b></td><td class="n"><b>${S.res.outputs.reduce((a, o) => a + o.count, 0)}</b></td><td class="n total">${money(grand)}</td><td></td><td></td></tr></table>`;
    const hasErr = errs.length || diffs.some((n) => n.type === 'iban-diff'), vdOk = /^\d{8}$/.test(S.valueDate);
    if (missing.length) h += `<div class="box err">${t('needTpl', missing.map((o) => esc(o.label)).join(', '))} <button class="link" data-tab="tpl">${t('tabTpl')}</button></div>`;
    if (!vdOk) h += `<div class="box err">${t('needVd')}</div>`;
    if (hasErr) h += `<p><label><input type="checkbox" data-act="ack" ${S.ack ? 'checked' : ''}> ${t('ack')}</label></p>`;
    h += `<p style="margin-top:14px"><button class="btn primary" data-act="create" ${missing.length || !vdOk || (hasErr && !S.ack) ? 'disabled' : ''}>${t('create')}</button></p></section>`;
    return h + createdBox();
  }
  function createdBox() {
    if (!S.created) return ''; const c = S.created;
    return `<section class="card"><div class="box ok"><h3>${t('done', esc((S.store.kind === 'folder' ? S.store.name + '/' : '') + c.dir + '/Accounting'))}</h3>${t('selfOk')}${c.added ? '<br>' + t('masterAdded', c.added) : ''}</div><table><tr><th>${t('file')}</th><th class="n">${t('people')}</th><th class="n">${t('sum')}</th><th></th></tr>${c.files.map((f, i) => `<tr><td>${esc(f.name)}</td><td class="n">${f.count}</td><td class="n">${money(f.total)}</td><td><button class="link" data-act="dl" data-i="${i}">${t('dl')}</button></td></tr>`).join('')}</table><p style="margin-top:12px"><button class="btn" data-act="dlAll">${t('dlAll')}</button></p></section>`;
  }
  function viewMaster() {
    const ids = Object.keys(S.master), q = S.mq.trim().toLowerCase(); const rows = ids.filter((id) => !q || id.includes(q) || (S.master[id].name || '').toLowerCase().includes(q) || (S.master[id].iban || '').toLowerCase().includes(q)).sort((a, b) => (S.master[a].name || '').localeCompare(S.master[b].name || ''));
    let h = `<section class="card"><h2>${t('mTitle')} <span class="muted" style="font-weight:400">· ${t('mCount', ids.length)}</span></h2><div class="row" style="margin-bottom:12px"><input type="search" data-act="mq" placeholder="${t('mSearch')}" value="${esc(S.mq)}" style="min-width:260px"><button class="btn" data-pick="master">${t('mImport')}</button><input type="file" accept=".xls,.xlsx" class="hide" data-file="master"><button class="btn" data-act="mExport">${t('mExport')}</button><button class="btn" data-act="mAdd">${t('mAdd')}</button></div>`;
    if (!ids.length && S.mEdit !== '') return specialCard() + h + `<div class="box new">${t('mEmpty')}</div></section>`;
    const editRow = (id, v) => `<tr><td><input class="cell mono" id="me_id" value="${esc(id)}" maxlength="10" ${id ? 'readonly' : ''} style="width:110px"></td><td><input class="cell" id="me_name" maxlength="35" value="${esc(v.name || '')}"></td><td><input class="cell" id="me_bank" maxlength="35" value="${esc(v.bank || '')}" style="width:90px"></td><td><input class="cell mono" id="me_iban" maxlength="34" value="${esc(v.iban || '')}"></td><td style="white-space:nowrap"><button class="link" data-act="mSave">${t('save')}</button> · <button class="link" data-act="mCancel">${t('cancel')}</button></td></tr>`;
    h = specialCard() + h;
    h += `<div class="scroll"><table><tr><th>${t('id')}</th><th>${t('name')}</th><th>${t('bank')}</th><th>${t('iban')}</th><th></th></tr>${S.mEdit === '' ? editRow('', {}) : ''}${rows.map((id) => S.mEdit === id ? editRow(id, S.master[id]) : `<tr><td class="mono">${esc(id)}</td><td>${esc(S.master[id].name)}</td><td>${esc(S.master[id].bank)}</td><td class="mono">${esc(S.master[id].iban)}</td><td style="white-space:nowrap"><button class="link" data-act="mEdit" data-id="${esc(id)}">${t('edit')}</button> · <button class="link" data-act="mDel" data-id="${esc(id)}">${t('del')}</button></td></tr>`).join('')}</table></div></section>`;
    return h;
  }

  function specialCard() {
    const rows = Object.entries(S.special);
    const opts = W.OUTPUTS.map((o) => `<option value="${o.key}">${esc(o.label)}</option>`).join('');
    return `<section class="card"><h2>Always in a separate file</h2><p class="hint">People listed here always go to the chosen output file, whichever HR file they come from (for example the D360 bank person and the PDH trainee). Saved in the shared folder, not in the app code.</p>
      <table><tr><th>${t('id')}</th><th>${t('name')}</th><th>${t('out')}</th><th></th></tr>${rows.map(([id, k]) => `<tr><td class="mono">${esc(id)}</td><td>${esc((S.master[id] || {}).name || '')}</td><td>${esc(outLabel(k))}</td><td><button class="link" data-act="spDel" data-id="${esc(id)}">${t('del')}</button></td></tr>`).join('')}
      <tr><td><input class="cell mono" id="sp_id" maxlength="10" placeholder="10-digit ID" style="width:120px"></td><td></td><td><select id="sp_out">${opts}</select></td><td><button class="link" data-act="spAdd">Add</button></td></tr></table></section>`;
  }
  async function saveSpecial(fn) { const cur = await readJson('settings.json', {}); cur.special = cur.special || {}; fn(cur.special); await writeJson('settings.json', cur); S.special = cur.special; recompute(); render(); }
  function viewTpl() {
    let h = `<section class="card"><h2>${t('tTitle')}</h2><p class="hint">${t('tHint')}</p><div class="drop" data-drop="tpl">${t('dropTpl')}<input type="file" multiple accept=".xls,.zip" class="hide" data-file="tpl"></div><table style="margin-top:14px"><tr><th>${t('out')}</th><th>${t('tpl')}</th><th class="n">${t('tPeople')}</th><th></th></tr>`;
    for (const o of W.OUTPUTS) { const tp = S.templates[o.key]; h += `<tr><td><b>${esc(o.label)}</b></td><td>${!tp ? `<span class="tag err">${t('tplMissing')}</span>` : tp.error ? `<span class="tag err">${esc(tp.error)}</span>` : tp.std ? `<span class="tag ok">${t('tStd')}</span>` : `<span class="tag warn">${t('tplOdd')}</span><div class="small" style="max-width:520px;margin-top:4px">${t('tOddMsg')}</div>`}</td><td class="n">${tp && tp.bytes ? tp.rows : ''}</td><td><button class="link" data-pick="tpl1" data-key="${o.key}">${t('tUpload')}</button></td></tr>`; }
    return h + `</table><input type="file" accept=".xls" class="hide" data-file="tpl1"></section>`;
  }
  async function viewHist() {
    const months = (await S.store.list('')).filter((n) => /^\d{4}-\d{2}$/.test(n)).sort().reverse(); const el = $('#hist'); if (!el) return;
    if (!months.length) { el.innerHTML = `<p class="muted">${t('hEmpty')}</p>`; return; }
    let h = '';
    for (const mth of months) { const log = await readJson(mth + '/run.json', null); if (!log) continue;
      h += `<details ${mth === months[0] ? 'open' : ''} style="margin-bottom:10px"><summary><b>${mth}</b> · ${money(log.grandTotal)} SAR · ${log.outputs.reduce((a, o) => a + o.count, 0)} ${t('people')} · Value date ${esc(log.valueDate)} · ${t('created')} ${esc(log.createdAt.slice(0, 16).replace('T', ' '))}</summary><table style="margin:8px 0 0"><tr><th>${t('out')}</th><th class="n">${t('people')}</th><th class="n">${t('sum')}</th><th>${t('file')}</th></tr>${log.outputs.map((o) => `<tr><td>${esc(o.label)}</td><td class="n">${o.count}</td><td class="n">${money(o.total)}</td><td><button class="link" data-act="hdl" data-p="${esc(mth + '/Accounting/' + o.fileName)}">${esc(o.fileName)}</button></td></tr>`).join('')}</table></details>`; }
    el.innerHTML = h || `<p class="muted">${t('hEmpty')}</p>`;
  }
  function viewHelp() {
    return `<section class="card"><h2>How it works</h2>
      <p><b>Rule:</b> amounts (salary, basic, housing, other, deduction) come from the HR file; name, bank code and IBAN come from the master list. People who are not in the master list are written from HR values and flagged.</p>
      <p><b>Template is preserved:</b> the macro, the Export button and the formatting of the bank template (.xls) stay byte-for-byte the same; only cell values change. The formulas in the salary column and in Total Salary are kept.</p>
      <p><b>Self-check:</b> every created file is read back and compared (people count, total, value date, IDs, IBANs). If anything differs, nothing is saved.</p>
      <p><b>Storage layout:</b> in the shared folder: <code>templates/</code>, <code>master.json</code>, and per month <code>2026-11/HR</code>, <code>2026-11/Accounting</code>, <code>2026-11/run.json</code>.</p>
      <p><b>Privacy:</b> payroll data stays in this browser and in the shared folder you choose. Nothing is sent to a server.</p><p class="muted small">v${VERSION}</p></section>`;
  }

  // ───────── 이벤트 ─────────
  document.addEventListener('click', async (ev) => {
    const el = ev.target.closest('[data-tab],[data-act],[data-pick],.drop'); if (!el) return;
    if (el.dataset.tab) { S.tab = el.dataset.tab; S.mEdit = null; render(); return; }
    if (el.classList.contains('drop')) { if (ev.target.tagName !== 'INPUT') $('input', el).click(); return; }
    if (el.dataset.pick) { const inp = $(`input[data-file="${el.dataset.pick}"]`); inp.dataset.key = el.dataset.key || ''; inp.click(); return; }
    const a = el.dataset.act;
    if (a === 'rmFile') { S.inputs.splice(+el.dataset.i, 1); S.created = null; recompute(); render(); }
    else if (a === 'clearFiles') { S.inputs = []; S.overrides = {}; S.addNew = {}; S.ack = false; S.created = null; recompute(); render(); }
    else if (a === 'create') { el.disabled = true; el.textContent = t('working'); await createFiles(); }
    else if (a === 'dl') { const f = S.created.files[+el.dataset.i]; download(f.name, f.data, 'application/vnd.ms-excel'); }
    else if (a === 'dlAll') download(`WPS_${S.created.dir}.zip`, X.zipStore(S.created.files.map((f) => ({ name: f.name, data: f.data }))), 'application/zip');
    else if (a === 'hdl') { const u = await S.store.read(el.dataset.p); if (u) download(el.dataset.p.split('/').pop(), u, 'application/vnd.ms-excel'); }
    else if (a === 'mExport') { const lines = ['ID,Name,Bank,IBAN'].concat(Object.keys(S.master).map((id) => [id, S.master[id].name, S.master[id].bank, S.master[id].iban].map((v) => '"' + String(v || '').replace(/"/g, '""') + '"').join(','))); download('master.csv', te.encode('﻿' + lines.join('\r\n')), 'text/csv'); }
    else if (a === 'spAdd') { const id = $('#sp_id').value.trim(), k = $('#sp_out').value; if (!/^\d{10}$/.test(id)) return toast(t('idBad')); await saveSpecial((sp) => { sp[id] = k; }); }
    else if (a === 'spDel') { const id = el.dataset.id; await saveSpecial((sp) => { delete sp[id]; }); }
    else if (a === 'mAdd') { S.mEdit = ''; render(); }
    else if (a === 'mEdit') { S.mEdit = el.dataset.id; render(); }
    else if (a === 'mCancel') { S.mEdit = null; render(); }
    else if (a === 'mDel') { const id = el.dataset.id; if (confirm(t('confirmDel', S.master[id].name))) { await masterApply((cur) => { delete cur[id]; }); recompute(); render(); } }
    else if (a === 'mSave') { const id = $('#me_id').value.trim(), name = W.clean($('#me_name').value), bank = W.clean($('#me_bank').value), iban = W.cleanIban($('#me_iban').value);
      if (!/^\d{10}$/.test(id)) return toast(t('idBad')); if (!/^SA[0-9A-Z]{22}$/.test(iban)) return toast(t('ibanBad'));
      await masterApply((cur) => { cur[id] = Object.assign({}, cur[id], { name, bank, iban }); }); S.mEdit = null; recompute(); render(); }
  });
  document.addEventListener('change', async (ev) => {
    const el = ev.target, a = el.dataset.act;
    if (el.dataset.file) { const files = [...el.files]; const kind = el.dataset.file, key = el.dataset.key; el.value = ''; if (!files.length) return; if (kind === 'hr') await addHr(files); else if (kind === 'tpl') await addTemplates(files); else if (kind === 'tpl1') await addTemplates(files, key); else if (kind === 'master') await importMaster(files); return; }
    if (a === 'setOut') { S.inputs[+el.dataset.i].output = el.value || null; S.created = null; recompute(); render(); }
    else if (a === 'month') { if (!/^\d{4}-\d{2}$/.test(el.value)) return; S.month = el.value; await loadPrev(); recompute(); render(); }
    else if (a === 'runDate') { S.runDate = el.value; const [y, m, d] = el.value.split('-').map(Number); if (y) { S.valueDate = W.valueDateFor(new Date(y, m - 1, d)); } render(); }
    else if (a === 'valueDate') { S.valueDate = el.value.trim(); render(); }
    else if (a === 'ov') { const id = el.dataset.id; S.overrides[id] = Object.assign({}, S.overrides[id], { [el.dataset.k]: el.value }); recompute(); render(); }
    else if (a === 'useHr') { const id = el.dataset.id; S.overrides[id] = Object.assign({}, S.overrides[id], { useHr: el.value === '1' }); recompute(); render(); }
    else if (a === 'addNew') { S.addNew[el.dataset.id] = el.checked; }
    else if (a === 'ack') { S.ack = el.checked; render(); }
  });
  document.addEventListener('input', (ev) => { if (ev.target.dataset.act === 'mq') { S.mq = ev.target.value; const pos = ev.target.selectionStart; render(); const n = $('input[data-act="mq"]'); n.focus(); n.setSelectionRange(pos, pos); } });
  ['dragover', 'dragleave', 'drop'].forEach((type) => document.addEventListener(type, async (ev) => {
    const z = ev.target.closest && ev.target.closest('.drop'); if (type !== 'dragleave') ev.preventDefault(); if (!z) return;
    z.classList.toggle('over', type === 'dragover');
    if (type === 'drop') { const files = [...ev.dataTransfer.files]; if (z.dataset.drop === 'hr') await addHr(files); else await addTemplates(files); }
  }));
  $('#btnFolder').addEventListener('click', () => connectFolder(S.pendingHandle));

  (async function init() {
    try { const h = window.showDirectoryPicker ? await idb.get('dirHandle') : null;
      if (h) { if ((await h.queryPermission({ mode: 'readwrite' })) === 'granted') S.store = folderStore(h); else S.pendingHandle = h; } } catch (e) { /* 저장된 폴더 없음 */ }
    await loadShared(); recompute(); render();
    if (S.pendingHandle) toast(t('perm'));
  })();
})();
