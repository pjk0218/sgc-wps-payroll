/* app.js — 화면과 저장. 급여 데이터는 이 브라우저와 선택한 공유 폴더 밖으로 나가지 않는다(서버 전송 없음). */
(function () {
  'use strict';
  const X = window.XlsBin, W = window.Wps, VERSION = '1.2.2';
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
    async list(dir) { const pre = 'f:' + (dir ? dir + '/' : ''); const set = new Set(); for (const k of await idb.keys()) if (typeof k === 'string' && k.startsWith(pre)) set.add(k.slice(pre.length).split('/')[0]); return [...set].sort(); },
    async remove(p) { const pre = 'f:' + p; for (const k of await idb.keys()) if (typeof k === 'string' && (k === pre || k.startsWith(pre + '/'))) await idb.del(k); } };
  function folderStore(root) {
    const dirOf = async (parts, create) => { let d = root; for (const p of parts) d = await d.getDirectoryHandle(p, { create }); return d; };
    return { kind: 'folder', name: root.name,
      async read(p) { try { const parts = p.split('/'), f = parts.pop(); const d = await dirOf(parts, false); const h = await d.getFileHandle(f); return new Uint8Array(await (await h.getFile()).arrayBuffer()); } catch (e) { if (e.name === 'NotFoundError' || e.name === 'TypeMismatchError') return null; throw e; } },
      async write(p, u8) { const parts = p.split('/'), f = parts.pop(); const d = await dirOf(parts, true); const h = await d.getFileHandle(f, { create: true }); const w = await h.createWritable(); await w.write(u8); await w.close(); },
      async list(dir) { try { const d = await dirOf(dir ? dir.split('/') : [], false); const out = []; for await (const [n] of d.entries()) out.push(n); return out.sort(); } catch (e) { return []; } },
      async remove(p) { const parts = p.split('/'), f = parts.pop(); try { const d = await dirOf(parts, false); await d.removeEntry(f, { recursive: true }); } catch (e) { if (e.name !== 'NotFoundError') throw e; } } };
  }
  const te = new TextEncoder(), td = new TextDecoder();
  const readJson = async (p, dflt) => { const u = await S.store.read(p); if (!u) return dflt; try { return JSON.parse(td.decode(u)); } catch (e) { return dflt; } };
  const writeJson = (p, o) => S.store.write(p, te.encode(JSON.stringify(o, null, 1)));

  // ───────── 상태 ─────────
  const now = new Date();
  const S = { tplMsg: [], live: null, lastChange: null, viewingHist: false, histSel: {}, pendingRender: false, hrCache: {}, connectMsg: '', special: {}, exportPath: '', store: localStore, pendingHandle: null, tab: 'run', master: {}, templates: {}, inputs: [], overrides: {}, addNew: {}, res: null, ack: false, prev: null, created: null,
    month: now.getFullYear() + '-' + pad2(now.getMonth() + 1), runDate: now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate()), valueDate: W.valueDateFor(now), vdTouched: false, mq: '', mEdit: null };
  window.__wps = S;
  const CLIENT = Math.random().toString(36).slice(2, 10);
  let userName = localStorage.getItem('wps.user') || '';

  function toast(msg) { const el = $('#toast'); el.textContent = msg; el.classList.add('on'); clearTimeout(toast.t); toast.t = setTimeout(() => el.classList.remove('on'), 3500); }
  function download(name, u8, type) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([u8], { type: type || 'application/octet-stream' })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }

  async function loadShared() {
    S.master = await readJson('master.json', {});
    const st = await readJson('settings.json', {}); S.special = st.special || {}; S.exportPath = st.exportPath || '';
    S.templates = {};
    for (const o of W.OUTPUTS.filter((x) => !x.fs)) { const u = await S.store.read('templates/' + o.key + '.xls'); if (u) { try { S.templates[o.key] = tplInfo(u); } catch (e) { S.templates[o.key] = { error: e.message }; } } }
    await loadPrev();
  }
  function tplInfo(u8) { const sh = X.readXls(u8, W.SHEET), lay = W.layout(sh); return { bytes: u8, rows: W.readEmployees(sh, lay).employees.length, std: lay.macroStandard }; }
  const prevMonthOf = (ym) => { const [y, m] = ym.split('-').map(Number); return m === 1 ? (y - 1) + '-12' : y + '-' + pad2(m - 1); };
  async function loadPrev() { S.prev = await readJson(prevMonthOf(S.month) + '/run.json', null); }

  // 연결 실패 이유를 화면에 남긴다 (조용히 끝나지 않게)
  function connectFail(msg) { S.connectMsg = msg; S.tab = 'run'; render(); }
  async function connectFolder(handle) {
    S.connectMsg = '';
    let step = 'pick';
    try {
      if (!handle) handle = await window.showDirectoryPicker({ mode: 'readwrite' });
      step = 'permission';
      const perm = await handle.requestPermission({ mode: 'readwrite' });
      if (perm !== 'granted') return connectFail(`The browser did not give permission to edit the folder "${handle.name}" (answer: ${perm}). Click Connect shared folder again and choose "Edit files" / "Allow". If no question appears, click the icon left of the address bar → Site settings → File editing → set to "Ask" or "Allow", then reload the page.`);
      step = 'test';
      // 실제로 쓰고 지울 수 있는지 확인
      const fh = await handle.getFileHandle('.wps-write-test', { create: true }); const w = await fh.createWritable(); await w.write('ok'); await w.close(); await handle.removeEntry('.wps-write-test');
      S.store = folderStore(handle); S.pendingHandle = null; await idb.set('dirHandle', handle); await loadShared(); S.hrCache = {};
      const live = await readJson('live.json', null); S.live = live; await openMonth(live && live.month ? live.month : S.month);
      toast('Connected: ' + handle.name);
    } catch (e) {
      if (e.name === 'AbortError' && step === 'pick') return;
      const why = e.name === 'SecurityError' || e.name === 'NotAllowedError' ? 'The browser or a company policy blocked access to this folder.' : e.name === 'NoModificationAllowedError' || e.name === 'InvalidModificationError' ? 'The folder is read-only for this Windows user (check the share permission is Read/Write).' : 'The folder could not be opened.';
      connectFail(`${why} Step: ${step}. Details: ${e.name} — ${e.message}. Try choosing the folder through Network → JKPC → WPS Payroll instead of the drive letter.`);
    }
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
    const added = [];
    for (const f of await expand(fileList)) {
      try { const sh = await X.readAny(f.data, W.SHEET); const { employees, header } = W.readEmployees(sh);
        await S.store.write(S.month + '/HR/' + f.name, f.data);
        S.hrCache[S.month + '/' + f.name + '/' + f.data.length] = { fileName: f.name, data: f.data, employees, header };
        added.push({ fileName: f.name, output: W.detectOutput(f.name), size: f.data.length }); }
      catch (e) { toast(t('readFail', f.name, e.message)); }
    }
    if (added.length) await editDraft((d) => { for (const a of added) { d.inputs = d.inputs.filter((i) => i.fileName !== a.fileName); d.inputs.push(a); } d.created = null; });
  }
  // 은행 양식 검사: 회계팀 최종본은 상단 인원수·합계가 행과 맞고, 매크로가 읽는 표준 위치를 쓴다.
  // HR 원본은 상단 값이 맞지 않거나(예: 186명, 2024년 Value Date) 열 위치가 달라서 걸러진다.
  function templateProblem(u8) {
    const sh = X.readXls(u8, W.SHEET), lay = W.layout(sh), r = W.readEmployees(sh, lay), n = r.employees.length;
    const sum = W.r2(r.employees.reduce((a, e) => a + e.salary, 0)), hc = Number(r.header.count), ht = Number(r.header.total);
    if (!lay.macroStandard) return 'its columns or header cells are not where the Export macro reads them (salary must be in column E, value date in F16)';
    if (hc !== n || Math.abs(ht - sum) > 0.01) return `its header says ${r.header.count} employees / total ${r.header.total}, but the sheet has ${n} rows / total ${sum} — this looks like an HR file, not the accounting team's final bank file`;
    return '';
  }
  async function addTemplates(fileList, forceKey) {
    let n = 0; const bad = [];
    for (const f of await expand(fileList)) {
      const key = forceKey || W.detectTemplate(f.name); if (!key) { bad.push(`${f.name}: could not tell which output this file is for`); continue; }
      try { const why = templateProblem(f.data); if (why) { bad.push(`${f.name} was NOT registered: ${why}.`); continue; }
        const info = tplInfo(f.data); await S.store.write('templates/' + key + '.xls', f.data); S.templates[key] = info; n++; } catch (e) { bad.push(`${f.name}: ${e.message}`); }
    }
    S.tplMsg = bad; if (n) toast(t('tplSaved', n)); recompute(); render();
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

  // ───────── 공유 작업 상태: <월>/draft.json + live.json ─────────
  // 한 사람이 바꾸면 공유 폴더에 바로 저장하고, 다른 사람 화면은 3초마다 확인해서 따라간다.
  const emptyDraft = (month) => ({ month, runDate: S.runDate, valueDate: S.valueDate, inputs: [], overrides: {}, addNew: {}, ack: false, created: null, rev: 0 });
  async function readDraft(month) {
    const d = await readJson(month + '/draft.json', null); if (d) return d;
    const log = await readJson(month + '/run.json', null); if (!log) return null;   // 예전 방식으로 저장된 달
    return { month, runDate: S.runDate, valueDate: log.valueDate, inputs: (log.inputs || []).map((i) => ({ fileName: i.fileName, output: i.output })), overrides: {}, addNew: {}, ack: false, rev: 0,
      created: { at: log.createdAt, by: log.createdBy || '', files: log.outputs.map((o) => ({ name: o.fileName, count: o.count, total: o.total })), added: 0 } };
  }
  async function hrInput(month, e) {
    const key = month + '/' + e.fileName + '/' + (e.size || '');
    if (!S.hrCache[key]) { const u = await S.store.read(month + '/HR/' + e.fileName); if (!u) return null; const { employees, header } = W.readEmployees(await X.readAny(u, W.SHEET)); S.hrCache[key] = { fileName: e.fileName, data: u, employees, header }; }
    return Object.assign({}, S.hrCache[key], { output: e.output });
  }
  async function applyDraft(d) {
    // HR 파일을 먼저 다 읽은 뒤 화면 상태를 한 번에 바꾼다 (반쯤 바뀐 화면이 보이지 않게)
    const ins = []; for (const e of d.inputs || []) { try { const i = await hrInput(d.month, e); if (i) ins.push(i); } catch (err) { toast(t('readFail', e.fileName, err.message)); } }
    S.month = d.month; S.draftRev = d.rev || 0;
    if (d.runDate) S.runDate = d.runDate; if (d.valueDate) S.valueDate = d.valueDate;
    S.overrides = d.overrides || {}; S.addNew = d.addNew || {}; S.ack = !!d.ack; S.created = d.created || null;
    S.inputs = ins; S.lastChange = d.by ? { by: d.by, at: d.at } : null;
    await loadPrev(); recompute(); safeRender();
  }
  // 저장·불러오기는 한 번에 하나씩 (빠르게 연달아 바꿔도 서로 덮어쓰지 않게)
  let chain = Promise.resolve();
  const withTimeout = (pr, ms, what) => Promise.race([pr, new Promise((_, rej) => setTimeout(() => rej(new Error(what + ' did not answer within ' + (ms / 1000) + ' s')), ms))]);
  // 한 단계가 멈춰도(네트워크 끊김) 다음 작업이 영원히 막히지 않게 90초 후 체인을 풀어 줌
  const serial = (fn) => { const r = chain.then(fn, fn); chain = withTimeout(r, 90000, 'A shared-folder step').catch(() => {}); return r; };
  const openMonth = (month) => serial(() => openMonthNow(month));
  async function openMonthNow(month) { S.hrCache = {}; const d = await readDraft(month); await applyDraft(d || emptyDraft(month)); }
  // 모든 변경은 '최신 draft를 다시 읽고 → 바꾸고 → 저장' (두 사람이 동시에 바꿔도 서로 지우지 않게)
  const editDraft = (fn) => serial(() => editDraftNow(fn));
  async function editDraftNow(fn) {
    const d = (await readJson(S.month + '/draft.json', null)) || (await readDraft(S.month)) || emptyDraft(S.month);
    d.month = S.month; fn(d); d.rev = (d.rev || 0) + 1; d.by = userName || 'Someone'; d.client = CLIENT; d.at = new Date().toISOString();
    await writeJson(S.month + '/draft.json', d);
    const live = { month: S.month, rev: Date.now() + '-' + CLIENT, by: d.by, client: CLIENT, at: d.at };
    await writeJson('live.json', live); S.live = live; S.viewingHist = false;
    await applyDraft(d);
  }
  const poll = () => serial(pollNow);
  async function pollNow() {
    try {
      const live = await readJson('live.json', null);
      if (!live || (S.live && live.rev === S.live.rev)) return;
      S.live = live;
      if (live.client === CLIENT) return;
      S.master = await readJson('master.json', S.master);
      const st = await readJson('settings.json', {}); S.special = st.special || {}; S.exportPath = st.exportPath || '';
      if (live.month === S.month || !S.viewingHist) { await openMonthNow(live.month); toast(`${live.by} updated ${live.month}`); }
      else safeRender();
    } catch (e) { /* 네트워크가 잠시 끊겨도 다음 확인 때 다시 시도 */ }
  }
  // 입력 중인 칸이 있으면 화면을 바로 다시 그리지 않고, 칸을 벗어날 때 그린다
  function safeRender() { const a = document.activeElement; if (a && a.matches && a.matches('main input[type=text], main input.cell, main input[type=search]')) { S.pendingRender = true; return; } render(); }

  function recompute() {
    const [y, m] = S.month.split('-').map(Number);
    const act = S.inputs.filter((i) => i.output !== 'ignore');
    S.res = act.length ? W.build(act, { year: y, month: m, master: S.master, special: S.special, overrides: S.overrides }) : null;
  }

  // ───────── 만들기 ─────────
  async function createFiles() {
    const r = S.res, files = []; const dir = S.month;
    try {
      for (const o of r.outputs) { const f = W.fill(S.templates[o.base].bytes, o, S.valueDate, { filePath: S.exportPath ? S.exportPath.replace(/\{month\}/gi, S.month) : '' }); files.push({ name: o.fileName, data: f.bytes, key: o.key, count: f.check.count, total: f.check.total }); }
      for (const f of files) await S.store.write(dir + '/Accounting/' + f.name, f.data);
      // 마스터 반영: 체크된 신규 인원, HR 계좌로 바꾸기로 한 사람
      const adds = [], upds = [];
      for (const o of r.outputs) for (const row of o.rows) { if (!row.inMaster && S.addNew[row.id] !== false) adds.push(row); else if (row.inMaster && (S.overrides[row.id] || {}).useHr) upds.push(row); }
      if (adds.length || upds.length) await masterApply((cur) => { for (const x of adds) cur[x.id] = { name: x.name, bank: x.bank, iban: x.iban, added: S.month }; for (const x of upds) cur[x.id] = Object.assign({}, cur[x.id], { bank: x.bank, iban: x.iban, updated: S.month }); });
      const log = { version: VERSION, month: S.month, valueDate: S.valueDate, createdAt: new Date().toISOString(), createdBy: userName || '', grandTotal: W.r2(r.outputs.reduce((a, o) => a + o.total, 0)),
        inputs: S.inputs.map((i) => ({ fileName: i.fileName, output: i.output, count: i.employees.length })),
        outputs: r.outputs.map((o) => ({ key: o.key, label: o.label, fileName: o.fileName, count: o.count, total: o.total, people: o.rows.map((x) => ({ id: x.id, name: x.name, salary: x.salary })) })),
        notices: r.notices.map((n) => ({ level: n.level, type: n.type, output: n.output, id: n.id, name: n.name, msg: n.msg })) };
      await writeJson(dir + '/run.json', log);
      const created = { at: log.createdAt, by: userName || '', added: adds.length, files: files.map((f) => ({ name: f.name, count: f.count, total: f.total })) };
      S.master = await readJson('master.json', S.master);
      await editDraft((d) => { d.created = created; d.overrides = {}; d.addNew = {}; d.ack = false; });
    } catch (e) { toast(e.message); console.error(e); }
  }

  // ───────── 화면 ─────────
  function render() {
    document.documentElement.lang = lang;
    const un = $('#userName'); if (un && document.activeElement !== un) un.value = userName;
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
    if (S.store.kind !== 'folder' && window.showDirectoryPicker) h += `<div class="box new"><h3>Connect the team's shared folder first</h3>Click <b>Connect shared folder</b> (top right). In the folder window, type <code>\\\\JKPC\\WPS Payroll</code> in the <b>Folder:</b> box and click <b>Select Folder</b>, then <b>Edit files</b> / <b>Allow</b>. Do not pick a drive letter such as Z: — the browser silently refuses whole drives.</div>`;
    if (S.connectMsg) h += `<div class="box err"><h3>Shared folder not connected</h3>${esc(S.connectMsg)}<div class="small muted" style="margin-top:6px">Browser: ${esc(navigator.userAgent.match(/(Edg|Chrome)\/[\d.]+/g)?.join(' ') || navigator.userAgent)}</div></div>`;
    if (S.viewingHist) h += `<div class="box new">Viewing saved month <b>${esc(S.month)}</b> from History. Changes you make here are saved to this month and shown to the team. ${S.live && S.live.month !== S.month ? `<button class="link" data-act="goLive">Go to the month the team is working on (${esc(S.live.month)})</button>` : ''}</div>`;
    if (S.lastChange) h += `<p class="small muted" style="margin:0 0 10px">Shared with the team · last change by <b>${esc(S.lastChange.by)}</b> at ${esc((S.lastChange.at || '').slice(0, 16).replace('T', ' '))} · screens update every 3 seconds</p>`;
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
    const missing = S.res.outputs.filter((o) => !S.templates[o.base] || !S.templates[o.base].bytes), grand = S.res.outputs.reduce((a, o) => a + o.total, 0);
    h += `<section class="card"><h2><span class="step">4</span>${t('s4')}</h2><table><tr><th>${t('out')}</th><th class="n">${t('people')}</th><th class="n">${t('sum')}</th><th>${t('tpl')}</th><th>${t('vsPrev')}</th></tr>`;
    for (const o of S.res.outputs) {
      const tp = S.templates[o.base]; const tag = !tp || !tp.bytes ? `<span class="tag err">${t('tplMissing')}</span>` : tp.std ? `<span class="tag ok">${t('tplOk')}</span>` : `<span class="tag warn">${t('tplOdd')}</span>`;
      let cmp = `<span class="muted small">${t('noPrev')}</span>`; const p = S.prev && S.prev.outputs.find((x) => x.key === o.key);
      if (p) { const pid = new Set(p.people.map((x) => x.id)), cid = new Set(o.rows.map((x) => x.id)); const joined = o.rows.filter((x) => !pid.has(x.id)), left = p.people.filter((x) => !cid.has(x.id)); const dc = o.count - p.count, dt = o.total - p.total;
        cmp = `${dc >= 0 ? '+' : ''}${dc} · ${dt >= 0 ? '+' : ''}${money(dt)}` + (joined.length || left.length ? ` <details style="display:inline-block;vertical-align:top"><summary>${t('joined')} ${joined.length} / ${t('left')} ${left.length}</summary><div class="small">${joined.map((x) => '+ ' + esc(x.name)).concat(left.map((x) => '− ' + esc(x.name))).join('<br>')}</div></details>` : ''); }
      h += `<tr${o.fs ? ' style="background:#fff8ec"' : ''}><td><b>${esc(o.label)}</b>${o.fs ? ' <span class="tag warn">separate payment</span>' : ''}<div class="small muted">${esc(o.fileName)}</div></td><td class="n">${o.count}</td><td class="n">${money(o.total)}</td><td>${tag}</td><td>${cmp}</td></tr>`;
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
    if (!S.created) return ''; const c = S.created, dir = S.month;
    const who = c.by ? ` · by ${esc(c.by)}` : '', when = c.at ? ` · ${esc(c.at.slice(0, 16).replace('T', ' '))}` : '';
    return `<section class="card"><div class="box ok"><h3>${t('done', esc((S.store.kind === 'folder' ? S.store.name + '/' : '') + dir + '/Accounting'))}</h3>${t('selfOk')}${c.added ? '<br>' + t('masterAdded', c.added) : ''}<div class="small muted">Created${when}${who}</div></div><table><tr><th>${t('file')}</th><th class="n">${t('people')}</th><th class="n">${t('sum')}</th><th></th></tr>${c.files.map((f, i) => `<tr><td>${esc(f.name)}</td><td class="n">${f.count}</td><td class="n">${money(f.total)}</td><td><button class="link" data-act="dl" data-i="${i}">${t('dl')}</button></td></tr>`).join('')}</table><p style="margin-top:12px"><button class="btn" data-act="dlAll">${t('dlAll')}</button></p><p class="small muted" style="margin-top:8px">Export button does not work in Excel? <button class="link" data-tab="help">See the 3 checks in How it works</button></p></section>`;
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
    const opts = W.OUTPUTS.filter((o) => !o.fs).map((o) => `<option value="${o.key}">${esc(o.label)}</option>`).join('');
    return `<section class="card"><h2>Always in a separate file</h2><p class="hint">People listed here always go to the chosen output file, whichever HR file they come from (for example the D360 bank person and the PDH trainee). Saved in the shared folder, not in the app code.</p>
      <table><tr><th>${t('id')}</th><th>${t('name')}</th><th>${t('out')}</th><th></th></tr>${rows.map(([id, k]) => `<tr><td class="mono">${esc(id)}</td><td>${esc((S.master[id] || {}).name || '')}</td><td>${esc(outLabel(k))}</td><td><button class="link" data-act="spDel" data-id="${esc(id)}">${t('del')}</button></td></tr>`).join('')}
      <tr><td><input class="cell mono" id="sp_id" maxlength="10" placeholder="10-digit ID" style="width:120px"></td><td></td><td><select id="sp_out">${opts}</select></td><td><button class="link" data-act="spAdd">Add</button></td></tr></table></section>`;
  }
  async function saveSpecial(fn) { const cur = await readJson('settings.json', {}); cur.special = cur.special || {}; fn(cur.special); await writeJson('settings.json', cur); S.special = cur.special; recompute(); render(); }
  function exportPathCard() {
    const ex = S.exportPath ? S.exportPath.replace(/\{month\}/gi, S.month) : '';
    return `<section class="card"><h2>Export folder (cell F18)</h2><p class="hint">The Export macro saves the bank text file in the folder written in cell F18. Set it once for the team; <code>{month}</code> becomes the payroll month. Leave empty to keep the template's own path. Example: <code>Z:\{month}\Accounting</code></p>
      <div class="row"><input type="text" id="exp_path" value="${esc(S.exportPath)}" placeholder="Z:\{month}\Accounting" style="min-width:320px;font-family:Consolas,monospace"><button class="btn" data-act="expSave">Save</button></div>
      ${ex ? `<p class="small muted" style="margin-top:8px">This month's files get F18 = <code>${esc(ex)}</code> (the D360 bank file has no F18 and is left as is).</p>` : ''}</section>`;
  }
  function viewTpl() {
    let h = exportPathCard() + `<section class="card"><h2>${t('tTitle')}</h2><p class="hint">${t('tHint')} <b>Use the accounting team's final bank files only, never the HR files.</b></p>${(S.tplMsg || []).length ? `<div class="box err"><h3>Some files were not registered</h3><ul>${S.tplMsg.map((m) => `<li>${esc(m)}</li>`).join('')}</ul></div>` : ''}<div class="drop" data-drop="tpl">${t('dropTpl')}<input type="file" multiple accept=".xls,.zip" class="hide" data-file="tpl"></div><table style="margin-top:14px"><tr><th>${t('out')}</th><th>${t('tpl')}</th><th class="n">${t('tPeople')}</th><th></th></tr>`;
    for (const o of W.OUTPUTS.filter((x) => !x.fs)) { const tp = S.templates[o.key]; h += `<tr><td><b>${esc(o.label)}</b></td><td>${!tp ? `<span class="tag err">${t('tplMissing')}</span>` : tp.error ? `<span class="tag err">${esc(tp.error)}</span>` : tp.std ? `<span class="tag ok">${t('tStd')}</span>` : `<span class="tag warn">${t('tplOdd')}</span><div class="small" style="max-width:520px;margin-top:4px">${t('tOddMsg')}</div>`}</td><td class="n">${tp && tp.bytes ? tp.rows : ''}</td><td><button class="link" data-pick="tpl1" data-key="${o.key}">${t('tUpload')}</button></td></tr>`; }
    return h + `</table><input type="file" accept=".xls" class="hide" data-file="tpl1"></section>`;
  }
  async function viewHist() {
    const months = (await S.store.list('')).filter((n) => /^\d{4}-\d{2}$/.test(n)).sort().reverse(); const el = $('#hist'); if (!el) return;
    const rows = [];
    for (const mth of months) { const log = await readJson(mth + '/run.json', null), d = await readJson(mth + '/draft.json', null); if (!log && !d) continue; rows.push({ mth, log, d }); }
    if (!rows.length) { el.innerHTML = `<p class="muted">${t('hEmpty')}</p>`; return; }
    const nSel = rows.filter((r) => S.histSel[r.mth]).length;
    let h = `<p class="hint">Click a month to open it in Monthly run. Tick months and press Delete to remove them from the shared folder (HR files, bank files and records of that month). The master list is not changed.</p>
      <p><button class="btn" data-act="histDel" ${nSel ? '' : 'disabled'} style="${nSel ? 'border-color:var(--err);color:var(--err)' : ''}">Delete selected (${nSel})</button></p>
      <table><tr><th style="width:28px"><input type="checkbox" data-act="histAll" ${nSel && nSel === rows.length ? 'checked' : ''}></th><th>Month</th><th>Status</th><th class="n">${t('people')}</th><th class="n">${t('sum')}</th><th>Value date</th><th>Last change</th></tr>`;
    for (const { mth, log, d } of rows) {
      const created = log || (d && d.created);
      const status = log ? `<span class="tag ok">Created</span>` : `<span class="tag warn">In progress</span>`;
      const ppl = log ? log.outputs.reduce((a, o) => a + o.count, 0) : '', tot = log ? money(log.grandTotal) : '';
      const lc = d && d.by ? `${esc(d.by)} · ${esc((d.at || '').slice(0, 16).replace('T', ' '))}` : log ? `${esc(log.createdBy || '')} · ${esc(log.createdAt.slice(0, 16).replace('T', ' '))}` : '';
      h += `<tr><td><input type="checkbox" data-act="histSel" data-m="${mth}" ${S.histSel[mth] ? 'checked' : ''}></td><td><button class="link" data-act="histOpen" data-m="${mth}"><b>${mth}</b></button>${S.live && S.live.month === mth ? ' <span class="tag new">team is here</span>' : ''}</td><td>${status}</td><td class="n">${ppl}</td><td class="n">${tot}</td><td class="mono">${esc((log && log.valueDate) || (d && d.valueDate) || '')}</td><td class="small">${lc}</td></tr>`;
      if (log) h += `<tr><td></td><td colspan="6" class="small muted" style="padding-top:0">${log.outputs.map((o) => `<button class="link" data-act="hdl" data-p="${esc(mth + '/Accounting/' + o.fileName)}">${esc(o.label)}</button> ${o.count} · ${money(o.total)}`).join(' &nbsp;|&nbsp; ')}</td></tr>`;
    }
    el.innerHTML = h + '</table>';
  }
  async function deleteMonths(list) {
    for (const m of list) { await S.store.remove(m); delete S.histSel[m]; }
    for (const k of Object.keys(S.hrCache)) if (list.some((m) => k.startsWith(m + '/'))) delete S.hrCache[k];
    const live = { month: S.month, rev: Date.now() + '-' + CLIENT, by: userName || 'Someone', client: CLIENT, at: new Date().toISOString(), deleted: list };
    await writeJson('live.json', live); S.live = live;
    if (list.includes(S.month)) await applyDraft(emptyDraft(S.month));
    toast('Deleted: ' + list.join(', ')); render();
  }
  function viewHelp() {
    const m = S.month, expDir = (S.exportPath || 'Z:\\{month}\\Accounting').replace('{month}', m);
    return `<section class="card" id="exportHelp"><h2>Export button does not work?</h2>
      <p>The Export button is a macro. Excel blocks it when a red bar says <b>"Microsoft blocked macros because the source of this file is not trusted"</b>. Do these 3 checks in order.</p>
      <ol style="line-height:1.7">
        <li><b>Use a team PC (Faisal or Jassim), not JKPC.</b> The bank macro contains Arabic text and only runs on Windows set up like the bank expects. On JKPC it always fails, even with the original bank file.</li>
        <li><b>Open the file from the shared folder, not from Downloads.</b> Close Excel, then open <code>${esc(expDir)}</code> in File Explorer and double-click the file there. Files in Downloads are always blocked.</li>
        <li><b>Still blocked? Unblock the file.</b> Close Excel → right-click the file → <b>Properties</b> → at the bottom tick <b>Unblock</b> → <b>OK</b> → open the file again.</li>
      </ol>
      <p class="small muted">Unblock all files of the month at once: right-click the Start button → <b>Terminal</b> (or PowerShell), paste this line and press Enter:<br><code>Get-ChildItem "${esc(expDir)}" | Unblock-File</code></p>
      <p class="small muted">After Export, the bank file <code>yyyymmdd-hhmm.txt</code> appears in the folder written in cell F18 (<code>${esc(expDir)}</code>). If the red bar comes back on a team PC every month, Excel's Trusted Locations on that PC must include <code>Z:\\</code> and <code>\\\\JKPC\\WPS Payroll\\</code> with "Allow Trusted Locations on my network" ticked — ask JK.</p></section>
    <section class="card"><h2>How it works</h2>
      <p><b>Rule:</b> amounts (salary, basic, housing, other, deduction) come from the HR file; name, bank code and IBAN come from the master list. People who are not in the master list are written from HR values and flagged.</p>
      <p><b>Template is preserved:</b> the macro, the Export button and the formatting of the bank template (.xls) stay byte-for-byte the same; only cell values change. The formulas in the salary column and in Total Salary are kept.</p>
      <p><b>Self-check:</b> every created file is read back and compared (people count, total, value date, IDs, IBANs). If anything differs, nothing is saved.</p>
      <p><b>Storage layout:</b> in the shared folder: <code>templates/</code>, <code>master.json</code>, and per month <code>2026-11/HR</code>, <code>2026-11/Accounting</code>, <code>2026-11/run.json</code>, <code>2026-11/draft.json</code> (the screen everyone shares) and <code>live.json</code> (which month the team is working on).</p>
      <p><b>Final salary:</b> rows whose Payment Description says "Final Settlement", "Final Salary" or "End of Service" (or a file whose name says so) are put automatically into a separate bank file per project, named "... -Final Salary.xls", with the description "Final Salary". Salary rows of the same person stay in the normal salary file.</p>
      <p><b>Working together:</b> every change (HR files, dates, choices) is saved to the shared folder at once; other screens check every 3 seconds and follow. Enter your name at the top so others see who changed what.</p>
      <p><b>Privacy:</b> payroll data stays in this browser and in the shared folder you choose. Nothing is sent to a server.</p><p class="muted small">v${VERSION}</p></section>`;
  }

  // ───────── 이벤트 ─────────
  document.addEventListener('click', async (ev) => {
    const el = ev.target.closest('[data-tab],[data-act],[data-pick],.drop'); if (!el) return;
    if (el.dataset.tab) { S.tab = el.dataset.tab; S.mEdit = null; render(); return; }
    if (el.classList.contains('drop')) { if (ev.target.tagName !== 'INPUT') $('input', el).click(); return; }
    if (el.dataset.pick) { const inp = $(`input[data-file="${el.dataset.pick}"]`); inp.dataset.key = el.dataset.key || ''; inp.click(); return; }
    const a = el.dataset.act;
    if (a === 'rmFile') { const fn = S.inputs[+el.dataset.i].fileName; await S.store.remove(S.month + '/HR/' + fn); await editDraft((d) => { d.inputs = d.inputs.filter((i) => i.fileName !== fn); d.created = null; }); }
    else if (a === 'clearFiles') { if (!confirm(`Remove all HR files of ${S.month} for everyone?`)) return; await S.store.remove(S.month + '/HR'); await editDraft((d) => { d.inputs = []; d.overrides = {}; d.addNew = {}; d.ack = false; d.created = null; }); }
    else if (a === 'histOpen') { S.tab = 'run'; await openMonth(el.dataset.m); S.viewingHist = true; render(); }
    else if (a === 'goLive') { S.viewingHist = false; await openMonth(S.live.month); }
    else if (a === 'histDel') { const list = Object.keys(S.histSel).filter((m) => S.histSel[m]).sort(); if (!list.length) return; if (!confirm(`Delete ${list.join(', ')} for the whole team?\n\nHR files, bank files and records of ${list.length === 1 ? 'this month' : 'these months'} will be removed from the shared folder. This cannot be undone.`)) return; await deleteMonths(list); }
    else if (a === 'create') { el.disabled = true; el.textContent = t('working'); await createFiles(); }
    else if (a === 'dl') { const f = S.created.files[+el.dataset.i]; const u = await S.store.read(S.month + '/Accounting/' + f.name); if (u) download(f.name, u, 'application/vnd.ms-excel'); else toast('File not found in the shared folder.'); }
    else if (a === 'dlAll') { const list = []; for (const f of S.created.files) { const u = await S.store.read(S.month + '/Accounting/' + f.name); if (u) list.push({ name: f.name, data: u }); } download(`WPS_${S.month}.zip`, X.zipStore(list), 'application/zip'); }
    else if (a === 'hdl') { const u = await S.store.read(el.dataset.p); if (u) download(el.dataset.p.split('/').pop(), u, 'application/vnd.ms-excel'); }
    else if (a === 'mExport') { const lines = ['ID,Name,Bank,IBAN'].concat(Object.keys(S.master).map((id) => [id, S.master[id].name, S.master[id].bank, S.master[id].iban].map((v) => '"' + String(v || '').replace(/"/g, '""') + '"').join(','))); download('master.csv', te.encode('﻿' + lines.join('\r\n')), 'text/csv'); }
    else if (a === 'expSave') { const v = $('#exp_path').value.trim(); if (!v && S.exportPath && !confirm('Clear the Export folder for the whole team? Cell F18 will then keep the path saved in each template.')) return; const cur = await readJson('settings.json', {}); cur.exportPath = v; await writeJson('settings.json', cur); S.exportPath = v; toast('Export folder saved.'); render(); }
    else if (a === 'spAdd') { const id = $('#sp_id').value.trim(), k = $('#sp_out').value; if (!/^\d{10}$/.test(id)) return toast(t('idBad')); await saveSpecial((sp) => { sp[id] = k; }); }
    else if (a === 'spDel') { const id = el.dataset.id; if (!confirm(`Remove ${(S.master[id] || {}).name || id} from "Always in a separate file" for the whole team?\n\nFrom now on this person goes into the main file of their HR project.`)) return; await saveSpecial((sp) => { delete sp[id]; }); }
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
    if (a === 'setOut') { const fn = S.inputs[+el.dataset.i].fileName, v = el.value || null; await editDraft((d) => { const e = d.inputs.find((i) => i.fileName === fn); if (e) e.output = v; d.created = null; }); }
    else if (a === 'month') { if (!/^\d{4}-\d{2}$/.test(el.value)) return; S.viewingHist = S.live ? el.value !== S.live.month : false; await openMonth(el.value); }
    else if (a === 'runDate') { const [y, m, dd] = el.value.split('-').map(Number); const vd = y ? W.valueDateFor(new Date(y, m - 1, dd)) : S.valueDate; await editDraft((d) => { d.runDate = el.value; d.valueDate = vd; }); }
    else if (a === 'valueDate') { const v = el.value.trim(); await editDraft((d) => { d.valueDate = v; }); }
    else if (a === 'ov') { const id = el.dataset.id, k = el.dataset.k, v = el.value; await editDraft((d) => { d.overrides[id] = Object.assign({}, d.overrides[id], { [k]: v }); }); }
    else if (a === 'useHr') { const id = el.dataset.id, v = el.value === '1'; await editDraft((d) => { d.overrides[id] = Object.assign({}, d.overrides[id], { useHr: v }); }); }
    else if (a === 'addNew') { const id = el.dataset.id, v = el.checked; await editDraft((d) => { d.addNew[id] = v; }); }
    else if (a === 'ack') { const v = el.checked; await editDraft((d) => { d.ack = v; }); }
    else if (a === 'histSel') { S.histSel[el.dataset.m] = el.checked; render(); }
    else if (a === 'histAll') { const boxes = document.querySelectorAll('input[data-act="histSel"]'); boxes.forEach((b) => { S.histSel[b.dataset.m] = el.checked; }); render(); }
    else if (a === 'userName') { userName = el.value.trim(); localStorage.setItem('wps.user', userName); toast(userName ? 'Name saved: ' + userName : 'Name cleared'); }
  });
  document.addEventListener('input', (ev) => { if (ev.target.dataset.act === 'mq') { S.mq = ev.target.value; const pos = ev.target.selectionStart; render(); const n = $('input[data-act="mq"]'); n.focus(); n.setSelectionRange(pos, pos); } });
  ['dragover', 'dragleave', 'drop'].forEach((type) => document.addEventListener(type, async (ev) => {
    const z = ev.target.closest && ev.target.closest('.drop'); if (type !== 'dragleave') ev.preventDefault(); if (!z) return;
    z.classList.toggle('over', type === 'dragover');
    if (type === 'drop') { const files = [...ev.dataTransfer.files]; if (z.dataset.drop === 'hr') await addHr(files); else await addTemplates(files); }
  }));
  $('#btnFolder').addEventListener('click', () => connectFolder(S.pendingHandle));
  document.addEventListener('focusout', () => { if (S.pendingRender) setTimeout(() => { if (S.pendingRender) { S.pendingRender = false; safeRender(); } }, 50); });

  (async function init() {
    try { render(); } catch (e) { console.error(e); }
    let h = null;
    try { h = window.showDirectoryPicker ? await withTimeout(idb.get('dirHandle'), 5000, 'Browser storage') : null;
      if (h) { if ((await withTimeout(h.queryPermission({ mode: 'readwrite' }), 8000, 'The shared folder')) === 'granted') S.store = folderStore(h); else S.pendingHandle = h; } } catch (e) { console.error(e); if (h) S.pendingHandle = h; }
    try {
      await withTimeout(loadShared(), 30000, 'The shared folder');
      const live = await withTimeout(readJson('live.json', null), 10000, 'The shared folder'); S.live = live;
      await withTimeout(openMonth(live && live.month ? live.month : S.month), 60000, 'Loading this month');
    } catch (e) {
      console.error(e);
      if (S.store.kind === 'folder') { S.pendingHandle = S.pendingHandle || h; S.store = localStore; }
      S.connectMsg = 'Could not read the shared folder (' + (e && e.message || e) + '). Check that you are on the office Wi-Fi and that JKPC is on, then click Reconnect (top right) or press Ctrl+F5.';
    }
    try { render(); } catch (e) { console.error(e); document.querySelector('main').innerHTML = '<div class="box err">Screen error: ' + esc(e.message) + '. Press Ctrl+F5.</div>'; }
    if (S.pendingHandle && !S.connectMsg) toast(t('perm'));
    setInterval(poll, 3000);
  })();
})();
