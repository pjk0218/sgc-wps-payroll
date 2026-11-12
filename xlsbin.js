/* xlsbin.js — 의존성 없는 .xls(BIFF8) 읽기/부분 수정 + .xlsx 읽기 라이브러리
   원칙: 양식 파일의 매크로(VBA)·버튼·서식 스트림은 바이트 그대로 두고, 시트의 셀 값만 교체한다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsBin = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const FREE = 0xFFFFFFFF, EOC = 0xFFFFFFFE, FATSECT = 0xFFFFFFFD, DIFSECT = 0xFFFFFFFC;
  const dv = (u8) => new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  function concat(parts) {
    let n = 0; for (const p of parts) n += p.length;
    const out = new Uint8Array(n); let o = 0;
    for (const p of parts) { out.set(p, o); o += p.length; }
    return out;
  }

  // ───────────── CFB (OLE2 복합 문서) ─────────────
  function cfbRead(u8) {
    const d = dv(u8);
    if (d.getUint32(0, true) !== 0xE011CFD0 || d.getUint32(4, true) !== 0xE11AB1A1) throw new Error('Not an .xls file');
    const ss = 1 << d.getUint16(30, true), mss = 1 << d.getUint16(32, true);
    const nfat = d.getUint32(44, true), dirStart = d.getUint32(48, true), cutoff = d.getUint32(56, true);
    const miniFatStart = d.getUint32(60, true); let difStart = d.getUint32(68, true);
    const secOff = (n) => (n + 1) * ss;
    const dif = []; for (let i = 0; i < 109; i++) dif.push(d.getUint32(76 + i * 4, true));
    let guard = 0;
    while (difStart < DIFSECT && guard++ < 100000) {
      const o = secOff(difStart);
      for (let i = 0; i < ss / 4 - 1; i++) dif.push(d.getUint32(o + i * 4, true));
      difStart = d.getUint32(o + ss - 4, true);
    }
    const fat = [];
    for (let k = 0; k < nfat; k++) { const s = dif[k]; if (s >= DIFSECT) continue; const o = secOff(s); for (let i = 0; i < ss / 4; i++) fat.push(d.getUint32(o + i * 4, true)); }
    const chain = (st) => { const parts = []; let g = 0; while (st < DIFSECT && g++ < 4e6) { parts.push(u8.subarray(secOff(st), secOff(st) + ss)); st = fat[st]; } return concat(parts); };
    const dirBytes = chain(dirStart), dd = dv(dirBytes);
    const entries = [];
    for (let o = 0; o + 128 <= dirBytes.length; o += 128) {
      const nl = dd.getUint16(o + 64, true);
      let name = ''; for (let i = 0; i < nl - 2; i += 2) name += String.fromCharCode(dd.getUint16(o + i, true));
      entries.push({ raw: dirBytes.slice(o, o + 128), name, type: dirBytes[o + 66], left: dd.getUint32(o + 68, true), right: dd.getUint32(o + 72, true), child: dd.getUint32(o + 76, true), start: dd.getUint32(o + 116, true), size: dd.getUint32(o + 120, true), data: null, path: null });
    }
    let miniFat = []; if (miniFatStart < DIFSECT) { const m = chain(miniFatStart), md = dv(m); for (let i = 0; i < m.length; i += 4) miniFat.push(md.getUint32(i, true)); }
    const miniStream = entries[0].start < DIFSECT ? chain(entries[0].start) : new Uint8Array(0);
    for (const e of entries) {
      if (e.type !== 2) continue;
      if (e.size < cutoff) { const parts = []; let st = e.start, g = 0; while (st < DIFSECT && g++ < 4e6) { parts.push(miniStream.subarray(st * mss, (st + 1) * mss)); st = miniFat[st]; } e.data = concat(parts).slice(0, e.size); }
      else e.data = chain(e.start).slice(0, e.size);
    }
    const walk = (i, p) => { if (i === FREE || i >= entries.length) return; const e = entries[i]; if (e.path !== null) return; e.path = p + e.name; walk(e.left, p); walk(e.right, p); if (e.type === 1 || e.type === 5) walk(e.child, p + e.name + '/'); };
    entries[0].path = ''; walk(entries[0].child, '');
    return { entries, header: u8.slice(0, 512), cutoff };
  }
  function cfbGet(cfb, path) { const e = cfb.entries.find((x) => x.path === path && x.type === 2); return e ? e.data : null; }
  function cfbSet(cfb, path, data) { const e = cfb.entries.find((x) => x.path === path && x.type === 2); if (!e) throw new Error('Stream not found: ' + path); e.data = data; }
  // 디렉터리 트리(이름·CLSID·시간)는 그대로 두고 섹터 배치만 새로 계산해 512바이트 섹터 파일로 기록
  function cfbWrite(cfb) {
    const SS = 512, MSS = 64, CUT = 4096, ents = cfb.entries;
    const mini = [], big = [];
    for (const e of ents) if (e.type === 2) (e.data.length < CUT ? mini : big).push(e);
    // 미니 스트림
    const miniFat = []; const miniParts = []; let mpos = 0;
    for (const e of mini) {
      const n = Math.ceil(e.data.length / MSS);
      e._start = n ? mpos : EOC;
      for (let i = 0; i < n; i++) miniFat.push(i === n - 1 ? EOC : mpos + i + 1);
      const pad = new Uint8Array(n * MSS); pad.set(e.data); miniParts.push(pad); mpos += n;
    }
    const miniStream = concat(miniParts);
    const nDir = Math.ceil(ents.length * 128 / SS), nMiniFat = Math.ceil(miniFat.length * 4 / SS), nMiniStream = Math.ceil(miniStream.length / SS);
    let nBig = 0; for (const e of big) nBig += Math.ceil(e.data.length / SS);
    const dataSecs = nDir + nMiniFat + nMiniStream + nBig;
    let nFat = 0, nDif = 0;
    for (;;) { const need = Math.ceil((dataSecs + nFat + nDif) / (SS / 4)); const needDif = need > 109 ? Math.ceil((need - 109) / (SS / 4 - 1)) : 0; if (need === nFat && needDif === nDif) break; nFat = need; nDif = needDif; }
    const total = dataSecs + nFat + nDif;
    const out = new Uint8Array((total + 1) * SS), od = dv(out);
    const fat = new Uint32Array(nFat * (SS / 4)).fill(FREE);
    let sec = 0;
    const difSecs = []; for (let i = 0; i < nDif; i++) { difSecs.push(sec); fat[sec++] = DIFSECT; }
    const fatSecs = []; for (let i = 0; i < nFat; i++) { fatSecs.push(sec); fat[sec++] = FATSECT; }
    const put = (bytes, nsec) => { if (!nsec) return EOC; const st = sec; out.set(bytes, (st + 1) * SS); for (let i = 0; i < nsec; i++) fat[sec + i] = i === nsec - 1 ? EOC : sec + i + 1; sec += nsec; return st; };
    const dirStartSec = sec; for (let i = 0; i < nDir; i++) fat[sec + i] = i === nDir - 1 ? EOC : sec + i + 1; sec += nDir;
    const mfBytes = new Uint8Array(nMiniFat * SS).fill(0xFF); const mfd = dv(mfBytes); miniFat.forEach((v, i) => mfd.setUint32(i * 4, v, true));
    const miniFatStartSec = put(mfBytes, nMiniFat);
    const miniStreamStartSec = put(miniStream, nMiniStream);
    for (const e of big) e._start = put(e.data, Math.ceil(e.data.length / SS));
    // 디렉터리
    const dirBytes = new Uint8Array(nDir * SS);
    ents.forEach((e, i) => {
      const r = e.raw.slice(); const rd = dv(r);
      if (i === 0) { rd.setUint32(116, miniStreamStartSec, true); rd.setUint32(120, miniStream.length, true); rd.setUint32(124, 0, true); }
      else if (e.type === 2) { rd.setUint32(116, e._start, true); rd.setUint32(120, e.data.length, true); rd.setUint32(124, 0, true); }
      else if (e.type === 1) { rd.setUint32(116, 0, true); rd.setUint32(120, 0, true); rd.setUint32(124, 0, true); }
      dirBytes.set(r, i * 128);
    });
    for (let i = ents.length; i < nDir * SS / 128; i++) { const o = i * 128; const x = dv(dirBytes); x.setUint32(o + 68, FREE, true); x.setUint32(o + 72, FREE, true); x.setUint32(o + 76, FREE, true); }
    out.set(dirBytes, (dirStartSec + 1) * SS);
    // FAT, DIFAT
    fatSecs.forEach((s, k) => { for (let i = 0; i < SS / 4; i++) od.setUint32((s + 1) * SS + i * 4, fat[k * (SS / 4) + i], true); });
    difSecs.forEach((s, k) => { const base = (s + 1) * SS; for (let i = 0; i < SS / 4 - 1; i++) { const idx = 109 + k * (SS / 4 - 1) + i; od.setUint32(base + i * 4, idx < nFat ? fatSecs[idx] : FREE, true); } od.setUint32(base + SS - 4, k === nDif - 1 ? EOC : difSecs[k + 1], true); });
    // 헤더
    out.set(cfb.header.subarray(0, 24), 0);
    od.setUint16(24, 0x003E, true); od.setUint16(26, 3, true); od.setUint16(28, 0xFFFE, true); od.setUint16(30, 9, true); od.setUint16(32, 6, true);
    od.setUint32(40, 0, true); od.setUint32(44, nFat, true); od.setUint32(48, dirStartSec, true); od.setUint32(52, 0, true); od.setUint32(56, CUT, true);
    od.setUint32(60, nMiniFat ? miniFatStartSec : EOC, true); od.setUint32(64, nMiniFat, true);
    od.setUint32(68, nDif ? difSecs[0] : EOC, true); od.setUint32(72, nDif, true);
    for (let i = 0; i < 109; i++) od.setUint32(76 + i * 4, i < nFat ? fatSecs[i] : FREE, true);
    return out;
  }

  // ───────────── BIFF8 ─────────────
  const R = { BOF: 0x809, EOF: 0x0A, BOUNDSHEET: 0x85, SST: 0xFC, CONTINUE: 0x3C, EXTSST: 0xFF, ROW: 0x208, DBCELL: 0xD7, INDEX: 0x20B, DIMENSIONS: 0x200, DEFCOLWIDTH: 0x55,
    LABELSST: 0xFD, NUMBER: 0x203, RK: 0x27E, MULRK: 0xBD, BLANK: 0x201, MULBLANK: 0xBE, FORMULA: 0x06, STRING: 0x207, SHRFMLA: 0x4BC, ARRAY: 0x221, TABLE: 0x236, BOOLERR: 0x205, LABEL: 0x204, RSTRING: 0xD6, FILEPASS: 0x2F };
  function records(u8, from, to) {
    const d = dv(u8), out = []; let p = from || 0; const end = to === undefined ? u8.length : to;
    while (p + 4 <= end) { const id = d.getUint16(p, true), len = d.getUint16(p + 2, true); if (p + 4 + len > end) break; out.push({ id, off: p, len }); p += 4 + len; }
    return out;
  }
  function rkVal(rk) { let v; if (rk & 2) v = rk >> 2; else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, rk & 0xFFFFFFFC, true); b.setUint32(0, 0, true); v = b.getFloat64(0, true); } return (rk & 1) ? v / 100 : v; }
  // SST: 문자열 목록 + 각 문자열의 스트림 위치(EXTSST 재생성용)
  function parseSST(u8, recs, iSst) {
    const d = dv(u8); const segs = [recs[iSst]]; let j = iSst + 1; while (j < recs.length && recs[j].id === R.CONTINUE) segs.push(recs[j++]);
    const total = d.getUint32(recs[iSst].off + 4, true), uniq = d.getUint32(recs[iSst].off + 8, true);
    let si = 0, p = segs[0].off + 12, segEnd = segs[0].off + 4 + segs[0].len;
    const strs = [], pos = [];
    const next = () => { si++; if (si >= segs.length) throw new Error('Damaged string table'); p = segs[si].off + 4; segEnd = segs[si].off + 4 + segs[si].len; };
    for (let n = 0; n < uniq; n++) {
      if (p >= segEnd) next();
      pos.push({ abs: p, rel: p - segs[si].off });
      const cch = d.getUint16(p, true); let flags = u8[p + 2]; p += 3;
      let cRun = 0, cbExt = 0; if (flags & 8) { cRun = d.getUint16(p, true); p += 2; } if (flags & 4) { cbExt = d.getUint32(p, true); p += 4; }
      let s = '', left = cch;
      while (left > 0) {
        if (p >= segEnd) { next(); flags = (flags & ~1) | (u8[p] & 1); p++; }
        const wide = flags & 1, avail = wide ? (segEnd - p) >> 1 : segEnd - p, take = Math.min(left, avail);
        if (wide) for (let i = 0; i < take; i++) s += String.fromCharCode(d.getUint16(p + i * 2, true)); else for (let i = 0; i < take; i++) s += String.fromCharCode(u8[p + i]);
        p += wide ? take * 2 : take; left -= take;
      }
      let skip = cRun * 4 + cbExt;
      while (skip > 0) { if (p >= segEnd) next(); const t = Math.min(skip, segEnd - p); p += t; skip -= t; }
      strs.push(s);
    }
    return { strs, pos, total, uniq, first: iSst, last: j - 1, segs };
  }
  function parseWorkbook(wb) {
    const recs = records(wb);
    if (!recs.length || recs[0].id !== R.BOF) throw new Error('Unexpected workbook format');
    const d = dv(wb); const sheets = []; let iSst = -1, iExt = -1, gEnd = -1, depth = 0;
    for (let i = 0; i < recs.length; i++) {
      const r = recs[i];
      if (r.id === R.BOF) depth++; else if (r.id === R.EOF) { depth--; if (depth === 0) { gEnd = i; break; } }
      else if (r.id === R.FILEPASS) throw new Error('Password-protected files cannot be processed');
      else if (r.id === R.BOUNDSHEET) { const cch = wb[r.off + 10], wide = wb[r.off + 11] & 1; let nm = ''; for (let k = 0; k < cch; k++) nm += String.fromCharCode(wide ? d.getUint16(r.off + 12 + k * 2, true) : wb[r.off + 12 + k]); sheets.push({ name: nm, off: d.getUint32(r.off + 4, true), recOff: r.off }); }
      else if (r.id === R.SST) iSst = i; else if (r.id === R.EXTSST) iExt = i;
    }
    const gEndOff = recs[gEnd].off + 4 + recs[gEnd].len;
    const sst = iSst >= 0 ? parseSST(wb, recs, iSst) : { strs: [], pos: [], total: 0, uniq: 0 };
    sheets.forEach((s, i) => { s.end = i + 1 < sheets.length ? sheets.slice().sort((a, b) => a.off - b.off).find((x) => x.off > s.off)?.off ?? wb.length : wb.length; });
    for (const s of sheets) { const later = sheets.filter((x) => x.off > s.off).map((x) => x.off); s.end = later.length ? Math.min(...later) : wb.length; }
    return { wb, recs, sheets, sst, iSst, iExt, gEndOff };
  }
  // 시트의 셀 값 읽기 → Map("r,c" → 값) (r,c는 0부터)
  function readSheetCells(pw, sheet) {
    const wb = pw.wb, d = dv(wb), recs = records(wb, sheet.off, sheet.end), cells = new Map(), xfs = new Map(), formulas = new Set(); let maxR = -1, maxC = -1;
    const set = (r, c, v, xf) => { cells.set(r + ',' + c, v); xfs.set(r + ',' + c, xf); if (r > maxR) maxR = r; if (c > maxC) maxC = c; };
    let depth = 0, pendStr = null;
    for (let i = 0; i < recs.length; i++) {
      const { id, off, len } = recs[i], p = off + 4;
      if (id === R.BOF) { depth++; continue; } if (id === R.EOF) { depth--; if (depth === 0) break; continue; } if (depth !== 1) continue;
      if (id === R.LABELSST) set(d.getUint16(p, true), d.getUint16(p + 2, true), pw.sst.strs[d.getUint32(p + 6, true)], d.getUint16(p + 4, true));
      else if (id === R.NUMBER) set(d.getUint16(p, true), d.getUint16(p + 2, true), d.getFloat64(p + 6, true), d.getUint16(p + 4, true));
      else if (id === R.RK) set(d.getUint16(p, true), d.getUint16(p + 2, true), rkVal(d.getInt32(p + 6, true)), d.getUint16(p + 4, true));
      else if (id === R.MULRK) { const r = d.getUint16(p, true), c0 = d.getUint16(p + 2, true), n = (len - 6) / 6; for (let k = 0; k < n; k++) set(r, c0 + k, rkVal(d.getInt32(p + 4 + k * 6 + 2, true)), d.getUint16(p + 4 + k * 6, true)); }
      else if (id === R.BOOLERR) { if (wb[p + 7] === 0) set(d.getUint16(p, true), d.getUint16(p + 2, true), wb[p + 6] ? true : false, d.getUint16(p + 4, true)); }
      else if (id === R.FORMULA) { const r = d.getUint16(p, true), c = d.getUint16(p + 2, true), xf = d.getUint16(p + 4, true); formulas.add(r + ',' + c); xfs.set(r + ',' + c, xf); if (d.getUint16(p + 12, true) === 0xFFFF) { const t = wb[p + 6]; if (t === 0) pendStr = [r, c, xf]; else if (t === 1) set(r, c, !!wb[p + 8], xf); else if (t === 3) set(r, c, '', xf); } else set(r, c, d.getFloat64(p + 6, true), xf); }
      else if (id === R.STRING && pendStr) { const cch = d.getUint16(p, true), wide = wb[p + 2] & 1; let s = ''; for (let k = 0; k < cch; k++) s += String.fromCharCode(wide ? d.getUint16(p + 3 + k * 2, true) : wb[p + 3 + k]); set(pendStr[0], pendStr[1], s, pendStr[2]); pendStr = null; }
      else if (id === R.LABEL || id === R.RSTRING) { const cch = d.getUint16(p + 6, true), wide = wb[p + 8] & 1; let s = ''; for (let k = 0; k < cch; k++) s += String.fromCharCode(wide ? d.getUint16(p + 9 + k * 2, true) : wb[p + 9 + k]); set(d.getUint16(p, true), d.getUint16(p + 2, true), s, d.getUint16(p + 4, true)); }
      else if (id === R.BLANK) xfs.set(d.getUint16(p, true) + ',' + d.getUint16(p + 2, true), d.getUint16(p + 4, true));
      else if (id === R.MULBLANK) { const r = d.getUint16(p, true), c0 = d.getUint16(p + 2, true), n = (len - 6) / 2; for (let k = 0; k < n; k++) xfs.set(r + ',' + (c0 + k), d.getUint16(p + 4 + k * 2, true)); }
    }
    return { cells, xfs, formulas, maxR, maxC };
  }
  function readXls(u8, sheetName) {
    const cfb = cfbRead(u8); const wb = cfbGet(cfb, 'Workbook') || cfbGet(cfb, 'Book'); if (!wb) throw new Error('No workbook found in the file');
    const pw = parseWorkbook(wb); const sh = sheetName ? pw.sheets.find((s) => s.name === sheetName) : pw.sheets[0]; if (!sh) throw new Error('Sheet not found: ' + sheetName);
    const r = readSheetCells(pw, sh); r.sheetNames = pw.sheets.map((s) => s.name); r.sheet = sh.name; return r;
  }

  // ───────────── 셀 값 교체(양식 보존) ─────────────
  const CELL_IDS = new Set([0xFD, 0x203, 0x27E, 0xBD, 0x201, 0xBE, 0x06, 0x205, 0x204, 0xD6]);
  const FOLLOW_IDS = new Set([0x207, 0x4BC, 0x221, 0x236, 0x3C]);
  function mkRec(id, len) { const b = new Uint8Array(4 + len), d = dv(b); d.setUint16(0, id, true); d.setUint16(2, len, true); return b; }
  function recNumber(r, c, xf, v) { const b = mkRec(0x203, 14), d = dv(b); d.setUint16(4, r, true); d.setUint16(6, c, true); d.setUint16(8, xf, true); d.setFloat64(10, v, true); return b; }
  function recLabelSst(r, c, xf, i) { const b = mkRec(0xFD, 10), d = dv(b); d.setUint16(4, r, true); d.setUint16(6, c, true); d.setUint16(8, xf, true); d.setUint32(10, i, true); return b; }
  function recBlank(r, c, xf) { const b = mkRec(0x201, 6), d = dv(b); d.setUint16(4, r, true); d.setUint16(6, c, true); d.setUint16(8, xf, true); return b; }
  function encStr(s) { let wide = false; for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) > 255) { wide = true; break; } const b = new Uint8Array(3 + s.length * (wide ? 2 : 1)), d = dv(b); d.setUint16(0, s.length, true); b[2] = wide ? 1 : 0; for (let i = 0; i < s.length; i++) { if (wide) d.setUint16(3 + i * 2, s.charCodeAt(i), true); else b[3 + i] = s.charCodeAt(i); } return b; }


  // 수식 토큰(rgce)을 다른 행으로 옮기기: 공유수식(RefN)·일반 상대참조 모두 지원. 모르는 토큰이면 null
  function relocateRgce(rg, shared, r, c, dRow) {
    const out = rg.slice(), d = dv(out); let p = 0;
    const fixRef = (o, isN) => { const cf = d.getUint16(o + 2, true); let row = d.getUint16(o, true), col = cf & 0xFF;
      if (isN) { if (cf & 0x8000) row = (r + ((row << 16) >> 16)) & 0xFFFF; if (cf & 0x4000) col = (c + ((col << 24) >> 24)) & 0xFF; }
      else if (cf & 0x8000) row = (row + dRow) & 0xFFFF;
      d.setUint16(o, row, true); d.setUint16(o + 2, (cf & 0xC000) | col, true); };
    while (p < out.length) {
      const t = out[p];
      if (t >= 0x03 && t <= 0x16) { p += 1; continue; }
      if (t === 0x17) { const cch = out[p + 1], wide = out[p + 2] & 1; p += 3 + cch * (wide ? 2 : 1); continue; }
      if (t === 0x19) { const g = out[p + 1]; if (g & 0x04) { p += 4 + (d.getUint16(p + 2, true) + 1) * 2; } else p += 4; continue; }
      if (t === 0x1C || t === 0x1D) { p += 2; continue; } if (t === 0x1E) { p += 3; continue; } if (t === 0x1F) { p += 9; continue; }
      const b = t & 0x1F; if (t < 0x20 || t > 0x7F) return null;
      if (b === 0x01) { p += 3; continue; } if (b === 0x02) { p += 4; continue; } if (b === 0x03) { p += 5; continue; }
      if (b === 0x04) { if (shared) return null; fixRef(p + 1, false); p += 5; continue; }
      if (b === 0x05) { if (shared) return null; const cf1 = d.getUint16(p + 5, true), cf2 = d.getUint16(p + 7, true); if (cf1 & 0x8000) d.setUint16(p + 1, (d.getUint16(p + 1, true) + dRow) & 0xFFFF, true); if (cf2 & 0x8000) d.setUint16(p + 3, (d.getUint16(p + 3, true) + dRow) & 0xFFFF, true); p += 9; continue; }
      if (b === 0x0C) { if (!shared) return null; fixRef(p + 1, true); out[p] = t - 8; p += 5; continue; }
      if (b === 0x0D) { if (!shared) return null; const cf1 = d.getUint16(p + 5, true), cf2 = d.getUint16(p + 7, true); let r1 = d.getUint16(p + 1, true), r2 = d.getUint16(p + 3, true), c1 = cf1 & 0xFF, c2 = cf2 & 0xFF;
        if (cf1 & 0x8000) r1 = (r + ((r1 << 16) >> 16)) & 0xFFFF; if (cf2 & 0x8000) r2 = (r + ((r2 << 16) >> 16)) & 0xFFFF; if (cf1 & 0x4000) c1 = (c + ((c1 << 24) >> 24)) & 0xFF; if (cf2 & 0x4000) c2 = (c + ((c2 << 24) >> 24)) & 0xFF;
        d.setUint16(p + 1, r1, true); d.setUint16(p + 3, r2, true); d.setUint16(p + 5, (cf1 & 0xC000) | c1, true); d.setUint16(p + 7, (cf2 & 0xC000) | c2, true); out[p] = t - 8; p += 9; continue; }
      return null;
    }
    return p === out.length ? out : null;
  }
  function recFormula(r, c, xf, v, rgce) { const b = mkRec(0x06, 22 + rgce.length), d = dv(b); d.setUint16(4, r, true); d.setUint16(6, c, true); d.setUint16(8, xf, true); d.setFloat64(10, v, true); d.setUint16(18, 0, true); d.setUint32(20, 0, true); d.setUint16(24, rgce.length, true); b.set(rgce, 26); return b; }

  /* edits: [{r, c, v, f}]  f='keep' → 양식 칸이 수식이면 수식은 두고 결과값만 갱신,  f=[행,열] → 수식이 없으면 그 칸의 수식을 복제
       v = 문자열 | 숫자 | null(빈칸)
     xfFallback(r,c): 양식에 해당 셀 서식이 없을 때 쓸 XF 번호 */
  function patchXls(fileU8, sheetName, edits, xfFallback) {
    const cfb = cfbRead(fileU8); const wbName = cfbGet(cfb, 'Workbook') ? 'Workbook' : 'Book'; const wb = cfbGet(cfb, wbName);
    const pw = parseWorkbook(wb), d = dv(wb); const sheet = pw.sheets.find((s) => s.name === sheetName); if (!sheet) throw new Error('Sheet not found: ' + sheetName);
    const recs = records(wb, sheet.off, sheet.end);
    // 깊이 1 레코드만 대상 (차트 등 내장 하위 스트림은 건드리지 않음)
    let depth = 0, i0 = -1, i1 = -1, iIndex = -1, iDim = -1, iDefCol = -1;
    const dep = recs.map((r) => { if (r.id === R.BOF) return ++depth; if (r.id === R.EOF) return depth--; return depth; });
    recs.forEach((r, i) => { if (dep[i] !== 1) return; if (r.id === R.INDEX && i0 < 0) iIndex = i; if (r.id === R.DIMENSIONS && i0 < 0) iDim = i; if (r.id === R.DEFCOLWIDTH && i0 < 0) iDefCol = i; if (r.id === R.ROW || CELL_IDS.has(r.id) || r.id === R.DBCELL) { if (i0 < 0) i0 = i; i1 = i; } });
    if (i0 < 0) throw new Error('The template sheet has no cells');
    while (i1 + 1 < recs.length && FOLLOW_IDS.has(recs[i1 + 1].id) && dep[i1 + 1] === 1) i1++;
    const rows = new Map(); const getRow = (r) => { let o = rows.get(r); if (!o) { o = { rec: null, cells: [] }; rows.set(r, o); } return o; };
    let last = null;
    for (let i = i0; i <= i1; i++) {
      const { id, off, len } = recs[i], p = off + 4, bytes = wb.subarray(off, off + 4 + len);
      if (dep[i] !== 1) throw new Error('Unexpected template structure (nested stream in cell area)');
      if (id === R.ROW) { getRow(d.getUint16(p, true)).rec = bytes.slice(); last = null; }
      else if (id === R.DBCELL) { last = null; }
      else if (CELL_IDS.has(id)) { const r = d.getUint16(p, true), c0 = d.getUint16(p + 2, true); const c1 = (id === R.MULRK || id === R.MULBLANK) ? d.getUint16(p + len - 2, true) : c0; last = { c0, c1, id, parts: [bytes] }; getRow(r).cells.push(last); }
      else if (FOLLOW_IDS.has(id) && last) last.parts.push(bytes);
      else throw new Error('Unexpected template structure (record 0x' + id.toString(16) + ')');
    }
    // SST 색인
    const sstMap = new Map(); pw.sst.strs.forEach((s, i) => { if (!sstMap.has(s)) sstMap.set(s, i); });
    const newStrs = []; let strRefs = 0;
    const isst = (s) => { let i = sstMap.get(s); if (i === undefined) { i = pw.sst.strs.length + newStrs.length; newStrs.push(s); sstMap.set(s, i); } strRefs++; return i; };
    const explode = (r, el) => { const b = el.parts[0], x = dv(b), out = []; if (el.id === R.MULRK) { for (let k = 0; k <= el.c1 - el.c0; k++) { const rec = mkRec(R.RK, 10), y = dv(rec); y.setUint16(4, r, true); y.setUint16(6, el.c0 + k, true); y.setUint16(8, x.getUint16(8 + k * 6, true), true); y.setUint32(10, x.getUint32(8 + k * 6 + 2, true), true); out.push({ c0: el.c0 + k, c1: el.c0 + k, id: R.RK, parts: [rec] }); } } else { for (let k = 0; k <= el.c1 - el.c0; k++) out.push({ c0: el.c0 + k, c1: el.c0 + k, id: R.BLANK, parts: [recBlank(r, el.c0 + k, x.getUint16(8 + k * 2, true))] }); } return out; };
    const touched = new Set(); let formulaKept = 0, formulaCloned = 0, formulaFail = 0;
    for (const e of edits) {
      const row = getRow(e.r); touched.add(e.r);
      let k = row.cells.findIndex((el) => el.c0 <= e.c && e.c <= el.c1);
      if (k >= 0 && row.cells[k].c1 > row.cells[k].c0) { const ex = explode(e.r, row.cells[k]); row.cells.splice(k, 1, ...ex); k = row.cells.findIndex((el) => el.c0 === e.c); }
      if (e.f && typeof e.v === 'number') {
        if (k >= 0 && row.cells[k].id === R.FORMULA) {   // 수식 유지, 결과값만 갱신
          const el = row.cells[k], b = el.parts[0].slice(); dv(b).setFloat64(10, e.v, true); el.parts = [b, ...el.parts.slice(1).filter((q) => dv(q).getUint16(0, true) !== R.STRING)]; formulaKept++; continue;
        }
        if (Array.isArray(e.f)) {   // 기준 칸의 수식을 이 행으로 복제
          const src = (rows.get(e.f[0]) || { cells: [] }).cells.find((el) => el.c0 === e.f[1] && el.id === R.FORMULA);
          if (src) {
            const fb = src.parts[0], fd = dv(fb), cce = fd.getUint16(24, true); let rg = fb.subarray(26, 26 + cce), shared = false, ok = fb.length === 26 + cce;
            if (cce === 5 && rg[0] === 0x01) { const br = fd.getUint16(27, true), bc = fd.getUint16(29, true); const base = (rows.get(br) || { cells: [] }).cells.find((el) => el.c0 === bc && el.id === R.FORMULA); const sh = base && base.parts.find((q) => dv(q).getUint16(0, true) === R.SHRFMLA); if (sh) { const cc = dv(sh).getUint16(12, true); rg = sh.subarray(14, 14 + cc); shared = true; ok = sh.length === 14 + cc; } else ok = false; }
            const ng = ok ? relocateRgce(rg, shared, e.r, e.c, e.r - e.f[0]) : null;
            if (ng) { const xf0 = k >= 0 ? dv(row.cells[k].parts[0]).getUint16(8, true) : (xfFallback ? xfFallback(e.r, e.c) : 15); const el = { c0: e.c, c1: e.c, id: R.FORMULA, parts: [recFormula(e.r, e.c, xf0, e.v, ng)] }; if (k >= 0) row.cells[k] = el; else { row.cells.push(el); row.cells.sort((a, b) => a.c0 - b.c0); } formulaCloned++; continue; }
            formulaFail++;
          }
        }
      }
      let xf;
      if (k >= 0 && row.cells[k].id === R.FORMULA && row.cells[k].parts.some((q) => dv(q).getUint16(0, true) === R.SHRFMLA)) throw new Error('A shared-formula base cell cannot be replaced by a value: row ' + (e.r + 1));
      if (k >= 0) xf = dv(row.cells[k].parts[0]).getUint16(8, true); else { if (e.v === null || e.v === undefined || e.v === '') continue; xf = xfFallback ? xfFallback(e.r, e.c) : 15; }
      let rec;
      if (e.v === null || e.v === undefined || e.v === '') rec = recBlank(e.r, e.c, xf);
      else if (typeof e.v === 'number') { if (!isFinite(e.v)) throw new Error('Not a number: row ' + (e.r + 1)); rec = recNumber(e.r, e.c, xf, e.v); }
      else rec = recLabelSst(e.r, e.c, xf, isst(String(e.v)));
      const el = { c0: e.c, c1: e.c, id: 0, parts: [rec] };
      if (k >= 0) row.cells[k] = el; else { row.cells.push(el); row.cells.sort((a, b) => a.c0 - b.c0); }
    }
    // ── 전역 영역: SST 추가 + EXTSST 재생성
    let globals = wb.subarray(0, pw.gEndOff), newPos = [];
    if (newStrs.length) {
      if (pw.iSst < 0) throw new Error('The template has no string table');
      const sstStart = pw.recs[pw.iSst].off, lastSeg = pw.recs[pw.sst.last], sstEnd = lastSeg.off + 4 + lastSeg.len;
      const head = wb.slice(0, sstEnd), hd = dv(head); hd.setUint32(sstStart + 4, pw.sst.total + strRefs, true); hd.setUint32(sstStart + 8, pw.sst.uniq + newStrs.length, true);
      const conts = []; let cur = [], curLen = 0, base = sstEnd;
      const flush = () => { if (!cur.length) return; const rec = mkRec(R.CONTINUE, curLen); let o = 4; for (const b of cur) { rec.set(b, o); o += b.length; } conts.push(rec); base += rec.length; cur = []; curLen = 0; };
      for (const s of newStrs) { const b = encStr(s); if (b.length > 8000) throw new Error('Text is too long'); if (curLen + b.length > 8224) flush(); newPos.push({ abs: base + 4 + curLen, rel: 4 + curLen }); cur.push(b); curLen += b.length; }
      flush();
      const allPos = pw.sst.pos.concat(newPos), uniq = allPos.length;
      let mid = [], tailStart = sstEnd;
      if (pw.iExt === pw.sst.last + 1) {
        const dsst = Math.max(8, Math.floor(uniq / 128) + 1), n = Math.ceil(uniq / dsst);
        const body = new Uint8Array(2 + n * 8), bd = dv(body); bd.setUint16(0, dsst, true);
        for (let i = 0; i < n; i++) { const q = allPos[i * dsst]; bd.setUint32(2 + i * 8, q.abs, true); bd.setUint16(2 + i * 8 + 4, q.rel, true); }
        // 8224바이트 초과 시 CONTINUE 분할
        let o = 0, first = true; while (o < body.length) { const take = Math.min(first ? 8218 : 8224, body.length - o); const rec = mkRec(first ? R.EXTSST : R.CONTINUE, take); rec.set(body.subarray(o, o + take), 4); mid.push(rec); o += take; first = false; }
        const er = pw.recs[pw.iExt]; tailStart = er.off + 4 + er.len; let j = pw.iExt + 1; while (pw.recs[j].id === R.CONTINUE) { tailStart = pw.recs[j].off + 4 + pw.recs[j].len; j++; }
      }
      globals = concat([head, ...conts, ...mid, wb.subarray(tailStart, pw.gEndOff)]);
    }
    const deltaG = globals.length - pw.gEndOff;
    // ── 시트 재조립
    const rowNums = [...rows.keys()].sort((a, b) => a - b);
    let tmplRow = null; for (const r of rowNums) if (rows.get(r).rec) { tmplRow = rows.get(r).rec; break; }
    let minR = Infinity, maxR = -1, minC = Infinity, maxC = -1;
    for (const r of rowNums) {
      const row = rows.get(r);
      if (!row.rec) { if (!row.cells.length) continue; if (tmplRow) { row.rec = tmplRow.slice(); dv(row.rec).setUint16(4, r, true); dv(row.rec).setUint16(6, row.cells[0].c0, true); dv(row.rec).setUint16(8, row.cells[row.cells.length - 1].c1 + 1, true); } }
      if (row.cells.length) { const a = row.cells[0].c0, b = row.cells[row.cells.length - 1].c1 + 1; if (row.rec && touched.has(r)) { const x = dv(row.rec); if (a < x.getUint16(6, true) || x.getUint16(8, true) === 0) x.setUint16(6, a, true); if (b > x.getUint16(8, true)) x.setUint16(8, b, true); } if (a < minC) minC = a; if (b > maxC) maxC = b; }
      if (r < minR) minR = r; if (r > maxR) maxR = r;
    }
    const blocks = []; let curB = null;
    for (const r of rowNums) { const row = rows.get(r); if (!row.rec && !row.cells.length) continue; if (!curB || curB.length === 32) { curB = []; blocks.push(curB); } curB.push(r); }
    const pre = []; for (let i = 0; i < i0; i++) { const rc = recs[i]; pre.push(i === iIndex ? mkRec(R.INDEX, 16 + 4 * blocks.length) : wb.slice(rc.off, rc.off + 4 + rc.len)); }
    let preLen = 0; const preOff = pre.map((b) => { const o = preLen; preLen += b.length; return o; });
    const sheetBase = sheet.off + (sheet.off >= pw.gEndOff ? deltaG : 0);
    const table = []; let pos = preLen; const dbPos = [];
    for (const blk of blocks) {
      const blockStart = pos; let firstRowEnd = -1;
      for (const r of blk) { const rec = rows.get(r).rec; if (rec) { table.push(rec); pos += rec.length; if (firstRowEnd < 0) firstRowEnd = pos; } }
      if (firstRowEnd < 0) firstRowEnd = pos;
      const offs = []; let prev = firstRowEnd;
      for (const r of blk) { offs.push(pos - prev); prev = pos; for (const el of rows.get(r).cells) for (const b of el.parts) { table.push(b); pos += b.length; } }
      const db = mkRec(R.DBCELL, 4 + 2 * offs.length), x = dv(db); x.setUint32(4, pos - blockStart, true); offs.forEach((v, i) => x.setUint16(8 + i * 2, Math.min(v, 0xFFFF), true));
      dbPos.push(pos); table.push(db); pos += db.length;
    }
    if (iIndex >= 0) { const x = dv(pre[iIndex]); x.setUint32(4, 0, true); x.setUint32(8, minR === Infinity ? 0 : minR, true); x.setUint32(12, maxR + 1, true); x.setUint32(16, iDefCol >= 0 ? sheetBase + preOff[iDefCol] : 0, true); dbPos.forEach((p, i) => x.setUint32(20 + i * 4, sheetBase + p, true)); }
    if (iDim >= 0 && maxR >= 0) { const x = dv(pre[iDim]); x.setUint32(4, Math.min(x.getUint32(4, true), minR), true); x.setUint32(8, Math.max(x.getUint32(8, true), maxR + 1), true); if (maxC >= 0) { x.setUint16(12, Math.min(x.getUint16(12, true), minC), true); x.setUint16(14, Math.max(x.getUint16(14, true), maxC), true); } }
    const postStart = recs[i1].off + 4 + recs[i1].len;
    const newSheet = concat([...pre, ...table, wb.subarray(postStart, sheet.end)]);
    const deltaS = newSheet.length - (sheet.end - sheet.off);
    const out = concat([globals, wb.subarray(pw.gEndOff, sheet.off), newSheet, wb.subarray(sheet.end)]);
    const od = dv(out);
    for (const s of pw.sheets) { let o = s.off; if (s.off >= pw.gEndOff) o += deltaG; if (s.off > sheet.off) o += deltaS; od.setUint32(s.recOff + 4, o, true); }
    cfbSet(cfb, wbName, out);
    const res = cfbWrite(cfb); res.info = { formulaKept, formulaCloned, formulaFail, newStrings: newStrs.length };
    return res;
  }

  // ───────────── .xlsx 읽기 (zip + XML, 의존성 없음) ─────────────
  async function inflateRaw(u8) { const ds = new DecompressionStream('deflate-raw'); const w = ds.writable.getWriter(); w.write(u8); w.close(); return new Uint8Array(await new Response(ds.readable).arrayBuffer()); }
  async function unzip(u8) {
    const d = dv(u8); let e = u8.length - 22; while (e >= 0 && d.getUint32(e, true) !== 0x06054b50) e--; if (e < 0) throw new Error('Not a zip file');
    const n = d.getUint16(e + 10, true); let p = d.getUint32(e + 16, true); const files = {}; const td = new TextDecoder();
    for (let i = 0; i < n; i++) {
      const method = d.getUint16(p + 10, true), csz = d.getUint32(p + 20, true), nl = d.getUint16(p + 28, true), xl = d.getUint16(p + 30, true), cl = d.getUint16(p + 32, true), lo = d.getUint32(p + 42, true);
      const utf8 = d.getUint16(p + 8, true) & 0x800; const nameBytes = u8.subarray(p + 46, p + 46 + nl); const name = utf8 ? td.decode(nameBytes) : Array.from(nameBytes, (c) => String.fromCharCode(c)).join('');
      const ds = lo + 30 + d.getUint16(lo + 26, true) + d.getUint16(lo + 28, true); const raw = u8.subarray(ds, ds + csz);
      if (!name.endsWith('/')) files[name] = method === 0 ? raw.slice() : await inflateRaw(raw);
      p += 46 + nl + xl + cl;
    }
    return files;
  }
  const xmlDec = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-fA-F]+);/g, (m, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, '&').replace(/_x([0-9A-Fa-f]{4})_/g, (m, n) => String.fromCharCode(parseInt(n, 16)));
  const tText = (x) => { let s = ''; const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g; let m; const y = x.replace(/<rPh[\s\S]*?<\/rPh>/g, ''); while ((m = re.exec(y))) s += xmlDec(m[1]); return s; };
  async function readXlsx(u8, sheetName) {
    const files = await unzip(u8), td = new TextDecoder(); const txt = (n) => (files[n] ? td.decode(files[n]) : null);
    const wbx = txt('xl/workbook.xml'); if (!wbx) throw new Error('Not an .xlsx file');
    const rels = txt('xl/_rels/workbook.xml.rels') || ''; const relMap = {}; for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) { const id = /Id="([^"]+)"/.exec(m[0]), tg = /Target="([^"]+)"/.exec(m[0]); if (id && tg) relMap[id[1]] = tg[1]; }
    const sheets = []; for (const m of wbx.matchAll(/<sheet\b[^>]*>/g)) { const nm = /name="([^"]*)"/.exec(m[0]), rid = /r:id="([^"]+)"/.exec(m[0]); sheets.push({ name: xmlDec(nm[1]), target: relMap[rid[1]] }); }
    const sh = sheetName ? sheets.find((s) => s.name === sheetName) : sheets[0]; if (!sh) throw new Error('Sheet not found: ' + sheetName);
    const path = sh.target.startsWith('/') ? sh.target.slice(1) : 'xl/' + sh.target; const sx = txt(path); if (!sx) throw new Error('Sheet file missing');
    const ssx = txt('xl/sharedStrings.xml') || ''; const sst = []; for (const m of ssx.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g)) sst.push(m[1] ? tText(m[1]) : '');
    const cells = new Map(); let maxR = -1, maxC = -1;
    for (const m of sx.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = /\br="([A-Z]+)(\d+)"/.exec(m[1]); if (!ref || m[2] === undefined) continue; const t = (/\bt="([^"]+)"/.exec(m[1]) || [])[1];
      let c = 0; for (const ch of ref[1]) c = c * 26 + ch.charCodeAt(0) - 64; c--; const r = +ref[2] - 1;
      const vm = /<v>([\s\S]*?)<\/v>/.exec(m[2]); let v;
      if (t === 'inlineStr') v = tText(m[2]); else if (!vm) continue; else if (t === 's') v = sst[+vm[1]]; else if (t === 'str' || t === 'e') v = xmlDec(vm[1]); else if (t === 'b') v = vm[1] === '1'; else v = Number(vm[1]);
      if (v === '' || v === undefined) continue; cells.set(r + ',' + c, v); if (r > maxR) maxR = r; if (c > maxC) maxC = c;
    }
    return { cells, xfs: new Map(), formulas: new Set(), maxR, maxC, sheetNames: sheets.map((s) => s.name), sheet: sh.name };
  }
  // 무압축 zip 만들기 (여러 결과 파일을 한 번에 내려받기용)
  function crc32(u8) { let t = crc32.t; if (!t) { t = crc32.t = new Uint32Array(256); for (let i = 0; i < 256; i++) { let c = i; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[i] = c >>> 0; } } let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = t[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zipStore(list) {
    const te = new TextEncoder(), parts = [], cen = []; let off = 0;
    for (const f of list) { const nm = te.encode(f.name), crc = crc32(f.data); const h = new Uint8Array(30 + nm.length), x = dv(h); x.setUint32(0, 0x04034b50, true); x.setUint16(4, 20, true); x.setUint16(6, 0x800, true); x.setUint32(14, crc, true); x.setUint32(18, f.data.length, true); x.setUint32(22, f.data.length, true); x.setUint16(26, nm.length, true); h.set(nm, 30);
      const c = new Uint8Array(46 + nm.length), y = dv(c); y.setUint32(0, 0x02014b50, true); y.setUint16(4, 20, true); y.setUint16(6, 20, true); y.setUint16(8, 0x800, true); y.setUint32(16, crc, true); y.setUint32(20, f.data.length, true); y.setUint32(24, f.data.length, true); y.setUint16(28, nm.length, true); y.setUint32(42, off, true); c.set(nm, 46);
      parts.push(h, f.data); cen.push(c); off += h.length + f.data.length; }
    const cl = cen.reduce((a, b) => a + b.length, 0); const e = new Uint8Array(22), z = dv(e); z.setUint32(0, 0x06054b50, true); z.setUint16(8, list.length, true); z.setUint16(10, list.length, true); z.setUint32(12, cl, true); z.setUint32(16, off, true);
    return concat([...parts, ...cen, e]);
  }
  async function readAny(u8, sheetName) { if (u8[0] === 0x50 && u8[1] === 0x4B) return readXlsx(u8, sheetName); return readXls(u8, sheetName); }
  return { readXlsx, readAny, unzip, zipStore, patchXls, cfbRead, cfbWrite, cfbGet, cfbSet, records, parseWorkbook, readSheetCells, readXls, R, concat, dv, rkVal };
});
