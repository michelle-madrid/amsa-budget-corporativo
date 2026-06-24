/* global React, ReactDOM, ADA, FilterBar, ViewTabs, KpiCards, Matrix, BarYears, ComplianceBars, Donut,
   useTweaks, TweaksPanel, TweakSection, TweakSlider, TweakRadio, TweakToggle */
const { useState, useMemo, useCallback, useEffect } = React;

const TWEAK_DEFAULTS = /*EDITMODE-BEGIN*/{
  "unit": "MUSD",
  "decimals": 2,
  "thrRed": 3,
  "thrYellow": 1,
  "showCharts": true,
  "density": "regular"
}/*EDITMODE-END*/;

/* ---------- Generador .xlsx mínimo (sin librerías) ---------- */
const _CRC = (() => { const t = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; } return t; })();
function _crc32(b) { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = _CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function _colName(i) { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
function _zipStore(files) {
  const enc = new TextEncoder();
  const u16 = n => [n & 0xFF, (n >> 8) & 0xFF];
  const u32 = n => [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF];
  const local = [], central = []; let offset = 0;
  for (const f of files) {
    const nm = enc.encode(f.name), crc = _crc32(f.data), sz = f.data.length;
    const h = [0x50, 0x4b, 0x03, 0x04, ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(crc), ...u32(sz), ...u32(sz), ...u16(nm.length), ...u16(0)];
    local.push(new Uint8Array(h), nm, f.data);
    central.push({ nm, crc, sz, offset });
    offset += h.length + nm.length + sz;
  }
  const cdir = []; let cdSize = 0;
  for (const c of central) {
    const h = [0x50, 0x4b, 0x01, 0x02, ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(c.crc), ...u32(c.sz), ...u32(c.sz), ...u16(c.nm.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(c.offset)];
    cdir.push(new Uint8Array(h), c.nm); cdSize += h.length + c.nm.length;
  }
  const eocd = [0x50, 0x4b, 0x05, 0x06, ...u16(0), ...u16(0), ...u16(central.length), ...u16(central.length), ...u32(cdSize), ...u32(offset), ...u16(0)];
  cdir.push(new Uint8Array(eocd));
  return new Blob([...local, ...cdir], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
function _xlsx(sheetName, headers, rows) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const cell = (c, ref) => {
    if (c && c.t === 'n' && c.v != null && isFinite(c.v)) return `<c r="${ref}" t="n"><v>${c.v}</v></c>`;
    const v = c == null ? '' : (c.v != null ? c.v : c);
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  };
  const all = [headers.map(h => ({ t: 's', v: h })), ...rows];
  const body = all.map((cells, ri) => `<row r="${ri + 1}">${cells.map((c, ci) => cell(c, _colName(ci) + (ri + 1))).join('')}</row>`).join('');
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
  const wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${esc(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const wbr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const e = new TextEncoder();
  return _zipStore([
    { name: '[Content_Types].xml', data: e.encode(ct) },
    { name: '_rels/.rels', data: e.encode(rels) },
    { name: 'xl/workbook.xml', data: e.encode(wb) },
    { name: 'xl/_rels/workbook.xml.rels', data: e.encode(wbr) },
    { name: 'xl/worksheets/sheet1.xml', data: e.encode(sheet) },
  ]);
}

/* ---------- Pestaña Diccionario: CECO→Gerencia/VP (editable) y catálogo de Ítem ---------- */
function DictView(props) {
  const A = window.ADA;
  const { vpov, setVpov, nameov, setNameov } = props; // estado elevado al App: al editar, el Dashboard se reagrupa solo
  const [q, setQ] = React.useState('');
  const [editCeco, setEditCeco] = React.useState(null);
  const [editGer, setEditGer] = React.useState(null);   // ceco de la fila cuya Gerencia se edita
  const [editItem, setEditItem] = React.useState(null); // clave del Ítem que se edita
  const ql = q.trim().toLowerCase();
  // Renombrar Gerencia / Ítem: el override va keyed por la clave estable (no por
  // CECO). Si el valor queda vacío o igual al nombre base, se elimina el override.
  const setGerName = (gerKey, val) => setNameov(o => {
    const ger = { ...(o.ger || {}) }; const v = (val || '').trim();
    if (!v || v === A.baseGer(gerKey)) delete ger[gerKey]; else ger[gerKey] = v;
    return { ...o, ger };
  });
  const setItemName = (itemKey, val) => setNameov(o => {
    const item = { ...(o.item || {}) }; const v = (val || '').trim();
    if (!v || v === A.baseItem(itemKey)) delete item[itemKey]; else item[itemKey] = v;
    return { ...o, item };
  });
  const cecos = (window.ADA_DICT && window.ADA_DICT.cecos) || [];
  const vpKeys = Array.from(new Set(cecos.map(r => r.v))).sort((a, b) => A.dispVP(a).localeCompare(A.dispVP(b), 'es'));
  const setVp = (ceco, vpKey, origV) => setVpov(o => { const n = { ...o }; if (!vpKey || vpKey === origV) delete n[ceco]; else n[ceco] = vpKey; return n; });
  const gerOv = nameov.ger || {};
  const cecoRows = cecos.map(r => {
    const vpKey = vpov[r.c] || r.v;
    return { ceco: r.c, gerKey: r.g, ger: A.dispGer(r.g), gerOverridden: !!gerOv[r.g], vpKey, vp: A.dispVP(vpKey), origV: r.v, overridden: !!vpov[r.c] };
  }).filter(r => !ql || r.ceco.toLowerCase().includes(ql) || r.ger.toLowerCase().includes(ql) || r.vp.toLowerCase().includes(ql));
  const nOver = Object.keys(vpov).length;
  // Catálogo de Ítem Relevante = unión de TODOS los ítems presentes en los
  // datos (corporativo + distribuibles), no solo el catálogo corporativo:
  // las distribuibles aportan ítems extra (p. ej. Reactivos, Agua, Subrepartos)
  // que sí aparecen en la tabla del Dashboard.
  const itemOv = nameov.item || {};
  const itemKeySet = new Set(A.D.items || []);
  (A.corpRecords || []).forEach(r => itemKeySet.add(r.item));
  (A.distRecords || []).forEach(r => itemKeySet.add(r.item));
  const itemRows = Array.from(itemKeySet)
    .map(k => ({ key: k, name: A.dispItem(k), overridden: !!itemOv[k] }))
    .filter(r => !ql || r.name.toLowerCase().includes(ql) || r.key.toLowerCase().includes(ql))
    .sort((a, b) => a.name.localeCompare(b.name, 'es'));
  const nItemOver = Object.keys(itemOv).length;
  const nGerOver = Object.keys(gerOv).length;
  const th = { textAlign: 'left', padding: '8px 12px', fontSize: 11, fontWeight: 700, letterSpacing: '.03em', color: '#fff', background: 'var(--amsa-teal-deep)', position: 'sticky', top: 0, zIndex: 1 };
  const td = { padding: '6px 12px', fontSize: 12.5, borderBottom: '1px solid var(--line-soft)', color: 'var(--ink)' };
  const foot = { padding: '7px 12px', fontSize: 11, color: 'var(--fg-muted)', borderTop: '1px solid var(--line-soft)' };
  const iconBtn = { border: 0, background: 'transparent', cursor: 'pointer', color: 'var(--fg-muted)', fontSize: 12, padding: '0 2px', lineHeight: 1 };
  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 12, margin: '4px 0 14px' }}>
        <div>
          <h2 style={{ fontFamily: 'var(--font-disp)', fontWeight: 800, fontSize: 18, color: 'var(--ink)', margin: 0 }}>Diccionario de códigos</h2>
          <div style={{ fontSize: 12, color: 'var(--fg-muted)', marginTop: 2 }}>CECO → Gerencia (renombrable) y Vicepresidencia (editable) · Ítem Relevante (renombrable)</div>
        </div>
        <input value={q} onChange={e => setQ(e.target.value)} placeholder="Buscar CECO, Gerencia, VP o Ítem…"
          style={{ height: 32, width: 300, padding: '0 12px', border: '1px solid var(--teal-border)', borderRadius: 6, fontSize: 12.5, outline: 'none', fontFamily: 'var(--font-sans)', color: 'var(--ink)' }} />
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 16, alignItems: 'start' }}>
        <div className="panel" style={{ padding: 0, overflow: 'hidden' }}>
          <div style={{ maxHeight: 580, overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr><th style={{ ...th, width: 118 }}>CECO</th><th style={th}>Gerencia</th><th style={th}>Vicepresidencia</th></tr></thead>
              <tbody>
                {cecoRows.map((r) => (
                  <tr key={r.ceco} style={r.overridden ? { background: 'var(--accent-wash)' } : undefined}>
                    <td style={{ ...td, fontVariantNumeric: 'tabular-nums', fontWeight: 600 }}>{r.ceco}</td>
                    <td style={td}>
                      {editGer === r.ceco
                        ? <input autoFocus type="text" defaultValue={r.ger}
                            onBlur={e => { setGerName(r.gerKey, e.target.value); setEditGer(null); }}
                            onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); if (e.key === 'Escape') setEditGer(null); }}
                            style={{ width: '100%', height: 26, border: '1px solid var(--amsa-teal)', borderRadius: 5, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', background: '#fff', padding: '0 6px', boxSizing: 'border-box' }} />
                        : <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            <span style={r.gerOverridden ? { fontWeight: 700, color: 'var(--accent-ink)' } : undefined}>{r.ger}</span>
                            {r.gerOverridden && <span title={'Original: ' + A.baseGer(r.gerKey)} style={{ fontSize: 9, color: 'var(--amsa-yellow)' }}>●</span>}
                            <button title="Renombrar Gerencia" onClick={() => setEditGer(r.ceco)} style={iconBtn}>✎</button>
                            {r.gerOverridden && <button title="Restablecer nombre" onClick={() => setGerName(r.gerKey, null)} style={iconBtn}>↺</button>}
                          </span>}
                    </td>
                    <td style={td}>
                      {editCeco === r.ceco
                        ? <select autoFocus value={r.vpKey}
                            onChange={e => { setVp(r.ceco, e.target.value, r.origV); setEditCeco(null); }}
                            onBlur={() => setEditCeco(null)}
                            style={{ width: '100%', height: 26, border: '1px solid var(--amsa-teal)', borderRadius: 5, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', background: '#fff' }}>
                            {vpKeys.map(k => <option key={k} value={k}>{A.dispVP(k)}</option>)}
                          </select>
                        : <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            <span style={r.overridden ? { fontWeight: 700, color: 'var(--accent-ink)' } : undefined}>{r.vp}</span>
                            {r.overridden && <span title={'Original: ' + A.dispVP(r.origV)} style={{ fontSize: 9, color: 'var(--amsa-yellow)' }}>●</span>}
                            <button title="Editar VP" onClick={() => setEditCeco(r.ceco)} style={iconBtn}>✎</button>
                            {r.overridden && <button title="Restablecer" onClick={() => setVp(r.ceco, null, r.origV)} style={iconBtn}>↺</button>}
                          </span>}
                    </td>
                  </tr>
                ))}
                {cecoRows.length === 0 && <tr><td style={td} colSpan="3">Sin resultados</td></tr>}
              </tbody>
            </table>
          </div>
          <div style={{ ...foot, display: 'flex', justifyContent: 'space-between' }}>
            <span>{cecoRows.length} CECO{nOver > 0 ? ` · ${nOver} con VP reasignada` : ''}{nGerOver > 0 ? ` · ${nGerOver} Gerencia renombrada${nGerOver > 1 ? 's' : ''}` : ''}</span>
            {(nOver > 0 || nGerOver > 0) && <button onClick={() => { setVpov({}); setNameov(o => ({ ...o, ger: {} })); }} style={{ ...iconBtn, color: 'var(--amsa-teal)', fontWeight: 600 }}>Restablecer todo</button>}
          </div>
        </div>
        <div className="panel" style={{ padding: 0, overflow: 'hidden' }}>
          <div style={{ maxHeight: 580, overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr><th style={th}>Ítem Relevante</th></tr></thead>
              <tbody>
                {itemRows.map((it, i) => (
                  <tr key={it.key} style={i % 2 ? { background: 'var(--card-alt2)' } : undefined}>
                    <td style={td}>
                      {editItem === it.key
                        ? <input autoFocus type="text" defaultValue={it.name}
                            onBlur={e => { setItemName(it.key, e.target.value); setEditItem(null); }}
                            onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); if (e.key === 'Escape') setEditItem(null); }}
                            style={{ width: '100%', height: 26, border: '1px solid var(--amsa-teal)', borderRadius: 5, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', background: '#fff', padding: '0 6px', boxSizing: 'border-box' }} />
                        : <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, width: '100%' }}>
                            <span style={it.overridden ? { fontWeight: 700, color: 'var(--accent-ink)' } : undefined}>{it.name}</span>
                            {it.overridden && <span title={'Original: ' + A.baseItem(it.key)} style={{ fontSize: 9, color: 'var(--amsa-yellow)' }}>●</span>}
                            <button title="Renombrar Ítem" onClick={() => setEditItem(it.key)} style={{ ...iconBtn, marginLeft: 'auto' }}>✎</button>
                            {it.overridden && <button title="Restablecer nombre" onClick={() => setItemName(it.key, null)} style={iconBtn}>↺</button>}
                          </span>}
                    </td>
                  </tr>
                ))}
                {itemRows.length === 0 && <tr><td style={td}>Sin resultados</td></tr>}
              </tbody>
            </table>
          </div>
          <div style={{ ...foot, display: 'flex', justifyContent: 'space-between' }}>
            <span>{itemRows.length} ítems{nItemOver > 0 ? ` · ${nItemOver} renombrado${nItemOver > 1 ? 's' : ''}` : ''}</span>
            {nItemOver > 0 && <button onClick={() => setNameov(o => ({ ...o, item: {} }))} style={{ ...iconBtn, color: 'var(--amsa-teal)', fontWeight: 600 }}>Restablecer nombres</button>}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------- Panel de colores en vivo (editable desde el front, persiste en este equipo) ---------- */
const THEME_DEFAULTS = {
  '--amsa-teal': '#2a8a96', '--amsa-teal-deep': '#14515a', '--amsa-teal-light': '#b9dde0',
  '--amsa-yellow': '#f0a929', '--amsa-red': '#e63b2e',
  '--ok': '#1f9d57', '--yellow': '#e0a800', '--red': '#dc3545', '--blue': '#4782b4',
  '--ink': '#1f2428', '--page': '#e9ecee', '--card': '#ffffff',
};
const THEME_KEYS = ['--amsa-teal', '--amsa-teal-deep', '--amsa-teal-light', '--amsa-yellow', '--amsa-red', '--ok', '--yellow', '--red', '--blue', '--ink', '--page', '--card'];
const THEME_GROUPS = [
  { label: 'Marca', items: [['--amsa-teal', 'Principal'], ['--amsa-teal-deep', 'Teal oscuro'], ['--amsa-teal-light', 'Teal claro'], ['--amsa-yellow', 'Amarillo'], ['--amsa-red', 'Rojo']] },
  { label: 'Semáforo', items: [['--ok', 'Verde (ok)'], ['--yellow', 'Amarillo'], ['--red', 'Rojo'], ['--blue', 'Azul']] },
  { label: 'Texto y fondo', items: [['--ink', 'Texto'], ['--page', 'Fondo página'], ['--card', 'Tarjetas']] },
];
// Resto del bloque <style id=ada-tema> (grises + derivados color-mix) para reproducirlo al "Copiar".
const THEME_TAIL = '--fg-1:#313131 !important;--fg-2:#4d4d4d !important;--fg-3:#5B5C64 !important;--fg-4:#858585 !important;--fg-soft:#9aa3a8 !important;--fg-muted:#8a9499 !important;--line:#E2E5E8 !important;--line-soft:#EEF1F3 !important;--card-alt:#F6F8F9 !important;--card-alt2:#fafbfb !important;--teal-wash:color-mix(in srgb,var(--amsa-teal) 9%,#fff) !important;--teal-wash2:color-mix(in srgb,var(--amsa-teal) 5%,#fff) !important;--teal-100:color-mix(in srgb,var(--amsa-teal) 14%,#fff) !important;--teal-200:color-mix(in srgb,var(--amsa-teal) 22%,#fff) !important;--teal-300:color-mix(in srgb,var(--amsa-teal) 30%,#fff) !important;--teal-border:color-mix(in srgb,var(--amsa-teal) 26%,#fff) !important;--teal-muted:color-mix(in srgb,var(--amsa-teal) 50%,#5B5C64) !important;--accent-wash:color-mix(in srgb,var(--amsa-yellow) 16%,#fff) !important;--accent-300:color-mix(in srgb,var(--amsa-yellow) 60%,#fff) !important;--accent-ink:color-mix(in srgb,var(--amsa-yellow) 78%,#000) !important;--accent-900:color-mix(in srgb,var(--amsa-yellow) 32%,#000) !important;';
function ada_buildBlock(ov) {
  var body = '';
  THEME_KEYS.forEach(function (k) { body += k + ':' + ((ov && ov[k]) || THEME_DEFAULTS[k]) + ' !important;'; });
  return '<style id=ada-tema>/* ====== COLORES DEL DASHBOARD - edita los hex #rrggbb y guarda (no toques !important) ====== */:root{' + body + THEME_TAIL + '}</style>';
}
function ColorPanel({ onApply, edit }) {
  // "Modo Edición" se controla desde la barra de pestañas (estado elevado al App);
  // aquí solo aparece el panel/botón 🎨 Colores mientras edit está activo.
  const [open, setOpen] = React.useState(false);
  React.useEffect(() => { if (!edit) setOpen(false); }, [edit]); // al salir de edición, cierra el panel
  const [ov, setOv] = React.useState(function () { try { const s = localStorage.getItem('ada_colors_v1'); return s ? JSON.parse(s) : {}; } catch (e) { return {}; } });
  const [copied, setCopied] = React.useState(false);
  const persist = (n) => { try { localStorage.setItem('ada_colors_v1', JSON.stringify(n)); } catch (e) {} };
  const refresh = () => { if (window.ADA && window.ADA.resetTheme) window.ADA.resetTheme(); if (onApply) onApply(); };
  const setColor = (k, val) => {
    setOv(prev => { const n = Object.assign({}, prev, { [k]: val }); persist(n); return n; });
    document.documentElement.style.setProperty(k, val, 'important');
    refresh();
  };
  const reset = () => {
    Object.keys(THEME_DEFAULTS).forEach(k => document.documentElement.style.removeProperty(k));
    setOv({}); persist({}); refresh();
  };
  const copyBlock = () => {
    const block = ada_buildBlock(ov);
    let done = false;
    try { if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(block); done = true; } } catch (e) {}
    if (!done) { try { const ta = document.createElement('textarea'); ta.value = block; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta); } catch (e) {} }
    setCopied(true); setTimeout(() => setCopied(false), 3200);
  };
  const nOv = Object.keys(ov).length;
  const sw = { width: 38, height: 22, border: '1px solid var(--teal-border)', borderRadius: 5, padding: 0, background: 'none', cursor: 'pointer', flexShrink: 0 };
  const rowS = { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, padding: '3px 0', fontSize: 12, color: 'var(--fg-2)' };
  return (
    <div style={{ position: 'fixed', right: 16, bottom: 16, zIndex: 9999, display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 9, fontFamily: 'var(--font-sans)' }}>
      {edit && open && (
        <div style={{ width: 244, maxHeight: 'calc(100vh - 90px)', overflowY: 'auto', background: 'var(--card)', border: '1px solid var(--line)', borderRadius: 12, boxShadow: '0 10px 36px rgba(20,40,50,.18)', padding: '12px 14px' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
            <b style={{ fontFamily: 'var(--font-disp)', fontSize: 13, color: 'var(--ink)' }}>Colores del dashboard</b>
            <button onClick={() => setOpen(false)} title="Cerrar" style={{ border: 0, background: 'none', cursor: 'pointer', color: 'var(--fg-soft)', fontSize: 14, lineHeight: 1 }}>✕</button>
          </div>
          {THEME_GROUPS.map(g => (
            <div key={g.label}>
              <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: '.05em', textTransform: 'uppercase', color: 'var(--fg-soft)', margin: '8px 0 2px' }}>{g.label}</div>
              {g.items.map(([k, lbl]) => (
                <label key={k} style={rowS}>
                  <span>{lbl}</span>
                  <input type="color" value={(ov[k] || THEME_DEFAULTS[k]).toLowerCase()} onChange={e => setColor(k, e.target.value)} style={sw} />
                </label>
              ))}
            </div>
          ))}
          <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
            <button onClick={copyBlock} style={{ flex: 1, height: 28, border: 0, borderRadius: 7, background: 'var(--amsa-teal)', color: '#fff', fontWeight: 600, fontSize: 12, cursor: 'pointer' }}>{copied ? '¡Guardado!' : 'Guardar'}</button>
            <button onClick={reset} disabled={!nOv} style={{ height: 28, padding: '0 10px', border: '1px solid var(--teal-border)', borderRadius: 7, background: 'var(--card)', color: nOv ? 'var(--fg-2)' : 'var(--fg-soft)', fontWeight: 600, fontSize: 12, cursor: nOv ? 'pointer' : 'default' }}>Restablecer</button>
          </div>
          <div style={{ fontSize: 10.5, color: 'var(--fg-soft)', marginTop: 8, lineHeight: 1.4 }}>
            {copied
              ? 'Pega lo copiado reemplazando el bloque que empieza con «<style id=ada-tema>» en el HTML (Ctrl+F) para fijarlo para todos.'
              : 'Los cambios se guardan en este equipo. Usa «Guardar» y pégalo en el HTML para dejarlos fijos para todos.'}
          </div>
        </div>
      )}
      {edit && (
        <button onClick={() => setOpen(o => !o)} style={{ display: 'flex', alignItems: 'center', gap: 7, height: 36, padding: '0 14px', border: 0, borderRadius: 20, background: 'var(--amsa-teal)', color: '#fff', fontWeight: 700, fontSize: 12.5, fontFamily: 'var(--font-disp)', cursor: 'pointer', boxShadow: '0 4px 14px rgba(20,40,50,.22)' }}>
          <span style={{ fontSize: 15 }}>🎨</span> Colores{nOv ? ' ·' + nOv : ''}
        </button>
      )}
    </div>
  );
}

function App() {
  const A = ADA;
  const TX = window.TEXTOS;
  // Carga inicial: mezcla los defaults con lo guardado en localStorage (persiste
  // preferencias de presentación en este equipo).
  // Solo se persisten preferencias de presentación; los umbrales del semáforo
  // son regla de negocio y vienen siempre del default (no se "congelan" por sesión).
  const PERSIST_KEYS = ['unit', 'decimals', 'density', 'showCharts'];
  const initialTweaks = useMemo(() => {
    try {
      const s = localStorage.getItem('ada_tweaks_v1');
      if (s) {
        const saved = JSON.parse(s), pick = {};
        PERSIST_KEYS.forEach(k => { if (saved[k] != null) pick[k] = saved[k]; });
        return { ...TWEAK_DEFAULTS, ...pick };
      }
    } catch (e) {}
    return TWEAK_DEFAULTS;
  }, []);
  const [t, setTweak] = useTweaks(initialTweaks);
  useEffect(() => {
    try {
      const pick = {};
      PERSIST_KEYS.forEach(k => { pick[k] = t[k]; });
      localStorage.setItem('ada_tweaks_v1', JSON.stringify(pick));
    } catch (e) {}
  }, [t]);

  const [st, setSt] = useState({
    month: 11,          // año completo (sin filtro mensual: solo acumulado anual)
    dataMode: 'both',   // corp | dist | both | none (chips Corporativo/Distribuible)
    items: [], vps: [], gers: [], companies: [],
    version: 'ORI',
    years: [2025],      // años reales seleccionados (2022..2025)
    showProp: false,    // Propuesta 2027 como capa de comparación aditiva
    donutMetric: 'real', // métrica del donut: 'real' | 'prop'
    yearAgg: 'sum',     // tabla con varios años: 'sum' (acumulado) | 'avg' (promedio)
    q: '',              // búsqueda en la tabla (VP / Gerencia / Ítem)
    groupMode: 'org',   // 'org' = VP›Ger›Ítem (foto) · 'item' = Ítem›VP›Ger
    sort: { key: 'real', dir: 'desc' }, // orden de la tabla por columna
  });
  const set = useCallback(patch => setSt(s => ({ ...s, ...patch })), []);

  const [expanded, setExpanded] = useState(() => new Set());
  // Propuesta 2027 editada: persiste en este equipo (localStorage) para no perderla al recargar.
  const [overrides, setOverrides] = useState(() => {
    try { const s = localStorage.getItem('ada_prop_v1'); return s ? JSON.parse(s) : {}; } catch (e) { return {}; }
  });
  useEffect(() => {
    try { localStorage.setItem('ada_prop_v1', JSON.stringify(overrides)); } catch (e) {}
  }, [overrides]);
  // VP reasignada por CECO desde el Diccionario (override editable, persiste en
  // este equipo). Al cambiar, reasignamos la VP de los registros y el Dashboard
  // se reagrupa automáticamente.
  const [vpov, setVpov] = useState(() => {
    try { const s = localStorage.getItem('ada_vpov_v1'); return s ? JSON.parse(s) : {}; } catch (e) { return {}; }
  });
  useEffect(() => { try { localStorage.setItem('ada_vpov_v1', JSON.stringify(vpov)); } catch (e) {} }, [vpov]);
  // Aplica el override a los registros ANTES de reconstruir árbol/series/donut.
  useMemo(() => A.applyVpOverrides(vpov), [vpov]);
  // Renombres EDITABLES de Gerencia e Ítem (nombres truncados de la base). Solo
  // afectan el display (dispGer/dispItem), keyed por la clave estable → un
  // refresco de la base conserva los renombres. Persisten en localStorage.
  const [nameov, setNameov] = useState(() => {
    try { const s = localStorage.getItem('ada_nameov_v1'); const o = s ? JSON.parse(s) : {}; return { ger: o.ger || {}, item: o.item || {} }; } catch (e) { return { ger: {}, item: {} }; }
  });
  useEffect(() => { try { localStorage.setItem('ada_nameov_v1', JSON.stringify(nameov)); } catch (e) {} }, [nameov]);
  // Aplica ANTES del primer render para que tabla/filtros muestren ya el nombre editado.
  useMemo(() => A.applyNameOverrides(nameov), [nameov]);
  // Colores guardados (panel de colores): aplicarlos ANTES del primer render para
  // que charts/donut resuelvan ya con el tema del usuario (sin parpadeo).
  const [, setColorTick] = useState(0);
  const onColorApply = useCallback(() => setColorTick(t => t + 1), []);
  // "Modo Edición": estado elevado aquí para que el botón viva en la barra de pestañas
  // (junto a Diccionario) y el panel 🎨 Colores siga abajo a la derecha.
  const [editMode, setEditMode] = useState(false);
  useMemo(() => {
    try {
      const s = localStorage.getItem('ada_colors_v1'); const ov = s ? JSON.parse(s) : {};
      const d = document.documentElement;
      Object.keys(ov).forEach(k => d.style.setProperty(k, ov[k], 'important'));
      if (A.resetTheme) A.resetTheme();
    } catch (e) {}
  }, []);
  // Responsivo: el diseño está pensado para ~1480px. En pantallas más anchas/grandes
  // escala todo proporcionalmente (zoom) para que no se vea chico; tope 1.6x.
  useEffect(() => {
    const fit = () => {
      let z = parseFloat(document.documentElement.style.zoom) || 1;
      const real = window.innerWidth * z;            // ancho aprox. a zoom 1
      z = Math.min(1.6, Math.max(1, real / 1480));
      document.documentElement.style.zoom = z;
    };
    fit();
    window.addEventListener('resize', fit);
    return () => window.removeEventListener('resize', fit);
  }, []);
  const [page, setPage] = useState('dashboard'); // 'dashboard' | 'dict' (pestañas)
  const [hiddenCols, setHiddenCols] = useState([]); // columnas ocultas de la tabla
  const [pinnedCols, setPinnedCols] = useState([]); // columnas fijadas (sticky)

  const showProp = !!st.showProp;
  const histYears = st.years.filter(y => typeof y === 'number').sort((a, b) => a - b);
  const yearsLabel = histYears.length ? histYears.join(' + ') : '—';
  const multiYear = histYears.length > 1; // con >1 año se puede elegir acumulado/promedio en la tabla
  // Con más de 2 años, las tarjetas no listan los años (no caben) → "N años".
  const periodLbl = histYears.length > 2 ? `${histYears.length} años` : yearsLabel;
  const thr = { red: t.thrRed, yellow: t.thrYellow };
  // Sin simulación de 2026/2027 (no hay datos). 'ori' ajusta el plan histórico en versión Original.
  const growth = { g26: 0, gReal26: 0, g27: 0, ori: 0.955 };

  const groupDims = st.groupMode === 'item' ? ['item', 'vp', 'ger'] : ['vp', 'ger', 'item'];

  const opts = useMemo(() => ({
    years: st.years, showProp, donutMetric: st.donutMetric, yearAgg: st.yearAgg, version: st.version, dataMode: st.dataMode,
    companies: st.companies, vps: st.vps, gers: st.gers, items: st.items,
    groupBy: groupDims, sort: st.sort, overrides, growth,
  }), [st, showProp, overrides, vpov]);

  const tree = useMemo(() => A.buildTree(opts), [opts]);
  const series = useMemo(() => A.annualSeries(opts), [opts]);
  const dist = useMemo(() => A.distribuible(opts), [opts]);
  const dims = useMemo(() => A.dimsFor(opts), [opts]);
  const modeLbl = { corp: 'Corporativo', dist: 'Distribuible', both: 'Corporativo + Distribuible', none: 'Sin datos seleccionados' }[st.dataMode || 'both'];
  const corpOn = st.dataMode === 'corp' || st.dataMode === 'both';
  const distOn = st.dataMode === 'dist' || st.dataMode === 'both';
  const companiesActive = st.companies.length > 0 && st.companies.length < A.COMPANIAS.length;

  // Al buscar, expandir las ramas (VP/Gerencia) que coinciden, dejando que
  // Expandir/Colapsar sigan operando sobre ese estado.
  useEffect(() => {
    const ql = st.q.trim().toLowerCase();
    if (!ql) return;
    const D = { vp: A.dispVP, ger: A.dispGer, item: A.dispItem };
    const d0 = D[groupDims[0]], d1 = D[groupDims[1]], d2 = D[groupDims[2]];
    const keys = new Set();
    tree.vpNodes.forEach(n1 => {
      const hit1 = d0(n1.name).toLowerCase().includes(ql);
      if (hit1) keys.add(n1.name);
      n1.children.forEach(n2 => {
        const hit2 = d1(n2.name).toLowerCase().includes(ql);
        const hit3 = n2.children.some(n3 => d2(n3.name).toLowerCase().includes(ql));
        if (hit1 || hit2 || hit3) { keys.add(n1.name); keys.add(n1.name + '|' + n2.name); }
      });
    });
    setExpanded(keys);
  }, [st.q, tree]);

  const onToggle = useCallback(key => {
    setExpanded(prev => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });
  }, []);
  // Clic en el nombre de una fila (VP / Gerencia / Ítem) → ese elemento pasa a
  // ser filtro GLOBAL del dashboard (KPIs, gráficos y tabla). Vuelve a hacer
  // clic en la misma fila para quitar el filtro. Al fijar un nivel se limpian
  // los niveles inferiores para enfocar exactamente ese elemento.
  const onRowFilter = useCallback((row, dims) => {
    const FK = { vp: 'vps', ger: 'gers', item: 'items' };
    const HIER = ['vp', 'ger', 'item'];
    const dim = dims[row.level - 1];
    const fkey = FK[dim];
    const val = row.node.name;
    setSt(s => {
      const cur = s[fkey] || [];
      if (cur.length === 1 && cur[0] === val) return { ...s, [fkey]: [] }; // toggle off
      const patch = { [fkey]: [val] };
      HIER.slice(HIER.indexOf(dim) + 1).forEach(d => { patch[FK[d]] = []; });
      return { ...s, ...patch };
    });
  }, []);
  // Limpiar todos los filtros de categoría (VP / Gerencia / Ítem) a la vez.
  const onClearFilter = useCallback(() => setSt(s => ({ ...s, vps: [], gers: [], items: [] })), []);
  const expandAll = () => {
    const n = new Set();
    tree.vpNodes.forEach(vp => { n.add(vp.name); vp.children.forEach(g => n.add(vp.name + '|' + g.name)); });
    setExpanded(n);
  };
  const collapseAll = () => setExpanded(new Set());

  const onSort = useCallback(key => {
    setSt(s => s.sort.key === key
      ? { ...s, sort: { key, dir: s.sort.dir === 'asc' ? 'desc' : 'asc' } }
      : { ...s, sort: { key, dir: key === 'name' ? 'asc' : 'desc' } });
  }, []);
  // Menú de columnas: orden explícito, fijar/liberar y ocultar.
  const onSortDir = useCallback((key, dir) => setSt(s => ({ ...s, sort: { key, dir } })), []);
  const onPinCol = useCallback(key => setPinnedCols(p => p.includes(key) ? p.filter(x => x !== key) : [...p, key]), []);
  const onHideCol = useCallback(key => setHiddenCols(h => h.includes(key) ? h : [...h, key]), []);
  const resetCols = () => { setHiddenCols([]); setPinnedCols([]); };

  const onEditProp = useCallback((node, newTotalUSD) => {
    // Distribuye el nuevo total entre los recIds del ítem. Si la base prop27 es 0
    // (caso actual: propuesta arranca en 0), reparte en partes iguales — si no,
    // proporcional a la base existente.
    setOverrides(prev => {
      const n = { ...prev };
      const ids = node.recIds || [];
      if (!ids.length) return n;
      const base = ids.map(id => A.derived(A.recordById(id), growth).prop27);
      const baseSum = base.reduce((a, b) => a + b, 0);
      ids.forEach((id, i) => {
        n[id] = baseSum > 0 ? newTotalUSD * (base[i] / baseSum) : newTotalUSD / ids.length;
      });
      return n;
    });
  }, [growth]);
  const resetProp = () => setOverrides({});

  /* ---------- KPI cards ---------- */
  const unit = t.unit;
  const dec = t.decimals != null ? t.decimals : 1;
  const kpis = useMemo(() => {
    const T = tree.total;
    // Con más de un año seleccionado, las tarjetas históricas muestran el PROMEDIO
    // anual (suma / nº años), no el acumulado, y lo dicen explícitamente.
    const nY = tree.nYears || 1;
    const avg = nY > 1;
    const realShown = avg ? T.ytdReal / nY : T.ytdReal;
    const verShown = avg ? T.ytdVersion / nY : T.ytdVersion;
    const difShown = realShown - verShown;
    const pct = verShown ? difShown / Math.abs(verShown) : 0; // mismo ratio (avg o suma)
    const cump = T.ytdVersion ? T.ytdReal / T.ytdVersion : 0; // ratio: no cambia
    const alarmaVPs = tree.vpNodes.filter(v => A.kpiColor(v.agg.ytdReal, v.agg.ytdVersion, thr) === 'rojo').map(v => A.dispVP(v.name));
    const alarmas = alarmaVPs.length;
    const cumpCol = A.kpiColor(T.ytdReal, T.ytdVersion, thr);
    const cards = [
      { label: avg ? 'Real promedio anual' : TX.kpi.realAnual, value: A.fmt(realShown, unit, dec), unit, sub: avg ? `promedio de ${periodLbl}` : `${yearsLabel} · acumulado anual` },
      { label: avg ? 'Presupuesto promedio' : TX.kpi.presupuestoAnual, value: A.fmt(verShown, unit, dec), unit, sub: A.VERSIONES.find(v => v.id === st.version).label },
      { label: avg ? 'Desviación promedio' : TX.kpi.desviacionAnual, value: (difShown > 0 ? '+' : '') + A.fmt(difShown, unit, dec), unit, color: difShown < 0 ? 'var(--ok)' : difShown > 0 ? 'var(--red)' : 'var(--fg-soft)', trend: (pct > 0 ? '+' : '') + A.fmtPct(pct, 1), trendDir: pct > 0 ? 'up' : 'down', sub: pct > 0 ? 'sobre presupuesto' : 'bajo presupuesto' },
      { label: TX.kpi.cumplimiento, value: A.fmtPct(cump, 0), color: A.kpiHex(cumpCol), sub: 'Real / Presupuesto anual' },
      { label: TX.kpi.alarmas, value: String(alarmas), color: A.kpiHex(alarmas ? 'rojo' : 'azul'), sub: `de ${tree.vpNodes.length} VP sobre umbral`, tooltip: alarmas ? 'VP sobre presupuesto (>' + thr.red + '%): ' + alarmaVPs.join(', ') : 'Ninguna VP sobre el umbral' },
    ];
    // Capa Propuesta 2027: tarjeta adicional comparando contra el promedio de los años seleccionados.
    if (showProp) {
      const prom = tree.nYears ? T.ytdReal / tree.nYears : 0;
      const prop = T.prop || 0;
      const dP = prom ? (prop - prom) / Math.abs(prom) : 0;
      cards.push({
        label: 'Propuesta 2027', value: A.fmt(prop, unit, dec), unit, status: 'amarillo', dot: 'amarillo',
        trend: (dP > 0 ? '+' : '') + A.fmtPct(dP, 1), trendDir: dP > 0 ? 'up' : 'down',
        sub: `vs promedio ${periodLbl}`,
      });
    }
    return cards;
  }, [tree, showProp, unit, dec, thr, st.version, yearsLabel]);

  /* ---------- export CSV ---------- */
  const exportExcel = () => {
    const div = unit === 'kUSD' ? 1e3 : 1e6;
    const f = Math.pow(10, dec);
    const num = v => ({ t: 'n', v: Math.round((v / div) * f) / f });
    const DLBL = { vp: 'Vicepresidencia', ger: 'Gerencia', item: '\u00CDtem Relevante' };
    const DISP = { vp: A.dispVP, ger: A.dispGer, item: A.dispItem };
    const cols = groupDims.map(d => DLBL[d]);
    const names = (n1, n2, n3) => [DISP[groupDims[0]](n1.name), DISP[groupDims[1]](n2.name), DISP[groupDims[2]](n3.name)].map(s => ({ t: 's', v: s }));
    const nY = tree.nYears || 0;
    const byYear = !!tree.byYear;
    const yrs = tree.years || [];
    // Acumulado o promedio según el selector de la tabla (solo con >1 año).
    const fAgg = (st.yearAgg === 'avg' && nY > 1) ? 1 / nY : 1;
    let headers;
    if (byYear) {
      headers = [...cols];
      yrs.forEach(y => headers.push(`Real ${y} (${unit})`, `Ppto ${y} (${unit})`));
    } else {
      const realLbl = fAgg !== 1 ? `Real promedio (${unit})` : `Real (${unit})`;
      const pptoLbl = fAgg !== 1 ? `Ppto promedio (${unit})` : `Ppto Original (${unit})`;
      headers = [...cols, realLbl, pptoLbl, `Dif (${unit})`, '% Dif'];
    }
    if (showProp) headers = [...headers, `Propuesta 2027 (${unit})`, `Δ vs Real %`];
    const rows = [];
    tree.vpNodes.forEach(n1 => n1.children.forEach(n2 => n2.children.forEach(n3 => {
      const a = n3.agg, difu = a.ytdReal - a.ytdVersion;
      const pct = a.ytdVersion ? Math.round((difu / Math.abs(a.ytdVersion)) * 100) : 0;
      let row;
      if (byYear) {
        row = [...names(n1, n2, n3)];
        yrs.forEach(y => { const yv = (a.yr && a.yr[y]) || { real: 0, ver: 0 }; row.push(num(yv.real), num(yv.ver)); });
      } else {
        row = [...names(n1, n2, n3), num(a.ytdReal * fAgg), num(a.ytdVersion * fAgg), num(difu * fAgg), { t: 'n', v: pct }];
      }
      if (showProp) {
        const prom = nY ? a.ytdReal / nY : 0;
        const dpct = prom ? Math.round((((a.prop || 0) - prom) / Math.abs(prom)) * 100) : 0;
        row.push(num(a.prop || 0), { t: 'n', v: dpct });
      }
      rows.push(row);
    })));
    const safe = ('Datos ' + yearsLabel).replace(/[\\\/?*\[\]:]/g, ' ').slice(0, 31);
    const url = URL.createObjectURL(_xlsx(safe, headers, rows));
    const a = document.createElement('a');
    a.href = url; a.download = `ActividadCorporativa_${yearsLabel.replace(/[^0-9A-Za-z]+/g, '_')}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  };

  const dens = t.density;
  const rowPad = dens === 'compact' ? 4 : dens === 'comfy' ? 11 : 7;

  return (
    <div id="app" style={{ '--rowpad': rowPad + 'px' }}>
      {/* Header */}
      <div className="hdr">
        <svg className="hdr-bg" viewBox="0 0 1480 90" preserveAspectRatio="none">
          <path d="M0 90 L120 10 L210 90 Z" fill={A.color('var(--amsa-teal)')} opacity="0.06" />
          <path d="M150 90 L255 25 L330 90 Z" fill={A.color('var(--amsa-yellow)')} opacity="0.07" />
          <path d="M1480 0 L1480 90 L1380 90 Z" fill={A.color('var(--amsa-teal)')} opacity="0.05" />
        </svg>
        <div className="hdr-row">
          <div className="hdr-titles">
            <h1>
              <span style={{ color: corpOn ? 'var(--amsa-teal)' : 'var(--amsa-teal-light)', transition: 'color .15s' }}>{TX.header.titulo}</span>
              <span className="plus" style={{ color: distOn ? 'var(--amsa-teal)' : 'var(--amsa-teal-light)', transition: 'color .15s' }}> {TX.header.tituloPlus}</span>
            </h1>
            <div className="sub">{TX.header.subtitulo}</div>
          </div>
          <img className="hdr-logo" src={window.ADA_LOGO || 'assets/logo_amsa.png'} alt="Antofagasta Minerals" />
        </div>
      </div>
      <div className="accent"><i className="t"></i><i className="r"></i><i className="y"></i></div>

      <div className="app-wrap">
        <div style={{ display: 'flex', gap: 4, marginBottom: 12, borderBottom: '2px solid var(--teal-100)' }}>
          {[['dashboard', 'Dashboard'], ['dict', 'Diccionario (CECO · Ítem)']].map(([id, lbl]) => (
            <button key={id} type="button" onClick={() => setPage(id)} style={{
              border: 0, background: 'transparent', cursor: 'pointer', padding: '8px 16px',
              fontFamily: 'var(--font-disp)', fontWeight: 700, fontSize: 13, marginBottom: -2,
              color: page === id ? 'var(--amsa-teal)' : 'var(--fg-soft)',
              borderBottom: page === id ? '2px solid var(--amsa-teal)' : '2px solid transparent',
            }}>{lbl}</button>
          ))}
          <button onClick={() => setEditMode(e => !e)} title="Activar/desactivar modo edición" style={{
            marginLeft: 'auto', alignSelf: 'center', display: 'flex', alignItems: 'center', gap: 6,
            height: 28, padding: '0 12px', borderRadius: 16,
            border: editMode ? '0' : '1px solid var(--teal-border)',
            background: editMode ? 'var(--amsa-teal-deep)' : 'rgba(255,255,255,.9)',
            color: editMode ? '#fff' : 'var(--fg-soft)', fontWeight: 600, fontSize: 11.5,
            fontFamily: 'var(--font-sans)', cursor: 'pointer',
            boxShadow: editMode ? '0 4px 14px rgba(20,40,50,.22)' : '0 1px 5px rgba(20,40,50,.12)',
            opacity: editMode ? 1 : 0.72,
          }}>
            <span style={{ fontSize: 13 }}>✏️</span> Modo Edición
          </button>
        </div>
        {page === 'dict' ? <DictView vpov={vpov} setVpov={setVpov} nameov={nameov} setNameov={setNameov} /> : <React.Fragment>
        <FilterBar st={st} set={set} gerOptions={dims.gers} itemOptions={dims.items} />

        <KpiCards kpis={kpis} unit={unit} />

        {t.showCharts && (() => {
          const barCard = (
            <React.Fragment>
              <h3>Real vs Presupuesto por año</h3>
              <div className="ph-sub">{showProp ? 'Histórico anual y Propuesta 2027' : 'Histórico anual'} · {unit}</div>
              <BarYears series={series} unit={unit} decimals={dec} showProp={showProp} />
              <div className="legend">
                <span><i style={{ background: 'var(--amsa-teal-deep)' }}></i>Real histórico</span>
                <span><i style={{ background: 'var(--amsa-teal-light)' }}></i>Presupuesto</span>
                {showProp && <span><i style={{ background: 'var(--amsa-yellow)' }}></i>Propuesta 27</span>}
              </div>
            </React.Fragment>
          );
          const compCard = (
            <React.Fragment>
              <h3>Cumplimiento del presupuesto</h3>
              <div className="ph-sub">Real / Presupuesto por año</div>
              <ComplianceBars series={series} thr={thr} unit={unit} />
            </React.Fragment>
          );
          // Sin donut: dos tarjetas separadas que llenan toda la fila con la
          // proporción 1.5:1. Las gráficas son de alto fijo (responsivas), así
          // que ensancharse para llenar el ancho no cambia el alto.
          if (!distOn) return (
            <div className="grid3" style={{ gridTemplateColumns: '1.5fr 1fr' }}>
              <div className="panel">{barCard}</div>
              <div className="panel">{compCard}</div>
            </div>
          );
          return (
            <div className="grid3">
              <div className="panel">{barCard}</div>
              <div className="panel">{compCard}</div>
              <div className="panel">
                <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 8 }}>
                  <div>
                    <h3>Distribuible por compañía</h3>
                    <div className="ph-sub">{st.donutMetric === 'prop' ? 'Propuesta 2027' : 'Real'} por compañía · {unit}</div>
                  </div>
                  {showProp && (
                    <div style={{ display: 'flex', gap: 2, background: '#fff', overflow: 'hidden',
                      border: '1px solid var(--teal-200)', borderRadius: 4, flex: '0 0 auto', marginTop: 2 }}>
                      {[['real', 'Real'], ['prop', 'Propuesta 27']].map(([m, lbl]) => (
                        <button key={m} type="button" onClick={() => set({ donutMetric: m })}
                          style={{ border: 0, padding: '4px 9px', cursor: 'pointer', fontSize: 11, fontWeight: 600,
                            fontFamily: 'var(--font-sans)',
                            background: st.donutMetric === m ? 'var(--amsa-teal)' : '#fff',
                            color: st.donutMetric === m ? '#fff' : 'var(--fg-soft)' }}>{lbl}</button>
                      ))}
                    </div>
                  )}
                </div>
                <Donut dist={dist} unit={unit} decimals={dec} />
              </div>
            </div>
          );
        })()}

        {/* Matrix */}
        <div className="matrix-card">
          <div className="matrix-top">
            <div>
              <h3>Ejecución {yearsLabel}</h3>
              <div className="mt-sub">{modeLbl}</div>
            </div>
            <div className="toolbar">
              <span style={{ position: 'relative', display: 'inline-flex', alignItems: 'center', marginRight: 6 }}>
                <input type="text" value={st.q} onChange={e => set({ q: e.target.value })}
                  placeholder="Buscar VP / Gerencia / Ítem…"
                  style={{ height: 30, width: 210, padding: '0 26px 0 10px', boxSizing: 'border-box',
                    border: '1px solid var(--teal-200)', borderRadius: 6, fontSize: 12, fontFamily: 'var(--font-sans)',
                    color: 'var(--ink)', outline: 'none' }} />
                {st.q && <button type="button" onClick={() => set({ q: '' })} aria-label="Limpiar búsqueda"
                  style={{ position: 'absolute', right: 6, border: 0, background: 'transparent', cursor: 'pointer',
                    color: 'var(--fg-soft)', fontSize: 14, lineHeight: 1, padding: 2 }}>×</button>}
              </span>
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginRight: 6, fontSize: 11.5, color: 'var(--fg-3)', fontWeight: 600 }}>
                Agrupar
                <select value={st.groupMode} onChange={e => { set({ groupMode: e.target.value }); setExpanded(new Set()); }}
                  style={{ height: 30, padding: '0 8px', border: '1px solid var(--teal-200)', borderRadius: 6, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', cursor: 'pointer' }}>
                  <option value="org">VP › Gerencia › Ítem</option>
                  <option value="item">Ítem › VP › Gerencia</option>
                </select>
              </span>
              {multiYear && (
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginRight: 6, fontSize: 11.5, color: 'var(--fg-3)', fontWeight: 600 }}>
                  Años
                  <select value={st.yearAgg} onChange={e => set({ yearAgg: e.target.value })}
                    style={{ height: 30, padding: '0 8px', border: '1px solid var(--teal-200)', borderRadius: 6, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', cursor: 'pointer' }}>
                    <option value="sum">Acumulado</option>
                    <option value="avg">Promedio</option>
                    <option value="byYear">Por año</option>
                  </select>
                </span>
              )}
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 7, marginRight: 6, fontSize: 11.5, color: 'var(--fg-3)', fontWeight: 600 }}>
                Decimales
                <span style={{ display: 'inline-flex', alignItems: 'stretch', height: 26, border: '1px solid var(--teal-200)', borderRadius: 6, overflow: 'hidden', background: '#fff' }}>
                  <span style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', minWidth: 24, fontWeight: 700, color: 'var(--ink)', fontVariantNumeric: 'tabular-nums' }}>{dec}</span>
                  <span style={{ display: 'flex', flexDirection: 'column', borderLeft: '1px solid var(--teal-100)' }}>
                    <button type="button" aria-label="Más decimales" onClick={() => setTweak('decimals', Math.min(3, dec + 1))}
                      style={{ border: 0, borderBottom: '1px solid var(--teal-100)', background: 'var(--teal-wash2)', color: 'var(--teal-muted)', cursor: 'pointer', padding: '0 7px', fontSize: 7, lineHeight: '12px', flex: 1 }}>▲</button>
                    <button type="button" aria-label="Menos decimales" onClick={() => setTweak('decimals', Math.max(0, dec - 1))}
                      style={{ border: 0, background: 'var(--teal-wash2)', color: 'var(--teal-muted)', cursor: 'pointer', padding: '0 7px', fontSize: 7, lineHeight: '12px', flex: 1 }}>▼</button>
                  </span>
                </span>
              </span>
              <button type="button" onClick={() => expanded.size ? collapseAll() : expandAll()}
                style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: 'var(--teal-wash)', color: 'var(--amsa-teal)', border: '1px solid var(--amsa-teal-light)', borderRadius: 7, padding: '6px 13px', fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: 'var(--font-sans)', whiteSpace: 'nowrap', marginRight: 6 }}>
                {expanded.size ? '⤒ Colapsar todo' : '⤓ Expandir todo'}
              </button>
              {(hiddenCols.length > 0 || pinnedCols.length > 0) && <button className="btn ghost" onClick={resetCols}>Restablecer columnas</button>}
              {showProp && Object.keys(overrides).length > 0 && <button className="btn ghost" onClick={resetProp}>Restablecer propuesta</button>}
              <button className="btn" style={{ whiteSpace: 'nowrap' }} onClick={exportExcel}>⭳ Exportar Excel</button>
            </div>
          </div>
          <div style={{ overflowX: 'auto' }}>
            <Matrix tree={tree} showProp={showProp} nYears={tree.nYears} yearAgg={st.yearAgg} yearsLabel={yearsLabel} unit={unit} thr={thr} decimals={dec}
              versionLabel={A.VERSIONES.find(v => v.id === st.version).label} q={st.q}
              sort={st.sort} onSort={onSort} onSortDir={onSortDir}
              hiddenCols={hiddenCols} pinnedCols={pinnedCols} onHide={onHideCol} onPin={onPinCol}
              expanded={expanded} onToggle={onToggle} onEditProp={onEditProp} companiesActive={companiesActive}
              onRowFilter={onRowFilter} filterSel={{ vps: st.vps, gers: st.gers, items: st.items }} onClearFilter={onClearFilter} />
          </div>
        </div>

        <div className="note">
          <b>KPI:</b> verde ≤ presupuesto · amarillo hasta {thr.red}% · rojo &gt; {thr.red}% sobre presupuesto · montos en {unit}{showProp ? (companiesActive ? ' · Propuesta 2027 no editable con filtro de compañía' : ' · Propuesta 2027 editable a nivel de ítem') : ''}.<br />
          <b>Fuente:</b> Consulta a SAP BPC de valores históricos 2022-2025. <b>Propuesta 2027:</b> Pendiente.
        </div>
        </React.Fragment>}
      </div>

      <TweaksPanel>
        <TweakSection label="Presentación" />
        <TweakRadio label="Unidad" value={t.unit} options={['MUSD', 'kUSD']} onChange={v => setTweak('unit', v)} />
        <TweakRadio label="Densidad" value={t.density} options={['compact', 'regular', 'comfy']} onChange={v => setTweak('density', v)} />
        <TweakToggle label="Mostrar gráficos" value={t.showCharts} onChange={v => setTweak('showCharts', v)} />
        <TweakSection label="Regla KPI (% sobre presupuesto)" />
        <TweakSlider label="Umbral rojo" value={t.thrRed} min={3} max={30} step={1} unit="%" onChange={v => setTweak('thrRed', v)} />
        <TweakSlider label="Umbral amarillo" value={t.thrYellow} min={1} max={t.thrRed - 1} step={1} unit="%" onChange={v => setTweak('thrYellow', v)} />
      </TweaksPanel>

      <ColorPanel onApply={onColorApply} edit={editMode} />
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
