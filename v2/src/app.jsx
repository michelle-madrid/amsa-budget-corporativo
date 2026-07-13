/* global React, ReactDOM, CORP, FilterBar, ViewTabs, KpiCards, Matrix, BarYears, ComplianceBars, Donut,
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
  const A = window.CORP;
  const { vpov, setVpov, nameov, setNameov } = props; // estado elevado al App: al editar, el Dashboard se reagrupa solo
  const cecoNames = props.cecoNames || { ger: {}, dceco: {} };   // renombres POR CECO (Gerencia y Desc. CECO)
  const setCecoNames = props.setCecoNames || (() => {});
  const cecoMode = props.cecoMode || 'new';            // Estructura CECOS (Nueva/Antigua) — el Diccionario refleja el mapa activo
  const onCecoMode = props.onCecoMode || (() => {});
  const [q, setQ] = React.useState('');
  const [editCeco, setEditCeco] = React.useState(null);
  const [editGer, setEditGer] = React.useState(null);   // ceco de la fila cuya Gerencia se edita
  const [editDceco, setEditDceco] = React.useState(null); // ceco cuya Desc. CECO se edita
  const [editItem, setEditItem] = React.useState(null); // clave del Ítem que se edita
  const [openCats, setOpenCats] = React.useState(() => new Set()); // divisiones expandidas (acordeón)
  const [catSel, setCatSel] = React.useState([]); // filtro por Clasificación del Gasto ([] = todas)
  const ql = q.trim().toLowerCase();
  // Renombrar Gerencia y Desc. CECO: override POR CECO (afecta solo a ese CECO). Si el valor
  // queda vacío o igual al original, se quita el override.
  const setGerName = (ceco, base, val) => setCecoNames(o => {
    const ger = { ...(o.ger || {}) }; const v = (val || '').trim();
    if (!v || v === base) delete ger[ceco]; else ger[ceco] = v;
    return { ...o, ger };
  });
  const setDcecoName = (ceco, base, val) => setCecoNames(o => {
    const dceco = { ...(o.dceco || {}) }; const v = (val || '').trim();
    if (!v || v === base) delete dceco[ceco]; else dceco[ceco] = v;
    return { ...o, dceco };
  });
  // Renombrar Ítem Relevante: override compartido por clave de Ítem (afecta a todos los CECO
  // con ese Ítem, que es el comportamiento deseado para los ítems).
  const setItemName = (itemKey, val) => setNameov(o => {
    const item = { ...(o.item || {}) }; const v = (val || '').trim();
    if (!v || v === A.baseItem(itemKey)) delete item[itemKey]; else item[itemKey] = v;
    return { ...o, item };
  });
  const cecos = (window.CORP_DICT && window.CORP_DICT.cecos) || [];
  const vpKeys = Array.from(new Set(cecos.map(r => r.v))).sort((a, b) => A.dispVP(a).localeCompare(A.dispVP(b), 'es'));
  const setVp = (ceco, vpKey, origV) => setVpov(o => { const n = { ...o }; if (!vpKey || vpKey === origV) delete n[ceco]; else n[ceco] = vpKey; return n; });
  const gerCeco = cecoNames.ger || {}, dcecoCeco = cecoNames.dceco || {};   // overrides POR CECO
  // Alias de Gerencia: clave canónica → [nombres crudos como venían en los archivos].
  // Se muestran como sub-filas bajo el CECO cuya Gerencia es esa clave canónica.
  const aliasByCanon = {};
  Object.keys(A.GER_ALIAS || {}).forEach(raw => {
    const c = A.GER_ALIAS[raw].ger;
    (aliasByCanon[c] = aliasByCanon[c] || []).push(raw);
  });
  // División = "Clasificación del Gasto" (campo de CECOS.xlsx, viene en CORP_DICT.cl).
  // Fallback por código si faltara: 1000… = Corporativo, resto = Distribuible.
  const cecoCat = r => r.cl || (String(r.c || '').startsWith('1000') ? 'Gastos Centro Corporativo' : 'Gastos Distribuibles');
  const allRows = cecos.map(r => {
    const vpKey = vpov[r.c] || r.v;
    const gerBase = A.dispGer(r.g), dcecoBase = r.d || r.c;
    return { ceco: r.c, cat: cecoCat(r), gerKey: r.g,
      ger: gerCeco[r.c] || gerBase, gerBase, gerOverridden: !!gerCeco[r.c],
      dceco: dcecoCeco[r.c] || dcecoBase, dcecoBase, dcecoOverridden: !!dcecoCeco[r.c],
      vpKey, vp: A.dispVP(vpKey), origV: r.v, overridden: !!vpov[r.c],
      aliases: aliasByCanon[r.g] || null, tc: r.tc || '—', ap: r.ap || '—' };
  });
  // Opciones del filtro por Clasificación del Gasto (todas las divisiones, con su conteo).
  const catCountAll = {}; allRows.forEach(r => { catCountAll[r.cat] = (catCountAll[r.cat] || 0) + 1; });
  const catOptions = Array.from(new Set(allRows.map(r => r.cat)))
    .sort((a, b) => catCountAll[b] - catCountAll[a] || a.localeCompare(b, 'es'))
    .map(c => ({ value: c, label: `${c} (${catCountAll[c]})` }));
  const catSet = catSel.length ? new Set(catSel) : null;
  const cecoRows = allRows.filter(r => (!catSet || catSet.has(r.cat)) &&
    (!ql || r.ceco.toLowerCase().includes(ql) || r.dceco.toLowerCase().includes(ql) || r.ger.toLowerCase().includes(ql) || r.vp.toLowerCase().includes(ql) || r.cat.toLowerCase().includes(ql) || (r.aliases && r.aliases.some(n => n.toLowerCase().includes(ql)))));
  // Agrupa por Clasificación del Gasto; divisiones ordenadas por cantidad de CECOs (desc).
  const catCount = {}; cecoRows.forEach(r => { catCount[r.cat] = (catCount[r.cat] || 0) + 1; });
  const cats = Array.from(new Set(cecoRows.map(r => r.cat))).sort((a, b) => catCount[b] - catCount[a] || a.localeCompare(b, 'es'));
  // Acordeón: colapsado por defecto. Al buscar o filtrar por categoría, se abre todo
  // para que se vean los resultados. Botones Expandir/Colapsar todo controlan openCats.
  const effOpen = (ql || catSel.length) ? new Set(cats) : openCats;
  const toggleCat = c => setOpenCats(s => { const n = new Set(s); n.has(c) ? n.delete(c) : n.add(c); return n; });
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
    .map(k => ({ key: k, name: A.dispItem(k), code: A.itemCode(k), overridden: !!itemOv[k] }))
    .filter(r => !ql || r.name.toLowerCase().includes(ql) || r.key.toLowerCase().includes(ql) || (r.code && String(r.code).toLowerCase().includes(ql)))
    .sort((a, b) => a.name.localeCompare(b.name, 'es'));
  const nItemOver = Object.keys(itemOv).length;
  const nGerOver = Object.keys(gerCeco).length;
  const nDcecoOver = Object.keys(dcecoCeco).length;
  const th = { textAlign: 'left', padding: '8px 12px', fontSize: 11, fontWeight: 700, letterSpacing: '.03em', color: '#fff', background: 'var(--amsa-teal-deep)', position: 'sticky', top: 0, zIndex: 1 };
  const td = { padding: '6px 12px', fontSize: 12.5, borderBottom: '1px solid var(--line-soft)', color: 'var(--ink)' };
  const foot = { padding: '7px 12px', fontSize: 11, color: 'var(--fg-muted)', borderTop: '1px solid var(--line-soft)' };
  const iconBtn = { border: 0, background: 'transparent', cursor: 'pointer', color: 'var(--fg-muted)', fontSize: 12, padding: '0 2px', lineHeight: 1 };
  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 12, margin: '4px 0 14px' }}>
        <div>
          <h2 style={{ fontFamily: 'var(--font-disp)', fontWeight: 800, fontSize: 18, color: 'var(--ink)', margin: 0 }}>Diccionario de códigos</h2>
          <div style={{ fontSize: 12, color: 'var(--fg-muted)', marginTop: 2 }}>CECO · Desc. CECO y Gerencia (renombrables por CECO) · Vicepresidencia (editable) · agrupado por Clasificación del Gasto · Ítem Relevante (renombrable)</div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <label style={{ fontSize: 11.5, color: 'var(--fg-muted)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
            Estructura CECOS
            <select value={cecoMode} onChange={e => onCecoMode(e.target.value)}
              style={{ height: 32, border: '1px solid var(--teal-border)', borderRadius: 6, fontSize: 12.5, padding: '0 8px', color: 'var(--ink)', background: '#fff', fontFamily: 'var(--font-sans)', outline: 'none' }}>
              <option value="new">Nuevos</option>
              <option value="old">Antiguos</option>
            </select>
          </label>
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Buscar CECO, Gerencia, VP o Ítem…"
            style={{ height: 32, width: 260, padding: '0 12px', border: '1px solid var(--teal-border)', borderRadius: 6, fontSize: 12.5, outline: 'none', fontFamily: 'var(--font-sans)', color: 'var(--ink)' }} />
        </div>
      </div>
      {/* Filtro por Clasificación del Gasto: chips toggle (on-brand), con conteo. */}
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, margin: '0 0 12px' }}>
        {[{ k: null, label: 'Todas', n: allRows.length }].concat(catOptions.map(o => ({ k: o.value, label: o.value, n: catCountAll[o.value] }))).map(c => {
          const active = c.k === null ? catSel.length === 0 : catSel.includes(c.k);
          return (
            <button key={c.k || '__all'} type="button"
              onClick={() => c.k === null ? setCatSel([]) : setCatSel(s => s.includes(c.k) ? s.filter(x => x !== c.k) : s.concat(c.k))}
              style={{ border: active ? '1px solid var(--amsa-teal)' : '1px solid var(--teal-border)', background: active ? 'var(--amsa-teal)' : '#fff', color: active ? '#fff' : 'var(--fg-soft)', borderRadius: 14, padding: '4px 10px', fontSize: 11.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'var(--font-sans)', display: 'inline-flex', alignItems: 'center', gap: 6, lineHeight: 1.3 }}>
              {c.label}
              <span style={{ opacity: .7, fontVariantNumeric: 'tabular-nums', fontWeight: 700 }}>{c.n}</span>
            </button>
          );
        })}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 16, alignItems: 'start' }}>
        <div className="panel" style={{ padding: 0, overflow: 'hidden' }}>
          <div style={{ maxHeight: 580, overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr><th style={{ ...th, width: 176 }}>CECO · Desc. CECO</th><th style={th}>Gerencia</th><th style={th}>Vicepresidencia</th><th style={{ ...th, width: 96 }}>Tipo Costo</th><th style={{ ...th, width: 74 }}>¿Aplica?</th></tr></thead>
              <tbody>
                {cats.map(cat => {
                  const open = effOpen.has(cat);
                  const rowsOfCat = cecoRows.filter(r => r.cat === cat);
                  return (
                  <React.Fragment key={cat}>
                    <tr onClick={() => toggleCat(cat)} style={{ cursor: 'pointer' }}>
                      <td colSpan="5" style={{ padding: '7px 12px', fontSize: 10.5, fontWeight: 800, letterSpacing: '.06em', textTransform: 'uppercase', color: 'var(--amsa-teal-deep)', background: 'var(--teal-wash2)', borderTop: '1px solid var(--line-soft)' }}>
                        <span style={{ display: 'inline-block', width: 14, color: 'var(--fg-muted)' }}>{open ? '▾' : '▸'}</span>
                        {cat} <span style={{ fontWeight: 600, opacity: .7 }}>· {catCount[cat]} CECO{catCount[cat] > 1 ? 's' : ''}</span>
                      </td>
                    </tr>
                    {open && rowsOfCat.map(r => (
                      <React.Fragment key={r.ceco}>
                      <tr style={r.overridden ? { background: 'var(--accent-wash)' } : undefined}>
                        <td style={td}>
                          <div style={{ fontVariantNumeric: 'tabular-nums', fontWeight: 600 }}>{r.ceco}</div>
                          {editDceco === r.ceco
                            ? <input autoFocus type="text" defaultValue={r.dceco}
                                onBlur={e => { setDcecoName(r.ceco, r.dcecoBase, e.target.value); setEditDceco(null); }}
                                onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); if (e.key === 'Escape') setEditDceco(null); }}
                                style={{ width: '100%', height: 24, marginTop: 3, border: '1px solid var(--amsa-teal)', borderRadius: 5, fontSize: 11.5, fontFamily: 'var(--font-sans)', color: 'var(--ink)', background: '#fff', padding: '0 6px', boxSizing: 'border-box' }} />
                            : <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, marginTop: 1, fontSize: 11, color: 'var(--fg-muted)' }}>
                                <span style={r.dcecoOverridden ? { fontWeight: 700, color: 'var(--accent-ink)' } : undefined}>{r.dceco}</span>
                                {r.dcecoOverridden && <span title={'Original: ' + r.dcecoBase} style={{ fontSize: 9, color: 'var(--amsa-yellow)' }}>●</span>}
                                <button title="Renombrar Desc. CECO (solo este CECO)" onClick={() => setEditDceco(r.ceco)} style={iconBtn}>✎</button>
                                {r.dcecoOverridden && <button title="Restablecer nombre" onClick={() => setDcecoName(r.ceco, r.dcecoBase, null)} style={iconBtn}>↺</button>}
                              </span>}
                        </td>
                        <td style={td}>
                          {editGer === r.ceco
                            ? <input autoFocus type="text" defaultValue={r.ger}
                                onBlur={e => { setGerName(r.ceco, r.gerBase, e.target.value); setEditGer(null); }}
                                onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); if (e.key === 'Escape') setEditGer(null); }}
                                style={{ width: '100%', height: 26, border: '1px solid var(--amsa-teal)', borderRadius: 5, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', background: '#fff', padding: '0 6px', boxSizing: 'border-box' }} />
                            : <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                                <span style={r.gerOverridden ? { fontWeight: 700, color: 'var(--accent-ink)' } : undefined}>{r.ger}</span>
                                {r.gerOverridden && <span title={'Original: ' + r.gerBase} style={{ fontSize: 9, color: 'var(--amsa-yellow)' }}>●</span>}
                                <button title="Renombrar Gerencia (solo este CECO)" onClick={() => setEditGer(r.ceco)} style={iconBtn}>✎</button>
                                {r.gerOverridden && <button title="Restablecer nombre" onClick={() => setGerName(r.ceco, r.gerBase, null)} style={iconBtn}>↺</button>}
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
                        <td style={{ ...td, fontVariantNumeric: 'tabular-nums', color: 'var(--fg-2)' }}>{r.tc}</td>
                        <td style={{ ...td, fontWeight: r.ap === 'No' ? 700 : undefined, color: r.ap === 'No' ? 'var(--red)' : 'var(--fg-2)' }}>{r.ap}</td>
                      </tr>
                      {/* Sub-filas: nombres crudos con los que esta Gerencia venía en los archivos (alias unificados). */}
                      {r.aliases && r.aliases.map(raw => (
                        <tr key={r.ceco + '|' + raw} style={{ background: 'var(--teal-wash2)' }}>
                          <td style={{ ...td, borderBottom: '1px solid var(--line-soft)' }}></td>
                          <td style={{ ...td, paddingLeft: 26, color: 'var(--fg-muted)', fontStyle: 'italic' }}>
                            <span title="Nombre con el que venía en los archivos (unificado bajo el nombre oficial de arriba)">↳ {raw}</span>
                          </td>
                          <td style={td}></td>
                          <td style={td}></td>
                          <td style={td}></td>
                        </tr>
                      ))}
                      </React.Fragment>
                    ))}
                  </React.Fragment>
                  );
                })}
                {cecoRows.length === 0 && <tr><td style={td} colSpan="5">Sin resultados</td></tr>}
              </tbody>
            </table>
          </div>
          <div style={{ ...foot, display: 'flex', justifyContent: 'space-between' }}>
            <span>{cecoRows.length} CECO{nOver > 0 ? ` · ${nOver} con VP reasignada` : ''}{nGerOver > 0 ? ` · ${nGerOver} Gerencia renombrada${nGerOver > 1 ? 's' : ''}` : ''}{nDcecoOver > 0 ? ` · ${nDcecoOver} Desc. CECO renombrada${nDcecoOver > 1 ? 's' : ''}` : ''}</span>
            {(nOver > 0 || nGerOver > 0 || nDcecoOver > 0) && <button onClick={() => { setVpov({}); setCecoNames({ ger: {}, dceco: {} }); }} style={{ ...iconBtn, color: 'var(--amsa-teal)', fontWeight: 600 }}>Restablecer todo</button>}
          </div>
        </div>
        <div className="panel" style={{ padding: 0, overflow: 'hidden' }}>
          <div style={{ maxHeight: 580, overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr><th style={th}>Ítem Relevante</th><th style={{ ...th, width: 128 }}>Código</th></tr></thead>
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
                    <td style={{ ...td, fontVariantNumeric: 'tabular-nums', fontFamily: 'ui-monospace, SFMono-Regular, Menlo, monospace', fontSize: 11, color: it.code ? 'var(--fg-2)' : 'var(--fg-muted)', whiteSpace: 'nowrap' }} title={it.code ? ('Cód_Agrupación2: ' + it.code) : 'Sin código en el diccionario CLACO'}>{it.code || '—'}</td>
                  </tr>
                ))}
                {itemRows.length === 0 && <tr><td style={td} colSpan="2">Sin resultados</td></tr>}
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

/* ---------- Pestaña Tabla Resumen: comparador configurable por período ---------- */
// Catálogo de columnas (períodos/métricas) comparables, AGRUPADO por subcategoría
// (Real → Ppto → Forecast). El orden aquí define el orden de despliegue (desplegable
// y columnas de la tabla); dentro de cada grupo van de más nuevo a más antiguo. La
// PRIMERA columna seleccionada es la "base" contra la que se calculan Dif y % Dif.
// El Forecast 5+7 (cuando se cargue) viene de window.CORP_FCST26 (por Ítem).
const RESUMEN_COL_CATALOG = [
  // ── Real ──
  { key: 'r2026', kind: 'real', y: 2026, cat: 'Real', label: 'Real 2026 YTD' },
  { key: 'r2025', kind: 'real', y: 2025, cat: 'Real', label: 'Real 2025' },
  { key: 'r2024', kind: 'real', y: 2024, cat: 'Real', label: 'Real 2024' },
  { key: 'r2023', kind: 'real', y: 2023, cat: 'Real', label: 'Real 2023' },
  { key: 'r2022', kind: 'real', y: 2022, cat: 'Real', label: 'Real 2022' },
  // ── Ppto ──
  { key: 'prop',  kind: 'prop',          cat: 'Ppto', label: 'Ppto 2027' },
  { key: 'pfy26', kind: 'planfy',        cat: 'Ppto', label: 'Ppto 2026 FY' },
  { key: 'p2026', kind: 'plan', y: 2026, cat: 'Ppto', label: 'Ppto 2026 YTD' },
  { key: 'p2025', kind: 'plan', y: 2025, cat: 'Ppto', label: 'Ppto 2025' },
  { key: 'p2024', kind: 'plan', y: 2024, cat: 'Ppto', label: 'Ppto 2024' },
  { key: 'p2023', kind: 'plan', y: 2023, cat: 'Ppto', label: 'Ppto 2023' },
  { key: 'p2022', kind: 'plan', y: 2022, cat: 'Ppto', label: 'Ppto 2022' },
  // ── Forecast ──
  { key: 'fcst26', kind: 'fcst', cat: 'Forecast', label: 'Forecast 5+7 2026' },
];
const RESUMEN_DIM_LBL = { vp: 'Vicepresidencia', ger: 'Gerencia', dceco: 'Desc. CECO', itemrel: 'Ítem Relevante', item: 'Ítem', claco: 'Desc. CLACO', ceco: 'CECO', contra: 'Contrapartida' };
// Estructuras de la Tabla Resumen. El Ítem Relevante (Agrupación3) va como padre del
// Ítem (Agrupación4). Las que terminan en CECO/Contrapartida = detalle máximo.
// Selección por defecto de Clasificación Cuenta: TODAS menos "Mano de Obra".
const defaultClases = () => (window.CORP.dimsFor({ dataMode: 'both' }).clases || []).filter(c => c !== 'Mano de Obra');
const RESUMEN_DIMS = {
  item:   ['itemrel', 'item', 'vp', 'ger'],
  org:    ['vp', 'ger', 'itemrel', 'item'],
  itemc:  ['itemrel', 'item', 'vp', 'ger', 'dceco', 'ceco'],
  orgc:   ['vp', 'ger', 'dceco', 'ceco', 'itemrel', 'item'],
  itemcc: ['itemrel', 'item', 'vp', 'ger', 'dceco', 'ceco', 'contra'],
  orgcc:  ['vp', 'ger', 'itemrel', 'item', 'dceco', 'ceco', 'contra'],
};

function ResumenView({ overrides, unit, dec, onDec, valMode, cecoMode, onValMode, onCecoMode, st, set }) {
  const A = window.CORP;
  valMode = valMode || 'n'; cecoMode = cecoMode || 'new';
  onValMode = onValMode || (() => {}); onCecoMode = onCecoMode || (() => {});
  onDec = onDec || (() => {});
  st = st || {}; set = set || (() => {});
  // Filtros COMPARTIDOS con la pestaña "Gastos Corporativos": viven en el estado del App
  // (no se reinician al cambiar de pestaña) y se mantienen sincronizados entre ambas vistas.
  const dataMode = st.dataMode || 'both';
  const setDataMode = v => set({ dataMode: v });
  const vps = st.vps || [], gers = st.gers || [], items = st.items || [], itemrels = st.itemrels || [];
  const tcs = st.tcs || [], clases = st.clases || [], aps = st.aps || [], companias = st.companias || [], stMode = st.stMode || '', cecos = st.cecos || [];
  const setVps = v => set({ vps: v });
  const setCecos = v => set({ cecos: v });
  const clacos = st.clacos || [];
  const setClacos = v => set({ clacos: v });
  // Ocultar valores puntuales de la estructura (ej. un CECO): { dim: [valores] }.
  const hidden = st.hidden || {};
  const hiddenCount = Object.keys(hidden).reduce((n, d) => n + (hidden[d] ? hidden[d].length : 0), 0);
  const hideVal = (dim, val) => set({ hidden: Object.assign({}, hidden, { [dim]: [...(hidden[dim] || []), val].filter((v, i, a) => a.indexOf(v) === i) }) });
  const unhideAll = () => set({ hidden: {} });
  const setGers = v => set({ gers: v });
  const setItems = v => set({ items: v });
  const setItemrels = v => set({ itemrels: v });
  const setTcs = v => set({ tcs: v });
  const setClases = v => set({ clases: v });
  const setAps = v => set({ aps: v });
  const setCompanias = v => set({ companias: v });
  const setStMode = v => set({ stMode: v });
  const [distPopOpen, setDistPopOpen] = React.useState(false);  // popover de compañías anclado a "Distribuible"
  const [draftComp, setDraftComp] = React.useState([]);         // selección en borrador dentro del popover
  const [groupMode, setGroupMode] = React.useState('orgcc'); // estructuras con detalle hasta Contrapartida
  // Estructura por defecto (pedida): VP › Desc. CECO › CECO › Ítem Relevante › Ítem › Desc. CLACO
  // (Gerencia y Contrapartida quedan disponibles como chips "+"). Es un orden custom (no coincide
  // con un preset), así que arranca fijado en dimOrder; el selector de Estructura lo resetea.
  const [dimOrder, setDimOrder] = React.useState(['vp', 'ger', 'dceco', 'ceco', 'itemrel', 'item', 'claco', 'contra']);
  const [hiddenDims, setHiddenDims] = React.useState(() => new Set(['ger', 'contra'])); // niveles excluidos (× ) — Gerencia y Contrapartida ocultos por defecto (se agregan con "+")
  const [dragDim, setDragDim] = React.useState(null);        // nivel que se está arrastrando
  const [detOrder, setDetOrder] = React.useState('td');      // detalle: 'td' Texto pedido›Denominación · 'dt' al revés
  const [showDetP, setShowDetP] = React.useState(true);      // mostrar el detalle Ppto/Fcst (Concepto Gasto › Actividad)
  const [showCeco, setShowCeco] = React.useState(true);      // permite quitar el CECO del desglose
  const [nameW, setNameW] = React.useState(null);            // ancho de la columna de nombres: null = automático (llena la pantalla, se achica al comparar más); px = fijado a mano
  const [selCols, setSelCols] = React.useState(['pfy26', 'fcst26']); // columnas a comparar (NO incluyen la base): Ppto 2026 FY + Forecast 5+7 2026
  const [baseKey, setBaseKey] = React.useState('prop'); // columna base (fija, va primero) para Dif/% Dif = Ppto 2027
  const [showDif, setShowDif] = React.useState(true);   // Mostrar Dif / % Dif marcado por defecto
  const [colsOpen, setColsOpen] = React.useState(false);
  const [expanded, setExpanded] = React.useState(() => new Set());
  const [collapsed, setCollapsed] = React.useState(() => new Set()); // ramas cerradas a mano durante la búsqueda
  const q = st.q || ''; const setQ = v => set({ q: v });   // buscador compartido con "Gastos Corporativos"
  const [sortKey, setSortKey] = React.useState(null); // id de columna a ordenar (null = orden natural)
  const [sortDir, setSortDir] = React.useState('asc'); // 'asc' (menor→mayor) | 'desc'
  const [dragCol, setDragCol] = React.useState(null); // columna que se está arrastrando (reordenar)
  const colsRef = React.useRef(null);
  React.useEffect(() => {
    if (!colsOpen) return;
    const f = e => { if (colsRef.current && !colsRef.current.contains(e.target)) setColsOpen(false); };
    document.addEventListener('mousedown', f);
    return () => document.removeEventListener('mousedown', f);
  }, [colsOpen]);
  // Overflow de la barra de filtros: los que no quepan en una línea se colapsan en
  // un botón "+" (igual que el Dashboard). Colapsa de derecha a izquierda, de a uno,
  // con histéresis para no oscilar.
  const [nHidden, setNHidden] = React.useState(0);
  const [moreOpen, setMoreOpen] = React.useState(false);
  const barRef = React.useRef(null);
  const moreRef = React.useRef(null);
  const distRef = React.useRef(null);   // contenedor del botón Distribuible + su popover
  React.useEffect(() => {
    if (!distPopOpen) return;
    const f = e => { if (distRef.current && !distRef.current.contains(e.target)) setDistPopOpen(false); };
    document.addEventListener('mousedown', f);
    return () => document.removeEventListener('mousedown', f);
  }, [distPopOpen]);
  const collapseW = React.useRef(0);
  // Scroll horizontal: barra superior sincronizada con el contenedor de la tabla.
  const tblWrapRef = React.useRef(null);   // contenedor con overflow-x de la tabla
  const topScrollRef = React.useRef(null); // barrita de scroll arriba
  const syncing = React.useRef(false);
  const [tblW, setTblW] = React.useState(0); // ancho real (scrollWidth) de la tabla
  React.useLayoutEffect(() => { const el = tblWrapRef.current; if (el) setTblW(el.scrollWidth); });
  React.useEffect(() => {
    const upd = () => { const el = tblWrapRef.current; if (el) setTblW(el.scrollWidth); };
    window.addEventListener('resize', upd);
    return () => window.removeEventListener('resize', upd);
  }, []);
  const onTopScroll = () => { if (syncing.current) return; syncing.current = true; if (tblWrapRef.current && topScrollRef.current) tblWrapRef.current.scrollLeft = topScrollRef.current.scrollLeft; syncing.current = false; };
  const onTblScroll = () => { if (syncing.current) return; syncing.current = true; if (tblWrapRef.current && topScrollRef.current) topScrollRef.current.scrollLeft = tblWrapRef.current.scrollLeft; syncing.current = false; };
  React.useEffect(() => {
    if (!moreOpen) return;
    const f = e => { if (moreRef.current && !moreRef.current.contains(e.target)) setMoreOpen(false); };
    document.addEventListener('mousedown', f);
    return () => document.removeEventListener('mousedown', f);
  }, [moreOpen]);
  const N_COLLAPSIBLE = 9; // VP, Gerencia, Ítem Relevante, Código CECO, Código CLACO, Tipo Costo, Clasif Cuenta, ¿Aplica? (van al "+")
  React.useLayoutEffect(() => {
    const el = barRef.current;
    if (!el) return;
    const measure = () => {
      // La barra envuelve (no encoge): detectamos salto de fila por la ALTURA total.
      // Si la barra es más alta que un solo control + padding, hay 2+ filas → colapsar.
      const maxH = Array.from(el.children).reduce((m, k) => Math.max(m, k.offsetHeight), 0);
      const wrapped = el.clientHeight > maxH + 36; // padding 28 + tolerancia
      if (wrapped && nHidden < N_COLLAPSIBLE) {
        collapseW.current = el.clientWidth; setNHidden(n => Math.min(N_COLLAPSIBLE, n + 1));
      } else if (!wrapped && nHidden > 0 && el.clientWidth > collapseW.current + 150) {
        setNHidden(n => Math.max(0, n - 1));
      }
    };
    measure();
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(measure) : null;
    if (ro) ro.observe(el);
    window.addEventListener('resize', measure);
    return () => { if (ro) ro.disconnect(); window.removeEventListener('resize', measure); };
  }, [nHidden, dataMode]);

  const corpOn = dataMode === 'corp' || dataMode === 'both';
  const distOn = dataMode === 'dist' || dataMode === 'both';
  const setFlags = (c, d) => setDataMode(c && d ? 'both' : c ? 'corp' : d ? 'dist' : 'none');

  // "Desc. CLACO" es un nivel OPCIONAL disponible en cualquier estructura: se inyecta tras
  // "Ítem" (su padre natural) si no está ya, y arranca oculto (ver hiddenDims) → aparece como
  // chip "+ Desc. CLACO" para agregarlo. Solo aquí (Tabla Resumen), no toca los presets compartidos.
  const _rawBase = dimOrder || RESUMEN_DIMS[groupMode] || RESUMEN_DIMS.item;
  const _baseDims = _rawBase.includes('claco') ? _rawBase
    : (() => { const ii = _rawBase.indexOf('item'); const b = _rawBase.slice(); b.splice(ii < 0 ? b.length : ii + 1, 0, 'claco'); return b; })();
  const _structDims = _baseDims.filter(d => showCeco || (d !== 'ceco' && d !== 'dceco'));
  const dims = _structDims.filter(d => !hiddenDims.has(d));   // niveles activos (excluye los quitados con ×)
  const ghostDims = _structDims.filter(d => hiddenDims.has(d)); // quitados → se muestran como chips para re-agregar
  // Reordenar niveles arrastrando el breadcrumb: mueve 'from' delante de 'to'.
  const reorderDim = (from, to) => {
    if (!from || from === to) return;
    const cur = _baseDims.slice();
    const fi = cur.indexOf(from); if (fi < 0) return;
    cur.splice(fi, 1);
    const ti = cur.indexOf(to);
    cur.splice(ti < 0 ? cur.length : ti, 0, from);
    setDimOrder(cur); setExpanded(new Set());
  };
  // La base se elige aparte (cualquier columna). Las "columnas a comparar" (selCols)
  // NO incluyen la base. La base SIEMPRE va primero; las demás siguen el orden del catálogo.
  const baseCol = RESUMEN_COL_CATALOG.find(c => c.key === baseKey)
    || RESUMEN_COL_CATALOG.find(c => selCols.includes(c.key)) || null;
  // Orden = el de selCols (reordenable arrastrando los encabezados). Excluye la base.
  const compareCols = selCols
    .map(k => RESUMEN_COL_CATALOG.find(c => c.key === k))
    .filter(c => c && (!baseCol || c.key !== baseCol.key));
  const cols = baseCol ? [baseCol, ...compareCols] : compareCols;
  const reorderCols = (fromKey, toKey) => {
    if (!fromKey || fromKey === toKey) return;
    setSelCols(prev => {
      const a = prev.filter(k => k !== fromKey);
      const i = a.indexOf(toKey);
      if (i < 0) a.push(fromKey); else a.splice(i, 0, fromKey);
      return a;
    });
  };
  const neededYears = [...new Set(cols.filter(c => c.y).map(c => c.y))];

  const tree = A.buildTree({
    years: neededYears, showProp: true, yearAgg: 'byYear', version: 'ORI',
    dataMode, companies: [], vps, gers, itemrels, items, tcs, clases, companias: distOn ? companias : [], cecos, clacos, st: stMode, aps, hidden, groupBy: dims,
    sort: { key: 'real', dir: 'desc' }, overrides, growth: A.DEF_GROWTH,
  });

  const yrOf = (agg, y) => (agg.yr && agg.yr[y]) || { real: 0, ver: 0 };
  const valOf = (col, node) => {
    const agg = node.agg;
    if (col.kind === 'prop') return agg.prop || 0;
    if (col.kind === 'fcst') return agg.fcst || 0;   // Forecast 5+7 2026 (anual)
    if (col.kind === 'planfy') return agg.fy26 || 0; // Ppto 2026 anual (FY)
    if (col.kind === 'real') return yrOf(agg, col.y).real;
    if (col.kind === 'plan') return yrOf(agg, col.y).ver;
    return 0;
  };

  // Ordenamiento por columna (clic en encabezados). Ordena el árbol recursivamente
  // (cada nivel por el mismo criterio) antes de aplanarlo, conservando la jerarquía.
  const colByKey = {}; cols.forEach(c => { colByKey[c.key] = c; });
  const sortValue = node => {
    if (!sortKey || sortKey === '__name') return 0;
    let base = sortKey, kind = 'val';
    if (sortKey.endsWith('_d')) { base = sortKey.slice(0, -2); kind = 'dif'; }
    else if (sortKey.endsWith('_p')) { base = sortKey.slice(0, -2); kind = 'pct'; }
    const c = colByKey[base]; if (!c) return 0;
    const v = valOf(c, node);
    if (kind === 'val') return v;
    const bv = baseCol ? valOf(baseCol, node) : 0;
    const dif = bv - v;
    if (kind === 'dif') return dif;
    return v ? dif / Math.abs(v) : NaN;   // % Dif indefinido si v=0
  };
  let sortCmp = null;   // comparador activo (se pasa a flattenTree para ordenar también el detalle)
  if (sortKey) {
    const dir = sortDir === 'asc' ? 1 : -1;
    const byName = (a, b) => String(a.name).localeCompare(String(b.name), 'es') * dir;
    const byNum = (a, b) => {
      const va = sortValue(a), vb = sortValue(b);
      const na = va == null || isNaN(va), nb = vb == null || isNaN(vb);
      if (na && nb) return 0; if (na) return 1; if (nb) return -1; // indefinidos al final
      return (va - vb) * dir;
    };
    sortCmp = sortKey === '__name' ? byName : byNum;
    const sortRec = ns => { ns.sort(sortCmp); ns.forEach(n => n.children && n.children.length && sortRec(n.children)); };
    sortRec(tree.vpNodes);
  }
  const clickSort = key => { if (sortKey === key) setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortKey(key); setSortDir('asc'); } };
  const sortArrow = key => sortKey === key ? (sortDir === 'asc' ? ' ▲' : ' ▼') : '';

  // v3: detalle (Texto pedido › Denominación) bajo cada Contrapartida, con valores por año.
  const flat = window.flattenTree(tree, expanded, q, dims, dims.includes('contra'), detOrder, collapsed, stMode, showDetP, sortCmp);
  // Total: normalmente tree.total (todo lo que pasa los filtros). Con búsqueda activa, el
  // buscador filtra filas en pantalla pero NO tree.total → el Total sumaría de más. Entonces
  // recalculamos el Total sumando los nodos de PRIMER nivel visibles (los que la búsqueda dejó).
  const _sumAggs = aggs => {
    const t = { real: 0, version: 0, ytdReal: 0, ytdVersion: 0, fy26: 0, fcst: 0, prop: 0, yr: {} };
    aggs.forEach(a => {
      if (!a) return;
      t.fy26 += a.fy26 || 0; t.fcst += a.fcst || 0; t.prop += a.prop || 0;
      if (a.yr) for (const y in a.yr) { const c = t.yr[y] || (t.yr[y] = { real: 0, ver: 0 }); c.real += (a.yr[y].real || 0); c.ver += (a.yr[y].ver || 0); }
    });
    return t;
  };
  const _topRows = flat.filter(r => r.level === 1);
  const totalAgg = q ? _sumAggs(_topRows.map(r => r.node.agg)) : tree.total;
  // Al buscar, las ramas que coinciden se abren solas: el toggle marca/desmarca "colapsado"
  // (permite esconder el resto). Sin búsqueda, opera sobre el set normal de expandidos.
  const onToggle = key => {
    if (q) setCollapsed(prev => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });
    else setExpanded(prev => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });
  };
  // Redimensionar la columna de nombres arrastrando el borde del encabezado. Doble clic = automático.
  const startResize = e => {
    e.preventDefault(); e.stopPropagation();
    const th = e.currentTarget.parentElement;
    const startW = th.offsetWidth, startX = e.clientX;
    const cap = Math.round((window.innerWidth || 1400) * 0.7);
    const move = ev => setNameW(Math.max(150, Math.min(cap, startW + (ev.clientX - startX))));
    const up = () => { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); document.body.style.userSelect = ''; };
    document.body.style.userSelect = 'none';
    document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
  };
  const spacerW = nameW != null ? '100%' : 0;   // cuando la columna está fijada, este spacer absorbe el sobrante
  React.useEffect(() => { setCollapsed(new Set()); }, [q]);   // cada búsqueda nueva parte sin colapsos manuales
  // Click-para-filtrar: VP/Gerencia/Ítem togglean su filtro; CECO/Contrapartida usan el buscador.
  const _FKset = { vp: setVps, ger: setGers, itemrel: setItemrels, item: setItems };
  const _FKcur = { vp: vps, ger: gers, itemrel: itemrels, item: items };
  const onPick = row => {
    const dim = row.dim || dims[row.level - 1];
    const val = row.node.name;
    if (_FKset[dim]) {
      const cur = _FKcur[dim] || [];
      if (cur.length === 1 && cur[0] === val) { _FKset[dim]([]); return; }   // toggle off
      _FKset[dim]([val]);
      dims.slice(dims.indexOf(dim) + 1).forEach(d => { if (_FKset[d]) _FKset[d]([]); }); // limpia dims más profundas
    } else {
      setQ(q === val ? '' : val);   // ceco / contrapartida → buscador (toggle)
    }
  };
  const rowActive = row => {
    const dim = row.dim || dims[row.level - 1];
    if (_FKset[dim]) { const cur = _FKcur[dim] || []; return cur.length === 1 && cur[0] === row.node.name; }
    return q === row.node.name;
  };
  const expandAll = () => {
    const n = new Set();
    const walk = (nodes, prefix) => nodes.forEach(nd => {
      const key = prefix ? prefix + '|' + nd.name : nd.name;
      if (nd.children && nd.children.length) { n.add(key); walk(nd.children, key); }
    });
    walk(tree.vpNodes, '');
    setExpanded(n);
    setCollapsed(new Set());   // quita colapsos manuales para que abra todo (incluso con búsqueda)
  };
  // Al buscar, el árbol se fuerza abierto: para colapsar de verdad, marcamos como colapsadas
  // TODAS las ramas visibles (el set 'collapsed' gana sobre el auto-abrir de la búsqueda).
  const collapseAll = () => { setExpanded(new Set()); setCollapsed(q ? new Set(flat.filter(r => r.expandable).map(r => r.key)) : new Set()); };

  const dimsOpt = A.dimsFor({ dataMode, companies: [], vps, gers });
  // VP solo de registros de GASTO (no Dotaciones, que usan otra convención de nombres).
  const vpOpts = [...new Set(A.activeRecords({ dataMode }).map(r => r.vp))]
    .sort((a, b) => A.dispVP(a).localeCompare(A.dispVP(b), 'es'))
    .map(v => ({ value: v, label: A.dispVP(v) }));
  const gerOpts = dimsOpt.gers.map(v => ({ value: v, label: A.dispGer(v) }));
  const itemrelOpts = (dimsOpt.itemrels || []).map(v => ({ value: v, label: A.dispItemRel(v) }));
  const tcOpts = (dimsOpt.tcs || []).map(v => ({ value: v, label: v }));
  const clasOpts = (dimsOpt.clases || []).map(v => ({ value: v, label: v }));
  // Toggle rápido "Incluir Mano de Obra" (coordinado con el filtro Clasificación Cuenta: ambos
  // leen/escriben `clases`). clases=[] significa TODAS (incluye Mano de Obra). Por defecto la
  // Mano de Obra está fuera (defaultClases), así que el check arranca desmarcado.
  const MDO = 'Mano de Obra';
  const hasMdO = clasOpts.some(o => o.value === MDO);
  const mdoIncluida = clases.length === 0 || clases.includes(MDO);
  const toggleMdO = () => {
    const all = clasOpts.map(o => o.value);
    if (!mdoIncluida) { setClases([...clases, MDO]); return; }
    // Excluir MdO. Si el resultado quedara vacío ([] = todas → volvería a incluirla), se cae
    // a "todas menos Mano de Obra" para respetar el destilde.
    const keep = (clases.length ? clases : all).filter(c => c !== MDO);
    setClases(keep.length ? keep : all.filter(c => c !== MDO));
  };
  // Filtro por código CECO: acotado a VP/Gerencia activas + los ya seleccionados (aunque queden fuera).
  const cecoOpts = [...new Set([...(dimsOpt.cecos || []), ...cecos])].sort((a, b) => a.localeCompare(b, 'es')).map(v => ({ value: v, label: v }));
  const clacoOpts = [...new Set([...(dimsOpt.clacos || []), ...clacos])].sort((a, b) => a.localeCompare(b, 'es')).map(v => ({ value: v, label: v }));   // filtro por código CLACO (Clase de Costo)
  // "Código CLACO": si hay algún otro filtro activo (VP/Gerencia/Ítem/CECO/…), se tildan los
  // CLACOs que realmente se están mostrando (presentes tras el resto de filtros); sin otro
  // filtro se deja "Todos" sin tildar (el usuario no eligió nada aún).
  const clacoNarrowed = !!(vps.length || gers.length || itemrels.length || items.length || cecos.length || tcs.length || aps.length || (distOn && companias.length));
  const clacoScope = React.useMemo(() => !clacoNarrowed ? [] : A.clacosInScope({
    dataMode, companies: [], vps, gers, itemrels, items, tcs, clases,
    companias: distOn ? companias : [], cecos, st: stMode, aps, hidden,
  }), [clacoNarrowed, dataMode, vps, gers, itemrels, items, tcs, clases, companias, distOn, cecos, stMode, aps, hidden]);
  const compOpts = (dimsOpt.companias || []).map(v => ({ value: v, label: v }));
  // Color de marca por compañía (nombre → abrev → color de A.COMPANIAS; MLP/ANT/CEN/CMZ tienen color).
  const compColor = React.useMemo(() => {
    const byAbrev = {}; (A.COMPANIAS || []).forEach(c => { byAbrev[c.id] = c.color; });
    const m = {}; Object.values((A.V && A.V.comps) || {}).forEach(c => { if (c && c.nombre) m[c.nombre] = byAbrev[c.abrev] || null; });
    return m;
  }, []);
  // Compañías distribuibles (clasif "Distribuibles") — el filtro Compañía es un concepto del distribuible.
  const distCompNames = React.useMemo(() => {
    const s = new Set(); Object.values((A.V && A.V.comps) || {}).forEach(c => { if (c && c.nombre && /distrib/i.test(c.clasif || '')) s.add(c.nombre); });
    return s;
  }, []);
  // Lista de compañías del popover Distribuible: SIEMPRE todas las distribuibles (no depende del
  // modo actual), si no en modo solo Corporativo quedaría vacía y no podrías reactivar Distribuible.
  const compChips = (A.dimsFor({ dataMode: 'both' }).companias || []).filter(v => distCompNames.has(v)).map(v => ({ value: v, label: v }));
  const compAbrev = React.useMemo(() => {
    const m = {}; Object.values((A.V && A.V.comps) || {}).forEach(c => { if (c && c.nombre) m[c.nombre] = c.abrev || ''; });
    return m;
  }, []);
  // Popover de compañías anclado al botón Distribuible (elegir a qué operaciones se reparte).
  const allCompVals = compChips.map(o => o.value);
  const openDistPop = () => { setDraftComp(companias.length ? companias.filter(c => allCompVals.includes(c)) : allCompVals.slice()); setDistPopOpen(true); };
  const toggleDraft = v => setDraftComp(d => d.includes(v) ? d.filter(x => x !== v) : [...d, v]);
  const allDraft = allCompVals.length > 0 && draftComp.length === allCompVals.length;
  const toggleAllDraft = () => setDraftComp(allDraft ? [] : allCompVals.slice());
  const applyDist = () => {
    if (draftComp.length === 0) { setFlags(corpOn, false); setCompanias([]); }                       // nada marcado → sin Distribuible
    else { setFlags(corpOn, true); setCompanias(draftComp.length === allCompVals.length ? [] : draftComp); }  // [] = todas
    setDistPopOpen(false);
  };
  const apOpts = (dimsOpt.aps || []).map(v => ({ value: v, label: v }));

  // v3: bajo cada Contrapartida se abren Texto pedido › Denominación (detalle del gasto).
  const detailOn = !!(A.hasDetail && A.hasDetail()) && dims[dims.length - 1] === 'contra';
  const DET_LBL = { texto: 'Texto pedido', denom: 'Denominación' };
  const detDims = detOrder === 'dt' ? ['denom', 'texto'] : ['texto', 'denom'];
  const detLbls = detailOn ? detDims.map(d => DET_LBL[d]) : [];
  // Detalle Ppto/Forecast (Concepto Gasto › Actividad) — cuelga del Ítem cuando existe DETP.
  const detailPOn = !!(A.hasDetailP && A.hasDetailP()) && dims.includes('item');
  const DETP_DIMS = ['concepto', 'actividad'];
  const DETP_LBL = { concepto: 'Concepto Gasto', actividad: 'Actividad' };
  const famBadge = fam => (
    <span style={{ marginLeft: 5, fontSize: 8, fontWeight: 800, letterSpacing: '.04em', padding: '1px 4px', borderRadius: 3, verticalAlign: 'middle',
      background: fam === 'real' ? 'var(--amsa-teal)' : 'var(--amsa-yellow)', color: fam === 'real' ? '#fff' : '#3a2e10' }}>{fam === 'real' ? 'REAL' : 'PPTO'}</span>
  );
  const headerLbl = dims.map(d => RESUMEN_DIM_LBL[d]).concat(detLbls).concat(detailPOn && showDetP ? DETP_DIMS.map(d => DETP_LBL[d]) : []).join(' › ');
  // Mismo alto (34px) y estilo que los <select> para que «Datos» quede alineado con «Estructura».
  const chipSt = on => ({ height: 34, boxSizing: 'border-box', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', border: on ? '1px solid var(--amsa-teal)' : '1px solid var(--teal-border)', padding: '0 14px', cursor: 'pointer', borderRadius: 6, background: on ? 'var(--amsa-teal)' : '#fff', color: on ? '#fff' : 'var(--fg-soft)', fontFamily: 'var(--font-sans)', fontWeight: 600, fontSize: 12.5 });
  const cap = { fontFamily: 'var(--font-disp)', fontWeight: 700, fontSize: 9.5, letterSpacing: '.07em', color: '#8a9499', textTransform: 'uppercase', marginBottom: 4 };
  const fctlSel = { height: 34, padding: '0 10px', border: '1px solid #cdd6d8', borderRadius: 6, fontSize: 12.5, fontFamily: 'var(--font-sans)', color: 'var(--ink)', cursor: 'pointer', background: '#fff', width: '100%', maxWidth: 200, boxSizing: 'border-box' };
  const valHdr = isBase => ({ background: isBase ? '#717981' : 'var(--amsa-yellow)', color: isBase ? '#fff' : '#3a2e10' });
  const num = (v, k, extra) => <td key={k} className="tnum" style={{ textAlign: 'right', ...extra }}>{A.fmt(v, unit, dec)}</td>;
  // Familia de cada columna: Real vs Ppto/Forecast. Sirve para atenuar las que no aplican al grano.
  const colFam = c => (c.kind === 'real' ? 'real' : 'ppto');
  const HATCH = 'repeating-linear-gradient(-45deg,transparent,transparent 5px,rgba(20,81,90,.05) 5px,rgba(20,81,90,.05) 6px)';
  const dashTd = k => <td key={k} className="tnum" style={{ textAlign: 'right', color: 'var(--fg-soft)', background: HATCH }} title="No aplica a este nivel de detalle">—</td>;
  // fam = familia del grano (row): 'real' (Contrapartida/Texto pedido/Denominación) o 'ppto'
  // (Concepto Gasto/Actividad). Si el grano pertenece a una familia, las columnas de la OTRA
  // se muestran como "—" (no 0), porque ese dato no existe a ese nivel.
  const rowCells = (node, fam) => {
    const baseOff = !!(fam && baseCol && colFam(baseCol) !== fam);
    const baseVal = baseCol ? valOf(baseCol, node) : 0;
    const out = [];
    cols.forEach(c => {
      const off = !!(fam && colFam(c) !== fam);
      out.push(off ? dashTd(c.key) : num(valOf(c, node), c.key));
      if (baseCol && c.key !== baseCol.key && showDif) {
        if (off || baseOff) { out.push(dashTd(c.key + '_d')); out.push(dashTd(c.key + '_p')); }
        else {
          const v = valOf(c, node);
          const dif = baseVal - v;
          const pct = v ? dif / Math.abs(v) : null;
          out.push(<td key={c.key + '_d'} className="tnum" style={{ textAlign: 'right' }}>{(dif > 0 ? '+' : '') + A.fmt(dif, unit, dec)}</td>);
          out.push(<td key={c.key + '_p'} className="tnum pct" style={{ textAlign: 'right' }}>{pct == null ? '—' : (pct > 0 ? '+' : '') + A.fmtPct(pct, 0)}</td>);
        }
      }
    });
    return out;
  };

  // Filtros que pueden colapsar al "+" cuando no caben (de derecha a izquierda).
  const collapsibleEls = [
    <div className="fgroup grow" style={{ minWidth: 130 }} key="vp">
      <div style={cap}>Vicepresidencia</div>
      <div className="fctl">
        <MultiSelect options={vpOpts} selected={vps} placeholder="Todas" searchable
          onChange={v => {
            // Al elegir VP(s), se tildan automáticamente sus CECOs en el filtro Código CECO
            // (la lista sigue completa; puedes destildar o agregar otros).
            const vpCecos = [...new Set(A.records.filter(r => v.includes(r.vp)).map(r => r.ceco))];
            set({ vps: v, gers: gers.filter(x => !v.length || A.records.some(r => v.includes(r.vp) && r.ger === x)), cecos: [...new Set([...cecos, ...vpCecos])] });
          }} />
      </div>
    </div>,
    <div className="fgroup grow" style={{ minWidth: 130 }} key="ger">
      <div style={cap}>Gerencia</div>
      <div className="fctl">
        <MultiSelect options={gerOpts} selected={gers} placeholder="Todas" searchable onChange={setGers} />
      </div>
    </div>,
    <div className="fgroup grow" style={{ minWidth: 130 }} key="itemrel">
      <div style={cap}>Ítem Relevante</div>
      <div className="fctl">
        <MultiSelect options={itemrelOpts} selected={itemrels} placeholder="Todas" searchable onChange={setItemrels} />
      </div>
    </div>,
    <div className="fgroup grow" style={{ minWidth: 130 }} key="ceco">
      <div style={cap}>Código CECO</div>
      <div className="fctl">
        <MultiSelect options={cecoOpts} selected={cecos} placeholder="Todos" searchable onChange={setCecos} />
      </div>
    </div>,
    <div className="fgroup grow" style={{ minWidth: 130 }} key="claco">
      <div style={cap}>Código CLACO</div>
      <div className="fctl">
        <MultiSelect options={clacoOpts} selected={clacos} placeholder="Todos" searchable onChange={setClacos}
          impliedAll={clacoNarrowed ? clacoScope : undefined} />
      </div>
    </div>,
    <div className="fgroup" style={{ minWidth: 128 }} key="tc">
      <div style={cap}>Tipo Costo</div>
      <div className="fctl">
        <MultiSelect options={tcOpts} selected={tcs} placeholder="Todos" onChange={setTcs} />
      </div>
    </div>,
    <div className="fgroup" style={{ minWidth: 150 }} key="clas">
      <div style={cap}>Clasificación Cuenta</div>
      <div className="fctl">
        <MultiSelect options={clasOpts} selected={clases} placeholder="Todas" onChange={setClases} />
      </div>
    </div>,
    <div className="fgroup" style={{ minWidth: 110 }} key="ap">
      <div style={cap}>¿Aplica?</div>
      <div className="fctl">
        <MultiSelect options={apOpts} selected={aps} placeholder="Todas" onChange={setAps} />
      </div>
    </div>,
  ];
  const nShown = N_COLLAPSIBLE - nHidden;

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 12, margin: '4px 0 12px' }}>
        <div>
          <h2 style={{ fontFamily: 'var(--font-disp)', fontWeight: 800, fontSize: 18, color: 'var(--ink)', margin: 0 }}>Tabla Resumen</h2>
          <div style={{ fontSize: 12, color: 'var(--fg-muted)', marginTop: 2 }}>Comparador por período · {headerLbl} · {unit}</div>
        </div>
      </div>

      {/* Barra de controles: envuelve; lo que pasaría a una 2ª fila se colapsa al "+". */}
      <div className="filters" style={{ flexWrap: 'wrap', gap: 14, alignItems: 'flex-end' }} ref={barRef}>
        <div className="fgroup" style={{ minWidth: 150 }}>
          <div style={cap}>Base Moneda</div>
          <select value={valMode} onChange={e => onValMode(e.target.value)} style={fctlSel}>
            <option value="n">Moneda original</option>
            <option value="a">Moneda Ajustada 2027</option>
          </select>
        </div>
        <div className="fgroup" style={{ minWidth: 120 }}>
          <div style={cap}>Estructura CECOS</div>
          <select value={cecoMode} onChange={e => onCecoMode(e.target.value)} style={fctlSel}>
            <option value="new">Nuevos</option>
            <option value="old">Antiguos</option>
          </select>
        </div>
        <div className="fgroup" style={{ minWidth: 200 }}>
          <div style={cap}>Datos</div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button type="button" style={chipSt(corpOn)} onClick={() => setFlags(!corpOn, distOn)}>Corporativo</button>
            <span style={{ position: 'relative', display: 'inline-flex' }} ref={distRef}>
              <button type="button" style={{ ...chipSt(distOn), gap: 6 }} onClick={openDistPop}
                title="Elegir a qué compañías se reparte el gasto distribuible">
                Distribuible{distOn && companias.length ? ' (' + companias.length + ')' : ''}
                <span style={{ fontSize: 9, opacity: 0.8 }}>▾</span>
              </button>
              {distPopOpen && (
                <div style={{ position: 'absolute', top: 'calc(100% + 6px)', left: 0, zIndex: 60, background: '#fff', border: '1px solid var(--line)', borderRadius: 10, boxShadow: 'var(--shadow-2)', padding: 14, width: 250, maxWidth: 'calc(100vw - 32px)' }}>
                  <div style={{ fontFamily: 'var(--font-disp)', fontWeight: 700, fontSize: 13, color: 'var(--ink)' }}>¿Qué compañía distribuir?</div>
                  <div style={{ fontSize: 11, color: 'var(--fg-muted)', margin: '2px 0 10px' }}>El gasto distribuible se reparte a operaciones</div>
                  <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', padding: '4px 0', fontWeight: 700, color: 'var(--ink)', fontSize: 12.5 }}>
                    <input type="checkbox" checked={allDraft} onChange={toggleAllDraft} /> Todas las compañías
                  </label>
                  <div style={{ height: 1, background: 'var(--line)', margin: '6px 0' }} />
                  {compChips.map(o => (
                    <label key={o.value} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', padding: '4px 0', fontSize: 12.5, color: 'var(--fg-1)' }}>
                      <input type="checkbox" checked={draftComp.includes(o.value)} onChange={() => toggleDraft(o.value)} />
                      <span style={{ width: 8, height: 8, borderRadius: '50%', background: compColor[o.value] || 'var(--amsa-teal)', flex: '0 0 auto' }} />
                      <span style={{ flex: 1 }}>{o.label}</span>
                      <span style={{ fontSize: 10.5, fontWeight: 700, color: 'var(--fg-soft)', letterSpacing: '.03em' }}>{compAbrev[o.value] || ''}</span>
                    </label>
                  ))}
                  {compChips.length === 0 && <div style={{ fontSize: 12, color: 'var(--fg-soft)', padding: '4px 0' }}>Sin compañías distribuibles.</div>}
                  <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
                    <button type="button" onClick={applyDist} style={{ flex: 1, height: 32, border: 'none', borderRadius: 7, background: 'var(--amsa-teal)', color: '#fff', fontFamily: 'var(--font-sans)', fontWeight: 700, fontSize: 12.5, cursor: 'pointer' }}>Aplicar</button>
                    <button type="button" onClick={() => setDistPopOpen(false)} style={{ height: 32, padding: '0 14px', border: '1px solid var(--teal-border)', borderRadius: 7, background: '#fff', color: 'var(--fg-soft)', fontFamily: 'var(--font-sans)', fontWeight: 600, fontSize: 12.5, cursor: 'pointer' }}>Cancelar</button>
                  </div>
                </div>
              )}
            </span>
          </div>
        </div>
        <div className="fgroup" style={{ minWidth: 150 }}>
          <div style={cap}>Estructura</div>
          <select value={groupMode} onChange={e => { setGroupMode(e.target.value); setDimOrder(null); setExpanded(new Set()); }} style={fctlSel}
            title="La tabla mantiene todo el detalle; esto define el orden principal (arrastra el encabezado para reordenar).">
            <option value="orgcc">VP › Ítem Relevante</option>
            <option value="itemcc">Ítem Relevante › VP</option>
          </select>
        </div>
        <div className="fgroup" style={{ minWidth: 160 }}>
          <div style={cap}>Columna Base</div>
          <select value={baseCol ? baseCol.key : ''} style={fctlSel}
            title="Columna fija (va primero) contra la que se calculan Dif y % Dif"
            onChange={e => {
              const k = e.target.value, old = baseCol ? baseCol.key : null;
              setBaseKey(k);
              // la base sale de «columnas a comparar»; la base anterior pasa a comparar.
              setSelCols(p => { let n = p.filter(x => x !== k); if (old && old !== k && !n.includes(old)) n = [...n, old]; return n; });
            }}>
            {[...new Set(RESUMEN_COL_CATALOG.map(c => c.cat))].map(cat => (
              <optgroup key={cat} label={cat}>
                {RESUMEN_COL_CATALOG.filter(c => c.cat === cat).map(c => <option key={c.key} value={c.key}>{c.label}</option>)}
              </optgroup>
            ))}
          </select>
        </div>
        <div className="fgroup" style={{ minWidth: 170, position: 'relative' }} ref={colsRef}>
          <div style={cap}>Columnas a comparar</div>
          <button type="button" onClick={() => setColsOpen(o => !o)} style={{ ...fctlSel, textAlign: 'left', width: '100%', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
            <span>{compareCols.length} período{compareCols.length === 1 ? '' : 's'}</span><span style={{ fontSize: 9, color: 'var(--amsa-teal)' }}>▾</span>
          </button>
          {colsOpen && (
            <div style={{ position: 'absolute', top: 'calc(100% + 3px)', left: 0, zIndex: 50, background: '#fff', border: '1px solid var(--line)', borderRadius: 6, boxShadow: 'var(--shadow-2)', padding: 6, minWidth: 220, maxHeight: 340, overflowY: 'auto' }}>
              {[...new Set(RESUMEN_COL_CATALOG.map(c => c.cat))].map(cat => {
                const opts = RESUMEN_COL_CATALOG.filter(c => c.cat === cat && (!baseCol || c.key !== baseCol.key)); // la base no se compara consigo misma
                if (!opts.length) return null;
                return (
                <div key={cat}>
                  <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: '.05em', textTransform: 'uppercase', color: 'var(--fg-soft)', margin: '6px 4px 2px' }}>{cat}</div>
                  {opts.map(c => (
                    <label key={c.key} className="ms-opt" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <input type="checkbox" checked={selCols.includes(c.key)}
                        onChange={() => setSelCols(p => p.includes(c.key) ? p.filter(x => x !== c.key) : [...p, c.key])} />
                      <span>{c.label}</span>
                    </label>
                  ))}
                </div>);
              })}
            </div>
          )}
        </div>

        {/* Los filtros secundarios (VP, Gerencia, Ítem Relevante, Tipo Costo, ¿Aplica?) van
            SIEMPRE agrupados en el "+" → la barra principal queda en una sola fila. */}
        <div className="fgroup" style={{ minWidth: 'auto', position: 'relative' }} ref={moreRef}>
          <div style={cap} aria-hidden="true">&nbsp;</div>
          <button type="button" style={{ ...fctlSel, width: 46, maxWidth: 46, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, fontWeight: 700, color: 'var(--amsa-teal)' }}
            title="Más filtros (VP, Gerencia, Ítem Relevante, Código CECO, Código CLACO, Tipo Costo, Clasificación Cuenta, ¿Aplica?)" onClick={() => setMoreOpen(o => !o)}>+</button>
          {moreOpen && (
            <div style={{ position: 'absolute', top: 'calc(100% + 3px)', left: 0, zIndex: 50, background: '#fff', border: '1px solid var(--line)', borderRadius: 6, boxShadow: 'var(--shadow-2)', padding: 12, display: 'flex', flexDirection: 'column', gap: 12, width: 260, maxWidth: 'calc(100vw - 32px)' }}>
              {collapsibleEls}
            </div>
          )}
        </div>
      </div>

      <div className="matrix-card" style={{ marginTop: 12 }}>
        <div className="matrix-top">
          <div>
            <h3>Comparación por período</h3>
            <div className="mt-sub">{({ both: 'Corporativo + Distribuible', corp: 'Corporativo', dist: 'Distribuible', none: 'Sin datos' })[dataMode]}{baseCol ? ' · base: ' + baseCol.label : ''}</div>
          </div>
          <div className="toolbar">
            <span style={{ position: 'relative', display: 'inline-flex', alignItems: 'center' }}>
              <input type="text" value={q} onChange={e => setQ(e.target.value)} placeholder="Buscar VP / Gerencia / Ítem…"
                style={{ height: 30, width: 200, padding: '0 26px 0 10px', boxSizing: 'border-box', border: '1px solid var(--teal-200)', borderRadius: 6, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', outline: 'none' }} />
              {q && <button type="button" onClick={() => setQ('')} style={{ position: 'absolute', right: 6, border: 0, background: 'transparent', cursor: 'pointer', color: 'var(--fg-soft)', fontSize: 14, padding: 2 }}>×</button>}
            </span>
            {(vps.length || gers.length || itemrels.length || items.length || cecos.length || clacos.length || tcs.length || clases.length || companias.length || stMode !== 'excl' || aps.length || q) ? (
              <button type="button" onClick={() => { setVps([]); setGers([]); setItemrels([]); setItems([]); setCecos([]); setClacos([]); setTcs([]); setClases(defaultClases()); setCompanias([]); setStMode('excl'); setAps([]); setQ(''); }}
                title="Quitar filtros por clic / buscador"
                style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: 'var(--accent-wash)', color: 'var(--accent-900)', border: '1px solid var(--accent-300)', borderRadius: 7, padding: '6px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}>
                ✕ Limpiar filtros
              </button>
            ) : null}
            {hiddenCount > 0 && (
              <button type="button" onClick={unhideAll}
                title={'Ocultos: ' + Object.keys(hidden).flatMap(d => (hidden[d] || []).map(v => v)).join(', ') + '\nClic para restaurar todos'}
                style={{ display: 'inline-flex', alignItems: 'center', gap: 5, background: '#fff', color: 'var(--fg-3)', border: '1px solid var(--teal-border)', borderRadius: 7, padding: '6px 12px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}>
                ⊘ {hiddenCount} oculto{hiddenCount === 1 ? '' : 's'} · restaurar
              </button>
            )}
            {(A.hasST && A.hasST()) && (
              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 11.5, color: 'var(--fg-3)', fontWeight: 600, cursor: 'pointer' }}
                title="Services & Tech (CLACOs 6125020/6125021). Desmarca para excluirlos del panel.">
                <input type="checkbox" checked={stMode !== 'excl'} onChange={e => setStMode(e.target.checked ? '' : 'excl')} /> Incluir Services &amp; Tech
              </label>
            )}
            {hasMdO && (
              <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 11.5, color: 'var(--fg-3)', fontWeight: 600, cursor: 'pointer' }}
                title="Mano de Obra (Clasificación Cuenta). Por defecto excluida; marca para incluirla. Coordinado con el filtro «Clasificación Cuenta».">
                <input type="checkbox" checked={mdoIncluida} onChange={toggleMdO} /> Incluir Mano de Obra
              </label>
            )}
            <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 11.5, color: 'var(--fg-3)', fontWeight: 600, cursor: 'pointer' }}>
              <input type="checkbox" checked={showDif} onChange={e => setShowDif(e.target.checked)} /> Mostrar Dif / % Dif
            </label>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 7, fontSize: 11.5, color: 'var(--fg-3)', fontWeight: 600 }}>
              Decimales
              <span style={{ display: 'inline-flex', alignItems: 'stretch', height: 26, border: '1px solid var(--teal-200)', borderRadius: 6, overflow: 'hidden', background: '#fff' }}>
                <span style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', minWidth: 24, fontWeight: 700, color: 'var(--ink)', fontVariantNumeric: 'tabular-nums' }}>{dec}</span>
                <span style={{ display: 'flex', flexDirection: 'column', borderLeft: '1px solid var(--teal-100)' }}>
                  <button type="button" aria-label="Más decimales" onClick={() => onDec(Math.min(6, dec + 1))}
                    style={{ border: 0, borderBottom: '1px solid var(--teal-100)', background: 'var(--teal-wash2)', color: 'var(--teal-muted)', cursor: 'pointer', padding: '0 7px', fontSize: 7, lineHeight: '12px', flex: 1 }}>▲</button>
                  <button type="button" aria-label="Menos decimales" onClick={() => onDec(Math.max(0, dec - 1))}
                    style={{ border: 0, background: 'var(--teal-wash2)', color: 'var(--teal-muted)', cursor: 'pointer', padding: '0 7px', fontSize: 7, lineHeight: '12px', flex: 1 }}>▼</button>
                </span>
              </span>
            </span>
            {/* Un solo botón toggle: expande si todo está colapsado, colapsa si hay algo abierto. */}
            <button type="button" onClick={() => expanded.size ? collapseAll() : expandAll()}
              title={expanded.size ? 'Colapsar todas las filas' : 'Expandir todas las filas'}
              style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: 'var(--teal-wash)', color: 'var(--amsa-teal)', border: '1px solid var(--amsa-teal-light)', borderRadius: 7, padding: '6px 13px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>
              {expanded.size ? '⤒ Colapsar todo' : '⤓ Expandir todo'}
            </button>
          </div>
        </div>
        {/* Barrita de scroll horizontal ARRIBA, sincronizada con la tabla. */}
        <div ref={topScrollRef} onScroll={onTopScroll} className="hscroll-top" style={{ overflowX: 'auto', overflowY: 'hidden' }}>
          <div style={{ width: tblW, height: 1 }} />
        </div>
        <div ref={tblWrapRef} onScroll={onTblScroll} style={{ overflowX: 'auto' }}>
          {cols.length === 0
            ? <div className="empty-msg">Elige al menos una columna en «Columnas a comparar».</div>
            : (
            <table className="mtable resumen">
              <thead>
                <tr className="cols">
                  <th className="left" style={{ textAlign: 'left', userSelect: 'none', position: 'relative', ...(nameW != null ? { width: nameW, minWidth: nameW, maxWidth: nameW } : {}) }}>
                    <span onClick={() => clickSort('__name')} title="Ordenar alfabéticamente" style={{ cursor: 'pointer', marginRight: 6 }}>⇅{sortArrow('__name')}</span>
                    {dims.map((d, i) => (
                      <React.Fragment key={d}>
                        {i > 0 && <span style={{ color: 'var(--fg-muted)', margin: '0 3px' }}>›</span>}
                        <span draggable
                          onDragStart={() => setDragDim(d)}
                          onDragEnd={() => setDragDim(null)}
                          onDragOver={e => e.preventDefault()}
                          onDrop={e => { e.preventDefault(); reorderDim(dragDim, d); setDragDim(null); }}
                          title="Arrastra para reordenar este nivel"
                          style={{ cursor: 'grab', padding: '2px 4px 2px 7px', borderRadius: 5, whiteSpace: 'nowrap', display: 'inline-flex', alignItems: 'center', gap: 4,
                            background: dragDim === d ? 'var(--amsa-teal)' : 'var(--teal-100)', color: dragDim === d ? '#fff' : 'var(--amsa-teal-deep)',
                            opacity: dragDim === d ? 0.7 : 1 }}>
                          {RESUMEN_DIM_LBL[d]}{d === 'contra' && famBadge('real')}
                          {dims.length > 1 && (
                            <span onClick={e => { e.stopPropagation(); setHiddenDims(s => { const n = new Set(s); n.add(d); return n; }); setExpanded(new Set()); }}
                              title="Quitar este nivel de la estructura"
                              style={{ cursor: 'pointer', fontWeight: 700, fontSize: 13, lineHeight: 1, padding: '0 3px', borderRadius: 3, opacity: 0.6 }}>×</span>
                          )}
                        </span>
                      </React.Fragment>
                    ))}
                    {detDims.filter(() => detailOn).map(d => (
                      <React.Fragment key={d}>
                        <span style={{ color: 'var(--fg-muted)', margin: '0 3px' }}>›</span>
                        <span draggable
                          onDragStart={() => setDragDim('DET:' + d)}
                          onDragEnd={() => setDragDim(null)}
                          onDragOver={e => e.preventDefault()}
                          onDrop={e => { e.preventDefault(); if (String(dragDim).startsWith('DET:')) setDetOrder(o => o === 'dt' ? 'td' : 'dt'); setDragDim(null); }}
                          title="Detalle Real (cuelga bajo Contrapartida). Arrastra para intercambiar Texto pedido ↔ Denominación."
                          style={{ cursor: 'grab', padding: '2px 7px', borderRadius: 5, whiteSpace: 'nowrap', display: 'inline-flex', alignItems: 'center',
                            background: dragDim === 'DET:' + d ? 'var(--amsa-teal)' : 'transparent', color: dragDim === 'DET:' + d ? '#fff' : 'var(--fg-muted)',
                            border: '1px dashed var(--teal-border)' }}>{DET_LBL[d]}{famBadge('real')}</span>
                      </React.Fragment>
                    ))}
                    {/* Detalle Ppto/Forecast: Concepto Gasto › Actividad (cuelga del Ítem). Con × para quitarlo. */}
                    {detailPOn && showDetP && DETP_DIMS.map((d, i) => (
                      <React.Fragment key={d}>
                        <span style={{ color: 'var(--fg-muted)', margin: '0 3px' }}>›</span>
                        <span title="Detalle de Ppto 2027 y Forecast (cuelga del Ítem)."
                          style={{ padding: '2px 4px 2px 7px', borderRadius: 5, whiteSpace: 'nowrap', display: 'inline-flex', alignItems: 'center', gap: 4,
                            background: 'transparent', color: 'var(--fg-muted)', border: '1px dashed var(--accent-300)' }}>
                          {DETP_LBL[d]}{famBadge('ppto')}
                          <span onClick={() => { setShowDetP(false); setExpanded(new Set()); }}
                            title="Quitar el detalle de presupuesto (Concepto Gasto y Actividad)"
                            style={{ cursor: 'pointer', fontWeight: 700, fontSize: 13, lineHeight: 1, padding: '0 3px', borderRadius: 3, opacity: 0.6 }}>×</span>
                        </span>
                      </React.Fragment>
                    ))}
                    {detailPOn && !showDetP && (
                      <React.Fragment>
                        <span style={{ color: 'var(--fg-muted)', margin: '0 3px' }}>·</span>
                        <span onClick={() => { setShowDetP(true); setExpanded(new Set()); }}
                          title="Volver a mostrar el detalle de presupuesto (Concepto Gasto › Actividad)"
                          style={{ cursor: 'pointer', padding: '2px 8px', borderRadius: 5, whiteSpace: 'nowrap', display: 'inline-flex', alignItems: 'center',
                            background: 'transparent', color: 'var(--fg-muted)', border: '1px dashed var(--accent-300)', opacity: 0.85 }}>
                          + Concepto Gasto · Actividad {famBadge('ppto')}
                        </span>
                      </React.Fragment>
                    )}
                    {/* Niveles quitados (×): se muestran atenuados para volver a incluirlos con un clic. */}
                    {ghostDims.map(d => (
                      <React.Fragment key={'g' + d}>
                        <span style={{ color: 'var(--fg-muted)', margin: '0 3px' }}>·</span>
                        <span onClick={() => { setHiddenDims(s => { const n = new Set(s); n.delete(d); return n; }); setExpanded(new Set()); }}
                          title="Volver a incluir este nivel en la estructura"
                          style={{ cursor: 'pointer', padding: '2px 8px', borderRadius: 5, whiteSpace: 'nowrap', display: 'inline-block',
                            background: 'transparent', color: 'var(--fg-muted)', border: '1px dashed var(--teal-border)', opacity: 0.85 }}>
                          + {RESUMEN_DIM_LBL[d]}
                        </span>
                      </React.Fragment>
                    ))}
                    {/* Arrastrar = ajustar ancho de la columna de nombres · doble clic = automático (llena la pantalla). */}
                    <span className="col-resize" onMouseDown={startResize} onDoubleClick={() => setNameW(null)}
                      title="Arrastra para ajustar el ancho de esta columna · doble clic = automático" />
                  </th>
                  {cols.map((c, i) => {
                    const isBase = baseCol && c.key === baseCol.key;
                    const blocks = [<th key={c.key}
                        draggable={!isBase}
                        onDragStart={!isBase ? (() => setDragCol(c.key)) : undefined}
                        onDragOver={!isBase ? (e => { e.preventDefault(); }) : undefined}
                        onDrop={!isBase ? (() => { reorderCols(dragCol, c.key); setDragCol(null); }) : undefined}
                        onClick={() => clickSort(c.key)}
                        title={isBase ? 'Columna base' : 'Clic: ordenar · Arrastra para reordenar la columna'}
                        style={{ ...valHdr(isBase), cursor: isBase ? 'pointer' : 'grab', userSelect: 'none', opacity: dragCol === c.key ? 0.5 : 1 }}>{c.label}{isBase ? ' (base)' : ''}{sortArrow(c.key)}</th>];
                    if (!isBase && showDif) {
                      blocks.push(<th key={c.key + '_d'} onClick={() => clickSort(c.key + '_d')} style={{ background: 'var(--amsa-teal-deep)', color: '#fff', cursor: 'pointer', userSelect: 'none' }}>Dif{sortArrow(c.key + '_d')}</th>);
                      blocks.push(<th key={c.key + '_p'} onClick={() => clickSort(c.key + '_p')} style={{ background: 'var(--amsa-teal-deep)', color: '#fff', cursor: 'pointer', userSelect: 'none' }}>% Dif{sortArrow(c.key + '_p')}</th>);
                    }
                    return blocks;
                  })}
                  <th className="spacer-fill" style={{ width: spacerW }} aria-hidden="true"></th>
                </tr>
              </thead>
              <tbody>
                {flat.map(row => (
                  <tr key={row.key} className={'row-' + row.type}>
                    <td className="name" style={nameW != null ? { width: nameW, maxWidth: nameW } : undefined}><Twig row={row} expanded={expanded} onToggle={onToggle} dims={dims} onPick={onPick} active={rowActive(row)} onHide={hideVal} /></td>
                    {rowCells(row.node, row.fam)}
                    <td className="spacer-fill" style={{ width: spacerW }}></td>
                  </tr>
                ))}
                {flat.length === 0 && <tr><td className="name" colSpan={2 + cols.length + (showDif ? (cols.length - 1) * 2 : 0)}>Sin resultados</td></tr>}
                {flat.truncated && <tr><td colSpan={20} style={{ padding: '8px 12px', fontSize: 12, color: 'var(--amsa-teal-deep)', background: 'var(--accent-wash)', fontWeight: 600 }}>Mostrando las primeras {flat.truncated.toLocaleString('es-CL')} filas. Afiná la búsqueda o expandí manualmente para ver el resto.</td></tr>}
                <tr className="row-total">
                  <td className="name" style={nameW != null ? { width: nameW, maxWidth: nameW } : undefined}>Total</td>
                  {rowCells({ agg: totalAgg })}
                  <td className="spacer-fill" style={{ width: spacerW }}></td>
                </tr>
              </tbody>
            </table>
          )}
        </div>
        <div className="note" style={{ padding: '8px 14px 12px' }}>
          Elige Datos, filtros (VP / Gerencia / Ítem), estructura y los períodos a comparar. La 1ª columna es la base; «Dif» = base − período, «% Dif» = Dif / período. Clic en un encabezado para ordenar (menor→mayor; otra vez, mayor→menor). Arrastra un encabezado de período para reordenar las columnas. Clic en un VP / Gerencia / Ítem filtra el panel por ese elemento (clic de nuevo para quitar); CECO / Contrapartida filtran vía el buscador.
          Valor: elige «Normal» o «Moneda Ajustada 2027» en la barra de filtros. Ppto 2027 = Presupuesto 2027 cargado del Excel.
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
  const refresh = () => { if (window.CORP && window.CORP.resetTheme) window.CORP.resetTheme(); if (onApply) onApply(); };
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
  if (!edit) return null;
  return (
    <div style={{ position: 'relative', display: 'inline-flex', alignItems: 'center', alignSelf: 'center', marginLeft: 8, fontFamily: 'var(--font-sans)' }}>
      <button onClick={() => setOpen(o => !o)} title="Editar colores del dashboard" style={{
        display: 'flex', alignItems: 'center', gap: 6, height: 28, padding: '0 12px', borderRadius: 16,
        border: 0, background: 'var(--amsa-teal)', color: '#fff', fontWeight: 600, fontSize: 11.5,
        fontFamily: 'var(--font-sans)', cursor: 'pointer', boxShadow: '0 1px 5px rgba(20,40,50,.18)' }}>
        <span style={{ fontSize: 13 }}>🎨</span> Colores{nOv ? ' ·' + nOv : ''}
      </button>
      {open && (
        <div style={{ position: 'absolute', top: 'calc(100% + 8px)', right: 0, zIndex: 9999, width: 244, maxHeight: 'calc(100vh - 160px)', overflowY: 'auto', background: 'var(--card)', border: '1px solid var(--line)', borderRadius: 12, boxShadow: '0 10px 36px rgba(20,40,50,.18)', padding: '12px 14px' }}>
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
    </div>
  );
}

/* ---------- Pestaña Dotaciones: dashboard de FTE (Propios / Contratista) ---------- */
function DotacionesView() {
  const A = window.CORP;
  const { useState } = React;
  const unit = 'num', dec = 1;                 // FTE (Nº), 1 decimal
  const [showProp, setShowProp] = useState(true);   // Propios
  const [showCont, setShowCont] = useState(true);   // Contratista
  const [years, setYears] = useState([2025]);
  const [vps, setVps] = useState([]);
  const [gers, setGers] = useState([]);
  const [expanded, setExpanded] = useState(() => new Set());
  const [q, setQ] = useState('');
  const MS = window.MultiSelect, Chip = window.ModeChip;

  const dataMode = (showProp && showCont) ? 'dot' : showProp ? 'propios' : showCont ? 'contratista' : 'none';
  const opts = { years, dataMode, vps, gers, groupBy: ['vp', 'ger'], sort: { key: 'real', dir: 'desc' },
                 showProp: false, version: 'ORI', growth: A.DEF_GROWTH, overrides: {} };
  const tree = A.buildTree(opts);
  const series = A.annualSeries(opts);
  const thr = { red: 10, yellow: 0 };
  const fmtN = v => A.fmt(v, unit, dec);

  // Opciones de VP / Gerencia según el modo (propios/contratista/ambos).
  const baseRecs = A.dotRecords.filter(r => dataMode === 'dot' ? true : r.src === dataMode);
  const vpOpts = [...new Set(baseRecs.map(r => r.vp))].sort((a, b) => a.localeCompare(b, 'es')).map(v => ({ value: v, label: v }));
  const gerVals = vps.length ? [...new Set(baseRecs.filter(r => vps.includes(r.vp)).map(r => r.ger))] : [...new Set(baseRecs.map(r => r.ger))];
  const gerOpts = gerVals.sort((a, b) => a.localeCompare(b, 'es')).map(v => ({ value: v, label: v }));
  const yearOpts = [{ value: 2022, label: '2022' }, { value: 2023, label: '2023' }, { value: 2024, label: '2024' },
                    { value: 2025, label: '2025' }, { value: 2026, label: '2026 YTD' }, { value: '2026fy', label: '2026 Ppto FY' }];

  const T = tree.total;
  const histYears = years.filter(y => typeof y === 'number');
  // FTE = dotación PROMEDIO, no acumulable: buildTree devuelve la SUMA de los años
  // seleccionados, así que aquí se promedia. El Real se divide entre los años con
  // dato real (numéricos); el Ppto entre esos años + el Ppto FY 2026 si está activo.
  const nReal = Math.max(1, histYears.length);
  const nVer = Math.max(1, histYears.length + (years.includes('2026fy') ? 1 : 0));
  const avgR = a => a.ytdReal / nReal;
  const avgV = a => a.ytdVersion / nVer;
  const Treal = avgR(T), Tver = avgV(T);
  const cump = Tver ? Treal / Tver : 0;
  const dif = Treal - Tver;
  const cumpCol = A.kpiColor(Treal, Tver, thr);
  const yLbl = [...histYears.map(String), ...(years.includes('2026fy') ? ['2026 Ppto FY'] : [])].join(' + ') || '—';
  const modeLbl = dataMode === 'dot' ? 'Propios + Contratista' : dataMode === 'propios' ? 'Propios' : dataMode === 'contratista' ? 'Contratista' : 'Sin datos';
  const kpis = [
    { label: 'Dotación Real', value: fmtN(Treal), unit: 'FTE', sub: yLbl + ' · promedio' },
    { label: 'Dotación Ppto', value: fmtN(Tver), unit: 'FTE', sub: 'Presupuesto' },
    { label: 'Desviación', value: (dif > 0 ? '+' : '') + fmtN(dif), unit: 'FTE', color: A.kpiHex(dif > 0 ? 'rojo' : 'azul'),
      sub: Tver ? (dif > 0 ? '+' : '') + A.fmtPct(dif / Math.abs(Tver), 1) + ' vs ppto' : '—' },
    { label: 'Cumplimiento', value: A.fmtPct(cump, 0), color: A.kpiHex(cumpCol), sub: 'Real / Ppto' },
  ];

  // Filas de la tabla: VP (subtotal) → Gerencia, con búsqueda y expandir/colapsar.
  const ql = q.trim().toLowerCase();
  const expandAll = () => setExpanded(new Set(tree.vpNodes.map(n => n.name)));
  const collapseAll = () => setExpanded(new Set());
  const cap = { fontFamily: 'var(--font-disp)', fontWeight: 700, fontSize: 9.5, letterSpacing: '.07em', color: '#8a9499', textTransform: 'uppercase', marginBottom: 4 };
  const anyFilter = !!(vps.length || gers.length || q.trim());
  const clearFilters = () => { setVps([]); setGers([]); setQ(''); };
  // Píldora de "Gerencia filtrada" (mismo look que la selección de fila de la tabla principal).
  const selPill = { borderRadius: 4, padding: '1px 6px', margin: '0 -6px', background: 'var(--teal-100)', color: 'var(--amsa-teal-deep)', fontWeight: 700, boxShadow: 'inset 0 0 0 1px var(--teal-border)' };
  // Bloque Real/Ppto/Dif/Cumpl. con las mismas clases que la tabla principal (mtable).
  const cellN = (real, ver) => {
    const d = real - ver, col = A.kpiHex(A.kpiColor(real, ver, thr));
    return [<td key="r" className="real tnum">{fmtN(real)}</td>,
            <td key="p" className="ver tnum">{fmtN(ver)}</td>,
            <td key="d" className="dif tnum" style={{ color: d < 0 ? 'var(--ok)' : d > 0 ? 'var(--red)' : 'var(--fg-soft)' }}>{(d > 0 ? '+' : '') + fmtN(d)}</td>,
            <td key="c" className="pct" style={{ color: col, fontWeight: 700 }}>{ver ? A.fmtPct(real / ver, 0) : '—'}</td>];
  };

  return (
    <div>
      <div style={{ margin: '4px 0 12px' }}>
        <h2 style={{ fontFamily: 'var(--font-disp)', fontWeight: 800, fontSize: 18, color: 'var(--ink)', margin: 0 }}>Dotaciones (FTE)</h2>
        <div style={{ fontSize: 12, color: 'var(--fg-muted)', marginTop: 2 }}>Dotación promedio anual (Nº) · Real vs Presupuesto · {modeLbl}</div>
      </div>

      <div className="filters" style={{ flexWrap: 'wrap', gap: 14, alignItems: 'flex-end' }}>
        <div className="fgroup" style={{ minWidth: 220 }}>
          <div style={cap}>Datos</div>
          <div style={{ display: 'flex', gap: 8 }}>
            <Chip label="Propios" on={showProp} onClick={() => setShowProp(v => !v)} />
            <Chip label="Contratista" on={showCont} onClick={() => setShowCont(v => !v)} />
          </div>
        </div>
        <div className="fgroup" style={{ minWidth: 140 }}>
          <div style={cap}>Año</div>
          <div className="fctl"><MS options={yearOpts} selected={years} onChange={setYears} placeholder="Año" /></div>
        </div>
        <div className="fgroup grow" style={{ minWidth: 180 }}>
          <div style={cap}>Vicepresidencia</div>
          <div className="fctl"><MS options={vpOpts} selected={vps} searchable placeholder="Todas"
            onChange={v => { setVps(v); setGers(g => g.filter(x => !v.length || baseRecs.some(r => v.includes(r.vp) && r.ger === x))); }} /></div>
        </div>
        <div className="fgroup grow" style={{ minWidth: 180 }}>
          <div style={cap}>Gerencia</div>
          <div className="fctl"><MS options={gerOpts} selected={gers} searchable placeholder="Todas" onChange={setGers} /></div>
        </div>
      </div>

      <KpiCards kpis={kpis} unit={unit} />

      <div className="grid3" style={{ gridTemplateColumns: '1.5fr 1fr', marginTop: 12 }}>
        <div className="panel">
          <h3>Dotación Real vs Presupuesto por año</h3>
          <div className="ph-sub">Histórico anual (FTE)</div>
          <BarYears series={series} unit={unit} decimals={dec} showProp={false} />
          <div className="legend">
            <span><i style={{ background: 'var(--amsa-teal-deep)' }}></i>Real</span>
            <span><i style={{ background: 'var(--amsa-teal-light)' }}></i>Presupuesto</span>
          </div>
        </div>
        <div className="panel">
          <h3>Cumplimiento del presupuesto</h3>
          <div className="ph-sub">Real / Presupuesto por año</div>
          <ComplianceBars series={series} thr={thr} unit={unit} />
        </div>
      </div>

      <div className="matrix-card" style={{ marginTop: 12 }}>
        <div className="matrix-top">
          <div><h3>Dotación {yLbl}</h3><div className="mt-sub">{modeLbl} · FTE promedio</div></div>
          <div className="toolbar">
            <input type="text" value={q} onChange={e => setQ(e.target.value)} placeholder="Buscar VP / Gerencia…"
              style={{ height: 30, width: 220, padding: '0 10px', border: '1px solid var(--teal-200)', borderRadius: 6, fontSize: 12, fontFamily: 'var(--font-sans)', color: 'var(--ink)', outline: 'none' }} />
            <button type="button" className="btn" onClick={() => expanded.size ? collapseAll() : expandAll()}>
              {expanded.size ? '⤒ Colapsar todo' : '⤓ Expandir todo'}
            </button>
          </div>
        </div>
        <div style={{ overflowX: 'auto', maxHeight: 520, overflowY: 'auto' }}>
          <table className="mtable">
            <thead>
              <tr className="grp">
                <th className="nameblk" rowSpan="2">VP / Gerencia
                  {anyFilter && (
                    <span role="button" title="Quitar la búsqueda y los filtros de VP / Gerencia" onClick={clearFilters}
                      style={{ marginLeft: 8, padding: '1px 7px', borderRadius: 4, background: 'rgba(255,255,255,.18)', color: '#fff',
                        fontSize: 10.5, fontWeight: 600, cursor: 'pointer', textTransform: 'none', letterSpacing: 'normal', whiteSpace: 'nowrap' }}>
                      Limpiar ×
                    </span>
                  )}
                </th>
                <th className="spacer" rowSpan="2"></th>
                <th colSpan="4">FTE promedio · {yLbl}</th>
              </tr>
              <tr className="cols">
                <th>Real</th><th className="ver">Ppto</th><th>Dif</th><th>Cumpl.</th>
              </tr>
            </thead>
            <tbody>
              {tree.vpNodes.map(vp => {
                const matchVP = !ql || vp.name.toLowerCase().includes(ql);
                const kids = (vp.children || []).filter(g => matchVP || g.name.toLowerCase().includes(ql));
                if (ql && !matchVP && kids.length === 0) return null;
                const open = expanded.has(vp.name) || (ql && kids.length > 0);
                return (
                  <React.Fragment key={vp.name}>
                    <tr className="row-vp" style={{ cursor: 'pointer' }}
                      onClick={() => setExpanded(p => { const n = new Set(p); n.has(vp.name) ? n.delete(vp.name) : n.add(vp.name); return n; })}>
                      <td className="name">
                        <span className="twig ind-1">
                          <button className="tog" type="button">{open ? '–' : '+'}</button>
                          <span>{vp.name}</span>
                        </span>
                      </td>
                      <td className="spacer"></td>
                      {cellN(avgR(vp.agg), avgV(vp.agg))}
                    </tr>
                    {open && kids.map(g => {
                      const selG = gers.includes(g.name);
                      return (
                      <tr key={vp.name + '|' + g.name} className="row-ger" style={{ cursor: 'pointer' }}
                        title="Clic para filtrar por esta Gerencia (toda la pestaña)"
                        onClick={() => setGers(p => p.includes(g.name) ? p.filter(x => x !== g.name) : [...p, g.name])}>
                        <td className="name">
                          <span className="twig ind-2">
                            <span className="tog empty"></span>
                            <span style={selG ? selPill : undefined}>{g.name}</span>
                          </span>
                        </td>
                        <td className="spacer"></td>
                        {cellN(avgR(g.agg), avgV(g.agg))}
                      </tr>
                      );
                    })}
                  </React.Fragment>
                );
              })}
              <tr className="row-total">
                <td className="name">Total {modeLbl}</td>
                <td className="spacer"></td>
                <td className="real tnum">{fmtN(Treal)}</td>
                <td className="ver tnum">{fmtN(Tver)}</td>
                <td className="dif tnum">{(dif > 0 ? '+' : '') + fmtN(dif)}</td>
                <td className="pct">{Tver ? A.fmtPct(cump, 0) : '—'}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


function App() {
  const A = window.CORP;
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
    items: [], itemrels: [], vps: [], gers: [], companies: [], companias: [], cecos: [], clacos: [], tcs: [],
    hidden: { claco: ['6125101'] },  // oculto por defecto: CLACO 6125101 "Fletes por venta cobre" (solo aparece en VP Comercialización) → no se considera por defecto; restaurable con "⊘ restaurar"
    clases: defaultClases(), // por defecto: todas menos Mano de Obra
    stMode: 'excl',     // Services & Tech: por defecto SIN S&T ('excl'). '' todos · 'only' solo · 'excl' sin
    aps: [],
    version: 'ORI',
    years: [2025],      // años reales seleccionados (2022..2025)
    showProp: false,    // Presupuesto 2027 como capa de comparación aditiva
    donutMetric: 'real', // métrica del donut: 'real' | 'prop'
    yearAgg: 'sum',     // tabla con varios años: 'sum' (acumulado) | 'avg' (promedio)
    q: '',              // búsqueda en la tabla (VP / Gerencia / Ítem)
    groupMode: 'org',   // 'org' = VP›Ger›Ítem (foto) · 'item' = Ítem›VP›Ger
    sort: { key: 'real', dir: 'desc' }, // orden de la tabla por columna
  });
  const set = useCallback(patch => setSt(s => ({ ...s, ...patch })), []);

  const [expanded, setExpanded] = useState(() => new Set());
  // Presupuesto 2027 editada: persiste en este equipo (localStorage) para no perderla al recargar.
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
  // Renombre POR CECO de Gerencia y Desc. CECO (editable en el Diccionario). Muta rec.ger /
  // rec.dceco solo del CECO editado → el cambio afecta únicamente a ese CECO. Persiste.
  const [cecoNames, setCecoNames] = useState(() => {
    try { const s = localStorage.getItem('ada_ceconames_v1'); const o = s ? JSON.parse(s) : {}; return { ger: o.ger || {}, dceco: o.dceco || {} }; } catch (e) { return { ger: {}, dceco: {} }; }
  });
  useEffect(() => { try { localStorage.setItem('ada_ceconames_v1', JSON.stringify(cecoNames)); } catch (e) {} }, [cecoNames]);
  useMemo(() => A.applyCecoNameOverrides(cecoNames), [cecoNames]);
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
  // Toggles globales del modelo: valor ('n' Normal | 'a' Ajustada 2027) · CECOS ('new'|'old').
  const [valMode, setValModeS] = useState('n');
  const [cecoMode, setCecoModeS] = useState('new');   // Estructura CECOS por defecto: Nueva
  const onValMode = useCallback(m => { A.setValMode(m); A.applyVpOverrides(vpov); A.applyCecoNameOverrides(cecoNames); setValModeS(m); }, [vpov, cecoNames]);
  const onCecoMode = useCallback(m => { A.setCecoMode(m); A.applyVpOverrides(vpov); A.applyCecoNameOverrides(cecoNames); setCecoModeS(m); }, [vpov, cecoNames]);

  const showProp = !!st.showProp;
  const histYears = st.years.filter(y => typeof y === 'number').sort((a, b) => a - b);
  const hasFY = st.years.includes('2026fy'); // 2026 Ppto FY (solo Ppto)
  const hasFcst = st.years.includes('2026fcst'); // 2026 Forecast 5+7 (suma al Real)
  const labelParts = [...histYears.map(String), ...(hasFY ? ['2026 Ppto FY'] : []), ...(hasFcst ? ['2026 Forecast 5+7'] : [])];
  const yearsLabel = labelParts.length ? labelParts.join(' + ') : '—';
  const multiYear = histYears.length > 1; // con >1 año se puede elegir acumulado/promedio en la tabla
  // Con más de 2 años, las tarjetas no listan los años (no caben) → "N años".
  const periodLbl = histYears.length > 2 ? `${histYears.length} años` : yearsLabel;
  const thr = { red: t.thrRed, yellow: t.thrYellow };
  // Sin simulación de 2026/2027 (no hay datos). 'ori' ajusta el plan histórico en versión Original.
  const growth = { g26: 0, gReal26: 0, g27: 0, ori: 0.955 };

  // Mismas agrupaciones que la Tabla Resumen (incluyen CECO; el detalle cuelga bajo Contrapartida).
  const groupDims = st.groupMode === 'item' ? RESUMEN_DIMS.itemcc : RESUMEN_DIMS.orgcc;

  const opts = useMemo(() => ({
    years: st.years, showProp, donutMetric: st.donutMetric, yearAgg: st.yearAgg, version: st.version, dataMode: st.dataMode,
    companies: st.companies, vps: st.vps, gers: st.gers, itemrels: st.itemrels, items: st.items, cecos: st.cecos, clacos: st.clacos, tcs: st.tcs, clases: st.clases, aps: st.aps, st: st.stMode, hidden: st.hidden,
    groupBy: groupDims, sort: st.sort, overrides, growth,
  }), [st, showProp, overrides, vpov, cecoNames, valMode, cecoMode]);

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
    const idn = x => x;
    const D = { vp: A.dispVP, ger: A.dispGer, itemrel: A.dispItemRel, item: A.dispItem, ceco: idn, contra: idn };
    const d0 = D[groupDims[0]] || idn, d1 = D[groupDims[1]] || idn, d2 = D[groupDims[2]] || idn;
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
    const FK = { vp: 'vps', ger: 'gers', itemrel: 'itemrels', item: 'items' };
    const dim = dims[row.level - 1];
    const fkey = FK[dim];
    const val = row.node.name;
    if (!fkey) { setSt(s => ({ ...s, q: s.q === val ? '' : val })); return; } // contra/ceco → buscador
    setSt(s => {
      const cur = s[fkey] || [];
      if (cur.length === 1 && cur[0] === val) return { ...s, [fkey]: [] }; // toggle off
      const patch = { [fkey]: [val] };
      dims.slice(dims.indexOf(dim) + 1).forEach(d => { if (FK[d]) patch[FK[d]] = []; });
      return { ...s, ...patch };
    });
  }, []);
  // Limpiar todos los filtros de categoría (VP / Gerencia / Ítem) a la vez.
  const onClearFilter = useCallback(() => setSt(s => ({ ...s, vps: [], gers: [], itemrels: [], items: [], tcs: [], clases: [], aps: [] })), []);
  const expandAll = () => {
    const n = new Set();
    const walk = (nodes, prefix) => nodes.forEach(nd => {
      const key = prefix ? prefix + '|' + nd.name : nd.name;
      if (nd.children && nd.children.length) { n.add(key); walk(nd.children, key); }
    });
    walk(tree.vpNodes, '');
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
    // Distribuye el nuevo total entre los recIds del ítem, proporcional al Ppto 2027
    // cargado (rec.prop27); si la base es 0, reparte en partes iguales.
    setOverrides(prev => {
      const n = { ...prev };
      const ids = node.recIds || [];
      if (!ids.length) return n;
      const base = ids.map(id => (A.recordById(id).prop27 || 0));
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
    // Con muchas tarjetas (Efecto Moneda + Propuesta) los montos no caben con 3 decimales:
    // se reducen decimales automáticamente (7+ tarjetas → 0 decimales; 6 → máx 1).
    const em = valMode === 'a' ? A.efectoMoneda(opts) : null;
    const nCards = 5 + (em ? ((em.realN ? 1 : 0) + (em.pptoN ? 1 : 0)) : 0) + (showProp ? 1 : 0);
    const kdec = nCards >= 7 ? 0 : nCards >= 6 ? Math.min(dec, 1) : dec;
    const cards = [
      { label: avg ? 'Real promedio anual' : TX.kpi.realAnual, value: A.fmt(realShown, unit, kdec), unit, sub: avg ? `promedio de ${periodLbl}` : `${yearsLabel} · acumulado anual` },
      { label: avg ? 'Presupuesto promedio' : TX.kpi.presupuestoAnual, value: A.fmt(verShown, unit, kdec), unit, sub: A.VERSIONES.find(v => v.id === st.version).label },
      { label: avg ? 'Desviación promedio' : TX.kpi.desviacionAnual, value: (difShown > 0 ? '+' : '') + A.fmt(difShown, unit, kdec), unit, color: difShown < 0 ? 'var(--ok)' : difShown > 0 ? 'var(--red)' : 'var(--fg-soft)', trend: (pct > 0 ? '+' : '') + A.fmtPct(pct, 1), trendDir: pct > 0 ? 'up' : 'down', sub: pct > 0 ? 'sobre presupuesto' : 'bajo presupuesto' },
      { label: TX.kpi.cumplimiento, value: A.fmtPct(cump, 0), color: A.kpiHex(cumpCol), sub: 'Real / Presupuesto anual' },
      { label: TX.kpi.alarmas, value: String(alarmas), color: A.kpiHex(alarmas ? 'rojo' : 'azul'), sub: `de ${tree.vpNodes.length} VP sobre umbral`, tooltip: alarmas ? 'VP sobre presupuesto (>' + thr.red + '%): ' + alarmaVPs.join(', ') : 'Ninguna VP sobre el umbral' },
    ];
    // Efecto Moneda: con la base "Moneda Ajustada 2027" activa, se agregan dos tarjetas
    // (acento amarillo = atención) — una para el Real y otra para el Ppto — con la plata
    // extra al reexpresar a 2027 vs la moneda original y el % sobre la original. Mismo
    // universo filtrado. Se omite la tarjeta cuya base original sea 0 (ej. 2026 FY no
    // tiene Real → solo se muestra el Efecto sobre Ppto).
    if (valMode === 'a' && em) {
      const tip = base => 'Diferencia del ' + base + ' al reexpresar de moneda original a moneda ajustada 2027 (años seleccionados). % sobre la moneda original.';
      const mkCard = (label, orig, adj, sub, base) => ({
        label, value: ((adj - orig) >= 0 ? '+' : '') + A.fmt(avg ? (adj - orig) / nY : (adj - orig), unit, kdec), unit,
        status: 'amarillo', dot: 'amarillo',
        trend: (((adj - orig) / (orig || 1)) >= 0 ? '+' : '') + A.fmtPct(orig ? (adj - orig) / Math.abs(orig) : 0, 1),
        trendDir: (adj - orig) >= 0 ? 'up' : 'down', sub, tooltip: tip(base),
      });
      const extra = [];
      if (em.realN) extra.push(mkCard('Efecto Moneda Real (2027)', em.realN, em.realA, 'vs Real original', 'Real'));
      if (em.pptoN) extra.push(mkCard('Efecto Moneda Ppto (2027)', em.pptoN, em.pptoA, 'vs Ppto original', 'Presupuesto'));
      cards.splice(2, 0, ...extra);
    }
    // Capa Presupuesto 2027: tarjeta adicional comparando contra el promedio de los años seleccionados.
    if (showProp) {
      const prom = tree.nYears ? T.ytdReal / tree.nYears : 0;
      const prop = T.prop || 0;
      const dP = prom ? (prop - prom) / Math.abs(prom) : 0;
      cards.push({
        label: 'Presupuesto 2027', value: A.fmt(prop, unit, kdec), unit, status: 'amarillo', dot: 'amarillo',
        trend: (dP > 0 ? '+' : '') + A.fmtPct(dP, 1), trendDir: dP > 0 ? 'up' : 'down',
        sub: `vs promedio real`,
      });
    }
    return cards;
  }, [tree, showProp, unit, dec, thr, st.version, yearsLabel, valMode, opts]);

  /* ---------- export CSV ---------- */
  const exportExcel = () => {
    const div = unit === 'kUSD' ? 1e3 : unit === 'USD' ? 1 : 1e6;
    const f = Math.pow(10, dec);
    const num = v => ({ t: 'n', v: Math.round((v / div) * f) / f });
    const DLBL = { vp: 'Vicepresidencia', ger: 'Gerencia', dceco: 'Desc. CECO', itemrel: '\u00CDtem Relevante', item: '\u00CDtem', ceco: 'CECO', contra: 'Contrapartida' };
    const DISP = { vp: A.dispVP, ger: A.dispGer, dceco: n => n, itemrel: A.dispItemRel, item: A.dispItem, ceco: n => n, contra: n => n };
    const cols = groupDims.map(d => DLBL[d] || d);
    const nameCells = path => path.map((node, i) => ({ t: 's', v: (DISP[groupDims[i]] || (x => x))(node.name) }));
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
    if (showProp) headers = [...headers, `Presupuesto 2027 (${unit})`, `Δ vs Real %`];
    const rows = [];
    const walkExport = (nodes, path) => nodes.forEach(nd => {
      const p = path.concat(nd);
      if (nd.children && nd.children.length) { walkExport(nd.children, p); return; }
      const a = nd.agg, difu = a.ytdReal - a.ytdVersion;
      const pct = a.ytdVersion ? Math.round((difu / Math.abs(a.ytdVersion)) * 100) : 0;
      let row;
      if (byYear) {
        row = [...nameCells(p)];
        yrs.forEach(y => { const yv = (a.yr && a.yr[y]) || { real: 0, ver: 0 }; row.push(num(yv.real), num(yv.ver)); });
      } else {
        row = [...nameCells(p), num(a.ytdReal * fAgg), num(a.ytdVersion * fAgg), num(difu * fAgg), { t: 'n', v: pct }];
      }
      if (showProp) {
        const prom = nY ? a.ytdReal / nY : 0;
        const dpct = prom ? Math.round((((a.prop || 0) - prom) / Math.abs(prom)) * 100) : 0;
        row.push(num(a.prop || 0), { t: 'n', v: dpct });
      }
      rows.push(row);
    });
    walkExport(tree.vpNodes, []);
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
            <div className="sub">{TX.header.subtitulo}{(window.V2_DATA && window.V2_DATA.pptoTs) ? ' · Ppto 2027 actualizado el ' + window.V2_DATA.pptoTs : ''}</div>
          </div>
          <img className="hdr-logo" src={window.CORP_LOGO || 'assets/logo_amsa.png'} alt="Antofagasta Minerals" />
        </div>
      </div>
      <div className="accent"><i className="t"></i><i className="r"></i><i className="y"></i></div>

      <div className="app-wrap">
        <div style={{ display: 'flex', gap: 4, marginBottom: 12, borderBottom: '2px solid var(--teal-100)' }}>
          {[['dashboard', 'Gastos Corporativos'], ['resumen', 'Tabla Resumen Gastos'],
            ...((A.dotRecords && A.dotRecords.length) ? [['dotaciones', 'Dotaciones AMSA (FTE)']] : []),  // se oculta si no hay dotaciones (ej. dashboards por VP)
            ['dict', 'Diccionario (CECO · Ítem)']].map(([id, lbl]) => (
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
          <ColorPanel onApply={onColorApply} edit={editMode} />
        </div>
        {page === 'dict' ? <DictView vpov={vpov} setVpov={setVpov} nameov={nameov} setNameov={setNameov} cecoNames={cecoNames} setCecoNames={setCecoNames} cecoMode={cecoMode} onCecoMode={onCecoMode} />
         : page === 'dotaciones' ? <DotacionesView />
         : page === 'resumen' ? <ResumenView overrides={overrides} unit={unit} dec={dec} onDec={v => setTweak('decimals', v)} valMode={valMode} cecoMode={cecoMode} onValMode={onValMode} onCecoMode={onCecoMode} st={st} set={set} />
         : <React.Fragment>
        <FilterBar st={st} set={set} gerOptions={dims.gers} itemrelOptions={dims.itemrels} cecoOptions={dims.cecos} clacoOptions={dims.clacos} tcOptions={dims.tcs} clasOptions={dims.clases} apOptions={dims.aps}
          valMode={valMode} cecoMode={cecoMode} onValMode={onValMode} onCecoMode={onCecoMode} />

        <KpiCards kpis={kpis} unit={unit} />

        {t.showCharts && (() => {
          const barCard = (
            <React.Fragment>
              <h3>Real vs Presupuesto por año</h3>
              <div className="ph-sub">{showProp ? 'Histórico anual y Presupuesto 2027' : 'Histórico anual'} · {unit}</div>
              <BarYears series={series} unit={unit} decimals={dec} showProp={showProp} />
              <div className="legend">
                <span><i style={{ background: 'var(--amsa-teal-deep)' }}></i>Real histórico</span>
                <span><i style={{ background: 'var(--amsa-teal-light)' }}></i>Presupuesto</span>
                {showProp && <span><i style={{ background: 'var(--amsa-yellow)' }}></i>Presupuesto 27</span>}
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
                    <div className="ph-sub">{st.donutMetric === 'prop' ? 'Presupuesto 2027' : 'Real'} por compañía · {unit}</div>
                  </div>
                  {showProp && (
                    <div style={{ display: 'flex', gap: 2, background: '#fff', overflow: 'hidden',
                      border: '1px solid var(--teal-200)', borderRadius: 4, flex: '0 0 auto', marginTop: 2 }}>
                      {[['real', 'Real'], ['prop', 'Presupuesto 27']].map(([m, lbl]) => (
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
                  <option value="org">VP › Ítem Relevante</option>
                  <option value="item">Ítem Relevante › VP</option>
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
                    <button type="button" aria-label="Más decimales" onClick={() => setTweak('decimals', Math.min(6, dec + 1))}
                      style={{ border: 0, borderBottom: '1px solid var(--teal-100)', background: 'var(--teal-wash2)', color: 'var(--teal-muted)', cursor: 'pointer', padding: '0 7px', fontSize: 7, lineHeight: '12px', flex: 1 }}>▲</button>
                    <button type="button" aria-label="Menos decimales" onClick={() => setTweak('decimals', Math.max(0, dec - 1))}
                      style={{ border: 0, background: 'var(--teal-wash2)', color: 'var(--teal-muted)', cursor: 'pointer', padding: '0 7px', fontSize: 7, lineHeight: '12px', flex: 1 }}>▼</button>
                  </span>
                </span>
              </span>
              <button type="button" onClick={expandAll}
                style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: 'var(--teal-wash)', color: 'var(--amsa-teal)', border: '1px solid var(--amsa-teal-light)', borderRadius: 7, padding: '6px 13px', fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: 'var(--font-sans)', whiteSpace: 'nowrap', marginRight: 6 }}>
                ⤓ Expandir todo
              </button>
              <button type="button" onClick={collapseAll}
                style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: 'var(--teal-wash)', color: 'var(--amsa-teal)', border: '1px solid var(--amsa-teal-light)', borderRadius: 7, padding: '6px 13px', fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: 'var(--font-sans)', whiteSpace: 'nowrap', marginRight: 6 }}>
                ⤒ Colapsar todo
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
              onRowFilter={onRowFilter} filterSel={{ vps: st.vps, gers: st.gers, itemrels: st.itemrels, items: st.items }} onClearFilter={onClearFilter} stMode={st.stMode} />
          </div>
        </div>

        <div className="note">
          <b>KPI:</b> verde ≤ presupuesto · amarillo hasta {thr.red}% · rojo &gt; {thr.red}% sobre presupuesto · montos en {unit}{showProp ? (companiesActive ? ' · Presupuesto 2027 no editable con filtro de compañía' : ' · Presupuesto 2027 editable a nivel de ítem') : ''}.<br />
          <b>Fuente:</b> Consulta a SAP BPC de valores históricos 2022-2025. <b>Presupuesto 2027:</b> Pendiente.
        </div>
        </React.Fragment>}
      </div>

      <TweaksPanel>
        <TweakSection label="Presentación" />
        <TweakRadio label="Unidad" value={t.unit} options={['MUSD', 'kUSD', 'USD']} onChange={v => setTweak('unit', v)} />
        <TweakRadio label="Densidad" value={t.density} options={['compact', 'regular', 'comfy']} onChange={v => setTweak('density', v)} />
        <TweakToggle label="Mostrar gráficos" value={t.showCharts} onChange={v => setTweak('showCharts', v)} />
        <TweakSection label="Regla KPI (% sobre presupuesto)" />
        <TweakSlider label="Umbral rojo" value={t.thrRed} min={3} max={30} step={1} unit="%" onChange={v => setTweak('thrRed', v)} />
        <TweakSlider label="Umbral amarillo" value={t.thrYellow} min={1} max={t.thrRed - 1} step={1} unit="%" onChange={v => setTweak('thrYellow', v)} />
      </TweaksPanel>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
