/* global React */

/* Construye filas visibles a partir del árbol y el set de expandidos. */
function parseLocaleNum(s) {
  let str = String(s).replace(/[^0-9.,-]/g, '');
  if (str.indexOf(',') !== -1) { str = str.replace(/\./g, '').replace(',', '.'); }
  return parseFloat(str);
}
const DISP_BY = { vp: 'dispVP', ger: 'dispGer', item: 'dispItem' };
const DIM_LBL = { vp: 'Vicepresidencia', ger: 'Gerencia', item: 'Ítem Relevante' };

function flattenTree(tree, expanded, q, dims) {
  const A = window.CORP;
  const d0 = A[DISP_BY[dims[0]]], d1 = A[DISP_BY[dims[1]]], d2 = A[DISP_BY[dims[2]]];
  const ql = (q || '').trim().toLowerCase();
  const has = (s) => s.toLowerCase().includes(ql);
  const rows = [];
  tree.vpNodes.forEach(n1 => {
    const k1 = n1.name;
    // Con búsqueda: incluir nivel 1 si coincide su nombre, o si tiene hijos/nietos que coinciden.
    let subs;
    if (!ql || has(d0(n1.name))) {
      subs = n1.children.map(g => ({ n2: g, items: g.children }));
    } else {
      subs = [];
      n1.children.forEach(g => {
        if (has(d1(g.name))) subs.push({ n2: g, items: g.children });
        else {
          const its = g.children.filter(it => has(d2(it.name)));
          if (its.length) subs.push({ n2: g, items: its });
        }
      });
      if (subs.length === 0) return;
    }
    rows.push({ type: 'vp', key: k1, node: n1, level: 1, expandable: n1.children.length > 0 });
    if (!expanded.has(k1)) return;
    subs.forEach(({ n2, items }) => {
      const k2 = k1 + '|' + n2.name;
      rows.push({ type: 'ger', key: k2, node: n2, level: 2, expandable: n2.children.length > 0 });
      if (expanded.has(k2)) {
        items.forEach(it => rows.push({ type: 'item', key: k2 + '|' + it.name, node: it, level: 3, expandable: false }));
      }
    });
  });
  return rows;
}

function Twig({ row, expanded, onToggle, dims, onPick, active }) {
  const A = window.CORP;
  const name = A[DISP_BY[dims[row.level - 1]]](row.node.name);
  const isOpen = expanded.has(row.key);
  return (
    <span className={'twig ind-' + row.level}>
      {row.expandable
        ? <button className="tog" onClick={() => onToggle(row.key)}>{isOpen ? '–' : '+'}</button>
        : <span className="tog empty"></span>}
      <span
        onClick={onPick ? () => onPick(row) : undefined}
        title={onPick ? 'Filtrar todo el panel por este elemento (clic de nuevo para quitar)' : undefined}
        style={onPick ? {
          cursor: 'pointer', borderRadius: 4, padding: '1px 6px', margin: '0 -6px',
          background: active ? 'var(--teal-100)' : 'transparent',
          color: active ? 'var(--amsa-teal-deep)' : undefined,
          fontWeight: active ? 700 : undefined,
          boxShadow: active ? 'inset 0 0 0 1px var(--teal-border)' : undefined,
        } : undefined}
      >{name}</span>
    </span>
  );
}

/* ---------------- Bloque de 5 columnas (Real/Versión/Dif/%Dif/KPI) ---------------- */
function execCells(real, version, unit, thr, A, keyPrefix, dec) {
  // Fila "vacía": Real y Ppto ambos se muestran como 0 al detalle de decimales actual.
  const div = unit === 'kUSD' ? 1e3 : 1e6;
  const zeroThr = 0.5 * div * Math.pow(10, -(dec || 0));
  const realZero = Math.abs(real) < zeroThr;
  const verZero = Math.abs(version) < zeroThr;
  const empty = realZero && verZero;          // ambos 0 → fila sin actividad
  const dif = real - version;
  // Real con presupuesto 0 → desviación 100% (signo del real); resto = dif/ppto.
  const pct = empty ? null : (verZero ? (real > 0 ? 1 : -1) : dif / Math.abs(version));
  // Bajo presupuesto (Dif < 0) = verde; sobre presupuesto (Dif > 0) = rojo; vacío = negro.
  const difColor = empty ? 'var(--ink)' : (dif < 0 ? 'var(--ok)' : dif > 0 ? 'var(--red)' : 'var(--fg-soft)');
  return [
    <td className="real tnum" key={keyPrefix + 'r'}>{A.fmt(real, unit, dec)}</td>,
    <td className="ver tnum" key={keyPrefix + 'v'}>{A.fmt(version, unit, dec)}</td>,
    <td className="dif tnum" style={{ color: difColor }} key={keyPrefix + 'd'}>{empty ? A.fmt(0, unit, dec) : (dif > 0 ? '+' : '') + A.fmt(dif, unit, dec)}</td>,
    <td className="pct" style={{ color: difColor }} key={keyPrefix + 'p'}>{pct == null ? '—' : (pct > 0 ? '+' : '') + A.fmtPct(pct, 0)}</td>,
    <td className="kpi" key={keyPrefix + 'k'}>{empty ? null : <span className="dot" style={{ background: A.kpiHex(A.kpiColor(real, version, thr)) }}></span>}</td>,
  ];
}

/* ---------------- Bloque Propuesta 2027 (2 columnas: valor + Δ vs Real) ---------------- */
// prom = promedio anual Real de los años seleccionados = ytdReal / nYears.
function propCells(node, editable, nYears, unit, dec, A, keyPrefix, onEditProp, companiesActive, yearsLabel) {
  const a = node.agg;
  const prop = a.prop || 0;
  const prom = nYears ? a.ytdReal / nYears : 0;
  const d = prop - prom;
  const pct = prom ? d / Math.abs(prom) : null;
  const div = unit === 'kUSD' ? 1e3 : 1e6;
  // El ámbar identifica la cifra de presupuesto propuesto (misma identidad que la barra "Propuesta 27").
  const tint = 'rgba(240,169,41,.09)';   // ámbar muy tenue para la columna del valor 2027
  const sep = '2px solid var(--accent-300)';
  return [
    <td className="tnum" style={{ background: tint, borderLeft: sep, color: 'var(--ink)', fontWeight: 600 }} key={keyPrefix + 'prop'}>
      {editable && !companiesActive
        ? <input className="prop-inp" type="text" defaultValue={(prop / div).toFixed(unit === 'kUSD' ? 0 : 2).replace('.', ',')}
            key={node.recIds.join(',') + ':' + yearsLabel}
            onBlur={e => { const val = parseLocaleNum(e.target.value); if (!isNaN(val)) onEditProp(node, val * div); }}
            onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); }} />
        : A.fmt(prop, unit, dec)}
    </td>,
    // Δ informativo (propuesta vs promedio histórico): neutro, sin semáforo verde/rojo.
    <td className="pct tnum" style={{ color: 'var(--fg-3)' }} key={keyPrefix + 'dp'}>{pct == null ? '—' : (pct > 0 ? '+' : '') + A.fmtPct(pct, 0)}</td>,
  ];
}

/* ---------------- Encabezado de columna con menú (ordenar/fijar/ocultar) ---------------- */
function ColHeader({ col, sort, onSortDir, onPin, onHide, pinned, style, className, align, rowSpan, extra }) {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef(null);
  React.useEffect(() => {
    if (!open) return;
    const f = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', f);
    return () => document.removeEventListener('mousedown', f);
  }, [open]);
  const arr = (sort && col.sortKey && sort.key === col.sortKey) ? (sort.dir === 'asc' ? ' ▲' : ' ▼') : '';
  const mi = { display: 'block', width: '100%', textAlign: 'left', border: 0, background: 'transparent', padding: '7px 14px', fontSize: 12, cursor: 'pointer', color: 'var(--ink)', whiteSpace: 'nowrap', fontFamily: 'var(--font-sans)', fontWeight: 500 };
  const st = { ...style, position: (style && style.position) || 'relative', cursor: 'pointer', userSelect: 'none' };
  return (
    <th className={className} style={st} ref={ref} rowSpan={rowSpan}>
      <span onClick={() => setOpen(o => !o)} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: align || 'flex-end', width: '100%' }}>
        {col.label}{arr}<span style={{ fontSize: 8, opacity: 0.55 }}>▾</span>
        {extra}
      </span>
      {open && (
        <div style={{ position: 'absolute', top: '100%', right: 0, marginTop: 2, background: '#fff', border: '1px solid var(--teal-200)', borderRadius: 6, boxShadow: '0 6px 20px rgba(0,0,0,.16)', zIndex: 80, minWidth: 170, padding: '4px 0', textTransform: 'none', letterSpacing: 'normal' }}>
          {col.sortKey && <button style={mi} onClick={() => { onSortDir(col.sortKey, 'desc'); setOpen(false); }}>↓&nbsp; Mayor a menor</button>}
          {col.sortKey && <button style={mi} onClick={() => { onSortDir(col.sortKey, 'asc'); setOpen(false); }}>↑&nbsp; Menor a mayor</button>}
          <div style={{ height: 1, background: 'var(--line-soft)', margin: '4px 0' }}></div>
          {col.pinnable !== false && <button style={mi} onClick={() => { onPin(col.key); setOpen(false); }}>{pinned ? '◳ Liberar columna' : '◰ Fijar columna'}</button>}
          {col.hideable !== false && <button style={mi} onClick={() => { onHide(col.key); setOpen(false); }}>⊘ Ocultar columna</button>}
        </div>
      )}
    </th>
  );
}

/* ============================ MATRIX ============================ */
function Matrix(props) {
  const A = window.CORP;
  const { tree, showProp, nYears, yearAgg, yearsLabel, unit, thr, decimals, versionLabel, q, sort, onSort, onSortDir,
    hiddenCols, pinnedCols, onHide, onPin, expanded, onToggle, onEditProp, companiesActive, onRowFilter, filterSel, onClearFilter } = props;
  // "Limpiar ×" en el encabezado de la columna de jerarquía: aparece solo si hay
  // alguna categoría filtrada (VP/Gerencia/Ítem), por clic en la tabla o por los
  // desplegables. stopPropagation para no abrir el menú de ordenar de la columna.
  const anyFilter = !!(filterSel && ((filterSel.vps && filterSel.vps.length) || (filterSel.gers && filterSel.gers.length) || (filterSel.items && filterSel.items.length)));
  const clearEl = (anyFilter && onClearFilter) ? (
    <span role="button" title="Quitar el filtro de categoría" onClick={(e) => { e.stopPropagation(); onClearFilter(); }}
      style={{ marginLeft: 8, padding: '1px 7px', borderRadius: 4, background: 'rgba(255,255,255,.18)', color: '#fff',
        fontSize: 10.5, fontWeight: 600, cursor: 'pointer', textTransform: 'none', letterSpacing: 'normal', whiteSpace: 'nowrap', flex: '0 0 auto' }}>
      Limpiar ×
    </span>
  ) : null;
  const dims = tree.dims || ['vp', 'ger', 'item'];
  const headerLbl = dims.map(d => DIM_LBL[d]).join(' / ');
  const rows = flattenTree(tree, expanded, q, dims);
  // Una fila está "activa" como filtro global si su dimensión está filtrada
  // exactamente por su valor (selección única).
  const FILT_KEY = { vp: 'vps', ger: 'gers', item: 'items' };
  const onPick = onRowFilter ? (row) => onRowFilter(row, dims) : undefined;
  const rowActive = (row) => {
    const sel = (filterSel && filterSel[FILT_KEY[dims[row.level - 1]]]) || [];
    return sel.length === 1 && sel[0] === row.node.name;
  };
  const fAgg = (yearAgg === 'avg' && nYears > 1) ? 1 / nYears : 1;
  const aggLbl = fAgg !== 1 ? 'Promedio anual' : 'Acumulado anual';
  const byYear = !!tree.byYear;
  const yrs = tree.years || [];

  /* ---------- Modo "Por año": Real + Ppto por año (sin menú de columnas) ---------- */
  if (byYear) {
    const arrowY = (k) => (sort && sort.key === k) ? (sort.dir === 'asc' ? ' ▲' : ' ▼') : '';
    const shY = (k) => ({ onClick: () => onSort && onSort(k), style: { cursor: 'pointer', userSelect: 'none' } });
    const pcY = (node, editable, kp) => propCells(node, editable, nYears, unit, decimals, A, kp, onEditProp, companiesActive, yearsLabel);
    const yearCells = (agg, kp) => yrs.flatMap((y, i) => {
      const yv = (agg.yr && agg.yr[y]) || { real: 0, ver: 0 };
      const sep = i ? { borderLeft: '1px solid var(--teal-200)' } : undefined;
      return [
        <td className="real tnum" style={sep} key={kp + y + 'r'}>{A.fmt(yv.real, unit, decimals)}</td>,
        <td className="ver tnum" key={kp + y + 'p'}>{A.fmt(yv.ver, unit, decimals)}</td>,
      ];
    });
    return (
      <table className="mtable">
        <thead>
          <tr className="grp">
            <th className="nameblk" rowSpan="2" {...shY('name')}>{headerLbl}{arrowY('name')}{clearEl}</th>
            <th className="spacer" rowSpan="2"></th>
            {yrs.map((y, i) => <th key={y} colSpan="2" style={i ? { borderLeft: '1px solid var(--teal-200)' } : undefined}>{y} · {unit}</th>)}
            {showProp && <th colSpan="2" style={{ borderLeft: '2px solid var(--accent-300)' }}>Propuesta 2027 · {unit}</th>}
          </tr>
          <tr className="cols">
            {yrs.flatMap((y, i) => [
              <th key={y + 'r'} style={i ? { borderLeft: '1px solid var(--teal-200)' } : undefined}>Real</th>,
              <th className="ver" key={y + 'p'}>Ppto</th>,
            ])}
            {showProp && <th style={{ background: 'var(--amsa-yellow)', color: 'var(--accent-900)', borderLeft: '2px solid var(--accent-300)' }} {...shY('prop')}>2027{arrowY('prop')}</th>}
            {showProp && <th {...shY('dprop')}>Δ vs Real{arrowY('dprop')}</th>}
          </tr>
        </thead>
        <tbody>
          {rows.map(row => (
            <tr key={row.key} className={'row-' + row.type}>
              <td className="name"><Twig row={row} expanded={expanded} onToggle={onToggle} dims={dims} onPick={onPick} active={rowActive(row)} /></td>
              <td className="spacer"></td>
              {yearCells(row.node.agg, row.key + ':')}
              {showProp && pcY(row.node, row.type === 'item', row.key + ':')}
            </tr>
          ))}
          <tr className="row-total">
            <td className="name">Total</td>
            <td className="spacer"></td>
            {yearCells(tree.total, 't:')}
            {showProp && pcY({ agg: tree.total, recIds: [] }, false, 't:')}
          </tr>
        </tbody>
      </table>
    );
  }

  /* ---------- Modo estándar: columnas con menú (ordenar/fijar/ocultar) ---------- */
  const hidden = new Set(hiddenCols || []);
  const pinnedAll = (pinnedCols || []).filter(k => !hidden.has(k));
  // Definición de columnas (clave estable, etiqueta, clave de orden).
  let cols = [
    { key: 'real', label: 'Real', sortKey: 'real' },
    { key: 'version', label: versionLabel, sortKey: 'version', cls: 'ver' },
    { key: 'dif', label: 'Dif', sortKey: 'dif' },
    { key: 'pct', label: '% Dif', sortKey: 'pct' },
    { key: 'kpi', label: 'KPI', sortKey: null, cls: 'kpi', align: 'center' },
  ];
  if (showProp) cols.push(
    { key: 'prop', label: '2027', sortKey: 'prop', amber: true },
    { key: 'dprop', label: 'Δ vs Real', sortKey: 'dprop' },
  );
  const visible = cols.filter(c => !hidden.has(c.key));
  const pinnedMetrics = pinnedAll.filter(k => k !== 'name' && visible.some(c => c.key === k));
  const pinActive = pinnedAll.includes('name') || pinnedMetrics.length > 0;
  // Orden: columnas fijadas primero, luego el resto.
  const ordered = [...pinnedMetrics.map(k => visible.find(c => c.key === k)), ...visible.filter(c => !pinnedMetrics.includes(c.key))];

  // Sticky: anchos fijos para las columnas congeladas (offsets deterministas).
  const NAMEW = 300, SPW = 14, PINW = 116;
  const pinIdx = k => pinnedMetrics.indexOf(k);
  const pinSt = (key, header) => {
    const i = pinIdx(key);
    if (i < 0) return null;
    return { position: 'sticky', left: NAMEW + SPW + i * PINW, width: PINW, minWidth: PINW, maxWidth: PINW, zIndex: header ? 5 : 2, background: header ? undefined : '#fff' };
  };
  const nameSt = (header) => pinActive
    ? { position: 'sticky', left: 0, width: NAMEW, minWidth: NAMEW, maxWidth: NAMEW, zIndex: header ? 6 : 3, overflow: 'hidden', textOverflow: 'ellipsis', background: header ? undefined : '#fff' }
    : undefined;
  const spSt = pinActive ? { position: 'sticky', left: NAMEW, zIndex: 3, background: '#fff' } : undefined;

  const ctxFor = (node, type) => {
    const agg = node.agg;
    const real = agg.ytdReal * fAgg, ver = agg.ytdVersion * fAgg;
    const div = unit === 'kUSD' ? 1e3 : 1e6, zt = 0.5 * div * Math.pow(10, -(decimals || 0));
    const empty = Math.abs(real) < zt && Math.abs(ver) < zt, verZero = Math.abs(ver) < zt, dif = real - ver;
    const pct = empty ? null : (verZero ? (real > 0 ? 1 : -1) : dif / Math.abs(ver));
    const difColor = empty ? 'var(--ink)' : (dif < 0 ? 'var(--ok)' : dif > 0 ? 'var(--red)' : 'var(--fg-soft)');
    const prom = nYears ? agg.ytdReal / nYears : 0, prop = agg.prop || 0;
    const dprop = prom ? (prop - prom) / Math.abs(prom) : null;
    return { node, type, real, ver, dif, pct, difColor, empty, prom, prop, dprop };
  };
  const cell = (col, c, isTotal) => {
    const ps = pinSt(col.key, false);
    const merge = (s) => ps ? { ...s, ...ps } : s;
    switch (col.key) {
      case 'real': return <td key="real" className="real tnum" style={ps || undefined}>{A.fmt(c.real, unit, decimals)}</td>;
      case 'version': return <td key="version" className="ver tnum" style={ps || undefined}>{A.fmt(c.ver, unit, decimals)}</td>;
      case 'dif': return <td key="dif" className="dif tnum" style={merge({ color: c.difColor })}>{c.empty ? A.fmt(0, unit, decimals) : (c.dif > 0 ? '+' : '') + A.fmt(c.dif, unit, decimals)}</td>;
      case 'pct': return <td key="pct" className="pct" style={merge({ color: c.difColor })}>{c.pct == null ? '—' : (c.pct > 0 ? '+' : '') + A.fmtPct(c.pct, 0)}</td>;
      case 'kpi': return <td key="kpi" className="kpi" style={ps || undefined}>{c.empty ? null : <span className="dot" style={{ background: A.kpiHex(A.kpiColor(c.real, c.ver, thr)) }}></span>}</td>;
      case 'prop': {
        const editable = c.type === 'item' && !isTotal, div = unit === 'kUSD' ? 1e3 : 1e6;
        return <td key="prop" className="tnum" style={merge({ background: ps ? '#fff' : 'rgba(240,169,41,.09)', borderLeft: '2px solid var(--accent-300)', color: 'var(--ink)', fontWeight: 600 })}>
          {editable && !companiesActive
            ? <input className="prop-inp" type="text" defaultValue={(c.prop / div).toFixed(unit === 'kUSD' ? 0 : 2).replace('.', ',')}
                key={c.node.recIds.join(',') + ':' + yearsLabel}
                onBlur={e => { const v = parseLocaleNum(e.target.value); if (!isNaN(v)) onEditProp(c.node, v * div); }}
                onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); }} />
            : A.fmt(c.prop, unit, decimals)}
        </td>;
      }
      case 'dprop': return <td key="dprop" className="pct tnum" style={merge({ color: 'var(--fg-3)' })}>{c.dprop == null ? '—' : (c.dprop > 0 ? '+' : '') + A.fmtPct(c.dprop, 0)}</td>;
      default: return null;
    }
  };

  const nameCol = { key: 'name', label: headerLbl, sortKey: 'name', pinnable: true, hideable: false };
  const hp = { sort, onSortDir, onPin, onHide };
  // Encabezado de grupo (solo sin columnas fijadas, para conservar el diseño agrupado).
  const stdVisible = visible.filter(c => c.key !== 'prop' && c.key !== 'dprop');
  const propVisible = visible.filter(c => c.key === 'prop' || c.key === 'dprop');

  return (
    <table className="mtable">
      <thead>
        {!pinActive ? (
          <React.Fragment>
            <tr className="grp">
              <ColHeader col={nameCol} {...hp} pinned={false} className="nameblk" align="flex-start" rowSpan="2" style={{ verticalAlign: 'bottom' }} extra={clearEl} />
              <th className="spacer" rowSpan="2"></th>
              {stdVisible.length > 0 && <th colSpan={stdVisible.length}>{aggLbl} · {yearsLabel} · {unit}</th>}
              {propVisible.length > 0 && <th colSpan={propVisible.length} style={{ borderLeft: '2px solid var(--accent-300)' }}>Propuesta 2027 · {unit}</th>}
            </tr>
            <tr className="cols">
              {visible.map(c => (
                <ColHeader key={c.key} col={c} {...hp} pinned={false}
                  className={(c.cls || '') + (c.amber ? '' : '')}
                  align={c.align}
                  style={c.amber ? { background: 'var(--amsa-yellow)', color: 'var(--accent-900)', borderLeft: '2px solid var(--accent-300)' } : undefined} />
              ))}
            </tr>
          </React.Fragment>
        ) : (
          <tr className="cols">
            <ColHeader col={nameCol} {...hp} pinned={pinnedAll.includes('name')} className="nameblk" align="flex-start" style={nameSt(true)} extra={clearEl} />
            <th className="spacer" style={spSt}></th>
            {ordered.map(c => (
              <ColHeader key={c.key} col={c} {...hp} pinned={pinnedMetrics.includes(c.key)}
                className={c.cls || ''} align={c.align}
                style={{ ...(c.amber ? { background: 'var(--amsa-yellow)', color: 'var(--accent-900)' } : {}), ...(pinSt(c.key, true) || {}) }} />
            ))}
          </tr>
        )}
      </thead>
      <tbody>
        {rows.map(row => {
          const c = ctxFor(row.node, row.type);
          return (
            <tr key={row.key} className={'row-' + row.type}>
              <td className="name" style={nameSt(false)}><Twig row={row} expanded={expanded} onToggle={onToggle} dims={dims} onPick={onPick} active={rowActive(row)} /></td>
              <td className="spacer" style={spSt}></td>
              {ordered.map(col => cell(col, c, false))}
            </tr>
          );
        })}
        <tr className="row-total">
          <td className="name" style={nameSt(false)}>Total</td>
          <td className="spacer" style={spSt}></td>
          {ordered.map(col => cell(col, ctxFor({ agg: tree.total, recIds: [] }, 'total'), true))}
        </tr>
      </tbody>
    </table>
  );
}

Object.assign(window, { Matrix, flattenTree });
