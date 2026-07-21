/* global React */

/* Construye filas visibles a partir del árbol y el set de expandidos. */
function parseLocaleNum(s) {
  let str = String(s).replace(/[^0-9.,-]/g, '');
  if (str.indexOf(',') !== -1) { str = str.replace(/\./g, '').replace(',', '.'); }
  return parseFloat(str);
}
const DISP_BY = { vp: 'dispVP', ger: 'dispGer', dceco: 'dispDceco', itemrel: 'dispItemRel', item: 'dispItem', ceco: 'dispCeco', contra: 'dispContra', claco: 'dispClaco' };
const DIM_LBL = { vp: 'Vicepresidencia', ger: 'Gerencia', dceco: 'Desc. CECO', itemrel: 'Ítem Relevante', item: 'Ítem', ceco: 'CECO', contra: 'Contrapartida', texto: 'Texto pedido', denom: 'Denominación', concepto: 'Concepto Gasto', actividad: 'Actividad' };
// Familia de cada dimensión de detalle (para atenuar columnas que no aplican): REAL vs PPTO.
const DIM_FAM = { contra: 'real', texto: 'real', denom: 'real', concepto: 'ppto', actividad: 'ppto' };
const DET_DIMS = ['texto', 'denom'];   // niveles de detalle (v3) que cuelgan de Contrapartida

// Aplana el árbol a filas visibles. Soporta N niveles (VP › Gerencia › Ítem › CECO…),
// respetando el set de expandidos. type = por PROFUNDIDAD (para el estilo por nivel).
function flattenTree(tree, expanded, q, dims, detail, detOrder, collapsed, stMode, detP, sortCmp, catNodes, realCol) {
  const A = window.CORP;
  const D = dims.map(d => A[DISP_BY[d]] || (x => x));
  const ql = (q || '').trim().toLowerCase();
  // Estilo por PROFUNDIDAD (jerarquía visual): 1º nivel = destacado (row-vp),
  // 2º = medio (row-ger), resto = normal (row-item). Independiente de qué dimensión sea.
  const clsFor = d => d === 0 ? 'vp' : d === 1 ? 'ger' : 'item';
  const useDet = !!(detail && A.hasDetail && A.hasDetail());   // v3: detalle Real bajo Contrapartida
  const useDetP = !!(detP && A.hasDetailP && A.hasDetailP());  // v3: detalle Ppto/Fcst (Concepto Gasto › Actividad)
  // Ancla del detalle Ppto = el ÚLTIMO nivel de la estructura que NO sea Contrapartida
  // (Concepto Gasto › Actividad cuelga de ahí, acotado a ese nivel, sea Ítem, CECO, Ítem
  // Relevante, Desc. CECO, Gerencia o VP — lo que el usuario haya dejado como último).
  const ppAnchor = useDetP ? dims.reduce((a, d, i) => (d !== 'contra' ? i : a), -1) : -1;
  const years = tree.years || [];
  const rows = [];
  // TOPE de filas: una búsqueda amplia (o expandir todo) puede revelar decenas de miles
  // de filas de detalle y congelar el navegador. Cortamos y avisamos (rows.truncated).
  const MAX_ROWS = 4000;
  const nameHit = (n, depth) => { const disp = D[depth] || (x => x); return String(disp(n.name)).toLowerCase().includes(ql); };
  // ¿Este nodo o alguno de sus descendientes (incl. el DETALLE: texto/denom/documento) coincide?
  const lastDim = dims[dims.length - 1];
  const subtreeMatch = (n, depth, ctx) => {
    if (nameHit(n, depth)) return true;
    const dim = dims[depth];
    // El detalle (texto/denom/doc) solo cuelga cuando Contrapartida es el ÚLTIMO nivel.
    if (dim === 'contra' && depth === dims.length - 1) return useDet && A.detailMatch ? A.detailMatch(Object.assign({}, ctx, { contra: n.name }), ql, stMode) : false;
    // El detalle Ppto/Fcst (Concepto Gasto › Actividad) cuelga del nivel más profundo entre Ítem y CECO.
    if (depth === ppAnchor && useDetP && A.detailMatchP && A.detailMatchP(Object.assign({}, ctx, { [dim]: n.name }), ql)) return true;
    if (n.children && n.children.length) {
      const c2 = Object.assign({}, ctx); c2[dim] = n.name;
      return n.children.some(c => subtreeMatch(c, depth + 1, c2));
    }
    return false;
  };
  // Niveles de detalle Real (Texto pedido › Denominación) generados al vuelo desde window.DET.
  function walkDetail(nodes, depth, prefix) {
    if (rows.length >= MAX_ROWS) return;
    if (sortCmp) nodes = nodes.slice().sort(sortCmp);   // ordenar el detalle por la columna activa
    nodes.forEach(n => {
      if (rows.length >= MAX_ROWS) return;
      const key = prefix + '|' + n.name + (n.doc ? '#' + n.doc : '');
      const expandable = !!(n.children && n.children.length);
      // Colapsar manual GANA sobre el auto-abrir de la búsqueda (permite esconder ramas al buscar).
      const open = (collapsed && collapsed.has(key)) ? false : (ql ? true : expanded.has(key));
      rows.push({ type: clsFor(depth), key, node: n, level: depth + 1, expandable, dim: n._dim, doc: n.doc, open, fam: 'real' });
      if (expandable && open) walkDetail(n.children, depth + 1, key);
    });
  }
  // Niveles de detalle Ppto/Fcst (Concepto Gasto › Actividad) desde window.DETP.
  function walkDetailP(nodes, depth, prefix) {
    if (rows.length >= MAX_ROWS) return;
    if (sortCmp) nodes = nodes.slice().sort(sortCmp);   // ordenar el detalle por la columna activa
    nodes.forEach(n => {
      if (rows.length >= MAX_ROWS) return;
      const key = prefix + '|P|' + n.name;
      const expandable = !!(n.children && n.children.length);
      const open = (collapsed && collapsed.has(key)) ? false : (ql ? true : expanded.has(key));
      rows.push({ type: clsFor(depth), key, node: n, level: depth + 1, expandable, dim: n._dim, open, fam: 'ppto' });
      if (expandable && open) walkDetailP(n.children, depth + 1, key);
    });
  }
  // Contrapartidas (nivel Real) bajo el nodo «Real»: se toman del ÁRBOL (n.children del ancla), por lo
  // que SUS VALORES CUADRAN con el total real del padre. Bajo cada Contrapartida cuelga Texto pedido ›
  // Denominación (window.DET). El nodo «Real» las agrupa; nace colapsado y aparece si hay columna Real.
  function walkContra(nodes, ctx, depth, prefix, force) {
    if (rows.length >= MAX_ROWS) return;
    if (sortCmp) nodes = nodes.slice().sort(sortCmp);
    nodes.forEach(n => {
      if (rows.length >= MAX_ROWS) return;
      if (n.name === '(Presupuesto)' || n.name === '(Forecast)' || n.name === '(Outlook)' || n.name === '(Ppto2027)') return;   // el Ppto/Fcst no tiene contrapartida
      const self = ql ? String(n.name).toLowerCase().includes(ql) : true;
      const ctxC = Object.assign({}, ctx, { contra: n.name });
      if (ql && !force && !self && !(A.detailMatch && A.detailMatch(ctxC, ql, stMode))) return;
      const key = prefix + '|C|' + n.name;
      const hasDet = !!(A.hasDetailFor && A.hasDetailFor(ctxC, stMode));
      const open = (collapsed && collapsed.has(key)) ? false : (ql ? true : expanded.has(key));
      rows.push({ type: clsFor(depth), key, node: n, level: depth + 1, expandable: hasDet, dim: 'contra', open, fam: 'real' });
      if (hasDet && open) walkDetail(A.detailNodes(ctxC, years, (force || self) ? null : ql, detOrder, stMode), depth + 1, key);
    });
  }
  // Nodo-CATEGORÍA (Real / Ppto·Forecast) que agrupa el detalle bajo el último nivel de la
  // estructura: nace COLAPSADO (clic para abrir; el detalle cuelga recién al abrirlo). Su valor
  // reutiliza el agg del nodo padre, enmascarado por familia (columnas de la otra familia → «—»).
  function pushCat(fam, node, parentKey, depth, emit) {
    if (rows.length >= MAX_ROWS) return;
    const catKey = parentKey + (fam === 'real' ? '|@R' : '|@P');
    const open = (collapsed && collapsed.has(catKey)) ? false : (ql ? true : expanded.has(catKey));
    rows.push({ type: clsFor(depth), key: catKey, node: node, level: depth + 1, expandable: true, cat: fam, fam: fam, open });
    if (open) emit(catKey, depth + 1);
  }
  // force = un ancestro coincidió por nombre → mostrar TODO su subárbol.
  function walk(nodes, depth, prefix, ctx, force) {
    if (rows.length >= MAX_ROWS) return;
    nodes.forEach(n => {
      if (rows.length >= MAX_ROWS) return;
      const dim = dims[depth];
      // El Ppto/Forecast no tienen contrapartida: no mostramos esos nodos a nivel
      // Contrapartida (confunde). El monto igual cuenta en el nivel padre.
      if (dim === 'contra' && (n.name === '(Presupuesto)' || n.name === '(Forecast)' || n.name === '(Outlook)' || n.name === '(Ppto2027)')) return;
      const self = ql ? nameHit(n, depth) : true;
      if (ql && !force && !self && !subtreeMatch(n, depth, ctx)) return;
      const key = prefix ? prefix + '|' + n.name : n.name;
      const ctx2 = Object.assign({}, ctx); ctx2[dim] = n.name;
      const detHere = useDet && dim === 'contra' && depth === dims.length - 1;   // detalle Real (texto/denom) bajo Contrapartida cuando ES el último nivel (Gastos Corporativos)
      const detPHere = useDetP && depth === ppAnchor;                            // detalle Ppto/Fcst cuelga del nivel más profundo entre Ítem y CECO
      const hasKids = !!(n.children && n.children.length);
      // «Real» (Tabla Resumen): las Contrapartidas del ÁRBOL (hijas de este nodo) se AGRUPAN bajo un
      // nodo «Real» cuando el siguiente nivel es Contrapartida y hay columna Real. Reutiliza los nodos
      // del árbol → sus valores CUADRAN con el total real del padre. Simétrico con «Ppto / Forecast».
      const realHere = catNodes && realCol && dims[depth + 1] === 'contra' && hasKids;
      const expandable = detHere ? A.hasDetailFor(ctx2, stMode) : (hasKids || (detPHere && A.hasDetailPFor(ctx2)));
      const childForce = force || self;
      // Colapsar manual GANA sobre el auto-abrir de la búsqueda (permite esconder ramas al buscar).
      const open = (collapsed && collapsed.has(key)) ? false : (ql ? true : expanded.has(key));
      rows.push({ type: clsFor(depth), key, node: n, level: depth + 1, expandable, dim, open, fam: DIM_FAM[dim] });
      if (expandable && open) {
        if (detHere) {
          walkDetail(A.detailNodes(ctx2, years, childForce ? null : ql, detOrder, stMode), depth + 1, key);
        } else {
          if (realHere) {
            // Envolver las Contrapartidas (hijas del árbol) bajo el nodo «Real»; el detalle texto/denom cuelga de cada una.
            const emitReal = (pfx, dpt) => walkContra(n.children, ctx2, dpt, pfx, childForce);
            pushCat('real', n, key, depth + 1, emitReal);
          } else if (hasKids) {
            walk(n.children, depth + 1, key, ctx2, childForce);
          }
          // «Ppto / Forecast»: Concepto Gasto › Actividad.
          if (detPHere && A.hasDetailPFor(ctx2)) {
            const emitP = (pfx, dpt) => walkDetailP(A.detailNodesP(ctx2, childForce ? null : ql), dpt, pfx);
            if (catNodes) pushCat('ppto', n, key, depth + 1, emitP); else emitP(key, depth + 1);
          }
        }
      }
    });
  }
  walk(tree.vpNodes, 0, '', {}, false);
  if (rows.length >= MAX_ROWS) rows.truncated = MAX_ROWS;   // se cortó: la UI muestra aviso
  return rows;
}

function Twig({ row, expanded, onToggle, dims, onPick, active, onHide }) {
  const A = window.CORP;
  const dim = row.dim || dims[row.level - 1];
  const name = (A[DISP_BY[dim]] || (x => x))(row.node.name);
  // Código al lado cuando corresponde: Ítem → Cód_Agrupación2 (CLACO) · CECO → código CECO ·
  // Desc. CLACO → código CLACO (la clave del nivel ya ES el código).
  const code = (dim === 'item' || dim === 'itemrel') ? (A.itemCode ? A.itemCode(row.node.name) : null)
             : (dim === 'ceco' || dim === 'claco') ? row.node.name
             : (dim === 'denom' || dim === 'texto') ? ((row.doc && row.doc !== '—') ? 'Doc ' + row.doc : null)   // Documento compra (en la hoja del detalle)
             : null;
  const showCode = code && String(code) !== String(name);
  const isOpen = row.open != null ? row.open : expanded.has(row.key);
  // Nodo-categoría (Real / Ppto·Forecast): agrupa el detalle bajo el último nivel. Sin código ni acciones.
  if (row.cat) {
    const catLbl = row.cat === 'real' ? 'Real' : 'Ppto / Forecast';
    return (
      <span className={'twig ind-' + row.level}>
        {row.expandable
          ? <button className="tog" onClick={() => onToggle(row.key)}>{isOpen ? '–' : '+'}</button>
          : <span className="tog empty"></span>}
        <span style={{ fontFamily: 'var(--font-disp)', fontWeight: 700, fontSize: 10.5, letterSpacing: '.05em', textTransform: 'uppercase',
          color: row.cat === 'real' ? 'var(--amsa-teal)' : 'var(--accent-ink)' }}>{catLbl}</span>
      </span>
    );
  }
  return (
    <span className={'twig ind-' + row.level}>
      {row.expandable
        ? <button className="tog" onClick={() => onToggle(row.key)}>{isOpen ? '–' : '+'}</button>
        : <span className="tog empty"></span>}
      <span
        onClick={onPick ? () => onPick(row) : undefined}
        title={onPick ? name + ' — clic para filtrar el panel por este elemento (clic de nuevo para quitar)' : name}
        style={onPick ? {
          cursor: 'pointer', borderRadius: 4, padding: '1px 6px', margin: '0 -6px',
          background: active ? 'var(--teal-100)' : 'transparent',
          color: active ? 'var(--amsa-teal-deep)' : undefined,
          fontWeight: active ? 700 : undefined,
          boxShadow: active ? 'inset 0 0 0 1px var(--teal-border)' : undefined,
        } : undefined}
      >{name}</span>
      {showCode ? <span style={{ marginLeft: 7, fontSize: 10.5, color: 'var(--fg-muted)', fontFamily: 'ui-monospace, SFMono-Regular, Menlo, monospace', fontVariantNumeric: 'tabular-nums' }}>{code}</span> : null}
      {(onHide && ['vp', 'ger', 'dceco', 'ceco', 'itemrel', 'item', 'claco', 'contra'].includes(dim)) ? (
        <button className="rowhide" title={'Ocultar «' + name + '» de la tabla'}
          onClick={e => { e.stopPropagation(); onHide(dim, row.node.name); }}>⊘</button>
      ) : null}
    </span>
  );
}

/* ---------------- Bloque de 5 columnas (Real/Versión/Dif/%Dif/KPI) ---------------- */
function execCells(real, version, unit, thr, A, keyPrefix, dec) {
  // Fila "vacía": Real y Ppto ambos se muestran como 0 al detalle de decimales actual.
  const div = unit === 'kUSD' ? 1e3 : unit === 'USD' ? 1 : 1e6;
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

/* ---------------- Bloque Presupuesto 2027 (2 columnas: valor + Δ vs Real) ---------------- */
// prom = promedio anual Real de los años seleccionados = ytdReal / nYears.
function propCells(node, editable, nYears, unit, dec, A, keyPrefix, onEditProp, companiesActive, yearsLabel) {
  const a = node.agg;
  const prop = a.prop || 0;
  const prom = nYears ? a.ytdReal / nYears : 0;
  const d = prop - prom;
  const pct = prom ? d / Math.abs(prom) : null;
  const div = unit === 'kUSD' ? 1e3 : unit === 'USD' ? 1 : 1e6;
  // El ámbar identifica la cifra de presupuesto propuesto (misma identidad que la barra "Presupuesto 27").
  const tint = 'rgba(240,169,41,.09)';   // ámbar muy tenue para la columna del valor 2027
  const sep = '2px solid var(--accent-300)';
  return [
    <td className="tnum" style={{ background: tint, borderLeft: sep, color: 'var(--ink)', fontWeight: 600 }} key={keyPrefix + 'prop'}>
      {editable && !companiesActive
        ? <input className="prop-inp" type="text" defaultValue={(prop / div).toFixed((unit === 'kUSD' || unit === 'USD') ? 0 : 2).replace('.', ',')}
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
    hiddenCols, pinnedCols, onHide, onPin, expanded, onToggle, onEditProp, companiesActive, onRowFilter, filterSel, onClearFilter, stMode } = props;
  // "Limpiar ×" en el encabezado de la columna de jerarquía: aparece solo si hay
  // alguna categoría filtrada (VP/Gerencia/Ítem), por clic en la tabla o por los
  // desplegables. stopPropagation para no abrir el menú de ordenar de la columna.
  const anyFilter = !!(filterSel && ((filterSel.vps && filterSel.vps.length) || (filterSel.gers && filterSel.gers.length) || (filterSel.itemrels && filterSel.itemrels.length) || (filterSel.items && filterSel.items.length)));
  const clearEl = (anyFilter && onClearFilter) ? (
    <span role="button" title="Quitar el filtro de categoría" onClick={(e) => { e.stopPropagation(); onClearFilter(); }}
      style={{ marginLeft: 8, padding: '1px 7px', borderRadius: 4, background: 'rgba(255,255,255,.18)', color: '#fff',
        fontSize: 10.5, fontWeight: 600, cursor: 'pointer', textTransform: 'none', letterSpacing: 'normal', whiteSpace: 'nowrap', flex: '0 0 auto' }}>
      Limpiar ×
    </span>
  ) : null;
  const dims = tree.dims || ['vp', 'ger', 'item'];
  // v3: si hay detalle embebido y NO estamos en modo "Por año", cuelgan Texto pedido ›
  // Denominación bajo cada Contrapartida (lazy). El encabezado lo refleja.
  const detailOn = !tree.byYear && dims[dims.length - 1] === 'contra' && !!(A.hasDetail && A.hasDetail());
  const headerLbl = dims.map(d => DIM_LBL[d]).join(' / ') + (detailOn ? ' / ' + DET_DIMS.map(d => DIM_LBL[d]).join(' / ') : '');
  const rows = flattenTree(tree, expanded, q, dims, detailOn, undefined, undefined, stMode);
  // Una fila está "activa" como filtro global si su dimensión está filtrada
  // exactamente por su valor (selección única).
  const FILT_KEY = { vp: 'vps', ger: 'gers', itemrel: 'itemrels', item: 'items' };
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
    const yearCells = (agg, kp, noPpto) => yrs.flatMap((y, i) => {
      const yv = (agg.yr && agg.yr[y]) || { real: 0, ver: 0 };
      const sep = i ? { borderLeft: '1px solid var(--teal-200)' } : undefined;
      return [
        <td className="real tnum" style={sep} key={kp + y + 'r'}>{A.fmt(yv.real, unit, decimals)}</td>,
        <td className="ver tnum" style={noPpto ? { color: 'var(--fg-muted)' } : undefined} key={kp + y + 'p'}>{noPpto ? 'N/A' : A.fmt(yv.ver, unit, decimals)}</td>,
      ];
    });
    return (
      <table className="mtable">
        <thead>
          <tr className="grp">
            <th className="nameblk" rowSpan="2" {...shY('name')}>{headerLbl}{arrowY('name')}{clearEl}</th>
            <th className="spacer" rowSpan="2"></th>
            {yrs.map((y, i) => <th key={y} colSpan="2" style={i ? { borderLeft: '1px solid var(--teal-200)' } : undefined}>{y} · {unit}</th>)}
            {showProp && <th colSpan="2" style={{ borderLeft: '2px solid var(--accent-300)' }}>Presupuesto 2027 · {unit}</th>}
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
              {yearCells(row.node.agg, row.key + ':', row.dim === 'contra')}
              {showProp && pcY(row.node, row.dim === 'item', row.key + ':')}
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

  const ctxFor = (node, type, dim) => {
    const agg = node.agg;
    const real = agg.ytdReal * fAgg, ver = agg.ytdVersion * fAgg;
    const div = unit === 'kUSD' ? 1e3 : unit === 'USD' ? 1 : 1e6, zt = 0.5 * div * Math.pow(10, -(decimals || 0));
    const empty = Math.abs(real) < zt && Math.abs(ver) < zt, verZero = Math.abs(ver) < zt, dif = real - ver;
    const pct = empty ? null : (verZero ? (real > 0 ? 1 : -1) : dif / Math.abs(ver));
    const difColor = empty ? 'var(--ink)' : (dif < 0 ? 'var(--ok)' : dif > 0 ? 'var(--red)' : 'var(--fg-soft)');
    const prom = nYears ? agg.ytdReal / nYears : 0, prop = agg.prop || 0;
    const dprop = prom ? (prop - prom) / Math.abs(prom) : null;
    // El presupuesto NO se abre por contrapartida (solo existe a nivel Ítem/CECO). A ese
    // detalle el Ppto es "n/A": no se muestra 0 ni se calcula Dif/%/KPI (no tiene sentido).
    const noPpto = dim === 'contra' || dim === 'texto' || dim === 'denom';   // detalle = solo Real
    return { node, type, dim, real, ver, dif, pct, difColor, empty, prom, prop, dprop, noPpto };
  };
  const cell = (col, c, isTotal) => {
    const ps = pinSt(col.key, false);
    const merge = (s) => ps ? { ...s, ...ps } : s;
    switch (col.key) {
      case 'real': return <td key="real" className="real tnum" style={ps || undefined}>{A.fmt(c.real, unit, decimals)}</td>;
      case 'version': return <td key="version" className="ver tnum" style={c.noPpto ? merge({ color: 'var(--fg-muted)' }) : (ps || undefined)}>{c.noPpto ? 'N/A' : A.fmt(c.ver, unit, decimals)}</td>;
      case 'dif': return <td key="dif" className="dif tnum" style={c.noPpto ? merge({ color: 'var(--fg-muted)' }) : merge({ color: c.difColor })}>{c.noPpto ? 'N/A' : (c.empty ? A.fmt(0, unit, decimals) : (c.dif > 0 ? '+' : '') + A.fmt(c.dif, unit, decimals))}</td>;
      case 'pct': return <td key="pct" className="pct" style={c.noPpto ? merge({ color: 'var(--fg-muted)' }) : merge({ color: c.difColor })}>{c.noPpto ? 'N/A' : (c.pct == null ? '—' : (c.pct > 0 ? '+' : '') + A.fmtPct(c.pct, 0))}</td>;
      case 'kpi': return <td key="kpi" className="kpi" style={ps || undefined}>{(c.noPpto || c.empty) ? null : <span className="dot" style={{ background: A.kpiHex(A.kpiColor(c.real, c.ver, thr)) }}></span>}</td>;
      case 'prop': {
        const editable = c.dim === 'item' && !isTotal && !c.noPpto, div = unit === 'kUSD' ? 1e3 : unit === 'USD' ? 1 : 1e6;
        return <td key="prop" className="tnum" style={merge({ background: ps ? '#fff' : 'rgba(240,169,41,.09)', borderLeft: '2px solid var(--accent-300)', color: c.noPpto ? 'var(--fg-muted)' : 'var(--ink)', fontWeight: 600 })}>
          {c.noPpto ? 'N/A' : (editable && !companiesActive
            ? <input className="prop-inp" type="text" defaultValue={(c.prop / div).toFixed((unit === 'kUSD' || unit === 'USD') ? 0 : 2).replace('.', ',')}
                key={c.node.recIds.join(',') + ':' + yearsLabel}
                onBlur={e => { const v = parseLocaleNum(e.target.value); if (!isNaN(v)) onEditProp(c.node, v * div); }}
                onKeyDown={e => { if (e.key === 'Enter') e.target.blur(); }} />
            : A.fmt(c.prop, unit, decimals))}
        </td>;
      }
      case 'dprop': return <td key="dprop" className="pct tnum" style={merge({ color: 'var(--fg-3)' })}>{c.noPpto ? 'N/A' : (c.dprop == null ? '—' : (c.dprop > 0 ? '+' : '') + A.fmtPct(c.dprop, 0))}</td>;
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
              {propVisible.length > 0 && <th colSpan={propVisible.length} style={{ borderLeft: '2px solid var(--accent-300)' }}>Presupuesto 2027 · {unit}</th>}
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
          const c = ctxFor(row.node, row.type, row.dim);
          return (
            <tr key={row.key} className={'row-' + row.type}>
              <td className="name" style={nameSt(false)}><Twig row={row} expanded={expanded} onToggle={onToggle} dims={dims} onPick={onPick} active={rowActive(row)} /></td>
              <td className="spacer" style={spSt}></td>
              {ordered.map(col => cell(col, c, false))}
            </tr>
          );
        })}
        {rows.truncated && (
          <tr><td colSpan={20} style={{ padding: '8px 12px', fontSize: 12, color: 'var(--amsa-teal-deep)', background: 'var(--accent-wash)', fontWeight: 600 }}>
            Mostrando las primeras {rows.truncated.toLocaleString('es-CL')} filas. Afiná la búsqueda o expandí manualmente para ver el resto.
          </td></tr>
        )}
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
