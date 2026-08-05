/* global React */
const { useState: useStateF, useRef: useRefF, useEffect: useEffectF } = React;

// Ver app.jsx: en los dashboards de UNA compañía distribuible se ocultan Alcance y Compañía.
const SOLO_DIST_F = !!(window.V2_DATA && window.V2_DATA.soloDist);

/* ---------------- Multi-select dropdown ---------------- */
function MultiSelect({ options, selected, onChange, placeholder, gold, searchable, impliedAll }) {
  const [open, setOpen] = useStateF(false);
  const [q, setQ] = useStateF('');
  const ref = useRefF(null);
  useEffectF(() => {
    function onDoc(e) { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }
    document.addEventListener('mousedown', onDoc);
    return () => document.removeEventListener('mousedown', onDoc);
  }, []);
  const sel = new Set(selected);
  // Sin selección explícita, si llega impliedAll (p.ej. los CLACOs en alcance según los OTROS
  // filtros) esos se muestran TILDADOS: son "lo que se está mostrando". El primer destilde
  // materializa la lista (deja de ser "todos" implícito) para poder excluir de a uno; si se
  // vuelve a cubrir impliedAll completo, regresa a [] (todos implícito).
  const implied = (!selected.length && impliedAll && impliedAll.length) ? impliedAll : null;
  const checkedSet = implied ? new Set(implied) : sel;
  const toggle = (v) => {
    const n = new Set(checkedSet);
    n.has(v) ? n.delete(v) : n.add(v);
    if (impliedAll && impliedAll.length && n.size === impliedAll.length && impliedAll.every(x => n.has(x))) onChange([]);
    else onChange([...n]);
  };
  const filtered = options.filter(o => String(o.label == null ? '' : o.label).toLowerCase().includes(q.toLowerCase()));
  let label = placeholder;
  if (selected.length === 1) label = options.find(o => o.value === selected[0])?.label || placeholder;
  else if (selected.length > 1) label = `${selected.length} seleccionadas`;
  return (
    <div className="ms-wrap" ref={ref}>
      <button className="ms-btn" onClick={() => setOpen(o => !o)} title={label}>
        <span className={selected.length ? '' : 'ph'} style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</span>
      </button>
      {open && (
        // El popup es una columna flex: buscador fijo arriba y acciones
        // (Todas/Limpiar) fijas abajo; SOLO la lista de opciones hace scroll,
        // así esas acciones quedan siempre visibles (no solo al final).
        <div className="ms-pop" style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          {searchable && <input className="ms-search" placeholder="Buscar…" value={q} onChange={e => setQ(e.target.value)} autoFocus style={{ flex: '0 0 auto' }} />}
          <div style={{ flex: '1 1 auto', minHeight: 0, overflowY: 'auto' }}>
            {filtered.map(o => (
              <label className="ms-opt" key={o.value}>
                <input type="checkbox" checked={checkedSet.has(o.value)} onChange={() => toggle(o.value)} />
                <span>{o.label}</span>
              </label>
            ))}
            {filtered.length === 0 && <div style={{ padding: 8, fontSize: 12, color: 'var(--fg-soft)' }}>Sin resultados</div>}
          </div>
          <div className="ms-actions" style={{ flex: '0 0 auto' }}>
            <button onClick={() => onChange(options.map(o => o.value))}>Todas</button>
            <button onClick={() => onChange([])}>Limpiar</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- Chip tipo píldora con check (Datos) ---------------- */
function ModeChip({ label, on, onClick }) {
  return (
    <button type="button" onClick={onClick} aria-pressed={on}
      style={{
        border: on ? '1px solid var(--amsa-teal)' : '1px solid var(--teal-border)', margin: 0,
        padding: '7px 12px', cursor: 'pointer', borderRadius: 7,
        appearance: 'none', WebkitAppearance: 'none', MozAppearance: 'none',
        boxSizing: 'border-box', outline: 'none', boxShadow: 'none',
        display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 7,
        backgroundColor: on ? 'var(--amsa-teal)' : '#ffffff',
        color: on ? '#ffffff' : 'var(--fg-soft)',
        fontFamily: 'var(--font-sans)', fontWeight: 600, fontSize: 12, lineHeight: 1.2,
        whiteSpace: 'nowrap',
      }}>
      <span style={{
        width: 13, height: 13, borderRadius: 3, flex: '0 0 auto',
        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
        fontSize: 10, lineHeight: 1, fontWeight: 800, color: '#fff',
        border: on ? 'none' : '1.5px solid var(--teal-border)',
        background: on ? 'rgba(255,255,255,.3)' : 'transparent',
      }}>{on ? '✓' : ''}</span>
      {label}
    </button>
  );
}

/* ---------------- Toggle deslizante (Comparar) ---------------- */
function SwitchToggle({ label, on, onClick, color }) {
  const c = color || 'var(--amsa-teal)';
  return (
    <button type="button" onClick={onClick} aria-pressed={on}
      style={{
        display: 'inline-flex', alignItems: 'center', gap: 9, cursor: 'pointer',
        border: '1px solid ' + (on ? c : 'var(--teal-border)'), background: '#fff', borderRadius: 7,
        padding: '6px 12px', outline: 'none', boxShadow: 'none', boxSizing: 'border-box',
        fontFamily: 'var(--font-sans)', fontWeight: 600, fontSize: 12,
        color: on ? 'var(--accent-900)' : 'var(--fg-soft)', whiteSpace: 'nowrap',
      }}>
      <span style={{
        width: 32, height: 17, borderRadius: 9, background: on ? c : 'var(--teal-border)',
        position: 'relative', flex: '0 0 auto', transition: 'background .15s',
      }}>
        <span style={{
          position: 'absolute', top: 2, left: on ? 17 : 2, width: 13, height: 13,
          borderRadius: '50%', background: '#fff', transition: 'left .15s',
          boxShadow: '0 1px 2px rgba(0,0,0,.3)',
        }}></span>
      </span>
      {label}
    </button>
  );
}

/* ---------------- Filter bar ---------------- */
function FilterBar(props) {
  const A = window.CORP;
  const { st, set, gerOptions, itemrelOptions, cecoOptions, clacoOptions, tcOptions, clasOptions, apOptions } = props;
  // Toggles globales (valor Normal/Ajustada 2027 · CECOS nuevos/antiguos).
  const valMode = props.valMode || 'n', cecoMode = props.cecoMode || 'new';
  const onValMode = props.onValMode || (() => {}), onCecoMode = props.onCecoMode || (() => {});
  const hideComercial = props.hideComercial !== false, onHideComercial = props.onHideComercial || (() => {});
  const hideFletes = props.hideFletes !== false, onHideFletes = props.onHideFletes || (() => {});
  const acTodoTC = props.acTodoTC !== false, onAcTodoTC = props.onAcTodoTC || (() => {});
  // Corte YTD 2026 en vivo (hasta qué mes suman el Real y el Ppto YTD).
  const ytdMes = props.ytdMes || 5, onYtdMes = props.onYtdMes || (() => {}), ytdMesMax = props.ytdMesMax || 12;
  const MESES_YTD_F = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  // Opciones de Ítem Relevante (Agrupación3) y Gerencia según el modo de datos activo.
  const itemrelVals = itemrelOptions || [];
  const itemrelOpts = itemrelVals.map(v => ({ value: v, label: A.dispItemRel(v) }));
  // Tipo Costo (C1/C2/C3) y ¿Aplica? (Sí/No). OJO: el Tipo Costo del filtro sale del CLACO
  // (Cód_Agrupación2 del CLACO), no del CECO.
  const tcOpts = (tcOptions || []).map(v => ({ value: v, label: v }));
  const cecoOpts = [...new Set([...(cecoOptions || []), ...(st.cecos || [])])].sort((a, b) => a.localeCompare(b, 'es')).map(v => ({ value: v, label: v }));
  const clacoOpts = [...new Set([...(clacoOptions || []), ...(st.clacos || [])])].sort((a, b) => a.localeCompare(b, 'es')).map(v => ({ value: v, label: v }));
  const clasOpts = (clasOptions || []).map(v => ({ value: v, label: v }));
  const apOpts = (apOptions || []).map(v => ({ value: v, label: v }));
  // VP/Gerencia SOLO de los registros de GASTO (corp/dist según el toggle Datos); NO de
  // Dotaciones (que usan otra convención "Vicepresidencia X" y duplicarían la lista).
  const gastoRecs = A.activeRecords({ dataMode: st.dataMode || 'both' });
  const vpOpts = [...new Set(gastoRecs.map(r => r.vp))]
    .sort((a, b) => A.dispVP(a).localeCompare(A.dispVP(b), 'es'))
    .map(v => ({ value: v, label: A.dispVP(v) }));
  const gamsaOpts = [...new Set(gastoRecs.map(r => r.gamsa))].filter(Boolean)
    .sort((a, b) => a.localeCompare(b, 'es')).map(v => ({ value: v, label: v }));   // Grupo AMSA
  const gerVals = gerOptions
    || (st.vps.length ? [...new Set(gastoRecs.filter(r => st.vps.includes(r.vp)).map(r => r.ger))] : A.D.gers);
  const gerOpts = gerVals.map(v => ({ value: v, label: A.dispGer(v) }));
  const compOpts = A.COMPANIAS.map(c => ({ value: c.id, label: `${c.id} · ${c.nombre}` }));
  const mode = st.dataMode || 'both';
  const corpOn = mode === 'corp' || mode === 'both';
  const distOn = mode === 'dist' || mode === 'both';
  const distMode = distOn;
  // "Código CLACO": si hay algún otro filtro activo, se tildan los CLACO que se están mostrando
  // (mismo criterio que en la Tabla Resumen). Aquí las compañías se filtran vía `companies`
  // (por código, como en esta vista), no vía `companias`.
  const clacoNarrowed = !!((st.vps && st.vps.length) || (st.gers && st.gers.length) || (st.itemrels && st.itemrels.length) || (st.items && st.items.length) || (st.cecos && st.cecos.length) || (st.tcs && st.tcs.length) || (st.aps && st.aps.length) || (st.companies && st.companies.length));
  const clacoScope = React.useMemo(() => !clacoNarrowed ? [] : A.clacosInScope({
    dataMode: mode, companies: st.companies || [], vps: st.vps || [], gers: st.gers || [], itemrels: st.itemrels || [],
    items: st.items || [], tcs: st.tcs || [], clases: st.clases || [], cecos: st.cecos || [],
    st: st.stMode || '', aps: st.aps || [], hidden: st.hidden || {},
  }), [clacoNarrowed, mode, st.companies, st.vps, st.gers, st.itemrels, st.items, st.tcs, st.clases, st.cecos, st.stMode, st.aps, st.hidden]);
  const setFlags = (c, d) => {
    const dataMode = c && d ? 'both' : c ? 'corp' : d ? 'dist' : 'none';
    // Al cambiar de modo, conservar las selecciones de Gerencia/Ítem que sigan
    // siendo válidas en el nuevo conjunto de datos (no borrar todo el filtro).
    const dims = A.dimsFor({ ...st, dataMode });
    const gset = new Set(dims.gers), iset = new Set(dims.items), irset = new Set(dims.itemrels);
    set({ dataMode, gers: st.gers.filter(g => gset.has(g)),
      items: st.items.filter(i => iset.has(i)),
      itemrels: (st.itemrels || []).filter(i => irset.has(i)) });
  };

  // Año: multi-selección de años reales (2022-2025). La Presupuesto 2027 ya no vive
  // aquí: es un switch aparte (capa de comparación aditiva).
  const yearOpts = [
    { value: 2022, label: '2022' }, { value: 2023, label: '2023' },
    { value: 2024, label: '2024' }, { value: 2025, label: '2025' },
    { value: 2026, label: '2026 YTD' },     // acumulado 2026 hasta el mes de corte (Real vs Ppto)
    { value: '2026fy', label: '2026 Ppto FY' }, // presupuesto anual 2026 (solo Ppto, sin Real)
    { value: '2026fcst', label: '2026 Forecast 5+7' }, // forecast anual 2026 (suma al Real)
  ];
  const onYears = (v) => set({ years: v }); // años reales; permite vacío (Limpiar)
  const showProp = !!st.showProp;

  // Tipo Costo y ¿Aplica? (filtros secundarios) van SIEMPRE dentro del botón "+".
  const moreRef = useRefF(null);
  const [moreOpen, setMoreOpen] = useStateF(false);
  useEffectF(() => {
    if (!moreOpen) return;
    const f = e => { if (moreRef.current && !moreRef.current.contains(e.target)) setMoreOpen(false); };
    document.addEventListener('mousedown', f);
    return () => document.removeEventListener('mousedown', f);
  }, [moreOpen]);
  const extras = [
    <div className="fgroup" style={{ minWidth: 150 }} key="tc">
      <div className="fcap">Tipo Costo (CLACO)</div>
      <div className="fctl">
        <MultiSelect options={tcOpts} selected={st.tcs || []} onChange={v => set({ tcs: v })} placeholder="Todos" />
      </div>
      <label style={{ display: 'flex', alignItems: 'center', gap: 5, marginTop: 5, fontSize: 12, color: 'var(--fg-soft)', cursor: 'pointer' }}
        title="Marcado (por defecto): VP Asuntos Corporativos muestra TODOS los Tipo Costo (C1/C2/C3), aunque el filtro esté acotado (p.ej. solo C1) para el resto.">
        <input type="checkbox" checked={acTodoTC} onChange={e => onAcTodoTC(e.target.checked)} /> Asuntos Corp.: todos los TC
      </label>
    </div>,
    <div className="fgroup" style={{ minWidth: 160 }} key="clas">
      <div className="fcap">Clasificación Cuenta</div>
      <div className="fctl">
        <MultiSelect options={clasOpts} selected={st.clases || []} onChange={v => set({ clases: v })} placeholder="Todas" />
      </div>
    </div>,
    (A.hasST && A.hasST()) ? (
    <div className="fgroup" style={{ minWidth: 150 }} key="st">
      <div className="fcap">Services &amp; Tech</div>
      <div className="fctl">
        <select value={st.stMode || ''} onChange={e => set({ stMode: e.target.value })} style={{ minWidth: 0, width: '100%' }}
          title="CLACOs 6125020/6125021. Solo Ppto y Forecast los distinguen; el Real no (queda 0).">
          <option value="">Todos</option>
          <option value="only">Solo Services &amp; Tech</option>
          <option value="excl">Sin Services &amp; Tech</option>
        </select>
      </div>
    </div>) : null,
    <div className="fgroup" style={{ minWidth: 150 }} key="ceco">
      <div className="fcap">Código CECO</div>
      <div className="fctl">
        <MultiSelect options={cecoOpts} selected={st.cecos || []} onChange={v => set({ cecos: v })} placeholder="Todos" searchable />
      </div>
    </div>,
    <div className="fgroup" style={{ minWidth: 150 }} key="claco">
      <div className="fcap">Código CLACO</div>
      <div className="fctl">
        <MultiSelect options={clacoOpts} selected={st.clacos || []} onChange={v => set({ clacos: v })} placeholder="Todos" searchable
          impliedAll={clacoNarrowed ? clacoScope : undefined} />
      </div>
    </div>,
    <div className="fgroup" style={{ minWidth: 150 }} key="ap">
      <div className="fcap">¿Aplica?</div>
      <div className="fctl">
        <MultiSelect options={apOpts} selected={st.aps || []} onChange={v => set({ aps: v })} placeholder="Todas" />
      </div>
    </div>,
  ];
  const nMore = ((st.tcs && st.tcs.length) ? 1 : 0) + ((st.clases && st.clases.length) ? 1 : 0) + ((st.aps && st.aps.length) ? 1 : 0) + ((st.cecos && st.cecos.length) ? 1 : 0) + ((st.clacos && st.clacos.length) ? 1 : 0) + (st.stMode ? 1 : 0);

  // Colapso dinámico: los filtros que no caben en una fila se mueven al "+".
  const [nHidden, setNHidden] = useStateF(0);
  const barRef = useRefF(null);
  const collapseW = useRefF(0);
  // Grupo AMSA, Ítem, VP, Gerencia, Compañía (colapsan desde el final). Tiene que coincidir con
  // el largo de collapsibleEls: sin Compañía (dashboards de una sola compañía) son 4.
  const N_COLLAPSIBLE = SOLO_DIST_F ? 4 : 5;
  useEffectF(() => {
    const el = barRef.current; if (!el) return;
    const measure = () => {
      const maxH = Array.from(el.children).reduce((m, k) => Math.max(m, k.offsetHeight), 0);
      const wrapped = el.clientHeight > maxH + 36;
      if (wrapped && nHidden < N_COLLAPSIBLE) { collapseW.current = el.clientWidth; setNHidden(n => Math.min(N_COLLAPSIBLE, n + 1)); }
      else if (!wrapped && nHidden > 0 && el.clientWidth > collapseW.current + 160) { setNHidden(n => Math.max(0, n - 1)); }
    };
    measure();
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(measure) : null;
    if (ro) ro.observe(el);
    window.addEventListener('resize', measure);
    return () => { if (ro) ro.disconnect(); window.removeEventListener('resize', measure); };
  }, [nHidden, mode]);
  const collapsibleEls = [
    <div className="fgroup grow" style={{ minWidth: 150 }} key="gamsa">
      <div className="fcap">Grupo AMSA</div>
      <div className="fctl">
        <MultiSelect options={gamsaOpts} selected={st.gamsas || []} onChange={v => set({ gamsas: v })} placeholder="Todos" />
      </div>
    </div>,
    <div className="fgroup grow" style={{ minWidth: 140 }} key="itemrel">
      <div className="fcap">Ítem Relevante</div>
      <div className="fctl">
        <MultiSelect options={itemrelOpts} selected={st.itemrels || []} onChange={v => set({ itemrels: v })} placeholder="Todas" searchable />
      </div>
    </div>,
    <div className="fgroup grow" style={{ minWidth: 140 }} key="vp">
      <div className="fcap">Vicepresidencia</div>
      <div className="fctl">
        <MultiSelect options={vpOpts} selected={st.vps}
          onChange={v => { const vpCecos = [...new Set(A.records.filter(r => v.includes(r.vp)).map(r => r.ceco))]; set({ vps: v, gers: st.gers.filter(g => !v.length || A.records.some(r => v.includes(r.vp) && r.ger === g)), cecos: [...new Set([...(st.cecos || []), ...vpCecos])] }); }}
          placeholder="Todas" searchable />
      </div>
    </div>,
    <div className="fgroup grow" style={{ minWidth: 140 }} key="ger">
      <div className="fcap">Gerencia</div>
      <div className="fctl">
        <MultiSelect options={gerOpts} selected={st.gers} onChange={v => set({ gers: v })} placeholder="Todas" searchable />
      </div>
    </div>,
    ...(SOLO_DIST_F ? [] : [
    <div className="fgroup" style={{ minWidth: 130, opacity: distMode ? 1 : 0.45 }}
      title={distMode ? null : 'Aplica en modo Distribuible / Ambos'} key="comp">
      <div className="fcap">Compañía</div>
      <div className="fctl" style={distMode ? null : { pointerEvents: 'none' }}>
        <MultiSelect options={compOpts} selected={st.companies} onChange={v => set({ companies: v })} placeholder="Todas" />
      </div>
    </div>]),
  ];
  const nShown = N_COLLAPSIBLE - nHidden;

  return (
    <div className="filters" style={{ flexWrap: 'wrap' }} ref={barRef}>
      <div className="fgroup" style={{ minWidth: 116, width: 116 }}>
        <div className="fcap">Base Moneda</div>
        <div className="fctl">
          <select value={valMode} onChange={e => onValMode(e.target.value)} style={{ minWidth: 0, width: '100%' }}>
            <option value="n">Moneda original</option>
            <option value="a">Moneda Ajustada 2027</option>
          </select>
        </div>
      </div>
      <div className="fgroup" style={{ minWidth: 140, width: 140 }}>
        <div className="fcap">Estructura CECOS</div>
        <div className="fctl">
          <select value={cecoMode} onChange={e => onCecoMode(e.target.value)} style={{ minWidth: 0, width: '100%' }}>
            <option value="new">Nuevos</option>
            <option value="old">Antiguos</option>
            <option value="ajustes">Nueva con ajustes</option>
          </select>
        </div>
      </div>
      <div className="fgroup" style={{ minWidth: 130, width: 130 }}>
        <div className="fcap">YTD 2026 hasta</div>
        <div className="fctl">
          <select value={ytdMes} onChange={e => onYtdMes(e.target.value)} style={{ minWidth: 0, width: '100%' }}
            title="Hasta qué mes suman el Real y el Ppto YTD de 2026">
            {MESES_YTD_F.slice(0, ytdMesMax).map((mn, i) => <option key={i} value={i + 1}>{mn}</option>)}
          </select>
        </div>
      </div>

      <div className="fgroup" style={{ minWidth: 210, display: SOLO_DIST_F ? 'none' : null }}>
        <div className="fcap">Datos</div>
        <div style={{ display: 'flex', gap: 8 }}>
          <ModeChip label="Corporativo" on={corpOn} onClick={() => setFlags(!corpOn, distOn)} />
          <ModeChip label="Distribuible" on={distOn} onClick={() => setFlags(corpOn, !distOn)} />
        </div>
      </div>

      <div className="fgroup" style={{ minWidth: 130 }}>
        <div className="fcap">Comercialización</div>
        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, height: 30, fontSize: 12, fontWeight: 600, color: 'var(--fg-2, #444)', cursor: 'pointer', whiteSpace: 'nowrap' }}
          title="Marcado: oculta el Ítem COMERCIALI (Cód. Agrupación4) en todo el dashboard.">
          <input type="checkbox" checked={hideComercial} onChange={e => onHideComercial(e.target.checked)} /> Ocultar
        </label>
      </div>

      <div className="fgroup" style={{ minWidth: 110 }}>
        <div className="fcap">Otros Fletes</div>
        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, height: 30, fontSize: 12, fontWeight: 600, color: 'var(--fg-2, #444)', cursor: 'pointer', whiteSpace: 'nowrap' }}
          title="Marcado: en la VP Comercialización muestra SOLO las combinaciones CECO/CLACO de la whitelist (construir.py → SHOW_COMERCIAL) más los CLACOs de Mano de Obra, y oculta el resto de la VP.">
          <input type="checkbox" checked={hideFletes} onChange={e => onHideFletes(e.target.checked)} /> Ocultar
        </label>
      </div>

      <div className="fgroup" style={{ minWidth: 150 }}>
        <div className="fcap">Comparar</div>
        <SwitchToggle label="Presupuesto 2027" on={showProp} color="var(--amsa-yellow)"
          onClick={() => set(showProp ? { showProp: false, donutMetric: 'real' } : { showProp: true })} />
      </div>

      <div className="fgroup" style={{ minWidth: 130 }}>
        <div className="fcap">Año</div>
        <div className="fctl">
          <MultiSelect options={yearOpts} selected={st.years} onChange={onYears} placeholder="Año" />
        </div>
      </div>

      {collapsibleEls.slice(0, nShown)}

      {/* "+" : filtros colapsados por espacio + Tipo Costo/¿Aplica? (siempre aquí). */}
      <div className="fgroup" style={{ minWidth: 'auto', position: 'relative' }} ref={moreRef}>
        <div className="fcap" aria-hidden="true">&nbsp;</div>
        <button type="button" title="Más filtros"
          onClick={() => setMoreOpen(o => !o)}
          style={{ width: 46, height: 30, boxSizing: 'border-box', padding: 0, position: 'relative',
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            background: '#fff', border: '1px solid var(--amsa-teal)', borderRadius: 3,
            cursor: 'pointer', fontSize: 17, fontWeight: 700, lineHeight: 1, color: 'var(--amsa-teal)' }}>
          +{(nMore + nHidden) > 0
            ? <span style={{ position: 'absolute', top: -4, right: -4, background: 'var(--amsa-yellow)', color: '#3a2e10', borderRadius: 8, fontSize: 9, fontWeight: 800, padding: '1px 5px' }}>{nMore + nHidden}</span>
            : null}
        </button>
        {moreOpen && (
          <div style={{ position: 'absolute', top: '100%', right: 0, marginTop: 6, zIndex: 60,
            background: '#fff', border: '1px solid var(--teal-border)', borderRadius: 8,
            boxShadow: '0 8px 24px rgba(0,0,0,.18)', padding: 12, display: 'flex',
            flexDirection: 'column', gap: 12, minWidth: 210 }}>
            {collapsibleEls.slice(nShown)}
            {extras}
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------- View tabs ---------------- */
function ViewTabs({ view, onChange }) {
  const A = window.CORP;
  return (
    <div className="viewtabs">
      {A.VISTAS.map(v => (
        <button key={v.id}
          className={(v.id === view ? 'on ' : '') + (v.tipo === 'prop' ? 'prop' : '')}
          onClick={() => onChange(v.id)}>{v.label}</button>
      ))}
    </div>
  );
}

/* ---------------- KPI summary cards ---------------- */
function KpiCards({ kpis, unit }) {
  const [hover, setHover] = React.useState(-1);
  return (
    <div className="kpis">
      {kpis.map((k, i) => (
        <div key={i} style={{ position: 'relative', display: 'flex' }}
          onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(-1)}>
          <div className={'kcard ' + (k.status || '')} style={{ flex: 1, ...(k.tooltip ? { cursor: 'help' } : {}) }}>
            <span className="strip" style={k.color ? { background: k.color } : undefined}></span>
            <div className="klbl">{k.label}</div>
            <div className="kval tnum" style={k.color ? { color: k.color } : undefined}>{k.value}{k.unit && <span className="u">{k.unit}</span>}</div>
            <div className="ksub">
              {(k.dot || k.color) && <span className={'dot ' + (typeof k.dot === 'string' ? k.dot : '')} style={k.color ? { background: k.color } : undefined}></span>}
              {k.trendDir && <span className={'trend ' + k.trendDir} style={k.color ? { color: k.color } : undefined}>{k.trend}</span>}
              <span>{k.sub}</span>
            </div>
          </div>
          {k.tooltip && hover === i && (
            <div style={{ position: 'absolute', top: '100%', left: 0, marginTop: 6, zIndex: 60,
              background: 'var(--ink)', color: '#fff', padding: '9px 12px', borderRadius: 8,
              fontSize: 11.5, lineHeight: 1.55, width: 'max-content', maxWidth: 340,
              boxShadow: '0 8px 24px rgba(0,0,0,.28)', pointerEvents: 'none' }}>
              {k.tooltip}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

Object.assign(window, { MultiSelect, FilterBar, ViewTabs, KpiCards });
