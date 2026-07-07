/* global React */
const { useState: useStateF, useRef: useRefF, useEffect: useEffectF } = React;

/* ---------------- Multi-select dropdown ---------------- */
function MultiSelect({ options, selected, onChange, placeholder, gold, searchable }) {
  const [open, setOpen] = useStateF(false);
  const [q, setQ] = useStateF('');
  const ref = useRefF(null);
  useEffectF(() => {
    function onDoc(e) { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }
    document.addEventListener('mousedown', onDoc);
    return () => document.removeEventListener('mousedown', onDoc);
  }, []);
  const sel = new Set(selected);
  const toggle = (v) => {
    const n = new Set(sel);
    n.has(v) ? n.delete(v) : n.add(v);
    onChange([...n]);
  };
  const filtered = options.filter(o => o.label.toLowerCase().includes(q.toLowerCase()));
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
                <input type="checkbox" checked={sel.has(o.value)} onChange={() => toggle(o.value)} />
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
  const { st, set, gerOptions, itemOptions, tcOptions, apOptions } = props;
  // Opciones de Ítem y Gerencia según el modo de datos activo (corp/dist/ambos).
  const itemVals = itemOptions || A.D.items;
  const itemOpts = itemVals.map(v => ({ value: v, label: A.dispItem(v) }));
  // Tipo Costo (C1/C3/Comercialización) y ¿Aplica? (Sí/No) desde CECOS, por código.
  const tcOpts = (tcOptions || []).map(v => ({ value: v, label: v }));
  const apOpts = (apOptions || []).map(v => ({ value: v, label: v }));
  // Incluye cualquier VP presente en los registros (p. ej. una gerencia reasignada
  // a otra VP desde el Diccionario), además del catálogo base.
  const vpOpts = [...new Set([...A.D.vps, ...A.records.map(r => r.vp)])].map(v => ({ value: v, label: A.dispVP(v) }));
  const gerVals = gerOptions
    || (st.vps.length ? [...new Set(A.records.filter(r => st.vps.includes(r.vp)).map(r => r.ger))] : A.D.gers);
  const gerOpts = gerVals.map(v => ({ value: v, label: A.dispGer(v) }));
  const compOpts = A.COMPANIAS.map(c => ({ value: c.id, label: `${c.id} · ${c.nombre}` }));
  const mode = st.dataMode || 'both';
  const corpOn = mode === 'corp' || mode === 'both';
  const distOn = mode === 'dist' || mode === 'both';
  const distMode = distOn;
  const setFlags = (c, d) => {
    const dataMode = c && d ? 'both' : c ? 'corp' : d ? 'dist' : 'none';
    // Al cambiar de modo, conservar las selecciones de Gerencia/Ítem que sigan
    // siendo válidas en el nuevo conjunto de datos (no borrar todo el filtro).
    const dims = A.dimsFor({ ...st, dataMode });
    const gset = new Set(dims.gers), iset = new Set(dims.items);
    set({ dataMode, gers: st.gers.filter(g => gset.has(g)), items: st.items.filter(i => iset.has(i)) });
  };

  // Año: multi-selección de años reales (2022-2025). La Propuesta 2027 ya no vive
  // aquí: es un switch aparte (capa de comparación aditiva).
  const yearOpts = [
    { value: 2022, label: '2022' }, { value: 2023, label: '2023' },
    { value: 2024, label: '2024' }, { value: 2025, label: '2025' },
    { value: 2026, label: '2026 YTD' },     // acumulado ene–may 2026 (Real vs Ppto)
    { value: '2026fy', label: '2026 Ppto FY' }, // presupuesto anual 2026 (solo Ppto, sin Real)
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
      <div className="fcap">Tipo Costo</div>
      <div className="fctl">
        <MultiSelect options={tcOpts} selected={st.tcs || []} onChange={v => set({ tcs: v })} placeholder="Todos" />
      </div>
    </div>,
    <div className="fgroup" style={{ minWidth: 150 }} key="ap">
      <div className="fcap">¿Aplica?</div>
      <div className="fctl">
        <MultiSelect options={apOpts} selected={st.aps || []} onChange={v => set({ aps: v })} placeholder="Todas" />
      </div>
    </div>,
  ];
  const nMore = ((st.tcs && st.tcs.length) ? 1 : 0) + ((st.aps && st.aps.length) ? 1 : 0);

  return (
    <div className="filters">
      <div className="fgroup" style={{ minWidth: 210 }}>
        <div className="fcap">Datos</div>
        <div style={{ display: 'flex', gap: 8 }}>
          <ModeChip label="Corporativo" on={corpOn} onClick={() => setFlags(!corpOn, distOn)} />
          <ModeChip label="Distribuible" on={distOn} onClick={() => setFlags(corpOn, !distOn)} />
        </div>
      </div>

      <div className="fgroup" style={{ minWidth: 150 }}>
        <div className="fcap">Comparar</div>
        <SwitchToggle label="Propuesta 2027" on={showProp} color="var(--amsa-yellow)"
          onClick={() => set(showProp ? { showProp: false, donutMetric: 'real' } : { showProp: true })} />
      </div>

      <div className="fgroup" style={{ minWidth: 130 }}>
        <div className="fcap">Año</div>
        <div className="fctl">
          <MultiSelect options={yearOpts} selected={st.years} onChange={onYears} placeholder="Año" />
        </div>
      </div>

      <div className="fgroup grow">
        <div className="fcap">Ítem Relevante (CLACO)</div>
        <div className="fctl">
          <MultiSelect options={itemOpts} selected={st.items} onChange={v => set({ items: v })}
            placeholder="Todas" searchable />
        </div>
      </div>

      <div className="fgroup grow">
        <div className="fcap">Vicepresidencia</div>
        <div className="fctl">
          <MultiSelect options={vpOpts} selected={st.vps}
            onChange={v => set({ vps: v, gers: st.gers.filter(g => !v.length || A.records.some(r => v.includes(r.vp) && r.ger === g)) })}
            placeholder="Todas" searchable />
        </div>
      </div>

      <div className="fgroup grow">
        <div className="fcap">Gerencia</div>
        <div className="fctl">
          <MultiSelect options={gerOpts} selected={st.gers} onChange={v => set({ gers: v })}
            placeholder="Todas" searchable />
        </div>
      </div>

      <div className="fgroup" style={distMode ? null : { opacity: 0.45 }}
        title={distMode ? null : 'Aplica en modo Distribuible / Ambos'}>
        <div className="fcap">Compañía</div>
        <div className="fctl" style={distMode ? null : { pointerEvents: 'none' }}>
          <MultiSelect options={compOpts} selected={st.companies} onChange={v => set({ companies: v })}
            placeholder="Todas" />
        </div>
      </div>

      <div className="fgroup" style={{ minWidth: 128 }}>
        <div className="fcap">Versión</div>
        <div className="fctl">
          <select value={st.version} onChange={e => set({ version: e.target.value })}>
            {A.VERSIONES.map(v => <option key={v.id} value={v.id}>{v.label}</option>)}
          </select>
        </div>
      </div>

      {/* Tipo Costo + ¿Aplica? siempre dentro del "+". */}
      <div className="fgroup" style={{ minWidth: 'auto', position: 'relative' }} ref={moreRef}>
        <div className="fcap" aria-hidden="true">&nbsp;</div>
        <button type="button" title="Más filtros (Tipo Costo, ¿Aplica?)"
          onClick={() => setMoreOpen(o => !o)}
          style={{ width: 46, height: 30, boxSizing: 'border-box', padding: 0, position: 'relative',
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            background: '#fff', border: '1px solid var(--amsa-teal)', borderRadius: 3,
            cursor: 'pointer', fontSize: 17, fontWeight: 700, lineHeight: 1, color: 'var(--amsa-teal)' }}>
          +{nMore > 0
            ? <span style={{ position: 'absolute', top: -4, right: -4, background: 'var(--amsa-yellow)', color: '#3a2e10', borderRadius: 8, fontSize: 9, fontWeight: 800, padding: '1px 5px' }}>{nMore}</span>
            : null}
        </button>
        {moreOpen && (
          <div style={{ position: 'absolute', top: '100%', right: 0, marginTop: 6, zIndex: 60,
            background: '#fff', border: '1px solid var(--teal-border)', borderRadius: 8,
            boxShadow: '0 8px 24px rgba(0,0,0,.18)', padding: 12, display: 'flex',
            flexDirection: 'column', gap: 12, minWidth: 190 }}>
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
        <div key={i} style={{ position: 'relative' }}
          onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(-1)}>
          <div className={'kcard ' + (k.status || '')} style={k.tooltip ? { cursor: 'help' } : undefined}>
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
