/* global React */
const { useState, useRef, useEffect } = React;

// Mide el ancho real del contenedor para que el gráfico llene el ancho
// disponible MANTENIENDO una altura fija (no escala verticalmente con el
// ancho). Así las tarjetas conservan su alto aunque las gráficas se ensanchen.
function useBoxWidth(fallback) {
  const ref = useRef(null);
  const [w, setW] = useState(fallback);
  useEffect(() => {
    const el = ref.current;
    if (!el || typeof ResizeObserver === 'undefined') return;
    const upd = () => setW(el.clientWidth || fallback);
    upd();
    const ro = new ResizeObserver(upd);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  return [ref, w];
}
// Escala unidad-viewBox → px (mantiene el tamaño visual de fuentes/barras).
const CHART_SCALE = 1.5;

/* ---------------- Bar chart: Real vs Presupuesto por año + 2027 ---------------- */
function BarYears({ series, unit, decimals, showProp }) {
  const A = window.CORP;
  const C = A.theme();
  const rows = [
    { y: '2022', real: series[2022].real, plan: series[2022].plan },
    { y: '2023', real: series[2023].real, plan: series[2023].plan },
    { y: '2024', real: series[2024].real, plan: series[2024].plan },
    { y: '2025', real: series[2025].real, plan: series[2025].plan },
  ];
  // 2026 YTD (acumulado ene–may): se muestra como barra aparte (color teal) si hay datos.
  if (series[2026] && (series[2026].real || series[2026].plan))
    rows.push({ y: '2026 YTD', real: series[2026].real, plan: series[2026].plan, ytd: true });
  // 2026 FY (presupuesto anual): SOLO Ppto, sin Real (barra de presupuesto).
  if (series['2026fy'] && series['2026fy'].plan)
    rows.push({ y: '2026 FY', plan: series['2026fy'].plan, fyBudget: true });
  // La barra de Propuesta 2027 solo aparece con el switch activado (capa aditiva).
  if (showProp) rows.push({ y: 'Prop 27', prop: series[2027].prop, prop27: true });
  const max = Math.max(1, ...rows.flatMap(r => [r.real || 0, r.plan || 0, r.prop || 0]));
  const [boxRef, boxW] = useBoxWidth(780);
  const W = Math.max(360, Math.round(boxW / CHART_SCALE)), H = 150, padL = 26, padB = 24, padT = 10;
  const innerH = H - padB - padT;
  const groupW = (W - padL) / rows.length;
  const div = unit === 'num' ? 1 : unit === 'kUSD' ? 1e3 : 1e6;

  const yTicks = 4;
  const ticks = Array.from({ length: yTicks + 1 }, (_, i) => (max / yTicks) * i);

  return (
    <div ref={boxRef} style={{ width: '100%' }}>
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" height={Math.round(H * CHART_SCALE)} preserveAspectRatio="xMidYMid meet" style={{ display: 'block' }}>
      {ticks.map((t, i) => {
        const yy = padT + innerH - (t / max) * innerH;
        return (
          <g key={i}>
            <line x1={padL} x2={W} y1={yy} y2={yy} stroke={C.line} strokeWidth="1" />
            <text x={3} y={yy - 3} fontSize="8.5" fill={C.fgSoft} fontFamily="Montserrat">{(t / div).toFixed(0)}</text>
          </g>
        );
      })}
      {rows.map((r, i) => {
        const gx = padL + i * groupW;
        const bw = r.prop27 ? groupW * 0.34 : groupW * 0.28;
        const gap = groupW * 0.08;
        const cx = gx + groupW / 2;
        const bar = (val, color, dx) => {
          const h = (val / max) * innerH;
          const x = cx + dx;
          const y = padT + innerH - h;
          return <g>
            <rect x={x} y={y} width={bw} height={Math.max(0, h)} fill={color} rx="1.5" />
            {val > 0 && <text x={x + bw / 2} y={y - 3} fontSize="7.5" fill={C.fg3} textAnchor="middle" fontFamily="Montserrat" fontWeight="700">{A.fmt(val, unit, 1)}</text>}
          </g>;
        };
        return (
          <g key={i}>
            {r.prop27
              ? bar(r.prop, C.yellow, -bw / 2)
              : r.fyBudget
                ? bar(r.plan, C.tealLight, -bw / 2)
                : <g>
                    {bar(r.real, C.tealDeep, -bw - gap / 2)}
                    {bar(r.plan, C.tealLight, gap / 2)}
                  </g>}
            <text x={cx} y={H - 9} fontSize="9.5" fill={C.fg3} textAnchor="middle" fontFamily="Montserrat" fontWeight="600">{r.y}</text>
          </g>
        );
      })}
    </svg>
    </div>
  );
}

/* ---------------- Cumplimiento histórico (% Real / Ppto) ---------------- */
function ComplianceBars({ series, thr, unit }) {
  const A = window.CORP;
  const C = A.theme();
  const rows = [2022, 2023, 2024, 2025].map(y => {
    const pct = series[y].plan ? series[y].real / series[y].plan : 0;
    return { y: String(y), pct };
  });
  // 2026 YTD: % de cumplimiento del acumulado ene–may (Real / Ppto), barra atenuada.
  if (series[2026] && (series[2026].real || series[2026].plan))
    rows.push({ y: '2026 YTD', pct: series[2026].plan ? series[2026].real / series[2026].plan : 0, ytd: true });
  const [boxRef, boxW] = useBoxWidth(480);
  const W = Math.max(280, Math.round(boxW / CHART_SCALE)), H = 150, padB = 24, padT = 16, padL = 8;
  const innerH = H - padB - padT;
  const maxPct = Math.max(1.2, ...rows.map(r => r.pct));
  const groupW = (W - padL) / rows.length;
  // Misma regla que el semáforo KPI: verde <= ppto, amarillo hasta thr.red%, rojo > thr.red%.
  const colorFor = (pct) => A.kpiHex(A.kpiColor(pct, 1, thr));
  const oneY = padT + innerH - (1 / maxPct) * innerH;
  return (
    <div ref={boxRef} style={{ width: '100%' }}>
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" height={Math.round(H * CHART_SCALE)} preserveAspectRatio="xMidYMid meet" style={{ display: 'block' }}>
      <line x1={padL} x2={W} y1={oneY} y2={oneY} stroke={C.grid} strokeWidth="1" strokeDasharray="3 3" />
      <text x={W - 2} y={oneY - 4} fontSize="8.5" fill={C.fgSoft} textAnchor="end" fontFamily="Montserrat">100% ppto</text>
      {rows.map((r, i) => {
        const gx = padL + i * groupW;
        const bw = groupW * 0.5;
        const cx = gx + groupW / 2;
        const h = (r.pct / maxPct) * innerH;
        const y = padT + innerH - h;
        const c = colorFor(r.pct);
        return (
          <g key={i}>
            <rect x={cx - bw / 2} y={y} width={bw} height={Math.max(0, h)} fill={c} rx="1.5" />
            <text x={cx} y={y - 5} fontSize="9" fill={c} textAnchor="middle" fontFamily="Montserrat" fontWeight="700">{(r.pct * 100).toFixed(0)}%</text>
            <text x={cx} y={H - 9} fontSize="9.5" fill={C.fg3} textAnchor="middle" fontFamily="Montserrat" fontWeight="600">{r.y}</text>
          </g>
        );
      })}
    </svg>
    </div>
  );
}

/* ---------------- Distribuible por compañía (donut) ---------------- */
function Donut({ dist, unit, decimals }) {
  const A = window.CORP;
  const C = A.theme();
  // val === null → compañía no seleccionada (N/A); cualquier número (incl. 0) es un valor real.
  const data = A.COMPANIAS.map(c => ({ ...c, val: dist[c.id] == null ? null : dist[c.id] })).sort((a, b) => (b.val || 0) - (a.val || 0));
  const total = data.reduce((a, b) => a + (b.val || 0), 0) || 1;
  const R = 64, r = 40, cx = 80, cy = 80;
  let acc = 0;
  const arcs = data.map(d => {
    const frac = (d.val || 0) / total;
    const a0 = acc * 2 * Math.PI - Math.PI / 2;
    acc += frac;
    const a1 = acc * 2 * Math.PI - Math.PI / 2;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const p = (ang, rad) => [cx + rad * Math.cos(ang), cy + rad * Math.sin(ang)];
    const [x0, y0] = p(a0, R), [x1, y1] = p(a1, R);
    const [x2, y2] = p(a1, r), [x3, y3] = p(a0, r);
    const dd = `M${x0},${y0} A${R},${R} 0 ${large} 1 ${x1},${y1} L${x2},${y2} A${r},${r} 0 ${large} 0 ${x3},${y3} Z`;
    return { ...d, dd, frac };
  });
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
      <svg viewBox="0 0 160 160" width="150" height="150" style={{ flex: '0 0 auto' }}>
        {arcs.map((a, i) => <path key={i} d={a.dd} fill={A.color(a.color)} stroke="#fff" strokeWidth="1.5" />)}
        <text x="80" y="76" textAnchor="middle" fontSize="11" fill={C.fg4} fontFamily="Montserrat" fontWeight="600">Total</text>
        <text x="80" y="92" textAnchor="middle" fontSize="15" fill={C.ink} fontFamily="Montserrat" fontWeight="700">{A.fmt(total, unit, decimals)}</text>
      </svg>
      <div style={{ flex: 1 }}>
        {arcs.map((a, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, padding: '4px 0', fontSize: 12 }}>
            <span style={{ display: 'flex', alignItems: 'center', gap: 7, color: 'var(--fg-2)', minWidth: 0 }}>
              <i style={{ width: 10, height: 10, borderRadius: 2, background: a.color, display: 'inline-block', flexShrink: 0 }}></i>
              <b style={{ color: 'var(--ink)', flexShrink: 0 }}>{a.id}</b>
              <span style={{ color: 'var(--fg-soft)', fontSize: 10.5, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{a.nombre}</span>
            </span>
            <span className="tnum" style={{ color: a.val == null ? 'var(--fg-soft)' : '#1f2428', fontWeight: 700, whiteSpace: 'nowrap', flexShrink: 0 }}>
              {a.val == null
                ? <span style={{ fontWeight: 600 }}>N/A</span>
                : <>{A.fmt(a.val, unit, decimals)} <span style={{ color: 'var(--fg-soft)', fontWeight: 600 }}>· {(a.frac * 100).toFixed(0)}%</span></>}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

Object.assign(window, { BarYears, ComplianceBars, Donut });
