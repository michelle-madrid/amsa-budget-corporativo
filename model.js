/* =============================================================
   CORP · Actividad Corporativa + Distribuibles — MODELO DE DATOS
   JS puro (sin JSX). Expone window.CORP (helpers + cómputo).
   Base real: 2022–2025 (Real vs Plan) leído del Excel corporativo.
   Derivados (transparentes, editables/configurables):
     · YTD 2026  — Ppto 2026 proyectado + Real mensual hasta el mes
     · Propuesta 2027 — base editable comparada contra años previos
     · Distribuible por compañía (MLP/ANT/CEN/CMZ) — claves de prorrateo
   ============================================================= */
(function () {
  const D = window.CORP_DATA;
  const DIST = window.DIST_DATA || { records: [] };
  const DOT = window.DOT_DATA || { records: [] };   // Dotaciones (FTE) — Propios/Contratista

  // ===== Textos editables del dashboard (los edita la herramienta "Editar textos") =====
  const TEXTOS = window.TEXTOS = /*TEXTOS-BEGIN*/{
    "header": {
      "titulo": "Actividad Corporativa",
      "tituloPlus": "+ Distribuibles",
      "subtitulo": "Gastos Corporativos por Ítem Relevante, Vicepresidencia y Gerencia"
    },
    "filtros": {
      "datos": "Datos",
      "corporativo": "Corporativo",
      "distribuible": "Distribuible",
      "anio": "Año",
      "item": "Ítem Relevante, Clase de Costo, CLACO",
      "vicepresidencia": "Vicepresidencia",
      "gerencia": "Gerencia",
      "compania": "Compañía (distribuible)",
      "version": "Versión",
      "todas": "Todas"
    },
    "kpi": {
      "realAnual": "Real Anual",
      "presupuestoAnual": "Presupuesto Anual",
      "desviacionAnual": "Desviación Anual",
      "cumplimiento": "Cumplimiento",
      "alarmas": "Alarmas",
      "propuesta2027": "Propuesta 2027",
      "promedio": "Promedio 22–25",
      "deltaPromedio": "Δ vs Prom 22–25"
    },
    "paneles": {
      "realVsPpto": "Real vs Presupuesto por año",
      "realVsPptoSub": "Histórico anual y Propuesta 2027",
      "cumplimiento": "Cumplimiento del presupuesto",
      "cumplimientoSub": "Real / Presupuesto por año",
      "distribuible": "Distribuible por compañía",
      "distribuibleSub": "por compañía",
      "legendRealHist": "Real histórico",
      "legendPresupuesto": "Presupuesto",
      "legendPropuesta": "Propuesta 27",
      "donutTotal": "Total"
    },
    "tabla": {
      "ejecucion": "Ejecución",
      "propuestaTitulo": "Propuesta de Presupuesto 2027",
      "acumuladoAnual": "Acumulado anual",
      "total": "Total Centro Corporativo",
      "colReal": "Real",
      "colDif": "Dif",
      "colPctDif": "% Dif",
      "colKpi": "KPI",
      "buscar": "Buscar VP / Gerencia / Ítem…",
      "agrupar": "Agrupar",
      "decimales": "Decimales",
      "expandir": "Expandir",
      "colapsar": "Colapsar",
      "restablecer": "Restablecer",
      "exportar": "Exportar Excel"
    },
    "footer": {
      "fuenteLabel": "Fuente:",
      "fuente": "Consulta a SAP BPC de valores históricos 2022-2025.",
      "propuestaLabel": "Propuesta 2027:",
      "propuesta": "Pendiente."
    },
    "version": { "original": "Ppto Original" },
    "companias": {
      "MLP": "Minera Los Pelambres",
      "ANT": "Minera Antucoya",
      "CEN": "Minera Centinela",
      "CMZ": "Minera Zaldívar"
    }
  }/*TEXTOS-END*/;

  const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                 'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  const COMPANIAS = [
    { id: 'MLP', nombre: TEXTOS.companias.MLP, color: 'var(--amsa-teal)' },
    { id: 'ANT', nombre: TEXTOS.companias.ANT, color: 'var(--amsa-yellow)' },
    { id: 'CEN', nombre: TEXTOS.companias.CEN, color: 'var(--amsa-red)' },
    { id: 'CMZ', nombre: TEXTOS.companias.CMZ, color: 'var(--blue)' },
  ];

  const VISTAS = [
    { id: '2022', label: '2022', tipo: 'hist', anio: 2022 },
    { id: '2023', label: '2023', tipo: 'hist', anio: 2023 },
    { id: '2024', label: '2024', tipo: 'hist', anio: 2024 },
    { id: '2025', label: '2025', tipo: 'hist', anio: 2025 },
    { id: 'PROP2027', label: TEXTOS.kpi.propuesta2027, tipo: 'prop', anio: 2027 },
  ];

  const VERSIONES = [
    { id: 'ORI', label: TEXTOS.version.original },
    // 'Ppto Ajustado' pendiente: se agregará cuando existan los datos ajustados.
  ];

  /* ---------- utilidades deterministas ---------- */
  function hash(str) {
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); }
    return (h >>> 0) / 4294967295; // 0..1
  }

  // perfil estacional mensual (suma 1). Diciembre algo más cargado.
  const BASE_PROFILE = [0.072,0.076,0.085,0.080,0.084,0.082,0.083,0.085,0.086,0.090,0.087,0.090];
  function monthWeights(seed) {
    const j = hash(seed);
    const w = BASE_PROFILE.map((b, i) => b * (0.9 + 0.2 * hash(seed + ':' + i)));
    const s = w.reduce((a, b) => a + b, 0);
    return w.map(x => x / s);
  }

  // claves de prorrateo distribuible por VP (asunción; suma 1)
  const BASE_KEYS = { MLP: 0.46, ANT: 0.22, CEN: 0.20, CMZ: 0.12 };
  function companyKeys(vp) {
    const raw = COMPANIAS.map(c => BASE_KEYS[c.id] * (0.75 + 0.5 * hash(vp + '#' + c.id)));
    const s = raw.reduce((a, b) => a + b, 0);
    const out = {};
    COMPANIAS.forEach((c, i) => { out[c.id] = raw[i] / s; });
    return out;
  }

  /* ---------- normalización de registros ---------- */
  // Cada registro: {vp, ger, item, y2022..y2025:{real,plan}}
  // Derivamos campos por registro.
  const DEF_GROWTH = { g26: 1.045, gReal26: 1.052, g27: 1.030, ori: 0.955 };

  // ============================================================
  //  UNIFICACIÓN DE GERENCIAS (alias)  ← edita aquí los "juntar bajo un nombre"
  // ------------------------------------------------------------
  //  Distintas exportaciones de SAP truncan el MISMO nombre de Gerencia de forma
  //  distinta (ej.: corporativo "Gcia.Riesgo y Ctrl I" vs distribuible
  //  "Grcia. Riesgos, Compliance y Control Interno"). Esta tabla dice qué nombres
  //  crudos se unen bajo UNA sola Gerencia (clave 'ger') y, opcionalmente, bajo qué
  //  Vicepresidencia ('vp'). El nombre que se MUESTRA se define en gerNames (o se
  //  edita en el Diccionario). Vive en model.js a propósito: NO se sobrescribe al
  //  refrescar los datos con el .bat, así las unificaciones se conservan.
  //  Formato:  'Nombre crudo tal cual viene en los datos': { ger: 'ClaveDestino', vp: 'ClaveVP' }
  const GER_ALIAS = {
    'Gcia.Riesgo y Ctrl I':                         { ger: 'Grcia.Riesg.CompyCIn', vp: 'VP Finanzas' },
    'Grcia. Riesgos, Compliance y Control Interno': { ger: 'Grcia.Riesg.CompyCIn', vp: 'VP Finanzas' },
  };

  // Normaliza un set de registros crudos {comp?,vp,ger,item,y2022..y2025}.
  function normalize(list, offset, src) {
    return list.map((r, i) => {
      const idx = offset + i;
      // Unifica Gerencias equivalentes (ver GER_ALIAS) antes de armar el registro.
      const al = GER_ALIAS[r.ger];
      const ger = al ? al.ger : r.ger;
      const vp = (al && al.vp) ? al.vp : r.vp;
      const seed = (r.comp || '') + '|' + vp + '|' + ger + '|' + r.item + '|' + idx;
      return {
        id: idx, src, comp: r.comp || null,
        vp: vp, _vp0: vp, ger: ger, item: r.item,
        // Clasificación por código (CECOS): Tipo Costo (C1/C3/Comercialización) y ¿Aplica? (Sí/No).
        tc: r.tc || null, ap: r.ap || null,
        hist: {
          2022: { real: r.y2022.real || 0, plan: r.y2022.plan || 0 },
          2023: { real: r.y2023.real || 0, plan: r.y2023.plan || 0 },
          2024: { real: r.y2024.real || 0, plan: r.y2024.plan || 0 },
          2025: { real: r.y2025.real || 0, plan: r.y2025.plan || 0 },
          // 2026 YTD: suma de los meses 2026.01–2026.05 (Real y Ppto) cargada del Excel.
          2026: { real: (r.y2026 && r.y2026.real) || 0, plan: (r.y2026 && r.y2026.plan) || 0 },
        },
        // 2026 FY: presupuesto ANUAL (2026.TOTAL Plan). Serie aparte, solo Ppto
        // (el Real del total no se usa). 0 si el registro no trae presupuesto anual.
        fy26: (r.y2026fy && r.y2026fy.plan) || 0,
        _seed: seed, _weights: monthWeights(seed),
      };
    });
  }
  // Dotaciones (FTE): VP → Gerencia, sin ítem ni compañía. Cada registro trae su
  // propio src ('propios' | 'contratista'). Misma estructura de años que el gasto.
  function normalizeDot(list, offset) {
    return list.map((r, i) => {
      const idx = offset + i;
      const seed = 'dot|' + r.src + '|' + r.vp + '|' + r.ger + '|' + idx;
      const h = (o) => ({ real: (o && o.real) || 0, plan: (o && o.plan) || 0 });
      return {
        id: idx, src: r.src, comp: null,
        vp: r.vp, _vp0: r.vp, ger: r.ger, item: r.ger, tc: null, ap: null,
        hist: { 2022: h(r.y2022), 2023: h(r.y2023), 2024: h(r.y2024), 2025: h(r.y2025), 2026: h(r.y2026) },
        fy26: (r.y2026fy && r.y2026fy.plan) || 0,
        _seed: seed, _weights: monthWeights(seed),
      };
    });
  }
  // Corporativo (BBDD), Distribuible (4 compañías) y Dotaciones en un único índice por id.
  const corpRecords = normalize(D.records, 0, 'corp');
  const distRecords = normalize(DIST.records, corpRecords.length, 'dist');
  const dotRecords = normalizeDot(DOT.records, corpRecords.length + distRecords.length);
  const records = corpRecords.concat(distRecords).concat(dotRecords);

  // Conjunto activo según modo de datos. Gasto: corp | dist | both (+ compañías).
  // Dotaciones: propios | contratista | dot (ambas).
  function activeRecords(opts) {
    const mode = opts.dataMode || 'corp';
    const cset = opts.companies && opts.companies.length ? new Set(opts.companies) : null;
    let recs = [];
    if (mode === 'corp' || mode === 'both') recs = recs.concat(corpRecords);
    if (mode === 'dist' || mode === 'both') recs = recs.concat(cset ? distRecords.filter(r => cset.has(r.comp)) : distRecords);
    if (mode === 'propios' || mode === 'dot') recs = recs.concat(dotRecords.filter(r => r.src === 'propios'));
    if (mode === 'contratista' || mode === 'dot') recs = recs.concat(dotRecords.filter(r => r.src === 'contratista'));
    return recs;
  }

  // Derivados 2026/2027 calculados al vuelo según supuestos de crecimiento.
  function derived(rec, g) {
    // Sin simulación: 2026 no tiene datos y la Propuesta 2027 se ingresa manual
    // (base 0; los valores editados viven en overrides).
    return { ppto26: 0, real26full: 0, prop27: 0 };
  }

  /* ---------- acceso a celdas según vista ---------- */
  // Devuelve métricas crudas (USD) de un registro para una vista.
  // monthIndex: 0..11  → bloque "mes"; ytd → acumulado hasta monthIndex.
  // version: 'AJU'|'ORI'. compFactor: fracción de prorrateo (1 = total).
  function recordMetrics(rec, view, version, monthIndex, compFactor, growth) {
    const f = compFactor;
    const w = rec._weights;
    const g = growth || DEF_GROWTH;
    const cum = (arr, upto) => { let s = 0; for (let i = 0; i <= upto; i++) s += arr[i]; return s; };

    function versionAdj(planAnnual) {
      return version === 'ORI' ? planAnnual * g.ori : planAnnual;
    }

    if (view.tipo === 'hist') {
      const h = rec.hist[view.anio];
      const realA = h.real, planA = versionAdj(h.plan);
      const realMonths = w.map(x => realA * x);
      const planMonths = w.map(x => planA * x);
      return {
        real: realMonths[monthIndex] * f,
        version: planMonths[monthIndex] * f,
        ytdReal: cum(realMonths, monthIndex) * f,
        ytdVersion: cum(planMonths, monthIndex) * f,
      };
    }
    if (view.tipo === 'ytd') {
      const dv = derived(rec, g);
      const realA = dv.real26full, planA = versionAdj(dv.ppto26);
      const realMonths = w.map(x => realA * x);
      const planMonths = w.map(x => planA * x);
      return {
        real: realMonths[monthIndex] * f,
        version: planMonths[monthIndex] * f,
        ytdReal: cum(realMonths, monthIndex) * f,
        ytdVersion: cum(planMonths, monthIndex) * f,
      };
    }
    // prop 2027 manejado aparte
    return { real: 0, version: 0, ytdReal: 0, ytdVersion: 0 };
  }

  /* ---------- agregación por jerarquía ---------- */
  // Filtra registros y arma árbol VP → Gerencia → Ítem con métricas sumadas.
  function buildTree(opts) {
    // opts: {years:[2022..2025], showProp, version, dataMode,
    //        companies, vps, gers, items, overrides}
    // Propuesta 2027 ya no es un modo excluyente: es una capa de comparación
    // aditiva. Con showProp, cada nodo lleva además 'prop' (propuesta) junto a
    // las métricas de ejecución; el promedio de comparación = ytdReal / nYears.
    const showProp = !!opts.showProp;
    const years = (opts.years || []).filter(y => typeof y === 'number');
    // '2026fy' (pseudo-año del filtro Año): suma el Ppto ANUAL 2026 (fy26) al
    // Presupuesto; NO aporta Real (es solo presupuesto).
    const wantFY = (opts.years || []).includes('2026fy');
    // El Plan del BBDD es el Presupuesto Original; no se aplica ningún ajuste.
    const versionAdj = p => p;

    const filtered = activeRecords(opts).filter(rec => {
      if (opts.vps && opts.vps.length && !opts.vps.includes(rec.vp)) return false;
      if (opts.gers && opts.gers.length && !opts.gers.includes(rec.ger)) return false;
      if (opts.items && opts.items.length && !opts.items.includes(rec.item)) return false;
      if (opts.tcs && opts.tcs.length && !opts.tcs.includes(rec.tc)) return false;
      if (opts.aps && opts.aps.length && !opts.aps.includes(rec.ap)) return false;
      return true;
    });

    // Modo "Por año": además del total, cada nodo lleva el desglose Real/Ppto año a año.
    const byYear = opts.yearAgg === 'byYear';

    // métrica por registro: suma de los años seleccionados (acumulado anual).
    // Con showProp se anexa la Propuesta 2027 (override o base). Sin prorrateo.
    function mFor(rec) {
      let real = 0, ver = 0;
      const yr = byYear ? {} : null;
      for (const y of years) {
        const r = rec.hist[y].real, v = versionAdj(rec.hist[y].plan);
        real += r; ver += v;
        if (byYear) yr[y] = { real: r, ver: v };
      }
      if (wantFY) {                       // 2026 Ppto FY: solo Ppto, sin Real
        ver += rec.fy26 || 0;
        if (byYear) yr['2026fy'] = { real: 0, ver: rec.fy26 || 0 };
      }
      const m = { real, version: ver, ytdReal: real, ytdVersion: ver, fy26: rec.fy26 || 0 };
      if (showProp) m.prop = (opts.overrides && opts.overrides[rec.id] != null) ? opts.overrides[rec.id] : derived(rec, opts.growth).prop27;
      if (byYear) m.yr = yr;
      return m;
    }
    const zero = () => {
      const z = showProp
        ? { real: 0, version: 0, ytdReal: 0, ytdVersion: 0, fy26: 0, prop: 0 }
        : { real: 0, version: 0, ytdReal: 0, ytdVersion: 0, fy26: 0 };
      if (byYear) z.yr = {};
      return z;
    };
    const add = (acc, m) => {
      for (const k in acc) {
        if (k === 'yr') continue;
        acc[k] += (m[k] || 0);
      }
      if (acc.yr && m.yr) {
        for (const y in m.yr) {
          if (!acc.yr[y]) acc.yr[y] = { real: 0, ver: 0 };
          acc.yr[y].real += m.yr[y].real;
          acc.yr[y].ver += m.yr[y].ver;
        }
      }
      return acc;
    };

    // Jerarquía configurable (ej. ['vp','ger','item'] o ['item','vp','ger']).
    const dims = (opts.groupBy && opts.groupBy.length) ? opts.groupBy : ['vp', 'ger', 'item'];
    const DISPM = { vp: dispVP, ger: dispGer, item: dispItem };

    // Orden configurable por columna (key: name|real|version|dif|pct).
    const sort = opts.sort || { key: 'real', dir: 'desc' };
    const dirMul = sort.dir === 'asc' ? 1 : -1;
    const promOf = a => (years.length ? a.ytdReal / years.length : 0); // promedio años sel.
    function sortVal(n) {
      const a = n.agg;
      switch (sort.key) {
        case 'version': return a.ytdVersion;
        case 'dif': return a.ytdReal - a.ytdVersion;
        case 'pct': return a.ytdVersion ? (a.ytdReal - a.ytdVersion) / Math.abs(a.ytdVersion) : 0;
        case 'prop': return a.prop || 0;
        case 'dprop': { const p = promOf(a); return p ? ((a.prop || 0) - p) / Math.abs(p) : 0; }
        default: return a.ytdReal; // 'real'
      }
    }
    function cmpFor(depth) {
      if (sort.key === 'name') {
        const disp = DISPM[dims[depth]];
        return (a, b) => dirMul * disp(a.name).localeCompare(disp(b.name), 'es');
      }
      return (a, b) => dirMul * (sortVal(a) - sortVal(b));
    }

    const metricOf = new Map();
    filtered.forEach(r => metricOf.set(r, mFor(r)));

    function group(recs, depth) {
      const map = new Map(), order = [];
      recs.forEach(rec => {
        const k = rec[dims[depth]];
        if (!map.has(k)) { map.set(k, []); order.push(k); }
        map.get(k).push(rec);
      });
      const nodes = order.map(k => {
        const sub = map.get(k);
        const agg = zero();
        sub.forEach(r => add(agg, metricOf.get(r)));
        const node = { name: k, agg };
        if (depth < dims.length - 1) node.children = group(sub, depth + 1);
        else { node.leaf = true; node.recIds = sub.map(r => r.id); }
        return node;
      });
      nodes.sort(cmpFor(depth));
      return nodes;
    }

    const vpNodes = group(filtered, 0);
    const total = zero();
    vpNodes.forEach(n => add(total, n.agg));

    return { vpNodes, total, showProp, nYears: years.length, byYear, years: years.slice(), dims };
  }

  /* ---------- series anuales (para gráficos) ---------- */
  function annualSeries(opts) {
    // suma anual sobre el conjunto activo (año completo, sin prorrateo).
    const filtered = activeRecords(opts).filter(rec => {
      if (opts.vps && opts.vps.length && !opts.vps.includes(rec.vp)) return false;
      if (opts.gers && opts.gers.length && !opts.gers.includes(rec.ger)) return false;
      if (opts.items && opts.items.length && !opts.items.includes(rec.item)) return false;
      if (opts.tcs && opts.tcs.length && !opts.tcs.includes(rec.tc)) return false;
      if (opts.aps && opts.aps.length && !opts.aps.includes(rec.ap)) return false;
      return true;
    });
    const g = opts.growth || DEF_GROWTH;
    const versionAdj = p => p; // Plan del BBDD = Ppto Original (sin ajuste)
    const out = {
      2022: { real: 0, plan: 0 }, 2023: { real: 0, plan: 0 },
      2024: { real: 0, plan: 0 }, 2025: { real: 0, plan: 0 },
      2026: { real: 0, plan: 0 },
      '2026fy': { plan: 0 },
      2027: { prop: 0 },
    };
    filtered.forEach(rec => {
      [2022, 2023, 2024, 2025].forEach(y => {
        out[y].real += rec.hist[y].real;
        out[y].plan += versionAdj(rec.hist[y].plan);
      });
      // 2026 YTD: valores reales cargados del Excel (no simulados).
      out[2026].real += rec.hist[2026].real;
      out[2026].plan += versionAdj(rec.hist[2026].plan);
      // 2026 FY: presupuesto anual (solo Ppto).
      out['2026fy'].plan += rec.fy26 || 0;
      const prop = (opts.overrides && opts.overrides[rec.id] != null) ? opts.overrides[rec.id] : derived(rec, g).prop27;
      out[2027].prop += prop;
    });
    return out;
  }

  function distribuible(opts) {
    // Distribuible REAL por compañía para la vista (año completo).
    // Filtra por compañía/VP/ítem; no por gerencia (los CECO del distribuible
    // no equivalen a las gerencias del corporativo).
    const cset = opts.companies && opts.companies.length ? new Set(opts.companies) : null;
    // Métrica del donut: 'real' (años seleccionados) o 'prop' (Propuesta 2027).
    const useProp = opts.donutMetric === 'prop';
    const years = (opts.years || []).filter(y => typeof y === 'number');
    const filtered = distRecords.filter(rec => {
      if (cset && !cset.has(rec.comp)) return false;
      if (opts.vps && opts.vps.length && !opts.vps.includes(rec.vp)) return false;
      if (opts.items && opts.items.length && !opts.items.includes(rec.item)) return false;
      if (opts.tcs && opts.tcs.length && !opts.tcs.includes(rec.tc)) return false;
      if (opts.aps && opts.aps.length && !opts.aps.includes(rec.ap)) return false;
      return true;
    });
    const out = { MLP: 0, ANT: 0, CEN: 0, CMZ: 0 };
    filtered.forEach(rec => {
      let b;
      if (useProp) b = (opts.overrides && opts.overrides[rec.id] != null) ? opts.overrides[rec.id] : derived(rec, opts.growth).prop27;
      else { b = 0; for (const y of years) b += rec.hist[y].real; }
      out[rec.comp] += b;
    });
    // Compañías excluidas por el filtro → null (N/A), para distinguirlas de un 0 real.
    if (cset) Object.keys(out).forEach(id => { if (!cset.has(id)) out[id] = null; });
    return out;
  }

  /* ---------- KPI semáforo ---------- */
  // pctDev = |dif| / version  (gasto: real sobre versión = sobregasto)
  function kpiColor(real, version, thr) {
    if (!version) return real > 0 ? 'rojo' : 'azul'; // gasto sin presupuesto = sobregasto (100%)
    const dev = (real - version) / Math.abs(version); // + = sobre presupuesto
    if (dev <= 0) return 'azul';                       // igual o bajo presupuesto = verde
    const a = dev * 100;                                // % sobre presupuesto
    return a > thr.red ? 'rojo' : 'amarillo';           // > red% = rojo; (0, red%] = amarillo
  }

  /* ---------- formato ---------- */
  function fmt(usd, unit, decimals) {
    // unit 'num' = valor crudo (FTE / Nº, sin dividir); 'kUSD' = miles; resto = millones.
    const div = unit === 'num' ? 1 : unit === 'kUSD' ? 1e3 : 1e6;
    let v = usd / div;
    const d = decimals != null ? decimals : (unit === 'num' ? 0 : unit === 'kUSD' ? 0 : 1);
    return v.toLocaleString('es-CL', { minimumFractionDigits: d, maximumFractionDigits: d });
  }
  function fmtPct(x, d) {
    if (x == null || !isFinite(x)) return '—';
    return (x * 100).toLocaleString('es-CL', { minimumFractionDigits: d || 0, maximumFractionDigits: d || 0 }) + '%';
  }

  // Los nombres llegan truncados a 20 chars desde SAP; cada mapa trae la versión
  // completa. Si una clave falta en su propio mapa (p. ej. "Presidencia Ejecutiv"
  // usada como Gerencia pero solo definida como VP), se busca en los demás mapas
  // antes de mostrar la clave cruda truncada.
  const _allNames = Object.assign({}, D.itemNames, D.gerNames, D.vpNames);
  // Overrides de nombre EDITABLES por el usuario (Gerencia / Ítem), keyed por la
  // clave estable (rec.ger / rec.item). Persisten en localStorage (App) y tienen
  // prioridad sobre el nombre de la base, así un refresco de datos conserva los
  // renombres mientras la clave no cambie. Vacíos por defecto.
  let _gerOv = {}, _itemOv = {};
  function applyNameOverrides(ov) {
    _gerOv = (ov && ov.ger) || {};
    _itemOv = (ov && ov.item) || {};
  }
  // Nombre "de base" (sin override) — para mostrar el original y permitir reset.
  function baseGer(n) { return D.gerNames[n] || _allNames[n] || n; }
  function baseItem(n) { return D.itemNames[n] || _allNames[n] || n; }
  function dispVP(n) { return D.vpNames[n] || _allNames[n] || n; }
  function dispGer(n) { return _gerOv[n] || D.gerNames[n] || _allNames[n] || n; }
  function dispItem(n) { return _itemOv[n] || D.itemNames[n] || _allNames[n] || n; }

  // Opciones de Gerencia (dependen de VP) e Ítem según el modo de datos activo.
  function dimsFor(opts) {
    const recs = activeRecords(opts);
    const vpset = opts.vps && opts.vps.length ? new Set(opts.vps) : null;
    const gers = [], gseen = new Set();
    const items = [], iseen = new Set();
    const tcs = [], tseen = new Set();   // Tipo Costo (C1/C3/Comercialización)
    const aps = [], aseen = new Set();   // ¿Aplica? (Sí/No)
    recs.forEach(r => {
      if (!vpset || vpset.has(r.vp)) { if (!gseen.has(r.ger)) { gseen.add(r.ger); gers.push(r.ger); } }
      if (!iseen.has(r.item)) { iseen.add(r.item); items.push(r.item); }
      if (r.tc && !tseen.has(r.tc)) { tseen.add(r.tc); tcs.push(r.tc); }
      if (r.ap && !aseen.has(r.ap)) { aseen.add(r.ap); aps.push(r.ap); }
    });
    gers.sort((a, b) => a.localeCompare(b, 'es'));
    items.sort((a, b) => a.localeCompare(b, 'es'));
    tcs.sort((a, b) => a.localeCompare(b, 'es'));
    aps.sort((a, b) => a.localeCompare(b, 'es'));
    return { gers, items, tcs, aps };
  }

  // Resolución de colores del tema: traduce cualquier expresión CSS (var(--x),
  // color-mix(...), hex) a un color concreto rgb() leyendo el valor computado.
  // Necesario para los SVG de las gráficas y el semáforo, donde var() NO funciona
  // como atributo (solo en CSS). Los colores se definen en el bloque <style id="ada-tema">.
  let _probe = null, _colorCache = {};
  function color(expr, fallback) {
    if (expr in _colorCache) return _colorCache[expr];
    try {
      if (!document || !document.body) return fallback || expr;
      if (!_probe) { _probe = document.createElement('span'); _probe.style.cssText = 'position:absolute;left:-9999px;top:-9999px;visibility:hidden;pointer-events:none'; }
      if (!_probe.isConnected) document.body.appendChild(_probe);
      _probe.style.color = ''; _probe.style.color = expr;
      const c = getComputedStyle(_probe).color;
      const out = (c && c !== 'rgba(0, 0, 0, 0)') ? c : (fallback || expr);
      _colorCache[expr] = out;
      return out;
    } catch (e) { return fallback || expr; }
  }
  // Paleta resuelta para las gráficas (un solo lugar lee las variables del tema).
  let _theme = null;
  function theme() {
    if (_theme) return _theme;
    const t = {
      teal: color('var(--amsa-teal)', '#2a8a96'),
      tealDeep: color('var(--amsa-teal-deep)', '#14515a'),
      tealLight: color('var(--amsa-teal-light)', '#b9dde0'),
      yellow: color('var(--amsa-yellow)', '#f0a929'),
      line: color('var(--line-soft)', '#EEF1F3'),
      grid: color('var(--teal-300)', '#c2c9cd'),
      fgSoft: color('var(--fg-soft)', '#9aa3a8'),
      fg2: color('var(--fg-2)', '#4d4d4d'),
      fg3: color('var(--fg-3)', '#5B5C64'),
      fg4: color('var(--fg-4)', '#858585'),
      ink: color('var(--ink)', '#1f2428'),
    };
    if (document && document.body) _theme = t; // cachea solo cuando ya se pudo resolver
    return t;
  }
  // Limpia las cachés de color: tras cambiar variables del tema en vivo (panel de
  // colores), fuerza que charts/donut/kpi vuelvan a resolver los colores.
  function resetTheme() { _colorCache = {}; _theme = null; }

  // Semáforo a color: ok = verde, advertencia = ámbar, alerta = rojo (del tema).
  function kpiHex(c) { return c === 'rojo' ? color('var(--red)', '#DC3545') : c === 'amarillo' ? color('var(--yellow)', '#E0A800') : color('var(--ok)', '#1f9d57'); }

  // VP reasignada desde el Diccionario (override editable por CECO). El override
  // viene por CECO con el nombre LARGO de la VP; aquí lo traducimos a la clave
  // corta que usan los registros y lo aplicamos a nivel de Gerencia (cada CECO
  // pertenece a una). Mutar rec.vp hace que TODO el modelo (árbol, series, donut,
  // filtros, KPIs) refleje el cambio sin tocar cada agregación. Los registros sin
  // override (o de Distribuible) vuelven a su VP original (_vp0).
  let _vpFullToShort = null;
  function applyVpOverrides(ovByCeco) {
    ovByCeco = ovByCeco || {};
    if (!_vpFullToShort) {
      _vpFullToShort = {};
      for (const k in D.vpNames) _vpFullToShort[D.vpNames[k]] = k;
    }
    const dict = (window.CORP_DICT && window.CORP_DICT.cecos) || [];
    const gerVp = {};
    dict.forEach(c => {
      const nv = ovByCeco[c.c];
      if (nv) gerVp[c.g] = _vpFullToShort[nv] || nv; // nombre largo → clave corta
    });
    records.forEach(r => { r.vp = gerVp[r.ger] || r._vp0; });
  }

  window.CORP = {
    D, records, corpRecords, distRecords, dotRecords, MESES, COMPANIAS, VISTAS, VERSIONES,
    buildTree, annualSeries, distribuible, activeRecords, dimsFor, applyVpOverrides, kpiColor, kpiHex, color, theme, resetTheme,
    fmt, fmtPct, dispVP, dispGer, dispItem, applyNameOverrides, baseGer, baseItem,
    recordById: (id) => records[id],
    derived,
    DEF_GROWTH,
    GER_ALIAS,
  };
})();
