/* =============================================================
   CORP · Actividad Corporativa + Distribuibles — MODELO DE DATOS
   JS puro (sin JSX). Expone window.CORP (helpers + cómputo).
   Base real: 2022–2025 (Real vs Plan) leído del Excel corporativo.
   Derivados (transparentes, editables/configurables):
     · YTD 2026  — Ppto 2026 proyectado + Real mensual hasta el mes
     · Presupuesto 2027 — base editable comparada contra años previos
     · Distribuible por compañía (MLP/ANT/CEN/CMZ) — claves de prorrateo
   ============================================================= */
(function () {
  // ===== Datos v2: window.V2_DATA (registros CECO×Ítem con doble valor n/a + mapas de CECOS) =====
  const V = window.V2_DATA || { records: [], items: [], itemNames: {}, clacoNames: {}, cecoNew: {}, cecoOld: {}, comps: {} };
  const DOT = window.DOT_DATA || { records: [] };   // Dotaciones (FTE) — Propios/Contratista
  // Modo de valor ('n' Normal | 'a' Ajustada 2027) y de CECOS ('new' | 'old'),
  // elegibles en la barra de filtros. Al cambiar se reconstruyen los registros.
  let _valMode = 'n', _cecoMode = 'new';   // Estructura CECOS por defecto: Nueva

  // ===== Textos editables del dashboard (los edita la herramienta "Editar textos") =====
  const TEXTOS = window.TEXTOS = /*TEXTOS-BEGIN*/{
    "header": {
      "titulo": "Actividad Corporativa",
      "tituloPlus": "+ Distribuibles Ppto 2027",
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
      "propuesta2027": "Presupuesto 2027",
      "promedio": "Promedio 22–25",
      "deltaPromedio": "Δ vs Prom 22–25"
    },
    "paneles": {
      "realVsPpto": "Real vs Presupuesto por año",
      "realVsPptoSub": "Histórico anual y Presupuesto 2027",
      "cumplimiento": "Cumplimiento del presupuesto",
      "cumplimientoSub": "Real / Presupuesto por año",
      "distribuible": "Distribuible por compañía",
      "distribuibleSub": "por compañía",
      "legendRealHist": "Real histórico",
      "legendPresupuesto": "Presupuesto",
      "legendPropuesta": "Presupuesto 27",
      "donutTotal": "Total"
    },
    "tabla": {
      "ejecucion": "Ejecución",
      "propuestaTitulo": "Presupuesto 2027",
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
      "propuestaLabel": "Presupuesto 2027:",
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

  // ---- v2: construye los registros desde V.records según los modos activos ----
  function _cmap() { return _cecoMode === 'old' ? V.cecoOld : V.cecoNew; }
  function _pick(cell) { return cell ? (cell[_valMode] || 0) : 0; }   // celda {n,a} → valor del modo
  function _srcOf(compCode) { const c = V.comps[compCode]; return (c && /distrib/i.test(c.clasif || '')) ? 'dist' : 'corp'; }
  function _abrev(compCode) { const c = V.comps[compCode]; return (c && c.abrev) ? c.abrev : compCode; }
  // Nombre completo de la compañía (para el filtro "Compañía") — aplica a corp y distribuible.
  function _compName(compCode) { const c = V.comps[compCode]; return (c && c.nombre) ? c.nombre : ((c && c.abrev) ? c.abrev : (compCode || '(sin compañía)')); }
  // Registro v1-like: {id,src,comp,ceco,vp,_vp0,ger,item,tc,ap,cl,hist,fy26,...}
  function buildRecords() {
    const cm = _cmap();
    const other = _cecoMode === 'old' ? (V.cecoNew || {}) : (V.cecoOld || {});
    return V.records.map((r, i) => {
      // Estructura CECOS: se usa el mapa ACTIVO. Si el CECO NO está en el activo pero SÍ
      // en el otro (ej. 1000AC0502 en Antigua pero no en Nueva), se OCULTA (no pertenece
      // a la estructura elegida). Si no está en ninguno, se deja "(sin VP)" (data-only).
      let m = cm[r.ceco], hidden = false;
      if (!m) {
        if (other[r.ceco]) { hidden = true; m = other[r.ceco]; }
        else m = { vp: '(sin VP)', ger: '(sin Gerencia)', tc: null, ap: null, cl: '(sin clasificación)', comp: (r.ceco || '').slice(0, 4) };
      }
      const src = _srcOf(m.comp);
      const comp = src === 'dist' ? _abrev(m.comp) : null;
      const y = o => ({ real: _pick(o && o.real), plan: _pick(o && o.plan) });
      // Valores CRUDOS {n,a} (independientes del modo activo) — para el "Efecto Moneda".
      const na = c => ({ n: (c && c.n) || 0, a: (c && c.a) || 0 });
      const seed = r.ceco + '|' + r.item + '|' + i;
      return {
        id: i, src, comp, compania: _compName(m.comp), ceco: r.ceco, contra: r.contra || '(sin contrapartida)',
        _hidden: hidden,   // CECO fuera de la estructura activa → excluido del árbol/filtros
        vp: m.vp, _vp0: m.vp, ger: m.ger, _ger0: m.ger, dceco: m.dceco || r.ceco, _dceco0: m.dceco || r.ceco, item: r.item,
        // Ítem Relevante = Agrupación3 del CLACO (nivel padre del ítem/Agrupación4).
        itemrel: (V.itemRel && V.itemRel[r.item]) || r.item,
        // Tipo Costo (C1/C2/C3/TRASPASOS) desde el CLACO (V.itemTc), no del CECO.
        tc: (V.itemTc && V.itemTc[r.item]) || null, ap: m.ap || null, cl: m.cl || '(sin clasificación)',
        // Clasificación Cuenta del CLACO (Total Opex / Mano de Obra / …) — para filtrar.
        clas: (V.itemClas && V.itemClas[r.item]) || null,
        st: r.st || 0,   // Services & Tech (CLACOs 6125020/6125021) — solo Ppto/Forecast (Real no los distingue)
        claco: r.claco || null,   // Clase de Costo (código) — para filtrar por CLACO
        hist: { 2022: y(r.y2022), 2023: y(r.y2023), 2024: y(r.y2024), 2025: y(r.y2025), 2026: y(r.y2026) },
        fy26: _pick(r.y2026fy && r.y2026fy.plan),
        fcst: _pick(r.fcst),   // Forecast 5+7 2026 (anual) — 0 salvo en registros "(Forecast)"
        prop27: r.prop27 || 0, // Ppto 2027 (anual, valor tal cual) — 0 salvo en "(Ppto2027)"
        _realNA: { 2022: na(r.y2022 && r.y2022.real), 2023: na(r.y2023 && r.y2023.real), 2024: na(r.y2024 && r.y2024.real), 2025: na(r.y2025 && r.y2025.real), 2026: na(r.y2026 && r.y2026.real) },
        _planNA: { 2022: na(r.y2022 && r.y2022.plan), 2023: na(r.y2023 && r.y2023.plan), 2024: na(r.y2024 && r.y2024.plan), 2025: na(r.y2025 && r.y2025.plan), 2026: na(r.y2026 && r.y2026.plan) },
        _fy26NA: na(r.y2026fy && r.y2026fy.plan),
        _fcstNA: na(r.fcst),
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
  // Registros vivos (se mutan in-place al cambiar de modo, para no invalidar A.records).
  const dotRecords = normalizeDot(DOT.records, V.records.length);
  const records = [], corpRecords = [], distRecords = [];
  function rebuildDict() {
    const cm = _cmap();
    window.CORP_DICT = { cecos: Object.keys(cm).map(ceco => {
      const m = cm[ceco];
      return { c: ceco, g: m.ger, v: m.vp, d: m.dceco || ceco, tc: m.tc || null, ap: m.ap || null, cl: m.cl || null };
    }) };
  }
  function rebuild() {
    const g = buildRecords();
    records.length = 0; corpRecords.length = 0; distRecords.length = 0;
    // records mantiene TODOS (índice = id, para recordById); corp/dist excluye los _hidden.
    g.forEach(r => { records.push(r); if (!r._hidden) (r.src === 'dist' ? distRecords : corpRecords).push(r); });
    dotRecords.forEach(r => records.push(r));
    rebuildDict();
  }
  function setValMode(m) { if ((m === 'n' || m === 'a') && m !== _valMode) { _valMode = m; rebuild(); } }
  function setCecoMode(m) { if ((m === 'new' || m === 'old') && m !== _cecoMode) { _cecoMode = m; rebuild(); } }
  rebuild();  // init

  // Conjunto activo según modo de datos. Gasto: corp | dist | both (+ compañías).
  // Dotaciones: propios | contratista | dot (ambas).
  // Predicado de filtro por categoría (VP/Ger/Ítem/TipoCosto/Aplica) — compartido
  // por buildTree, annualSeries y efectoMoneda para que no se desincronicen.
  function _matchOpts(opts) {
    return rec => {
      if (opts.vps && opts.vps.length && !opts.vps.includes(rec.vp)) return false;
      if (opts.gers && opts.gers.length && !opts.gers.includes(rec.ger)) return false;
      if (opts.itemrels && opts.itemrels.length && !opts.itemrels.includes(rec.itemrel)) return false;
      if (opts.cecos && opts.cecos.length && !opts.cecos.includes(rec.ceco)) return false;
      if (opts.clacos && opts.clacos.length && !opts.clacos.includes(rec.claco)) return false;
      if (opts.items && opts.items.length && !opts.items.includes(rec.item)) return false;
      if (opts.tcs && opts.tcs.length && !opts.tcs.includes(rec.tc)) return false;
      if (opts.clases && opts.clases.length && rec.clas && !opts.clases.includes(rec.clas)) return false;   // solo filtra los que TIENEN clasificación
      if (opts.companias && opts.companias.length && !opts.companias.includes(rec.compania)) return false;
      if (opts.st === 'only' && !rec.st) return false;   // Solo Services & Tech
      if (opts.st === 'excl' && rec.st) return false;     // Sin Services & Tech
      if (opts.aps && opts.aps.length && !opts.aps.includes(rec.ap)) return false;
      // Ocultar valores puntuales: opts.hidden = { dim: [valores] } → excluye esos registros.
      if (opts.hidden) { for (const d in opts.hidden) { const a = opts.hidden[d]; if (a && a.length && a.indexOf(rec[d]) >= 0) return false; } }
      return true;
    };
  }
  // Efecto Moneda: suma Real (y Ppto) en AMBAS bases {n,a} sobre el MISMO universo
  // filtrado que la tabla, para reportar cuánto cambia el total al reexpresar a 2027.
  function efectoMoneda(opts) {
    const years = (opts.years || []).filter(y => typeof y === 'number');
    const wantFY = (opts.years || []).includes('2026fy');
    const wantFcst = (opts.years || []).includes('2026fcst');
    const filtered = activeRecords(opts).filter(_matchOpts(opts));
    let realN = 0, realA = 0, pptoN = 0, pptoA = 0;
    for (const rec of filtered) {
      if (!rec._realNA) continue;   // dotaciones (FTE) no tienen doble moneda
      for (const y of years) {
        const rc = rec._realNA[y]; if (rc) { realN += rc.n; realA += rc.a; }
        const pc = rec._planNA[y]; if (pc) { pptoN += pc.n; pptoA += pc.a; }
      }
      if ((wantFY || wantFcst) && rec._fy26NA) { pptoN += rec._fy26NA.n; pptoA += rec._fy26NA.a; }   // Ppto FY una vez
      if (wantFcst && rec._fcstNA) { realN += rec._fcstNA.n; realA += rec._fcstNA.a; }   // Forecast → Real
    }
    return { realN, realA, pptoN, pptoA };
  }

  // ===== Detalle del gasto (v3): niveles lazy Contrapartida › Texto pedido › Denominación =====
  // window.DET = { s:[strings], k:{ 'item\u0001contra': [[cecoIdx,textoIdx,denomIdx,docIdx,bk,n,a], …] } }
  // Solo Real (el Ppto no tiene detalle). Se materializa bajo demanda al expandir una
  // Contrapartida, respetando años seleccionados, valMode y (si la estructura lo tiene) CECO.
  const DET = window.DET || null;
  const _BKIDX = { 2022: 0, 2023: 1, 2024: 2, 2025: 3, 2026: 4 };   // año → bucket del detalle
  const _IDX2Y = [2022, 2023, 2024, 2025, 2026];                    // bucket → año (2026 = YTD)
  function hasDetail() { return !!DET; }
  function _stSkip(r, stMode) { return (stMode === 'only' && !r[7]) || (stMode === 'excl' && !!r[7]); }
  function hasDetailFor(ctx, stMode) {
    if (!DET || !ctx) return false;
    const rows = DET.k[ctx.item + '\u0001' + ctx.contra];
    if (!rows) return false;
    if (!stMode) return true;
    for (const r of rows) { if (!_stSkip(r, stMode)) return true; }
    return false;
  }
  function _detZero() { return { real: 0, version: 0, ytdReal: 0, ytdVersion: 0, fy26: 0, prop: 0, yr: {} }; }
  function _addYr(agg, y, v) { const c = agg.yr[y] || (agg.yr[y] = { real: 0, ver: 0 }); c.real += v; }
  // Búsqueda dentro del detalle: precalcula (y cachea por query) qué claves ítem×contra
  // tienen alguna línea cuyo Texto pedido / Denominación / Documento contenga el texto.
  let _detHit = { q: undefined, withCeco: null, noCeco: null };
  function detailHitSets(q, stMode) {
    const _hk = q + '|' + (stMode || '');
    if (_detHit.q === _hk) return _detHit;
    const withCeco = new Set(), noCeco = new Set();
    const ql = (q || '').toLowerCase();
    if (DET && ql) {
      const S = DET.s;
      for (const key in DET.k) {
        let any = false;
        for (const r of DET.k[key]) {
          if (_stSkip(r, stMode)) continue;
          if (S[r[1]].toLowerCase().includes(ql) || S[r[2]].toLowerCase().includes(ql) || String(S[r[3]]).toLowerCase().includes(ql)) {
            any = true; withCeco.add(S[r[0]] + '\u0001' + key);
          }
        }
        if (any) noCeco.add(key);
      }
    }
    _detHit = { q: _hk, withCeco, noCeco };
    return _detHit;
  }
  function detailMatch(ctx, q, stMode) {
    if (!DET || !ctx) return false;
    const h = detailHitSets(q, stMode);
    const key = ctx.item + '\u0001' + ctx.contra;
    return ctx.ceco ? h.withCeco.has(ctx.ceco + '\u0001' + key) : h.noCeco.has(key);
  }
  // order: 'td' = Texto pedido > Denominacion (default) . 'dt' = Denominacion > Texto pedido.
  function detailNodes(ctx, years, q, order, stMode) {
    if (!DET || !ctx) return [];
    const ql = (q || '').toLowerCase();
    const rows = DET.k[ctx.item + '\u0001' + ctx.contra];
    if (!rows) return [];
    const S = DET.s;
    const cecoC = ctx.ceco || null;
    const yb = new Set((years || []).filter(y => typeof y === 'number').map(y => _BKIDX[y]));
    const vi = _valMode === 'a' ? 6 : 5;
    const dt = order === 'dt';
    const i1 = dt ? 2 : 1, d1 = dt ? 'denom' : 'texto';   // nivel 1 (padre)
    const i2 = dt ? 1 : 2, d2 = dt ? 'texto' : 'denom';   // nivel 2 (hoja, lleva el Documento)
    const byL1 = new Map();
    for (const r of rows) {
      if (_stSkip(r, stMode)) continue;
      if (cecoC && S[r[0]] !== cecoC) continue;
      if (yb.size && !yb.has(r[4])) continue;
      const v = r[vi]; if (!v) continue;
      const y = _IDX2Y[r[4]];
      const texto = S[r[1]], denom = S[r[2]], doc = S[r[3]];
      if (ql && !(texto.toLowerCase().includes(ql) || denom.toLowerCase().includes(ql) || String(doc).toLowerCase().includes(ql))) continue;
      const n1 = S[r[i1]], n2 = S[r[i2]];
      let g1 = byL1.get(n1);
      if (!g1) { g1 = { name: n1, _dim: d1, agg: _detZero(), recIds: [], _c: new Map() }; byL1.set(n1, g1); }
      g1.agg.ytdReal += v; g1.agg.real += v; _addYr(g1.agg, y, v);
      const dk = n2 + '\u0001' + doc;
      let g2 = g1._c.get(dk);
      if (!g2) { g2 = { name: n2, _dim: d2, doc: doc, agg: _detZero(), recIds: [], children: [] }; g1._c.set(dk, g2); }
      g2.agg.ytdReal += v; g2.agg.real += v; _addYr(g2.agg, y, v);
    }
    const out = [];
    byL1.forEach(g1 => {
      g1.children = [...g1._c.values()].sort((a, b) => b.agg.ytdReal - a.agg.ytdReal);
      delete g1._c; out.push(g1);
    });
    out.sort((a, b) => b.agg.ytdReal - a.agg.ytdReal);
    return out;
  }

  // ---- Detalle Ppto / Forecast (Concepto Gasto › Actividad) — window.DETP ----
  // Línea: [cecoIdx, cgIdx, actIdx, medida, n, a]. medida: 0 Ppto2025·1 Ppto2026YTD·2 Ppto2026FY·3 Fcst·4 Ppto2027.
  const DETP = window.DETP || null;
  function hasDetailP() { return !!DETP; }
  // El detalle Ppto puede colgar de CUALQUIER último nivel (Ítem, CECO, Ítem Relevante, Desc.
  // CECO, Gerencia, VP). Deriva la ruta de cada línea (item→itemrel, ceco→vp/ger/dceco desde el
  // mapa activo) y la acota al contexto del nodo donde cuelga.
  function _detpItemKeys(ctx) {
    if (ctx.item) return DETP.k[ctx.item] ? [ctx.item] : [];
    if (ctx.itemrel) return Object.keys(DETP.k).filter(itk => ((V.itemRel && V.itemRel[itk]) || itk) === ctx.itemrel);
    return Object.keys(DETP.k);
  }
  function _detpCecoOk(ceco, ctx) {
    if (ctx.ceco && ceco !== ctx.ceco) return false;
    if (ctx.vp || ctx.ger || ctx.dceco) {
      const m = _cmap()[ceco]; if (!m) return false;   // CECO fuera de la estructura activa → no cuenta
      if (ctx.vp && m.vp !== ctx.vp) return false;
      if (ctx.ger && m.ger !== ctx.ger) return false;
      if (ctx.dceco && (m.dceco || ceco) !== ctx.dceco) return false;
    }
    return true;
  }
  function hasDetailPFor(ctx) {
    if (!DETP || !ctx) return false;
    const S = DETP.s;
    for (const itk of _detpItemKeys(ctx)) {
      for (const r of DETP.k[itk]) { if (_detpCecoOk(S[r[0]], ctx)) return true; }
    }
    return false;
  }
  function _detpZero() { return { real: 0, version: 0, ytdReal: 0, ytdVersion: 0, fy26: 0, fcst: 0, prop: 0, yr: {}, _sort: 0 }; }
  function _detpAdd(agg, meas, v) {
    if (meas === 0) { (agg.yr[2025] || (agg.yr[2025] = { real: 0, ver: 0 })).ver += v; }
    else if (meas === 1) { (agg.yr[2026] || (agg.yr[2026] = { real: 0, ver: 0 })).ver += v; }
    else if (meas === 2) { agg.fy26 += v; }
    else if (meas === 3) { agg.fcst += v; }
    else if (meas === 4) { agg.prop += v; }
    agg._sort += v;
  }
  function detailMatchP(ctx, q) {
    if (!DETP || !ctx || !q) return false;
    const S = DETP.s, ql = q.toLowerCase();
    for (const itk of _detpItemKeys(ctx)) {
      for (const r of DETP.k[itk]) {
        if (!_detpCecoOk(S[r[0]], ctx)) continue;
        if (S[r[1]].toLowerCase().includes(ql) || S[r[2]].toLowerCase().includes(ql)) return true;
      }
    }
    return false;
  }
  function detailNodesP(ctx, q) {
    if (!DETP || !ctx) return [];
    const S = DETP.s;
    const ql = (q || '').toLowerCase();
    const vi = _valMode === 'a' ? 5 : 4;   // n=4 · a=5
    const byCG = new Map();
    for (const itk of _detpItemKeys(ctx)) for (const r of DETP.k[itk]) {
      if (!_detpCecoOk(S[r[0]], ctx)) continue;
      const v = r[vi]; if (!v) continue;
      const cg = S[r[1]], act = S[r[2]];
      if (ql && !(cg.toLowerCase().includes(ql) || act.toLowerCase().includes(ql))) continue;
      let g1 = byCG.get(cg);
      if (!g1) { g1 = { name: cg, _dim: 'concepto', fam: 'ppto', agg: _detpZero(), recIds: [], _c: new Map() }; byCG.set(cg, g1); }
      _detpAdd(g1.agg, r[3], v);
      let g2 = g1._c.get(act);
      if (!g2) { g2 = { name: act, _dim: 'actividad', fam: 'ppto', agg: _detpZero(), recIds: [], children: [] }; g1._c.set(act, g2); }
      _detpAdd(g2.agg, r[3], v);
    }
    const out = [];
    byCG.forEach(g1 => { g1.children = [...g1._c.values()].sort((a, b) => b.agg._sort - a.agg._sort); delete g1._c; out.push(g1); });
    out.sort((a, b) => b.agg._sort - a.agg._sort);
    return out;
  }

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
    // Sin simulación: 2026 no tiene datos y la Presupuesto 2027 se ingresa manual
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
    // Presupuesto 2027 ya no es un modo excluyente: es una capa de comparación
    // aditiva. Con showProp, cada nodo lleva además 'prop' (propuesta) junto a
    // las métricas de ejecución; el promedio de comparación = ytdReal / nYears.
    const showProp = !!opts.showProp;
    const years = (opts.years || []).filter(y => typeof y === 'number');
    // '2026fy' (pseudo-año del filtro Año): suma el Ppto ANUAL 2026 (fy26) al
    // Presupuesto; NO aporta Real (es solo presupuesto).
    const wantFY = (opts.years || []).includes('2026fy');
    // '2026fcst' (pseudo-año): suma el Forecast 5+7 2026 (anual) al REAL (es la estimación
    // del año); permite comparar Forecast vs Ppto FY. NO aporta Ppto.
    const wantFcst = (opts.years || []).includes('2026fcst');
    // El Plan del BBDD es el Presupuesto Original; no se aplica ningún ajuste.
    const versionAdj = p => p;

    const filtered = activeRecords(opts).filter(_matchOpts(opts));

    // Modo "Por año": además del total, cada nodo lleva el desglose Real/Ppto año a año.
    const byYear = opts.yearAgg === 'byYear';

    // métrica por registro: suma de los años seleccionados (acumulado anual).
    // Con showProp se anexa la Presupuesto 2027 (override o base). Sin prorrateo.
    function mFor(rec) {
      let real = 0, ver = 0;
      const yr = byYear ? {} : null;
      for (const y of years) {
        const r = rec.hist[y].real, v = versionAdj(rec.hist[y].plan);
        real += r; ver += v;
        if (byYear) yr[y] = { real: r, ver: v };
      }
      if (wantFcst) {                     // Forecast 5+7 2026 → Real, comparado contra el
        real += rec.fcst || 0;            // Ppto FY 2026 (ambos año completo). fy26 una sola
        ver += rec.fy26 || 0;             // vez: el else evita doble si además está wantFY.
        if (byYear) yr['2026fcst'] = { real: rec.fcst || 0, ver: rec.fy26 || 0 };
      } else if (wantFY) {                // 2026 Ppto FY: solo Ppto, sin Real
        ver += rec.fy26 || 0;
        if (byYear) yr['2026fy'] = { real: 0, ver: rec.fy26 || 0 };
      }
      const m = { real, version: ver, ytdReal: real, ytdVersion: ver, fy26: rec.fy26 || 0, fcst: rec.fcst || 0 };
      // Ppto 2027 = valor cargado del Excel (rec.prop27); un override editado tiene prioridad.
      if (showProp) m.prop = (opts.overrides && opts.overrides[rec.id] != null) ? opts.overrides[rec.id] : (rec.prop27 || 0);
      if (byYear) m.yr = yr;
      return m;
    }
    const zero = () => {
      const z = showProp
        ? { real: 0, version: 0, ytdReal: 0, ytdVersion: 0, fy26: 0, fcst: 0, prop: 0 }
        : { real: 0, version: 0, ytdReal: 0, ytdVersion: 0, fy26: 0, fcst: 0 };
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
    const DISPM = { vp: dispVP, ger: dispGer, dceco: n => n, itemrel: dispItemRel, item: dispItem, ceco: n => n, contra: n => n, claco: dispClaco };

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
    const filtered = activeRecords(opts).filter(_matchOpts(opts));
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
      const prop = (opts.overrides && opts.overrides[rec.id] != null) ? opts.overrides[rec.id] : (rec.prop27 || 0);
      out[2027].prop += prop;
    });
    return out;
  }

  function distribuible(opts) {
    // Distribuible REAL por compañía para la vista (año completo).
    // Filtra por compañía/VP/ítem; no por gerencia (los CECO del distribuible
    // no equivalen a las gerencias del corporativo).
    const cset = opts.companies && opts.companies.length ? new Set(opts.companies) : null;
    // Métrica del donut: 'real' (años seleccionados) o 'prop' (Presupuesto 2027).
    const useProp = opts.donutMetric === 'prop';
    const years = (opts.years || []).filter(y => typeof y === 'number');
    const filtered = distRecords.filter(rec => {
      if (cset && !cset.has(rec.comp)) return false;
      if (opts.vps && opts.vps.length && !opts.vps.includes(rec.vp)) return false;
      if (opts.items && opts.items.length && !opts.items.includes(rec.item)) return false;
      if (opts.tcs && opts.tcs.length && !opts.tcs.includes(rec.tc)) return false;
      if (opts.clases && opts.clases.length && rec.clas && !opts.clases.includes(rec.clas)) return false;
      if (opts.aps && opts.aps.length && !opts.aps.includes(rec.ap)) return false;
      return true;
    });
    const out = { MLP: 0, ANT: 0, CEN: 0, CMZ: 0 };
    filtered.forEach(rec => {
      let b;
      if (useProp) b = (opts.overrides && opts.overrides[rec.id] != null) ? opts.overrides[rec.id] : (rec.prop27 || 0);
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
    // unit 'num' = valor crudo (FTE / Nº); 'USD' = dólares (sin dividir); 'kUSD' = miles; resto = millones.
    const div = (unit === 'num' || unit === 'USD') ? 1 : unit === 'kUSD' ? 1e3 : 1e6;
    let v = usd / div;
    const d = decimals != null ? decimals : (unit === 'num' ? 0 : (unit === 'USD' || unit === 'kUSD') ? 0 : 1);
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
  // En v2 la VP y la Gerencia ya vienen con nombre completo desde CECOS; el Ítem se
  // muestra por V.itemNames (Cód_Agrupación2 → nombre). Overrides de Gerencia/Ítem
  // editables por el usuario (persisten en localStorage).
  let _gerOv = {}, _itemOv = {};
  function applyNameOverrides(ov) {
    _gerOv = (ov && ov.ger) || {};
    _itemOv = (ov && ov.item) || {};
  }
  function baseGer(n) { return n; }
  function baseItem(n) { return V.itemNames[n] || n; }
  function dispVP(n) { return n; }
  function dispGer(n) { return _gerOv[n] || n; }
  function dispItem(n) { return _itemOv[n] || V.itemNames[n] || n; }
  // Ítem Relevante = Agrupación3 (nombre por V.relNames; la clave es el código Ag3).
  function dispItemRel(n) { return (V.relNames && V.relNames[n]) || V.itemNames[n] || n; }
  // Desc. CLACO (Clase de Costo): la clave es el código; el nombre viene de V.clacoNames.
  function dispClaco(n) { return (n && V.clacoNames && V.clacoNames[n]) || n || '(sin CLACO)'; }
  // El "código" del Ítem/Ítem Relevante ES la clave. '(sin ítem)' → null.
  function itemCode(n) { return (n && n !== '(sin ítem)') ? n : null; }

  // Opciones de Gerencia (dependen de VP) e Ítem según el modo de datos activo.
  function dimsFor(opts) {
    const recs = activeRecords(opts);
    const vpset = opts.vps && opts.vps.length ? new Set(opts.vps) : null;
    const gers = [], gseen = new Set();
    const items = [], iseen = new Set();
    const itemrels = [], irseen = new Set();   // Ítem Relevante (Agrupación3)
    const tcs = [], tseen = new Set();   // Tipo Costo (C1/C3/Comercialización)
    const clases = [], cseen = new Set();   // Clasificación Cuenta (Total Opex / Mano de Obra / …)
    const companias = [], compseen = new Set();   // Compañía (Corporativo / Los Pelambres / …)
    const cecos = [], ececoseen = new Set();   // Código CECO
    const clacos = [], clseen = new Set();   // Código CLACO (Clase de Costo)
    const aps = [], aseen = new Set();   // ¿Aplica? (Sí/No)
    recs.forEach(r => {
      if (!vpset || vpset.has(r.vp)) { if (!gseen.has(r.ger)) { gseen.add(r.ger); gers.push(r.ger); } }
      if (!iseen.has(r.item)) { iseen.add(r.item); items.push(r.item); }
      if (r.itemrel && !irseen.has(r.itemrel)) { irseen.add(r.itemrel); itemrels.push(r.itemrel); }
      if (r.tc && !tseen.has(r.tc)) { tseen.add(r.tc); tcs.push(r.tc); }
      if (r.clas && !cseen.has(r.clas)) { cseen.add(r.clas); clases.push(r.clas); }
      if (r.compania && !compseen.has(r.compania)) { compseen.add(r.compania); companias.push(r.compania); }
      if (r.ceco && !ececoseen.has(r.ceco)) { ececoseen.add(r.ceco); cecos.push(r.ceco); }   // lista COMPLETA de CECOs
      if (r.claco && !clseen.has(r.claco)) { clseen.add(r.claco); clacos.push(r.claco); }
      if (r.ap && !aseen.has(r.ap)) { aseen.add(r.ap); aps.push(r.ap); }
    });
    gers.sort((a, b) => a.localeCompare(b, 'es'));
    items.sort((a, b) => dispItem(a).localeCompare(dispItem(b), 'es'));
    itemrels.sort((a, b) => dispItemRel(a).localeCompare(dispItemRel(b), 'es'));
    tcs.sort((a, b) => a.localeCompare(b, 'es'));
    clases.sort((a, b) => a.localeCompare(b, 'es'));
    companias.sort((a, b) => a.localeCompare(b, 'es'));
    cecos.sort((a, b) => a.localeCompare(b, 'es'));
    clacos.sort((a, b) => a.localeCompare(b, 'es'));
    aps.sort((a, b) => a.localeCompare(b, 'es'));
    return { gers, items, itemrels, tcs, clases, companias, cecos, clacos, aps };
  }

  // CLACOs presentes tras aplicar TODOS los filtros MENOS el de CLACO (mismo predicado que la
  // tabla). Sirve para tildar en el dropdown "Código CLACO" justo los que se están mostrando.
  function clacosInScope(opts) {
    const o = Object.assign({}, opts, { clacos: [] });
    const set = new Set();
    activeRecords(o).filter(_matchOpts(o)).forEach(r => { if (r.claco) set.add(r.claco); });
    return [...set].sort((a, b) => a.localeCompare(b, 'es'));
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
  // Override de VP por CECO (editable en el Diccionario). La VP ya es el nombre
  // completo, así que se aplica directo por CECO. Sin override → VP original (_vp0).
  function applyVpOverrides(ovByCeco) {
    ovByCeco = ovByCeco || {};
    records.forEach(r => { r.vp = (r.ceco && ovByCeco[r.ceco]) || r._vp0; });
  }

  // Renombre POR CECO de Gerencia y Desc. CECO (editable en el Diccionario). A diferencia del
  // rename por nombre (compartido), acá el valor es texto libre y afecta SOLO al CECO editado:
  // muta rec.ger / rec.dceco de ese CECO (así el árbol lo agrupa/muestra con su nuevo nombre).
  // ov = { ger: {ceco: nombre}, dceco: {ceco: nombre} }. Sin override → original (_ger0/_dceco0).
  function applyCecoNameOverrides(ov) {
    const g = (ov && ov.ger) || {}, d = (ov && ov.dceco) || {};
    records.forEach(r => {
      r.ger = (r.ceco && g[r.ceco]) || r._ger0;
      r.dceco = (r.ceco && d[r.ceco]) || r._dceco0;
    });
  }

  window.CORP = {
    V, D: Object.assign({ vps: [], gers: [] }, V),
    records, corpRecords, distRecords, dotRecords, MESES, COMPANIAS, VISTAS, VERSIONES,
    buildTree, annualSeries, distribuible, activeRecords, efectoMoneda, dimsFor, clacosInScope, applyVpOverrides, applyCecoNameOverrides, kpiColor, kpiHex, color, theme, resetTheme,
    hasDetail, hasDetailFor, detailNodes, detailMatch,
    hasDetailP, hasDetailPFor, detailNodesP, detailMatchP,   // detalle Ppto/Forecast (Concepto Gasto › Actividad)
    hasST: () => records.some(r => r.st),   // ¿hay registros Services & Tech? (para mostrar el filtro)
    fmt, fmtPct, dispVP, dispGer, dispItem, dispItemRel, dispClaco, itemCode, applyNameOverrides, baseGer, baseItem,
    setValMode, setCecoMode, getValMode: () => _valMode, getCecoMode: () => _cecoMode,
    recordById: (id) => records[id],
    derived, DEF_GROWTH, GER_ALIAS,
  };
})();
