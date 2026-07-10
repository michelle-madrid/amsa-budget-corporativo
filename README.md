# AMSA · Dashboards de Presupuesto Corporativo

Dashboards **autocontenidos** (un solo `.html` que abre offline en Chrome/Edge, sin instalar
nada). El repo tiene la versión antigua congelada (`v1_antiguo/`) y la actual (`v2/` → genera
`v2` y `v3`).

```
/
├── v1_antiguo/   Versión antigua congelada (no se toca).
│
├── v2/           Versión actual. UN pipeline regenera TODO.
│   ├── construir_v2.py        ← ÚNICO script: Excel → parquet → data.js → embebe en HTML
│   ├── src/                   ← código de la app (se embebe en el HTML)
│   │   ├── model.js             datos + cálculo (window.CORP): árbol, filtros, detalle, forecast
│   │   ├── app.jsx              pestañas, KPIs, Tabla Resumen, Diccionario, export
│   │   ├── matrix.jsx           tabla jerárquica (flattenTree + render) y detalle lazy
│   │   ├── filters.jsx          barra de filtros del Dashboard + KpiCards
│   │   ├── charts.jsx           gráficos (barras / cumplimiento / donut)
│   │   └── tweaks-panel.jsx     panel de colores/textos editable
│   ├── uploads/              ← DATA de entrada (Excel) + diccionario/
│   ├── assets/logo_amsa.png  ← se embebe como data: URI
│   ├── Dashboard Corporativo v2.html   ← SALIDA liviana (y a la vez "shell" plantilla, ver abajo)
│   ├── bbdd_v2.parquet, data_v2.js     ← intermedios regenerables
│   └── ...
│
└── v3/           SALIDA con DETALLE del gasto (la genera construir_v2.py; no se edita a mano)
    └── Dashboard Corporativo v3.html
```

## Cómo regenerar todo (replicable desde la base)

```bash
cd v2
python construir_v2.py
```

Requisitos: **Python** (openpyxl, pandas, pyarrow) y **Node.js** (solo para validar el JSX con
el Babel embebido durante el build). Un solo comando regenera: `bbdd_v2.parquet`, `data_v2.js`,
`v2/Dashboard Corporativo v2.html` (+zip) y `v3/Dashboard Corporativo v3.html` (+zip).

### v2 vs v3
| | Archivo | Peso | Contenido |
|---|---|---|---|
| **v2** | `v2/Dashboard Corporativo v2.html` | ~2.3 MB | Liviano, para compartir. Detalle hasta Contrapartida. |
| **v3** | `v3/Dashboard Corporativo v3.html` | ~7 MB | Igual + **detalle línea**: Contrapartida › Texto pedido › Denominación (+ Documento). |

Ambos son **un solo archivo** que abre con doble clic (Chrome/Edge). No dependen de `uploads/`
ni de `assets/` (logo y fuentes embebidos como `data:`; scripts inline, sin `blob:`).

## Entradas (en `v2/uploads/`)
- **Reales Históricos v2.xlsx** — Real corporativo (hoja *Data Consolidada*).
- **Reales Distribuibles Histórico v2.xlsx** — Real distribuibles (otra disposición de columnas).
- **Planes Históricos v2.xlsx** — Presupuesto (hoja *Budget Consolidado Unpivot*).
- **EXPORT_FORECAST_5+7 2026.xlsx** — Forecast 5+7 2026 (hoja *Forecast 5+7 Unpivot*).
- **Dotaciones Histórico AMSA.xlsx** — FTE Propios/Contratista.
- **diccionario/** — `CECOS.xlsx`, `CLACOS.xlsx`, `COMPAÑÍAS.xlsx`.

> Para actualizar los datos: reemplazá los Excel en `v2/uploads/` (mismos nombres) y volvé a
> correr `python construir_v2.py`.

## Cómo funciona (resumen técnico)
1. **Lee** los Excel y cruza CECO→VP/Gerencia y CLACO→Ítem (ver docstring de `construir_v2.py`).
2. **Agrega** por `(CECO, Ítem, Contrapartida, bucket)` → `bbdd_v2.parquet` (BBDD regenerable).
3. **Serializa** a `data_v2.js` (`window.V2_DATA` + dotaciones + logo + forecast; v3 agrega `window.DET`).
4. **Valida** el JSX con el Babel embebido (aborta si hay error).
5. **Embebe** `src/*` + `data.js` en el HTML (gzip+base64 en el manifiesto) y aplica parches
   idempotentes (bootstrap inline, pantalla de carga, CSS de indentación).

**Autocontenido / inline:** al abrir, un bootstrap descomprime los assets y los ejecuta **inline**
(sin `blob:` ni `fetch`) para que funcione desde `file://` aunque el archivo esté descargado. En
la vista previa de SharePoint (que bloquea JS) se muestra una pantalla estática con instrucciones.

**"Shell" / plantilla:** `v2/Dashboard Corporativo v2.html` es a la vez la salida y la plantilla
base del bundler. `construir_v2.py` lo **parchea en sitio** (deja respaldo `.bak`). Está
versionado en git; si se pierde, restaurar desde git o `.bak`.

## Funcionalidades del dashboard
- **Base Moneda:** Normal (USD) vs Ajustada 2027.
- **Estructura CECOS:** Nueva vs Antigua (por defecto Antigua).
- **Datos:** Corporativo / Distribuible.
- **Año:** 2022–2025, 2026 YTD, 2026 Ppto FY, **2026 Forecast 5+7** (suma al Real).
- **Jerarquía:** VP › Gerencia › (CECO) › **Ítem Relevante** (Agrup3) › **Ítem** (Agrup4) ›
  Contrapartida (› Texto pedido › Denominación + Documento, solo v3).
- **Tabla Resumen:** comparador por período con columnas Real / Ppto / **Forecast**; niveles
  reordenables arrastrando el encabezado; buscador que también busca en el detalle (texto,
  denominación, documento) y revela la rama.

## Editar la lógica (puntos centralizados en `construir_v2.py`)
- Columnas de Reales → `COLS_CORP` / `COLS_DIST`
- Columnas del Detalle (v3) → `DET_COLS_CORP` / `DET_COLS_DIST`
- Columnas del Forecast → `FC_CECO` / `FC_CLACO` / `FC_VALN` / `FC_VALA`
- Años/buckets → `YEARS_HIST`, `YTD_2026`, `BUCKET_IDX`
- Bootstrap / pantalla de carga → `NEW_BOOTSTRAP` / `NEW_THUMB`

## v1 (congelado)
`cd v1_antiguo && python construir_bbdd.py` — no modificar salvo necesidad.
