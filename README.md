# Dashboard de Presupuesto Corporativo · AMSA

Dashboard **autocontenido**: un solo archivo `.html` que abre offline en Chrome/Edge con doble
clic, sin instalar nada. Todo el dashboard (código y datos) se genera desde Excel con un único
comando.

```
/
├── tablero/                 El motor: scripts que leen los Excel y arman el dashboard.
│   ├── construir.py           ← script principal (Excel → datos → HTML). Empezá por acá.
│   ├── actualizar.py          Reconstruye el dashboard + copias extra + dashboards por VP.
│   ├── dividir_por_vp.py      Parte el dashboard en un HTML por Vicepresidencia.
│   ├── publicar.py            Copia los HTML por VP a la carpeta compartida (OneDrive).
│   ├── server.py              Panel web para hacer todo lo anterior sin recordar comandos.
│   ├── agregar_real.py        Agrega un Excel de Real nuevo al maestro.
│   ├── unpivot_ejercicio.py   Prepara la hoja «… Unpivot» de un Forecast/Outlook/Ppto.
│   ├── bbdd_a_excel.py        Exporta la base de datos a Excel para compartir.
│   ├── plantilla.html         Molde HTML del que se parte (no se comparte; se restaura de git).
│   ├── src/                   Código de la aplicación (se embebe dentro del HTML).
│   ├── uploads/               DATOS DE ENTRADA (todos los Excel). Ver uploads/LEEME.txt.
│   ├── assets/                Logo (se embebe como data: URI).
│   ├── bbdd/                  La misma data en tablas parquet (para Power BI / DuckDB).
│   └── utilidades/            Scripts sueltos de mantenimiento/recuperación.
│
└── salida/                  Lo que se comparte.
    ├── Dashboard Corporativo.html   ← el dashboard terminado.
    └── por_vp/                      Un HTML por Vicepresidencia.
```

## Regenerar el dashboard

```bash
cd tablero
python construir.py
```

Necesita **Python** (openpyxl, pandas, pyarrow) y **Node.js** (solo para validar el código de la
app durante el armado). El comando deja el dashboard en `salida/Dashboard Corporativo.html`.

> Para actualizar los datos: reemplazá el Excel que cambió en `tablero/uploads/` (mismo nombre) y
> volvé a correr `python construir.py`. Para cambiar el Presupuesto 2027, editá antes `PPTO27_FILE`
> en `construir.py`.

Variante rápida para previsualizar: `python construir.py --rapido` (rearma solo el dashboard, sin
reescribir la plantilla ni los archivos intermedios).

## Panel web (opcional)

En vez de recordar comandos:

```bash
cd tablero
python server.py       # abre http://localhost:5050
```

Permite actualizar Ppto / Forecast / Outlook, agregar Real, y publicar (local, o a OneDrive).

## Dónde tocar para los cambios más frecuentes

Todo está centralizado arriba de `tablero/construir.py`:

| Cuando… | Editá |
|---|---|
| Cambia el nombre del Excel de Ppto 2027 | `PPTO27_FILE` |
| Cambia el Forecast / Outlook | `FCST_FILE` / `OUT_FILE` |
| Se corren las columnas de los Reales | `COLS_CORP` / `COLS_DIST` |
| Columnas del detalle del gasto | `DET_COLS_CORP` / `DET_COLS_DIST` |
| Columnas del Forecast | `FC_CECO` / `FC_CLACO` / `FC_VALN` / `FC_VALA` |
| Años y períodos a mostrar | `YEARS_HIST`, `YTD_2026` |
| Renombrar un CLACO dentro de un CECO | `CLACO_REMAP` |
| Forzar todo un CECO a un único CLACO | `CLACO_FORCE_CECO` — `{CECO: CLACO}` |
| Comercialización a mostrar (filtro "Otros Fletes") | `SHOW_COMERCIAL` — `{CECO: [CLACO, …]}` (whitelist de la VP Comercialización) |
| Reclasificar la Clasificación Cuenta | `CLAS_CUENTA_OVERRIDE` |

Los destinos de OneDrive están en `EXTRA_DESTS` (`actualizar.py`) y `DEST` (`publicar.py`).

## De dónde salen los datos

Todos los Excel viven en `tablero/uploads/` (detalle en `uploads/LEEME.txt`):

- **Reales Históricos v2.xlsx** — Real corporativo.
- **Reales Distribuibles Histórico v2.xlsx** — Real de los distribuibles (columnas en otro orden).
- **Planes Históricos v2.xlsx** — Presupuesto histórico.
- **EXPORT_FORECAST_5+7 2026.xlsx** — Forecast 5+7 del año.
- **Dotaciones Histórico AMSA.xlsx** — dotación (FTE).
- **diccionario/** — `CECOS.xlsx`, `CLACOS.xlsx`, `COMPAÑÍAS.xlsx`.

## Cómo funciona (resumen técnico)

1. **Lee** los Excel y cruza CECO → VP/Gerencia y CLACO → Ítem.
2. **Agrega** por `(CECO, Ítem, Contrapartida, período)` y guarda la base (`bbdd/`).
3. **Serializa** los datos a JavaScript.
4. **Valida** el código de la app y lo **embebe** —junto con los datos— dentro del HTML.

El HTML resultante es **un solo archivo**: al abrirlo, descomprime todo y corre sin depender de
`uploads/` ni de internet (logo, fuentes y scripts van incrustados). La `plantilla.html` es el molde
del que se parte; `construir.py` la actualiza en sitio dejando un respaldo `.bak`.

## Qué muestra el dashboard

- **Base Moneda:** Normal (USD) vs Ajustada 2027.
- **Estructura CECOS:** Nueva vs Antigua.
- **Datos:** Corporativo / Distribuible.
- **Año:** 2022–2025, 2026 YTD, 2026 Ppto, **2026 Forecast 5+7**.
- **Jerarquía:** VP › Gerencia › (CECO) › Ítem Relevante › Ítem › Contrapartida › Texto pedido ›
  Denominación (+ Documento).
- **Tabla Resumen:** comparador por período (Real / Ppto / Forecast), niveles reordenables
  arrastrando el encabezado, y buscador que también busca dentro del detalle del gasto.
