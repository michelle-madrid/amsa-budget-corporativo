# AMSA · Dashboards de Presupuesto Corporativo

Este repositorio contiene **dos versiones** del dashboard, separadas para no mezclarlas.

```
/
├── v1_antiguo/     Dashboard ACTUAL (versión antigua) — funciona, no se toca
│   ├── Dashboard Actividad Corporativa.html   ← entregable v1
│   ├── uploads/                                ← data v1 (estructura antigua)
│   ├── construir_bbdd.py                       ← pipeline canónico v1 (Excel → data.js → HTML)
│   ├── actualizar_dashboard.py, reembedir_codigo.py, nombres_claco.py,
│   │   forecast.py, cruce_claco.py, agregar_cecos_c3.py
│   ├── app.jsx, model.js, filters.jsx, charts.jsx, matrix.jsx, tweaks-panel.jsx, styles.css
│   ├── assets/, *.parquet, Actualizar Dashboard.bat, COMO *.txt
│
├── v2/             Dashboard NUEVO (en construcción)
│   ├── uploads/                                ← AQUÍ va la data nueva (estructura nueva)
│   │   └── Dotaciones Histórico AMSA.xlsx      ← dotaciones (igual que v1)
│   └── construir_v2.py                         ← (pendiente) ÚNICO python que regenera el v2
│
└── README.md
```

## Reglas
- **v1_antiguo/** es la versión antigua congelada. No se modifica.
- **v2/** es el dashboard nuevo. Todo el tratamiento de datos va en **un solo** `v2/construir_v2.py`, reproducible (Excel de `v2/uploads/` → dashboard).
- **Dotaciones** se mantiene igual: el v2 reutiliza `Dotaciones Histórico AMSA.xlsx` (copiado en `v2/uploads/`).

## Cómo regenerar
- v1: `cd v1_antiguo && python construir_bbdd.py`
- v2: `cd v2 && python construir_v2.py`  *(cuando esté creado)*
