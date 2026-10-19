# MDI Armas Ilícitas

Proyecto de análisis de datos sobre registros de armas ilícitas y delitos asociados en Ecuador.

## Contenido

- `mdi_armasilicitas_pm_historico_2017_2025.xlsx` - histórico principal de 2017 a 2025.
- `mdi_armasilicitas_pm_2026_enero_septiembre.xlsx` - corte 2026 de enero a septiembre.
- `mdi_armasilicitas_dd_2024.xlsx` - diccionario de datos 2024.
- `mdi_armasilicitas_dd_2025.xlsx` - diccionario de datos 2025.
- `src/build_report.py` - generador reproducible del notebook y del reporte final.
- `notebooks/01_analisis_armas_ilicitas.ipynb` - notebook con graficos y hallazgos.
- `reports/reporte_final.html` - reporte ejecutivo final en HTML.

## Uso

Estos archivos sirven como base para:

- análisis temporal,
- análisis geográfico,
- análisis por tipo de delito,
- análisis por tipo de arma u objeto,
- control de calidad de datos.

## Reproducir entregables

Si quieres regenerar el notebook y el reporte:

```bash
python src/build_report.py
```

## Notas

- El entorno local `env/` no se versiona en Git.
- Los Excel se mantienen en la raíz del proyecto como insumo fuente.
