# Metodologías para tableros de BI – Actividad grupal

**Descarga directa:** `entregables/Tablero_BSC_Pedidos.zip` (proyecto de Power BI + datos + PDF).

| Punto | Entregable | Ubicación |
|---|---|---|
| 1 | Resumen de metodologías (PDF, 3 páginas) | `entregables/Resumen_Metodologias_BI.pdf` |
| 2 | ETL de `olist_orders_dataset.csv` (Power Query, revisión del tipo de cada columna) | Proyecto de Power BI, consulta **Pedidos** |
| 3 | Tablero **Balanced Scorecard** (5 páginas, relaciones uno a muchos) | `tablero/Tablero_Pedidos_BSC.pbip` |
| — | Guía del ETL, diccionario de columnas, modelo, indicadores y cómo abrirlo | `entregables/Guia_ETL_y_Tablero_BSC.pdf` |

## Cómo obtener el archivo .pbix

1. Descomprimir `Tablero_BSC_Pedidos.zip` en `C:\` (queda `C:\Trabajos\...`).
   Si se usa otra carpeta, cambiar el parámetro **RutaDatos** en *Transformar datos → Editar parámetros*.
2. Abrir `C:\Trabajos\tablero\Tablero_Pedidos_BSC.pbip` con Power BI Desktop y pulsar **Actualizar**.
3. **Archivo → Guardar como → Archivo de Power BI (.pbix)**.

## Modelo

`Pedidos` (hechos) ← muchos a uno → `Calendario`, `Estados`, `Rangos`.
Tablas desconectadas: `Indicadores` (14 KPI del BSC con meta y tolerancia), `Etapas` y `Medidas` (41 medidas DAX).

## Fuentes

- `fuentes/generar_pbip.py` – genera el proyecto de Power BI (modelo TMDL + reporte PBIR).
- `fuentes/etl_pedidos_python.py` – réplica del ETL en Python con los valores esperados de los indicadores.
- `fuentes/generar_resumen_pdf.py`, `fuentes/generar_guia_pdf.py`, `fuentes/empaquetar_zip.py` – PDF y ZIP.
