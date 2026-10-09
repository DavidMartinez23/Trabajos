# Metodologías para tableros de BI – Actividad grupal

| Punto | Entregable | Ubicación |
|---|---|---|
| 1 | Resumen de metodologías (PDF, 3 páginas) | `entregables/Resumen_Metodologias_BI.pdf` |
| 2 | ETL de `olist_orders_dataset.csv` (Power Query) | Dentro del proyecto de Power BI, consulta **Pedidos** |
| 3 | Tablero de control con metodología **BPM** (4 páginas) | `tablero/Tablero_Pedidos_BPM.pbip` |
| — | Guía del ETL, del modelo y de cómo abrir el tablero | `entregables/Guia_ETL_y_Tablero.pdf` |

## Cómo obtener el archivo .pbix

1. Descomprimir el repositorio para que los datos queden en `C:\Trabajos\datos\originales\`
   (o cambiar el parámetro **RutaDatos** en *Transformar datos → Editar parámetros*).
2. Abrir `tablero\Tablero_Pedidos_BPM.pbip` con Power BI Desktop y pulsar **Actualizar**.
3. **Archivo → Guardar como → Archivo de Power BI (.pbix)**.

## Estructura

- `datos/originales/` – CSV original de pedidos (Olist).
- `tablero/` – proyecto de Power BI (modelo TMDL + reporte PBIR).
- `fuentes/` – scripts que generan los PDF y el proyecto, y la réplica del ETL en Python
  (`python fuentes/etl_pedidos_python.py datos/originales/olist_orders_dataset.csv`) con los valores esperados.
