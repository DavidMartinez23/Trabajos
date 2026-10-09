"""Réplica en Python (pandas) del ETL de Power Query aplicado a olist_orders_dataset.csv.

Sirve para verificar los valores que debe mostrar el tablero en Power BI.
Uso: python etl_pedidos_python.py ../datos/originales/olist_orders_dataset.csv
"""
import sys
import pandas as pd

ESTADOS = {"delivered": "Entregado", "shipped": "Enviado", "canceled": "Cancelado",
           "unavailable": "No disponible", "invoiced": "Facturado", "processing": "En proceso",
           "created": "Creado", "approved": "Aprobado"}


def etl(ruta):
    o = pd.read_csv(ruta, dtype=str)
    fechas = ["order_purchase_timestamp", "order_approved_at", "order_delivered_carrier_date",
              "order_delivered_customer_date", "order_estimated_delivery_date"]
    for c in fechas:
        o[c] = pd.to_datetime(o[c])
    o = o.drop_duplicates("order_id")
    o = o[(o.order_purchase_timestamp >= "2017-01-01") & (o.order_purchase_timestamp < "2018-09-01")].copy()
    o["Estado"] = o.order_status.map(ESTADOS)
    P, A, C, D, E = (o[c] for c in fechas)

    calidad = pd.Series("Sin inconsistencias", index=o.index)
    reglas = [  # se evalúan en orden inverso para que gane la primera regla que se cumpla
        ("Entregado sin fecha de entrega", (o.Estado == "Entregado") & D.isna()),
        ("Cancelado con fecha de entrega", (o.Estado == "Cancelado") & D.notna()),
        ("Transportista antes de la compra", C < P),
        ("Transportista antes de la aprobación", C < A),
        ("Entrega antes del transportista", D < C),
        ("Entregado sin fecha de aprobación", (o.Estado == "Entregado") & A.isna()),
    ]
    for nombre, m in reversed(reglas):
        calidad[m.fillna(False)] = nombre
    o["Calidad del registro"] = calidad

    o["Horas de aprobación"] = ((A - P).dt.total_seconds() / 3600).round(2)
    prep = (C - A).dt.total_seconds() / 86400
    o["Días de preparación"] = prep.where(prep >= 0).round(2)
    trans = (D - C).dt.total_seconds() / 86400
    o["Días de transporte"] = trans.where(trans >= 0).round(2)
    entregado = (o.Estado == "Entregado") & D.notna()
    o["Días de ciclo total"] = ((D - P).dt.total_seconds() / 86400).where(entregado).round(2)
    o["Días vs. estimado"] = (D.dt.normalize() - E.dt.normalize()).dt.days.where(entregado)
    o["Estado de entrega"] = "No entregado"
    o.loc[entregado & (o["Días vs. estimado"] <= 0), "Estado de entrega"] = "A tiempo"
    o.loc[entregado & (o["Días vs. estimado"] > 0), "Estado de entrega"] = "Con retraso"
    return o


if __name__ == "__main__":
    o = etl(sys.argv[1])
    ent = o[o["Estado de entrega"] != "No entregado"]
    tarde = o[o["Estado de entrega"] == "Con retraso"]
    print("Pedidos después del ETL:", len(o))
    print("% entregados:", round(len(ent) / len(o) * 100, 2))
    print("% entregas a tiempo:", round((ent["Estado de entrega"] == "A tiempo").mean() * 100, 2))
    print("Pedidos con retraso:", len(tarde), "| días promedio de retraso:", round(tarde["Días vs. estimado"].mean(), 2))
    print("Días promedio vs. estimado (entregados):", round(ent["Días vs. estimado"].mean(), 2))
    for c in ["Horas de aprobación", "Días de preparación", "Días de transporte", "Días de ciclo total"]:
        print(f"{c} (promedio):", round(o[c].mean(), 2))
    print("% cancelados + no disponibles:", round(o.Estado.isin(["Cancelado", "No disponible"]).mean() * 100, 2))
    print("\nEstados:\n", o.Estado.value_counts().to_string())
    print("\nCalidad del registro:\n", o["Calidad del registro"].value_counts().to_string())
    print("\nEtapas alcanzadas: compra", len(o), "| aprobado", o.order_approved_at.notna().sum(),
          "| transportista", o.order_delivered_carrier_date.notna().sum(), "| entregado", len(ent))
    m = o.groupby(o.order_purchase_timestamp.dt.to_period("M"))
    print("\nPor mes (pedidos, % a tiempo):")
    print(pd.DataFrame({"pedidos": m.size(),
                        "%a_tiempo": m.apply(lambda g: round((g["Estado de entrega"] == "A tiempo").sum() /
                                                            max((g["Estado de entrega"] != "No entregado").sum(), 1) * 100, 1))}).to_string())


def indicadores_bsc(o):
    """Valores esperados de los indicadores del Balanced Scorecard (periodo completo)."""
    ent = o[o["Estado de entrega"] != "No entregado"]
    anio = o.order_purchase_timestamp.dt.year
    mes = o.order_purchase_timestamp.dt.month
    a18, a17 = o[(anio == 2018) & (mes <= 8)], o[(anio == 2017) & (mes <= 8)]
    a_tiempo = lambda d: (d["Estado de entrega"] == "A tiempo").sum() / (d["Estado de entrega"] != "No entregado").sum()
    dias = (o.order_purchase_timestamp.max().normalize() - pd.Timestamp("2017-01-01")).days + 1
    return {
        "1.1 Crecimiento interanual de pedidos (2018 vs 2017, ene-ago)": len(a18) / len(a17) - 1,
        "1.2 % pedidos entregados": len(ent) / len(o),
        "1.3 % pedidos perdidos": o.Estado.isin(["Cancelado", "No disponible"]).mean(),
        "2.1 % entregas a tiempo": a_tiempo(o),
        "2.2 Días promedio de retraso": o.loc[o["Estado de entrega"] == "Con retraso", "Días vs. estimado"].mean(),
        "2.3 % retrasos mayores a 7 días": (ent["Días vs. estimado"] > 7).sum() / len(ent),
        "2.4 Días de holgura promedio": -ent["Días vs. estimado"].mean(),
        "3.1 Días promedio de ciclo": o["Días de ciclo total"].mean(),
        "3.2 Horas promedio de aprobación": o["Horas de aprobación"].mean(),
        "3.3 Días promedio de preparación": o["Días de preparación"].mean(),
        "3.4 Días promedio de transporte": o["Días de transporte"].mean(),
        "4.1 % registros consistentes": (o["Calidad del registro"] == "Sin inconsistencias").mean(),
        "4.2 Pedidos promedio por día": len(o) / 608,
        "4.3 Mejora % a tiempo vs año anterior (pp)": (a_tiempo(a18) - a_tiempo(a17)) * 100,
        "(días del calendario)": dias,
    }
