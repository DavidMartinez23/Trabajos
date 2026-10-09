"""Genera el proyecto de Power BI (PBIP: modelo TMDL + reporte PBIR) del tablero Balanced Scorecard de pedidos.

Uso: python generar_pbip.py <carpeta_destino>
"""
import hashlib
import json
import shutil
import sys
import uuid
from pathlib import Path

NOMBRE = "Tablero_Pedidos_BSC"
RUTA_DATOS_DEFECTO = "C:\\Trabajos\\datos\\originales\\"
AQUI = Path(__file__).resolve().parent


def gid(*partes):
    """Identificador estable de 20 caracteres hexadecimales (como los que usa Power BI)."""
    return hashlib.sha1("|".join(partes).encode()).hexdigest()[:20]


def guid(*partes):
    return str(uuid.UUID(hashlib.md5("|".join(partes).encode()).hexdigest()))


def escribir(ruta, contenido):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    if not isinstance(contenido, str):
        contenido = json.dumps(contenido, ensure_ascii=False, indent=2)
    ruta.write_text(contenido, encoding="utf-8")


def bloque(texto, tabs):
    """Indenta un bloque de código (M o DAX) con tabulaciones para TMDL."""
    pre = "\t" * tabs
    return "\n".join(pre + l if l.strip() else "" for l in texto.strip("\n").split("\n"))


# ---------------------------------------------------------------------------
# Modelo semántico (TMDL)
# ---------------------------------------------------------------------------

M_PEDIDOS = r'''
let
    // Fuente: archivo CSV de pedidos de Olist (UTF-8, separado por comas)
    Origen = Csv.Document(File.Contents(RutaDatos & "olist_orders_dataset.csv"), [Delimiter = ",", Columns = 8, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    #"Encabezados promovidos" = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),
    // Las fechas faltantes llegan como texto vacío: se reemplazan por null
    #"Vacíos reemplazados por null" = Table.ReplaceValue(#"Encabezados promovidos", "", null, Replacer.ReplaceValue, {"order_approved_at", "order_delivered_carrier_date", "order_delivered_customer_date"}),
    // Limpieza de texto: espacios sobrantes y estados en minúscula
    #"Textos limpiados" = Table.TransformColumns(#"Vacíos reemplazados por null", {{"order_id", Text.Trim, type text}, {"customer_id", Text.Trim, type text}, {"order_status", each Text.Lower(Text.Trim(_)), type text}}),
    // Tipos de fecha con cultura en-US para leer el formato ISO sin depender de la configuración regional
    #"Tipos cambiados" = Table.TransformColumnTypes(#"Textos limpiados", {{"order_purchase_timestamp", type datetime}, {"order_approved_at", type datetime}, {"order_delivered_carrier_date", type datetime}, {"order_delivered_customer_date", type datetime}, {"order_estimated_delivery_date", type datetime}}, "en-US"),
    #"Duplicados quitados" = Table.Distinct(#"Tipos cambiados", {"order_id"}),
    // Solo meses completos: se excluyen 2016 (datos escasos) y sep-oct 2018 (mes incompleto)
    #"Filtrado periodo completo" = Table.SelectRows(#"Duplicados quitados", each [order_purchase_timestamp] >= #datetime(2017, 1, 1, 0, 0, 0) and [order_purchase_timestamp] < #datetime(2018, 9, 1, 0, 0, 0)),
    MapaEstados = [delivered = "Entregado", shipped = "Enviado", canceled = "Cancelado", unavailable = "No disponible", invoiced = "Facturado", processing = "En proceso", created = "Creado", approved = "Aprobado"],
    #"Estados traducidos" = Table.TransformColumns(#"Filtrado periodo completo", {{"order_status", each Record.FieldOrDefault(MapaEstados, _, _), type text}}),
    #"Agregada fecha de compra" = Table.AddColumn(#"Estados traducidos", "purchase_date", each DateTime.Date([order_purchase_timestamp]), type date),
    #"Agregada hora de compra" = Table.AddColumn(#"Agregada fecha de compra", "purchase_hour", each Time.Hour([order_purchase_timestamp]), Int64.Type),
    // Etapa 1 del proceso: compra -> aprobación del pago (horas)
    #"Agregadas horas de aprobación" = Table.AddColumn(#"Agregada hora de compra", "approval_hours", each if [order_approved_at] = null then null else Number.Round(Duration.TotalHours([order_approved_at] - [order_purchase_timestamp]), 2), type nullable number),
    // Etapa 2: aprobación -> entrega al transportista (días); valores negativos son inconsistentes y quedan en null
    #"Agregados días de preparación" = Table.AddColumn(#"Agregadas horas de aprobación", "prep_days", each
        let d = if [order_approved_at] = null or [order_delivered_carrier_date] = null then null else Duration.TotalDays([order_delivered_carrier_date] - [order_approved_at])
        in if d = null or d < 0 then null else Number.Round(d, 2), type nullable number),
    // Etapa 3: transportista -> cliente (días); valores negativos quedan en null
    #"Agregados días de transporte" = Table.AddColumn(#"Agregados días de preparación", "transit_days", each
        let d = if [order_delivered_carrier_date] = null or [order_delivered_customer_date] = null then null else Duration.TotalDays([order_delivered_customer_date] - [order_delivered_carrier_date])
        in if d = null or d < 0 then null else Number.Round(d, 2), type nullable number),
    #"Agregada marca de entrega" = Table.AddColumn(#"Agregados días de transporte", "is_delivered", each [order_status] = "Entregado" and [order_delivered_customer_date] <> null, type logical),
    // Tiempo de ciclo completo: compra -> entrega al cliente
    #"Agregados días de ciclo" = Table.AddColumn(#"Agregada marca de entrega", "cycle_days", each if [is_delivered] then Number.Round(Duration.TotalDays([order_delivered_customer_date] - [order_purchase_timestamp]), 2) else null, type nullable number),
    // Diferencia contra la fecha prometida: negativo = antes de lo estimado, positivo = retraso
    #"Agregados días vs estimado" = Table.AddColumn(#"Agregados días de ciclo", "days_vs_estimate", each if [is_delivered] then Duration.Days(DateTime.Date([order_delivered_customer_date]) - DateTime.Date([order_estimated_delivery_date])) else null, Int64.Type),
    #"Agregado estado de entrega" = Table.AddColumn(#"Agregados días vs estimado", "delivery_status", each if not [is_delivered] then "No entregado" else if [days_vs_estimate] <= 0 then "A tiempo" else "Con retraso", type text),
    #"Agregado rango de retraso" = Table.AddColumn(#"Agregado estado de entrega", "delay_range", each
        if [delivery_status] <> "Con retraso" then [delivery_status]
        else if [days_vs_estimate] <= 3 then "1 a 3 días"
        else if [days_vs_estimate] <= 7 then "4 a 7 días"
        else if [days_vs_estimate] <= 14 then "8 a 14 días"
        else "Más de 14 días", type text),
    // Validación de consistencia de las fechas del proceso (se marca la primera regla que se incumple)
    #"Agregada calidad del registro" = Table.AddColumn(#"Agregado rango de retraso", "record_quality", each
        if [order_status] = "Entregado" and [order_delivered_customer_date] = null then "Entregado sin fecha de entrega"
        else if [order_status] = "Cancelado" and [order_delivered_customer_date] <> null then "Cancelado con fecha de entrega"
        else if [order_delivered_carrier_date] <> null and [order_delivered_carrier_date] < [order_purchase_timestamp] then "Transportista antes de la compra"
        else if [order_delivered_carrier_date] <> null and [order_approved_at] <> null and [order_delivered_carrier_date] < [order_approved_at] then "Transportista antes de la aprobación"
        else if [order_delivered_customer_date] <> null and [order_delivered_carrier_date] <> null and [order_delivered_customer_date] < [order_delivered_carrier_date] then "Entrega antes del transportista"
        else if [order_status] = "Entregado" and [order_approved_at] = null then "Entregado sin fecha de aprobación"
        else "Sin inconsistencias", type text),
    #"Columna auxiliar quitada" = Table.RemoveColumns(#"Agregada calidad del registro", {"is_delivered"}),
    #"Columnas renombradas" = Table.RenameColumns(#"Columna auxiliar quitada", {
        {"order_id", "ID Pedido"}, {"customer_id", "ID Cliente"}, {"order_status", "Estado del pedido"},
        {"order_purchase_timestamp", "Fecha y hora de compra"}, {"order_approved_at", "Fecha de aprobación"},
        {"order_delivered_carrier_date", "Fecha entrega a transportista"}, {"order_delivered_customer_date", "Fecha entrega al cliente"},
        {"order_estimated_delivery_date", "Fecha estimada de entrega"}, {"purchase_date", "Fecha de compra"}, {"purchase_hour", "Hora de compra"},
        {"approval_hours", "Horas de aprobación"}, {"prep_days", "Días de preparación"}, {"transit_days", "Días de transporte"},
        {"cycle_days", "Días de ciclo total"}, {"days_vs_estimate", "Días vs estimado"}, {"delivery_status", "Estado de entrega"},
        {"delay_range", "Rango de retraso"}, {"record_quality", "Calidad del registro"}}),
    // Revisión final del formato de cada columna: texto, fecha/hora, fecha, entero o decimal
    #"Tipos verificados por columna" = Table.TransformColumnTypes(#"Columnas renombradas", {
        {"ID Pedido", type text}, {"ID Cliente", type text}, {"Estado del pedido", type text},
        {"Fecha y hora de compra", type datetime}, {"Fecha de aprobación", type datetime},
        {"Fecha entrega a transportista", type datetime}, {"Fecha entrega al cliente", type datetime},
        {"Fecha estimada de entrega", type date}, {"Fecha de compra", type date}, {"Hora de compra", Int64.Type},
        {"Horas de aprobación", type number}, {"Días de preparación", type number}, {"Días de transporte", type number},
        {"Días de ciclo total", type number}, {"Días vs estimado", Int64.Type}, {"Estado de entrega", type text},
        {"Rango de retraso", type text}, {"Calidad del registro", type text}})
in
    #"Tipos verificados por columna"
'''

M_CALENDARIO = r'''
let
    // Calendario continuo para el periodo analizado (enero 2017 - agosto 2018)
    FechaInicio = #date(2017, 1, 1),
    FechaFin = #date(2018, 8, 31),
    ListaFechas = List.Dates(FechaInicio, Duration.Days(FechaFin - FechaInicio) + 1, #duration(1, 0, 0, 0)),
    #"Convertido en tabla" = Table.FromList(ListaFechas, Splitter.SplitByNothing(), {"Fecha"}),
    #"Tipo cambiado" = Table.TransformColumnTypes(#"Convertido en tabla", {{"Fecha", type date}}),
    #"Agregado año" = Table.AddColumn(#"Tipo cambiado", "Año", each Date.Year([Fecha]), Int64.Type),
    #"Agregado mes número" = Table.AddColumn(#"Agregado año", "Mes número", each Date.Month([Fecha]), Int64.Type),
    #"Agregado mes" = Table.AddColumn(#"Agregado mes número", "Mes", each Text.Proper(Date.MonthName([Fecha], "es-ES")), type text),
    #"Agregado año-mes número" = Table.AddColumn(#"Agregado mes", "Año-Mes número", each [Año] * 100 + [Mes número], Int64.Type),
    #"Agregado mes y año" = Table.AddColumn(#"Agregado año-mes número", "Mes y año", each Text.Start([Mes], 3) & " " & Text.From([Año]), type text),
    #"Agregado trimestre" = Table.AddColumn(#"Agregado mes y año", "Trimestre", each "T" & Text.From(Date.QuarterOfYear([Fecha])), type text),
    #"Agregado número día semana" = Table.AddColumn(#"Agregado trimestre", "Número día semana", each Date.DayOfWeek([Fecha], Day.Monday) + 1, Int64.Type),
    #"Agregado día de la semana" = Table.AddColumn(#"Agregado número día semana", "Día de la semana", each Text.Proper(Date.DayOfWeekName([Fecha], "es-ES")), type text)
in
    #"Agregado día de la semana"
'''

M_ETAPAS = r'''
let
    // Etapas del proceso de pedido para el embudo de la perspectiva de procesos internos
    Etapas = #table(type table [Orden = Int64.Type, Etapa = text], {
        {1, "1. Compra registrada"},
        {2, "2. Pago aprobado"},
        {3, "3. Entregado al transportista"},
        {4, "4. Entregado al cliente"}})
in
    Etapas
'''

M_ESTADOS = r'''
let
    // Dimensión de estados: cada estado agrupa muchos pedidos (relación uno a muchos)
    Estados = #table(type table [#"Estado del pedido" = text, #"Código original" = text, #"Fase del proceso" = text, Orden = Int64.Type], {
        {"Creado", "created", "En curso", 1},
        {"Aprobado", "approved", "En curso", 2},
        {"Facturado", "invoiced", "En curso", 3},
        {"En proceso", "processing", "En curso", 4},
        {"Enviado", "shipped", "En curso", 5},
        {"Entregado", "delivered", "Completado", 6},
        {"Cancelado", "canceled", "Perdido", 7},
        {"No disponible", "unavailable", "Perdido", 8}})
in
    Estados
'''

M_RANGOS = r'''
let
    // Dimensión de rangos de retraso: cada rango agrupa muchos pedidos (relación uno a muchos)
    Rangos = #table(type table [Orden = Int64.Type, #"Rango de retraso" = text, #"Tipo de entrega" = text], {
        {1, "A tiempo", "A tiempo"},
        {2, "1 a 3 días", "Con retraso"},
        {3, "4 a 7 días", "Con retraso"},
        {4, "8 a 14 días", "Con retraso"},
        {5, "Más de 14 días", "Con retraso"},
        {6, "No entregado", "No entregado"}})
in
    Rangos
'''

M_INDICADORES = r'''
let
    // Indicadores del Balanced Scorecard: perspectiva, meta, tolerancia (zona amarilla) y sentido de mejora
    Indicadores = #table(type table [#"Código" = text, Orden = Int64.Type, Perspectiva = text, #"Orden perspectiva" = Int64.Type, Indicador = text, Meta = number, Tolerancia = number, Sentido = text, Formato = text], {
        {"1.1", 1, "Financiera", 1, "Crecimiento interanual de pedidos", 0.20, 0.05, "Mayor es mejor", "%"},
        {"1.2", 2, "Financiera", 1, "% de pedidos entregados (ventas completadas)", 0.97, 0.01, "Mayor es mejor", "%"},
        {"1.3", 3, "Financiera", 1, "% de pedidos perdidos (cancelados o no disponibles)", 0.01, 0.005, "Menor es mejor", "%"},
        {"2.1", 4, "Clientes", 2, "% de entregas a tiempo", 0.95, 0.03, "Mayor es mejor", "%"},
        {"2.2", 5, "Clientes", 2, "Días promedio de retraso", 7, 2, "Menor es mejor", "días"},
        {"2.3", 6, "Clientes", 2, "% de retrasos mayores a 7 días", 0.02, 0.01, "Menor es mejor", "%"},
        {"2.4", 7, "Clientes", 2, "Días de anticipación frente a la fecha prometida", 10, 2, "Mayor es mejor", "días"},
        {"3.1", 8, "Procesos internos", 3, "Días promedio de ciclo (compra a entrega)", 10, 2, "Menor es mejor", "días"},
        {"3.2", 9, "Procesos internos", 3, "Horas promedio de aprobación del pago", 12, 4, "Menor es mejor", "horas"},
        {"3.3", 10, "Procesos internos", 3, "Días promedio de preparación", 3, 1, "Menor es mejor", "días"},
        {"3.4", 11, "Procesos internos", 3, "Días promedio de transporte", 8, 2, "Menor es mejor", "días"},
        {"4.1", 12, "Aprendizaje y crecimiento", 4, "% de registros consistentes (calidad de datos)", 0.99, 0.01, "Mayor es mejor", "%"},
        {"4.2", 13, "Aprendizaje y crecimiento", 4, "Pedidos procesados por día (capacidad)", 150, 20, "Mayor es mejor", "num"},
        {"4.3", 14, "Aprendizaje y crecimiento", 4, "Mejora del % a tiempo frente al año anterior", 0, 0.02, "Mayor es mejor", "pp"}})
in
    Indicadores
'''

COLS_ESTADOS = [
    ("Estado del pedido", "string", None, {"sortByColumn": "Orden", "isKey": True}),
    ("Código original", "string", None, {}),
    ("Fase del proceso", "string", None, {}),
    ("Orden", "int64", "0", {"isHidden": True}),
]

COLS_RANGOS = [
    ("Orden", "int64", "0", {"isHidden": True}),
    ("Rango de retraso", "string", None, {"sortByColumn": "Orden", "isKey": True}),
    ("Tipo de entrega", "string", None, {}),
]

COLS_INDICADORES = [
    ("Código", "string", None, {"isKey": True}),
    ("Orden", "int64", "0", {"isHidden": True}),
    ("Perspectiva", "string", None, {"sortByColumn": "Orden perspectiva"}),
    ("Orden perspectiva", "int64", "0", {"isHidden": True}),
    ("Indicador", "string", None, {"sortByColumn": "Orden"}),
    ("Meta", "double", "0.00", {}),
    ("Tolerancia", "double", "0.00", {}),
    ("Sentido", "string", None, {}),
    ("Formato", "string", None, {"isHidden": True}),
]

# (nombre, tipo, formato, extra) - extra: dict con isHidden / sortByColumn / isKey / summarizeBy
COLS_PEDIDOS = [
    ("ID Pedido", "string", None, {}),
    ("ID Cliente", "string", None, {"isHidden": True}),
    ("Estado del pedido", "string", None, {"isHidden": True}),
    ("Fecha y hora de compra", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha de aprobación", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha entrega a transportista", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha entrega al cliente", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha estimada de entrega", "dateTime", "dd/mm/yyyy", {"date": True}),
    ("Fecha de compra", "dateTime", "dd/mm/yyyy", {"date": True}),
    ("Hora de compra", "int64", "0", {}),
    ("Horas de aprobación", "double", "0.00", {}),
    ("Días de preparación", "double", "0.00", {}),
    ("Días de transporte", "double", "0.00", {}),
    ("Días de ciclo total", "double", "0.00", {}),
    ("Días vs estimado", "int64", "0", {}),
    ("Estado de entrega", "string", None, {}),
    ("Rango de retraso", "string", None, {"isHidden": True}),
    ("Calidad del registro", "string", None, {}),
]

COLS_CALENDARIO = [
    ("Fecha", "dateTime", "dd/mm/yyyy", {"isKey": True, "date": True}),
    ("Año", "int64", "0", {}),
    ("Mes número", "int64", "0", {"isHidden": True}),
    ("Mes", "string", None, {"sortByColumn": "Mes número"}),
    ("Año-Mes número", "int64", "0", {"isHidden": True}),
    ("Mes y año", "string", None, {"sortByColumn": "Año-Mes número"}),
    ("Trimestre", "string", None, {}),
    ("Número día semana", "int64", "0", {"isHidden": True}),
    ("Día de la semana", "string", None, {"sortByColumn": "Número día semana"}),
]

COLS_ETAPAS = [
    ("Orden", "int64", "0", {"isHidden": True}),
    ("Etapa", "string", None, {"sortByColumn": "Orden"}),
]

# (nombre, formato, descripción, DAX)
MEDIDAS = [
    ("Total Pedidos", "#,0", "Cantidad de pedidos en el contexto actual.",
     "COUNTROWS ( Pedidos )"),
    ("Pedidos Entregados", "#,0", "Pedidos con estado Entregado y fecha de entrega al cliente.",
     'CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Estado de entrega] <> "No entregado" )'),
    ("% Pedidos Entregados", "0.0%", "Proporción de pedidos que llegaron al cliente.",
     "DIVIDE ( [Pedidos Entregados], [Total Pedidos] )"),
    ("Pedidos A Tiempo", "#,0", "Pedidos entregados en o antes de la fecha estimada.",
     'CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Estado de entrega] = "A tiempo" )'),
    ("Pedidos Con Retraso", "#,0", "Pedidos entregados después de la fecha estimada (defectos del proceso).",
     'CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Estado de entrega] = "Con retraso" )'),
    ("% Entregas A Tiempo", "0.0%", "Cumplimiento del acuerdo de nivel de servicio (SLA) de entrega.",
     "DIVIDE ( [Pedidos A Tiempo], [Pedidos Entregados] )"),
    ("Meta % A Tiempo", "0.0%", "Meta del SLA de entrega definida para el proceso.",
     "0.95"),
    ("Brecha vs Meta", "+0.0%;-0.0%;0.0%", "Diferencia entre el cumplimiento real y la meta.",
     "IF ( NOT ISBLANK ( [% Entregas A Tiempo] ), [% Entregas A Tiempo] - [Meta % A Tiempo] )"),
    ("Estado de la Meta", None, "Semáforo del SLA: cumple, en riesgo (a menos de 3 puntos) o no cumple.",
     """
VAR Valor = [% Entregas A Tiempo]
VAR Meta = [Meta % A Tiempo]
RETURN
    SWITCH (
        TRUE (),
        ISBLANK ( Valor ), BLANK (),
        Valor >= Meta, "🟢 Cumple",
        Valor >= Meta - 0.03, "🟡 En riesgo",
        "🔴 No cumple"
    )
"""),
    ("Horas Promedio de Aprobación", "0.0", "Horas promedio entre la compra y la aprobación del pago.",
     "AVERAGE ( Pedidos[Horas de aprobación] )"),
    ("Días Promedio de Aprobación", "0.00", "Tiempo de aprobación expresado en días (para comparar etapas).",
     "DIVIDE ( [Horas Promedio de Aprobación], 24 )"),
    ("Días Promedio de Preparación", "0.00", "Días promedio entre la aprobación y la entrega al transportista.",
     "AVERAGE ( Pedidos[Días de preparación] )"),
    ("Días Promedio de Transporte", "0.00", "Días promedio entre el transportista y la entrega al cliente.",
     "AVERAGE ( Pedidos[Días de transporte] )"),
    ("Días Promedio de Ciclo", "0.0", "Días promedio desde la compra hasta la entrega al cliente.",
     "AVERAGE ( Pedidos[Días de ciclo total] )"),
    ("Días Promedio de Retraso", "0.0", "Días promedio de retraso de los pedidos entregados tarde.",
     'CALCULATE ( AVERAGE ( Pedidos[Días vs estimado] ), Pedidos[Estado de entrega] = "Con retraso" )'),
    ("Días de Holgura Promedio", "0.0", "Días promedio en que la entrega se adelanta a la fecha estimada.",
     """
VAR Diferencia = AVERAGE ( Pedidos[Días vs estimado] )
RETURN
    IF ( NOT ISBLANK ( Diferencia ), -Diferencia )
"""),
    ("DPMO Entregas", "#,0", "Defectos (entregas con retraso) por millón de oportunidades - Lean Six Sigma.",
     "DIVIDE ( [Pedidos Con Retraso], [Pedidos Entregados] ) * 1000000"),
    ("Nivel Sigma", "0.00", "Nivel sigma del proceso de entrega (con desplazamiento de 1,5).",
     """
VAR TasaDefectos = DIVIDE ( [Pedidos Con Retraso], [Pedidos Entregados] )
RETURN
    IF (
        NOT ISBLANK ( TasaDefectos ) && TasaDefectos > 0 && TasaDefectos < 1,
        NORM.S.INV ( 1 - TasaDefectos ) + 1.5
    )
"""),
    ("Pedidos Cancelados o No Disponibles", "#,0", "Pedidos que no se completaron por cancelación o falta de producto.",
     'CALCULATE ( COUNTROWS ( Pedidos ), Estados[Fase del proceso] = "Perdido" )'),
    ("% Pedidos Perdidos", "0.00%", "Proporción de pedidos cancelados o no disponibles.",
     "DIVIDE ( [Pedidos Cancelados o No Disponibles], [Total Pedidos] )"),
    ("Pedidos en Curso", "#,0", "Pedidos que aún no terminan el proceso (creados, aprobados, en proceso, facturados o enviados).",
     'CALCULATE ( COUNTROWS ( Pedidos ), Estados[Fase del proceso] = "En curso" )'),
    ("Registros con Inconsistencias", "#,0", "Pedidos con fechas del proceso que no son coherentes (detectados en el ETL).",
     'CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Calidad del registro] <> "Sin inconsistencias" )'),
    ("% Registros con Inconsistencias", "0.00%", "Proporción de registros con fechas inconsistentes.",
     "DIVIDE ( [Registros con Inconsistencias], [Total Pedidos] )"),
    ("% del Total de Pedidos", "0.00%", "Participación de cada estado del pedido sobre el total.",
     "DIVIDE ( [Total Pedidos], CALCULATE ( [Total Pedidos], ALL ( Estados ) ) )"),
    ("Pedidos que Alcanzan la Etapa", "#,0", "Pedidos que llegaron al menos a cada etapa del proceso (embudo).",
     """
SWITCH (
    SELECTEDVALUE ( Etapas[Orden] ),
    1, [Total Pedidos],
    2, COUNT ( Pedidos[Fecha de aprobación] ),
    3, COUNT ( Pedidos[Fecha entrega a transportista] ),
    4, [Pedidos Entregados]
)
"""),
    ("% Avance del Proceso", "0.0%", "Proporción de pedidos que alcanzan cada etapa respecto al total.",
     "DIVIDE ( [Pedidos que Alcanzan la Etapa], [Total Pedidos] )"),
    ("% Retrasos Mayores a 7 Días", "0.00%", "Pedidos entregados con más de 7 días de retraso sobre el total entregado.",
     "DIVIDE ( CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Días vs estimado] > 7 ), [Pedidos Entregados] )"),
    ("% Registros Consistentes", "0.00%", "Pedidos cuyas fechas cumplen todas las reglas de calidad.",
     'DIVIDE ( CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Calidad del registro] = "Sin inconsistencias" ), [Total Pedidos] )'),
    ("Pedidos Promedio por Día", "#,0.0", "Capacidad operativa: pedidos por día calendario del periodo.",
     "DIVIDE ( [Total Pedidos], COUNTROWS ( Calendario ) )"),
    ("Crecimiento Interanual de Pedidos", "+0.0%;-0.0%;0.0%", "Pedidos del año frente a los mismos meses del año anterior (por defecto 2018 vs 2017, enero-agosto).",
     """
VAR UltimaFecha = CALCULATE ( MAX ( Calendario[Fecha] ), REMOVEFILTERS ( Calendario ) )
VAR AnioActual = IF ( HASONEVALUE ( Calendario[Año] ), VALUES ( Calendario[Año] ), YEAR ( UltimaFecha ) )
VAR MesFin = IF ( AnioActual = YEAR ( UltimaFecha ), MONTH ( UltimaFecha ), 12 )
VAR Actual = CALCULATE ( [Total Pedidos], REMOVEFILTERS ( Calendario ), Calendario[Año] = AnioActual, Calendario[Mes número] <= MesFin )
VAR Anterior = CALCULATE ( [Total Pedidos], REMOVEFILTERS ( Calendario ), Calendario[Año] = AnioActual - 1, Calendario[Mes número] <= MesFin )
RETURN
    DIVIDE ( Actual - Anterior, Anterior )
"""),
    ("Mejora A Tiempo vs Año Anterior", "+0.0%;-0.0%;0.0%", "Puntos porcentuales de mejora del % a tiempo frente a los mismos meses del año anterior.",
     """
VAR UltimaFecha = CALCULATE ( MAX ( Calendario[Fecha] ), REMOVEFILTERS ( Calendario ) )
VAR AnioActual = IF ( HASONEVALUE ( Calendario[Año] ), VALUES ( Calendario[Año] ), YEAR ( UltimaFecha ) )
VAR MesFin = IF ( AnioActual = YEAR ( UltimaFecha ), MONTH ( UltimaFecha ), 12 )
VAR Actual = CALCULATE ( [% Entregas A Tiempo], REMOVEFILTERS ( Calendario ), Calendario[Año] = AnioActual, Calendario[Mes número] <= MesFin )
VAR Anterior = CALCULATE ( [% Entregas A Tiempo], REMOVEFILTERS ( Calendario ), Calendario[Año] = AnioActual - 1, Calendario[Mes número] <= MesFin )
RETURN
    IF ( NOT ISBLANK ( Actual ) && NOT ISBLANK ( Anterior ), Actual - Anterior )
"""),
    ("Valor Indicador", "0.00", "Valor actual del indicador BSC seleccionado.",
     """
SWITCH (
    SELECTEDVALUE ( Indicadores[Código] ),
    "1.1", [Crecimiento Interanual de Pedidos],
    "1.2", [% Pedidos Entregados],
    "1.3", [% Pedidos Perdidos],
    "2.1", [% Entregas A Tiempo],
    "2.2", [Días Promedio de Retraso],
    "2.3", [% Retrasos Mayores a 7 Días],
    "2.4", [Días de Holgura Promedio],
    "3.1", [Días Promedio de Ciclo],
    "3.2", [Horas Promedio de Aprobación],
    "3.3", [Días Promedio de Preparación],
    "3.4", [Días Promedio de Transporte],
    "4.1", [% Registros Consistentes],
    "4.2", [Pedidos Promedio por Día],
    "4.3", [Mejora A Tiempo vs Año Anterior]
)
"""),
    ("Meta Indicador", "0.00", "Meta definida para el indicador BSC seleccionado.",
     "SELECTEDVALUE ( Indicadores[Meta] )"),
    ("Estado Numérico", "0.0", "1 = cumple, 0,5 = en riesgo (dentro de la tolerancia), 0 = no cumple.",
     """
VAR Valor = [Valor Indicador]
VAR Meta = [Meta Indicador]
VAR Tolerancia = SELECTEDVALUE ( Indicadores[Tolerancia] )
VAR MayorEsMejor = SELECTEDVALUE ( Indicadores[Sentido] ) = "Mayor es mejor"
RETURN
    IF (
        NOT ISBLANK ( Valor ) && NOT ISBLANK ( Meta ),
        IF (
            MayorEsMejor,
            IF ( Valor >= Meta, 1, IF ( Valor >= Meta - Tolerancia, 0.5, 0 ) ),
            IF ( Valor <= Meta, 1, IF ( Valor <= Meta + Tolerancia, 0.5, 0 ) )
        )
    )
"""),
    ("Semáforo", None, "Semáforo del indicador BSC.",
     """
VAR E = [Estado Numérico]
RETURN
    IF ( NOT ISBLANK ( E ), IF ( E = 1, "🟢 Cumple", IF ( E = 0.5, "🟡 En riesgo", "🔴 No cumple" ) ) )
"""),
    ("Valor Actual", None, "Valor del indicador con su formato (%, días, horas, pp o número).",
     """
VAR Valor = [Valor Indicador]
VAR Formato = SELECTEDVALUE ( Indicadores[Formato] )
RETURN
    IF (
        NOT ISBLANK ( Valor ),
        SWITCH (
            Formato,
            "%", FORMAT ( Valor, "0.0%" ),
            "días", FORMAT ( Valor, "0.0" ) & " días",
            "horas", FORMAT ( Valor, "0.0" ) & " h",
            "pp", FORMAT ( Valor * 100, "+0.0;-0.0;0.0" ) & " pp",
            FORMAT ( Valor, "#,0" )
        )
    )
"""),
    ("Meta Objetivo", None, "Meta del indicador con su formato y sentido (≥ o ≤).",
     """
VAR Meta = [Meta Indicador]
VAR Formato = SELECTEDVALUE ( Indicadores[Formato] )
VAR Signo = IF ( SELECTEDVALUE ( Indicadores[Sentido] ) = "Mayor es mejor", "≥ ", "≤ " )
RETURN
    IF (
        NOT ISBLANK ( Meta ),
        Signo
            & SWITCH (
                Formato,
                "%", FORMAT ( Meta, "0.0%" ),
                "días", FORMAT ( Meta, "0" ) & " días",
                "horas", FORMAT ( Meta, "0" ) & " h",
                "pp", FORMAT ( Meta * 100, "0.0" ) & " pp",
                FORMAT ( Meta, "#,0" )
            )
    )
"""),
    ("Cumplimiento BSC", "0.0%", "Promedio del estado de los indicadores (verde = 100 %, amarillo = 50 %, rojo = 0 %).",
     "AVERAGEX ( VALUES ( Indicadores[Código] ), [Estado Numérico] )"),
    ("Indicadores en Verde", "0", "Cantidad de indicadores que cumplen la meta.",
     "COUNTROWS ( FILTER ( VALUES ( Indicadores[Código] ), [Estado Numérico] = 1 ) ) + 0"),
    ("Indicadores en Amarillo", "0", "Cantidad de indicadores en riesgo (dentro de la tolerancia).",
     "COUNTROWS ( FILTER ( VALUES ( Indicadores[Código] ), [Estado Numérico] = 0.5 ) ) + 0"),
    ("Indicadores en Rojo", "0", "Cantidad de indicadores que no cumplen la meta.",
     "COUNTROWS ( FILTER ( VALUES ( Indicadores[Código] ), NOT ISBLANK ( [Estado Numérico] ) && [Estado Numérico] = 0 ) ) + 0"),
]


def q(nombre):
    """Nombre TMDL entre comillas simples cuando hace falta."""
    if all(c.isalnum() or c == "_" for c in nombre) and nombre.isascii():
        return nombre
    return "'" + nombre.replace("'", "''") + "'"


def tmdl_columna(nombre, tipo, formato, extra):
    lineas = [f"\tcolumn {q(nombre)}", f"\t\tdataType: {tipo}"]
    if extra.get("isKey"):
        lineas.append("\t\tisKey")
    if extra.get("isHidden"):
        lineas.append("\t\tisHidden")
    if formato:
        lineas.append(f"\t\tformatString: {formato}")
    lineas.append("\t\tsummarizeBy: none")
    lineas.append(f"\t\tsourceColumn: {nombre}")
    if extra.get("sortByColumn"):
        lineas.append(f"\t\tsortByColumn: {q(extra['sortByColumn'])}")
    lineas.append("")
    lineas.append("\t\tannotation SummarizationSetBy = Automatic")
    if extra.get("date"):
        lineas.append("")
        lineas.append("\t\tannotation UnderlyingDateTimeDataType = Date")
    return "\n".join(lineas) + "\n"


def tmdl_tabla(nombre, columnas, m, descripcion, data_category=None):
    out = [f"/// {descripcion}", f"table {q(nombre)}"]
    if data_category:
        out.append(f"\tdataCategory: {data_category}")
    out.append("")
    for c in columnas:
        out.append(tmdl_columna(*c))
    out.append(f"\tpartition {q(nombre)} = m")
    out.append("\t\tmode: import")
    out.append("\t\tsource =")
    out.append(bloque(m, 4))
    out.append("")
    out.append("\tannotation PBI_ResultType = Table")
    out.append("")
    return "\n".join(out)


def tmdl_medidas():
    out = ["/// Tabla contenedora de las medidas DAX del tablero.", "table Medidas", ""]
    for nombre, formato, desc, dax in MEDIDAS:
        out.append(f"\t/// {desc}")
        dax = dax.strip("\n")
        if "\n" in dax:
            out.append(f"\tmeasure {q(nombre)} = ```")
            out.append(bloque(dax, 3))
            out.append("\t\t\t```")
        else:
            out.append(f"\tmeasure {q(nombre)} = {dax}")
        if formato:
            out.append(f"\t\tformatString: {formato}")
        out.append("")
    out += [
        "\tcolumn Value",
        "\t\tisHidden",
        "\t\tformatString: 0",
        "\t\tsummarizeBy: none",
        "\t\tisNameInferred",
        "\t\tsourceColumn: [Value]",
        "",
        "\t\tannotation SummarizationSetBy = Automatic",
        "",
        "\tpartition Medidas = calculated",
        "\t\tmode: import",
        "\t\tsource = ```",
        "\t\t\t\t{0}",
        "\t\t\t\t```",
        "",
    ]
    return "\n".join(out)


def generar_modelo(base):
    sm = base / f"{NOMBRE}.SemanticModel"
    escribir(sm / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2",
        "settings": {"qnaEnabled": False}})
    escribir(sm / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NOMBRE},
        "config": {"version": "2.0", "logicalId": guid(NOMBRE, "modelo")}})
    d = sm / "definition"
    escribir(d / "database.tmdl", "database\n\tcompatibilityLevel: 1567\n\n")
    escribir(d / "model.tmdl", "\n".join([
        "model Model",
        "\tculture: en-US",
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tdiscourageImplicitMeasures",
        "\tsourceQueryCulture: en-US",
        "\tdataAccessOptions",
        "\t\tlegacyRedirects",
        "\t\treturnErrorValuesAsNull",
        "",
        "annotation __PBI_TimeIntelligenceEnabled = 0",
        "",
        'annotation PBI_QueryOrder = ["RutaDatos","Pedidos","Calendario","Estados","Rangos","Indicadores","Etapas"]',
        "",
        "ref table Pedidos",
        "ref table Calendario",
        "ref table Estados",
        "ref table Rangos",
        "ref table Indicadores",
        "ref table Etapas",
        "ref table Medidas",
        "",
    ]))
    escribir(d / "expressions.tmdl", "\n".join([
        "/// Carpeta donde están los archivos CSV originales (debe terminar en \\).",
        f'expression RutaDatos = "{RUTA_DATOS_DEFECTO}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
        "",
        "\tannotation PBI_ResultType = Text",
        "",
    ]))
    escribir(d / "relationships.tmdl", "\n".join([
        f"relationship {guid(NOMBRE, 'rel-pedidos-calendario')}",
        "\tfromColumn: Pedidos.'Fecha de compra'",
        "\ttoColumn: Calendario.Fecha",
        "",
        f"relationship {guid(NOMBRE, 'rel-pedidos-estados')}",
        "\tfromColumn: Pedidos.'Estado del pedido'",
        "\ttoColumn: Estados.'Estado del pedido'",
        "",
        f"relationship {guid(NOMBRE, 'rel-pedidos-rangos')}",
        "\tfromColumn: Pedidos.'Rango de retraso'",
        "\ttoColumn: Rangos.'Rango de retraso'",
        "",
    ]))
    t = d / "tables"
    escribir(t / "Pedidos.tmdl", tmdl_tabla(
        "Pedidos", COLS_PEDIDOS, M_PEDIDOS,
        "Tabla de hechos: un registro por pedido (lado muchos de las relaciones con Calendario, Estados y Rangos)."))
    escribir(t / "Calendario.tmdl", tmdl_tabla(
        "Calendario", COLS_CALENDARIO, M_CALENDARIO,
        "Dimensión de fechas (enero 2017 - agosto 2018).", data_category="Time"))
    escribir(t / "Estados.tmdl", tmdl_tabla(
        "Estados", COLS_ESTADOS, M_ESTADOS,
        "Dimensión de estados del pedido: un estado se relaciona con muchos pedidos (1 a *)."))
    escribir(t / "Rangos.tmdl", tmdl_tabla(
        "Rangos", COLS_RANGOS, M_RANGOS,
        "Dimensión de rangos de retraso: un rango se relaciona con muchos pedidos (1 a *)."))
    escribir(t / "Indicadores.tmdl", tmdl_tabla(
        "Indicadores", COLS_INDICADORES, M_INDICADORES,
        "Indicadores del Balanced Scorecard con su perspectiva, meta y tolerancia (tabla desconectada)."))
    escribir(t / "Etapas.tmdl", tmdl_tabla(
        "Etapas", COLS_ETAPAS, M_ETAPAS,
        "Etapas del proceso de pedido usadas en el embudo (tabla desconectada)."))
    escribir(t / "Medidas.tmdl", tmdl_medidas())


# ---------------------------------------------------------------------------
# Reporte (PBIR)
# ---------------------------------------------------------------------------

SCH = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
V_SCHEMA = f"{SCH}/visualContainer/1.5.0/schema.json"
AZUL = "#1F4E79"
GRIS = "#5D6D7E"


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def texto_lit(s):
    return lit("'" + s.replace("'", "''") + "'")


def color(hex_):
    return {"solid": {"color": texto_lit(hex_)}}


def campo(tipo, entidad, prop):
    return {tipo: {"Expression": {"SourceRef": {"Entity": entidad}}, "Property": prop}}


def col(entidad, prop, activo=True):
    """Proyección de columna; las tablas no usan "active" (igual que las exportaciones de Power BI Desktop)."""
    c = {"field": campo("Column", entidad, prop), "queryRef": f"{entidad}.{prop}", "nativeQueryRef": prop}
    if activo:
        c["active"] = True
    return c


def med(prop, entidad="Medidas"):
    return {"field": campo("Measure", entidad, prop), "queryRef": f"{entidad}.{prop}", "nativeQueryRef": prop}


def orden(tipo, entidad, prop, direccion="Ascending"):
    return {"sort": [{"field": campo(tipo, entidad, prop), "direction": direccion}]}


def contenedor(titulo=None, borde=True):
    o = {}
    if titulo:
        o["title"] = [{"properties": {"show": lit("true"), "text": texto_lit(titulo),
                                       "fontSize": lit("12D"), "fontColor": color(AZUL)}}]
    else:
        o["title"] = [{"properties": {"show": lit("false")}}]
    if borde:
        o["border"] = [{"properties": {"show": lit("true"), "color": color("#D5DBE1"), "radius": lit("8D")}}]
        o["dropShadow"] = [{"properties": {"show": lit("false")}}]
    return o


class Pagina:
    def __init__(self, clave, nombre):
        self.clave = clave
        self.id = gid("pagina", clave)
        self.nombre = nombre
        self.visuales = []

    def agregar(self, clave, x, y, w, h, visual, titulo=None, borde=True):
        z = len(self.visuales) * 1000
        self.visuales.append({
            "$schema": V_SCHEMA,
            "name": gid(self.clave, clave),
            "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z},
            "visual": {**visual, "visualContainerObjects": contenedor(titulo, borde), "drillFilterOtherVisuals": True},
        })


def v_texto(parrafos):
    """parrafos: lista de listas de (texto, tamaño, negrita, color)."""
    ps = []
    for p in parrafos:
        runs = []
        for txt, size, bold, colr in p:
            estilo = {"fontSize": f"{size}pt", "color": colr}
            if bold:
                estilo["fontFamily"] = "Segoe (Bold)"
            runs.append({"value": txt, "textStyle": estilo})
        ps.append({"textRuns": runs})
    return {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": ps}}]}}


def v_slicer(entidad, prop):
    return {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [col(entidad, prop)]}},
                  "sortDefinition": {**orden("Column", entidad, prop), "isDefaultSort": True}},
        "objects": {"data": [{"properties": {"mode": texto_lit("Dropdown")}}],
                    "header": [{"properties": {"show": lit("true"), "fontColor": color(AZUL)}}]},
    }


def v_tarjetas(medidas, columnas=None):
    sel = {"id": "default"}
    return {
        "visualType": "cardVisual",
        "query": {"queryState": {"Data": {"projections": [med(m) for m in medidas]}}},
        "objects": {
            "layout": [{"properties": {"orientation": lit("0D"), "columnCount": lit(f"{columnas or len(medidas)}L"),
                                       "alignment": texto_lit("middle"), "style": texto_lit("Cards")}}],
            "value": [{"properties": {"fontColor": color(AZUL), "fontSize": lit("26D")}, "selector": sel}],
            "label": [{"properties": {"fontColor": color(GRIS)}, "selector": sel}],
            "accentBar": [{"properties": {"show": lit("false")}, "selector": sel}],
            "shadowCustom": [{"properties": {"show": lit("false")}, "selector": sel}],
        },
    }


def v_grafico(tipo, categoria, medidas, sort=None, etiquetas=True, leyenda=None, serie=None):
    qs = {"Category": {"projections": [col(*categoria)]}, "Y": {"projections": [med(m) for m in medidas]}}
    if serie:
        qs["Series"] = {"projections": [col(*serie)]}
    objetos = {"labels": [{"properties": {"show": lit("true" if etiquetas else "false")}}]}
    if tipo != "donutChart" and tipo != "funnel":
        objetos["categoryAxis"] = [{"properties": {"showAxisTitle": lit("false")}}]
        objetos["valueAxis"] = [{"properties": {"showAxisTitle": lit("false")}}]
    if leyenda is not None:
        objetos["legend"] = [{"properties": {"show": lit("true" if leyenda else "false"),
                                             "position": texto_lit("Top")}}]
    v = {"visualType": tipo, "query": {"queryState": qs}, "objects": objetos}
    if sort:
        v["query"]["sortDefinition"] = orden(*sort)
    return v


def v_tabla(columnas, medidas):
    return {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [col(*c, activo=False) for c in columnas] + [med(m) for m in medidas]}},
                  "sortDefinition": {**orden("Column", *columnas[0]), "isDefaultSort": True}},
        "objects": {"columnHeaders": [{"properties": {"backColor": color(AZUL), "fontColor": color("#FFFFFF")}}]},
    }


def encabezado(p, titulo, subtitulo):
    p.agregar("titulo", 16, 10, 620, 70, v_texto([
        [(titulo, 20, True, AZUL)],
        [(subtitulo, 10, False, GRIS)]]), borde=False)
    p.agregar("f_anio", 652, 6, 196, 80, v_slicer("Calendario", "Año"))
    p.agregar("f_mes", 860, 6, 196, 80, v_slicer("Calendario", "Mes"))
    p.agregar("f_estado", 1068, 6, 196, 80, v_slicer("Estados", "Estado del pedido"))


def parrafos(titulo, lineas, size=10):
    return v_texto([[(titulo, 13, True, AZUL)]] + [[(l, size, l.startswith("Hallazgo"), "#1B2631")] for l in lineas])


def construir_paginas():
    MES = ("Calendario", "Mes y año")
    ORDEN_MES = ("Column", "Calendario", "Mes y año")
    paginas = []

    p = Pagina("mapa", "0. Cuadro de mando BSC")
    encabezado(p, "Balanced Scorecard – Proceso de pedidos Olist",
               "14 indicadores en 4 perspectivas con meta y semáforo · enero 2017 – agosto 2018")
    p.agregar("kpis", 16, 92, 1248, 100, v_tarjetas(
        ["Cumplimiento BSC", "Indicadores en Verde", "Indicadores en Amarillo", "Indicadores en Rojo"]))
    p.agregar("scorecard", 16, 204, 836, 508, v_tabla(
        [("Indicadores", "Código"), ("Indicadores", "Perspectiva"), ("Indicadores", "Indicador")],
        ["Meta Objetivo", "Valor Actual", "Semáforo"]), titulo="Cuadro de mando: meta, valor actual y semáforo")
    p.agregar("perspectivas", 864, 204, 400, 248, v_grafico(
        "clusteredBarChart", ("Indicadores", "Perspectiva"), ["Cumplimiento BSC"],
        sort=("Column", "Indicadores", "Perspectiva")), titulo="Cumplimiento por perspectiva")
    p.agregar("mapa", 864, 464, 400, 248, parrafos("Mapa estratégico (causa → efecto)", [
        "1. Financiera: crecer en pedidos completados y reducir pérdidas.",
        "▲ 2. Clientes: cumplir la fecha de entrega prometida.",
        "▲ 3. Procesos internos: aprobar, preparar y transportar más rápido.",
        "▲ 4. Aprendizaje y crecimiento: datos confiables y más capacidad.",
        "Hallazgo: el transporte (9,3 días) es el cuello de botella que baja la puntualidad (93,2 % vs meta 95 %).",
    ]))
    paginas.append(p)

    p = Pagina("financiera", "1. Perspectiva financiera")
    encabezado(p, "Perspectiva financiera",
               "Objetivo: crecer en volumen de ventas (pedidos) y reducir los pedidos perdidos")
    p.agregar("kpis", 16, 92, 1248, 100, v_tarjetas(
        ["Total Pedidos", "Crecimiento Interanual de Pedidos", "% Pedidos Entregados", "% Pedidos Perdidos"]))
    p.agregar("interanual", 16, 204, 1248, 248, v_grafico(
        "lineChart", ("Calendario", "Mes"), ["Total Pedidos"], sort=("Column", "Calendario", "Mes"),
        etiquetas=True, leyenda=True, serie=("Calendario", "Año")),
        titulo="Pedidos por mes: comparación 2017 vs 2018")
    p.agregar("fases", 16, 464, 400, 248, v_grafico(
        "donutChart", ("Estados", "Fase del proceso"), ["Total Pedidos"],
        sort=("Measure", "Medidas", "Total Pedidos", "Descending"), leyenda=True),
        titulo="Pedidos por fase: completados, en curso y perdidos")
    p.agregar("perdidos_mes", 428, 464, 836, 248, v_grafico(
        "lineChart", MES, ["% Pedidos Perdidos"], sort=ORDEN_MES, etiquetas=False),
        titulo="% de pedidos perdidos (cancelados o no disponibles) por mes")
    paginas.append(p)

    p = Pagina("clientes", "2. Perspectiva clientes")
    encabezado(p, "Perspectiva de clientes",
               "Objetivo: cumplir la promesa de entrega · meta 95 % de pedidos a tiempo")
    p.agregar("kpis", 16, 92, 924, 100, v_tarjetas(
        ["% Entregas A Tiempo", "Días Promedio de Retraso", "% Retrasos Mayores a 7 Días", "Días de Holgura Promedio"]))
    p.agregar("meta", 952, 92, 312, 100, v_tarjetas(["Estado de la Meta", "Brecha vs Meta"]))
    p.agregar("tendencia", 16, 204, 1248, 248, v_grafico(
        "lineChart", MES, ["% Entregas A Tiempo", "Meta % A Tiempo"], sort=ORDEN_MES, etiquetas=False, leyenda=True),
        titulo="% de entregas a tiempo por mes frente a la meta")
    p.agregar("rangos", 16, 464, 620, 248, v_grafico(
        "clusteredColumnChart", ("Rangos", "Rango de retraso"), ["Pedidos Entregados"],
        sort=("Column", "Rangos", "Rango de retraso")),
        titulo="Pedidos entregados según días de retraso")
    p.agregar("sla_dia", 648, 464, 616, 248, v_grafico(
        "clusteredBarChart", ("Calendario", "Día de la semana"), ["% Entregas A Tiempo"],
        sort=("Column", "Calendario", "Día de la semana")),
        titulo="% de entregas a tiempo según el día de compra")
    paginas.append(p)

    p = Pagina("procesos", "3. Perspectiva procesos internos")
    encabezado(p, "Perspectiva de procesos internos",
               "Objetivo: reducir el tiempo de ciclo compra → aprobación → transportista → cliente")
    p.agregar("kpis", 16, 92, 1248, 100, v_tarjetas(
        ["Días Promedio de Ciclo", "Horas Promedio de Aprobación", "Días Promedio de Preparación", "Días Promedio de Transporte"]))
    p.agregar("composicion", 16, 204, 1248, 248, v_grafico(
        "columnChart", MES, ["Días Promedio de Aprobación", "Días Promedio de Preparación", "Días Promedio de Transporte"],
        sort=ORDEN_MES, etiquetas=False, leyenda=True),
        titulo="Composición del tiempo de ciclo por mes (días promedio por etapa)")
    p.agregar("embudo", 16, 464, 400, 248, v_grafico(
        "funnel", ("Etapas", "Etapa"), ["Pedidos que Alcanzan la Etapa"], sort=("Column", "Etapas", "Etapa")),
        titulo="Embudo: pedidos que alcanzan cada etapa")
    p.agregar("detalle", 428, 464, 836, 248, v_tabla(
        [MES], ["Total Pedidos", "Horas Promedio de Aprobación", "Días Promedio de Preparación",
                "Días Promedio de Transporte", "Días Promedio de Ciclo"]),
        titulo="Detalle mensual de tiempos")
    paginas.append(p)

    p = Pagina("aprendizaje", "4. Perspectiva aprendizaje y crecimiento")
    encabezado(p, "Perspectiva de aprendizaje y crecimiento",
               "Objetivo: información confiable y mayor capacidad para procesar pedidos")
    p.agregar("kpis", 16, 92, 1248, 100, v_tarjetas(
        ["% Registros Consistentes", "Pedidos Promedio por Día", "Mejora A Tiempo vs Año Anterior", "Registros con Inconsistencias"]))
    p.agregar("calidad", 16, 204, 620, 248, v_grafico(
        "clusteredBarChart", ("Pedidos", "Calidad del registro"), ["Registros con Inconsistencias"],
        sort=("Measure", "Medidas", "Registros con Inconsistencias", "Descending")),
        titulo="Registros con inconsistencias por regla de calidad")
    p.agregar("capacidad", 648, 204, 616, 248, v_grafico(
        "lineChart", MES, ["Pedidos Promedio por Día"], sort=ORDEN_MES, etiquetas=False),
        titulo="Capacidad: pedidos procesados por día, por mes")
    p.agregar("estados", 16, 464, 620, 248, v_tabla(
        [("Estados", "Estado del pedido"), ("Estados", "Fase del proceso")], ["Total Pedidos", "% del Total de Pedidos"]),
        titulo="Pedidos por estado (dimensión Estados, relación 1 a muchos)")
    p.agregar("reglas", 648, 464, 616, 248, parrafos("Revisión de datos en el ETL (Power Query)", [
        "• Cada columna con su tipo: textos (IDs, estado), fecha/hora (5 fechas), enteros y decimales (tiempos).",
        "• Duplicados por ID eliminados; meses incompletos filtrados (2016 y sep-oct 2018); vacíos → null.",
        "• Se marca cada pedido cuyas fechas no siguen el orden compra → aprobación → transportista → cliente.",
        "• Las duraciones negativas no entran en los promedios (quedan en null).",
    ]))
    paginas.append(p)
    return paginas


def generar_reporte(base):
    rp = base / f"{NOMBRE}.Report"
    escribir(rp / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NOMBRE}.SemanticModel"}}})
    escribir(rp / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NOMBRE},
        "config": {"version": "2.0", "logicalId": guid(NOMBRE, "reporte")}})
    temas = rp / "StaticResources" / "SharedResources" / "BaseThemes"
    temas.mkdir(parents=True, exist_ok=True)
    shutil.copy(AQUI / "recursos" / "CY24SU10.json", temas / "CY24SU10.json")
    escribir(rp / "StaticResources" / "RegisteredResources" / "TemaBSC.json", {
        "name": "TemaBSC.json",
        "dataColors": [AZUL, "#2E86C1", "#F39C12", "#27AE60", "#C0392B", "#8E44AD", "#16A085", "#7F8C8D"],
        "background": "#FFFFFF", "foreground": "#1B2631", "tableAccent": AZUL})
    d = rp / "definition"
    escribir(d / "version.json", {"$schema": f"{SCH}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    escribir(d / "report.json", {
        "$schema": f"{SCH}/report/1.2.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": "CY24SU10", "reportVersionAtImport": "5.61", "type": "SharedResources"},
            "customTheme": {"name": "TemaBSC.json", "reportVersionAtImport": "5.61", "type": "RegisteredResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": "TemaBSC.json", "path": "TemaBSC.json", "type": "CustomTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True,
                     "useDefaultAggregateDisplayName": True}})
    paginas = construir_paginas()
    escribir(d / "pages" / "pages.json", {
        "$schema": f"{SCH}/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [p.id for p in paginas], "activePageName": paginas[0].id})
    for p in paginas:
        escribir(d / "pages" / p.id / "page.json", {
            "$schema": f"{SCH}/page/1.3.0/schema.json",
            "name": p.id, "displayName": p.nombre, "displayOption": "FitToPage", "height": 720, "width": 1280})
        for v in p.visuales:
            escribir(d / "pages" / p.id / "visuals" / v["name"] / "visual.json", v)


def main():
    base = Path(sys.argv[1]).resolve()
    if base.exists():
        shutil.rmtree(base)
    base.mkdir(parents=True)
    generar_modelo(base)
    generar_reporte(base)
    escribir(base / f"{NOMBRE}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NOMBRE}.Report"}}],
        "settings": {"enableAutoRecovery": True}})
    escribir(base / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    print("Proyecto generado en", base)


if __name__ == "__main__":
    main()
