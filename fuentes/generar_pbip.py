"""Genera el proyecto de Power BI (PBIP: modelo TMDL + reporte PBIR) del tablero de pedidos.

Uso: python generar_pbip.py <carpeta_destino>
"""
import hashlib
import json
import shutil
import sys
import uuid
from pathlib import Path

NOMBRE = "Tablero_Pedidos_BPM"
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
    OrdenRangos = [#"A tiempo" = 1, #"1 a 3 días" = 2, #"4 a 7 días" = 3, #"8 a 14 días" = 4, #"Más de 14 días" = 5, #"No entregado" = 6],
    #"Agregado orden del rango" = Table.AddColumn(#"Agregado rango de retraso", "delay_range_order", each Record.Field(OrdenRangos, [delay_range]), Int64.Type),
    // Validación de consistencia de las fechas del proceso (se marca la primera regla que se incumple)
    #"Agregada calidad del registro" = Table.AddColumn(#"Agregado orden del rango", "record_quality", each
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
        {"delay_range", "Rango de retraso"}, {"delay_range_order", "Orden rango de retraso"}, {"record_quality", "Calidad del registro"}})
in
    #"Columnas renombradas"
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
    // Etapas del proceso de pedido (modelo BPM) para el embudo
    Etapas = #table(type table [Orden = Int64.Type, Etapa = text], {
        {1, "1. Compra registrada"},
        {2, "2. Pago aprobado"},
        {3, "3. Entregado al transportista"},
        {4, "4. Entregado al cliente"}})
in
    Etapas
'''

# (nombre, tipo, formato, extra) - extra: dict con isHidden / sortByColumn / isKey / summarizeBy
COLS_PEDIDOS = [
    ("ID Pedido", "string", None, {}),
    ("ID Cliente", "string", None, {"isHidden": True}),
    ("Estado del pedido", "string", None, {}),
    ("Fecha y hora de compra", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha de aprobación", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha entrega a transportista", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha entrega al cliente", "dateTime", "dd/mm/yyyy hh:nn", {}),
    ("Fecha estimada de entrega", "dateTime", "dd/mm/yyyy", {}),
    ("Fecha de compra", "dateTime", "dd/mm/yyyy", {"date": True}),
    ("Hora de compra", "int64", "0", {}),
    ("Horas de aprobación", "double", "0.00", {}),
    ("Días de preparación", "double", "0.00", {}),
    ("Días de transporte", "double", "0.00", {}),
    ("Días de ciclo total", "double", "0.00", {}),
    ("Días vs estimado", "int64", "0", {}),
    ("Estado de entrega", "string", None, {}),
    ("Rango de retraso", "string", None, {"sortByColumn": "Orden rango de retraso"}),
    ("Orden rango de retraso", "int64", "0", {"isHidden": True}),
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
     'CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Estado del pedido] IN { "Cancelado", "No disponible" } )'),
    ("% Cancelación", "0.00%", "Proporción de pedidos cancelados o no disponibles.",
     "DIVIDE ( [Pedidos Cancelados o No Disponibles], [Total Pedidos] )"),
    ("Pedidos en Curso", "#,0", "Pedidos que aún no terminan el proceso (creados, aprobados, en proceso, facturados o enviados).",
     'CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Estado del pedido] IN { "Creado", "Aprobado", "En proceso", "Facturado", "Enviado" } )'),
    ("Registros con Inconsistencias", "#,0", "Pedidos con fechas del proceso que no son coherentes (detectados en el ETL).",
     'CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Calidad del registro] <> "Sin inconsistencias" )'),
    ("% Registros con Inconsistencias", "0.00%", "Proporción de registros con fechas inconsistentes.",
     "DIVIDE ( [Registros con Inconsistencias], [Total Pedidos] )"),
    ("% del Total de Pedidos", "0.00%", "Participación de cada estado del pedido sobre el total.",
     "DIVIDE ( [Total Pedidos], CALCULATE ( [Total Pedidos], ALL ( Pedidos[Estado del pedido] ) ) )"),
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
        'annotation PBI_QueryOrder = ["RutaDatos","Pedidos","Calendario","Etapas"]',
        "",
        "ref table Pedidos",
        "ref table Calendario",
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
    ]))
    t = d / "tables"
    escribir(t / "Pedidos.tmdl", tmdl_tabla(
        "Pedidos", COLS_PEDIDOS, M_PEDIDOS,
        "Tabla de hechos: un registro por pedido, con tiempos por etapa del proceso y validaciones de calidad."))
    escribir(t / "Calendario.tmdl", tmdl_tabla(
        "Calendario", COLS_CALENDARIO, M_CALENDARIO,
        "Dimensión de fechas (enero 2017 - agosto 2018).", data_category="Time"))
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


def col(entidad, prop):
    return {"field": campo("Column", entidad, prop), "queryRef": f"{entidad}.{prop}",
            "nativeQueryRef": prop, "active": True}


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


def v_grafico(tipo, categoria, medidas, sort=None, etiquetas=True, leyenda=None):
    qs = {"Category": {"projections": [col(*categoria)]}, "Y": {"projections": [med(m) for m in medidas]}}
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
        "query": {"queryState": {"Values": {"projections": [col(*c) for c in columnas] + [med(m) for m in medidas]}},
                  "sortDefinition": orden("Column", *columnas[0])},
        "objects": {"columnHeaders": [{"properties": {"backColor": color(AZUL), "fontColor": color("#FFFFFF")}}]},
    }


def encabezado(p, titulo, subtitulo):
    p.agregar("titulo", 16, 10, 620, 70, v_texto([
        [(titulo, 20, True, AZUL)],
        [(subtitulo, 10, False, GRIS)]]), borde=False)
    p.agregar("f_anio", 652, 6, 196, 80, v_slicer("Calendario", "Año"))
    p.agregar("f_mes", 860, 6, 196, 80, v_slicer("Calendario", "Mes"))
    p.agregar("f_estado", 1068, 6, 196, 80, v_slicer("Pedidos", "Estado del pedido"))


def construir_paginas():
    MES = ("Calendario", "Mes y año")
    ORDEN_MES = ("Column", "Calendario", "Mes y año")
    paginas = []

    p = Pagina("resumen", "1. Resumen del proceso")
    encabezado(p, "Proceso de pedidos: visión general",
               "Metodología BPM - fase de monitoreo del proceso compra → aprobación → transportista → cliente")
    p.agregar("kpis", 16, 92, 1248, 100, v_tarjetas(
        ["Total Pedidos", "% Pedidos Entregados", "% Entregas A Tiempo", "Días Promedio de Ciclo"]))
    p.agregar("embudo", 16, 204, 400, 248, v_grafico(
        "funnel", ("Etapas", "Etapa"), ["Pedidos que Alcanzan la Etapa"], sort=("Column", "Etapas", "Etapa")),
        titulo="Embudo del proceso: pedidos que alcanzan cada etapa")
    p.agregar("pedidos_mes", 428, 204, 836, 248, v_grafico(
        "clusteredColumnChart", MES, ["Total Pedidos"], sort=ORDEN_MES),
        titulo="Volumen de pedidos por mes")
    p.agregar("estados", 16, 464, 400, 248, v_grafico(
        "donutChart", ("Pedidos", "Estado del pedido"), ["Total Pedidos"], sort=("Measure", "Medidas", "Total Pedidos", "Descending"),
        etiquetas=True, leyenda=True), titulo="Pedidos por estado")
    p.agregar("lectura", 428, 464, 836, 248, v_texto([
        [("¿Qué nos dice el proceso? (enero 2017 - agosto 2018)", 13, True, AZUL)],
        [("• 99.092 pedidos analizados; el 97,1 % llegó al cliente y solo el 1,2 % se canceló o no tuvo producto disponible.", 11, False, "#1B2631")],
        [("• El ciclo completo dura en promedio 12,5 días: la etapa de transporte (9,3 días) es el cuello de botella, muy por encima de la preparación (2,8 días) y la aprobación del pago (10,3 horas).", 11, False, "#1B2631")],
        [("• El 93,2 % de las entregas se cumple a tiempo, por debajo de la meta del 95 %. Las caídas se concentran en noviembre 2017 (Black Friday) y febrero-marzo 2018.", 11, False, "#1B2631")],
        [("• Recomendación: priorizar la mejora del transporte en temporadas de alta demanda.", 11, True, "#1B2631")],
    ]), titulo=None)
    paginas.append(p)

    p = Pagina("tiempos", "2. Tiempos por etapa")
    encabezado(p, "Tiempos por etapa del proceso",
               "¿Dónde se va el tiempo? Identificación del cuello de botella del proceso")
    p.agregar("kpis", 16, 92, 1248, 100, v_tarjetas(
        ["Horas Promedio de Aprobación", "Días Promedio de Preparación", "Días Promedio de Transporte", "Días Promedio de Ciclo"]))
    p.agregar("composicion", 16, 204, 1248, 248, v_grafico(
        "columnChart", MES, ["Días Promedio de Aprobación", "Días Promedio de Preparación", "Días Promedio de Transporte"],
        sort=ORDEN_MES, etiquetas=False, leyenda=True),
        titulo="Composición del tiempo de ciclo por mes (días promedio por etapa)")
    p.agregar("dia_semana", 16, 464, 400, 248, v_grafico(
        "clusteredBarChart", ("Calendario", "Día de la semana"), ["Horas Promedio de Aprobación"],
        sort=("Column", "Calendario", "Día de la semana")),
        titulo="Horas de aprobación según el día de compra")
    p.agregar("detalle", 428, 464, 836, 248, v_tabla(
        [MES], ["Total Pedidos", "Horas Promedio de Aprobación", "Días Promedio de Preparación",
                "Días Promedio de Transporte", "Días Promedio de Ciclo"]),
        titulo="Detalle mensual de tiempos")
    paginas.append(p)

    p = Pagina("sla", "3. Cumplimiento de entregas")
    encabezado(p, "Cumplimiento de la fecha de entrega (SLA)",
               "Meta: 95 % de pedidos entregados en la fecha prometida · retraso = defecto (Lean Six Sigma)")
    p.agregar("kpis", 16, 92, 924, 100, v_tarjetas(
        ["% Entregas A Tiempo", "Pedidos Con Retraso", "Días Promedio de Retraso", "Nivel Sigma"]))
    p.agregar("meta", 952, 92, 312, 100, v_tarjetas(["Estado de la Meta", "Brecha vs Meta"]))
    p.agregar("tendencia", 16, 204, 1248, 248, v_grafico(
        "lineChart", MES, ["% Entregas A Tiempo", "Meta % A Tiempo"], sort=ORDEN_MES, etiquetas=False, leyenda=True),
        titulo="% de entregas a tiempo por mes frente a la meta")
    p.agregar("rangos", 16, 464, 620, 248, v_grafico(
        "clusteredColumnChart", ("Pedidos", "Rango de retraso"), ["Pedidos Entregados"],
        sort=("Column", "Pedidos", "Rango de retraso")),
        titulo="Pedidos entregados según días de retraso")
    p.agregar("sla_dia", 648, 464, 616, 248, v_grafico(
        "clusteredBarChart", ("Calendario", "Día de la semana"), ["% Entregas A Tiempo"],
        sort=("Column", "Calendario", "Día de la semana")),
        titulo="% de entregas a tiempo según el día de compra")
    paginas.append(p)

    p = Pagina("excepciones", "4. Excepciones y calidad")
    encabezado(p, "Excepciones del proceso y calidad de los datos",
               "Pedidos que no completan el proceso y registros con fechas inconsistentes detectados en el ETL")
    p.agregar("kpis", 16, 92, 1248, 100, v_tarjetas(
        ["Pedidos Cancelados o No Disponibles", "% Cancelación", "Pedidos en Curso", "Registros con Inconsistencias"]))
    p.agregar("calidad", 16, 204, 620, 248, v_grafico(
        "clusteredBarChart", ("Pedidos", "Calidad del registro"), ["Registros con Inconsistencias"],
        sort=("Measure", "Medidas", "Registros con Inconsistencias", "Descending")),
        titulo="Registros con inconsistencias por tipo de regla")
    p.agregar("cancel_mes", 648, 204, 616, 248, v_grafico(
        "lineChart", MES, ["% Cancelación"], sort=ORDEN_MES, etiquetas=False),
        titulo="% de cancelación por mes")
    p.agregar("estados", 16, 464, 620, 248, v_tabla(
        [("Pedidos", "Estado del pedido")], ["Total Pedidos", "% del Total de Pedidos"]),
        titulo="Pedidos por estado")
    p.agregar("reglas", 648, 464, 616, 248, v_texto([
        [("Reglas de calidad aplicadas en el ETL (Power Query)", 13, True, AZUL)],
        [("• Se eliminaron duplicados por ID de pedido y se filtraron los meses incompletos (2016 y sep-oct 2018).", 10, False, "#1B2631")],
        [("• Fechas vacías → null; fechas leídas con cultura en-US; estados traducidos al español.", 10, False, "#1B2631")],
        [("• Se marca cada pedido cuyas fechas no siguen el orden compra → aprobación → transportista → cliente.", 10, False, "#1B2631")],
        [("• Las duraciones negativas no se usan en los promedios (quedan en null) para no distorsionar los tiempos.", 10, False, "#1B2631")],
    ]), titulo=None)
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
    escribir(rp / "StaticResources" / "RegisteredResources" / "TemaBPM.json", {
        "name": "TemaBPM.json",
        "dataColors": [AZUL, "#2E86C1", "#F39C12", "#27AE60", "#C0392B", "#8E44AD", "#16A085", "#7F8C8D"],
        "background": "#FFFFFF", "foreground": "#1B2631", "tableAccent": AZUL})
    d = rp / "definition"
    escribir(d / "version.json", {"$schema": f"{SCH}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    escribir(d / "report.json", {
        "$schema": f"{SCH}/report/1.2.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": "CY24SU10", "reportVersionAtImport": "5.61", "type": "SharedResources"},
            "customTheme": {"name": "TemaBPM.json", "reportVersionAtImport": "5.61", "type": "RegisteredResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": "TemaBPM.json", "path": "TemaBPM.json", "type": "CustomTheme"}]}],
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
