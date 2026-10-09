"""Genera la guía del ETL y del tablero (PDF). Uso: python generar_guia_pdf.py <salida.pdf>"""
import sys
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, Preformatted, SimpleDocTemplate, Spacer, Table, TableStyle

ACC = colors.HexColor("#1F4E79")
ss = getSampleStyleSheet()
title = ParagraphStyle("t", parent=ss["Title"], fontSize=16, leading=19, textColor=ACC, spaceAfter=4)
h1 = ParagraphStyle("h1", parent=ss["Heading2"], fontSize=12.5, leading=15, textColor=ACC, spaceBefore=9, spaceAfter=3, keepWithNext=1)
h2 = ParagraphStyle("h2", parent=ss["Heading3"], fontSize=10.5, leading=13, textColor=ACC, spaceBefore=5, spaceAfter=2)
body = ParagraphStyle("b", parent=ss["Normal"], fontSize=9.4, leading=12.2, alignment=TA_JUSTIFY, spaceAfter=3)
bul = ParagraphStyle("bl", parent=body, leftIndent=12, bulletIndent=2, spaceAfter=1.5)
cell = ParagraphStyle("c", parent=body, fontSize=8, leading=9.8, alignment=0, spaceAfter=0)
cellh = ParagraphStyle("ch", parent=cell, fontName="Helvetica-Bold", textColor=colors.white)
code = ParagraphStyle("code", parent=ss["Code"], fontSize=6.6, leading=8, backColor=colors.HexColor("#F4F6F8"),
                      borderPadding=4, leftIndent=2)

S = []
P = lambda t, st=body: S.append(Paragraph(t, st))
B = lambda t: S.append(Paragraph(t, bul, bulletText="•"))


def tabla(filas, anchos):
    data = [[Paragraph(str(c), cellh if i == 0 else cell) for c in f] for i, f in enumerate(filas)]
    t = Table(data, colWidths=[w * cm for w in anchos], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), ACC),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF3F8")]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B0BCC8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2)]))
    S.append(t)
    S.append(Spacer(1, 5))


P("Guía del ETL y del tablero Balanced Scorecard – Pedidos Olist", title)
P("Puntos 2 y 3 de la actividad · Fuente única: <i>olist_orders_dataset.csv</i> · Metodología del tablero: "
  "<b>Balanced Scorecard</b> (cuadro de mando integral).")

P("1. Perfilado de la base de datos (antes del ETL)", h1)
P("La base tiene 99.441 pedidos y 8 columnas. Todas llegan como texto desde el CSV, así que el primer trabajo del ETL es "
  "revisar cada columna y darle el tipo y formato correctos.")
tabla([
    ["Aspecto revisado", "Resultado", "Tratamiento en el ETL"],
    ["Duplicados por ID de pedido", "0 duplicados; 99.441 IDs únicos de 32 caracteres", "Paso de quitar duplicados como control"],
    ["Periodo cubierto", "04/09/2016 a 17/10/2018; 2016 tiene solo 329 pedidos y sep-oct 2018 solo 20",
     "Filtro a meses completos: ene 2017 – ago 2018 (se excluyen 349 pedidos)"],
    ["Fechas vacías", "Aprobación: 160 · Transportista: 1.783 · Entrega al cliente: 2.965 (pedidos no terminados)",
     "Texto vacío → null; los promedios ignoran los nulos"],
    ["Formatos", "Fechas ISO aaaa-mm-dd hh:mm:ss; IDs con comillas mezcladas; estados en inglés",
     "Tipos con cultura en-US, recorte de espacios, estados en español"],
    ["Orden lógico de las fechas", "1.359 transportista antes de la aprobación, 166 antes de la compra, 23 entregas antes "
     "del transportista, 8 'delivered' sin fecha de entrega, 14 sin aprobación",
     "Se marcan en <i>Calidad del registro</i>; duraciones negativas → null"],
], [3.4, 6.8, 6.8])

P("2. Revisión del formato de cada columna", h1)
P("El último paso de Power Query (<i>Tipos verificados por columna</i>) fija el tipo de las 18 columnas de la tabla Pedidos, "
  "y el modelo les asigna un formato de visualización:")
tabla([
    ["Columna original (CSV)", "Columna en el modelo", "Tipo final", "Formato", "Revisión"],
    ["order_id", "ID Pedido", "Texto", "—", "Sin vacíos ni duplicados; 32 caracteres"],
    ["customer_id", "ID Cliente", "Texto", "—", "Sin vacíos; un ID por pedido"],
    ["order_status", "Estado del pedido", "Texto", "—", "8 valores traducidos; llave hacia Estados"],
    ["order_purchase_timestamp", "Fecha y hora de compra", "Fecha/hora", "dd/mm/aaaa hh:mm", "Sin vacíos"],
    ["order_approved_at", "Fecha de aprobación", "Fecha/hora", "dd/mm/aaaa hh:mm", "160 vacíos → null"],
    ["order_delivered_carrier_date", "Fecha entrega a transportista", "Fecha/hora", "dd/mm/aaaa hh:mm", "1.783 vacíos → null"],
    ["order_delivered_customer_date", "Fecha entrega al cliente", "Fecha/hora", "dd/mm/aaaa hh:mm", "2.965 vacíos → null"],
    ["order_estimated_delivery_date", "Fecha estimada de entrega", "Fecha", "dd/mm/aaaa", "Siempre a las 00:00 → solo fecha"],
    ["(calculada)", "Fecha de compra", "Fecha", "dd/mm/aaaa", "Llave hacia Calendario"],
    ["(calculada)", "Hora de compra", "Entero", "0", "0 a 23"],
    ["(calculada)", "Horas de aprobación", "Decimal", "0,00", "Compra → aprobación"],
    ["(calculada)", "Días de preparación / de transporte", "Decimal", "0,00", "Negativos → null"],
    ["(calculada)", "Días de ciclo total", "Decimal", "0,00", "Solo pedidos entregados"],
    ["(calculada)", "Días vs estimado", "Entero", "0", "Negativo = antes de lo prometido"],
    ["(calculada)", "Estado de entrega / Rango de retraso", "Texto", "—", "Rango: llave hacia Rangos"],
    ["(calculada)", "Calidad del registro", "Texto", "—", "6 reglas de consistencia"],
], [3.9, 3.9, 1.8, 2.6, 4.8])

P("3. Operaciones ETL (Power Query)", h1)
P("<b>Extracción:</b> lectura del CSV con <i>Csv.Document</i> (UTF-8, coma, comillas estándar) desde la carpeta del "
  "parámetro <b>RutaDatos</b> y promoción de encabezados.", bul)
P("<b>Transformación:</b> vacíos → null, recorte de espacios, tipos con cultura en-US, duplicados quitados, filtro del "
  "periodo completo, traducción de estados, columnas de tiempo por etapa, clasificación de la entrega (A tiempo / Con "
  "retraso / No entregado y rango de retraso), reglas de calidad, renombrado al español y verificación final de tipos.", bul)
P("<b>Carga:</b> modelo en estrella. La tabla de hechos <b>Pedidos</b> (99.092 filas) se relaciona con tres dimensiones "
  "mediante relaciones <b>uno a muchos</b>:", bul)
tabla([
    ["Dimensión (lado 1)", "Filas", "Columna llave", "Tabla de hechos (lado muchos)", "Significado"],
    ["Calendario", "608 fechas", "Fecha ↔ Fecha de compra", "Pedidos", "Una fecha tiene muchos pedidos"],
    ["Estados", "8 estados", "Estado del pedido", "Pedidos", "Un estado agrupa muchos pedidos (ej. Entregado: 96.211)"],
    ["Rangos", "6 rangos", "Rango de retraso", "Pedidos", "Un rango agrupa muchos pedidos (ej. A tiempo: 89.672)"],
], [3.0, 1.8, 3.8, 3.4, 5.0])
P("Además hay dos tablas desconectadas: <b>Indicadores</b> (los 14 indicadores del BSC con meta y tolerancia) y "
  "<b>Etapas</b> (para el embudo), y la tabla <b>Medidas</b> con 41 medidas DAX.")

P("4. Metodología: Balanced Scorecard", h1)
P("El BSC traduce la estrategia en indicadores equilibrados en cuatro perspectivas unidas por un <b>mapa estratégico</b> "
  "de causa y efecto: con datos confiables y más capacidad (<i>aprendizaje</i>) se agilizan la aprobación, la preparación y "
  "el transporte (<i>procesos</i>); así se cumple la fecha prometida (<i>clientes</i>) y crecen los pedidos completados con "
  "menos pérdidas (<i>financiera</i>). Cada indicador tiene una meta y una tolerancia: <b>verde</b> si cumple la meta, "
  "<b>amarillo</b> si está dentro de la tolerancia y <b>rojo</b> si está fuera.")
tabla([
    ["Perspectiva", "Indicador", "Meta", "Valor esperado", "Semáforo"],
    ["Financiera", "1.1 Crecimiento interanual de pedidos (2018 vs 2017, ene-ago)", "≥ 20 %", "+135,1 %", "Verde"],
    ["Financiera", "1.2 % de pedidos entregados", "≥ 97 %", "97,1 %", "Verde"],
    ["Financiera", "1.3 % de pedidos perdidos (cancelados o no disponibles)", "≤ 1 %", "1,19 %", "Amarillo"],
    ["Clientes", "2.1 % de entregas a tiempo", "≥ 95 %", "93,2 %", "Amarillo"],
    ["Clientes", "2.2 Días promedio de retraso", "≤ 7 días", "10,6 días", "Rojo"],
    ["Clientes", "2.3 % de retrasos mayores a 7 días", "≤ 2 %", "2,97 %", "Amarillo"],
    ["Clientes", "2.4 Días de anticipación frente a la fecha prometida", "≥ 10 días", "11,8 días", "Verde"],
    ["Procesos internos", "3.1 Días promedio de ciclo", "≤ 10 días", "12,5 días", "Rojo"],
    ["Procesos internos", "3.2 Horas promedio de aprobación", "≤ 12 h", "10,3 h", "Verde"],
    ["Procesos internos", "3.3 Días promedio de preparación", "≤ 3 días", "2,8 días", "Verde"],
    ["Procesos internos", "3.4 Días promedio de transporte", "≤ 8 días", "9,3 días", "Amarillo"],
    ["Aprendizaje", "4.1 % de registros consistentes", "≥ 99 %", "98,6 %", "Amarillo"],
    ["Aprendizaje", "4.2 Pedidos procesados por día", "≥ 150", "163", "Verde"],
    ["Aprendizaje", "4.3 Mejora del % a tiempo vs año anterior", "≥ 0 pp", "−4,2 pp", "Rojo"],
], [2.6, 7.6, 1.8, 2.4, 1.8])
P("Resultado esperado del cuadro de mando: <b>6 indicadores en verde, 5 en amarillo y 3 en rojo</b>; cumplimiento global "
  "del 60,7 % (financiera 83 %, procesos 63 %, clientes 50 %, aprendizaje 50 %). Conclusión: el negocio crece, pero el "
  "transporte es el cuello de botella que afecta la puntualidad, y en 2018 la puntualidad empeoró frente a 2017.")

P("5. Páginas del tablero", h1)
tabla([
    ["Página", "Contenido"],
    ["0. Cuadro de mando BSC", "Cumplimiento global, conteo verde/amarillo/rojo, tabla de los 14 indicadores con meta, valor y "
     "semáforo, cumplimiento por perspectiva y mapa estratégico"],
    ["1. Perspectiva financiera", "Pedidos, crecimiento interanual, % entregados y % perdidos; comparación mensual 2017 vs 2018; "
     "pedidos por fase; % perdidos por mes"],
    ["2. Perspectiva clientes", "% a tiempo con semáforo y brecha vs meta, días de retraso, retrasos > 7 días, anticipación; "
     "tendencia vs meta; pedidos por rango de retraso; % a tiempo por día"],
    ["3. Perspectiva procesos internos", "Tiempos de ciclo, aprobación, preparación y transporte; composición del ciclo por mes; "
     "embudo de etapas; detalle mensual"],
    ["4. Perspectiva aprendizaje y crecimiento", "% registros consistentes, capacidad diaria, mejora vs año anterior, "
     "inconsistencias por regla, pedidos por estado y reglas del ETL"],
], [4.6, 12.4])
P("Todas las páginas tienen filtros de Año, Mes y Estado del pedido.")

P("6. Cómo abrir el tablero y guardar el .pbix", h1)
for i, t in enumerate([
    "Descargar <b>Tablero_BSC_Pedidos.zip</b> y descomprimirlo directamente en <b>C:\\</b>: así se crea "
    "<b>C:\\Trabajos\\</b> con las carpetas <i>datos</i> y <i>tablero</i>.",
    "Abrir <b>C:\\Trabajos\\tablero\\Tablero_Pedidos_BSC.pbip</b> con Power BI Desktop (versión actualizada).",
    "Si se descomprimió en otra carpeta: <i>Inicio → Transformar datos → Editar parámetros</i>, escribir en <b>RutaDatos</b> "
    "la carpeta donde está <i>olist_orders_dataset.csv</i> (terminada en \\) y pulsar <i>Aplicar cambios</i>.",
    "Pulsar <b>Inicio → Actualizar</b> y comparar el cuadro de mando con la tabla del punto 4.",
    "Guardar el entregable con <b>Archivo → Guardar como → Archivo de Power BI (.pbix)</b>."], 1):
    S.append(Paragraph(t, bul, bulletText=f"{i}."))
P("Si Power BI Desktop no abre el .pbip, activar en <i>Archivo → Opciones y configuración → Opciones → Características en "
  "versión preliminar</i>: “Opción de guardar proyecto de Power BI (.pbip)”, “Almacenar el modelo semántico con formato "
  "TMDL” y “Almacenar informes con formato de metadatos mejorado (PBIR)”, y reiniciar. En versiones recientes pueden venir "
  "activadas.")

P("Anexo – Medidas DAX del semáforo BSC", h1)
S.append(Preformatted("""Estado Numérico =
    VAR Valor = [Valor Indicador]                      -- SWITCH por Indicadores[Código]
    VAR Meta = [Meta Indicador]
    VAR Tolerancia = SELECTEDVALUE ( Indicadores[Tolerancia] )
    VAR MayorEsMejor = SELECTEDVALUE ( Indicadores[Sentido] ) = "Mayor es mejor"
    RETURN IF ( NOT ISBLANK ( Valor ) && NOT ISBLANK ( Meta ),
        IF ( MayorEsMejor,
             IF ( Valor >= Meta, 1, IF ( Valor >= Meta - Tolerancia, 0.5, 0 ) ),
             IF ( Valor <= Meta, 1, IF ( Valor <= Meta + Tolerancia, 0.5, 0 ) ) ) )
Cumplimiento BSC = AVERAGEX ( VALUES ( Indicadores[Código] ), [Estado Numérico] )
Crecimiento Interanual de Pedidos =
    VAR UltimaFecha = CALCULATE ( MAX ( Calendario[Fecha] ), REMOVEFILTERS ( Calendario ) )
    VAR AnioActual = IF ( HASONEVALUE ( Calendario[Año] ), VALUES ( Calendario[Año] ), YEAR ( UltimaFecha ) )
    VAR MesFin = IF ( AnioActual = YEAR ( UltimaFecha ), MONTH ( UltimaFecha ), 12 )
    VAR Actual = CALCULATE ( [Total Pedidos], REMOVEFILTERS ( Calendario ),
                             Calendario[Año] = AnioActual, Calendario[Mes número] <= MesFin )
    VAR Anterior = CALCULATE ( [Total Pedidos], REMOVEFILTERS ( Calendario ),
                               Calendario[Año] = AnioActual - 1, Calendario[Mes número] <= MesFin )
    RETURN DIVIDE ( Actual - Anterior, Anterior )""", code))

doc = SimpleDocTemplate(sys.argv[1], pagesize=letter, leftMargin=2.1 * cm, rightMargin=2.1 * cm,
                        topMargin=1.7 * cm, bottomMargin=1.7 * cm,
                        title="Guía del ETL y del tablero BSC – Pedidos Olist", author="Grupo de trabajo")
doc.build(S)
