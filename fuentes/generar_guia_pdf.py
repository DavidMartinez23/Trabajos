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
h1 = ParagraphStyle("h1", parent=ss["Heading2"], fontSize=12.5, leading=15, textColor=ACC, spaceBefore=9, spaceAfter=3)
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


P("Guía del ETL y del tablero de control – Proceso de pedidos (Olist)", title)
P("Puntos 2 y 3 de la actividad · Fuente: <i>olist_orders_dataset.csv</i> · Metodología del tablero: <b>BPM</b> "
  "(Gestión de Procesos de Negocio) con indicadores de Lean Six Sigma.")

P("1. Perfilado de la base de datos (antes del ETL)", h1)
P("La base contiene 99.441 pedidos y 8 columnas: identificador del pedido y del cliente, estado del pedido y cinco "
  "fechas que marcan las etapas del proceso (compra, aprobación del pago, entrega al transportista, entrega al cliente y "
  "fecha estimada de entrega). Hallazgos de calidad:")
tabla([
    ["Aspecto revisado", "Resultado", "Tratamiento en el ETL"],
    ["Duplicados por ID de pedido", "0 registros duplicados", "Se deja el paso de quitar duplicados como control"],
    ["Periodo cubierto", "04/09/2016 a 17/10/2018; 2016 tiene solo 329 pedidos (sin datos en nov.) y "
     "sep-oct 2018 solo 20", "Se filtra a meses completos: ene 2017 – ago 2018 (se excluyen 349 pedidos)"],
    ["Fechas vacías", "Aprobación: 160 · Transportista: 1.783 · Entrega al cliente: 2.965 (pedidos no terminados)",
     "Texto vacío → null; las métricas de tiempo ignoran los nulos"],
    ["Formato de fechas y textos", "Fechas ISO (aaaa-mm-dd hh:mm:ss); IDs con comillas mezcladas; estados en inglés",
     "Tipos con cultura en-US, recorte de espacios y traducción de estados al español"],
    ["Orden lógico de las fechas", "1.359 pedidos con transportista antes de la aprobación, 166 antes de la compra, "
     "23 entregas antes del transportista, 8 'delivered' sin fecha de entrega, 14 sin fecha de aprobación",
     "Se marcan en la columna <i>Calidad del registro</i>; las duraciones negativas quedan en null"],
], [3.4, 6.8, 6.8])

P("2. Operaciones ETL (Power Query)", h1)
P("Todo el ETL está en la consulta <b>Pedidos</b> del archivo de Power BI (Inicio → Transformar datos), con un nombre "
  "descriptivo por paso. El parámetro <b>RutaDatos</b> indica la carpeta de los CSV.")
P("<b>Extracción (E):</b> lectura del CSV con <i>Csv.Document</i> (UTF-8, separador coma, comillas estándar) y promoción "
  "de encabezados.", bul)
P("<b>Transformación (T):</b>", bul)
for t in [
    "Reemplazo de textos vacíos por null en las fechas opcionales; recorte de espacios en los identificadores.",
    "Conversión de tipos (fecha/hora) con cultura en-US para que funcione con cualquier configuración regional de Windows.",
    "Eliminación de duplicados por ID de pedido y filtro del periodo completo (ene 2017 – ago 2018).",
    "Traducción de los 8 estados: delivered → Entregado, shipped → Enviado, canceled → Cancelado, unavailable → "
    "No disponible, invoiced → Facturado, processing → En proceso, created → Creado, approved → Aprobado.",
    "Columnas calculadas del proceso: <i>Fecha de compra</i> (llave con el calendario), <i>Hora de compra</i>, "
    "<i>Horas de aprobación</i>, <i>Días de preparación</i> (aprobación → transportista), <i>Días de transporte</i> "
    "(transportista → cliente), <i>Días de ciclo total</i>, <i>Días vs estimado</i> (negativo = antes de lo prometido).",
    "Clasificación: <i>Estado de entrega</i> (A tiempo / Con retraso / No entregado) y <i>Rango de retraso</i> "
    "(1-3, 4-7, 8-14, más de 14 días) con su columna de orden.",
    "Validación: <i>Calidad del registro</i> aplica 6 reglas de consistencia de fechas y conserva la primera que se incumple.",
    "Renombrado de todas las columnas al español."]:
    S.append(Paragraph(t, ParagraphStyle("b2", parent=bul, leftIndent=24, bulletIndent=14), bulletText="–"))
P("<b>Carga (L):</b> modelo estrella con la tabla de hechos <b>Pedidos</b> (99.092 filas) relacionada (muchos a uno) "
  "con la dimensión <b>Calendario</b> (generada en Power Query, marcada como tabla de fechas). Se agregan la tabla "
  "desconectada <b>Etapas</b> (para el embudo) y la tabla <b>Medidas</b> con 26 medidas DAX.", bul)

P("3. Metodología elegida: BPM", h1)
P("La base describe un <b>proceso</b>: cada pedido pasa por las etapas compra → aprobación del pago → entrega al "
  "transportista → entrega al cliente. Por eso se eligió BPM, cuya fase de <b>monitoreo</b> mide el desempeño del "
  "proceso (volumen, tiempos por etapa, cuellos de botella y cumplimiento del acuerdo de servicio, SLA) para "
  "<b>optimizarlo</b>. Se complementa con Lean Six Sigma: una entrega con retraso es un <i>defecto</i>, y con él se "
  "calculan los DPMO y el nivel sigma del proceso.")
tabla([
    ["Página del tablero", "Pregunta BPM", "Contenido"],
    ["1. Resumen del proceso", "¿Cuántos pedidos entran y hasta dónde llegan?",
     "Tarjetas KPI, embudo de etapas, pedidos por mes, pedidos por estado y lectura de hallazgos"],
    ["2. Tiempos por etapa", "¿Dónde está el cuello de botella?",
     "Composición del tiempo de ciclo por mes (aprobación, preparación, transporte), horas de aprobación por día y detalle mensual"],
    ["3. Cumplimiento de entregas", "¿Se cumple la fecha prometida (SLA)?",
     "% a tiempo vs meta del 95 %, semáforo, brecha, días de retraso, nivel sigma, rangos de retraso"],
    ["4. Excepciones y calidad", "¿Qué pedidos no completan el proceso y qué datos fallan?",
     "Cancelaciones, pedidos en curso, inconsistencias por regla, % de cancelación por mes y reglas del ETL"],
], [3.6, 4.6, 8.8])
P("Todas las páginas tienen filtros de Año, Mes y Estado del pedido.")

P("4. Indicadores (valores esperados para todo el periodo)", h1)
P("Estos valores se calcularon con una réplica del ETL en Python (<i>fuentes/etl_pedidos_python.py</i>) y sirven para "
  "comprobar que el tablero carga bien los datos:")
tabla([
    ["Indicador (medida DAX)", "Definición", "Valor esperado"],
    ["Total Pedidos", "Pedidos del periodo después del ETL", "99.092"],
    ["% Pedidos Entregados", "Entregados con fecha de entrega / total", "97,1 %"],
    ["% Entregas A Tiempo", "Entregados en o antes de la fecha estimada / entregados (meta 95 %)", "93,2 % (No cumple)"],
    ["Pedidos Con Retraso", "Entregados después de la fecha estimada", "6.531"],
    ["Días Promedio de Retraso", "Promedio de días tarde, solo pedidos con retraso", "10,6 días"],
    ["Horas Promedio de Aprobación", "Compra → aprobación del pago", "10,3 horas"],
    ["Días Promedio de Preparación", "Aprobación → transportista", "2,83 días"],
    ["Días Promedio de Transporte", "Transportista → cliente (cuello de botella)", "9,34 días"],
    ["Días Promedio de Ciclo", "Compra → entrega al cliente", "12,5 días"],
    ["Nivel Sigma", "NORM.S.INV(1 − tasa de defectos) + 1,5", "≈ 2,99"],
    ["% Cancelación", "(Cancelados + No disponibles) / total", "1,19 %"],
    ["Registros con Inconsistencias", "Pedidos que incumplen alguna regla de fechas", "1.401"],
], [4.6, 8.4, 4.0])

P("5. Cómo abrir el tablero y obtener el archivo .pbix", h1)
for i, t in enumerate([
    "Instalar o actualizar <b>Power BI Desktop</b> (Microsoft Store o microsoft.com/power-bi).",
    "Descargar el repositorio (botón <i>Code → Download ZIP</i>) y descomprimirlo de modo que los datos queden en "
    "<b>C:\\Trabajos\\datos\\originales\\</b>. Si se usa otra carpeta, ver el paso 4.",
    "Abrir el archivo <b>tablero\\Tablero_Pedidos_BPM.pbip</b> con doble clic.",
    "Si la carpeta es distinta: <i>Inicio → Transformar datos → Editar parámetros</i>, escribir la ruta de la carpeta de "
    "los CSV en <b>RutaDatos</b> (terminada en \\) y pulsar <i>Aplicar cambios</i>.",
    "Pulsar <b>Inicio → Actualizar</b> para cargar los datos. Comparar las tarjetas con la tabla del punto 4.",
    "Guardar el entregable con <b>Archivo → Guardar como</b>, tipo <b>Archivo de Power BI (.pbix)</b>."], 1):
    S.append(Paragraph(t, bul, bulletText=f"{i}."))
P("Si Power BI Desktop no abre el .pbip, se debe activar en <i>Archivo → Opciones y configuración → Opciones → "
  "Características en versión preliminar</i>: “Opción de guardar proyecto de Power BI (.pbip)”, “Almacenar el modelo "
  "semántico con formato TMDL” y “Almacenar informes con formato de metadatos mejorado (PBIR)”. Después se reinicia "
  "Power BI Desktop. En las versiones recientes estas opciones pueden venir activadas por defecto.")

P("Anexo – Medidas DAX principales", h1)
S.append(Preformatted("""% Entregas A Tiempo = DIVIDE ( [Pedidos A Tiempo], [Pedidos Entregados] )
Pedidos A Tiempo    = CALCULATE ( COUNTROWS ( Pedidos ), Pedidos[Estado de entrega] = "A tiempo" )
Días Promedio de Ciclo = AVERAGE ( Pedidos[Días de ciclo total] )
Estado de la Meta =
    VAR Valor = [% Entregas A Tiempo]
    VAR Meta = [Meta % A Tiempo]          -- 0,95
    RETURN SWITCH ( TRUE (), ISBLANK ( Valor ), BLANK (),
                    Valor >= Meta, "[verde] Cumple", Valor >= Meta - 0.03, "[amarillo] En riesgo", "[rojo] No cumple" )
Nivel Sigma =
    VAR TasaDefectos = DIVIDE ( [Pedidos Con Retraso], [Pedidos Entregados] )
    RETURN IF ( TasaDefectos > 0 && TasaDefectos < 1, NORM.S.INV ( 1 - TasaDefectos ) + 1.5 )
Pedidos que Alcanzan la Etapa =
    SWITCH ( SELECTEDVALUE ( Etapas[Orden] ), 1, [Total Pedidos],
             2, COUNT ( Pedidos[Fecha de aprobación] ), 3, COUNT ( Pedidos[Fecha entrega a transportista] ),
             4, [Pedidos Entregados] )""", code))
P("El código M completo de cada paso puede verse en Power Query o en "
  
  "<i>tablero/Tablero_Pedidos_BPM.SemanticModel/definition/tables/Pedidos.tmdl</i>.", ParagraphStyle("izq", parent=body, alignment=0))

doc = SimpleDocTemplate(sys.argv[1], pagesize=letter, leftMargin=2.1 * cm, rightMargin=2.1 * cm,
                        topMargin=1.7 * cm, bottomMargin=1.7 * cm,
                        title="Guía del ETL y del tablero – Proceso de pedidos", author="Grupo de trabajo")
doc.build(S)
