import sys
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER

out = sys.argv[1]
doc = SimpleDocTemplate(out, pagesize=letter, leftMargin=2.2*cm, rightMargin=2.2*cm,
                        topMargin=1.8*cm, bottomMargin=1.8*cm,
                        title="Metodologías para el desarrollo de tableros de BI",
                        author="Grupo de trabajo")
ss = getSampleStyleSheet()
ACC = colors.HexColor("#1F4E79")
title = ParagraphStyle("t", parent=ss["Title"], fontSize=15, leading=18, textColor=ACC, spaceAfter=2)
sub = ParagraphStyle("s", parent=ss["Normal"], fontSize=9, alignment=TA_CENTER, textColor=colors.HexColor("#555555"), spaceAfter=8)
h = ParagraphStyle("h", parent=ss["Heading2"], fontSize=11.5, leading=14, textColor=ACC, spaceBefore=7, spaceAfter=2)
body = ParagraphStyle("b", parent=ss["Normal"], fontName="Helvetica", fontSize=9.6, leading=12.4, alignment=TA_JUSTIFY, spaceAfter=3)
bul = ParagraphStyle("bl", parent=body, leftIndent=11, bulletIndent=2, spaceAfter=1.5)
cell = ParagraphStyle("c", parent=body, fontSize=8.2, leading=10, alignment=0, spaceAfter=0)
cellh = ParagraphStyle("ch", parent=cell, fontName="Helvetica-Bold", textColor=colors.white)
ref = ParagraphStyle("r", parent=body, fontSize=8, leading=10, leftIndent=10, firstLineIndent=-10, spaceAfter=1, alignment=0)

S = []
P = lambda t, st=body: S.append(Paragraph(t, st))
B = lambda t: S.append(Paragraph(t, bul, bulletText="•"))

P("Metodologías para el desarrollo de tableros de control en Inteligencia de Negocios", title)
P("Integrantes: ______________________ · ______________________ · ______________________", sub)

P("Un tablero de control (<i>dashboard</i>) no es solo un conjunto de gráficos: es una herramienta que traduce datos en "
  "decisiones. Para que cumpla ese propósito necesita un marco que defina <b>qué medir</b> (estrategia y objetivos), "
  "<b>cómo medirlo</b> (indicadores y procesos), <b>para quién</b> (usuarios) y <b>cómo comunicarlo</b> (narrativa visual). "
  "A continuación se resumen seis metodologías que responden a esas preguntas.")

# 1 BSC
P("1. Balanced Scorecard (Cuadro de Mando Integral)", h)
P("Propuesto por Robert Kaplan y David Norton (Harvard Business Review, 1992), el Balanced Scorecard traduce la misión y la "
  "estrategia de la organización en un conjunto equilibrado de objetivos e indicadores. Su idea central es que los indicadores "
  "financieros son <i>rezagados</i> (muestran el pasado) y deben complementarse con indicadores <i>adelantados</i> que anticipen el "
  "desempeño futuro. Organiza la medición en cuatro perspectivas:")
B("<b>Financiera:</b> ¿cómo nos ven los accionistas? (rentabilidad, crecimiento de ingresos, margen, ROI).")
B("<b>Clientes:</b> ¿cómo nos ven los clientes? (satisfacción, retención, participación de mercado).")
B("<b>Procesos internos:</b> ¿en qué procesos debemos sobresalir? (eficiencia, calidad, tiempos de ciclo).")
B("<b>Aprendizaje y crecimiento:</b> ¿cómo seguimos mejorando? (capacitación, tecnología, clima laboral).")
P("Las perspectivas se conectan mediante un <b>mapa estratégico</b> de relaciones causa-efecto: personas capacitadas mejoran los "
  "procesos, que satisfacen al cliente, que genera resultados financieros. Cada objetivo tiene indicador, meta e iniciativa. "
  "<b>En BI</b> se materializa como un tablero con una página o sección por perspectiva y semáforos meta vs. real. "
  "<b>Ventaja:</b> alinea la operación con la estrategia. <b>Limitación:</b> su implementación es exigente y puede volverse rígida si no se revisa.")

# 2 OKR y KPI
P("2. OKR y KPI", h)
P("Los <b>KPI</b> (<i>Key Performance Indicators</i>) son métricas cuantificables que miden el desempeño de un proceso u "
  "objetivo frente a una meta (p. ej. ventas mensuales, tasa de conversión, tiempo medio de entrega). Un buen KPI cumple el "
  "criterio <b>SMART</b>: específico, medible, alcanzable, relevante y acotado en el tiempo. Monitorean la <i>salud</i> continua del negocio.")
P("Los <b>OKR</b> (<i>Objectives and Key Results</i>) fueron creados por Andy Grove en Intel y popularizados por John Doerr en "
  "Google. Combinan un <b>Objetivo</b> cualitativo e inspirador (“Ser la marca preferida en la región”) con 3 a 5 <b>Resultados "
  "Clave</b> medibles (“Aumentar el NPS de 30 a 50”). Se fijan en ciclos cortos (normalmente trimestrales), son públicos y "
  "transparentes, y suelen ser ambiciosos: lograr el 70 % se considera un buen resultado.")
P("<b>Relación:</b> los KPI miden lo que ya funciona (“mantener el rumbo”); los OKR impulsan el cambio (“cambiar el rumbo”), y un "
  "resultado clave a menudo es un KPI con una meta retadora. <b>En BI</b> se usan tarjetas KPI, medidores (<i>gauges</i>) de "
  "avance, tendencias frente a metas y barras de progreso por resultado clave.")

# 3 LSS
P("3. Lean Six Sigma", h)
P("Integra dos enfoques: <b>Lean</b> (derivado del Sistema de Producción Toyota), que busca eliminar desperdicios y todo lo que no "
  "agrega valor al cliente, y <b>Six Sigma</b> (Motorola, década de 1980, impulsado por General Electric), que reduce la "
  "variabilidad de los procesos con estadística, hasta 3,4 defectos por millón de oportunidades. Lean ataca los ocho desperdicios "
  "(sobreproducción, esperas, transporte, sobreprocesamiento, inventario, movimientos, defectos y talento no aprovechado). "
  "Su ciclo de mejora es <b>DMAIC</b>:")
B("<b>Definir</b> el problema, el alcance y lo que es crítico para la calidad (CTQ) según la voz del cliente.")
B("<b>Medir</b> el desempeño actual del proceso y recolectar datos confiables (línea base).")
B("<b>Analizar</b> las causas raíz (Pareto, diagrama de Ishikawa, 5 porqués, regresión).")
B("<b>Mejorar</b> implementando y validando soluciones. <b>Controlar</b> para sostener las mejoras (cartas de control, estandarización).")
P("<b>En BI</b> el tablero es la herramienta de las fases Medir y Controlar: muestra defectos, tiempos de ciclo, rendimiento, "
  "diagramas de Pareto y cartas de control con límites. <b>Ventaja:</b> decisiones basadas en datos y mejora medible. "
  "<b>Limitación:</b> requiere datos de proceso de calidad y personal capacitado (cinturones verdes/negros).")

# 4 UCD
P("4. UCD – Diseño Centrado en el Usuario", h)
P("El <i>User-Centered Design</i> (Don Norman, años 80; norma ISO 9241-210) es un proceso iterativo en el que las necesidades, "
  "tareas y limitaciones de los usuarios finales guían cada decisión de diseño. Sus fases son: (1) <b>entender el contexto de uso</b> "
  "mediante entrevistas, observación y <i>personas</i>; (2) <b>especificar los requisitos</b>: qué preguntas debe responder el "
  "tablero y qué decisiones apoya; (3) <b>diseñar soluciones</b> con bocetos, <i>wireframes</i> y prototipos; y (4) <b>evaluar</b> con "
  "usuarios reales (pruebas de usabilidad), repitiendo el ciclo hasta cumplir los requisitos.")
P("<b>En BI</b> evita tableros saturados que nadie usa: define la jerarquía visual (lo más importante arriba a la izquierda), "
  "el nivel de detalle por tipo de usuario (directivo vs. operativo), filtros intuitivos y accesibilidad (contraste, color). "
  "<b>Ventaja:</b> alta adopción y utilidad real. <b>Limitación:</b> exige tiempo y acceso a los usuarios.")

# 5 BPM
P("5. BPM – Gestión de Procesos de Negocio", h)
P("El <i>Business Process Management</i> es una disciplina que gestiona la organización a partir de sus procesos de punta a punta "
  "(no por departamentos aislados) para hacerlos más eficientes, flexibles y alineados con el cliente. Su ciclo de vida comprende: "
  "<b>diseñar</b> (identificar el proceso actual <i>AS-IS</i> y el deseado <i>TO-BE</i>), <b>modelar</b> (comúnmente con la notación "
  "estándar <b>BPMN</b>), <b>ejecutar</b> (implementar, a menudo con software BPMS que automatiza flujos), <b>monitorear</b> (medir "
  "el desempeño con indicadores de proceso) y <b>optimizar</b> (rediseñar con base en la evidencia).")
P("<b>En BI</b> la fase de monitoreo se apoya en tableros de <i>Business Activity Monitoring</i>: tiempos por etapa, cuellos de "
  "botella, volumen de casos, cumplimiento de acuerdos de nivel de servicio (SLA) y costo por proceso. <b>Ventaja:</b> visión "
  "transversal y mejora continua. <b>Limitación:</b> requiere madurez organizacional y procesos documentados.")

# 6 Storytelling
P("6. Storytelling con datos", h)
P("El <i>data storytelling</i> (popularizado por Cole Nussbaumer Knaflic en <i>Storytelling with Data</i>, 2015) combina tres "
  "elementos: <b>datos</b> (la evidencia), <b>visualización</b> (la forma de mostrarla) y <b>narrativa</b> (el hilo que le da "
  "sentido y lleva a la acción). Sus principios son: entender el contexto y la audiencia (¿quién?, ¿qué debe saber o hacer?); elegir "
  "el gráfico adecuado para cada mensaje; eliminar el desorden visual (bordes, colores y decimales innecesarios); enfocar la atención "
  "con color, tamaño y posición; pensar como diseñador; y contar una historia con estructura de <b>inicio</b> (contexto), <b>nudo</b> "
  "(hallazgo o conflicto) y <b>desenlace</b> (recomendación o llamado a la acción).")
P("<b>En BI</b> se traduce en páginas que siguen un orden lógico (resumen → análisis → detalle), títulos que afirman el hallazgo "
  "(“Las ventas cayeron 12 % en la región norte”) en lugar de describir el gráfico, y anotaciones que guían la lectura. "
  "<b>Ventaja:</b> la información se recuerda y se convierte en acción. <b>Limitación:</b> una narrativa sesgada puede manipular la interpretación.")

# Comparativo
P("Cuadro comparativo", h)
rows = [["Metodología", "Pregunta que responde", "Foco", "Aporte al tablero"],
        ["Balanced Scorecard", "¿Se cumple la estrategia?", "Estratégico", "4 perspectivas, mapa causa-efecto, meta vs. real"],
        ["OKR y KPI", "¿Avanzamos hacia las metas?", "Estratégico / táctico", "Tarjetas, medidores, % de avance, tendencias"],
        ["Lean Six Sigma", "¿El proceso es eficiente y estable?", "Operativo", "Defectos, Pareto, cartas de control"],
        ["UCD", "¿El tablero sirve a quien lo usa?", "Diseño / usabilidad", "Requisitos, prototipos, pruebas con usuarios"],
        ["BPM", "¿Cómo fluyen los procesos?", "Procesos", "Tiempos por etapa, cuellos de botella, SLA"],
        ["Storytelling", "¿Se entiende y motiva a actuar?", "Comunicación", "Narrativa, jerarquía visual, títulos con hallazgos"]]
data = [[Paragraph(c, cellh if i == 0 else cell) for c in r] for i, r in enumerate(rows)]
t = Table(data, colWidths=[3.1*cm, 4.3*cm, 3.0*cm, 6.6*cm], repeatRows=1)
t.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), ACC),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF3F8")]),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B0BCC8")),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]))
S.append(KeepTogether([t]))
S.append(Spacer(1, 4))
P("<b>Conclusión:</b> las metodologías son complementarias. BSC y OKR/KPI definen <i>qué</i> medir; Lean Six Sigma y BPM aportan "
  "la mirada de <i>procesos</i>; UCD y Storytelling garantizan que el tablero sea <i>usable y comprensible</i>. Un buen tablero de "
  "BI suele combinar un marco de medición con principios de diseño y narrativa.")

P("Referencias", h)
for r in [
    "Kaplan, R. S. y Norton, D. P. (1992). The Balanced Scorecard: Measures that drive performance. <i>Harvard Business Review</i>, 70(1), 71-79.",
    "Doerr, J. (2018). <i>Measure What Matters: OKRs</i>. Portfolio/Penguin.",
    "George, M. L. (2002). <i>Lean Six Sigma: Combining Six Sigma Quality with Lean Speed</i>. McGraw-Hill.",
    "ISO 9241-210:2019. <i>Ergonomics of human-system interaction – Human-centred design for interactive systems</i>.",
    "Dumas, M., La Rosa, M., Mendling, J. y Reijers, H. A. (2018). <i>Fundamentals of Business Process Management</i> (2.ª ed.). Springer.",
    "Knaflic, C. N. (2015). <i>Storytelling with Data</i>. Wiley.",
    "Few, S. (2006). <i>Information Dashboard Design</i>. O'Reilly."]:
    P(r, ref)

doc.build(S)
