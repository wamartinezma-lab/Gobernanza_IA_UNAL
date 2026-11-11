#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Documento Word (.docx): PROPUESTA DE CREACIÓN DE LA MESA TÉCNICA DE INTELIGENCIA
ARTIFICIAL COMO SALA ADSCRITA AL CNED — versión ACTUALIZADA.

Mantiene la estructura del documento original (contexto normativo, marco multinivel,
funciones, ciclo de vida PLASI, cuatro núcleos, Directrices Técnicas DT 1-25,
lineamientos y guías / anexos) e INTEGRA los conceptos nuevos derivados del gemelo
digital 3D y el dashboard interactivo de Gobernanza de la IA:
  Plenaria de Delegados · quórum y votación · matriz RACI · taxonomía de IA ·
  niveles de riesgo (Bajo/Medio/Alto/Inaceptable) · subsidiariedad ·
  participación 129 (108+21) · referentes del Sur Global · hoja de ruta 2025-2027.

Incluye tablas, gráficos NATIVOS editables de Word y diagramas vectoriales DrawingML.
Sin dependencias externas (OOXML con zipfile).
Salida: propuesta-mesa-tecnica-ia-cned.docx
"""
import time
from ooxml_docx import (Doc, para, run, h1, h2, h3, bullet, numitem, pagebreak,
    spacer, nota, caption, table, reset_captions,
    VERDE, DORADO, DORADO2, AZUL, TEAL, CONTROL, SOCIO, CNED, GRIS, ROJO, PUR,
    NEGRO, AMBAR, CELESTE)

d = Doc("Propuesta de creacion de la Mesa Tecnica de IA como sala adscrita al CNED (actualizada)")
B = d.add

# ====================================================================
#   PORTADA
# ====================================================================
B(para([run("UNIVERSIDAD NACIONAL DE COLOMBIA", b=True, sz=22, color=GRIS)], align="center", before=560, after=40))
B(para([run("Dirección Nacional de Estrategia Digital (DNED) · División Universidad Laboratorio (DIVULAB)", sz=20, color=GRIS)], align="center", after=380))
B(para([run("Propuesta de creación de la", b=True, sz=34, color=VERDE)], align="center", after=20))
B(para([run("MESA TÉCNICA DE INTELIGENCIA ARTIFICIAL", b=True, sz=46, color=VERDE)], align="center", after=20))
B(para([run("como sala especializada adscrita al Comité Nacional de Estrategia Digital (CNED)", b=True, sz=26, color=TEAL)], align="center", after=360))
B(para([run("Documento de propuesta · versión actualizada con el modelo operativo de gobernanza", sz=22, color=NEGRO)], align="center", after=40))
B(para([run("(Plenaria de Delegados, quórum y votación, matriz RACI, taxonomía y niveles de riesgo)", i=True, sz=20, color=GRIS)], align="center", after=460))
B(para([run("Elaboró", b=True, sz=22, color=DORADO)], align="center", after=40))
B(para([run("Ing. Wilver Alexander Martínez", b=True, sz=24, color=VERDE)], align="center", after=20))
B(para([run("Líder DNED de la iniciativa institucional de Gobernanza de la IA", sz=20, color=NEGRO)], align="center", after=20))
B(para([run("Ingeniero de Sistemas · Especialista en Gobierno Electrónico · Estudiante de Maestría en Ingeniería Industrial", sz=18, color=GRIS)], align="center", after=360))
B(para([run("Con la coordinación de la DNED–DIVULAB y el acompañamiento de la Mesa Técnica de IA", i=True, sz=18, color=GRIS)], align="center", after=120))
B(para([run("Bogotá D.C. · " + time.strftime("%Y"), sz=20, color=GRIS)], align="center"))
B(pagebreak())

# ====================================================================
#   0. PRESENTACIÓN Y NOTA DE ACTUALIZACIÓN
# ====================================================================
B(h1("Presentación"))
B(para("El presente documento propone la creación de la Mesa Técnica de Inteligencia Artificial (en adelante, la Mesa Técnica) "
  "como sala especializada adscrita al Comité Nacional de Estrategia Digital (CNED) de la Universidad Nacional de Colombia. "
  "La propuesta responde a la necesidad institucional de contar con una instancia estable, colegiada y multiestamentaria que "
  "oriente el uso responsable y la apropiación ética de la inteligencia artificial (IA) en la docencia, la investigación, la "
  "extensión y la gestión, a lo largo de las nueve sedes de la Universidad."))
B(para("Esta versión actualiza la propuesta original incorporando el modelo operativo de gobernanza que se ha venido "
  "co-construyendo y que hoy se encuentra representado en un gemelo digital 3D y en un dashboard interactivo. En particular, "
  "se precisan: el funcionamiento de la Plenaria de Delegados, las reglas de quórum y votación, la matriz RACI de roles de "
  "decisión, la taxonomía de sistemas de IA, la graduación del riesgo por afectación a derechos fundamentales, el principio "
  "de subsidiariedad, la participación alcanzada (129 personas) y la hoja de ruta 2025–2027."))
B(nota("Nota de actualización. Los apartados sombreados como «Novedad del modelo operativo» corresponden a conceptos "
  "incorporados a partir del gemelo digital y el dashboard de Gobernanza de la IA; el resto conserva y precisa el contenido "
  "de la propuesta original."))

# ====================================================================
#   1. CONTEXTO NORMATIVO E INSTITUCIONAL
# ====================================================================
B(h1("1. Contexto normativo e institucional"))
B(para("La creación de la Mesa Técnica se enmarca en la institucionalidad de la estrategia digital de la Universidad y en los "
  "lineamientos que ya orientan el uso de las tecnologías de la información. La gobernanza de la IA no sustituye ni duplica el "
  "gobierno de las TI: lo complementa, ubicándose en el ámbito del gobierno corporativo institucional, con intersecciones "
  "claras con el gobierno de TI y el gobierno de datos."))
B(table(
  [[("Referente", {}), ("Alcance para la Mesa Técnica", {})],
   [("Resolución de Rectoría 614 de 2020", {}), ("Crea y reglamenta el Comité Nacional de Estrategia Digital (CNED) como instancia de gobierno de la estrategia digital, al cual se adscribe la Mesa Técnica como sala especializada.", {})],
   [("Circular 01 de 2024", {}), ("Orienta el uso responsable de herramientas de IA en la comunidad universitaria y motiva una instancia técnica permanente de acompañamiento y evaluación.", {})],
   [("Autonomía universitaria", {}), ("Fundamenta un modelo propio, público, nacional y pluriétnico, que armoniza eficiencia operativa con valores éticos y compromisos socioambientales.", {})],
   [("Habeas Data y protección de datos", {}), ("Marco de obligatorio cumplimiento para cualquier sistema de IA que trate datos personales de la comunidad.", {})]],
  widths=[3200, 6200]))
B(caption("Tabla", "Marco normativo que sustenta la adscripción de la Mesa Técnica al CNED. Cada fila relaciona un referente "
  "(izquierda) con su implicación concreta para la Mesa (derecha)."))
B(h2("Ubicación de la gobernanza de la IA"))
B(para("La gobernanza de la IA supera al gobierno de TI: es un asunto institucional de derechos, misión y sostenibilidad. El "
  "siguiente esquema sitúa a la Mesa Técnica en el gobierno corporativo, en intersección con los gobiernos de TI y de datos."))
B(d.diagram(
  boxes=[
    {"x":320,"y":80,"w":360,"h":150,"text":"Gobierno corporativo\nGobernanza de la IA","fill":VERDE,"fs":850},
    {"x":60,"y":560,"w":340,"h":150,"text":"Gobierno de TI\nInfraestructura y operación","fill":TEAL,"fs":800},
    {"x":600,"y":560,"w":340,"h":150,"text":"Gobierno de datos\nCalidad y Habeas Data","fill":AZUL,"fs":800},
  ],
  arrows=[(430,230,230,560,DORADO2,{"label":"orienta","tail":"stealth"}),
          (570,230,770,560,DORADO2,{"label":"orienta","tail":"stealth"}),
          (400,635,600,635,GRIS,{"tail":"stealth","head":"stealth","dash":True,"label":"se articula"})],
  w_in=6.2, h_in=2.9, title="¿Dónde se ubica la gobernanza de la IA?"))
B(caption("Diagrama", "Ubicación de la gobernanza de la IA. El nodo superior (gobierno corporativo) orienta y se interseca "
  "con el gobierno de TI y el gobierno de datos; la Mesa Técnica opera en esa intersección."))
B(h2("Antecedentes: aspectos principales del benchmarking internacional"))
B(para("El diseño de la Mesa Técnica no parte de cero: se apoya en un ejercicio de referenciación (benchmarking) sobre cómo "
  "universidades de alto nivel gobiernan la inteligencia artificial. Se analizaron tres referentes del Norte Global y, por "
  "su mayor cercanía al contexto de una universidad pública, nacional y multicampus, se ampliaron con referentes del Sur "
  "Global. El propósito no fue adoptar un modelo externo, sino extraer aprendizajes transferibles para construir un modelo "
  "propio. El documento complementario de benchmarking desarrolla este análisis en detalle; a continuación se sintetizan sus "
  "aspectos principales."))
B(h3("Referentes del Norte Global"))
B(table(
  [[("Institución", {}), ("Modelo de gobernanza", {}), ("Carácter de la decisión", {}), ("Aprendizaje clave", {})],
   [("MIT", {"color":"FFFFFF","fill":"A31F34"}), ("Descentralizado por núcleos (SERC; comités ad hoc).", {}), ("Consultivo / orientador.", {}), ("Cercanía al problema mediante núcleos de especialización.", {})],
   [("UCLA", {"color":"FFFFFF","fill":"2774AE"}), ("Matricial-centralizado (UC AI Council; HAIC en salud).", {}), ("Mandatorio en el dominio clínico.", {}), ("Control fuerte y veto donde el riesgo es alto.", {})],
   [("Stanford", {"color":"FFFFFF","fill":"8C1515"}), ("Hub-and-spoke articulado por el instituto HAI.", {}), ("Orientador (estándares y evidencia).", {}), ("Producción de estándares, guías y métricas (AI Index).", {})]],
  widths=[1500, 3400, 2300, 2200]))
B(caption("Tabla", "Síntesis de los referentes del Norte Global. Cada fila relaciona la institución, su modelo de "
  "distribución de la decisión, el carácter de sus decisiones y el aprendizaje que la UNAL incorpora."))
B(h3("Referentes del Sur Global"))
B(table(
  [[("Institución / red", {}), ("Rasgo", {}), ("Aprendizaje para la UNAL", {})],
   [("Tec de Monterrey (México)", {}), ("Multicampus; integridad académica con IA.", {}), ("Coordinar criterios comunes entre sedes distribuidas.", {})],
   [("USP (Brasil)", {}), ("Gran IES pública con enfoque normativo.", {}), ("Regular preservando la misión pública y la escala.", {})],
   [("UNAM (México)", {}), ("Pública nacional; proceso en construcción.", {}), ("Modelo propio, participativo y situado.", {})],
   [("INTEGRITY-AI", {}), ("Red LatAm (U. del Cauca; U. de San Buenaventura).", {}), ("Trabajo en red regional y validación entre pares.", {})]],
  widths=[2600, 3400, 3400]))
B(caption("Tabla", "Síntesis de los referentes del Sur Global, más cercanos al contexto de la UNAL. Cada fila relaciona el "
  "referente, su rasgo distintivo y el aprendizaje transferible al modelo propio."))
B(para([run("Conclusión de los antecedentes. ", b=True, color=DORADO),
  run("La UNAL no adopta ningún modelo tal cual: construye uno propio, público, nacional, pluriétnico y multicampus, que "
      "combina la cercanía al problema (MIT), el control fuerte donde el riesgo es alto (UCLA), la producción de estándares "
      "y evidencia (Stanford) y la coordinación multicampus y la escala pública del Sur Global. Estos aprendizajes "
      "fundamentan la estructura de cuatro núcleos y la operación colegiada que se describen más adelante.")]))
B(pagebreak())

# ====================================================================
#   2. JUSTIFICACIÓN
# ====================================================================
B(h1("2. Justificación"))
B(para("La IA generativa y agéntica ya está presente en aulas, laboratorios y procesos administrativos. Ante esta realidad, la "
  "pregunta determinante no es si usar la IA, sino quién decide qué se puede hacer, con qué datos y bajo qué límites. Los "
  "marcos rígidos o excesivamente centralizados fragmentan la apropiación tecnológica o limitan su impacto pedagógico y "
  "social; por eso la Universidad requiere una instancia técnica permanente, colegiada y cercana a las sedes."))
for t in [
    "Habilitar la apropiación ética y responsable de la IA al servicio de las misiones de la Universidad.",
    "Proteger los derechos fundamentales de la comunidad frente a usos de IA de alto riesgo.",
    "Dar trazabilidad a la decisión: hacer explícito «quién decide qué» mediante un órgano colegiado.",
    "Articular la sostenibilidad (Green AI) y la soberanía del conocimiento como ejes transversales.",
    "Asegurar coherencia multicampus respetando las capacidades y realidades de cada sede."]:
    B(bullet(t))

# ====================================================================
#   3. MARCO MULTINIVEL Y SUBSIDIARIEDAD  (incluye concepto nuevo)
# ====================================================================
B(h1("3. Marco de gobernanza multinivel y subsidiariedad"))
B(para("La gobernanza de la IA en la UNAL se organiza en tres niveles articulados que conectan la decisión política de alto "
  "nivel con la operación cotidiana en las sedes. En la cúspide, el CNED define la política y aprueba la normativa; en el "
  "nivel intermedio, la Mesa Técnica ejecuta, evalúa y audita a través de sus cuatro núcleos; en la base, la comunidad "
  "universitaria —nueve sedes, facultades, grupos y laboratorios vivos— es, a la vez, destinataria y co-creadora del modelo."))
B(d.diagram(
  boxes=[
    {"x":270,"y":40,"w":460,"h":160,"text":"CNED\nComité Nacional de Estrategia Digital\nPolítica y aprobación normativa","fill":CNED,"fs":850},
    {"x":250,"y":410,"w":500,"h":160,"text":"MESA TÉCNICA DE IA\nEjecución · evaluación · auditoría · 4 núcleos","fill":VERDE,"fs":850},
    {"x":230,"y":780,"w":540,"h":160,"text":"COMUNIDAD UNIVERSITARIA\n9 sedes · facultades · grupos · laboratorios vivos","fill":TEAL,"fs":850},
  ],
  arrows=[(410,200,410,410,AZUL,{"tail":"stealth","label":"política"}),
          (410,570,410,780,AZUL,{"tail":"stealth","label":"lineamiento"}),
          (590,410,590,200,DORADO2,{"tail":"stealth","label":"ponencia"}),
          (590,780,590,570,DORADO2,{"tail":"stealth","label":"radica caso"})],
  w_in=6.2, h_in=3.3, title="Arquitectura multinivel de la gobernanza"))
B(caption("Diagrama", "Arquitectura multinivel. Las flechas azules (izquierda) muestran la bajada de política y "
  "lineamientos; las doradas (derecha), el flujo ascendente de co-construcción —radicación de casos y ponencias— desde la "
  "comunidad hacia la Mesa y el CNED."))
B(h2("Principio de subsidiariedad  ·  Novedad del modelo operativo"))
B(para("La decisión se toma en el nivel más cercano al lugar donde se genera el riesgo y el valor; se centraliza solo por "
  "excepción: cuando afecta a toda la Universidad, compromete información sensible, implica riesgos elevados o requiere "
  "estándares institucionales comunes. Así, lo cotidiano y de bajo riesgo se resuelve en la facultad o la sede, y solo escala "
  "cuando el impacto lo justifica."))
B(table(
  [[("Ámbito de decisión", {}), ("Qué decide", {})],
   [("Nacional (CNED / Mesa)", {}), ("Política y normativa; alto riesgo; datos sensibles; estándares comunes.", {})],
   [("Sede", {}), ("Capacidades e infraestructura locales; riesgos territoriales.", {})],
   [("Facultad / Grupo", {}), ("IA en aula, investigación y extensión; bajo riesgo local.", {})],
   [("Compartido", {}), ("Riesgo medio multi-sede; herramientas transversales; Habeas Data.", {})]],
  widths=[3000, 6400]))
B(caption("Tabla", "Reparto de decisiones por subsidiariedad. Cada fila es un ámbito y la columna derecha enumera los asuntos "
  "que le corresponden según su cercanía al riesgo y al valor."))
B(pagebreak())

# ====================================================================
#   4. FUNCIONES DE LA MESA TÉCNICA
# ====================================================================
B(h1("4. Funciones de la Mesa Técnica"))
B(para("La Mesa Técnica cumple funciones estratégicas y tácticas que, en conjunto, le permiten orientar, evaluar y acompañar "
  "el despliegue de la IA en la Universidad. Las funciones estratégicas conectan con la política institucional; las tácticas, "
  "con la operación y el acompañamiento a las sedes."))
B(h3("Funciones estratégicas"))
for t in [
    "Proponer al CNED lineamientos, políticas y normativa para el uso responsable y la apropiación ética de la IA.",
    "Mantener y actualizar el marco de principios (subsidiariedad, derechos, Green AI, soberanía del conocimiento).",
    "Articular el benchmarking internacional y la validación académica externa del modelo.",
    "Priorizar y aprobar la hoja de ruta institucional de IA (plan plurianual 2025–2027)."]:
    B(bullet(t))
B(h3("Funciones tácticas"))
for t in [
    "Evaluar iniciativas y sistemas de IA según su nivel de riesgo y emitir decisión colegiada.",
    "Acompañar a facultades y sedes en la implementación de guías y lineamientos.",
    "Auditar y hacer seguimiento al ciclo de vida de los sistemas (PLASI).",
    "Consolidar y custodiar la documentación del modelo (guías, actas y productos de las mesas)."]:
    B(bullet(t))

# ====================================================================
#   5. CICLO DE VIDA PLASI
# ====================================================================
B(h1("5. Ciclo de vida de los sistemas de IA (PLASI)"))
B(para("La Mesa Técnica evalúa y acompaña los sistemas de IA a lo largo de todo su ciclo de vida —no solo en su aprobación "
  "inicial— mediante el procedimiento PLASI (Planeación, Lineamiento, Aprobación, Seguimiento e Implementación). Cada fase "
  "define responsables, entregables y puntos de control, de modo que la gobernanza acompañe al sistema desde su concepción "
  "hasta su retiro."))
B(table(
  [[("Fase", {"align":"center"}), ("Qué ocurre", {}), ("Resultado / control", {})],
   [("P · Planeación", {}), ("Formulación de la iniciativa; identificación de propósito, datos y usuarios.", {}), ("Ficha del caso y clasificación preliminar de riesgo.", {})],
   [("L · Lineamiento", {}), ("Aplicación de guías y requisitos (ética, datos, seguridad, Green AI).", {}), ("Checklist de cumplimiento por núcleo.", {})],
   [("A · Aprobación", {}), ("Deliberación en la Plenaria de Delegados y decisión colegiada.", {}), ("Acta de decisión (favorable / con condiciones / veto / escalar).", {})],
   [("S · Seguimiento", {}), ("Monitoreo de desempeño, incidentes y afectación a derechos.", {}), ("Reportes periódicos y alertas.", {})],
   [("I · Implementación", {}), ("Despliegue controlado y acompañamiento a la sede o facultad.", {}), ("Puesta en operación con salvaguardas.", {})]],
  widths=[1900, 4100, 3400]))
B(caption("Tabla", "Procedimiento PLASI. Las filas recorren las cinco fases del ciclo de vida; para cada una se indica la "
  "actividad principal y el entregable o punto de control que la cierra."))
B(pagebreak())

# ====================================================================
#   6. LOS CUATRO NÚCLEOS
# ====================================================================
B(h1("6. Los cuatro núcleos funcionales"))
B(para("La Mesa Técnica no es un cuerpo homogéneo: se estructura en cuatro núcleos funcionales que, en conjunto, cubren de "
  "forma multidimensional la gobernanza de la IA. Esta composición busca que ninguna dimensión quede desatendida y que las "
  "tensiones —por ejemplo, entre viabilidad técnica y protección de derechos— se resuelvan de forma deliberada y colegiada."))
B(table(
  [[("Núcleo", {}), ("Composición (ejemplos)", {}), ("Enfoque", {})],
   [("1.1 Estratégico Misional", {"color":"FFFFFF","fill":AZUL}), ("Vicerrectorías; Dir. Extensión y P.I.", {}), ("Coherencia misional, soberanía epistémica.", {})],
   [("1.2 Técnico & Green Computing", {"color":"FFFFFF","fill":TEAL}), ("DNED; ciberseguridad; HPC; TI de sedes.", {}), ("Infraestructura, DT 1-25, eficiencia energética.", {})],
   [("1.3 Control & Jurídico-Fiscal", {"color":"FFFFFF","fill":CONTROL}), ("Planeación; Jurídica; Control Interno.", {}), ("Legalidad, RAEE, Habeas Data, software libre.", {})],
   [("1.4 Socio-Ambiental & Territorial", {"color":"FFFFFF","fill":SOCIO}), ("Investigadores; sedes; estudiantes.", {}), ("Ética, saberes territoriales, salud mental.", {})]],
  widths=[3000, 3400, 3000]))
B(caption("Tabla", "Los cuatro núcleos. El color de la primera columna identifica a cada núcleo; «Composición» da ejemplos de "
  "sus integrantes y «Enfoque» resume su ámbito de responsabilidad."))

# ---- detalle por núcleo ----
B(h2("Detalle de cada núcleo"))

B(h3("1.1 · Núcleo Estratégico Misional"))
B(para("Vela por que el despliegue de la IA sea coherente con las misiones sustantivas de la Universidad —docencia, "
  "investigación y extensión— y con su proyecto educativo. Es el punto de entrada estratégico de los casos: traduce la "
  "política del CNED en criterios de priorización y cuida la soberanía epistémica y el sentido público de la institución."))
B(table(
  [[("Dimensión", {}), ("Detalle", {})],
   [("Composición típica", {}), ("Vicerrectoría General, Académica y de Investigación; Dirección Nacional de Extensión y Propiedad Intelectual.", {})],
   [("Funciones específicas", {}), ("Priorizar iniciativas según aporte misional; velar por la pertinencia pedagógica e investigativa; proteger la soberanía epistémica.", {})],
   [("Decisiones típicas", {}), ("¿La iniciativa aporta a las misiones? ¿Es coherente con el proyecto educativo? ¿Qué prioridad institucional tiene?", {})]],
  widths=[2600, 6800]))

B(h3("1.2 · Núcleo Técnico & Green Computing"))
B(para("Responde por la viabilidad técnica, la seguridad y la sostenibilidad ambiental de los sistemas de IA. Es el núcleo "
  "que materializa las Directrices Técnicas (DT 1–25) y que vela por que la Universidad crezca en capacidades de cómputo sin "
  "comprometer su huella ambiental (Green Computing)."))
B(table(
  [[("Dimensión", {}), ("Detalle", {})],
   [("Composición típica", {}), ("Dirección Nacional de Estrategia Digital (DNED); coordinadores de ciberseguridad; administración del centro de datos / HPC; representantes de TI de las sedes.", {})],
   [("Funciones específicas", {}), ("Evaluar infraestructura y seguridad; aplicar las DT 1–25; medir y reducir la huella energética; asegurar interoperabilidad y MLOps.", {})],
   [("Decisiones típicas", {}), ("¿Hay capacidad de cómputo y seguridad suficientes? ¿Cumple las DT? ¿Cuál es su costo energético y ambiental?", {})]],
  widths=[2600, 6800]))

B(h3("1.3 · Núcleo Control & Jurídico-Fiscal"))
B(para("Cuida la legalidad, la protección de datos y la viabilidad presupuestal y fiscal. Es la primera línea de control: "
  "junto con el núcleo Socio-Ambiental, puede objetar una iniciativa cuando hay riesgo de vulneración de derechos o de "
  "incumplimiento normativo, y activar el escalamiento al CNED."))
B(table(
  [[("Dimensión", {}), ("Detalle", {})],
   [("Composición típica", {}), ("Dirección Nacional de Planeación; Oficina Jurídica Nacional; Control Interno; delegado de gestión ambiental (RAEE).", {})],
   [("Funciones específicas", {}), ("Verificar legalidad y licenciamiento; garantizar Habeas Data; evaluar viabilidad presupuestal; promover software libre y manejo de RAEE.", {})],
   [("Decisiones típicas", {}), ("¿Es legal y presupuestalmente viable? ¿Protege los datos personales? ¿Requiere objeción o escalamiento?", {})]],
  widths=[2600, 6800]))

B(h3("1.4 · Núcleo Socio-Ambiental & Territorial"))
B(para("Aporta la mirada ética, de derechos y de saberes territoriales. Actúa como control ético transversal: su objeción por "
  "riesgo a derechos fundamentales prevalece sobre la mayoría y obliga a vetar la iniciativa o a escalarla al CNED. Integra "
  "la voz de la comunidad, las sedes y los estudiantes."))
B(table(
  [[("Dimensión", {}), ("Detalle", {})],
   [("Composición típica", {}), ("Investigadores en IA y sostenibilidad; delegados de facultades y grupos; Sedes de Presencia Nacional; estudiantes, docentes y administrativos.", {})],
   [("Funciones específicas", {}), ("Evaluar afectación a derechos y a poblaciones; incorporar saberes territoriales e interculturales; atender salud mental y bienestar digital.", {})],
   [("Decisiones típicas", {}), ("¿Afecta derechos fundamentales o a grupos vulnerables? ¿Respeta los saberes territoriales? ¿Debe objetarse o escalarse?", {})]],
  widths=[2600, 6800]))
B(nota("En conjunto, los cuatro núcleos cubren las dimensiones misional, técnica, jurídico-fiscal y socio-ambiental. Dos de "
  "ellos (1.3 y 1.4) tienen capacidad de objeción por riesgo a derechos, lo que convierte la protección de las personas en "
  "un control explícito dentro de la decisión colegiada."))

B(h2("Interacción entre núcleos  ·  Novedad del modelo operativo"))
B(para("Los núcleos no operan de forma aislada: interactúan a través de la plenaria y mediante consultas directas. El Núcleo "
  "Estratégico (1.1) deriva los casos a los núcleos Técnico (1.2) y de Control (1.3), que se consultan entre sí "
  "—seguridad/infraestructura frente a datos/licencias—, mientras el Socio-Ambiental (1.4) actúa como control ético "
  "transversal con capacidad de objeción cuando hay riesgo a derechos."))
B(d.diagram(
  boxes=[
    {"x":370,"y":40,"w":260,"h":140,"text":"1.1 Estratégico\nMisional","fill":AZUL,"fs":820},
    {"x":60,"y":430,"w":260,"h":150,"text":"1.2 Técnico\n& Green Computing","fill":TEAL,"fs":760},
    {"x":680,"y":430,"w":260,"h":150,"text":"1.3 Control\n& Jurídico-Fiscal","fill":CONTROL,"fs":760},
    {"x":370,"y":820,"w":260,"h":150,"text":"1.4 Socio-Ambiental\nControl ético","fill":SOCIO,"fs":760},
  ],
  arrows=[(440,180,210,430,DORADO2,{"tail":"stealth","label":"deriva"}),
          (560,180,790,430,DORADO2,{"tail":"stealth","label":"deriva"}),
          (320,500,680,500,GRIS,{"tail":"stealth","head":"stealth","dash":True,"label":"consulta"}),
          (500,820,500,180,ROJO,{"tail":"triangle","label":"objeción / veto"})],
  w_in=6.2, h_in=3.4, title="Interacción de los cuatro núcleos"))
B(caption("Diagrama", "Interacción de los cuatro núcleos. Las flechas doradas muestran la derivación desde el núcleo "
  "estratégico; la gris punteada, la consulta mutua Técnico↔Control; la flecha roja ascendente, la objeción ética "
  "transversal del núcleo 1.4 con capacidad de veto."))
B(pagebreak())

# ====================================================================
#   7. OPERACIÓN: PLENARIA, QUÓRUM Y VOTACIÓN  (NUEVO)
# ====================================================================
B(h1("7. Operación de la Mesa: Plenaria de Delegados, quórum y votación"))
B(para([run("Novedad del modelo operativo. ", b=True, color=DORADO),
  run("La Mesa Técnica opera como un órgano colegiado y deliberante, no como una ventanilla técnica. Su instancia central es "
      "la Plenaria de Delegados, integrada por un delegado de cada uno de los cuatro núcleos. La plenaria recibe y radica los "
      "casos (iniciativas o sistemas de IA), los traslada a trámite a los núcleos para su evaluación especializada y luego "
      "delibera con las ponencias que cada delegado trae de regreso.")]))
B(d.diagram(
  boxes=[
    {"x":210,"y":40,"w":580,"h":130,"text":"FACULTADES / SEDES","fill":GRIS,"fs":850},
    {"x":320,"y":370,"w":360,"h":160,"text":"PLENARIA DE DELEGADOS\n1 por núcleo · delibera y decide","fill":VERDE,"fs":820},
    {"x":30,"y":760,"w":210,"h":120,"text":"Núcleo 1.1","fill":AZUL,"fs":800},
    {"x":280,"y":760,"w":210,"h":120,"text":"Núcleo 1.2","fill":TEAL,"fs":800},
    {"x":530,"y":760,"w":210,"h":120,"text":"Núcleo 1.3","fill":CONTROL,"fs":800},
    {"x":780,"y":760,"w":210,"h":120,"text":"Núcleo 1.4","fill":SOCIO,"fs":800},
  ],
  arrows=[(500,170,500,370,DORADO2,{"tail":"stealth","label":"radica caso"}),
          (430,530,150,760,GRIS,{"tail":"stealth"}),(470,530,390,760,GRIS,{"tail":"stealth"}),
          (530,530,630,760,GRIS,{"tail":"stealth"}),(570,530,880,760,GRIS,{"tail":"stealth","label":"traslado"}),
          (180,760,440,530,DORADO2,{"tail":"stealth"}),(840,760,560,530,DORADO2,{"tail":"stealth","label":"ponencia"})],
  w_in=6.2, h_in=3.5, title="Flujo operativo de la Mesa Técnica"))
B(caption("Diagrama", "Flujo operativo. Las facultades y sedes radican el caso en la Plenaria (flecha dorada superior); la "
  "plenaria lo traslada a trámite a los cuatro núcleos (flechas grises descendentes) y estos regresan con ponencia (flechas "
  "doradas ascendentes) para la deliberación y la decisión colegiada."))
B(h2("Procedimiento secuencial"))
for i, t in enumerate([
    "Radicación del caso en la Plenaria de Delegados y análisis preliminar.",
    "Traslado a trámite: cada delegado lleva el caso a su núcleo para evaluación especializada.",
    "Consulta entre núcleos (1.2 Técnico ↔ 1.3 Jurídico: seguridad/infraestructura ↔ datos/licencias).",
    "Regreso con ponencia de cada núcleo y deliberación en plenaria.",
    "Decisión colegiada por quórum (≥ 3 de 4 delegados)."], 1):
    B(numitem(i, t))
B(h2("Quórum y regla de decisión"))
B(para("La votación se rige por reglas explícitas: existe un quórum mínimo para sesionar y una regla de prevalencia que "
  "protege los derechos fundamentales. El veto no depende del dominio de aplicación (académico, administrativo, clínico), "
  "sino del nivel de riesgo a derechos."))
B(table(
  [[("Condición de la votación", {}), ("Resultado", {})],
   [("Objeción de 1.3 o 1.4 por riesgo a derechos", {"fill":"FBE9E7"}), ("Prevalece: veto / escalar al CNED", {"fill":"FBE9E7","color":ROJO,"bold":True})],
   [("Unanimidad a favor", {}), ("Favorable", {})],
   [("Mayoría con salvaguardas", {}), ("Favorable con condiciones (nivel sede)", {})],
   [("Sin quórum (< 3 de 4 delegados)", {}), ("Sesión aplazada", {})]],
  widths=[5200, 4200]))
B(caption("Tabla", "Regla de decisión de la plenaria. Cada fila cruza una condición de votación con su resultado. La fila "
  "resaltada expresa la prevalencia: una objeción por riesgo a derechos pesa más que la mayoría y obliga a vetar o escalar."))
B(nota("Ejemplo: un predictor de deserción que trata datos sensibles genera objeción del núcleo jurídico (1.3) y del "
  "socio-ambiental (1.4), por lo que escala al CNED aunque exista mayoría a favor."))
B(pagebreak())

# ====================================================================
#   8. RIESGO Y TAXONOMÍA  (NUEVO)
# ====================================================================
B(h1("8. Gestión de riesgo y taxonomía de sistemas de IA"))
B(para([run("Novedad del modelo operativo. ", b=True, color=DORADO),
  run("No todas las iniciativas de IA representan el mismo riesgo. La gobernanza se gradúa en función de la posible "
      "afectación a los derechos fundamentales de las personas: el control se define por el riesgo y no por el dominio de "
      "aplicación. La gestión combina cuatro niveles de riesgo —cada uno con un tratamiento y una autoridad de aprobación "
      "distintos— y una taxonomía que distingue los tipos de sistemas de IA según su propósito.")]))
B(h2("Niveles de riesgo"))
B(table(
  [[("Nivel", {"align":"center"}), ("Tratamiento", {}), ("Autoridad de aprobación", {})],
   [("Inaceptable", {"color":"FFFFFF","fill":"7A1414"}), ("Prohibido: vulnera derechos fundamentales.", {}), ("No se aprueba.", {})],
   [("Alto", {"color":"FFFFFF","fill":ROJO}), ("Evaluación exhaustiva, salvaguardas y auditoría.", {}), ("CNED (con ponencia de la Mesa).", {})],
   [("Medio", {"color":"FFFFFF","fill":AMBAR}), ("Condiciones y seguimiento.", {}), ("Mesa Técnica / nivel compartido.", {})],
   [("Bajo", {"color":"FFFFFF","fill":SOCIO}), ("Buenas prácticas y registro.", {}), ("Facultad / sede (subsidiariedad).", {})]],
  widths=[1800, 4400, 3200]))
B(caption("Tabla", "Niveles de riesgo. El color indica la severidad (de verde = bajo a vino = inaceptable); el tratamiento y "
  "la autoridad de aprobación aumentan en exigencia conforme crece el riesgo a derechos."))
B(d.chart("bar", ["Bajo","Medio","Alto","Inaceptable"], [1,2,3,4],
          [SOCIO, AMBAR, ROJO, "7A1414"], "Exigencia de control por nivel de riesgo", w_in=6.0, h_in=2.8))
B(caption("Gráfico", "Escalonamiento del control (objeto de gráfico editable de Word). La barra crece con el nivel de riesgo, "
  "ilustrando que a mayor afectación potencial a derechos corresponde mayor exigencia y una autoridad de aprobación superior."))
B(h2("Taxonomía de sistemas de IA"))
for t in [
    "IA de productividad y asistencia (copilotos, generación de texto/código) — uso general con buenas prácticas.",
    "IA analítica y predictiva (predicción, clasificación, scoring) — atención especial cuando decide sobre personas.",
    "IA sobre datos sensibles (salud, bienestar, datos personales) — tratamiento reforzado y Habeas Data.",
    "IA de automatización de decisiones (procesos académicos o administrativos) — requiere supervisión humana.",
    "IA de investigación y laboratorios vivos — gobernanza adaptada al contexto experimental."]:
    B(bullet(t))
B(pagebreak())

# ====================================================================
#   9. MATRIZ RACI  (NUEVO)
# ====================================================================
B(h1("9. Matriz RACI de roles de decisión"))
B(para([run("Novedad del modelo operativo. ", b=True, color=DORADO),
  run("Para hacer trazable «quién decide qué», el modelo adopta una matriz RACI (Responsable, Aprobador, Consultado, "
      "Informado) que asigna, para cada tipo de actividad, el rol de cada instancia. Así se evita que una sola área concentre "
      "la decisión y se clarifican las responsabilidades a lo largo del ciclo PLASI.")]))
B(table(
  [[("Actividad", {}), ("CNED", {"align":"center"}), ("Plenaria", {"align":"center"}), ("Núcleos", {"align":"center"}), ("Sede/Facultad", {"align":"center"})],
   [("Definir política y normativa", {}), ("A", {"align":"center","bold":True,"color":AZUL}), ("R", {"align":"center","bold":True,"color":SOCIO}), ("C", {"align":"center"}), ("I", {"align":"center"})],
   [("Clasificar riesgo del caso", {}), ("I", {"align":"center"}), ("A", {"align":"center","bold":True,"color":AZUL}), ("R", {"align":"center","bold":True,"color":SOCIO}), ("C", {"align":"center"})],
   [("Evaluar por núcleo", {}), ("I", {"align":"center"}), ("C", {"align":"center"}), ("R", {"align":"center","bold":True,"color":SOCIO}), ("I", {"align":"center"})],
   [("Decisión sobre riesgo alto", {}), ("A", {"align":"center","bold":True,"color":AZUL}), ("R", {"align":"center","bold":True,"color":SOCIO}), ("C", {"align":"center"}), ("I", {"align":"center"})],
   [("Decisión sobre riesgo bajo", {}), ("I", {"align":"center"}), ("C", {"align":"center"}), ("C", {"align":"center"}), ("A/R", {"align":"center","bold":True,"color":SOCIO})],
   [("Seguimiento e implementación", {}), ("I", {"align":"center"}), ("C", {"align":"center"}), ("R", {"align":"center","bold":True,"color":SOCIO}), ("R", {"align":"center","bold":True,"color":SOCIO})]],
  widths=[3400, 1500, 1500, 1500, 1500]))
B(caption("Tabla", "Matriz RACI. Las filas son actividades del ciclo de gobernanza; las columnas, las instancias. "
  "R = Responsable (ejecuta), A = Aprobador (decide), C = Consultado, I = Informado. Nótese cómo el riesgo bajo recae en la "
  "sede/facultad (subsidiariedad) y el riesgo alto escala al CNED."))

# ====================================================================
#   10. DIRECTRICES TÉCNICAS DT 1-25
# ====================================================================
B(h1("10. Directrices Técnicas (DT 1–25)"))
B(para("El núcleo Técnico & Green Computing (1.2) mantiene un conjunto de 25 Directrices Técnicas (DT 1–25) que operacionalizan "
  "los requisitos de infraestructura, seguridad, datos, interoperabilidad y sostenibilidad para cualquier sistema de IA de la "
  "Universidad. Se agrupan en cinco dominios; la tabla resume el alcance de cada grupo."))
B(table(
  [[("Grupo", {}), ("Rango", {"align":"center"}), ("Alcance", {})],
   [("Infraestructura y cómputo", {}), ("DT 1–5", {"align":"center"}), ("HPC, nube, aprovisionamiento y capacidad.", {})],
   [("Seguridad y acceso", {}), ("DT 6–10", {"align":"center"}), ("Ciberseguridad, identidad, control de acceso.", {})],
   [("Datos e interoperabilidad", {}), ("DT 11–15", {"align":"center"}), ("Calidad de datos, estándares, integración.", {})],
   [("Green Computing", {}), ("DT 16–20", {"align":"center"}), ("Eficiencia energética, huella de carbono, RAEE.", {})],
   [("Operación y MLOps", {}), ("DT 21–25", {"align":"center"}), ("Despliegue, monitoreo, trazabilidad y software libre.", {})]],
  widths=[3000, 1600, 4800]))
B(caption("Tabla", "Agrupación de las 25 Directrices Técnicas por dominio. Cada grupo reúne cinco directrices; el rango DT "
  "indica su numeración y la columna de alcance resume su objeto."))
B(pagebreak())

# ====================================================================
#   11. LINEAMIENTOS, GUÍAS Y PARTICIPACIÓN
# ====================================================================
B(h1("11. Lineamientos, guías y participación de la comunidad"))
B(para("La Mesa Técnica se nutre de un ecosistema de lineamientos y guías elaborados por 14 mesas de trabajo temáticas, que "
  "aportan conocimiento situado de la comunidad en materia de integridad académica, protección de datos, salud mental y "
  "bienestar digital, enfoque de género, diversidad cultural, entre otros. Este cuerpo de orientaciones es el insumo con el "
  "que la Mesa evalúa y acompaña los casos."))
B(h2("Participación alcanzada  ·  Novedad del modelo operativo"))
B(para("El proceso ha logrado una participación amplia y verificable: 129 personas en total, distribuidas en 108 participantes "
  "de la comunidad universitaria y 21 facilitadores de la DNED, a lo largo de las nueve sedes. Esta base multiestamentaria "
  "sostiene la legitimidad del modelo."))
B(table(
  [[("Indicador", {}), ("Valor", {"align":"center"}), ("Descripción", {})],
   [("Mesas de trabajo", {}), ("14", {"align":"center"}), ("Mesas temáticas que estructuran lineamientos y guías.", {})],
   [("Participantes", {}), ("108", {"align":"center"}), ("Miembros de la comunidad universitaria.", {})],
   [("Facilitadores (DNED)", {}), ("21", {"align":"center"}), ("Equipo que coordina y acompaña las mesas.", {})],
   [("Total personas", {}), ("129", {"align":"center","bold":True}), ("108 participantes + 21 facilitadores.", {})],
   [("Sedes integradas", {}), ("9", {"align":"center"}), ("Alcance multicampus y territorial.", {})]],
  widths=[3000, 1400, 5000]))
B(caption("Tabla", "Participación en la co-construcción. La fila «Total personas» consolida 129 = 108 participantes + 21 "
  "facilitadores de la DNED, sobre 14 mesas de trabajo y 9 sedes."))
B(d.chart("dough", ["Participantes (108)", "Facilitadores DNED (21)"], [108, 21], [TEAL, DORADO2],
          "Participación: 129 personas (108 + 21)", w_in=5.4, h_in=2.9))
B(caption("Gráfico", "Composición de la participación (objeto de gráfico editable de Word). El sector mayor corresponde a los "
  "108 participantes de la comunidad y el menor a los 21 facilitadores de la DNED, para un total de 129 personas."))
B(pagebreak())

# ====================================================================
#   12. HOJA DE RUTA 2025-2027  (NUEVO)
# ====================================================================
B(h1("12. Hoja de ruta 2025–2027"))
B(para([run("Novedad del modelo operativo. ", b=True, color=DORADO),
  run("La puesta en marcha de la Mesa Técnica se inscribe en un plan plurianual que avanza del diagnóstico a la "
      "implementación y evaluación, con validación en sedes y facultades. La siguiente tabla y la línea de tiempo sintetizan "
      "las fases.")]))
B(table(
  [[("Año", {"align":"center"}), ("Fase", {}), ("Hitos", {})],
   [("2025", {"align":"center","bold":True,"color":TEAL}), ("Diagnóstico y co-construcción", {}), ("Benchmarking; marco conceptual; 14 mesas; lineamientos y guías.", {})],
   [("2026", {"align":"center","bold":True,"color":AZUL}), ("Validación y aprobación", {}), ("Validación en sedes y facultades; revisión académica; aprobación CNED.", {})],
   [("2027", {"align":"center","bold":True,"color":PUR}), ("Implementación y evaluación", {}), ("Operación de la Mesa; PLASI en marcha; evaluación y mejora continua.", {})]],
  widths=[1200, 3400, 4800]))
B(caption("Tabla", "Hoja de ruta por año. Cada fila es un año del plan plurianual, con su fase y sus hitos principales."))
B(d.diagram(
  boxes=[
    {"x":20,"y":330,"w":290,"h":220,"text":"2025\nDiagnóstico y\nco-construcción","fill":TEAL,"fs":880},
    {"x":355,"y":330,"w":290,"h":220,"text":"2026\nValidación y\naprobación","fill":AZUL,"fs":880},
    {"x":690,"y":330,"w":290,"h":220,"text":"2027\nImplementación\ny evaluación","fill":PUR,"fs":880},
  ],
  arrows=[(310,440,355,440,DORADO2,{"tail":"stealth"}),(645,440,690,440,DORADO2,{"tail":"stealth"})],
  w_in=6.2, h_in=2.7, title="Hoja de ruta 2025–2027"))
B(caption("Diagrama", "Línea de tiempo 2025–2027. Las flechas marcan la progresión de una fase a la siguiente, del "
  "diagnóstico (2025) a la implementación y evaluación (2027)."))

# ====================================================================
#   13. RECOMENDACIONES Y DECISIÓN SOLICITADA
# ====================================================================
B(h1("13. Recomendaciones y decisión solicitada al CNED"))
B(para("Con base en lo expuesto, se recomienda al Comité Nacional de Estrategia Digital crear la Mesa Técnica de Inteligencia "
  "Artificial como sala especializada adscrita al CNED, adoptando el modelo operativo descrito —Plenaria de Delegados, cuatro "
  "núcleos, reglas de quórum y votación, matriz RACI, graduación del riesgo y ciclo PLASI— y la hoja de ruta 2025–2027."))
for t in [
    "Formalizar la Mesa Técnica y la designación de delegados por núcleo.",
    "Adoptar la graduación del riesgo por afectación a derechos como criterio rector de la decisión.",
    "Habilitar la operación de la Plenaria de Delegados con sus reglas de quórum y prevalencia.",
    "Avalar la hoja de ruta 2025–2027 y los recursos para su implementación.",
    "Mantener la validación académica externa y la participación multiestamentaria como prácticas permanentes."]:
    B(bullet(t))

# ====================================================================
#   14. REFERENCIAS
# ====================================================================
B(h1("14. Referencias y documentación"))
B(para("La documentación completa del proceso —marco conceptual, actas y productos de las mesas de trabajo, guías y "
  "documentos de lineamientos— se encuentra en el repositorio institucional de la DNED–DIVULAB:"))
B(para([run("Repositorio de documentación: ", b=True, sz=22, color=NEGRO),
  d.hyperlink("https://drive.google.com/drive/folders/1zZ3nCgpte0JUMN6NLxNNcKHmAxAID6Jj",
              "https://drive.google.com/drive/folders/1zZ3nCgpte0JUMN6NLxNNcKHmAxAID6Jj")], align="left"))
for r in [
    "Resolución de Rectoría 614 de 2020 (creación y reglamentación del CNED).",
    "Circular 01 de 2024 sobre uso responsable de herramientas de IA.",
    "Documento de marco conceptual, normativo y regulatorio para el uso ético de la IA (estado del arte). Ing. Wilver Alexander Martínez, DNED.",
    "Presentación institucional «Gobernanza IA». DNED–DIVULAB.",
    "Benchmarking internacional (ver documento complementario): MIT; UCLA; Stanford; Tec de Monterrey; USP; UNAM; proyecto INTEGRITY-AI.",
    "Guías y documentos de lineamientos de las 14 mesas de trabajo.",
    "Gemelo digital 3D y dashboard interactivo del modelo de gobernanza de la IA (DNED–DIVULAB)."]:
    B(bullet(r))
B(nota("Documento elaborado con apoyo de herramientas de IA generativa; contenido validado contra los documentos "
  "institucionales de la UNAL. DNED–DIVULAB."))

# ---- guardar ----
doc = d.save("propuesta-mesa-tecnica-ia-cned.docx")
import os
print("DOCX OK:", os.path.getsize("propuesta-mesa-tecnica-ia-cned.docx"), "bytes")
print("Parrafos:", doc.count("<w:p>"), "| Tablas:", doc.count("<w:tbl>"),
      "| Graficos:", len(d._charts), "| Saltos:", doc.count('w:type="page"'))
