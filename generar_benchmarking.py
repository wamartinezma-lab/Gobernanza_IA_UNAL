#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Documento Word (.docx): BENCHMARKING INTERNACIONAL DE GOBERNANZA DE IA — versión
ACTUALIZADA Y AMPLIADA.

Mantiene la estructura del documento original (marco teórico, los tres comités —MIT,
UCLA, Stanford—, matriz comparativa, grafo de relaciones, diagnóstico UNAL,
conclusiones y referencias) y lo AMPLÍA con:
  - referentes del Sur Global: Tec de Monterrey, USP, UNAM y el proyecto INTEGRITY-AI
    (con pares colombianos U. del Cauca y U. de San Buenaventura);
  - los conceptos del modelo UNAL resultante (multinivel, cuatro núcleos, Plenaria de
    Delegados, quórum/votación, RACI, niveles de riesgo, subsidiariedad, hoja de ruta).

Incluye tablas, gráficos NATIVOS editables de Word y un grafo de relaciones DrawingML.
Sin dependencias externas (OOXML con zipfile).
Salida: benchmarking-gobernanza-ia-internacional.docx
"""
import time
from ooxml_docx import (Doc, para, run, h1, h2, h3, bullet, numitem, pagebreak,
    spacer, nota, caption, table,
    VERDE, DORADO, DORADO2, AZUL, TEAL, CONTROL, SOCIO, CNED, GRIS, ROJO, PUR,
    NEGRO, AMBAR, CELESTE)

MIT_C   = "A31F34"   # rojo MIT
UCLA_C  = "2774AE"   # azul UCLA
STAN_C  = "8C1515"   # cardinal Stanford

d = Doc("Benchmarking internacional de gobernanza de IA (MIT, UCLA, Stanford) y referentes del Sur Global")
B = d.add

# ====================================================================
#   PORTADA
# ====================================================================
B(para([run("UNIVERSIDAD NACIONAL DE COLOMBIA", b=True, sz=22, color=GRIS)], align="center", before=560, after=40))
B(para([run("Dirección Nacional de Estrategia Digital (DNED) · División Universidad Laboratorio (DIVULAB)", sz=20, color=GRIS)], align="center", after=380))
B(para([run("Benchmarking internacional de", b=True, sz=32, color=VERDE)], align="center", after=20))
B(para([run("GOBERNANZA DE LA INTELIGENCIA ARTIFICIAL", b=True, sz=42, color=VERDE)], align="center", after=20))
B(para([run("MIT · UCLA · Stanford, ampliado con referentes del Sur Global", b=True, sz=26, color=TEAL)], align="center", after=40))
B(para([run("y lectura hacia el modelo propio de la UNAL", b=True, sz=24, color=TEAL)], align="center", after=360))
B(para([run("Documento de referenciación y diagnóstico · versión actualizada", sz=22, color=NEGRO)], align="center", after=40))
B(para([run("(Sur Global: Tec de Monterrey · USP · UNAM · INTEGRITY-AI, y conceptos del modelo UNAL resultante)", i=True, sz=20, color=GRIS)], align="center", after=460))
B(para([run("Elaboró", b=True, sz=22, color=DORADO)], align="center", after=40))
B(para([run("Ing. Wilver Alexander Martínez", b=True, sz=24, color=VERDE)], align="center", after=20))
B(para([run("Líder DNED de la iniciativa institucional de Gobernanza de la IA", sz=20, color=NEGRO)], align="center", after=20))
B(para([run("Ingeniero de Sistemas · Especialista en Gobierno Electrónico · Estudiante de Maestría en Ingeniería Industrial", sz=18, color=GRIS)], align="center", after=360))
B(para([run("Con la coordinación de la DNED–DIVULAB", i=True, sz=18, color=GRIS)], align="center", after=120))
B(para([run("Bogotá D.C. · " + time.strftime("%Y"), sz=20, color=GRIS)], align="center"))
B(pagebreak())

# ====================================================================
#   1. INTRODUCCIÓN Y MARCO TEÓRICO
# ====================================================================
B(h1("1. Introducción y marco teórico"))
B(para("Antes de diseñar un modelo propio de gobernanza de la inteligencia artificial (IA), la Universidad Nacional de "
  "Colombia realizó un ejercicio de referenciación (benchmarking) sobre la forma en que universidades de alto nivel gobiernan "
  "la IA. El objetivo no fue copiar un modelo externo, sino comprender las lógicas de decisión disponibles y extraer "
  "aprendizajes transferibles a un contexto público, nacional, pluriétnico y multicampus."))
B(para("Esta versión actualiza y amplía el benchmarking original —centrado en tres referentes del Norte Global (MIT, UCLA y "
  "Stanford)— incorporando deliberadamente referentes del Sur Global, más cercanos a la realidad de la UNAL, y conectando los "
  "hallazgos con el modelo de gobernanza que la Universidad construyó a partir de ellos."))
B(h2("Marco teórico: gobernanza frente a gestión de la IA"))
B(para("La gobernanza de la IA supera al gobierno de las TI. Se ubica en el gobierno corporativo institucional —donde se "
  "definen política, derechos, misión y sostenibilidad— con intersecciones con el gobierno de TI (infraestructura y "
  "operación) y el gobierno de datos (calidad, protección y Habeas Data). El benchmarking se lee bajo cinco criterios de "
  "análisis que luego estructuran la matriz comparativa."))
B(table(
  [[("Criterio", {}), ("Pregunta que responde", {})],
   [("Modelo de gobernanza", {}), ("¿Cómo se distribuye la decisión (centralizada, descentralizada, matricial, hub-and-spoke)?", {})],
   [("Órganos y rasgos", {}), ("¿Qué instancias existen y qué las distingue?", {})],
   [("Carácter de la decisión", {}), ("¿Es consultiva, orientadora o mandatoria?", {})],
   [("Enfoque de riesgo", {}), ("¿Cómo se gradúa el control según el impacto?", {})],
   [("Dimensión ambiental", {}), ("¿Hay consideraciones de sostenibilidad o Green AI?", {})]],
  widths=[2800, 6600]))
B(caption("Tabla", "Criterios de análisis del benchmarking. Cada fila es un criterio y la pregunta que busca responder al "
  "comparar los referentes."))
B(pagebreak())

# ====================================================================
#   2. REFERENTES DEL NORTE GLOBAL
# ====================================================================
B(h1("2. Referentes del Norte Global"))
B(h2("2.1 MIT — modelo descentralizado por núcleos"))
B(para("El Massachusetts Institute of Technology articula su gobernanza de la IA de forma descentralizada, apoyada en el "
  "Schwarzman College of Computing (SERC — Social and Ethical Responsibilities of Computing), comités ad hoc y programas "
  "semilla. La decisión se distribuye entre núcleos de especialización, con un carácter predominantemente consultivo: orienta "
  "y habilita, más que imponer. Su fortaleza es la cercanía al problema; su reto, la coherencia institucional global."))
B(h2("2.2 UCLA — modelo matricial-centralizado"))
B(para("La Universidad de California (UCLA) opera un esquema matricial-centralizado, articulado por el UC AI Council a nivel de "
  "sistema y por instancias como el HAIC (Health AI Committee) en el dominio clínico. En salud, el carácter de la decisión es "
  "mandatorio: ningún sistema clínico se despliega sin aval. Su fortaleza es el control fuerte donde el riesgo es alto; su "
  "reto, extender esa gobernanza a dominios no clínicos sin perder agilidad."))
B(h2("2.3 Stanford — modelo hub-and-spoke"))
B(para("Stanford organiza su gobernanza como un hub-and-spoke articulado por el instituto HAI (Human-Centered AI), con nodos "
  "como AIMES y productos de referencia como el AI Index. El carácter es orientador: fija dirección, estándares y evidencia, "
  "y deja la ejecución a las unidades. Su fortaleza es la producción de conocimiento y estándares; su reto, traducir la "
  "orientación en obligaciones verificables."))
B(h2("Síntesis de los tres referentes del Norte"))
B(table(
  [[("Institución", {}), ("Modelo", {}), ("Órganos / rasgos", {}), ("Carácter", {})],
   [("MIT", {"color":"FFFFFF","fill":MIT_C}), ("Descentralizado", {}), ("SERC; comités ad hoc; programa semilla", {}), ("Consultivo", {})],
   [("UCLA", {"color":"FFFFFF","fill":UCLA_C}), ("Matricial-centralizado", {}), ("UC AI Council; HAIC (salud)", {}), ("Mandatorio (clínico)", {})],
   [("Stanford", {"color":"FFFFFF","fill":STAN_C}), ("Hub-and-spoke", {}), ("HAI; AIMES; AI Index", {}), ("Orientador", {})]],
  widths=[1900, 2600, 3300, 1600]))
B(caption("Tabla", "Comparación de los referentes del Norte Global. El color identifica a cada institución; las columnas "
  "indican su modelo de distribución de la decisión, sus órganos distintivos y el carácter de sus decisiones."))
B(pagebreak())

# ====================================================================
#   3. REFERENTES DEL SUR GLOBAL  (AMPLIACIÓN)
# ====================================================================
B(h1("3. Referentes del Sur Global (ampliación)"))
B(para([run("Ampliación de esta versión. ", b=True, color=DORADO),
  run("Los referentes del Norte operan en contextos y puntos de partida distintos al colombiano. Por ello el ejercicio se "
      "amplió con referentes del Sur Global, más cercanos a la realidad de la UNAL —universidades públicas, nacionales, "
      "multicampus y con restricciones de recursos—, que permiten aprendizajes más transferibles.")]))
B(h3("3.1 Tec de Monterrey (México) — multicampus"))
B(para("Referente multicampus latinoamericano con lineamientos de integridad académica para el uso de IA. Aporta la "
  "experiencia de coordinar criterios comunes entre sedes distribuidas, un reto análogo al de las nueve sedes de la UNAL."))
B(h3("3.2 Universidad de São Paulo — USP (Brasil) — gran IES pública"))
B(para("La mayor institución de educación superior pública de la región, con un enfoque normativo. Muestra cómo una "
  "universidad pública de gran escala puede regular el uso de la IA preservando autonomía y misión pública."))
B(h3("3.3 UNAM (México) — proceso en construcción"))
B(para("Universidad pública nacional cuyo proceso de gobernanza de IA está en construcción. Es un espejo cercano a la UNAL: "
  "comparte la condición de universidad pública nacional y la necesidad de un modelo propio y participativo."))
B(h3("3.4 Proyecto INTEGRITY-AI — red latinoamericana con pares colombianos"))
B(para("Red latinoamericana de integridad académica y gobernanza de IA que incluye pares colombianos como la Universidad del "
  "Cauca y la Universidad de San Buenaventura. Aporta el valor del trabajo en red regional y la validación entre pares del Sur."))
B(h2("Síntesis de los referentes del Sur Global"))
B(table(
  [[("Institución / red", {}), ("Rasgo", {}), ("Aprendizaje para la UNAL", {})],
   [("Tec de Monterrey", {}), ("Multicampus (LatAm)", {}), ("Coordinar criterios comunes entre sedes.", {})],
   [("USP (Brasil)", {}), ("Normativo · gran IES pública", {}), ("Regular preservando misión pública y escala.", {})],
   [("UNAM (México)", {}), ("Público nacional · en construcción", {}), ("Modelo propio, participativo y situado.", {})],
   [("INTEGRITY-AI", {}), ("Red LatAm (U. Cauca, U. San Buenaventura)", {}), ("Trabajo en red y validación entre pares.", {})]],
  widths=[2600, 3400, 3400]))
B(caption("Tabla", "Referentes del Sur Global. Cada fila relaciona un referente con su rasgo distintivo y el aprendizaje "
  "transferible al contexto de la UNAL."))
B(pagebreak())

# ====================================================================
#   4. MATRIZ COMPARATIVA INTEGRADA
# ====================================================================
B(h1("4. Matriz comparativa integrada"))
B(para("La siguiente matriz integra referentes del Norte y del Sur bajo los criterios de análisis. La UNAL no adopta ninguno "
  "tal cual: construye un modelo propio que toma la cercanía al problema (MIT), el control fuerte donde el riesgo es alto "
  "(UCLA), la producción de estándares y evidencia (Stanford) y la coordinación multicampus y la escala pública del Sur."))
B(table(
  [[("Institución", {}), ("Modelo", {}), ("Carácter", {}), ("Enfoque de riesgo", {}), ("Ambiental", {})],
   [("MIT", {"color":"FFFFFF","fill":MIT_C}), ("Descentralizado", {}), ("Consultivo", {}), ("Por dominio", {}), ("Incipiente", {})],
   [("UCLA", {"color":"FFFFFF","fill":UCLA_C}), ("Matricial-central.", {}), ("Mandatorio (salud)", {}), ("Alto en clínico", {}), ("Incipiente", {})],
   [("Stanford", {"color":"FFFFFF","fill":STAN_C}), ("Hub-and-spoke", {}), ("Orientador", {}), ("Por estándares", {}), ("Incipiente", {})],
   [("Tec Monterrey", {}), ("Multicampus", {}), ("Normativo", {}), ("Integridad acad.", {}), ("Parcial", {})],
   [("USP", {}), ("Normativo", {}), ("Mandatorio", {}), ("Institucional", {}), ("Parcial", {})],
   [("UNAM", {}), ("En construcción", {}), ("Mixto", {}), ("En definición", {}), ("Parcial", {})],
   [("UNAL (propuesto)", {"color":"FFFFFF","fill":VERDE,"bold":True}), ("Multinivel + 4 núcleos", {"bold":True}), ("Colegiado (quórum)", {"bold":True}), ("Por derechos", {"bold":True}), ("Green AI (eje)", {"bold":True})]],
  widths=[2200, 2200, 1900, 1800, 1300]))
B(caption("Tabla", "Matriz comparativa integrada. Las primeras filas son referentes del Norte (con color) y del Sur Global; "
  "la última fila resalta el modelo propuesto por la UNAL, que gradúa el riesgo por derechos y adopta Green AI como eje, no "
  "como consideración incipiente."))
B(d.chart("col", ["MIT","UCLA","Stanford","Tec Mty","USP","UNAM","UNAL"],
          [1,1,1,2,2,2,3],
          [MIT_C, UCLA_C, STAN_C, AMBAR, SOCIO, PUR, VERDE],
          "Madurez de la dimensión ambiental (1 incipiente · 2 parcial · 3 eje)", w_in=6.2, h_in=2.9))
B(caption("Gráfico", "Madurez de la dimensión ambiental por referente (objeto de gráfico editable de Word). La escala va de "
  "1 (incipiente) a 3 (eje transversal); la UNAL es la única que eleva Green AI a eje del modelo."))
B(pagebreak())

# ====================================================================
#   5. GRAFO DE RELACIONES
# ====================================================================
B(h1("5. Grafo de relaciones entre referentes y el modelo UNAL"))
B(para("El benchmarking no es una suma de casos aislados: existen influencias y afinidades entre referentes y hacia el modelo "
  "de la UNAL. El siguiente grafo representa esas relaciones. En el centro, el modelo UNAL recibe aportes del Norte (lógicas "
  "de decisión) y del Sur (escala pública y coordinación multicampus)."))
B(d.diagram(
  boxes=[
    {"x":390,"y":420,"w":240,"h":170,"text":"MODELO\nUNAL","fill":VERDE,"fs":950},
    {"x":40,"y":40,"w":210,"h":130,"text":"MIT","fill":MIT_C,"fs":1000},
    {"x":395,"y":30,"w":230,"h":120,"text":"UCLA","fill":UCLA_C,"fs":1000},
    {"x":760,"y":40,"w":210,"h":130,"text":"Stanford","fill":STAN_C,"fs":900},
    {"x":40,"y":830,"w":210,"h":130,"text":"Tec Mty / USP","fill":SOCIO,"fs":800},
    {"x":395,"y":840,"w":230,"h":120,"text":"UNAM","fill":PUR,"fs":950},
    {"x":760,"y":830,"w":210,"h":130,"text":"INTEGRITY-AI","fill":AMBAR,"fs":780},
  ],
  arrows=[
    (145,170,440,420,MIT_C,{"tail":"stealth","label":"núcleos"}),
    (510,150,510,420,UCLA_C,{"tail":"stealth","label":"veto"}),
    (865,170,580,440,STAN_C,{"tail":"stealth","label":"estándares"}),
    (145,830,440,590,SOCIO,{"tail":"stealth","label":"multicampus"}),
    (510,840,510,590,PUR,{"tail":"stealth","label":"público"}),
    (865,830,580,560,AMBAR,{"tail":"stealth","label":"red de pares"}),
  ],
  w_in=6.3, h_in=3.6, title="Aportes de los referentes al modelo UNAL"))
B(caption("Diagrama", "Grafo de relaciones. Las flechas superiores (Norte: MIT, UCLA, Stanford) aportan lógicas de decisión "
  "al modelo UNAL; las inferiores (Sur Global) aportan escala pública, coordinación multicampus y validación entre pares. "
  "Cada etiqueta nombra el aporte principal."))
B(pagebreak())

# ====================================================================
#   6. DIAGNÓSTICO UNAL
# ====================================================================
B(h1("6. Diagnóstico de la UNAL a la luz del benchmarking"))
B(para("Contrastada con los referentes, la UNAL presenta fortalezas, oportunidades y necesidades propias de una universidad "
  "pública, nacional, pluriétnica y multicampus. El diagnóstico orienta las decisiones de diseño del modelo."))
B(table(
  [[("Dimensión", {}), ("Situación de la UNAL", {}), ("Decisión de diseño", {})],
   [("Escala y distribución", {}), ("Nueve sedes con capacidades heterogéneas.", {}), ("Subsidiariedad: decidir cerca del riesgo y el valor.", {})],
   [("Carácter público", {}), ("Misión pública y autonomía universitaria.", {}), ("Modelo propio, no adopción directa de referentes.", {})],
   [("Riesgo", {}), ("Casos académicos pueden afectar derechos (becas, deserción).", {}), ("Graduar el control por derechos, no por dominio.", {})],
   [("Sostenibilidad", {}), ("Compromisos socioambientales institucionales.", {}), ("Green AI como eje transversal del modelo.", {})],
   [("Participación", {}), ("Comunidad amplia y multiestamentaria.", {}), ("Co-construcción: 129 personas, 14 mesas, 9 sedes.", {})]],
  widths=[2400, 3800, 3200]))
B(caption("Tabla", "Diagnóstico y decisiones de diseño. Cada fila cruza una dimensión del diagnóstico con la situación de la "
  "UNAL y la decisión de diseño que de ella se deriva."))

# ====================================================================
#   7. EL MODELO UNAL RESULTANTE
# ====================================================================
B(h1("7. El modelo UNAL resultante del benchmarking"))
B(para([run("Lectura hacia el modelo propio. ", b=True, color=DORADO),
  run("Los aprendizajes del benchmarking se tradujeron en un modelo de gobernanza multinivel, participativo y multicampus, "
      "cuyos rasgos operativos —hoy representados en un gemelo digital 3D y un dashboard interactivo— se resumen a "
      "continuación.")]))
for t in [
    "Arquitectura multinivel: CNED → Mesa Técnica de IA → Comunidad universitaria (nueve sedes).",
    "Cuatro núcleos funcionales: Estratégico (1.1), Técnico & Green (1.2), Control & Jurídico (1.3) y Socio-Ambiental (1.4).",
    "Operación colegiada: Plenaria de Delegados que radica casos, los traslada a los núcleos y decide por quórum (≥ 3 de 4).",
    "Matriz RACI que hace trazable «quién decide qué» a lo largo del ciclo PLASI.",
    "Graduación del riesgo por afectación a derechos (Bajo · Medio · Alto · Inaceptable), no por dominio de aplicación.",
    "Subsidiariedad: lo cotidiano y de bajo riesgo se resuelve en la facultad o la sede; se centraliza por excepción.",
    "Green AI y soberanía del conocimiento como ejes transversales.",
    "Hoja de ruta 2025–2027: diagnóstico y co-construcción → validación y aprobación → implementación y evaluación."]:
    B(bullet(t))
B(h2("Qué toma la UNAL de cada referente"))
B(table(
  [[("Referente", {}), ("Aporte incorporado al modelo UNAL", {})],
   [("MIT", {"color":"FFFFFF","fill":MIT_C}), ("Cercanía al problema: núcleos especializados.", {})],
   [("UCLA", {"color":"FFFFFF","fill":UCLA_C}), ("Control fuerte donde el riesgo es alto: veto y escalamiento.", {})],
   [("Stanford", {"color":"FFFFFF","fill":STAN_C}), ("Estándares y evidencia: guías, DT y métricas.", {})],
   [("Sur Global", {"color":"FFFFFF","fill":SOCIO}), ("Escala pública y coordinación multicampus; red de pares.", {})]],
  widths=[2400, 7000]))
B(caption("Tabla", "Aportes incorporados. Cada fila indica qué toma el modelo UNAL de cada referente, sintetizando la lectura "
  "del benchmarking hacia un modelo propio."))
B(pagebreak())

# ====================================================================
#   8. CONCLUSIONES
# ====================================================================
B(h1("8. Conclusiones"))
for i, t in enumerate([
    "No existe un modelo único: los referentes resuelven de forma distinta «quién decide qué», con qué carácter y bajo qué enfoque de riesgo.",
    "El Sur Global aporta aprendizajes más transferibles a la UNAL (escala pública, multicampus, trabajo en red) que la adopción directa del Norte.",
    "La UNAL construye un modelo propio: multinivel, colegiado, con graduación del riesgo por derechos y Green AI como eje.",
    "La operación colegiada (Plenaria + quórum) y la matriz RACI dan trazabilidad y legitimidad a la decisión.",
    "La validación académica externa y la participación multiestamentaria (129 personas) sostienen la robustez del modelo."], 1):
    B(numitem(i, t))

# ====================================================================
#   9. REFERENCIAS
# ====================================================================
B(h1("9. Referencias y documentación"))
B(para("La documentación completa del proceso de referenciación y del modelo resultante se encuentra en el repositorio "
  "institucional de la DNED–DIVULAB:"))
B(para([run("Repositorio de documentación: ", b=True, sz=22, color=NEGRO),
  d.hyperlink("https://drive.google.com/drive/folders/1zZ3nCgpte0JUMN6NLxNNcKHmAxAID6Jj",
              "https://drive.google.com/drive/folders/1zZ3nCgpte0JUMN6NLxNNcKHmAxAID6Jj")], align="left"))
for r in [
    "MIT Schwarzman College of Computing — Social and Ethical Responsibilities of Computing (SERC).",
    "University of California / UCLA — UC AI Council; Health AI Committee (HAIC).",
    "Stanford HAI (Human-Centered AI); AIMES; Stanford AI Index.",
    "Tec de Monterrey — lineamientos de integridad académica con IA.",
    "Universidad de São Paulo (USP) — regulación de IA en una gran IES pública.",
    "UNAM — proceso de gobernanza de IA (en construcción).",
    "Proyecto INTEGRITY-AI — red latinoamericana (incluye U. del Cauca y U. de San Buenaventura).",
    "Documento de marco conceptual y estado del arte para el uso ético de la IA. Ing. Wilver Alexander Martínez, DNED.",
    "Propuesta de creación de la Mesa Técnica de IA como sala adscrita al CNED (documento complementario).",
    "Gemelo digital 3D y dashboard interactivo del modelo de gobernanza de la IA (DNED–DIVULAB)."]:
    B(bullet(r))
B(nota("Documento elaborado con apoyo de herramientas de IA generativa; contenido validado contra los documentos "
  "institucionales de la UNAL. DNED–DIVULAB."))

# ---- guardar ----
doc = d.save("benchmarking-gobernanza-ia-internacional.docx")
import os
print("DOCX OK:", os.path.getsize("benchmarking-gobernanza-ia-internacional.docx"), "bytes")
print("Parrafos:", doc.count("<w:p>"), "| Tablas:", doc.count("<w:tbl>"),
      "| Graficos:", len(d._charts), "| Saltos:", doc.count('w:type="page"'))
