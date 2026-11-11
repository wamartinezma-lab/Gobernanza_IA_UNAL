# Gobernanza de la IA en la UNAL · Mesa Técnica de IA

Dos piezas interactivas (HTML autocontenido) del **Modelo de Gobernanza Multinivel de la Inteligencia Artificial en la Universidad Nacional de Colombia**, centradas en la **Mesa Técnica de IA** (Sala Adscrita al CNED).

Elaborado para **DNED–DIVULAB** (Dirección Nacional de Estrategia Digital · División Universidad Laboratorio).

## 📂 Contenido

| Archivo | Qué es |
|---------|--------|
| `gemelo-digital-gobernanza-ia-unal.html` | **Gemelo digital 3D** interactivo (Three.js) — modelo navegable en 3D. |
| `cerebro-digital-gobernanza-ia-unal.html` | **Cerebro digital** — red de conocimiento interactiva (SVG + JS, sin dependencias) que conecta todos los conceptos y entregables del repositorio en lóbulos temáticos; nodos clicables, panel de detalle, pan/zoom y pulsos de "pensamiento". |
| `dashboard-gobernanza-ia-unal.html` | **Dashboard** por secciones (Chart.js) — KPIs, gráficos, tablas, matrices y simulador. |
| `roadmap-exposicion-30min.html` | **Roadmap/guion de exposición (30 min)** en modo presentador, con cronómetro por bloque. |
| `roadmap-exposicion-30min.md` | El mismo guion en Markdown (editable, con checklist y preguntas frecuentes). |
| `brief-ponencia-gobernanza-ia-unal.md` | **Brief de la ponencia** (abstract ES/EN, bio, cupo, objetivos, agenda, metodología, aportes y requerimientos). |
| `brief-ponencia-gobernanza-ia-unal.pdf` | El brief en **PDF** (2 páginas) listo para envío. |
| `brief-ponencia-gobernanza-ia-unal.docx` | El brief en **Word (.docx)** editable. |
| `generar_brief.py` | Script Python (sin dependencias) que genera el PDF y el DOCX del brief. |
| `documento-modelo-gobernanza-ia-unal.docx` | **Documento Word institucional**: cuerpo principal (avance y rol de la DNED) + 9 anexos (A–I) del modelo, con tablas y gráficos. |
| `generar_documento.py` | Script Python (sin dependencias) que genera el documento Word anterior. |
| `propuesta-mesa-tecnica-ia-cned.docx` | **Documento Word — Propuesta de creación de la Mesa Técnica de IA** (sala adscrita al CNED), versión actualizada con el modelo operativo (plenaria, quórum, RACI, riesgo, subsidiariedad, hoja de ruta), con tablas, gráficos y diagramas nativos. |
| `benchmarking-gobernanza-ia-internacional.docx` | **Documento Word — Benchmarking internacional** (MIT · UCLA · Stanford) ampliado con el **Sur Global** (Tec de Monterrey · USP · UNAM · INTEGRITY-AI) y la lectura hacia el modelo UNAL. |
| `ooxml_docx.py` | Módulo (sin dependencias) con el *toolkit* OOXML compartido (textos, tablas, gráficos nativos, hipervínculos y **diagramas vectoriales DrawingML**). |
| `generar_propuesta_mesa_tecnica.py` / `generar_benchmarking.py` | Scripts Python que generan los dos documentos Word anteriores. |

## 📘 Documento Word del modelo (principal + anexos)

`documento-modelo-gobernanza-ia-unal.docx` presenta todo el modelo y el trabajo adelantado:
- **Documento principal:** portada · presentación · **avance institucional** (tabla de indicadores + gráfico) · **rol articulador de la DNED–DIVULAB** · síntesis e índice de anexos.
- Reconoce al **ing. Wilver Alexander Martínez** como quien inició el trabajo con el **estado del arte**, sintetizado en el **documento de marco conceptual** que delineó el alcance del uso de la IA y los principios de uso responsable y apropiación ética.
- **Anexos (A–J), máx. 2 páginas c/u, con tablas/gráficos:** A) benchmarking · B) estructura multinivel y subsidiariedad · C) los 4 núcleos · D) operación (plenaria, quórum, votación) · E) riesgo y taxonomía · F) matriz RACI · G) participación (129 personas = 108 + 21, por rol) · H) recomendaciones académicas · I) hoja de ruta 2025–2027 · **J) listado de guías y documentos de lineamientos** de IA construidos.

Cada anexo incluye **al menos 2 párrafos descriptivos** y un **pie de figura (caption)** que explica cómo leer la tabla o el gráfico. Se genera con `python3 generar_documento.py` (OOXML sin librerías externas). Los gráficos de **avance de documentos** (dona) y **participación por rol** (barras) son **objetos de gráfico nativos de Word (editables)**: al hacer doble clic se pueden modificar datos, colores y tipo.

## 📄 Documentos Word actualizados (propuesta + benchmarking)

Dos documentos Word que **actualizan los PDF institucionales** con los conceptos nuevos del gemelo 3D y el dashboard, integrando **tablas, gráficos nativos editables y diagramas vectoriales (DrawingML)**:

- **`propuesta-mesa-tecnica-ia-cned.docx`** — *Propuesta de creación de la Mesa Técnica de IA como sala adscrita al CNED.* Conserva la estructura original (contexto normativo —Resolución 614/2020, Circular 01/2024—, marco multinivel, funciones, ciclo de vida **PLASI**, cuatro núcleos, Directrices Técnicas **DT 1–25**, lineamientos y guías) e **integra**: Plenaria de Delegados, quórum y votación, matriz **RACI**, taxonomía y **niveles de riesgo** (Bajo/Medio/Alto/Inaceptable), **subsidiariedad**, participación **129** (108+21) y **hoja de ruta 2025–2027**. Incluye 10 tablas, 2 gráficos nativos y 5 diagramas (ubicación de la gobernanza, arquitectura multinivel, interacción de núcleos, flujo de la plenaria y línea de tiempo).
- **`benchmarking-gobernanza-ia-internacional.docx`** — *Benchmarking internacional de gobernanza de IA.* Conserva la estructura original (marco teórico, MIT/UCLA/Stanford, matriz comparativa, grafo de relaciones, diagnóstico UNAL, conclusiones) y lo **amplía** con el **Sur Global** (Tec de Monterrey, USP, UNAM e **INTEGRITY-AI** con la U. del Cauca y la U. de San Buenaventura) y la **lectura hacia el modelo UNAL resultante** (multinivel, 4 núcleos, plenaria, RACI, riesgo por derechos, subsidiariedad, Green AI, hoja de ruta). Incluye 6 tablas, 1 gráfico nativo y el **grafo de relaciones** como diagrama vectorial.

Ambos comparten el *toolkit* `ooxml_docx.py` y se generan sin librerías externas:

```bash
python3 generar_propuesta_mesa_tecnica.py   # -> propuesta-mesa-tecnica-ia-cned.docx
python3 generar_benchmarking.py             # -> benchmarking-gobernanza-ia-internacional.docx
```

Los diagramas vectoriales se envuelven en `mc:AlternateContent` para asegurar su visualización en versiones recientes y antiguas de Word (y en LibreOffice).
| `presentacion-dashboard-gobernanza-ia-unal.pptx` | **Presentación PowerPoint** (14 diapositivas, 16:9) del dashboard para el evento de 30 min, basada en el brief, con diagramas y **2 gráficos nativos editables**. |
| `generar_pptx.py` | Script Python (sin dependencias) que genera la presentación `.pptx`. |

## 🖥️ Presentación PowerPoint (evento 30 min)

`presentacion-dashboard-gobernanza-ia-unal.pptx` — 14 diapositivas en 16:9, basadas en la agenda del brief:
portada · apertura · el problema (diagrama de conjuntos IA⊃TI∩datos) · ecosistema institucional · benchmarking (Norte/Sur Global) · modelo multinivel · 4 núcleos · operación de la plenaria · **votación (gráfico de barras editable)** · riesgo y taxonomía · matriz RACI · hoja de ruta 2025–2027 · **avance (gráfico de dona editable)** · cierre.

Los diagramas se construyen con formas nativas y los dos gráficos son objetos de gráfico de PowerPoint **totalmente editables** (datos, colores, tipo). Se genera con `python3 generar_pptx.py` (OOXML sin librerías externas).

**Animaciones (selectivas):** 9 de las 14 diapositivas tienen animación donde aporta (diagramas que se construyen por pasos); las densas en texto quedan estáticas para facilitar la lectura. La **apertura** y el **flujo de la plenaria** usan una narrativa **tipo cómic** —viñetas, globos de diálogo y estrellas "POW"— que aparecen en secuencia. Hay transición de desvanecido entre diapositivas. El contenido de cada diapositiva está enriquecido con detalle explicativo (órganos de benchmarking, composición de núcleos, descripción de niveles de riesgo, leyenda RACI, etc.).
| `roadmap-exposicion-gobernanza-ia-unal.pdf` | **PDF listo para envío** (7 páginas): brief + gráficos del roadmap (línea de tiempo, minutos por bloque, flujo de la demo, plenaria y hoja de ruta). |
| `generar_pdf_roadmap.py` | Script Python (sin dependencias) que genera el PDF anterior. |

## 📑 Brief de la ponencia

El documento ejecutivo de la charla, listo para el programa académico o para enviar a los organizadores. Incluye: ficha (ponente, formato, duración, nivel, **cupo/aforo**), **resumen/abstract en español y en inglés**, palabras clave, contexto y problema, objetivos (general + específicos), público objetivo, **agenda de 30 min**, metodología (con la demo en vivo), mensajes clave/aportes, resultados esperados, requerimientos técnicos, **bio del ponente** y referencias.

Disponible en tres formatos:
- **Markdown** (`.md`) — editable.
- **PDF** (`.pdf`, 2 páginas) — listo para envío; se genera con `python3 generar_brief.py`.
- **Word** (`.docx`) — editable en Microsoft Word / LibreOffice; mismo generador.

Ambos (PDF y DOCX) se producen **sin librerías externas** (el `.docx` es OOXML construido con `zipfile`). El brief también aparece como página 2 del PDF del roadmap.

## 🎤 Roadmap de exposición (30 min)

Para presentar el modelo en un evento académico. Abre `roadmap-exposicion-30min.html` (modo presentador):
- **9 bloques** que suman exactamente **30 min**: apertura, problema, benchmarking, modelo UNAL, operación de la Mesa, riesgo/RACI, **demo en vivo**, recomendaciones/hoja de ruta y cierre.
- **Cronómetro doble:** tiempo del bloque actual (se pone en rojo si te excedes) y total restante; barra de progreso proporcional por bloque.
- Cada bloque indica el **mensaje clave**, los puntos a decir, **qué mostrar en pantalla** (gemelo 3D / dashboard) y notas para el expositor.
- Navegación con flechas del teclado (`←` `→`) y `Espacio` para iniciar/pausar.
- La versión `.md` incluye además checklist del expositor y respuestas rápidas a preguntas frecuentes.

### 📄 PDF para envío
`roadmap-exposicion-gobernanza-ia-unal.pdf` (6 páginas A4) es un documento imprimible/enviable con **gráficos vectoriales**:
- **Portada** + línea de tiempo proporcional de los 9 bloques y gráfico de barras de minutos por bloque.
- **Guion por bloques** (mensaje clave, puntos a decir y qué mostrar en pantalla).
- **Gráficos de apoyo:** flujo de la demo (Facultad → … → CNED), diagrama de la Plenaria con los 4 núcleos y hoja de ruta 2025–2027.

Se genera con `python3 generar_pdf_roadmap.py` (no requiere librerías externas; produce un PDF 1.4 estándar).

## 📊 Dashboard

Abre `dashboard-gobernanza-ia-unal.html` en cualquier navegador (Chart.js se carga desde CDN). Panel con navegación lateral y 8 secciones:

1. **Resumen general** — KPIs (4 núcleos, 14 mesas, 11 entregados, 3 pendientes, 9 sedes, 10 recomendaciones) + gráficos.
2. **Estructura de gobernanza** — niveles CNED → Mesa → Comunidad, los 4 núcleos y la arquitectura de decisiones (subsidiariedad).
3. **Procesos y ciclo** — plan 2025–2027 (5 fases), ciclo de vida de modelos, PLASI, Green Computing.
4. **Gestión de riesgo** — 4 niveles de riesgo, taxonomía de sistemas de IA, gobernanza corporativa (IA ⊃ TI ∩ datos).
5. **Roles (RACI)** — matriz quién propone/evalúa/autoriza/implementa/supervisa/suspende.
6. **Modelos de referencia** — benchmarking Norte y Sur Global + matriz comparativa.
7. **Simulador de operación** — recorre el procedimiento de una iniciativa de IA por todo el modelo (Facultad → Sede → Mesa Técnica → CNED → Operación).
8. **Sesiones y documentos** — estado de las 14 mesas (S1–S4) con gráficos.
9. **Participantes por mesa** — modela la matriz **persona × mesa** del archivo oficial (cada “X” = una persona participa en una mesa). Incluye: **totales de control** (107 participantes + 21 facilitadores DNED = **128 personas**), selector por mesa con la tabla de participantes (nombre, rol, dependencia/sede), gráfico de participantes por mesa, **conteo por rol** (gráfico de dona + tabla, suma = 128), **matriz de participación con ✓** y resumen general. Los facilitadores se cuentan como participantes de cada mesa. Nombres como placeholders en `PERSONAS` y conteos en `ROLES`/`TOTALES` (en el `<script>`), listos para reemplazar con el listado oficial.
10. **Recomendaciones** — las 10 de J.M. Sánchez Torres con su incorporación.

### 🎮 Simulador de operación
Permite **configurar una iniciativa de IA** (tipo de sistema, origen, datos y alcance) y simular su recorrido por el modelo de gobernanza:
- Calcula automáticamente el **nivel de riesgo** (Bajo / Medio / Alto / Inaceptable) por su afectación a derechos fundamentales.
- Aplica **subsidiariedad**: resuelve en el nivel más cercano (Facultad / Sede) o **escala al CNED** cuando el riesgo lo exige.
- Anima el flujo **paso a paso** mostrando en cada etapa quién actúa según la **matriz RACI** (R/A/C/I/S).
- Entrega un **veredicto** (autorizada local, con condiciones, escalada o vetada) y 5 escenarios de ejemplo (deserción, becas, tutor en aula, agente autónomo, chatbot).

**Operación de la Mesa — Plenaria de Delegados:** al llegar a la Mesa Técnica (etapa 3), el simulador modela su funcionamiento como un **órgano colegiado**:
- Los casos/iniciativas se radican en la **Plenaria de Delegados** (un delegado por núcleo: D1–D4), que hace el **análisis preliminar**.
- La plenaria **traslada a trámite**: cada delegado lleva el caso a su núcleo para la evaluación especializada.
  - **1.1 Estratégico Misional** → pertinencia misional y soberanía epistémica.
  - **1.2 Técnico & Green** → viabilidad técnica, HPC, huella de carbono.
  - **1.3 Control & Jurídico-Fiscal** → legalidad, Habeas Data, RAEE, presupuesto.
  - **1.4 Socio-Ambiental & Territorial** → ética, derechos, territorio, salud mental.
  - Durante el trámite, **1.2 ↔ 1.3** se consultan mutuamente.
- Los delegados **regresan a la plenaria con su ponencia**; la plenaria **delibera con quórum** y consolida el **dictamen colegiado** (Favorable / Con salvaguardas / Objeción) que alimenta la decisión por subsidiariedad.

**Quórum y votación:** la plenaria muestra un panel de votación con:
- **Quórum** (presentes / 4, válido si ≥ 3 delegados) con indicador visual.
- **Conteo de votos** en barra proporcional: a favor · con salvaguardas · objeción.
- **Resultado** y **regla de decisión** aplicada. Regla jerárquica por protección de derechos: una objeción de **1.3 (Jurídico)** o **1.4 (Socio-Ambiental)** por riesgo a derechos **prevalece** (veto / escalar al CNED); en los demás casos se resuelve por salvaguardas, unanimidad o mayoría.
- **Mini-diagrama SVG animado:** la operación de la plenaria se muestra como un grafo con la **Plenaria** al centro y los 4 núcleos; una **animación del flujo de trabajo** ilumina paso a paso la radicación, el traslado a trámite, la consulta 1.2↔1.3, el regreso con ponencias y la deliberación, con un pulso que recorre las conexiones y el veredicto final por núcleo. Incluye botón **▶ Reproducir animación**.
- **Tooltips de interacción:** al pasar el mouse (o tocar) una flecha se muestra el detalle del hand-off. En la flecha bidireccional **1.2 ↔ 1.3** se despliega lo que cada núcleo le pregunta al otro: el Técnico consulta al Jurídico sobre legalidad del tratamiento de datos, licencias y presupuesto; el Jurídico consulta al Técnico sobre ciberseguridad, alojamiento de datos (Habeas Data), consumo energético y RAEE.

## 🧠 Cerebro digital (red de conocimiento + modelo multi-agente)

`cerebro-digital-gobernanza-ia-unal.html` es un **mapa de conocimiento interactivo** que fusiona el modelo de gobernanza de la UNAL con el **modelo multi-agente** del vault `AI-Governance-Brain` (patrón *LLM Wiki* de Karpathy). Es **100 % autocontenido** (SVG + JavaScript puro, sin librerías ni CDN: funciona sin conexión).

- **35 nodos** y **65 sinapsis** organizados en **8 lóbulos temáticos:** Agentes · Núcleo de decisión · Estándares y frameworks · Conceptos · Gobernanza UNAL · Mesas de co-creación · Skills · Entregables.
- **Modelo multi-agente incorporado** (del ZIP): los 4 agentes de valor **Freedom · Equality · Safety · Culture**, el **Protocolo de decisión**, **Escalada y deliberación** y los **umbrales de daño (0–3)**.
- **Estándares y conceptos:** **AI IAM Standard**, **NIST Mapping**, **Non-Human Identity**, **Least Privilege** y **Runtime Control Plane**.
- **Skills del vault** como nodos: `deliberate-agents`, `risk-assessment`, `ingest-source`, `generate-report`.
- **Capas del grafo (Karpathy)** en la leyenda: Core · Estándares · Conceptos · Gobernanza UNAL · Entregables (para aislar niveles), además del encendido/apagado por lóbulo.
- **⚖ Deliberar:** modal que aplica el *Decision Protocol* con **6 casos de ejemplo precargados** (predictor de deserción, triage clínico autónomo, detección de IA en ensayos, moderación de contenido, agente con acceso a sistemas, asistente creativo masivo). Describes o eliges un caso y el cerebro activa los 4 agentes, muestra la posición de cada uno, estima el nivel de riesgo (0–3) y da un nivel de confianza.
- **Interacción entre agentes:** un diagrama muestra los 4 agentes y traza sus **alianzas** (coinciden, en verde) y **conflictos** de valor (trade-offs, en ámbar punteado) según el caso deliberado.
- **🧬 Evolución persistente (el cerebro aprende y crece):** cada deliberación puede **registrarse como un nodo *Insight*** que se añade en vivo al grafo, conectado al protocolo y a los agentes activados. Los insights se **guardan en `localStorage`**, de modo que el cerebro **conserva lo aprendido entre sesiones** (al recargar vuelven a aparecer). Botones **🧽 Olvidar** / enlace "Olvidar aprendizajes" para reiniciar la memoria. Un **contador de crecimiento** muestra nodos · sinapsis · insights en tiempo real.
- **Integrado al dashboard:** el dashboard incluye la pestaña **🧠 Cerebro digital**, que embebe el cerebro (iframe con carga diferida) y un enlace para abrirlo a pantalla completa.
- **Gráficos profesionales:** nodos con **gradientes radiales** por lóbulo, brillo especular (volumen), resplandor suave (filtro SVG), halos y viñeta de fondo.
- **Interactividad:** clic en un nodo resalta conexiones y abre un **panel de detalle** (descripción, cita del agente cuando aplica, puntos clave y *chips* navegables); **pan**, **zoom**, **⟳ Centrar** y **⚡ Pensar / ⏸ Pausa** (pulsos de "pensamiento").
- El nodo **Repositorio GitHub** enlaza al repo del proyecto.

> El contenido, diseño (capas por color) y funcionalidad (deliberación, skills) provienen de `AI-Governance-Brain.zip`, que también está en el repositorio como vault de Obsidian/Claude.

## 🧊 Gemelo digital 3D

Abre `gemelo-digital-gobernanza-ia-unal.html` en cualquier navegador moderno (Chrome, Firefox, Edge, Safari). Es un **archivo autocontenido**; Three.js se carga desde CDN (requiere conexión a internet la primera vez).

## 🏛️ Qué representa

Estructura 3D de 4 niveles jerárquicos:

| Nivel | Elemento | Contenido |
|-------|----------|-----------|
| **1 — Alto nivel** | CNED | Definición política y aprobación normativa |
| **2 — Ejecución** ⭐ | Mesa Técnica de IA | Sala adscrita · 4 Núcleos · PLASI · Ciclo trienal |
| **3 — Núcleos** | 4 bloques | Estratégico Misional · Técnico & Green Computing · Control & Jurídico-Fiscal · Socio-Ambiental & Territorial |
| **4 — Base** | Comunidad Universitaria | Laboratorios vivos · 9 sedes · facultades · grupos |

### Extensiones comparativas
- **🛰 Satélites de referencia (dos órbitas):**
  - *Órbita Norte Global:* MIT (descentralizado), UCLA (matricial-centralizado), Stanford (hub-and-spoke).
  - *Órbita Sur Global / Latinoamérica (Rec. 1):* Tec de Monterrey, USP, UNAM y el proyecto **INTEGRITY-AI** (con la U. del Cauca y la U. de San Buenaventura como pares colombianos).
- **📅 Línea de tiempo 2025–2027:** las 5 fases del plan de trabajo (Diagnóstico → Co-construcción → Validación → Revisión y aprobación → Implementación), con responsables DNED y Mesa Técnica · GIIA-UNAL.

### Recomendaciones académicas incorporadas
Comentarios de la **Prof. Jenny Marcela Sánchez Torres** (Facultad de Ingeniería, 01-oct-2026) traducidos a vistas interactivas:

| Botón | Vista | Recomendaciones |
|-------|-------|-----------------|
| 🏛 | **Arquitectura de decisiones** (nacional / sede / facultad / compartido) + subsidiariedad | 2, 9 |
| 🧭 | **Matriz RACI por riesgo** (propone / evalúa / autoriza / implementa / supervisa / suspende) | 8, 4 |
| 🧩 | **Tipos de IA y riesgo** (taxonomía por propósito-origen + 4 niveles de riesgo) | 4, 5, 6 |
| ⊚ | **Gobernanza corporativa** (IA ⊃ gobierno de TI ∩ gobierno de datos) | 3, 7 |
| 📝 | **Recomendaciones aplicadas** (trazabilidad de las 10) | todas |

Además, un **nodo transversal 3D "Capacidades Habilitantes & Legitimidad"** (Rec. 10) sostiene lateralmente el modelo, y el Núcleo Técnico incorpora la nota de **marcos ambientales reconocidos** (Rec. 7, prof. Moreno Mantilla / Of. Gestión Ambiental sede Bogotá).

## 🖱️ Interactividad

- **Arrastrar** = rotar · **rueda** = zoom · **giro automático** activable.
- **Clic en cualquier bloque, satélite o fase** → panel con su composición orgánica y funciones.
- Barra de herramientas: vista orbital, cenital, enfoque a la Mesa, modelos de referencia, línea de tiempo, mostrar/ocultar satélites y pausar giro.

## 📚 Fuentes

Contenido extraído de los documentos institucionales:
- Presentación *Gobernanza IA* (DNED–DIVULAB).
- *Propuesta de creación de la Mesa Técnica de Inteligencia Artificial como Sala Adscrita al CNED* (benchmarking MIT / UCLA / Stanford).

---

*Elaborado con apoyo de herramientas de IA generativa; contenido validado contra los documentos institucionales de la UNAL.*
