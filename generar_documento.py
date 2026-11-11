#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un documento Word (.docx) del Modelo de Gobernanza de la IA en la UNAL:
documento principal (avance institucional + rol de la DNED) y anexos por item
(máx. 2 páginas c/u), con tablas y gráficos NATIVOS de Word (objetos chart editables).
Sin dependencias externas: OOXML construido con zipfile.
Salida: documento-modelo-gobernanza-ia-unal.docx
"""
import zipfile, html, time, os

def xe(t): return html.escape(str(t),quote=True)

# ---------- colores (hex sin #) ----------
VERDE="1F4037"; DORADO="7A6321"; DORADO2="B89347"; AZUL="2962FF"; TEAL="0F8C82"
CONTROL="E8591C"; SOCIO="17A06A"; CNED="1B355B"; GRIS="55615B"; ROJO="C0392B"; PUR="6B4FA0"; NEGRO="202B26"

# ====================================================================
#   Primitivas OOXML
# ====================================================================
def run(t,b=False,i=False,sz=22,color=NEGRO,font="Calibri"):
    rpr=f'<w:rPr>{"<w:b/>" if b else ""}{"<w:i/>" if i else ""}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:color w:val="{color}"/><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/></w:rPr>'
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xe(t)}</w:t></w:r>'

def para(runs,align="both",after=120,before=0,ind=0,line=276):
    if isinstance(runs,str): runs=[run(runs)]
    ppr=f'<w:pPr><w:jc w:val="{align}"/>'
    if ind: ppr+=f'<w:ind w:left="{ind}"/>'
    ppr+=f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/></w:pPr>'
    return f'<w:p>{ppr}{"".join(runs)}</w:p>'

def h1(t):
    return (f'<w:p><w:pPr><w:spacing w:before="280" w:after="120"/>'
            f'<w:pBdr><w:bottom w:val="single" w:sz="18" w:space="2" w:color="{DORADO2}"/></w:pBdr></w:pPr>'
            f'{run(t,b=True,sz=30,color=VERDE)}</w:p>')

def h2(t):
    return para([run(t,b=True,sz=24,color=TEAL)],align="left",before=200,after=80)

def bullet(t,color=NEGRO):
    return para([run("•  ",b=True,sz=22,color=DORADO2),run(t,sz=22,color=color)],align="left",after=60,ind=340)

def pagebreak():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def spacer(after=120): return para([run("")],after=after)

# ---- tabla ----
def cell(txt_runs, w, fill=None, align="left", vmerge=None, bold=False):
    shd=f'<w:shd w:val="clear" w:fill="{fill}"/>' if fill else ''
    tcpr=f'<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}<w:vAlign w:val="center"/></w:tcPr>'
    if isinstance(txt_runs,str):
        txt_runs=[run(txt_runs,b=bold,sz=20,color=(NEGRO))]
    body=f'<w:p><w:pPr><w:jc w:val="{align}"/><w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>{"".join(txt_runs)}</w:p>'
    return f'<w:tc>{tcpr}{body}</w:tc>'

def table(rows, widths, header_fill=VERDE):
    # rows: lista de filas; cada fila: lista de (texto|runs, {opts})
    tblW=sum(widths)
    grid="".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    out=(f'<w:tbl><w:tblPr><w:tblW w:w="{tblW}" w:type="dxa"/>'
         f'<w:tblBorders>'
         f'<w:top w:val="single" w:sz="4" w:color="D9DED9"/><w:left w:val="single" w:sz="4" w:color="D9DED9"/>'
         f'<w:bottom w:val="single" w:sz="4" w:color="D9DED9"/><w:right w:val="single" w:sz="4" w:color="D9DED9"/>'
         f'<w:insideH w:val="single" w:sz="4" w:color="D9DED9"/><w:insideV w:val="single" w:sz="4" w:color="D9DED9"/>'
         f'</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>')
    for ri,rowcells in enumerate(rows):
        rpr='<w:trPr><w:tblHeader/></w:trPr>' if ri==0 else ''
        tcs=""
        for ci,(content,opts) in enumerate(rowcells):
            o=opts or {}
            fill=o.get('fill', header_fill if ri==0 else None)
            color=o.get('color', "FFFFFF" if ri==0 else NEGRO)
            bold=o.get('bold', ri==0)
            align=o.get('align','left')
            runs=content if isinstance(content,list) else [run(content,b=bold,sz=20,color=color)]
            tcs+=cell(runs, widths[ci], fill=fill, align=align)
        out+=f'<w:tr>{rpr}{tcs}</w:tr>'
    out+='</w:tbl>'
    return out

# ---- gráfico de barras horizontal con celdas (nativo, sin imagen) ----
def barchart(data, maxw=5600, color=TEAL, unit=""):
    # data: [(label, value), ...]
    mx=max(v for _,v in data) or 1
    rows=[]
    for label,v in data:
        barw=int(maxw*v/mx) or 60
        bar_cell=(f'<w:tc><w:tcPr><w:tcW w:w="{barw}" w:type="dxa"/>'
                  f'<w:shd w:val="clear" w:fill="{color}"/></w:tcPr>'
                  f'<w:p><w:pPr><w:spacing w:before="10" w:after="10"/></w:pPr>'
                  f'{run(" ",sz=14,color=color)}</w:p></w:tc>')
        rest=maxw-barw
        rest_cell=(f'<w:tc><w:tcPr><w:tcW w:w="{rest if rest>0 else 1}" w:type="dxa"/></w:tcPr>'
                   f'<w:p><w:pPr><w:spacing w:before="10" w:after="10"/></w:pPr>{run(" "+str(v)+unit,sz=18,color=GRIS)}</w:p></w:tc>')
        lbl_cell=(f'<w:tc><w:tcPr><w:tcW w:w="2600" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
                  f'<w:p><w:pPr><w:spacing w:before="10" w:after="10"/></w:pPr>{run(label,sz=18,color=NEGRO)}</w:p></w:tc>')
        rows.append(f'<w:tr>{lbl_cell}{bar_cell}{rest_cell}</w:tr>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{2600+maxw}" w:type="dxa"/>'
            f'<w:tblBorders><w:none/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="2600"/><w:gridCol w:w="{maxw}"/></w:tblGrid>'
            + "".join(rows) + '</w:tbl>')

def anexo_titulo(letra, titulo):
    return (f'<w:p><w:pPr><w:spacing w:before="120" w:after="40"/>'
            f'<w:pBdr><w:bottom w:val="single" w:sz="18" w:space="2" w:color="{DORADO2}"/></w:pBdr></w:pPr>'
            f'{run("ANEXO "+letra+" · ",b=True,sz=28,color=DORADO)}{run(titulo,b=True,sz=28,color=VERDE)}</w:p>')

def nota(t):
    return para([run(t,i=True,sz=18,color=GRIS)],align="left",before=80,after=120)

# pie de figura / tabla (caption): explica qué representa la imagen o tabla
_fign=[0]
def caption(tipo, texto):
    _fign[0]+=1
    return para([run(f"{tipo} {_fign[0]}. ",b=True,sz=18,color=DORADO),
                 run(texto, i=True, sz=18, color=GRIS)],align="left",before=40,after=160)

# ====================================================================
#   GRÁFICOS NATIVOS DE WORD (objetos chart editables)
#   Cada chart => parte /word/charts/chartN.xml + rel + <w:drawing> inline
# ====================================================================
EMU=914400
_charts=[]  # [(filename, xml)]
_chart_rid=[100]  # rIds de charts en document.xml.rels (desde 100 para no chocar)

def _bar_chart_xml(cats, vals, colors, titulo):
    pts="".join(f'<c:pt idx="{i}"><c:v>{xe(c)}</c:v></c:pt>' for i,c in enumerate(cats))
    vpts="".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
    dpts="".join(f'<c:dPt><c:idx val="{i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>'
                 f'<c:spPr><a:solidFill><a:srgbClr val="{colors[i%len(colors)]}"/></a:solidFill></c:spPr></c:dPt>'
                 for i in range(len(cats)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<c:chart>'
      f'<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="es-CO" sz="1100" b="1"/><a:t>{xe(titulo)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
      '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>'
      '<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="1"/>'
      '<c:ser><c:idx val="0"/><c:order val="0"/>'
      '<c:tx><c:strRef><c:f>Hoja1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Personas</c:v></c:pt></c:strCache></c:strRef></c:tx>'
      f'{dpts}'
      '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
      f'<c:cat><c:strRef><c:f>Hoja1!$A$2:$A${len(cats)+1}</c:f><c:strCache><c:ptCount val="{len(cats)}"/>{pts}</c:strCache></c:strRef></c:cat>'
      f'<c:val><c:numRef><c:f>Hoja1!$B$2:$B${len(vals)+1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>{vpts}</c:numCache></c:numRef></c:val>'
      '</c:ser><c:axId val="111"/><c:axId val="222"/></c:barChart>'
      '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="222"/></c:catAx>'
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="111"/></c:valAx>'
      '</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')

def _doughnut_chart_xml(cats, vals, colors, titulo):
    pts="".join(f'<c:pt idx="{i}"><c:v>{xe(c)}</c:v></c:pt>' for i,c in enumerate(cats))
    vpts="".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
    dpts="".join(f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/>'
                 f'<c:spPr><a:solidFill><a:srgbClr val="{colors[i%len(colors)]}"/></a:solidFill></c:spPr></c:dPt>'
                 for i in range(len(cats)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      '<c:chart>'
      f'<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="es-CO" sz="1100" b="1"/><a:t>{xe(titulo)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
      '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>'
      '<c:doughnutChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>'
      f'{dpts}'
      '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
      f'<c:cat><c:strRef><c:f>Hoja1!$A$2:$A${len(cats)+1}</c:f><c:strCache><c:ptCount val="{len(cats)}"/>{pts}</c:strCache></c:strRef></c:cat>'
      f'<c:val><c:numRef><c:f>Hoja1!$B$2:$B${len(vals)+1}</c:f><c:numCache><c:ptCount val="{len(vals)}"/>{vpts}</c:numCache></c:numRef></c:val>'
      '</c:ser><c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>'
      '</c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')

_hyperlinks=[]  # [(rid, url)]
_hl_n=[0]
def hyperlink(url, text):
    _hl_n[0]+=1
    rid=f"rIdLink{_hl_n[0]}"
    _hyperlinks.append((rid,url))
    r=(f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="{AZUL}"/><w:u w:val="single"/>'
       f'<w:sz w:val="22"/><w:szCs w:val="22"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>'
       f'<w:t xml:space="preserve">{xe(text)}</w:t></w:r>')
    return f'<w:hyperlink r:id="{rid}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{r}</w:hyperlink>'

def chart(kind, cats, vals, colors, titulo, w_in=6.0, h_in=3.1):
    """Devuelve un <w:p> con un objeto chart de Word inline (editable)."""
    idx=len(_charts)+1
    fname=f"charts/chart{idx}.xml"
    xml=_bar_chart_xml(cats,vals,colors,titulo) if kind=="bar" else _doughnut_chart_xml(cats,vals,colors,titulo)
    _charts.append((fname,xml))
    rid=f"rIdChart{idx}"
    cx=int(w_in*EMU); cy=int(h_in*EMU)
    docpr_id=1000+idx
    drawing=(f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
      f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
      f'<wp:docPr id="{docpr_id}" name="Grafico{idx}"/>'
      '<wp:cNvGraphicFramePr/>'
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
      f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/>'
      '</a:graphicData></a:graphic></wp:inline></w:drawing>')
    return f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr><w:r>{drawing}</w:r></w:p>'

# ====================================================================
#   CONTENIDO
# ====================================================================
body=[]

# ---------- PORTADA ----------
body.append(para([run("UNIVERSIDAD NACIONAL DE COLOMBIA",b=True,sz=22,color=GRIS)],align="center",before=600,after=40))
body.append(para([run("Dirección Nacional de Estrategia Digital (DNED) · División Universidad Laboratorio (DIVULAB)",sz=20,color=GRIS)],align="center",after=400))
body.append(para([run("Modelo de Gobernanza de la Inteligencia Artificial",b=True,sz=48,color=VERDE)],align="center",after=80))
body.append(para([run("Hacia un ecosistema para el uso responsable y la apropiación ética de la IA",b=True,sz=30,color=TEAL)],align="center",after=400))
body.append(para([run("Documento institucional · avance en la formulación y elaboración de lineamientos y guías",sz=22,color=NEGRO)],align="center",after=80))
body.append(para([run("Documento principal y anexos del modelo",i=True,sz=20,color=GRIS)],align="center",after=520))
# ---- Elaboró ----
body.append(para([run("Elaboró",b=True,sz=22,color=DORADO)],align="center",after=40))
body.append(para([run("Ing. Wilver Alexander Martínez",b=True,sz=24,color=VERDE)],align="center",after=20))
body.append(para([run("Ingeniero de Sistemas · Especialista en Gobierno Electrónico",sz=20,color=NEGRO)],align="center",after=20))
body.append(para([run("Estudiante de Maestría en Ingeniería Industrial",sz=20,color=NEGRO)],align="center",after=20))
body.append(para([run("Dirección Nacional de Estrategia Digital",sz=20,color=GRIS)],align="center",after=560))
body.append(para([run("Bogotá D.C. · "+time.strftime("%Y"),sz=20,color=GRIS)],align="center"))
body.append(pagebreak())

# ---------- 1. PRESENTACIÓN ----------
body.append(h1("1. Presentación"))
body.append(para("La Universidad Nacional de Colombia adelanta un esfuerzo institucional de alcance nacional para definir un "
  "ecosistema de gobernanza que oriente el uso responsable y la apropiación ética de la inteligencia artificial (IA) en "
  "toda su comunidad. Este documento presenta el modelo de gobernanza y el trabajo adelantado en la formulación y "
  "elaboración de los lineamientos y guías, resaltando el avance institucional alcanzado y el rol articulador de la "
  "Dirección Nacional de Estrategia Digital (DNED), a través de la División Universidad Laboratorio (DIVULAB)."))
body.append(para("Más que una estructura administrativa, se trata de un proceso colectivo y multiestamentario que articula "
  "lineamientos, estructuras de decisión, procesos y capacidades, para que la IA se despliegue al servicio de las misiones "
  "de la Universidad y de la protección de los derechos de las personas. El documento se organiza en un cuerpo principal "
  "—que sintetiza el avance y el rol de la DNED— y un conjunto de anexos, uno por cada componente del modelo."))

# ---------- 2. AVANCE INSTITUCIONAL (resaltado) ----------
body.append(h1("2. Avance institucional alcanzado"))
body.append(para("El proceso ha logrado un avance significativo y verificable, soportado en la participación activa de la "
  "comunidad universitaria y en la articulación de múltiples esfuerzos bajo la coordinación de la DNED–DIVULAB:"))
body.append(table(
  [[("Indicador",{}),("Valor",{"align":"center"}),("Descripción",{})],
   [("Mesas de trabajo",{}),("14",{"align":"center"}),("Mesas temáticas que estructuran los lineamientos y guías",{})],
   [("Documentos entregados",{}),("11",{"align":"center"}),("Documentos finales entregados (79% de avance)",{})],
   [("Documentos pendientes",{}),("3",{"align":"center"}),("En proceso de consolidación",{})],
   [("Personas participantes",{}),("129",{"align":"center"}),("108 participantes + 21 facilitadores (DNED)",{})],
   [("Sedes integradas",{}),("9",{"align":"center"}),("Alcance multicampus y territorial",{})],
   [("Recomendaciones académicas",{}),("10",{"align":"center"}),("Comentarios externos incorporados al modelo",{})]],
  widths=[3000,1400,5000]))
body.append(spacer())
body.append(h2("Avance de documentos por mesa"))
body.append(chart("dough", ["Entregados","Pendientes"], [11,3], [SOCIO,CONTROL],
                  "Documentos: 11 entregados · 3 pendientes", w_in=5.0, h_in=2.8))
body.append(caption("Gráfico","Avance documental (objeto de gráfico editable de Word). Cada sector de la dona representa una "
  "categoría: verde = documentos entregados (11); naranja = pendientes (3), sobre un total de 14 mesas de trabajo."))

# ---------- 3. ROL DE LA DNED ----------
body.append(h1("3. El rol articulador de la DNED–DIVULAB"))
body.append(para("La Dirección Nacional de Estrategia Digital orienta y articula la estrategia digital de la Universidad en sus "
  "nueve sedes. A través de la División Universidad Laboratorio (DIVULAB), lidera la co-construcción multiestamentaria del "
  "marco de uso ético y responsable de la IA, en articulación con la Mesa Técnica de IA y la comunidad universitaria. Su "
  "función es integrar y dar coherencia a esfuerzos que, de otro modo, estarían dispersos."))
body.append(h2("Origen del trabajo: estado del arte y marco conceptual"))
body.append(para([run("Este proceso fue iniciado por el ingeniero ",sz=22),
  run("Wilver Alexander Martínez",b=True,sz=22,color=VERDE),
  run(", líder en la DNED de la iniciativa institucional de Gobernanza de la IA, quien formuló el ",sz=22),
  run("estado del arte",b=True,sz=22),
  run(" y lo sintetizó en el ",sz=22),
  run("documento de marco conceptual",b=True,sz=22),
  run(". Este marco permitió delinear el alcance del uso de la IA en la Universidad y los principios de ",sz=22),
  run("uso responsable",b=True,sz=22),
  run(" y ",sz=22),
  run("apropiación ética",b=True,sz=22),
  run(" de la IA, sentando las bases conceptuales sobre las que se construyeron las mesas de trabajo, los lineamientos y el "
      "presente modelo de gobernanza.",sz=22)]))
body.append(h2("Funciones de articulación"))
for t in ["Coordinar las 14 mesas de trabajo y consolidar sus productos en un marco común.",
          "Facilitar la participación (108 participantes + 21 facilitadores) y el diálogo entre estamentos y sedes.",
          "Conducir el benchmarking internacional y la validación académica externa del modelo.",
          "Articular los componentes del modelo (estructura, operación, riesgo, participación, hoja de ruta).",
          "Integrar la sostenibilidad (Green AI) y la soberanía del conocimiento como ejes transversales."]:
    body.append(bullet(t))
body.append(h2("Ejes rectores del ecosistema"))
body.append(table(
  [[("Eje",{}),("Qué significa",{})],
   [("Participativo y multiestamentario",{}),("La comunidad como co-creadora, no destinataria pasiva",{})],
   [("Multinivel y multicampus",{}),("CNED → Mesa Técnica → Comunidad; 9 sedes; subsidiariedad",{})],
   [("Centrado en derechos",{}),("El riesgo y el veto se gradúan por afectación a derechos fundamentales",{})],
   [("Sostenible",{}),("Green AI: eficiencia y huella de carbono con marcos reconocidos",{})],
   [("Soberanía del conocimiento",{}),("Equidad territorial y diálogo intercultural",{})]],
  widths=[3400,6000]))

# ---------- 4. SÍNTESIS + ÍNDICE DE ANEXOS ----------
body.append(h1("4. Síntesis del modelo e índice de anexos"))
body.append(para("El modelo de gobernanza se compone de los siguientes elementos, desarrollados en los anexos. Cada anexo "
  "ocupa un máximo de dos páginas e incluye tablas o gráficos cuando aportan a la comprensión."))
body.append(table(
  [[("Anexo",{"align":"center"}),("Componente del modelo",{})],
   [("A",{"align":"center"}),("Benchmarking internacional (Norte y Sur Global)",{})],
   [("B",{"align":"center"}),("Estructura multinivel y subsidiariedad",{})],
   [("C",{"align":"center"}),("Los cuatro núcleos funcionales",{})],
   [("D",{"align":"center"}),("Operación: plenaria de delegados, quórum y votación",{})],
   [("E",{"align":"center"}),("Gestión de riesgo y taxonomía de sistemas de IA",{})],
   [("F",{"align":"center"}),("Matriz RACI de roles de decisión",{})],
   [("G",{"align":"center"}),("Participación: 129 personas por mesa y por rol",{})],
   [("H",{"align":"center"}),("Recomendaciones académicas incorporadas",{})],
   [("I",{"align":"center"}),("Hoja de ruta 2025–2027",{})],
   [("J",{"align":"center"}),("Listado de guías y documentos de lineamientos de IA",{})]],
  widths=[1200,8200]))

# ---------- 5. REFERENCIAS Y DOCUMENTACIÓN ----------
body.append(h1("5. Referencias y documentación"))
body.append(para("El presente documento se sustenta en el trabajo institucional adelantado por la DNED–DIVULAB y en las "
  "fuentes que se relacionan a continuación. La documentación completa del proceso —marco conceptual, actas y productos de "
  "las mesas de trabajo, guías y documentos de lineamientos— se encuentra desplegada en el repositorio institucional:"))
body.append(para([run("Repositorio de documentación: ",b=True,sz=22,color=NEGRO),
  hyperlink("https://drive.google.com/drive/folders/1zZ3nCgpte0JUMN6NLxNNcKHmAxAID6Jj",
            "https://drive.google.com/drive/folders/1zZ3nCgpte0JUMN6NLxNNcKHmAxAID6Jj")],
  align="left"))
body.append(h2("Fuentes"))
refs=[
  "Documento de marco conceptual, normativo y regulatorio para el uso ético de la IA (estado del arte). Ing. Wilver Alexander Martínez, DNED.",
  "Propuesta de creación de la Mesa Técnica de Inteligencia Artificial como sala adscrita al CNED. DNED.",
  "Presentación institucional «Gobernanza IA». DNED–DIVULAB.",
  "Guías y documentos de lineamientos de las 14 mesas de trabajo (ver Anexo J).",
  "Comentarios de revisión académica externa incorporados al modelo (ver Anexo H).",
  "Benchmarking internacional: MIT Schwarzman College of Computing; UCLA / UC AI Council; Stanford HAI / AI Index; Tec de Monterrey; USP; UNAM; proyecto INTEGRITY-AI (ver Anexo A).",
]
for r in refs:
    body.append(bullet(r))
body.append(nota("La versión vigente y completa de cada documento está disponible en el repositorio institucional "
  "indicado. Para acceso, consultar con la Dirección Nacional de Estrategia Digital (DNED)."))
body.append(pagebreak())

# ====================================================================
#   ANEXOS (máx 2 páginas c/u)
# ====================================================================

# ---- ANEXO A: Benchmarking ----
body.append(anexo_titulo("A","Benchmarking internacional de gobernanza de IA"))
body.append(para("El diseño del modelo partió de un ejercicio de referenciación (benchmarking) de la forma en que universidades "
  "de alto nivel gobiernan la inteligencia artificial. Se analizaron tres referentes del Norte Global —el MIT, con un "
  "esquema descentralizado por núcleos de especialización; la Universidad de California (UCLA), con un modelo "
  "matricial-centralizado y control clínico; y Stanford, con un enfoque hub-and-spoke articulado por su instituto HAI—, "
  "cada uno con una lógica distinta de decisión y una dimensión ambiental incipiente."))
body.append(para("Dado que esos referentes operan en contextos y puntos de partida diferentes al colombiano, el ejercicio se "
  "amplió deliberadamente con referentes del Sur Global, más cercanos a la realidad de la UNAL: el Tec de Monterrey "
  "(multicampus), la Universidad de São Paulo (regulación en una gran IES pública), la UNAM (proceso en construcción) y "
  "el proyecto INTEGRITY-AI, red latinoamericana con pares colombianos (U. del Cauca y U. de San Buenaventura). La tabla "
  "siguiente sintetiza y compara estos referentes."))
body.append(table(
  [[("Institución",{}),("Modelo",{}),("Órganos / rasgos",{}),("Carácter",{})],
   [("MIT",{"color":NEGRO}),("Descentralizado",{}),("SERC; Ad Hoc Committee; programa semilla",{}),("Consultivo",{})],
   [("UCLA",{}),("Matricial-centralizado",{}),("UC AI Council; HAIC (salud)",{}),("Mandatorio (clínico)",{})],
   [("Stanford",{}),("Hub-and-spoke",{}),("HAI; AIMES; AI Index",{}),("Orientador",{})],
   [("Tec de Monterrey",{}),("Multicampus (LatAm)",{}),("Integridad académica con IA",{}),("Referente Sur",{})],
   [("USP (Brasil)",{}),("Normativo (LatAm)",{}),("Mayor IES pública de la región",{}),("Referente Sur",{})],
   [("UNAM (México)",{}),("En construcción",{}),("Universidad pública nacional",{}),("Referente Sur",{})],
   [("INTEGRITY-AI",{}),("Red colombiana/LatAm",{}),("U. del Cauca; U. San Buenaventura",{}),("Red de pares",{})]],
  widths=[2200,2400,3200,1600]))
body.append(caption("Tabla","Comparación de referentes. Cada fila es una institución; las columnas indican su modelo de "
  "gobernanza (cómo se distribuye la decisión), sus órganos o rasgos distintivos y el carácter de sus decisiones "
  "(consultivo, mandatorio u orientador). Las tres primeras filas son referentes del Norte Global y las cuatro últimas, "
  "del Sur Global."))
body.append(nota("La UNAL no adopta un modelo externo tal cual: construye uno propio, público, nacional, pluriétnico y multicampus."))
body.append(pagebreak())

# ---- ANEXO B: Estructura multinivel ----
body.append(anexo_titulo("B","Estructura multinivel y subsidiariedad"))
body.append(para("La gobernanza de la IA en la UNAL se organiza en tres niveles articulados que conectan la decisión política "
  "de alto nivel con la operación cotidiana en las sedes. En la cúspide, el Comité Nacional de Estrategia Digital (CNED) "
  "define la política y aprueba la normativa; en el nivel intermedio, la Mesa Técnica de IA ejecuta, evalúa y audita a "
  "través de sus cuatro núcleos; y en la base, la comunidad universitaria —nueve sedes, facultades, grupos y laboratorios "
  "vivos— es, a la vez, destinataria y co-creadora del modelo."))
body.append(para("Este esquema evita tanto la fragmentación de un modelo totalmente descentralizado como la rigidez de uno "
  "excesivamente centralizado. La primera tabla describe los tres niveles y su rol; la segunda operacionaliza el principio "
  "de subsidiariedad, indicando qué tipo de decisiones corresponde a cada ámbito."))
body.append(table(
  [[("Nivel",{}),("Instancia",{}),("Rol",{})],
   [("1 · Alto nivel",{}),("CNED",{}),("Definición política de alto nivel y aprobación normativa",{})],
   [("2 · Ejecución",{}),("Mesa Técnica de IA",{}),("Ejecución, evaluación, auditoría; 4 núcleos; PLASI; ciclo trienal",{})],
   [("3 · Base",{}),("Comunidad Universitaria",{}),("9 sedes, facultades, grupos, laboratorios vivos",{})]],
  widths=[2200,3000,4200]))
body.append(caption("Tabla","Los tres niveles del modelo. La columna 'Nivel' va de la decisión política (arriba) a la base "
  "comunitaria (abajo); 'Instancia' nombra el órgano responsable y 'Rol' resume su función principal."))
body.append(h2("Principio de subsidiariedad"))
body.append(para("La decisión se toma en el nivel más cercano al lugar donde se genera el riesgo y el valor; se centraliza solo "
  "por excepción: cuando afecta a toda la Universidad, compromete información sensible, implica riesgos elevados o requiere "
  "estándares institucionales comunes. Así, lo cotidiano y de bajo riesgo se resuelve en la facultad o la sede, y solo "
  "escala cuando el impacto lo justifica."))
body.append(table(
  [[("Ámbito de decisión",{}),("Qué decide",{})],
   [("Nacional (CNED/Mesa)",{}),("Política y normativa; alto riesgo; datos sensibles; PLASI",{})],
   [("Sede",{}),("Capacidades e infraestructura locales; riesgos territoriales",{})],
   [("Facultad / Grupo",{}),("IA en aula, investigación y extensión; bajo riesgo local",{})],
   [("Compartido",{}),("Riesgo medio multi-sede; herramientas transversales; Habeas Data",{})]],
  widths=[3000,6400]))
body.append(caption("Tabla","Reparto de decisiones por subsidiariedad. Cada fila es un ámbito de decisión y la columna de la "
  "derecha enumera los asuntos que le corresponden según su cercanía al riesgo y al valor."))
body.append(pagebreak())

# ---- ANEXO C: 4 núcleos ----
body.append(anexo_titulo("C","Los cuatro núcleos funcionales de la Mesa Técnica"))
body.append(para("La Mesa Técnica de IA no es un cuerpo homogéneo: se estructura en cuatro núcleos funcionales que, en conjunto, "
  "cubren de forma multidimensional la gobernanza de la IA. El Núcleo Estratégico Misional vela por la coherencia con las "
  "misiones de la Universidad; el Técnico & Green Computing atiende la infraestructura, la seguridad y la sostenibilidad; "
  "el de Control & Jurídico-Fiscal cuida la legalidad, la protección de datos y la viabilidad presupuestal; y el "
  "Socio-Ambiental & Territorial aporta la mirada ética, de derechos y de saberes territoriales."))
body.append(para("Esta composición busca que ninguna dimensión quede desatendida y que las tensiones (por ejemplo, entre "
  "viabilidad técnica y protección de derechos) se resuelvan de forma deliberada y colegiada. La tabla relaciona cada "
  "núcleo con su composición típica y su enfoque."))
body.append(table(
  [[("Núcleo",{}),("Composición (ejemplos)",{}),("Enfoque",{})],
   [("1.1 Estratégico Misional",{"color":"FFFFFF","fill":AZUL}),("Vicerrectorías; Dir. Extensión y P.I.",{}),("Coherencia misional, soberanía epistémica",{})],
   [("1.2 Técnico & Green Computing",{"color":"FFFFFF","fill":TEAL}),("DNED; ciberseguridad; HPC; TI de sedes",{}),("Infraestructura, DT 1-25, eficiencia energética",{})],
   [("1.3 Control & Jurídico-Fiscal",{"color":"FFFFFF","fill":CONTROL}),("Planeación; Jurídica; Control Interno",{}),("Legalidad, RAEE, Habeas Data, software libre",{})],
   [("1.4 Socio-Ambiental & Territorial",{"color":"FFFFFF","fill":SOCIO}),("Investigadores; sedes; estudiantes",{}),("Ética, saberes territoriales, salud mental",{})]],
  widths=[3000,3400,3000]))
body.append(caption("Tabla","Los cuatro núcleos. El color de la primera columna identifica a cada núcleo (azul = estratégico, "
  "verde azulado = técnico, naranja = control/jurídico, verde = socio-ambiental); 'Composición' da ejemplos de sus "
  "integrantes y 'Enfoque' resume su ámbito de responsabilidad."))
body.append(h2("Funciones específicas y decisiones típicas"))
body.append(table(
  [[("Núcleo",{}),("Funciones específicas",{}),("Pregunta que decide",{})],
   [("1.1 Estratégico",{"color":"FFFFFF","fill":AZUL}),("Priorizar por aporte misional; velar por pertinencia y soberanía epistémica.",{}),("¿Aporta a las misiones y al proyecto educativo?",{})],
   [("1.2 Técnico & Green",{"color":"FFFFFF","fill":TEAL}),("Evaluar infraestructura y seguridad; aplicar DT 1-25; medir huella energética.",{}),("¿Hay cómputo y seguridad? ¿Cumple las DT?",{})],
   [("1.3 Control & Jurídico",{"color":"FFFFFF","fill":CONTROL}),("Verificar legalidad y licencias; garantizar Habeas Data; viabilidad presupuestal.",{}),("¿Es legal, viable y protege los datos?",{})],
   [("1.4 Socio-Ambiental",{"color":"FFFFFF","fill":SOCIO}),("Evaluar afectación a derechos y a poblaciones; saberes territoriales; bienestar.",{}),("¿Afecta derechos? ¿Debe objetarse o escalarse?",{})]],
  widths=[2400,4400,2600]))
body.append(caption("Tabla","Profundización de los núcleos. Para cada núcleo se detallan sus funciones específicas y la "
  "pregunta central que resuelve en la evaluación de un caso. Los núcleos 1.3 y 1.4 pueden objetar por riesgo a derechos."))
body.append(nota("Los núcleos interactúan: el Estratégico deriva a Técnico y Control (que se consultan entre sí) y el "
  "Socio-Ambiental actúa como control ético transversal con capacidad de objeción."))
body.append(pagebreak())

# ---- ANEXO D: Operación / plenaria ----
body.append(anexo_titulo("D","Operación: plenaria de delegados, quórum y votación"))
body.append(para("La Mesa Técnica opera como un órgano colegiado y deliberante, no como una ventanilla técnica. Su instancia "
  "central es la Plenaria de Delegados, integrada por un delegado de cada uno de los cuatro núcleos. La plenaria recibe y "
  "radica los casos (iniciativas o sistemas de IA), los traslada a trámite a los núcleos para su evaluación especializada, "
  "y luego delibera con las ponencias que cada delegado trae de regreso."))
body.append(para("Este diseño hace trazable 'quién decide qué' y evita que una sola área concentre la decisión. La votación se "
  "rige por reglas explícitas: existe un quórum mínimo para sesionar y una regla de prevalencia que protege los derechos "
  "fundamentales. A continuación se describe el procedimiento paso a paso y, luego, la regla de decisión."))
body.append(h2("Procedimiento"))
for t in ["Radicación del caso en la Plenaria de Delegados y análisis preliminar.",
          "Traslado a trámite: cada delegado lleva el caso a su núcleo para evaluación especializada.",
          "Consulta entre núcleos (1.2 Técnico ↔ 1.3 Jurídico: seguridad/infra ↔ datos/licencias).",
          "Regreso con ponencia de cada núcleo y deliberación en plenaria.",
          "Decisión colegiada por quórum (≥ 3 de 4 delegados)."]:
    body.append(bullet(t))
body.append(h2("Regla de decisión y votación"))
body.append(table(
  [[("Condición",{}),("Resultado",{})],
   [("Objeción de 1.3 o 1.4 por riesgo a derechos",{}),("Prevalece: veto / escalar al CNED",{})],
   [("Mayoría con salvaguardas",{}),("Favorable con condiciones (nivel sede)",{})],
   [("Unanimidad a favor",{}),("Favorable",{})],
   [("Sin quórum (< 3 de 4)",{}),("Sesión aplazada",{})]],
  widths=[5200,4200]))
body.append(caption("Tabla","Regla de decisión de la plenaria. Cada fila cruza una condición de la votación (izquierda) con "
  "su resultado (derecha). La primera fila expresa la regla de prevalencia: una objeción por riesgo a derechos pesa más "
  "que la mayoría y obliga a vetar o escalar al CNED."))
body.append(nota("Ejemplo: un predictor de deserción (datos sensibles) genera objeción del núcleo jurídico y escala al CNED."))
body.append(pagebreak())

# ---- ANEXO E: Riesgo y taxonomía ----
body.append(anexo_titulo("E","Gestión de riesgo y taxonomía de sistemas de IA"))
body.append(para("Un principio central del modelo es que no todas las iniciativas de IA son iguales ni representan el mismo "
  "riesgo. Por ello, la gobernanza se gradúa en función del nivel de riesgo, entendido como la posible afectación a los "
  "derechos fundamentales de las personas. Un sistema que decide sobre becas o predice la deserción puede comprometer "
  "derechos tanto como uno clínico, aunque sea 'académico'; de ahí que el control se defina por el riesgo y no por el "
  "dominio de aplicación."))
body.append(para("La gestión combina dos lentes complementarios: por un lado, cuatro niveles de riesgo —de bajo a inaceptable— "
  "cada uno con un tratamiento y una autoridad de aprobación distintos; por otro, una taxonomía que distingue los tipos de "
  "sistemas de IA según su propósito y origen. La primera tabla presenta los niveles y la lista posterior, la taxonomía."))
body.append(h2("Niveles de riesgo"))
body.append(table(
  [[("Nivel",{}),("Tratamiento",{})],
   [("Inaceptable",{"color":"FFFFFF","fill":"7A1414"}),("Prohibido: vulnera derechos fundamentales",{})],
   [("Alto",{"color":"FFFFFF","fill":ROJO}),("Autorización CNED + auditoría previa/posterior + veto",{})],
   [("Medio",{"color":"FFFFFF","fill":CONTROL}),("Autorización de sede + supervisión de la Mesa",{})],
   [("Bajo",{"color":"FFFFFF","fill":SOCIO}),("Autorización local (facultad) + registro",{})]],
  widths=[2400,7000]))
body.append(caption("Tabla","Niveles de riesgo. El color va de rojo oscuro (inaceptable) a verde (bajo); a mayor riesgo, "
  "más alta la instancia que autoriza y más estrictos los controles. La columna derecha indica el tratamiento exigido en "
  "cada nivel."))
body.append(h2("Taxonomía por propósito y origen"))
for t in ["Decisión académico-administrativa (deserción, becas, evaluación) — riesgo típico alto.",
          "IA en el aula (enseñanza-aprendizaje) — riesgo bajo-medio.",
          "IA en investigación y extensión — riesgo medio.",
          "IA desarrollada en la UNAL (modelos propios) — riesgo variable.",
          "IA generativa de terceros — riesgo medio-alto (dependencia, fuga de datos).",
          "IA agéntica (autónoma) — riesgo alto (emergente)."]:
    body.append(bullet(t))
body.append(nota("El veto se gradúa por riesgo a derechos, no por el dominio de aplicación."))
body.append(pagebreak())

# ---- ANEXO F: RACI ----
body.append(anexo_titulo("F","Matriz RACI de roles de decisión"))
body.append(para("Para que la gobernanza sea operativa no basta con definir quién participa: hay que precisar quién decide "
  "qué en cada etapa del ciclo de vida de una iniciativa de IA. La matriz RACI cumple esa función, asignando a cada actor "
  "(Facultad, Sede, Mesa Técnica, CNED y la instancia Jurídica/Control) un papel específico en cada actividad, según el "
  "riesgo y la misión involucrados."))
body.append(para("La matriz emplea cinco papeles: R (Responsable, propone o implementa), A (Autoriza y rinde cuentas), "
  "C (Consultado o evalúa), I (Informado) y S (Supervisa y puede suspender). Es una propuesta inicial para la discusión "
  "institucional e incorpora la recomendación de graduar el veto por el riesgo a los derechos, visible en las filas de "
  "autorización de alto riesgo y de suspensión."))
body.append(table(
  [[("Actividad",{}),("Facultad",{"align":"center"}),("Sede",{"align":"center"}),("Mesa Téc.",{"align":"center"}),("CNED",{"align":"center"}),("Jurídico",{"align":"center"})],
   [("Proponer",{}),("R",{"align":"center"}),("C",{"align":"center"}),("I",{"align":"center"}),("—",{"align":"center"}),("—",{"align":"center"})],
   [("Evaluar riesgo",{}),("C",{"align":"center"}),("C",{"align":"center"}),("R",{"align":"center"}),("I",{"align":"center"}),("C",{"align":"center"})],
   [("Autorizar (alto)",{}),("I",{"align":"center"}),("C",{"align":"center"}),("C",{"align":"center"}),("A",{"align":"center"}),("C",{"align":"center"})],
   [("Implementar",{}),("R",{"align":"center"}),("R",{"align":"center"}),("C",{"align":"center"}),("—",{"align":"center"}),("—",{"align":"center"})],
   [("Supervisar",{}),("I",{"align":"center"}),("C",{"align":"center"}),("S",{"align":"center"}),("I",{"align":"center"}),("S",{"align":"center"})],
   [("Suspender (veto)",{}),("I",{"align":"center"}),("I",{"align":"center"}),("S",{"align":"center"}),("A",{"align":"center"}),("S",{"align":"center"})]],
  widths=[2800,1340,1340,1340,1340,1240]))
body.append(caption("Tabla","Matriz RACI. Cada fila es una actividad del ciclo de vida de una iniciativa; cada columna, un "
  "actor. La celda indica el papel de ese actor en esa actividad: R responsable, A autoriza, C consultado, I informado, "
  "S supervisa/suspende; '—' significa que no interviene."))
body.append(nota("Propuesta inicial para la discusión institucional (incorpora la recomendación de graduar el veto por derechos)."))
body.append(pagebreak())

# ---- ANEXO G: Participación ----
body.append(anexo_titulo("G","Participación: 129 personas (108 participantes + 21 facilitadores)"))
body.append(para("El carácter multiestamentario del proceso se refleja en sus cifras de participación: 129 personas "
  "intervinieron en la construcción del marco, distribuidas en las 14 mesas de trabajo. De ellas, 108 son participantes "
  "de la comunidad universitaria y 21 son facilitadores de la DNED, que acompañan metodológicamente cada mesa y se cuentan "
  "también como participantes."))
body.append(para("Esta amplitud —que integra docentes, investigadores, personal administrativo, estudiantes y directivos de "
  "las nueve sedes— es la que da legitimidad al modelo y asegura que los lineamientos recojan conocimiento situado. La "
  "primera tabla presenta los totales de control y el gráfico siguiente muestra cómo se distribuyen las personas por rol."))
body.append(h2("Totales"))
body.append(table(
  [[("Categoría",{}),("N.º personas",{"align":"center"})],
   [("Participantes",{}),("108",{"align":"center"})],
   [("Facilitadores (DNED)",{}),("21",{"align":"center"})],
   [("Total",{"fill":"EEF3F1","color":NEGRO,"bold":True}),("129",{"align":"center","fill":"EEF3F1","color":NEGRO,"bold":True})]],
  widths=[5200,4200]))
body.append(caption("Tabla","Totales de participación. Suma de participantes (108) y facilitadores de la DNED (21) para un "
  "total de 129 personas; la fila sombreada es el total."))
body.append(h2("Distribución por rol"))
ROLES=[("Facilitador(a) DNED",21),("Docente",34),("Investigador(a)",22),("Administrativo(a)",19),("Estudiante",15),("Directivo/Delegado",18)]
body.append(chart("bar", [r[0] for r in ROLES], [r[1] for r in ROLES],
                  [TEAL,AZUL,SOCIO,CONTROL,PUR,DORADO2],
                  "Distribución de 129 personas por rol", w_in=6.3, h_in=3.3))
body.append(caption("Gráfico","Distribución de las 129 personas por rol (objeto de gráfico editable de Word). Cada barra es "
  "un rol y su longitud es proporcional al número de personas, mostrado al final de la barra. La suma de todas las barras "
  "es 129."))
body.append(nota("Cifras de control del archivo de participación; el desglose por rol se ajusta al listado oficial."))
body.append(pagebreak())

# ---- ANEXO H: Recomendaciones ----
body.append(anexo_titulo("H","Recomendaciones académicas incorporadas"))
body.append(para("El modelo fue sometido a revisión académica externa, cuyas observaciones se incorporaron para fortalecer su "
  "pertinencia, rigor y contextualización. Entre los aportes más relevantes figuran la inclusión de referentes "
  "latinoamericanos, la graduación del veto por riesgo a derechos fundamentales (y no por dominio), la necesidad de una "
  "matriz RACI explícita y el principio de subsidiariedad."))
body.append(para("La siguiente tabla sintetiza las diez recomendaciones acogidas. Cada una se tradujo en un elemento concreto "
  "del modelo —presente en los demás anexos—, lo que evidencia que la revisión externa no fue un trámite formal sino una "
  "fuente efectiva de mejora."))
recs=[("1","Referentes del Sur Global / Latinoamérica"),
      ("2","Arquitectura de decisiones (nacional/sede/facultad/compartido)"),
      ("3","La gobernanza de IA supera al gobierno de TI (gobierno corporativo)"),
      ("4","Veto graduado por riesgo a derechos fundamentales"),
      ("5","Tipologías de sistemas de IA (no solo 'iniciativas')"),
      ("6","Definición de riesgo (proyecto de ley Colombia / marco UE adaptado)"),
      ("7","Marcos ambientales reconocidos (no indicadores propios)"),
      ("8","Matriz RACI (propone/evalúa/autoriza/implementa/supervisa/suspende)"),
      ("9","Subsidiariedad (decidir cerca del riesgo y el valor)"),
      ("10","Legitimidad institucional y capacidades habilitantes")]
body.append(table(
  [[("#",{"align":"center"}),("Recomendación incorporada",{})]]+
  [[(n,{"align":"center"}),(t,{})] for n,t in recs],
  widths=[900,8500]))
body.append(caption("Tabla","Recomendaciones acogidas. La columna '#' numera cada recomendación y la de la derecha la "
  "enuncia; todas fueron incorporadas al modelo y se desarrollan en los anexos correspondientes."))
body.append(pagebreak())

# ---- ANEXO I: Hoja de ruta ----
body.append(anexo_titulo("I","Hoja de ruta 2025–2027"))
body.append(para("La construcción del ecosistema de gobernanza no es un evento puntual sino un proceso plurianual estructurado "
  "en cinco fases que van del diagnóstico a la implementación. El plan ordena los esfuerzos en el tiempo y reparte el "
  "liderazgo: la DNED conduce las fases iniciales de diagnóstico, co-construcción y validación, mientras que la Mesa "
  "Técnica de IA toma el relevo en la revisión, aprobación e implementación."))
body.append(para("Este horizonte 2025–2027 permite que los lineamientos se validen con las sedes y facultades antes de su "
  "adopción, y que la implementación incluya evaluación y ajustes. La tabla detalla cada fase, su periodo y sus hitos "
  "principales."))
body.append(table(
  [[("Fase",{}),("Periodo",{"align":"center"}),("Hitos",{})],
   [("1. Diagnóstico",{}),("2.º sem. 2025",{"align":"center"}),("Análisis de resultados del taller inicial",{})],
   [("2. Co-construcción",{}),("Feb–Sep 2026",{"align":"center"}),("Taller focalizado (mesas de trabajo); co-creación",{})],
   [("3. Validación",{}),("Oct 2026",{"align":"center"}),("Validación externa; diálogos en sedes y facultades",{})],
   [("4. Revisión y aprobación",{}),("Nov 2026 – Feb 2027",{"align":"center"}),("Revisión legal; aprobación institucional",{})],
   [("5. Implementación",{}),("Mar–Jun 2027",{"align":"center"}),("Lanzamiento; evaluación y ajustes",{})]],
  widths=[3000,2400,4000]))
body.append(caption("Tabla","Hoja de ruta 2025–2027. Cada fila es una fase secuencial del plan; las columnas indican su "
  "periodo y sus hitos principales. Las fases 1 a 3 son lideradas por la DNED y las fases 4 y 5, por la Mesa Técnica de IA."))
body.append(nota("La DNED lidera diagnóstico–validación; la Mesa Técnica de IA conduce revisión–implementación."))
body.append(pagebreak())

# ---- ANEXO J: Listado de guías y documentos de lineamientos ----
body.append(anexo_titulo("J","Listado de guías y documentos de lineamientos de IA"))
body.append(para([run("El producto tangible del ecosistema es un cuerpo articulado de guías y documentos de lineamientos. "
  "Todo el conjunto parte del ",sz=22),
  run("documento de marco conceptual",b=True,sz=22),
  run(" (estado del arte, formulado por el ing. Wilver Alexander Martínez), que delineó el alcance del uso de la IA y los "
      "principios de uso responsable y apropiación ética, y actúa como pieza base que da coherencia a los demás.",sz=22)]))
body.append(para("Cada guía corresponde a una mesa de trabajo temática y aborda una dimensión específica (integridad "
  "académica, protección de datos, salud mental, enfoque de género, diversidad cultural, entre otras). La primera tabla "
  "destaca el documento base y la segunda lista las 14 guías con su estado de avance (entregado o pendiente), de modo que "
  "se aprecie el estado global del trabajo documental."))
# marco conceptual como pieza base
body.append(table(
  [[("Documento base",{}),("Autoría / origen",{}),("Estado",{"align":"center"})],
   [("Marco conceptual, normativo y regulatorio para el uso ético de la IA (estado del arte)",{}),
    ("Ing. Wilver Alexander Martínez — DNED",{}),("Base del proceso",{"align":"center","fill":"EEF3F1","color":NEGRO,"bold":True})]],
  widths=[5400,2600,1400]))
body.append(caption("Tabla","Documento base del conjunto. Identifica el marco conceptual (estado del arte), su autoría y su "
  "condición de pieza fundacional sobre la que se articulan las demás guías."))
body.append(spacer(80))
GUIAS=[
 ("1","Gestión del cambio y apropiación ética de la IA en la UNAL","Entregado"),
 ("2","Principios, usos y límites éticos de la IA en la UNAL","Entregado"),
 ("3","Integridad académica y citación responsable del uso de IA","Pendiente"),
 ("4","Consentimiento informado y protección de datos personales","Entregado"),
 ("5","Identificación, evaluación y gestión de riesgos éticos","Entregado"),
 ("6","Gobernanza institucional y registro de modelos (PLASI)","Pendiente"),
 ("7","Salud mental, bienestar digital y uso responsable de la IA","Entregado"),
 ("8","Aplicaciones éticas y pedagógicas de la IA por área disciplinar","Entregado"),
 ("9","Marco conceptual, normativo y regulatorio para el uso ético de la IA","Entregado"),
 ("10","IA, diversidad cultural y enfoque etnoeducativo","Pendiente"),
 ("11","Lineamientos estratégicos para la política institucional de IA","Entregado"),
 ("12","Gestión estratégica de información, datos y conocimiento","Entregado"),
 ("13","Identidad digital, confianza institucional y responsabilidad","Pendiente"),
 ("14","Brecha de género e IA","Entregado"),
]
def estado_cell(e):
    fill = SOCIO if e=="Entregado" else CONTROL
    return [run(e, b=True, sz=18, color="FFFFFF")], fill
rows=[[("#",{"align":"center"}),("Guía / documento de lineamientos",{}),("Estado",{"align":"center"})]]
for n,t,e in GUIAS:
    runs,fill=estado_cell(e)
    rows.append([(n,{"align":"center"}),(t,{}),(runs,{"align":"center","fill":fill})])
body.append(table(rows, widths=[700,7100,1600]))
body.append(caption("Tabla","Guías y documentos de lineamientos. Cada fila es una guía temática; la columna 'Estado' usa "
  "verde para 'Entregado' y naranja para 'Pendiente'. De 14 guías, 11 están entregadas y 3 en consolidación."))
body.append(nota("El documento de marco conceptual (estado del arte) es la pieza base que articula todo el conjunto."))
body.append(spacer(240))
body.append(para([run("Documento elaborado con apoyo de herramientas de IA generativa; contenido validado contra los documentos "
  "institucionales de la UNAL. DNED–DIVULAB.",i=True,sz=18,color=GRIS)],align="left"))

# ====================================================================
#   EMPAQUETADO DOCX
# ====================================================================
document=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
  + ''.join(body)
  + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'
    '</w:sectPr></w:body></w:document>')

ct=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
 '<Default Extension="xml" ContentType="application/xml"/>'
 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
 '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
 '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
 + "".join(f'<Override PartName="/word/{fn}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' for fn,_ in _charts)
 + '</Types>')
rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
 '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
now=time.strftime("%Y-%m-%dT%H:%M:%SZ",time.gmtime())
core=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
 '<dc:title>Modelo de Gobernanza de la IA en la UNAL - documento y anexos</dc:title>'
 '<dc:creator>DNED-DIVULAB - Universidad Nacional de Colombia</dc:creator>'
 f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
 f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
app=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
 '<Application>DNED-DIVULAB</Application></Properties>')
chart_rels="".join(
  f'<Relationship Id="rIdChart{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{fn}"/>'
  for i,(fn,_) in enumerate(_charts))
link_rels="".join(
  f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{xe(url)}" TargetMode="External"/>'
  for rid,url in _hyperlinks)
word_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 + chart_rels + link_rels + '</Relationships>')

with zipfile.ZipFile("documento-modelo-gobernanza-ia-unal.docx","w",zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml",ct)
    z.writestr("_rels/.rels",rels)
    z.writestr("word/document.xml",document)
    z.writestr("word/_rels/document.xml.rels",word_rels)
    z.writestr("docProps/core.xml",core)
    z.writestr("docProps/app.xml",app)
    for fn,xml in _charts:
        z.writestr(f"word/{fn}", xml)

print("DOCX OK:", os.path.getsize("documento-modelo-gobernanza-ia-unal.docx"), "bytes")
print("Parrafos:", document.count("<w:p>"), "| Tablas:", document.count("<w:tbl>"), "| Saltos de pagina:", document.count('w:type="page"'))
