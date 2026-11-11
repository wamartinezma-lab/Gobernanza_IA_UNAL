#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera una presentacion PowerPoint (.pptx) del dashboard de Gobernanza de IA de la UNAL
para un evento academico de 30 minutos, basada en el brief.
Sin dependencias externas: construye el OOXML (presentationml) con zipfile.
Incluye diagramas con formas y DOS graficos NATIVOS editables (barras y dona).
Salida: presentacion-dashboard-gobernanza-ia-unal.pptx
"""
import zipfile, html

EMU=914400            # 1 pulgada
PW=12192000           # 13.333 in (16:9)
PH=6858000            # 7.5 in
def inch(v): return int(v*EMU)

# Paleta (hex sin #)
VERDE="1f4037"; VERDEPF="132a24"; DORADO="b89347"; NARANJA="e8933f"
AZUL="2962ff"; TEAL="0f8c82"; CONTROL="e8591c"; SOCIO="17a06a"; CNED="1b355b"
BLANCO="ffffff"; CREMA="f4f1e9"; GRIS="7a7f88"; ROJO="c0392b"; PURPURA="6b4fa0"; TXT="1d2b26"

def xe(t): return html.escape(str(t),quote=True)

# ---------- helpers de shapes ----------
_sid=[1]
_last=[0]          # id del ultimo shape creado
def nid():
    _sid[0]+=1; _last[0]=_sid[0]; return _sid[0]
def last_id():
    return _last[0]
def peek_next():
    "id que recibira el proximo shape creado"
    return _sid[0]+1

def rect(x,y,w,h,fill,line=None,lw=12700,rich=None,round_=False):
    geom="roundRect" if round_ else "rect"
    ln=""
    if line: ln=f'<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    txt=rich or ""
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="r{_sid[0]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}</p:spPr>
<p:txBody><a:bodyPr anchor="ctr" lIns="68580" rIns="68580" tIns="27432" bIns="27432"/><a:lstStyle/>{txt}</p:txBody></p:sp>'''

def rect_alpha(x,y,w,h,fill,alpha=40000,line=None,lw=12700,round_=True):
    "rectangulo con relleno semitransparente (alpha en milesimas: 40000 = 40%)"
    geom="roundRect" if round_ else "rect"
    ln=f'<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else ''
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="ra{_sid[0]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>{ln}</p:spPr>
<p:txBody><a:bodyPr anchor="ctr" lIns="68580" rIns="68580" tIns="27432" bIns="27432"/><a:lstStyle/></p:txBody></p:sp>'''

def shape(prst,x,y,w,h,fill,line=None,lw=12700,rich=None):
    "forma preestablecida (prstGeom) de color solido; util para iconos editables"
    ln=f'<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
    txt=rich or ""
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="ic{_sid[0]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{ln}</p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>{txt}</p:txBody></p:sp>'''

def icon_badge(cx_in,cy_in,d_in,ring,glyph):
    """Icono circular (badge) con un glifo de texto centrado. cx,cy = centro en pulgadas.
    Devuelve (xml, [ids]) para poder animarlo."""
    r=d_in/2.0; x=inch(cx_in-r); y=inch(cy_in-r); d=inch(d_in)
    i1=peek_next(); xml=shape("ellipse",x,y,d,d,BLANCO,line=ring,lw=28575)
    i2=peek_next(); xml+=textbox(x,inch(cy_in-r-0.02),d,d,[para([run(glyph,int(d_in*2100),ring,b=True)],align="ctr")],anchor="ctr")
    return xml,[i1,i2]

def para(runs, align="l", space=0):
    r="".join(runs)
    return f'<a:p><a:pPr algn="{align}"/>{r}</a:p>'

def run(t,sz=1800,color=TXT,b=False,i=False,font="Calibri"):
    bb=' b="1"' if b else ""; ii=' i="1"' if i else ""
    return f'<a:r><a:rPr lang="es-CO" sz="{sz}"{bb}{ii}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{font}"/></a:rPr><a:t>{xe(t)}</a:t></a:r>'

def tb(paras):  # construye txBody-content
    return "".join(paras)

def textbox(x,y,w,h,paras,anchor="t"):
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="t{_sid[0]}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr anchor="{anchor}" wrap="square"/><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>'''

# ----- helpers estilo COMIC -----
def callout(x,y,w,h,fill,line_,paras,prst="wedgeRoundRectCallout",adj=None):
    "globo de dialogo tipo historieta (con cola)"
    av=f'<a:avLst>{adj}</a:avLst>' if adj else '<a:avLst/>'
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="cal{_sid[0]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="{prst}">{av}</a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
<a:ln w="28575"><a:solidFill><a:srgbClr val="{line_}"/></a:solidFill></a:ln></p:spPr>
<p:txBody><a:bodyPr anchor="ctr" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>'''

def panel(x,y,w,h,fill,label=None,lsz=1200):
    "vineta (panel) de comic con borde grueso negro"
    s=f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="pan{_sid[0]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
<a:ln w="38100"><a:solidFill><a:srgbClr val="132a24"/></a:solidFill></a:ln></p:spPr>
<p:txBody><a:bodyPr anchor="b" lIns="54864" rIns="54864" bIns="27432"/><a:lstStyle/>'''
    s+=(para([run(label,lsz,BLANCO,b=True)]) if label else '<a:p/>')
    s+='</p:txBody></p:sp>'
    return s

def star(x,y,w,h,fill,line_,txt,sz=1400):
    "estrella de comic tipo POW! (explosion)"
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="st{_sid[0]}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
<a:prstGeom prst="irregularSeal1"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
<a:ln w="28575"><a:solidFill><a:srgbClr val="{line_}"/></a:solidFill></a:ln></p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>{para([run(txt,sz,line_,b=True)],align="ctr")}</p:txBody></p:sp>'''

def line(x1,y1,x2,y2,color=DORADO,w=19050):
    x=min(x1,x2);y=min(y1,y2);cx=abs(x2-x1) or 1;cy=abs(y2-y1) or 1
    flipH="1" if x2<x1 else "0"; flipV="1" if y2<y1 else "0"
    return f'''<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{nid()}" name="c{_sid[0]}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm flipH="{flipH}" flipV="{flipV}"><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="{w}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>'''

# banda de encabezado + pie reutilizable
def chrome(titulo, kicker=None):
    s=rect(0,0,PW,inch(1.05),VERDEPF)
    s+=rect(0,0,inch(0.12),inch(1.05),DORADO)
    paras=[para([run(titulo,2600,BLANCO,b=True)])]
    if kicker: paras=[para([run(kicker,1100,DORADO,b=True)]),para([run(titulo,2600,BLANCO,b=True)])]
    s+=textbox(inch(0.45),inch(0.08),inch(12),inch(0.95),paras,anchor="ctr")
    s+=rect(0,PH-inch(0.28),PW,inch(0.28),VERDE)
    s+=textbox(inch(0.45),PH-inch(0.27),inch(11),inch(0.26),[para([run("DNED-DIVULAB  ·  Universidad Nacional de Colombia",1000,CREMA)])],anchor="ctr")
    return s

_card_ids=[]   # ids del ultimo card creado
def card(x,y,w,h,accent,titulo,lineas,tsz=1400,lsz=1100):
    global _card_ids; _card_ids=[]
    i1=peek_next(); s=rect(x,y,w,h,BLANCO,line="e2e7e3",round_=True); _card_ids.append(i1)
    i2=peek_next(); s+=rect(x,y,w,inch(0.12),accent,round_=False); _card_ids.append(i2)
    paras=[para([run(titulo,tsz,accent,b=True)])]
    for ln in lineas:
        paras.append(para([run("•  "+ln,lsz,TXT)]))
    i3=peek_next(); s+=textbox(x+inch(0.12),y+inch(0.2),w-inch(0.24),h-inch(0.3),paras); _card_ids.append(i3)
    return s
def card_anim(effect='flyB'):
    "devuelve un paso de animacion con todos los shapes del ultimo card"
    return [(i,effect) for i in _card_ids]

# ======================================================================
#   GRAFICOS NATIVOS (editables) — barras y dona
# ======================================================================
def chart_rels(idx): # rels del grafico (ninguno externo necesario salvo colors/style opcionales)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')

def bar_chart_xml():
    # votacion de la plenaria: a favor / salvaguardas / objecion (ejemplo caso desercion)
    cats=["A favor","Con salvaguardas","Objeción"]
    vals=[0,3,1]
    colors=[SOCIO,NARANJA,ROJO]
    pts="".join(f'<c:pt idx="{i}"><c:v>{xe(c)}</c:v></c:pt>' for i,c in enumerate(cats))
    vpts="".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
    dpts="".join(
        f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="{colors[i]}"/></a:solidFill></c:spPr></c:dPt>'
        for i in range(len(cats)))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="es-CO" sz="1200" b="1"/><a:t>Votación de la plenaria (caso: deserción)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>
<c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Hoja1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Votos</c:v></c:pt></c:strCache></c:strRef></c:tx>
{dpts}
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>
<c:cat><c:strRef><c:f>Hoja1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>{pts}</c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Hoja1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>{vpts}</c:numCache></c:numRef></c:val>
</c:ser><c:axId val="111"/><c:axId val="222"/></c:barChart>
<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>
<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>
</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'''

def doughnut_chart_xml():
    cats=["Entregados","Pendientes"]; vals=[11,3]; colors=[SOCIO,NARANJA]
    pts="".join(f'<c:pt idx="{i}"><c:v>{xe(c)}</c:v></c:pt>' for i,c in enumerate(cats))
    vpts="".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
    dpts="".join(f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="{colors[i]}"/></a:solidFill></c:spPr></c:dPt>' for i in range(2))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="es-CO" sz="1200" b="1"/><a:t>Documentos: 11 entregados · 3 pendientes</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>
<c:plotArea><c:layout/>
<c:doughnutChart><c:varyColors val="1"/>
<c:ser><c:idx val="0"/><c:order val="0"/>{dpts}
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>
<c:cat><c:strRef><c:f>Hoja1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>{pts}</c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Hoja1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/>{vpts}</c:numCache></c:numRef></c:val>
</c:ser><c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>
</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'''

def graphicframe(rid,x,y,w,h):
    return f'''<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{nid()}" name="chart{_sid[0]}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/></a:graphicData></a:graphic></p:graphicFrame>'''

# ======================================================================
#   DIAPOSITIVAS
# ======================================================================
slides=[]   # cada entry: (shapes_xml, charts, anim_steps)

def add_slide(shapes, charts=None, anims=None, no_transition=False):
    slides.append((shapes, charts or [], anims or [], no_transition))

# ----------------------------------------------------------------------
#   ANIMACIONES (timing tree OOXML)
#   anims = lista de "pasos"; cada paso = lista de (shape_id, efecto)
#   efecto: 'fade','flyL','flyR','flyB','flyT','wipeL','wipeU','zoom'
#   El primer paso arranca con la diapositiva; los demás por clic.
# ----------------------------------------------------------------------
EFFECTS={
 'fade'  :('fade',None),
 'flyL'  :('wipe','fromLeft'),   # entra deslizando desde la izquierda
 'flyR'  :('wipe','fromRight'),
 'flyB'  :('wipe','fromBottom'),
 'flyT'  :('wipe','fromTop'),
 'wipeL' :('wipe','fromLeft'),
 'wipeU' :('wipe','fromBottom'),
 'zoom'  :('fade',None),         # fallback simple y robusto
}

# Tiempos de animación (ms) — ajustados para transiciones ágiles conservando los efectos.
ANIM_DUR=320      # duración de cada efecto de entrada (antes 500)
ANIM_STAGGER=90   # retardo escalonado entre shapes de un mismo paso (antes 180)

def build_timing(anim_steps):
    """Genera <p:timing> con una secuencia principal de clicks."""
    if not anim_steps: return ""
    click_pars=[]
    for si,step in enumerate(anim_steps):
        # dentro de un paso, los shapes entran casi juntos (escalonado leve)
        inner=[]
        for k,(spid,effect) in enumerate(step):
            filt,_dir=EFFECTS.get(effect,('fade',None))
            delay = 0 if k==0 else ANIM_STAGGER*k   # escalonado dentro del paso
            # nodeType: primer efecto de cada paso espera clic (afterEffect en 1er paso = withGroup)
            node = "afterEffect" if k>0 else ("afterEffect" if si==0 else "onClick")
            animeff = (f'<p:animEffect transition="in" filter="{filt}">'
                       f'<p:cBhvr><p:cTn id="{nid()}" dur="{ANIM_DUR}"/>'
                       f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>')
            setvis = (f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                      f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                      f'<p:to><p:strVal val="visible"/></p:to></p:set>')
            inner.append(
              f'<p:par><p:cTn id="{nid()}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" '
              f'grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
              f'<p:childTnLst>{setvis}{animeff}</p:childTnLst></p:cTn></p:par>')
        click_pars.append(
          f'<p:par><p:cTn id="{nid()}" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="{"0" if si==0 else "indefinite"}"/></p:stCondLst>'
          f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>')
    seq=(f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="{nid()}" dur="indefinite" nodeType="mainSeq">'
         f'<p:childTnLst>{"".join(click_pars)}</p:childTnLst></p:cTn>'
         f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
         f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>')
    return (f'<p:timing><p:tnLst><p:par><p:cTn id="{nid()}" dur="indefinite" restart="never" nodeType="tmRoot">'
            f'<p:childTnLst>{seq}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')

# ---- 1. Portada ----
def s_portada():
    s=rect(0,0,PW,PH,VERDEPF)
    s+=rect(0,0,PW,inch(0.14),DORADO)
    s+=rect(0,PH-inch(0.14),PW,inch(0.14),NARANJA)
    s+=rect(inch(0.6),inch(1.0),inch(0.9),inch(0.9),VERDE,line=DORADO)
    s+=textbox(inch(0.6),inch(1.25),inch(0.9),inch(0.5),[para([run("UNAL",1400,DORADO,b=True)],align="ctr")])
    tid=peek_next()
    s+=textbox(inch(0.6),inch(2.3),inch(12),inch(2.2),[
        para([run("Ponencia",2000,DORADO,b=True)]),
        para([run("Hacia un ecosistema de gobernanza para el uso",3200,BLANCO,b=True)]),
        para([run("responsable y la apropiación ética de la IA en la UNAL",3200,BLANCO,b=True)]),
    ])
    aid=peek_next()
    s+=textbox(inch(0.6),inch(5.2),inch(12),inch(1.4),[
        para([run("Prof. Francisco Albeiro Gómez Jaramillo   ·   Ing. Wilver Alexander Martínez Martínez",1500,CREMA)]),
        para([run("Dirección Nacional de Estrategia Digital (DNED) · División Universidad Laboratorio (DIVULAB)",1300,DORADO)]),
    ])
    add_slide(s, anims=[[(tid,'fade')],[(aid,'flyB')]])

# ---- 2. Apertura (ESTILO COMIC, animado por paneles) ----
def s_apertura():
    s=chrome("¿Quién decide qué puede hacer la IA?","Sección 0 · Apertura")
    s+=textbox(inch(0.6),inch(1.15),inch(12),inch(0.5),[
        para([run("Un caso cotidiano contado en tres viñetas:",1500,GRIS,i=True)])])
    # tres paneles de comic
    pw=3.95; px=[0.6,4.75,8.9]; py=1.75; ph=2.5
    p1=peek_next(); s+=panel(inch(px[0]),inch(py),inch(pw),inch(ph),"eef3f1","1. Un profesor usa IA")
    p2=peek_next(); s+=panel(inch(px[1]),inch(py),inch(pw),inch(ph),"eaf0ff","2. ¿Y los datos sensibles?")
    p3=peek_next(); s+=panel(inch(px[2]),inch(py),inch(pw),inch(ph),"fdeef0","3. ¿Quién lo autoriza?")
    # globos de dialogo (callouts) dentro de cada panel
    b1=peek_next(); s+=callout(inch(0.85),inch(1.95),inch(3.4),inch(1.35),BLANCO,AZUL,
        [para([run("“Un sistema predice qué estudiantes van a desertar.”",1200,TXT,b=True)])])
    b2=peek_next(); s+=callout(inch(5.0),inch(1.95),inch(3.4),inch(1.35),BLANCO,TEAL,
        [para([run("“Usa notas, salud y datos socioeconómicos.”",1200,TXT,b=True)])])
    b3=peek_next(); s+=callout(inch(9.15),inch(1.95),inch(3.4),inch(1.35),BLANCO,CONTROL,
        [para([run("“¿Puede desplegarse? ¿Quién responde?”",1200,TXT,b=True)])])
    # cierre: estrella comic + mensaje
    stx=peek_next(); s+=star(inch(0.7),inch(4.55),inch(2.3),inch(1.9),NARANJA,CONTROL,"¡La pregunta clave!",1200)
    msg=peek_next(); s+=textbox(inch(3.3),inch(4.8),inch(9.2),inch(1.4),[
        para([run("No es SI usar la IA, sino QUIÉN decide qué se puede hacer,",2000,TXT,b=True)]),
        para([run("con qué datos y bajo qué límites — de forma trazable y responsable.",1800,TXT)])])
    # animacion tipo comic: panel -> globo, uno a uno; luego la estrella y el mensaje
    add_slide(s, anims=[
        [(p1,'fade'),(b1,'zoom')],
        [(p2,'fade'),(b2,'zoom')],
        [(p3,'fade'),(b3,'zoom')],
        [(stx,'zoom'),(msg,'fade')]])

# ---- 3. El problema ----
def s_problema():
    s=chrome("La gobernanza de la IA supera al gobierno de TI","Sección 1 · El problema")

    # ===== COLUMNA IZQUIERDA: gobierno institucional + ejes misionales + ambiente =====
    L1=peek_next(); s+=rect(inch(0.5),inch(1.35),inch(4.0),inch(4.75),CNED,round_=True)
    L1b=peek_next(); s+=textbox(inch(0.5),inch(1.5),inch(4.0),inch(0.8),[
        para([run("Gobierno institucional",1250,BLANCO,b=True)],align="ctr"),
        para([run("(corporativo · UNAL)",1050,CREMA)],align="ctr")])
    # ejes misionales (3) + ambiente
    ejes=[("Docencia",SOCIO),("Investigación",AZUL),("Extensión",CONTROL),("Componente ambiental · Green AI","2e8b57")]
    Lsteps=[]; ey=2.45
    for t,c in ejes:
        r=peek_next(); s+=rect(inch(0.75),inch(ey),inch(3.5),inch(0.62),c,round_=True)
        tx=peek_next(); s+=textbox(inch(0.75),inch(ey+0.12),inch(3.5),inch(0.4),[para([run(t,1100,BLANCO,b=True)],align="ctr")])
        Lsteps.append([(r,'flyL'),(tx,'fade')]); ey+=0.78
    Lcap=peek_next(); s+=textbox(inch(0.6),inch(5.68),inch(3.8),inch(0.4),[para([run("Ejes misionales + ética y ambiente",950,CREMA,i=True)],align="ctr")])

    # ===== COLUMNA CENTRAL: transición TI/Datos -> Gobernanza de IA =====
    # Gobernanza de IA (emerge arriba)
    G=peek_next(); s+=rect(inch(4.9),inch(1.5),inch(3.5),inch(1.5),TEAL,round_=True)
    Gb=peek_next(); s+=textbox(inch(4.9),inch(1.72),inch(3.5),inch(1.1),[
        para([run("Gobernanza de la IA",1300,BLANCO,b=True)],align="ctr"),
        para([run("surge la necesidad de gobierno",1000,CREMA,i=True)],align="ctr")])
    # Transición desde los modelos de TI y de Datos (cajas base, semitransparentes)
    TI=peek_next(); s+=rect_alpha(inch(4.9),inch(4.9),inch(1.65),inch(1.1),AZUL,alpha=45000,line=AZUL)
    TIb=peek_next(); s+=textbox(inch(4.9),inch(5.15),inch(1.65),inch(0.7),[
        para([run("Gobierno",1000,VERDEPF,b=True)],align="ctr"),para([run("de TI",1000,VERDEPF,b=True)],align="ctr")])
    DA=peek_next(); s+=rect_alpha(inch(6.75),inch(4.9),inch(1.65),inch(1.1),PURPURA,alpha=45000,line=PURPURA)
    DAb=peek_next(); s+=textbox(inch(6.75),inch(5.15),inch(1.65),inch(0.7),[
        para([run("Gobierno",1000,VERDEPF,b=True)],align="ctr"),para([run("de datos",1000,VERDEPF,b=True)],align="ctr")])
    # flechas de "alimentación" (transición) hacia la gobernanza
    aTI=peek_next(); s+=line(inch(5.7),inch(4.9),inch(6.0),inch(3.05),AZUL,22225)
    aDA=peek_next(); s+=line(inch(7.55),inch(4.9),inch(7.2),inch(3.05),PURPURA,22225)
    almcap=peek_next(); s+=textbox(inch(4.9),inch(4.2),inch(3.5),inch(0.55),[para([run("alimentan / transición →",1000,GRIS,i=True)],align="ctr")])

    # ===== COLUMNA DERECHA: cómo aparece la necesidad de gobierno =====
    s+=card(inch(8.85),inch(1.45),inch(3.95),inch(4.6),DORADO,"¿Por qué aparece la necesidad de gobierno?",
            ["Opacidad (caja negra) y sesgos","Fuga de información estratégica","Pérdida de supervisión humana","Huella de carbono (Green AI)","Decisiones que afectan derechos","Abarca las 3 misiones + ética y ambiente"],1300,1100)
    ac=card_anim('flyR')

    add_slide(s, anims=[
        [(L1,'zoom'),(L1b,'fade')],
        Lsteps[0], Lsteps[1], Lsteps[2], Lsteps[3]+[(Lcap,'fade')],
        [(TI,'fade'),(TIb,'fade'),(DA,'fade'),(DAb,'fade')],
        [(aTI,'fade'),(aDA,'fade'),(almcap,'fade'),(G,'zoom'),(Gb,'fade')],
        ac])

# ---- 4. Ecosistema institucional ----
def s_ecosistema():
    s=chrome("El ecosistema de gobernanza: actores y flujo","Esfuerzo institucional")
    s+=textbox(inch(0.6),inch(1.12),inch(12.2),inch(0.45),[para([run("Ejemplo del modelo: cómo se relacionan los actores y fluye una iniciativa de IA por el ecosistema.",1300,GRIS,i=True)])])

    # ===== FILA DE ACTORES con iconos (badges) =====
    #   CNED --- Mesa Técnica (4 núcleos) --- Comunidad (9 sedes) ; abajo: mesas/guías alimentan
    cy=2.55
    cn_x, me_x, co_x = 2.1, 6.65, 11.2   # centros en pulgadas
    # cajas de fondo por actor
    A1=peek_next(); s+=rect(inch(cn_x-1.5),inch(1.75),inch(3.0),inch(1.9),CNED,round_=True)
    A2=peek_next(); s+=rect(inch(me_x-1.6),inch(1.75),inch(3.2),inch(1.9),TEAL,round_=True)
    A3=peek_next(); s+=rect(inch(co_x-1.5),inch(1.75),inch(3.0),inch(1.9),SOCIO,round_=True)
    # iconos (glifos sobrios dentro de badge circular blanco)
    ic1,ids1=icon_badge(cn_x,2.25,0.78,CNED,"\u2302")      # casa/institución (CNED)
    ic2,ids2=icon_badge(me_x,2.25,0.78,TEAL,"\u2699")      # engranaje (Mesa Técnica)
    ic3,ids3=icon_badge(co_x,2.25,0.78,SOCIO,"\u25CF\u25CF")  # personas/comunidad
    s+=ic1+ic2+ic3
    # etiquetas de actor
    L1=peek_next(); s+=textbox(inch(cn_x-1.5),inch(2.82),inch(3.0),inch(0.8),[para([run("CNED",1350,BLANCO,b=True)],align="ctr"),para([run("define política y aprueba",950,CREMA)],align="ctr")])
    L2=peek_next(); s+=textbox(inch(me_x-1.6),inch(2.82),inch(3.2),inch(0.8),[para([run("Mesa Técnica de IA",1300,BLANCO,b=True)],align="ctr"),para([run("4 núcleos · evalúa y decide",950,CREMA)],align="ctr")])
    L3=peek_next(); s+=textbox(inch(co_x-1.5),inch(2.82),inch(3.0),inch(0.8),[para([run("Comunidad · 9 sedes",1250,BLANCO,b=True)],align="ctr"),para([run("usa y co-crea",950,CREMA)],align="ctr")])

    # ===== FLECHAS DE FLUJO (ida: política/lineamiento ; vuelta: caso/co-creación) =====
    f1=peek_next(); s+=line(inch(cn_x+1.5),inch(2.25),inch(me_x-1.65),inch(2.25),AZUL,22225)   # CNED -> Mesa
    f1l=peek_next(); s+=textbox(inch(cn_x+1.4),inch(1.9),inch(2.9),inch(0.3),[para([run("política / norma",900,AZUL,b=True)],align="ctr")])
    f2=peek_next(); s+=line(inch(me_x+1.6),inch(2.25),inch(co_x-1.55),inch(2.25),AZUL,22225)   # Mesa -> Comunidad
    f2l=peek_next(); s+=textbox(inch(me_x+1.5),inch(1.9),inch(2.9),inch(0.3),[para([run("lineamientos / guías",900,AZUL,b=True)],align="ctr")])
    # retorno (abajo): comunidad radica caso -> mesa -> ponencia -> CNED
    r1=peek_next(); s+=line(inch(co_x-1.55),inch(3.25),inch(me_x+1.6),inch(3.25),DORADO,22225)  # Comunidad -> Mesa
    r2=peek_next(); s+=line(inch(me_x-1.65),inch(3.25),inch(cn_x+1.5),inch(3.25),DORADO,22225)  # Mesa -> CNED
    rl=peek_next(); s+=textbox(inch(cn_x+1.4),inch(3.32),inch(5.8),inch(0.3),[para([run("← radica caso / co-creación / ponencia",900,DORADO,b=True)],align="ctr")])

    # ===== ALIMENTACIÓN: 14 mesas y guías que nutren a la Mesa Técnica =====
    MB=peek_next(); s+=rect_alpha(inch(me_x-1.7),inch(4.35),inch(3.4),inch(0.95),PURPURA,alpha=50000,line=PURPURA)
    MBi,idsMB=icon_badge(me_x-1.15,4.82,0.55,PURPURA,"\u2630")  # mesas/documentos
    s+=MBi
    MBt=peek_next(); s+=textbox(inch(me_x-0.75),inch(4.45),inch(2.4),inch(0.8),[para([run("14 mesas + guías",1050,VERDEPF,b=True)],align="ctr"),para([run("conocimiento situado",900,VERDEPF)],align="ctr")])
    fa=peek_next(); s+=line(inch(me_x),inch(4.35),inch(me_x),inch(3.68),PURPURA,22225)  # mesas -> Mesa Técnica
    faL=peek_next(); s+=textbox(inch(me_x+0.1),inch(3.85),inch(2.2),inch(0.3),[para([run("alimentan",900,PURPURA,b=True,i=True)])])

    # ===== ROL DE LOS ELEMENTOS (franja inferior) =====
    roles=[("CNED",CNED,"Dirige y aprueba"),("Mesa Técnica",TEAL,"Evalúa, decide, acompaña"),
           ("Comunidad",SOCIO,"Usa, co-crea, radica casos"),("Mesas + guías",PURPURA,"Nutren el marco"),
           ("Validación acad.",DORADO,"Legitima el modelo")]
    rx=0.6; rw=2.44; rsteps=[]
    for t,c,d in roles:
        rr=peek_next(); s+=rect(inch(rx),inch(5.55),inch(rw),inch(1.05),c,round_=True)
        rt=peek_next(); s+=textbox(inch(rx),inch(5.66),inch(rw),inch(0.9),[para([run(t,1050,BLANCO,b=True)],align="ctr"),para([run(d,900,CREMA)],align="ctr")])
        rsteps.append([(rr,'flyB'),(rt,'fade')]); rx+=rw+0.08

    add_slide(s, anims=[
        [(A1,'zoom')]+[(i,'fade') for i in ids1]+[(L1,'fade')],
        [(A2,'zoom')]+[(i,'fade') for i in ids2]+[(L2,'fade')],
        [(A3,'zoom')]+[(i,'fade') for i in ids3]+[(L3,'fade')],
        [(f1,'fade'),(f1l,'fade'),(f2,'fade'),(f2l,'fade')],
        [(r1,'fade'),(r2,'fade'),(rl,'fade')],
        [(MB,'fade')]+[(i,'fade') for i in idsMB]+[(MBt,'fade'),(fa,'fade'),(faL,'fade')],
    ]+rsteps)

# ---- 5. Benchmarking ----
def s_benchmarking():
    s=chrome("Qué hacen otras universidades","Sección 2 · Benchmarking")
    s+=textbox(inch(0.6),inch(1.15),inch(12),inch(0.5),[para([run("Tres modelos del Norte Global + referentes del Sur Global, más cercanos a la realidad de la UNAL.",1400,GRIS,i=True)])])
    s+=textbox(inch(0.6),inch(1.65),inch(5),inch(0.4),[para([run("NORTE GLOBAL (referencia)",1100,CNED,b=True)])])
    norte=[("MIT",AZUL,["Descentralizado","SERC · programa semilla USD 100k"]),
           ("UCLA",DORADO,["Matricial-centralizado","UC AI Council · veto clínico HIPAA"]),
           ("Stanford",ROJO,["Hub-and-spoke","HAI · AI Index · huella de carbono"])]
    x=0.6
    for t,c,d in norte:
        s+=card(inch(x),inch(2.05),inch(3.9),inch(1.75),c,t,d,1600,1050); x+=4.1
    s+=textbox(inch(0.6),inch(3.95),inch(8),inch(0.4),[para([run("SUR GLOBAL / LATINOAMÉRICA",1100,SOCIO,b=True)])])
    sur=[("Tec Monterrey",SOCIO,["México · multicampus","Integridad académica con IA"]),
         ("USP",SOCIO,["Brasil · normativo","Mayor IES pública de LatAm"]),
         ("UNAM",SOCIO,["México · en construcción","Universidad pública nacional"]),
         ("INTEGRITY-AI",DORADO,["Red colombiana/LatAm","U. del Cauca · U. San Buenaventura"])]
    x=0.6; w=(12.8-0.3*3)/4
    for t,c,d in sur:
        s+=card(inch(x),inch(4.35),inch(w),inch(1.75),c,t,d,1350,1000); x+=w+0.3
    # sin animacion (comparativa para lectura)
    add_slide(s)

# ---- 6. Modelo multinivel ----
def s_multinivel():
    s=chrome("El modelo UNAL: multinivel","Sección 3 · El modelo")
    n1=peek_next(); s+=rect(inch(3.6),inch(1.5),inch(6.1),inch(1.0),CNED,round_=True)
    n1b=peek_next(); s+=textbox(inch(3.6),inch(1.6),inch(6.1),inch(0.8),[para([run("CNED — definición política y aprobación normativa",1400,BLANCO,b=True)],align="ctr")])
    n2=peek_next(); s+=rect(inch(2.6),inch(2.9),inch(8.1),inch(1.0),TEAL,round_=True)
    n2b=peek_next(); s+=textbox(inch(2.6),inch(3.0),inch(8.1),inch(0.8),[para([run("Mesa Técnica de IA — ejecución · 4 núcleos · PLASI · ciclo trienal",1400,BLANCO,b=True)],align="ctr")])
    n3=peek_next(); s+=rect(inch(1.6),inch(4.3),inch(10.1),inch(1.0),SOCIO,round_=True)
    n3b=peek_next(); s+=textbox(inch(1.6),inch(4.4),inch(10.1),inch(0.8),[para([run("Comunidad Universitaria — 9 sedes · facultades · grupos · laboratorios vivos",1400,BLANCO,b=True)],align="ctr")])
    # descripciones laterales (mas detalle)
    s+=textbox(inch(1.0),inch(1.55),inch(2.4),inch(0.9),[para([run("Alto nivel: política y norma",1000,GRIS)])])
    s+=textbox(inch(1.0),inch(2.95),inch(1.4),inch(0.9),[para([run("Ejecución",1000,GRIS)])])
    s+=textbox(inch(0.3),inch(4.35),inch(1.2),inch(0.9),[para([run("Base",1000,GRIS)])])
    n4=peek_next(); s+=textbox(inch(1.6),inch(5.55),inch(10.1),inch(0.9),[
        para([run("Subsidiariedad: la decisión se toma en el nivel más cercano al riesgo y al valor;",1250,TXT,b=True)],align="ctr"),
        para([run("se centraliza solo por excepción (alto riesgo, datos sensibles o estándares comunes).",1200,GRIS)],align="ctr")])
    # animacion ligera: solo se arma la piramide por niveles
    add_slide(s, anims=[[(n1,'flyT'),(n1b,'fade')],[(n2,'flyL'),(n2b,'fade')],[(n3,'flyB'),(n3b,'fade')]])

# ---- 7. 4 nucleos ----
def s_nucleos():
    s=chrome("Los cuatro núcleos de la Mesa Técnica","Composición")
    n=[("1.1 Estratégico Misional",AZUL,["Vicerrectorías","Coherencia misional, soberanía epistémica"]),
       ("1.2 Técnico & Green Computing",TEAL,["DNED, ciberseguridad, HPC","Infraestructura, DT 1-25, eficiencia"]),
       ("1.3 Control & Jurídico-Fiscal",CONTROL,["Planeación, jurídica, control interno","Legalidad, RAEE, Habeas Data"]),
       ("1.4 Socio-Ambiental & Territorial",SOCIO,["Investigadores, sedes, estudiantes","Ética, saberes territoriales, salud mental"])]
    s+=textbox(inch(0.6),inch(1.15),inch(12),inch(0.45),[para([run("Cada núcleo aporta una mirada; juntos cubren la gobernanza de forma multidimensional.",1350,GRIS,i=True)])])
    xs=[0.6,6.6]; ys=[1.7,4.05]; i=0
    for ry in ys:
        for rx in xs:
            t,c,d=n[i]; s+=card(inch(rx),inch(ry),inch(5.9),inch(2.15),c,t,d,1500,1100); i+=1
    # sin animacion (composicion para lectura)
    add_slide(s)

# ---- 8. Operacion plenaria (ESTILO COMIC: el caso recorre el proceso) ----
def s_plenaria():
    s=chrome("Cómo opera la Mesa: la plenaria","Sección 4 · Operación")
    s+=textbox(inch(0.6),inch(1.12),inch(12),inch(0.45),[
        para([run("Seguimos el caso “predictor de deserción” paso a paso dentro de la Mesa:",1400,GRIS,i=True)])])
    # Panel central: la plenaria recibe el caso
    pl=peek_next(); s+=rect(inch(4.6),inch(1.7),inch(4.1),inch(0.95),TEAL,round_=True)
    plb=peek_next(); s+=textbox(inch(4.6),inch(1.8),inch(4.1),inch(0.75),[
        para([run("PLENARIA DE DELEGADOS",1200,BLANCO,b=True)],align="ctr"),
        para([run("recibe y radica el caso",1000,"d6fff8")],align="ctr")])
    # globo: la plenaria reparte
    gb0=peek_next(); s+=callout(inch(8.9),inch(1.55),inch(3.6),inch(1.0),BLANCO,TEAL,
        [para([run("“Trasladamos a trámite a los 4 núcleos.”",1100,TXT,b=True)])],prst="wedgeRoundRectCallout")
    # 4 nucleos, cada uno con su dictamen (globo) estilo comic
    nuc=[("1.1 Estratégico",AZUL,"“Es pertinente a la misión.”","Favorable"),
         ("1.2 Técnico/Green",TEAL,"“Viable, cuidando la huella.”","Con salvaguardas"),
         ("1.3 Jurídico/Control",CONTROL,"“¡Datos sensibles! Riesgo a derechos.”","OBJECIÓN"),
         ("1.4 Socio-Ambiental",SOCIO,"“Exige salvaguardas éticas.”","Con salvaguardas")]
    x=0.65; w=2.95; nsteps=[]
    for t,c,frase,dict_ in nuc:
        r=peek_next(); s+=rect(inch(x),inch(3.55),inch(w),inch(0.95),BLANCO,line=c,round_=True)
        tx=peek_next(); s+=textbox(inch(x),inch(3.62),inch(w),inch(0.8),[
            para([run(t,1150,c,b=True)],align="ctr"),para([run(dict_,1000,c)],align="ctr")])
        gb=peek_next(); s+=callout(inch(x),inch(4.65),inch(w),inch(1.05),"fbf6ea",c,
            [para([run(frase,1000,TXT)],align="ctr")],prst="wedgeRoundRectCallout")
        ln=peek_next(); s+=line(inch(6.65),inch(2.65),inch(x+w/2),inch(3.55),GRIS,9525)
        nsteps.append([(r,'flyB'),(tx,'fade'),(gb,'zoom'),(ln,'fade')])
        x+=w+0.08
    # estrella comic: resultado
    stx=peek_next(); s+=star(inch(9.3),inch(3.5),inch(3.0),inch(2.1),NARANJA,CONTROL,"¡Escala al CNED!",1250)
    note=peek_next(); s+=textbox(inch(0.6),inch(5.95),inch(12),inch(0.5),[
        para([run("Quórum ≥ 3 de 4. La objeción de 1.3/1.4 por derechos prevalece → decisión colegiada: elevar al CNED con veto.",1150,GRIS)])])
    add_slide(s, anims=[
        [(pl,'zoom'),(plb,'fade'),(gb0,'zoom')],
        nsteps[0],nsteps[1],nsteps[2],nsteps[3],
        [(stx,'zoom'),(note,'fade')]])

# ---- 9. Quorum y votacion (GRAFICO) ----
def s_votacion():
    s=chrome("Quórum y votación de la plenaria","Sección 4 · Decisión")
    s+=graphicframe("rId2",inch(0.6),inch(1.4),inch(7.2),inch(4.6))
    s+=card(inch(8.1),inch(1.5),inch(4.3),inch(4.5),DORADO,"Regla de decisión",
            ["Quórum: ≥ 3 de 4 delegados para sesionar","Objeción de 1.3 (jurídico) o 1.4 (ético)","por riesgo a derechos → prevalece",
             "Mayoría o salvaguardas → condicionado","Unanimidad a favor → favorable",
             "Caso deserción: Alto → escala al CNED"],1500,1050)
    # grafico editable; sin animacion para que se lea de inmediato
    add_slide(s, charts=[("rId2","bar")])

# ---- 10. Riesgo ----
def s_riesgo():
    s=chrome("Riesgo y taxonomía de sistemas de IA","Sección 5 · Riesgo")
    niveles=[("Inaceptable","7a1414","Prohibido"),("Alto",ROJO,"CNED + auditoría + veto"),
             ("Medio",NARANJA,"Sede + supervisión Mesa"),("Bajo",SOCIO,"Facultad + registro")]
    y=1.5; rsteps=[]
    for t,c,d in niveles:
        r=peek_next(); s+=rect(inch(0.6),inch(y),inch(2.2),inch(0.8),c,round_=True)
        t1=peek_next(); s+=textbox(inch(0.6),inch(y+0.2),inch(2.2),inch(0.5),[para([run(t,1300,BLANCO,b=True)],align="ctr")])
        t2=peek_next(); s+=textbox(inch(2.95),inch(y+0.18),inch(4.0),inch(0.5),[para([run(d,1250,TXT)])])
        rsteps.append([(r,'flyL'),(t1,'fade'),(t2,'fade')]); y+=1.0
    s+=card(inch(7.4),inch(1.5),inch(5.0),inch(4.4),PURPURA,"Taxonomía (por tipo)",
            ["Decisión académico-administrativa","IA en el aula","Investigación y extensión","IA propia de la UNAL","IA generativa de terceros","IA agéntica"],1500,1200); ac=card_anim('flyR')
    nid_=peek_next(); s+=textbox(inch(0.6),inch(5.7),inch(6.5),inch(0.6),[para([run("El veto se gradúa por riesgo a derechos, no por el dominio (Rec. 4).",1200,GRIS,i=True)])])
    add_slide(s, anims=[ac]+rsteps+[[(nid_,'fade')]], no_transition=True)  # Sección 5 · Riesgo

# ---- 11. RACI (tabla con shapes) ----
def s_raci():
    s=chrome("Matriz RACI: quién decide qué","Sección 5 · Roles")
    acts=["Proponer","Evaluar riesgo","Autorizar (alto)","Implementar","Supervisar","Suspender (veto)"]
    cols=["Facultad","Sede","Mesa Téc.","CNED","Jurídico"]
    data=[["R","C","I","—","—"],["C","C","R","I","C"],["I","C","C","A","C"],["R","R","C","—","—"],["I","C","S","I","S"],["I","I","S","A","S"]]
    cc={"R":AZUL,"A":ROJO,"C":TEAL,"I":GRIS,"S":CONTROL,"—":"c5ccc8"}
    x0=0.6; y0=1.5; cw0=3.0; cw=1.75; rh=0.62
    # encabezado
    hdr=[]
    hid=peek_next(); s+=rect(inch(x0),inch(y0),inch(cw0),inch(rh),VERDEPF); hdr.append((hid,'fade'))
    s+=textbox(inch(x0),inch(y0+0.13),inch(cw0),inch(0.4),[para([run("Actividad",1150,BLANCO,b=True)])])
    for j,cn in enumerate(cols):
        r=peek_next(); s+=rect(inch(x0+cw0+j*cw),inch(y0),inch(cw),inch(rh),VERDEPF); hdr.append((r,'fade'))
        s+=textbox(inch(x0+cw0+j*cw),inch(y0+0.13),inch(cw),inch(0.4),[para([run(cn,1050,BLANCO,b=True)],align="ctr")])
    rowsteps=[]
    for i,act in enumerate(acts):
        yy=y0+rh*(i+1); step=[]
        r=peek_next(); s+=rect(inch(x0),inch(yy),inch(cw0),inch(rh),CREMA); step.append((r,'flyL'))
        s+=textbox(inch(x0+0.1),inch(yy+0.13),inch(cw0-0.1),inch(0.4),[para([run(act,1050,TXT)])])
        for j,v in enumerate(data[i]):
            r2=peek_next(); s+=rect(inch(x0+cw0+j*cw),inch(yy),inch(cw),inch(rh),cc[v]); step.append((r2,'fade'))
            s+=textbox(inch(x0+cw0+j*cw),inch(yy+0.13),inch(cw),inch(0.4),[para([run(v,1100,BLANCO if v!="—" else GRIS,b=True)],align="ctr")])
        rowsteps.append(step)
    # leyenda RACI (mas detalle)
    leg=peek_next()
    s+=textbox(inch(0.6),inch(5.75),inch(12),inch(0.9),[
        para([run("R",1100,AZUL,b=True),run(" Responsable/propone-implementa    ",1050,TXT),
              run("A",1100,ROJO,b=True),run(" Autoriza (rinde cuentas)    ",1050,TXT),
              run("C",1100,TEAL,b=True),run(" Consultado/evalúa    ",1050,TXT),
              run("I",1100,GRIS,b=True),run(" Informado    ",1050,TXT),
              run("S",1100,CONTROL,b=True),run(" Supervisa/puede suspender",1050,TXT)]),
        para([run("Define no solo quién participa, sino quién decide qué según riesgo y misión; el veto responde al riesgo a derechos.",1050,GRIS,i=True)])])
    # Sección 5 · Roles: sin animaciones internas (todos los elementos visibles de una vez) y sin transición.
    add_slide(s, no_transition=True)

# ---- 12. Hoja de ruta ----
def s_ruta():
    s=chrome("Hoja de ruta 2025–2027","Sección 6 · Plan")
    fases=[("Diagnóstico","2S 2025",SOCIO),("Co-construcción","2026",AZUL),("Validación","Oct 2026",CONTROL),
           ("Revisión/aprob.","26–27",PURPURA),("Implementación","2027","4b3b8f")]
    x=0.6; w=2.25; steps=[]
    for i,(t,d,c) in enumerate(fases):
        r=peek_next(); s+=rect(inch(x),inch(3.0),inch(w),inch(1.5),c,round_=True)
        tx=peek_next(); s+=textbox(inch(x),inch(3.25),inch(w),inch(1.0),[para([run(t,1300,BLANCO,b=True)],align="ctr"),para([run(d,1100,CREMA)],align="ctr")])
        step=[(r,'flyL'),(tx,'fade')]
        if i<4:
            ln=peek_next(); s+=line(inch(x+w),inch(3.75),inch(x+w+0.3),inch(3.75),DORADO,19050); step.append((ln,'fade'))
        steps.append(step); x+=w+0.3
    fid=peek_next(); s+=textbox(inch(0.6),inch(5.0),inch(12),inch(0.6),[para([run("DNED lidera diagnóstico–validación · Mesa Técnica domina revisión–implementación",1250,GRIS)])])
    add_slide(s, anims=steps+[[(fid,'fade')]], no_transition=True)  # Sección 6 · Plan

# ---- 13. Avance (GRAFICO dona) ----
def s_avance():
    s=chrome("Estado de avance del ecosistema","Indicadores")
    s+=graphicframe("rId2",inch(0.6),inch(1.4),inch(6.6),inch(4.6))
    kpis=[("14","Mesas de trabajo temáticas"),("11","Documentos entregados (79%)"),("3","Documentos pendientes"),("9","Sedes integradas")]
    y=1.6
    for n,l in kpis:
        s+=rect(inch(7.6),inch(y),inch(4.8),inch(0.95),BLANCO,line="e2e7e3",round_=True)
        s+=textbox(inch(7.8),inch(y+0.12),inch(1.5),inch(0.7),[para([run(n,2600,TEAL,b=True)])])
        s+=textbox(inch(9.3),inch(y+0.28),inch(3.0),inch(0.5),[para([run(l,1250,TXT)])])
        y+=1.08
    # grafico editable; sin animacion
    add_slide(s, charts=[("rId2","dough")])

# ---- 14. Cierre ----
def s_cierre():
    s=rect(0,0,PW,PH,VERDEPF)
    s+=rect(0,0,PW,inch(0.14),DORADO); s+=rect(0,PH-inch(0.14),PW,inch(0.14),NARANJA)
    g1=peek_next()
    s+=textbox(inch(0.8),inch(1.6),inch(11.7),inch(2.6),[
        para([run("Gracias",4000,DORADO,b=True)]),
        para([run("La gobernanza de la IA en la UNAL es un mecanismo de soberanía del",2000,BLANCO)]),
        para([run("conocimiento, equidad territorial y protección de derechos.",2000,BLANCO)])])
    g2=peek_next()
    s+=textbox(inch(0.8),inch(4.6),inch(11.7),inch(1.6),[
        para([run("Prof. Francisco Albeiro Gómez Jaramillo · fagomezj@unal.edu.co",1500,CREMA)]),
        para([run("Ing. Wilver Alexander Martínez Martínez — DNED",1500,CREMA)]),
        para([run("DNED-DIVULAB · Universidad Nacional de Colombia",1300,DORADO)])])
    add_slide(s, anims=[[(g1,'fade')],[(g2,'flyB')]])

for fn in [s_portada,s_apertura,s_problema,s_ecosistema,s_benchmarking,s_multinivel,
           s_nucleos,s_plenaria,s_votacion,s_riesgo,s_raci,s_ruta,s_avance,s_cierre]:
    fn()

# ======================================================================
#   EMPAQUETADO OOXML
# ======================================================================
NS=('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')

def slide_xml(shapes, timing="", no_transition=False):
    trans='' if no_transition else '<p:transition spd="fast"><p:fade/></p:transition>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p:sld {NS}><p:cSld><p:spTree>'
            f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            f'{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping '
            'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'
            f'{trans}{timing}</p:sld>')

z=zipfile.ZipFile("presentacion-dashboard-gobernanza-ia-unal.pptx","w",zipfile.ZIP_DEFLATED)

# [Content_Types]
overrides=['<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
 '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
 '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
 '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
 '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
 '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
chart_count=0
for idx,(sh,charts,anims,no_trans) in enumerate(slides,1):
    overrides.append(f'<Override PartName="/ppt/slides/slide{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    for rid,kind in charts:
        chart_count+=1
        overrides.append(f'<Override PartName="/ppt/charts/chart{chart_count}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
ct=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    +"".join(overrides)+'</Types>')
z.writestr("[Content_Types].xml",ct)

# _rels/.rels
z.writestr("_rels/.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
 '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
 '</Relationships>')

# docProps
z.writestr("docProps/core.xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'
 '<dc:title>Ponencia - Dashboard Gobernanza de la IA en la UNAL</dc:title><dc:creator>DNED-DIVULAB</dc:creator></cp:coreProperties>')
z.writestr("docProps/app.xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
 f'<Application>DNED-DIVULAB</Application><Slides>{len(slides)}</Slides></Properties>')

# presentation.xml + rels
sldIds="".join(f'<p:sldId id="{255+i}" r:id="rId{i+2}"/>' for i in range(len(slides)))
z.writestr("ppt/presentation.xml",f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 f'<p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
 f'<p:sldIdLst>{sldIds}</p:sldIdLst><p:sldSz cx="{PW}" cy="{PH}"/><p:notesSz cx="{PH}" cy="{PW}"/></p:presentation>')
prels=['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
for i in range(len(slides)):
    prels.append(f'<Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>')
prels.append(f'<Relationship Id="rId{len(slides)+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
z.writestr("ppt/_rels/presentation.xml.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+"".join(prels)+'</Relationships>')

# theme minimo
theme=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">'
 '<a:themeElements><a:clrScheme name="Office">'
 '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
 '<a:dk2><a:srgbClr val="1F4037"/></a:dk2><a:lt2><a:srgbClr val="F4F1E9"/></a:lt2>'
 '<a:accent1><a:srgbClr val="2962FF"/></a:accent1><a:accent2><a:srgbClr val="0F8C82"/></a:accent2>'
 '<a:accent3><a:srgbClr val="E8591C"/></a:accent3><a:accent4><a:srgbClr val="17A06A"/></a:accent4>'
 '<a:accent5><a:srgbClr val="B89347"/></a:accent5><a:accent6><a:srgbClr val="6B4FA0"/></a:accent6>'
 '<a:hlink><a:srgbClr val="2962FF"/></a:hlink><a:folHlink><a:srgbClr val="6B4FA0"/></a:folHlink></a:clrScheme>'
 '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
 '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
 '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
 '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
 '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
 '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
 '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
 '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
 '</a:fmtScheme></a:themeElements></a:theme>')
z.writestr("ppt/theme/theme1.xml",theme)

# slideMaster + layout (minimos)
master=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 f'<p:sldMaster {NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
 f'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
 f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
 f'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
 f'<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>')
z.writestr("ppt/slideMasters/slideMaster1.xml",master)
z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
 '</Relationships>')
layout=(f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 f'<p:sldLayout {NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>'
 f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
 f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
 f'<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sldLayout>')
z.writestr("ppt/slideLayouts/slideLayout1.xml",layout)
z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
 '</Relationships>')

# slides + sus rels + charts
cc=0
for idx,(sh,charts,anims,no_trans) in enumerate(slides,1):
    z.writestr(f"ppt/slides/slide{idx}.xml",slide_xml(sh, build_timing(anims), no_trans))
    rels=['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>']
    for rid,kind in charts:
        cc+=1
        rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{cc}.xml"/>')
        z.writestr(f"ppt/charts/chart{cc}.xml", bar_chart_xml() if kind=="bar" else doughnut_chart_xml())
    z.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+"".join(rels)+'</Relationships>')

z.close()
import os
print("PPTX OK:", os.path.getsize("presentacion-dashboard-gobernanza-ia-unal.pptx"),"bytes,",len(slides),"diapositivas, graficos:",cc)
