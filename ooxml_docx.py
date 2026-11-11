#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ooxml_docx.py — Toolkit mínimo para construir documentos Word (.docx) a mano
(OOXML con zipfile, sin dependencias externas).

Provee una clase Doc que acumula el cuerpo del documento, registra gráficos
nativos (c:chart editables), hipervínculos externos y diagramas DrawingML
(formas + conectores) e implementa el empaquetado completo del .docx.

Reutilizado por los generadores de documentos del Modelo de Gobernanza de la IA
en la UNAL (DNED–DIVULAB).
"""
import zipfile, html, time, os

def xe(t): return html.escape(str(t), quote=True)

# ---------- paleta institucional (hex sin #) ----------
VERDE="1F4037"; DORADO="7A6321"; DORADO2="B89347"; AZUL="2962FF"; TEAL="0F8C82"
CONTROL="E8591C"; SOCIO="17A06A"; CNED="1B355B"; GRIS="55615B"; ROJO="C0392B"
PUR="6B4FA0"; NEGRO="202B26"; AMBAR="C9A227"; CELESTE="4FA3C7"
EMU=914400

# ====================================================================
#   Primitivas de texto
# ====================================================================
def run(t, b=False, i=False, sz=22, color=NEGRO, font="Calibri"):
    rpr=(f'<w:rPr>{"<w:b/>" if b else ""}{"<w:i/>" if i else ""}'
         f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:color w:val="{color}"/>'
         f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/></w:rPr>')
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xe(t)}</w:t></w:r>'

def para(runs, align="both", after=120, before=0, ind=0, line=276):
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
    return para([run(t,b=True,sz=24,color=TEAL)], align="left", before=200, after=80)

def h3(t):
    return para([run(t,b=True,sz=22,color=DORADO)], align="left", before=140, after=60)

def bullet(t, color=NEGRO):
    if isinstance(t,list):
        return para([run("•  ",b=True,sz=22,color=DORADO2)]+t, align="left", after=60, ind=340)
    return para([run("•  ",b=True,sz=22,color=DORADO2), run(t,sz=22,color=color)], align="left", after=60, ind=340)

def numitem(n, t, color=NEGRO):
    return para([run(f"{n}.  ",b=True,sz=22,color=TEAL), run(t,sz=22,color=color)], align="left", after=60, ind=340)

def pagebreak():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def spacer(after=120):
    return para([run("")], after=after)

def nota(t):
    return para([run(t,i=True,sz=18,color=GRIS)], align="left", before=80, after=120)

# ---- caption (pie de figura / tabla) con numeración global ----
_fign=[0]
def caption(tipo, texto):
    _fign[0]+=1
    return para([run(f"{tipo} {_fign[0]}. ",b=True,sz=18,color=DORADO),
                 run(texto, i=True, sz=18, color=GRIS)], align="left", before=40, after=160)

def reset_captions(): _fign[0]=0

# ====================================================================
#   Tablas
# ====================================================================
def _cell(txt_runs, w, fill=None, align="left"):
    shd=f'<w:shd w:val="clear" w:fill="{fill}"/>' if fill else ''
    tcpr=f'<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}<w:vAlign w:val="center"/></w:tcPr>'
    if isinstance(txt_runs,str):
        txt_runs=[run(txt_runs,sz=20)]
    body=(f'<w:p><w:pPr><w:jc w:val="{align}"/>'
          f'<w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>'
          f'{"".join(txt_runs)}</w:p>')
    return f'<w:tc>{tcpr}{body}</w:tc>'

def table(rows, widths, header_fill=VERDE, zebra=None):
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
            if ri==0:
                fill=o.get('fill', header_fill); color=o.get('color',"FFFFFF"); bold=o.get('bold',True)
            else:
                fill=o.get('fill', (zebra if (zebra and ri%2==0) else None))
                color=o.get('color', NEGRO); bold=o.get('bold', False)
            align=o.get('align','left'); sz=o.get('sz',20)
            runs=content if isinstance(content,list) else [run(content,b=bold,sz=sz,color=color)]
            tcs+=_cell(runs, widths[ci], fill=fill, align=align)
        out+=f'<w:tr>{rpr}{tcs}</w:tr>'
    out+='</w:tbl>'
    return out

# ====================================================================
#   Clase Doc — acumula cuerpo, gráficos, hipervínculos y diagramas
# ====================================================================
class Doc:
    def __init__(self, title, creator="DNED-DIVULAB - Universidad Nacional de Colombia"):
        self.title=title; self.creator=creator
        self.body=[]
        self._charts=[]      # [(filename, xml)]
        self._hyperlinks=[]  # [(rid, url)]
        self._hl_n=0

    # ---- API de contenido ----
    def add(self, *xml):
        for x in xml: self.body.append(x)
        return self

    # ---- hipervínculo externo ----
    def hyperlink(self, url, text):
        self._hl_n+=1
        rid=f"rIdLink{self._hl_n}"
        self._hyperlinks.append((rid,url))
        r=(f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="{AZUL}"/><w:u w:val="single"/>'
           f'<w:sz w:val="22"/><w:szCs w:val="22"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>'
           f'<w:t xml:space="preserve">{xe(text)}</w:t></w:r>')
        return (f'<w:hyperlink r:id="{rid}" '
                f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{r}</w:hyperlink>')

    # ---- gráfico nativo (editable) ----
    def chart(self, kind, cats, vals, colors, titulo, w_in=6.0, h_in=3.1):
        idx=len(self._charts)+1
        fname=f"charts/chart{idx}.xml"
        if kind=="bar":
            xml=self._bar_chart_xml(cats,vals,colors,titulo,bar_dir="bar")
        elif kind=="col":
            xml=self._bar_chart_xml(cats,vals,colors,titulo,bar_dir="col")
        elif kind=="pie":
            xml=self._pie_chart_xml(cats,vals,colors,titulo)
        else:
            xml=self._doughnut_chart_xml(cats,vals,colors,titulo)
        self._charts.append((fname,xml))
        rid=f"rIdChart{idx}"
        cx=int(w_in*EMU); cy=int(h_in*EMU); docpr_id=1000+idx
        drawing=(f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
          f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
          f'<wp:docPr id="{docpr_id}" name="Grafico{idx}"/><wp:cNvGraphicFramePr/>'
          '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
          '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
          f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
          f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/>'
          '</a:graphicData></a:graphic></wp:inline></w:drawing>')
        return (f'<w:p><w:pPr><w:jc w:val="center"/>'
                f'<w:spacing w:before="60" w:after="60"/></w:pPr><w:r>{drawing}</w:r></w:p>')

    def _bar_chart_xml(self, cats, vals, colors, titulo, bar_dir="bar"):
        pts="".join(f'<c:pt idx="{i}"><c:v>{xe(c)}</c:v></c:pt>' for i,c in enumerate(cats))
        vpts="".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
        dpts="".join(f'<c:dPt><c:idx val="{i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>'
                     f'<c:spPr><a:solidFill><a:srgbClr val="{colors[i%len(colors)]}"/></a:solidFill></c:spPr></c:dPt>'
                     for i in range(len(cats)))
        catpos="l" if bar_dir=="bar" else "b"; valpos="b" if bar_dir=="bar" else "l"
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>'
          f'<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="es-CO" sz="1100" b="1"/>'
          f'<a:t>{xe(titulo)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
          '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>'
          f'<c:barChart><c:barDir val="{bar_dir}"/><c:grouping val="clustered"/><c:varyColors val="1"/>'
          '<c:ser><c:idx val="0"/><c:order val="0"/>'
          '<c:tx><c:strRef><c:f>Hoja1!$B$1</c:f><c:strCache><c:ptCount val="1"/>'
          '<c:pt idx="0"><c:v>Valor</c:v></c:pt></c:strCache></c:strRef></c:tx>'
          f'{dpts}'
          '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
          '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
          f'<c:cat><c:strRef><c:f>Hoja1!$A$2:$A${len(cats)+1}</c:f><c:strCache><c:ptCount val="{len(cats)}"/>{pts}</c:strCache></c:strRef></c:cat>'
          f'<c:val><c:numRef><c:f>Hoja1!$B$2:$B${len(vals)+1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>{vpts}</c:numCache></c:numRef></c:val>'
          '</c:ser><c:axId val="111"/><c:axId val="222"/></c:barChart>'
          f'<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{catpos}"/><c:crossAx val="222"/></c:catAx>'
          f'<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{valpos}"/><c:crossAx val="111"/></c:valAx>'
          '</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')

    def _doughnut_chart_xml(self, cats, vals, colors, titulo):
        return self._pie_like(cats,vals,colors,titulo,"doughnutChart",'<c:firstSliceAng val="0"/><c:holeSize val="55"/>')

    def _pie_chart_xml(self, cats, vals, colors, titulo):
        return self._pie_like(cats,vals,colors,titulo,"pieChart",'<c:firstSliceAng val="0"/>')

    def _pie_like(self, cats, vals, colors, titulo, tag, tail):
        pts="".join(f'<c:pt idx="{i}"><c:v>{xe(c)}</c:v></c:pt>' for i,c in enumerate(cats))
        vpts="".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i,v in enumerate(vals))
        dpts="".join(f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/>'
                     f'<c:spPr><a:solidFill><a:srgbClr val="{colors[i%len(colors)]}"/></a:solidFill></c:spPr></c:dPt>'
                     for i in range(len(cats)))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>'
          f'<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="es-CO" sz="1100" b="1"/>'
          f'<a:t>{xe(titulo)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
          '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>'
          f'<c:{tag}><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>'
          f'{dpts}'
          '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
          '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
          f'<c:cat><c:strRef><c:f>Hoja1!$A$2:$A${len(cats)+1}</c:f><c:strCache><c:ptCount val="{len(cats)}"/>{pts}</c:strCache></c:strRef></c:cat>'
          f'<c:val><c:numRef><c:f>Hoja1!$B$2:$B${len(vals)+1}</c:f><c:numCache><c:ptCount val="{len(vals)}"/>{vpts}</c:numCache></c:numRef></c:val>'
          f'</c:ser>{tail}</c:{tag}>'
          '</c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')

    # ---- diagrama DrawingML profesional: formas (cajas) + conectores ----
    @staticmethod
    def _shade(hexc, factor):
        """Aclara (factor>1) u oscurece (factor<1) un color hex para gradientes."""
        hexc=hexc.lstrip('#')
        r=int(hexc[0:2],16); g=int(hexc[2:4],16); b=int(hexc[4:6],16)
        if factor>=1:
            r=int(r+(255-r)*(factor-1)); g=int(g+(255-g)*(factor-1)); b=int(b+(255-b)*(factor-1))
        else:
            r=int(r*factor); g=int(g*factor); b=int(b*factor)
        clamp=lambda v: max(0,min(255,v))
        return f"{clamp(r):02X}{clamp(g):02X}{clamp(b):02X}"

    def diagram(self, boxes, arrows, w_in=6.3, h_in=3.4, title=None,
                panel=True, panel_fill="F6F8F7", panel_line="E2E7E4"):
        """
        Diagrama vectorial inline con formas y conectores (nativo, 100% editable en Word).
          boxes:  [{x,y,w,h,text,fill,line,tcolor,shape,fs,bold,shadow,grad}]
                  coords 0..1000 x, 0..1000 y (lienzo lógico).
          arrows: [(x1,y1,x2,y2[,color[,opts]])]  coords 0..1000.
                  opts (dict): {label, dash, head, tail, width, color}
                    head/tail: 'triangle'|'stealth'|'arrow'|'oval'|'none'
                    dash: True para línea punteada.
          panel: dibuja un panel de fondo sutil con sombra (contenedor).
        """
        cx=int(w_in*EMU); cy=int(h_in*EMU)
        def X(v): return int(v/1000*cx)
        def Y(v): return int(v/1000*cy)
        shapes=[]; sid=[1]
        def nid():
            sid[0]+=1; return sid[0]

        # ---- panel de fondo (contenedor con sombra suave) ----
        if panel:
            pad=X(14)
            shapes.append(
              f'<wps:wsp><wps:cNvPr id="{nid()}" name="panel"/><wps:cNvSpPr/>'
              f'<wps:spPr><a:xfrm><a:off x="{-pad}" y="{-pad}"/>'
              f'<a:ext cx="{cx+2*pad}" cy="{cy+2*pad}"/></a:xfrm>'
              f'<a:prstGeom prst="roundRect"><a:avLst>'
              f'<a:gd name="adj" fmla="val 4000"/></a:avLst></a:prstGeom>'
              f'<a:solidFill><a:srgbClr val="{panel_fill}"/></a:solidFill>'
              f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{panel_line}"/></a:solidFill></a:ln>'
              f'<a:effectLst><a:outerShdw blurRad="60000" dist="25000" dir="5400000" rotWithShape="0">'
              f'<a:srgbClr val="9AA7A0"><a:alpha val="30000"/></a:srgbClr></a:outerShdw></a:effectLst>'
              f'</wps:spPr><wps:bodyPr/></wps:wsp>')

        # ---- cajas ----
        for b in boxes:
            x=X(b['x']); y=Y(b['y']); w=X(b['w']); h=Y(b['h'])
            fill=b.get('fill',TEAL); tcolor=b.get('tcolor','FFFFFF')
            line=b.get('line', self._shade(fill,0.78))
            shp=b.get('shape','roundRect'); fs=b.get('fs',900); bold=b.get('bold',True)
            shadow=b.get('shadow',True); grad=b.get('grad',True)
            # relleno: gradiente diagonal (claro->base) para dar volumen, o sólido
            if grad:
                top=self._shade(fill,1.18); bot=self._shade(fill,0.9)
                fillxml=(f'<a:gradFill rotWithShape="1"><a:gsLst>'
                  f'<a:gs pos="0"><a:srgbClr val="{top}"/></a:gs>'
                  f'<a:gs pos="100000"><a:srgbClr val="{bot}"/></a:gs></a:gsLst>'
                  f'<a:lin ang="5400000" scaled="0"/></a:gradFill>')
            else:
                fillxml=f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
            shadowxml=('<a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" '
              'rotWithShape="0"><a:srgbClr val="5A6660"><a:alpha val="38000"/></a:srgbClr>'
              '</a:outerShdw></a:effectLst>') if shadow else ''
            avlst='<a:gd name="adj" fmla="val 14000"/>' if shp=='roundRect' else ''
            lines=b['text'].split("\n")
            tp=""
            for li,ln in enumerate(lines):
                emph = bold and li==0
                sz = fs if li==0 else int(fs*0.82)
                col = tcolor
                sp = ' spc="20"' if emph else ''
                tp+=(f'<a:p><a:pPr algn="ctr"/><a:r>'
                     f'<a:rPr lang="es-CO" sz="{sz}" b="{1 if emph else 0}"{sp}>'
                     f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
                     f'<a:latin typeface="Calibri"/></a:rPr>'
                     f'<a:t>{xe(ln)}</a:t></a:r></a:p>')
            shapes.append(
              f'<wps:wsp><wps:cNvPr id="{nid()}" name="box"/><wps:cNvSpPr/>'
              f'<wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
              f'<a:prstGeom prst="{shp}"><a:avLst>{avlst}</a:avLst></a:prstGeom>'
              f'{fillxml}'
              f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
              f'{shadowxml}</wps:spPr>'
              f'<wps:txbx><w:txbxContent>{tp}</w:txbxContent></wps:txbx>'
              f'<wps:bodyPr rot="0" anchor="ctr" lIns="36000" rIns="36000" '
              f'tIns="18000" bIns="18000"/></wps:wsp>')

        # ---- conectores (flechas) ----
        _HEADS={'triangle','stealth','arrow','oval','none','diamond'}
        for ar in arrows:
            x1,y1,x2,y2=ar[0],ar[1],ar[2],ar[3]
            col = ar[4] if len(ar)>4 and isinstance(ar[4],str) else GRIS
            opts = {}
            for part in ar[4:]:
                if isinstance(part,dict): opts=part
            col = opts.get('color', col)
            width = opts.get('width', 15875)
            head = opts.get('head', 'none')
            tail = opts.get('tail', 'triangle')
            dash = opts.get('dash', False)
            label = opts.get('label')
            ox=X(min(x1,x2)); oy=Y(min(y1,y2))
            ew=X(abs(x2-x1)) or 1; eh=Y(abs(y2-y1)) or 1
            flipH="1" if x2<x1 else "0"; flipV="1" if y2<y1 else "0"
            headxml=f'<a:headEnd type="{head}" w="med" len="med"/>' if head in _HEADS and head!='none' else ''
            tailxml=f'<a:tailEnd type="{tail}" w="med" len="med"/>' if tail in _HEADS and tail!='none' else ''
            dashxml='<a:prstDash val="dash"/>' if dash else '<a:prstDash val="solid"/>'
            shapes.append(
              f'<wps:wsp><wps:cNvPr id="{nid()}" name="conector"/>'
              f'<wps:cNvCnPr/><wps:spPr><a:xfrm flipH="{flipH}" flipV="{flipV}">'
              f'<a:off x="{ox}" y="{oy}"/><a:ext cx="{ew}" cy="{eh}"/></a:xfrm>'
              f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
              f'<a:ln w="{width}" cap="rnd"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
              f'{dashxml}<a:round/>{headxml}{tailxml}</a:ln></wps:spPr>'
              f'<wps:bodyPr/></wps:wsp>')
            # etiqueta sobre el conector (caja pequeña sin borde, centrada)
            if label:
                mx=(x1+x2)/2; my=(y1+y2)/2
                lw=X(max(90,len(label)*22)); lh=Y(60)
                lx=X(mx)-lw//2; ly=Y(my)-lh//2
                shapes.append(
                  f'<wps:wsp><wps:cNvPr id="{nid()}" name="etiqueta"/><wps:cNvSpPr/>'
                  f'<wps:spPr><a:xfrm><a:off x="{lx}" y="{ly}"/><a:ext cx="{lw}" cy="{lh}"/></a:xfrm>'
                  f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst></a:prstGeom>'
                  f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
                  f'<a:ln w="6350"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill></a:ln></wps:spPr>'
                  f'<wps:txbx><w:txbxContent><a:p><a:pPr algn="ctr"/><a:r>'
                  f'<a:rPr lang="es-CO" sz="700" b="1"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
                  f'<a:latin typeface="Calibri"/></a:rPr><a:t>{xe(label)}</a:t></a:r></a:p></w:txbxContent></wps:txbx>'
                  f'<wps:bodyPr lIns="18000" rIns="18000" tIns="9000" bIns="9000" anchor="ctr"/></wps:wsp>')

        # ---- título opcional dentro del lienzo ----
        if title:
            shapes.append(
              f'<wps:wsp><wps:cNvPr id="{nid()}" name="titulo"/><wps:cNvSpPr/>'
              f'<wps:spPr><a:xfrm><a:off x="0" y="{Y(2)}"/><a:ext cx="{cx}" cy="{Y(70)}"/></a:xfrm>'
              f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
              f'<wps:txbx><w:txbxContent><a:p><a:pPr algn="ctr"/><a:r>'
              f'<a:rPr lang="es-CO" sz="950" b="1"><a:solidFill><a:srgbClr val="{VERDE}"/></a:solidFill>'
              f'<a:latin typeface="Calibri"/></a:rPr><a:t>{xe(title)}</a:t></a:r></a:p></w:txbxContent></wps:txbx>'
              f'<wps:bodyPr anchor="t"/></wps:wsp>')

        docpr_id=2000+len(self._charts)+self._hl_n+len(self.body)
        inner=(f'<wp:inline distT="0" distB="0" distL="0" distR="0" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
          f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
          f'<wp:docPr id="{docpr_id}" name="Diagrama"/><wp:cNvGraphicFramePr/>'
          '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
          '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">'
          '<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
          'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
          'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
          f'<a:ext cx="{cx}" cy="{cy}"/><a:chOff x="0" y="0"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></wpg:grpSpPr>'
          + "".join(shapes) +
          '</wpg:wgp></a:graphicData></a:graphic></wp:inline>')
        # Envoltura mc:AlternateContent para máxima compatibilidad entre versiones de Word.
        grp=('<mc:AlternateContent '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
          '<mc:Choice Requires="wpg" '
          'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">'
          f'<w:drawing>{inner}</w:drawing></mc:Choice>'
          f'<mc:Fallback><w:drawing>{inner}</w:drawing></mc:Fallback>'
          '</mc:AlternateContent>')
        return (f'<w:p><w:pPr><w:jc w:val="center"/>'
                f'<w:spacing w:before="80" w:after="80"/></w:pPr><w:r>{grp}</w:r></w:p>')

    # ---- empaquetado ----
    def save(self, path):
        document=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
          + ''.join(self.body)
          + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            'w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>')
        ct=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
         '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
         '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
         + "".join(f'<Override PartName="/word/{fn}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                   for fn,_ in self._charts)
         + '</Types>')
        rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
         '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
        now=time.strftime("%Y-%m-%dT%H:%M:%SZ",time.gmtime())
        core=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
         'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
         'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
         f'<dc:title>{xe(self.title)}</dc:title><dc:creator>{xe(self.creator)}</dc:creator>'
         f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
         f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
        app=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
         '<Application>DNED-DIVULAB</Application></Properties>')
        chart_rels="".join(
          f'<Relationship Id="rIdChart{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{fn}"/>'
          for i,(fn,_) in enumerate(self._charts))
        link_rels="".join(
          f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{xe(url)}" TargetMode="External"/>'
          for rid,url in self._hyperlinks)
        word_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         + chart_rels + link_rels + '</Relationships>')
        with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml",ct)
            z.writestr("_rels/.rels",rels)
            z.writestr("word/document.xml",document)
            z.writestr("word/_rels/document.xml.rels",word_rels)
            z.writestr("docProps/core.xml",core)
            z.writestr("docProps/app.xml",app)
            for fn,xml in self._charts:
                z.writestr(f"word/{fn}", xml)
        return document
