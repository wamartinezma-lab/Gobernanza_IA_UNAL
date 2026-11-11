#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera el brief / ponencia en dos formatos, sin dependencias externas:
  - brief-ponencia-gobernanza-ia-unal.pdf   (PDF con fuente TrueType incrustada -> tildes y ñ)
  - brief-ponencia-gobernanza-ia-unal.docx  (OOXML / Word)
Español de Latinoamérica, con ortografía correcta (tildes, ñ, signos ¿ ¡).
"""
import zlib, textwrap, zipfile, html, time, struct, os

# ---------------- Paleta ----------------
VERDE=(0.122,0.251,0.216); VERDEPF=(0.075,0.165,0.141); DORADO=(0.722,0.576,0.278)
NEGRO=(0.11,0.17,0.15); GRIST=(0.40,0.33,0.14); CREMA=(0.957,0.945,0.914); BLANCO=(1,1,1)
TXT=(0.17,0.22,0.20)
PW,PH=595.28,841.89

import os as _os
def _find_font(name):
    # busca la fuente en varias rutas conocidas (el entorno puede cambiar)
    for base in ("/usr/share/fonts/google-noto",
                 "/inner_container/usr/share/fonts/google-noto",
                 "/usr/share/fonts/truetype/noto"):
        p=_os.path.join(base,name)
        if _os.path.exists(p): return p
    return _os.path.join("/usr/share/fonts/google-noto",name)  # fallback
FONT_REG=_find_font("NotoSans-Regular.ttf")
FONT_BOLD=_find_font("NotoSans-Bold.ttf")

# ======================================================================
#   Contenido (acentos correctos)
# ======================================================================
TITULO="Hacia un ecosistema de gobernanza para el uso responsable y la apropiación ética de la IA en la UNAL"
HEADER_TTL="Ponencia"
HEADER_SUB="Propuesta de Gobernanza de la IA en la UNAL  ·  DNED-DIVULAB"
FOOTER="DNED-DIVULAB"

FICHA=[("Ponentes","Prof. Francisco Albeiro Gómez Jaramillo · Ing. Wilver Alexander Martínez Martínez"),
       ("Entidad","DNED · División Universidad Laboratorio (DIVULAB)"),
       ("Formato","Ponencia con apoyo visual interactivo"),
       ("Modalidad","Presencial / híbrida (requiere proyección)"),
       ("Nivel","Introductorio-intermedio (sin requisitos técnicos)"),
       ("Idioma","Español")]

ABSTRACT_P1=("La Universidad Nacional de Colombia adelanta un esfuerzo institucional de alcance nacional para definir un "
 "ecosistema de gobernanza que oriente el uso responsable y la apropiación ética de la IA en toda la comunidad "
 "universitaria. No es un comité técnico aislado, sino un proceso colectivo y multiestamentario —liderado por "
 "DNED-DIVULAB con la Mesa Técnica de IA— que articula lineamientos, estructuras de decisión, procesos y capacidades "
 "para que la IA sirva a las misiones de la Universidad y a la protección de los derechos de las personas.")
ABSTRACT_P2=("Ante una tecnología que ya incide en la docencia, la investigación, la extensión y la gestión, la pregunta "
 "determinante no es si usar la IA, sino quién decide qué se puede hacer, con qué datos y bajo qué límites. La "
 "ponencia presenta el componente de gobernanza de ese ecosistema: un esquema multinivel, participativo y multicampus "
 "que articula al CNED, una Mesa Técnica de IA con cuatro núcleos y a la comunidad de las nueve sedes, nutrido por 14 "
 "mesas de trabajo temáticas y un plan plurianual 2025-2027 de co-construcción, validación e implementación.")
ABSTRACT_P3=("A partir de un benchmarking internacional (MIT, UCLA, Stanford) ampliado con referentes del Sur Global "
 "(Tec de Monterrey, USP, UNAM, INTEGRITY-AI), se muestra cómo el ecosistema traduce principios —subsidiariedad, "
 "gobierno corporativo de la IA, riesgo graduado por derechos, soberanía del conocimiento y sostenibilidad— en un "
 "procedimiento operativo: una plenaria de delegados que recibe los casos, los traslada a trámite, delibera por quórum "
 "y emite una decisión colegiada. La charla se apoya en un gemelo digital 3D del modelo y un dashboard analítico.")
ABSTRACT_PARRS=[ABSTRACT_P1,ABSTRACT_P2,ABSTRACT_P3]

KW="ecosistema de gobernanza de IA, uso responsable, apropiación ética, educación superior, subsidiariedad, gestión de riesgo, soberanía del conocimiento, multicampus, Green AI."

CONTEXTO=["La IA generativa y agéntica ya está presente en aulas, laboratorios y procesos administrativos.",
 "Los marcos rígidos o excesivamente centralizados fragmentan la apropiación o limitan el impacto pedagógico y social.",
 "La UNAL —pública, nacional, pluriétnica y con nueve sedes— requiere un modelo propio que armonice eficiencia, ética, autonomía y compromisos socioambientales.",
 "La gobernanza de la IA supera al gobierno de TI: se ubica en el gobierno corporativo, con intersección a TI y datos."]

ECO_INTRO=("Más que una estructura, la UNAL impulsa un ecosistema que alinea estructuras de decisión, políticas, "
 "procesos y capacidades para el uso responsable y la apropiación ética de la IA. Sus rasgos distintivos:")
ECOSISTEMA=["Proceso colectivo y multiestamentario: la comunidad participa como co-creadora, no como destinataria pasiva.",
 "Ecosistema de lineamientos y guías: orientaciones éticas, pedagógicas, jurídicas y ambientales (integridad académica, datos, salud mental, género, diversidad cultural).",
 "14 mesas de trabajo temáticas que aportan conocimiento situado de la comunidad.",
 "Plan plurianual 2025-2027: diagnóstico, co-construcción, validación, aprobación e implementación.",
 "Carácter público, nacional, pluriétnico y multicampus: equidad territorial, soberanía del conocimiento y diálogo intercultural.",
 "Validación académica externa que retroalimenta y legitima el modelo."]
ECO_CIERRE="El propósito no es solo controlar la IA, sino habilitar su apropiación ética y responsable al servicio de las misiones de la Universidad y la protección de los derechos."

OBJ_GEN="Mostrar el esfuerzo institucional de la UNAL por definir un ecosistema de gobernanza para el uso responsable y la apropiación ética de la IA, y cómo opera en la práctica."
OBJ_ESP=["Situar el problema y contrastar tres modelos internacionales + referentes del Sur Global.",
 "Presentar el ecosistema institucional (proceso multiestamentario, lineamientos y plan 2025-2027).",
 "Explicar la arquitectura multinivel (CNED → Mesa Técnica → Comunidad) y sus cuatro núcleos.",
 "Mostrar el procedimiento operativo de la Mesa (plenaria, trámite, quórum y votación).",
 "Ilustrar la gestión diferenciada por tipo de sistema y nivel de riesgo (taxonomía, RACI, veto)."]

PUBLICO=("Autoridades y equipos universitarios, docentes, investigadores, personal administrativo, estudiantes y "
 "comunidad interesada en política institucional, ética y gestión de la IA. No requiere conocimientos técnicos avanzados.")

AGENDA=[("0","Apertura y gancho"),("1","El problema: por qué gobernar la IA"),
 ("2","Benchmarking (MIT, UCLA, Stanford + Sur Global)"),("3","El modelo UNAL: multinivel y 4 núcleos"),
 ("4","Cómo opera la Mesa: plenaria, quórum y votación"),("5","Riesgo, taxonomía y matriz RACI"),
 ("6","Recomendaciones y hoja de ruta 2025-2027"),("7","Cierre + preguntas")]

APORTES=["La UNAL construye un ecosistema —no solo una estructura— para el uso responsable y la apropiación ética de la IA.",
 "Es un esfuerzo colectivo y multiestamentario, con la comunidad como co-creadora.",
 "La gobernanza de la IA es un asunto institucional de derechos, misión y sostenibilidad, no solo técnico.",
 "Subsidiariedad: decidir en el nivel más cercano al riesgo y al valor; centralizar por excepción.",
 "El veto se gradúa por riesgo a derechos fundamentales, no por el dominio de aplicación.",
 "La operación colegiada (plenaria + quórum) hace trazable quién decide qué.",
 "El ecosistema integra referentes latinoamericanos y revisión académica externa."]

RESULTADOS=("Al finalizar, la audiencia podrá: comprender el ecosistema institucional y su propósito ético; distinguir "
 "los niveles del modelo y los tipos de sistemas de IA y su riesgo; identificar roles de decisión (RACI); y visualizar "
 "el recorrido operativo de una iniciativa de principio a fin.")

REQUISITOS=["Proyector o pantalla de proyección.","Salida de audio / sistema de sonido."]

PONENTES=[("Prof. Francisco Albeiro Gómez Jaramillo",
           "Coordinador de la División Universidad Laboratorio (DIVULAB), DNED, Universidad Nacional de Colombia. Correo: fagomezj@unal.edu.co"),
          ("Ing. Wilver Alexander Martínez Martínez",
           "Líder en la DNED de la iniciativa institucional de Gobernanza de la IA en la UNAL.")]
BIO=("La Dirección Nacional de Estrategia Digital (DNED), a través de la División Universidad Laboratorio (DIVULAB), "
 "orienta y articula la estrategia digital de la UNAL en sus nueve sedes, y lidera la co-construcción multiestamentaria "
 "del marco de uso ético y responsable de la IA, junto con la Mesa Técnica de IA. Su trabajo integra gobernanza "
 "institucional, acompañamiento pedagógico, sostenibilidad (Green AI) y soberanía del conocimiento, con enfoque "
 "participativo, territorial y pluriétnico.")

REFS=["Propuesta de creación de la Mesa Técnica de IA como sala adscrita al CNED (DNED).",
 "Presentación institucional Gobernanza IA (DNED-DIVULAB).",
 "Benchmarking: MIT Schwarzman College; UCLA / UC AI Council; Stanford HAI / AI Index; Tec de Monterrey; USP; UNAM; INTEGRITY-AI."]

# ======================================================================
#   Lector mínimo de TTF: cmap (unicode->gid) y hmtx (anchos)
# ======================================================================
class TTF:
    def __init__(self, path):
        self.data=open(path,'rb').read()
        self._parse()
    def u16(self,o): return struct.unpack(">H",self.data[o:o+2])[0]
    def s16(self,o): return struct.unpack(">h",self.data[o:o+2])[0]
    def u32(self,o): return struct.unpack(">I",self.data[o:o+4])[0]
    def _parse(self):
        numTables=self.u16(4); self.tables={}
        o=12
        for i in range(numTables):
            tag=self.data[o:o+4].decode('latin-1'); off=self.u32(o+8); ln=self.u32(o+12)
            self.tables[tag]=(off,ln); o+=16
        head=self.tables['head'][0]; self.unitsPerEm=self.u16(head+18)
        self.indexToLocFormat=self.s16(head+50)
        hhea=self.tables['hhea'][0]; self.numHMetrics=self.u16(hhea+34)
        maxp=self.tables['maxp'][0]; self.numGlyphs=self.u16(maxp+4)
        self._parse_hmtx(); self._parse_cmap()
    def _parse_hmtx(self):
        off=self.tables['hmtx'][0]; self.aw=[]
        last=0
        for i in range(self.numGlyphs):
            if i<self.numHMetrics:
                last=self.u16(off+i*4); self.aw.append(last)
            else:
                self.aw.append(last)
    def _parse_cmap(self):
        off=self.tables['cmap'][0]; n=self.u16(off+2); sub=None
        for i in range(n):
            pid=self.u16(off+4+i*8); eid=self.u16(off+6+i*8); so=self.u32(off+8+i*8)
            if (pid==3 and eid in (1,10)) or (pid==0):
                sub=off+so
        self.cmap={}
        if sub is None: return
        fmt=self.u16(sub)
        if fmt==4:
            segX2=self.u16(sub+6); segc=segX2//2
            endO=sub+14; startO=endO+segX2+2; deltaO=startO+segX2; rangeO=deltaO+segX2
            for s in range(segc):
                end=self.u16(endO+s*2); start=self.u16(startO+s*2)
                delta=self.u16(deltaO+s*2); ro=self.u16(rangeO+s*2)
                for c in range(start,end+1):
                    if c==0xFFFF: continue
                    if ro==0:
                        g=(c+delta)&0xFFFF
                    else:
                        gi=rangeO+s*2+ro+(c-start)*2
                        g=self.u16(gi)
                        if g!=0: g=(g+delta)&0xFFFF
                    if g: self.cmap[c]=g
        elif fmt==12:
            ng=self.u32(sub+12)
            for i in range(ng):
                b=sub+16+i*12
                sc=self.u32(b); ec=self.u32(b+4); sg=self.u32(b+8)
                for c in range(sc,ec+1): self.cmap[c]=sg+(c-sc)
    def gid(self,ch):
        return self.cmap.get(ord(ch),0)
    def width_em(self,ch):
        g=self.gid(ch)
        return self.aw[g]/self.unitsPerEm if g<len(self.aw) else 0.5

REG=TTF(FONT_REG); BLD=TTF(FONT_BOLD)

# ======================================================================
#   Motor PDF con fuentes TrueType incrustadas (Type0 / Identity-H)
# ======================================================================
class PDF:
    def __init__(self):
        self.objs=[]; self.pages=[]; self.cur=""; self.y=0
        self.used_reg=set(); self.used_bld=set()
    # --- medición ---
    def tw(self,t,size,bold=False):
        f=BLD if bold else REG
        return sum(f.width_em(c) for c in t)*size
    # --- primitivas ---
    def rect(self,x,y,w,h,fill=None,stroke=None,lw=1):
        s=""
        if fill:s+=f"{fill[0]:.3f} {fill[1]:.3f} {fill[2]:.3f} rg\n"
        if stroke:s+=f"{stroke[0]:.3f} {stroke[1]:.3f} {stroke[2]:.3f} RG\n{lw} w\n"
        s+=f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re\n"
        s+=("B\n" if fill and stroke else "f\n" if fill else "S\n"); return s
    def _hexstr(self,t,bold):
        f=BLD if bold else REG
        used=self.used_bld if bold else self.used_reg
        out=""
        for ch in t:
            g=f.gid(ch); used.add(ch)
            out+=f"{g:04X}"
        return out
    def txt(self,x,y,t,size=10,color=NEGRO,bold=False,wordspace=0.0):
        fn="FB" if bold else "FR"
        # wordspace via posicionamiento manual no aplica en Identity-H con Tw;
        # para justificar insertamos ajustes con TJ.
        hexs=self._hexstr(t,bold)
        s=(f"BT /{fn} {size} Tf {color[0]:.3f} {color[1]:.3f} {color[2]:.3f} rg "
           f"{x:.2f} {y:.2f} Td <{hexs}> Tj ET\n")
        return s
    def txt_justify(self,x,y,words,size,color,bold,width):
        # construye un TJ con espacios ensanchados entre palabras
        f=BLD if bold else REG
        fn="FB" if bold else "FR"
        natural=sum(self.tw(w,size,bold) for w in words)
        spacew=f.width_em(' ')*size
        gaps=len(words)-1
        total_space=width-natural
        if gaps>0 and total_space>spacew*gaps:
            per=total_space/gaps
        else:
            per=spacew
        # per en unidades de texto; en TJ el ajuste se da en milésimas de em (restando)
        # ancho de un espacio deseado = per; equivalente glifo espacio = spacew.
        # Insertamos el glifo espacio y un ajuste negativo para ampliar: adj = -(per-spacew)/size*1000
        adj=-(per-spacew)/size*1000.0
        parts=[]
        sgid=f.gid(' ')
        for i,w in enumerate(words):
            parts.append("<"+self._hexstr(w,bold)+">")
            if i<len(words)-1:
                parts.append("<%04X>"%sgid)
                parts.append(f"{adj:.1f}")
        tj="[ "+" ".join(parts)+" ] TJ"
        return (f"BT /{fn} {size} Tf {color[0]:.3f} {color[1]:.3f} {color[2]:.3f} rg "
                f"{x:.2f} {y:.2f} Td {tj} ET\n")
    def esc(self,t): return t  # no usado (hex)
    # --- layout ---
    def new_page(self):
        if self.cur: self.pages.append(self.cur)
        self.cur=""
        self.cur+=self.rect(0,0,PW,PH,CREMA)
        self.cur+=self.rect(0,PH-54,PW,54,VERDEPF)+self.rect(0,PH-54,6,54,DORADO)
        self.cur+=self.txt(40,PH-26,HEADER_TTL,16,BLANCO,bold=True)
        self.cur+=self.txt(40,PH-44,HEADER_SUB,9.5,(0.78,0.84,0.80))
        self.cur+=self.rect(0,0,PW,20,VERDE)
        self.cur+=self.txt(40,7,FOOTER,8.5,CREMA)
        self.y=PH-78
    def ensure(self,need):
        if self.y-need<34: self.new_page()
    def h2(self,t):
        self.ensure(26)
        self.cur+=self.txt(40,self.y,t,12.5,VERDE,bold=True); self.y-=8
        self.cur+=self.rect(40,self.y,120,2,DORADO); self.y-=14
    def para(self,t,size=9.6,color=TXT,chars=104,lead=12.6,indent=40,justify=True,width=PW-80):
        lines=textwrap.wrap(t,chars); 
        for i,ln in enumerate(lines):
            self.ensure(lead); last=(i==len(lines)-1); words=ln.split(" ")
            avail=width-(indent-40)
            if justify and not last and len(words)>1:
                self.cur+=self.txt_justify(indent,self.y,words,size,color,False,avail)
            else:
                self.cur+=self.txt(indent,self.y,ln,size,color)
            self.y-=lead
        self.y-=3
    def bullet(self,t,size=9.6,chars=98):
        lines=textwrap.wrap(t,chars)
        for i,ln in enumerate(lines):
            self.ensure(12.6)
            if i==0: self.cur+=self.txt(44,self.y,"•",size,DORADO,bold=True)
            self.cur+=self.txt(56,self.y,ln,size,TXT); self.y-=12.6
    def kv(self,k,v):
        self.ensure(14)
        self.cur+=self.txt(44,self.y,k+":",9.4,GRIST,bold=True)
        self.cur+=self.txt(150,self.y,v,9.4,TXT); self.y-=14
    # --- embedding ---
    def _font_objs(self, ttf, used, basefont):
        # subconjunto no: incrustamos la fuente completa (simple y robusto)
        fontfile=len(self.objs)+1; self.objs.append(("S",zlib.compress(ttf.data)))
        # W array (anchos por GID usados), CIDToGID=Identity => CID=GID
        widths={}
        for ch in used:
            g=ttf.gid(ch); widths[g]=round(ttf.aw[g]*1000/ttf.unitsPerEm)
        wparts=[]
        for g in sorted(widths):
            wparts.append(f"{g} [{widths[g]}]")
        warr="[ "+" ".join(wparts)+" ]"
        descr=len(self.objs)+1
        self.objs.append(
            f"<< /Type /FontDescriptor /FontName /{basefont} /Flags 32 "
            f"/FontBBox [-600 -400 1500 1200] /ItalicAngle 0 /Ascent 1069 /Descent -293 "
            f"/CapHeight 714 /StemV 90 /FontFile2 {fontfile} 0 R >>")
        cidfont=len(self.objs)+1
        self.objs.append(
            f"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{basefont} "
            f"/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> "
            f"/FontDescriptor {descr} 0 R /CIDToGIDMap /Identity /DW 500 /W {warr} >>")
        type0=len(self.objs)+1
        self.objs.append(
            f"<< /Type /Font /Subtype /Type0 /BaseFont /{basefont} /Encoding /Identity-H "
            f"/DescendantFonts [{cidfont} 0 R] >>")
        return type0
    def build(self):
        if self.cur:self.pages.append(self.cur)
        fr=self._font_objs(REG,self.used_reg,"NotoSans")
        fb=self._font_objs(BLD,self.used_bld,"NotoSans-Bold")
        cids=[]
        for c in self.pages:
            cids.append(len(self.objs)+1); self.objs.append(("S",zlib.compress(c.encode('latin-1','replace'))))
        parent=len(self.objs)+len(self.pages)+1; pids=[]
        for cid in cids:
            pids.append(len(self.objs)+1)
            self.objs.append(f"<< /Type /Page /Parent {parent} 0 R /MediaBox [0 0 {PW:.2f} {PH:.2f}] "
                             f"/Resources << /Font << /FR {fr} 0 R /FB {fb} 0 R >> >> /Contents {cid} 0 R >>")
        po=len(self.objs)+1; self.objs.append("<< /Type /Pages /Kids ["+" ".join(f"{p} 0 R" for p in pids)+f"] /Count {len(pids)} >>")
        cat=len(self.objs)+1; self.objs.append(f"<< /Type /Catalog /Pages {po} 0 R >>")
        out=b"%PDF-1.6\n%\xe2\xe3\xcf\xd3\n"; offs=[]
        for i,o in enumerate(self.objs,1):
            offs.append(len(out))
            if isinstance(o,tuple):
                out+=f"{i} 0 obj\n<< /Length {len(o[1])} /Filter /FlateDecode >>\nstream\n".encode('latin-1')+o[1]+b"\nendstream\nendobj\n"
            else:
                out+=f"{i} 0 obj\n{o}\nendobj\n".encode('latin-1')
        x=len(out); out+=f"xref\n0 {len(self.objs)+1}\n".encode('latin-1')+b"0000000000 65535 f \n"
        for o in offs: out+=f"{o:010d} 00000 n \n".encode('latin-1')
        out+=f"trailer\n<< /Size {len(self.objs)+1} /Root {cat} 0 R >>\nstartxref\n{x}\n%%EOF".encode('latin-1')
        return out

def build_pdf():
    p=PDF(); p.new_page()
    for ln in textwrap.wrap(TITULO,66):
        p.ensure(18); p.cur+=p.txt(40,p.y,ln,13.5,VERDE,bold=True); p.y-=17
    p.y-=6
    p.h2("Ficha")
    for k,v in FICHA: p.kv(k,v)
    p.y-=4
    p.h2("Resumen (abstract)")
    for parr in ABSTRACT_PARRS: p.para(parr)
    p.ensure(14); p.cur+=p.txt(40,p.y,"Palabras clave: "+KW,8.6,GRIST); p.y-=16
    p.h2("1. Contexto y problema")
    for b in CONTEXTO: p.bullet(b)
    p.h2("2. El esfuerzo institucional: hacia un ecosistema de gobernanza")
    p.para(ECO_INTRO)
    for b in ECOSISTEMA: p.bullet(b)
    p.para(ECO_CIERRE,size=9.4,color=GRIST)
    p.h2("3. Objetivos")
    p.para("General: "+OBJ_GEN)
    for b in OBJ_ESP: p.bullet(b)
    p.h2("4. Público objetivo"); p.para(PUBLICO)
    p.h2("5. Agenda")
    for n,t in AGENDA:
        p.ensure(13); p.cur+=p.txt(44,p.y,f"Sección {n}",8.8,DORADO,bold=True)
        p.cur+=p.txt(120,p.y,t,9.2,TXT); p.y-=13.5
    p.y-=4
    p.h2("6. Mensajes clave / aportes")
    for b in APORTES: p.bullet(b)
    p.h2("7. Resultados esperados"); p.para(RESULTADOS)
    p.h2("8. Sobre los ponentes")
    for nom,desc in PONENTES:
        p.ensure(13); p.cur+=p.txt(44,p.y,nom,10,VERDE,bold=True); p.y-=13
        p.para(desc,indent=44)
    p.para(BIO)
    p.h2("9. Referencias")
    for b in REFS: p.bullet(b)
    data=p.build()
    open("brief-ponencia-gobernanza-ia-unal.pdf","wb").write(data)
    return len(data),len(p.pages)

# ======================================================================
#   DOCX (OOXML) — UTF-8, acentos nativos
# ======================================================================
def xe(t): return html.escape(t,quote=True)
def run(t,bold=False,size=22,color="202b26"):
    rpr=f'<w:rPr>{"<w:b/>" if bold else ""}<w:sz w:val="{size}"/><w:szCs w:val="{size}"/><w:color w:val="{color}"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>'
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xe(t)}</w:t></w:r>'
def para(runs,space_after=120,ind=0,align="both"):
    ppr='<w:pPr>'
    if align: ppr+=f'<w:jc w:val="{align}"/>'
    if ind: ppr+=f'<w:ind w:left="{ind}"/>'
    ppr+=f'<w:spacing w:after="{space_after}" w:line="276" w:lineRule="auto"/></w:pPr>'
    return f'<w:p>{ppr}{"".join(runs)}</w:p>'
def heading(t):
    return (f'<w:p><w:pPr><w:spacing w:before="200" w:after="80"/>'
            f'<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="b89347"/></w:pBdr></w:pPr>'
            f'{run(t,bold=True,size=26,color="1f4037")}</w:p>')
def bullet(t):
    return para([run("•  ",size=21,color="b89347"),run(t,size=21)],space_after=60,ind=360,align="left")
def kvrow(k,v):
    return para([run(k+":  ",bold=True,size=21,color="7a6321"),run(v,size=21)],space_after=40,align="left")

def build_docx():
    body=[]
    body.append(para([run(TITULO,bold=True,size=30,color="1f4037")],space_after=60,align="left"))
    body.append(para([run("Ponencia · Propuesta de Gobernanza de la IA en la UNAL · DNED-DIVULAB",size=19,color="7a6321")],space_after=200,align="left"))
    body.append(heading("Ficha"))
    for k,v in FICHA: body.append(kvrow(k,v))
    body.append(heading("Resumen (abstract)"))
    for parr in ABSTRACT_PARRS: body.append(para([run(parr,size=21)]))
    body.append(para([run("Palabras clave: ",bold=True,size=20,color="7a6321"),run(KW,size=20)]))
    body.append(heading("1. Contexto y problema"))
    for b in CONTEXTO: body.append(bullet(b))
    body.append(heading("2. El esfuerzo institucional: hacia un ecosistema de gobernanza"))
    body.append(para([run(ECO_INTRO,size=21)]))
    for b in ECOSISTEMA: body.append(bullet(b))
    body.append(para([run(ECO_CIERRE,bold=True,size=21,color="1f4037")],ind=360))
    body.append(heading("3. Objetivos"))
    body.append(para([run("General: ",bold=True,size=21),run(OBJ_GEN,size=21)]))
    for b in OBJ_ESP: body.append(bullet(b))
    body.append(heading("4. Público objetivo")); body.append(para([run(PUBLICO,size=21)]))
    body.append(heading("5. Agenda"))
    for n,t in AGENDA:
        body.append(para([run(f"Sección {n}  ",bold=True,size=21,color="7a6321"),run(t,size=21)],space_after=40,align="left"))
    body.append(heading("6. Mensajes clave / aportes"))
    for b in APORTES: body.append(bullet(b))
    body.append(heading("7. Resultados esperados")); body.append(para([run(RESULTADOS,size=21)]))
    body.append(heading("8. Sobre los ponentes"))
    for nom,desc in PONENTES:
        body.append(para([run(nom+" — ",bold=True,size=21,color="1f4037"),run(desc,size=21)],space_after=80))
    body.append(para([run(BIO,size=21)]))
    body.append(heading("9. Referencias"))
    for b in REFS: body.append(bullet(b))

    document=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
      +''.join(body)+
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
      '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
      '</w:sectPr></w:body></w:document>')
    content_types=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')
    rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
    now=time.strftime("%Y-%m-%dT%H:%M:%SZ",time.gmtime())
    core=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
      '<dc:title>Ponencia - Propuesta de Gobernanza de la IA en la UNAL</dc:title>'
      '<dc:creator>DNED-DIVULAB - Universidad Nacional de Colombia</dc:creator>'
      f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
      f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
    app=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
      '<Application>DNED-DIVULAB</Application></Properties>')
    word_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    with zipfile.ZipFile("brief-ponencia-gobernanza-ia-unal.docx","w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",content_types)
        z.writestr("_rels/.rels",rels)
        z.writestr("word/document.xml",document)
        z.writestr("word/_rels/document.xml.rels",word_rels)
        z.writestr("docProps/core.xml",core)
        z.writestr("docProps/app.xml",app)
    return os.path.getsize("brief-ponencia-gobernanza-ia-unal.docx")

if __name__=="__main__":
    n,pg=build_pdf(); print(f"PDF  OK: {n} bytes, {pg} paginas (fuente TTF incrustada)")
    d=build_docx();   print(f"DOCX OK: {d} bytes")
