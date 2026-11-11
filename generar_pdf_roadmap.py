#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de PDF (sin dependencias externas) del Roadmap de exposición
"Modelo de Gobernanza de la IA en la UNAL" — incluye gráficos vectoriales.
Produce: roadmap-exposicion-gobernanza-ia-unal.pdf
"""
import zlib, textwrap

# ---------------- Paleta ----------------
VERDE   = (0.122, 0.251, 0.216)   # 1f4037
VERDEPF = (0.075, 0.165, 0.141)   # 132a24
DORADO  = (0.722, 0.576, 0.278)   # b89347
NARANJA = (0.910, 0.576, 0.247)   # e8933f
AZUL    = (0.161, 0.384, 1.0)     # 2962ff
TEAL    = (0.059, 0.549, 0.510)   # 0f8c82
CONTROL = (0.909, 0.349, 0.110)   # e8591c
SOCIO   = (0.090, 0.627, 0.416)   # 17a06a
CNED    = (0.106, 0.208, 0.357)   # 1b355b
GRIS    = (0.44, 0.47, 0.47)
BLANCO  = (1,1,1)
CREMA   = (0.957,0.945,0.914)
NEGRO   = (0.11,0.17,0.15)

PW, PH = 595.28, 841.89  # A4 vertical en puntos

# ---------------- Datos (bloques del roadmap) ----------------
BLOCKS = [
    (0,"Apertura y gancho",2, SOCIO,
     "La IA ya esta en aulas, investigacion y gestion. La pregunta no es SI usarla, sino QUIEN decide que, con que datos y con que limites.",
     ["Presentacion: DNED-DIVULAB y proposito de la charla.",
      "Gancho: un predictor de desercion suena inofensivo... pero decide sobre el futuro de personas.",
      "Promesa: en 30 min veran el modelo completo y una simulacion en vivo."],
     "Portada / titulo."),
    (1,"El problema: por que gobernar la IA",3, AZUL,
     "La gobernanza de la IA SUPERA al gobierno de TI: es un asunto institucional de derechos, mision y sostenibilidad.",
     ["Riesgos: caja negra, sesgos, fuga de informacion, perdida de supervision humana, huella de carbono.",
      "Gobernanza de IA dentro del gobierno corporativo, con interseccion a TI y datos (Rec. 3).",
      "Abarca las 3 misiones + dimension etica y ambiental."],
     "Dashboard > Gestion de riesgo > tarjeta Gobernanza corporativa de la IA."),
    (2,"Benchmarking: que hacen otras universidades",4, CNED,
     "Tres modelos de referencia, pero la UNAL debe mirar tambien al Sur Global.",
     ["MIT - descentralizado (SERC, programa semilla).",
      "UCLA - matricial-centralizado (UC AI Council, veto clinico HIPAA).",
      "Stanford - hub-and-spoke (HAI, AI Index, huella de carbono).",
      "Sur Global (Rec. 1): Tec de Monterrey, USP, UNAM e INTEGRITY-AI (U. del Cauca, U. San Buenaventura)."],
     "GEMELO 3D > boton Modelos de referencia (dos orbitas: Norte y Sur Global)."),
    (3,"El modelo UNAL: multinivel y 4 nucleos",5, TEAL,
     "Modelo multinivel, participativo y multicampus, guiado por las misiones.",
     ["Tres niveles: CNED (politica) > Mesa Tecnica de IA (ejecucion) > Comunidad (9 sedes, facultades, grupos).",
      "4 nucleos: 1.1 Estrategico Misional, 1.2 Tecnico & Green, 1.3 Control & Juridico-Fiscal, 1.4 Socio-Ambiental & Territorial.",
      "Subsidiariedad (Rec. 9): decidir en el nivel mas cercano al riesgo y al valor.",
      "Capacidades habilitantes y legitimidad (Rec. 10): gobernanza no es operacion."],
     "GEMELO 3D (vista orbital) > Dashboard > Estructura (arquitectura de decisiones)."),
    (4,"Como opera la Mesa: plenaria, quorum y votacion",5, DORADO,
     "La Mesa opera como organo colegiado: una Plenaria de Delegados (uno por nucleo) analiza, traslada a tramite y decide por quorum.",
     ["Flujo: radicacion en plenaria > traslado a tramite a los 4 nucleos > regreso con ponencia > deliberacion > dictamen.",
      "Interaccion 1.2 <-> 1.3: tecnico y juridico se consultan (datos/licencias <-> seguridad/infra).",
      "Quorum (>= 3 de 4) y votacion: a favor / con salvaguardas / objecion.",
      "Regla: una objecion de 1.3 o 1.4 por riesgo a derechos prevalece (veto / escalar al CNED)."],
     "Dashboard > Simulador > mini-diagrama animado; flecha 1.2<->1.3."),
    (5,"Riesgo, taxonomia y matriz RACI",4, CONTROL,
     "No todas las iniciativas de IA son iguales: se gobiernan por tipo y por riesgo a derechos.",
     ["Taxonomia (Rec. 5): administrativo, aula, investigacion, propio, generativo de terceros, agentico.",
      "Niveles de riesgo (Rec. 6): bajo / medio / alto / inaceptable.",
      "Veto por derechos (Rec. 4): desercion o becas pueden lesionar derechos tanto como un sistema clinico.",
      "Matriz RACI (Rec. 8): quien propone, evalua, autoriza, implementa, supervisa y suspende."],
     "Dashboard > Gestion de riesgo y > Roles (RACI)."),
    (6,"Demo en vivo del simulador",4, NARANJA,
     "Veamos el modelo FUNCIONANDO con un caso real.",
     ["1. Dashboard > Simulador.  2. Escenario 'Desercion estudiantil'.",
      "3. Riesgo calculado: Alto.  4. Correr la animacion: Facultad > Sede > Plenaria > 4 nucleos > votacion.",
      "5. 1.3 Juridico objeta por datos sensibles > objecion prevalente > escala al CNED.",
      "6. Veredicto final: Escalada al CNED."],
     "Dashboard > Simulador. Plan B: capturas de pantalla si falla la proyeccion."),
    (7,"Recomendaciones y hoja de ruta 2025-2027",2, SOCIO,
     "El modelo recoge revision academica y tiene un plan de implementacion.",
     ["Las 10 recomendaciones de la Prof. Sanchez Torres ya estan incorporadas.",
      "Plan 2025-2027: Diagnostico > Co-construccion > Validacion > Revision/aprobacion > Implementacion."],
     "Dashboard > Recomendaciones y > Procesos y ciclo (linea de tiempo)."),
    (8,"Cierre + preguntas",1, VERDE,
     "La gobernanza de la IA en la UNAL es un mecanismo de soberania del conocimiento, equidad territorial y proteccion de derechos.",
     ["Invitar a probar el gemelo 3D y el dashboard.",
      "Contacto DNED-DIVULAB.",
      "Abrir preguntas (reservar ~4 min)."],
     "Portada / contacto."),
]
TOTAL = sum(b[2] for b in BLOCKS)

# Fases del plan (para gráfico adicional)
PLAN = [("Diagnostico","2S 2025",SOCIO),("Co-construccion","2026",AZUL),
        ("Validacion","Oct 2026",CONTROL),("Revision/aprob.","26-27",(0.42,0.3,0.55)),
        ("Implementacion","2027",(0.29,0.23,0.56))]

# ============================================================
#  Mini motor PDF (content streams vectoriales)
# ============================================================
class PDF:
    def __init__(self):
        self.objs=[]; self.pages=[]
    def _add(self,data):
        self.objs.append(data); return len(self.objs)
    def add_page(self, content):
        self.pages.append(content)
    # --- primitivas de dibujo (devuelven string de operadores) ---
    @staticmethod
    def rect(x,y,w,h,fill=None,stroke=None,lw=1):
        s=""
        if fill: s+=f"{fill[0]:.3f} {fill[1]:.3f} {fill[2]:.3f} rg\n"
        if stroke: s+=f"{stroke[0]:.3f} {stroke[1]:.3f} {stroke[2]:.3f} RG\n{lw} w\n"
        s+=f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re\n"
        if fill and stroke: s+="B\n"
        elif fill: s+="f\n"
        elif stroke: s+="S\n"
        return s
    @staticmethod
    def line(x1,y1,x2,y2,color=(0,0,0),lw=1):
        return (f"{color[0]:.3f} {color[1]:.3f} {color[2]:.3f} RG\n{lw} w\n"
                f"{x1:.2f} {y1:.2f} m {x2:.2f} {y2:.2f} l S\n")
    @staticmethod
    def _esc(t): return t.replace("\\","\\\\").replace("(","\\(").replace(")","\\)")
    def text(self,x,y,t,size=11,color=(0,0,0),font="F1"):
        return (f"BT /{font} {size} Tf {color[0]:.3f} {color[1]:.3f} {color[2]:.3f} rg "
                f"{x:.2f} {y:.2f} Td ({self._esc(t)}) Tj ET\n")
    def text_center(self,cx,y,t,size=11,color=(0,0,0),font="F1"):
        w=self.text_width(t,size,font)
        return self.text(cx-w/2,y,t,size,color,font)
    @staticmethod
    def text_width(t,size,font="F1"):
        # aproximación: Helvetica ~0.52em, Bold ~0.56em
        k=0.56 if font=="F2" else 0.52
        return len(t)*size*k
    def wrap(self,x,y,t,size,color,maxchars,lead=None,font="F1"):
        lead=lead or size+3
        out=""; 
        for i,line in enumerate(textwrap.wrap(t,maxchars)):
            out+=self.text(x,y-i*lead,line,size,color,font)
        return out, y-(max(1,len(textwrap.wrap(t,maxchars))))*lead
    def build(self):
        # fuentes
        f1=self._add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        f2=self._add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")
        kids=[]; page_ids=[]
        content_ids=[]
        for c in self.pages:
            comp=zlib.compress(c.encode('latin-1','replace'))
            cid=len(self.objs)+1
            self.objs.append(("STREAM",comp))
            content_ids.append(cid)
        # páginas (se crean después para referenciar Parent)
        parent_id = len(self.objs)+len(self.pages)+1
        for i,cid in enumerate(content_ids):
            pid=self._add(f"<< /Type /Page /Parent {parent_id} 0 R /MediaBox [0 0 {PW:.2f} {PH:.2f}] "
                          f"/Resources << /Font << /F1 {f1} 0 R /F2 {f2} 0 R >> >> /Contents {cid} 0 R >>")
            page_ids.append(pid)
        pages_obj=self._add("<< /Type /Pages /Kids ["+" ".join(f"{p} 0 R" for p in page_ids)+
                            f"] /Count {len(page_ids)} >>")
        catalog=self._add(f"<< /Type /Catalog /Pages {pages_obj} 0 R >>")
        # serializar
        out=b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
        offsets=[]
        for i,o in enumerate(self.objs,1):
            offsets.append(len(out))
            if isinstance(o,tuple) and o[0]=="STREAM":
                data=o[1]
                out+=f"{i} 0 obj\n<< /Length {len(data)} /Filter /FlateDecode >>\nstream\n".encode('latin-1')
                out+=data+b"\nendstream\nendobj\n"
            else:
                out+=f"{i} 0 obj\n{o}\nendobj\n".encode('latin-1')
        xref=len(out)
        out+=f"xref\n0 {len(self.objs)+1}\n".encode('latin-1')
        out+=b"0000000000 65535 f \n"
        for off in offsets:
            out+=f"{off:010d} 00000 n \n".encode('latin-1')
        out+=f"trailer\n<< /Size {len(self.objs)+1} /Root {catalog} 0 R >>\nstartxref\n{xref}\n%%EOF".encode('latin-1')
        return out

pdf=PDF()
def topy(y):  # convierte y "desde arriba" a coordenadas PDF
    return PH-y

# ---------------- PÁGINA 1: Portada + gráficos ----------------
def page_cover():
    c=""
    # fondo
    c+=pdf.rect(0,0,PW,PH,fill=VERDEPF)
    # franja superior dorada
    c+=pdf.rect(0,PH-10,PW,10,fill=DORADO)
    c+=pdf.rect(0,0,PW,8,fill=NARANJA)
    # escudo (placeholder)
    c+=pdf.rect(60,PH-150,54,54,fill=VERDE,stroke=DORADO,lw=1.5)
    c+=pdf.text_center(87,PH-128,"UNAL",12,DORADO,"F2")
    # título
    c+=pdf.text(130,PH-120,"Roadmap de exposicion (30 min)",21,BLANCO,"F2")
    c+=pdf.text(130,PH-144,"Modelo de Gobernanza de la IA en la UNAL",14,DORADO,"F1")
    c+=pdf.text(60,PH-185,"Mesa Tecnica de IA  -  DNED-DIVULAB  -  Documento para envio",11,(0.74,0.80,0.77))
    c+=pdf.line(60,PH-200,PW-60,PH-200,DORADO,1)

    # ---- Gráfico 1: línea de tiempo proporcional (barra apilada) ----
    y0=PH-250
    c+=pdf.text(60,y0+14,"Linea de tiempo de la exposicion (proporcional a los minutos)",12,BLANCO,"F2")
    bx,bw=60,PW-120; bh=34; x=bx
    palette=[b[3] for b in BLOCKS]
    for i,b in enumerate(BLOCKS):
        seg=bw*b[2]/TOTAL
        c+=pdf.rect(x,y0-bh,seg,bh,fill=palette[i],stroke=VERDEPF,lw=1)
        # etiqueta de minutos dentro
        c+=pdf.text_center(x+seg/2,y0-bh/2-3,f"{b[2]}m",9,BLANCO,"F2")
        # número de bloque arriba
        c+=pdf.text_center(x+seg/2,y0+3,f"{b[0]}",8,DORADO,"F2")
        x+=seg
    c+=pdf.text(60,y0-bh-14,"0 min",8,(0.74,0.80,0.77))
    c+=pdf.text(PW-90,y0-bh-14,f"{TOTAL} min",8,(0.74,0.80,0.77))

    # ---- Gráfico 2: barras de minutos por bloque ----
    y1=y0-bh-70
    c+=pdf.text(60,y1+14,"Minutos por bloque",12,BLANCO,"F2")
    gx,gw=70,PW-150; gh=150; base=y1-gh
    maxm=max(b[2] for b in BLOCKS)
    colw=gw/len(BLOCKS)
    # eje
    c+=pdf.line(gx,base,gx+gw,base,(0.4,0.5,0.46),1)
    for i,b in enumerate(BLOCKS):
        bar_h=gh*b[2]/maxm
        bxx=gx+i*colw+colw*0.18
        bww=colw*0.64
        c+=pdf.rect(bxx,base,bww,bar_h,fill=palette[i])
        c+=pdf.text_center(bxx+bww/2,base+bar_h+4,f"{b[2]}",9,BLANCO,"F2")
        c+=pdf.text_center(bxx+bww/2,base-12,f"B{b[0]}",8,(0.74,0.80,0.77))
    # pie
    c+=pdf.text(60,70,"Contenido: 9 bloques - apertura, problema, benchmarking, modelo UNAL, operacion de la Mesa,",10,(0.80,0.85,0.82))
    c+=pdf.text(60,56,"riesgo/RACI, demo en vivo, recomendaciones/hoja de ruta y cierre.",10,(0.80,0.85,0.82))
    pdf.add_page(c)

# ---------------- PÁGINAS DE BLOQUES ----------------
def page_blocks(chunk, idx):
    c=""
    c+=pdf.rect(0,0,PW,PH,fill=CREMA)
    c+=pdf.rect(0,PH-46,PW,46,fill=VERDE)
    c+=pdf.rect(0,PH-46,6,46,fill=DORADO)
    c+=pdf.text(40,PH-30,f"Guion por bloques  ({idx})",15,BLANCO,"F2")
    y=PH-80
    for b in chunk:
        n,title,mins,color,key,items,screen=b
        # cabecera del bloque
        c+=pdf.rect(40,y-26,PW-80,26,fill=color)
        c+=pdf.text(50,y-18,f"Bloque {n}  -  {title}",12.5,BLANCO,"F2")
        c+=pdf.text(PW-110,y-18,f"{mins} min",11,BLANCO,"F2")
        y-=26
        # mensaje clave
        box_lines=textwrap.wrap("Mensaje clave: "+key, 92)
        bh=len(box_lines)*13+12
        c+=pdf.rect(40,y-bh,PW-80,bh,fill=BLANCO,stroke=(0.88,0.90,0.88),lw=0.8)
        c+=pdf.rect(40,y-bh,4,bh,fill=DORADO)
        ty=y-16
        for ln in box_lines:
            c+=pdf.text(52,ty,ln,10.5,NEGRO,"F2" if ln.startswith("Mensaje") else "F1")
            ty-=13
        y-=bh+8
        # items
        for it in items:
            lines=textwrap.wrap(it,96)
            c+=pdf.text(50,y-10,chr(149),10,color,"F2")  # bullet
            for j,ln in enumerate(lines):
                c+=pdf.text(62,y-10-j*12,ln,10,(0.17,0.22,0.20))
            y-=10+len(lines)*12+2
        # en pantalla
        slines=textwrap.wrap("En pantalla: "+screen,96)
        sh=len(slines)*12+10
        c+=pdf.rect(40,y-sh,PW-80,sh,fill=(0.96,0.93,0.86),stroke=DORADO,lw=0.7)
        sy=y-14
        for ln in slines:
            c+=pdf.text(52,sy,ln,9.5,(0.40,0.33,0.14),"F1")
            sy-=12
        y-=sh+18
    # pie de página
    c+=pdf.rect(0,0,PW,22,fill=VERDE)
    c+=pdf.text(40,8,"Modelo de Gobernanza de la IA en la UNAL  -  Mesa Tecnica de IA  -  DNED-DIVULAB",8.5,CREMA)
    pdf.add_page(c)

# ---------------- PÁGINA FINAL: flujo demo + plan ----------------
def page_demo_and_plan():
    c=""
    c+=pdf.rect(0,0,PW,PH,fill=CREMA)
    c+=pdf.rect(0,PH-46,PW,46,fill=VERDE); c+=pdf.rect(0,PH-46,6,46,fill=DORADO)
    c+=pdf.text(40,PH-30,"Graficos de apoyo: flujo de la demo y hoja de ruta",15,BLANCO,"F2")

    # ---- Flujo de la demo (cadena de etapas) ----
    y=PH-90
    c+=pdf.text(40,y,"Flujo de operacion que recorre la demo (caso: Prediccion de desercion)",12,VERDE,"F2")
    y-=20
    etapas=[("Facultad\npropone",AZUL),("Sede\ncapacidades",TEAL),("Plenaria\nrecibe",TEAL),
            ("4 Nucleos\ntramite",DORADO),("Votacion\nquorum",CONTROL),("CNED\nescala+veto",CNED)]
    n=len(etapas); gap=14; bw=(PW-80-gap*(n-1))/n; bh=52; x=40
    for i,(lab,col) in enumerate(etapas):
        c+=pdf.rect(x,y-bh,bw,bh,fill=col,stroke=VERDEPF,lw=0.8)
        for j,ln in enumerate(lab.split("\n")):
            c+=pdf.text_center(x+bw/2,y-20-j*13,ln,9,BLANCO,"F2" if j==0 else "F1")
        if i<n-1:
            ax=x+bw+gap/2
            c+=pdf.line(x+bw+2,y-bh/2,x+bw+gap-2,y-bh/2,VERDE,1.5)
            c+=pdf.text(ax-3,y-bh/2-10,">",11,VERDE,"F2")
        x+=bw+gap
    c+=pdf.text(40,y-bh-16,"Resultado: riesgo Alto -> el nucleo Juridico (1.3) objeta por datos sensibles -> objecion prevalente -> escala al CNED.",9.5,(0.17,0.22,0.20))

    # ---- Diagrama plenaria (nodo central + 4 nucleos) ----
    y2=y-bh-55
    c+=pdf.text(40,y2,"Operacion colegiada: Plenaria de Delegados y los 4 nucleos",12,VERDE,"F2")
    cx=PW/2; cy=y2-150
    # plenaria
    c+=pdf.rect(cx-85,cy+70,170,34,fill=TEAL,stroke=VERDEPF,lw=1)
    c+=pdf.text_center(cx,cy+90,"PLENARIA DE DELEGADOS",10,BLANCO,"F2")
    c+=pdf.text_center(cx,cy+78,"D1 - D2 - D3 - D4",8,(0.9,0.97,0.95),"F1")
    nucs=[("1.1 Estrategico",AZUL),("1.2 Tecnico/Green",TEAL),("1.3 Control/Jur.",CONTROL),("1.4 Socio-Amb.",SOCIO)]
    nw=120; nx0=(PW-(nw*2+30))/2
    pos=[(nx0,cy),(nx0+nw+30,cy),(nx0,cy-50),(nx0+nw+30,cy-50)]
    for (lab,col),(px,py) in zip(nucs,pos):
        c+=pdf.rect(px,py,nw,36,fill=BLANCO,stroke=col,lw=1.5)
        c+=pdf.rect(px,py,4,36,fill=col)
        c+=pdf.text(px+10,py+14,lab,9.5,col,"F2")
        # flecha plenaria->nucleo
        c+=pdf.line(cx,cy+70,px+nw/2,py+36,(0.6,0.6,0.6),0.8)
    c+=pdf.text_center(cx,cy-70,"La plenaria traslada a tramite, recibe ponencias y decide por quorum (regla: objecion de 1.3/1.4 prevalece).",8.5,(0.3,0.35,0.33),"F1")

    # ---- Hoja de ruta 2025-2027 ----
    y3=cy-100
    c+=pdf.text(40,y3,"Hoja de ruta 2025-2027",12,VERDE,"F2")
    y3-=16; px=40; pw=(PW-80-10*4)/5
    for i,(lab,dt,col) in enumerate(PLAN):
        c+=pdf.rect(px,y3-40,pw,40,fill=col,stroke=VERDEPF,lw=0.8)
        c+=pdf.text_center(px+pw/2,y3-18,lab,8.5,BLANCO,"F2")
        c+=pdf.text_center(px+pw/2,y3-30,dt,8,(0.92,0.92,0.92),"F1")
        if i<4:
            c+=pdf.text(px+pw+1,y3-24,">",11,VERDE,"F2")
        px+=pw+10
    c+=pdf.rect(0,0,PW,22,fill=VERDE)
    c+=pdf.text(40,8,"Fuentes: Propuesta Mesa Tecnica de IA, presentacion Gobernanza IA y comentarios J.M. Sanchez Torres (2026).",8.5,CREMA)
    pdf.add_page(c)

# ---------------- PÁGINA BRIEF DE LA PONENCIA ----------------
def page_brief():
    c=""
    c+=pdf.rect(0,0,PW,PH,fill=CREMA)
    c+=pdf.rect(0,PH-56,PW,56,fill=VERDEPF); c+=pdf.rect(0,PH-56,6,56,fill=DORADO)
    c+=pdf.text(40,PH-28,"Brief de la ponencia",17,BLANCO,"F2")
    c+=pdf.text(40,PH-46,"Para programa academico / envio a organizadores",10,(0.78,0.84,0.80),"F1")
    y=PH-82
    # título de la charla
    tlines=textwrap.wrap("Gobernanza de la Inteligencia Artificial en la UNAL: un modelo multinivel, participativo y multicampus",72)
    for ln in tlines:
        c+=pdf.text(40,y,ln,13,VERDE,"F2"); y-=17
    y-=4
    # ficha
    ficha=[("Entidad","DNED - Division Universidad Laboratorio (DIVULAB)"),
           ("Formato","Ponencia con demostracion interactiva en vivo"),
           ("Duracion","30 min (~26 exposicion + 4 preguntas)"),
           ("Nivel","Introductorio-intermedio (sin requisitos tecnicos)"),
           ("Idioma","Espanol (abstract tambien en ingles)"),
           ("Cupo","Hasta 120 presenciales + transmision hibrida")]
    for k,v in ficha:
        c+=pdf.text(44,y,k+":",9.5,DORADO,"F2")
        c+=pdf.text(120,y,v,9.5,(0.17,0.22,0.20),"F1"); y-=14
    y-=6
    # abstract
    c+=pdf.text(40,y,"Resumen",12,VERDE,"F2"); y-=16
    ab,y=pdf.wrap(44,y,
      "La adopcion de modelos de lenguaje, sistemas predictivos y agentes autonomos ya incide en la docencia, la "
      "investigacion, la extension y la gestion universitaria. La pregunta determinante no es SI usar la IA, sino QUIEN "
      "decide que se puede hacer, con que datos y bajo que limites. Esta ponencia presenta el modelo de gobernanza de la "
      "IA de la UNAL: un esquema multinivel, participativo y multicampus que articula al CNED, una Mesa Tecnica de IA con "
      "cuatro nucleos y a la comunidad de las nueve sedes. A partir de un benchmarking internacional (MIT, UCLA, Stanford) "
      "ampliado con referentes del Sur Global (Tec de Monterrey, USP, UNAM, INTEGRITY-AI), se muestra como el modelo "
      "traduce principios -subsidiariedad, gobierno corporativo de la IA, riesgo graduado por derechos y sostenibilidad- "
      "en un procedimiento operativo: una plenaria de delegados que recibe los casos, los traslada a tramite, delibera "
      "por quorum y emite un dictamen colegiado. Incluye una demostracion en vivo del ciclo de una iniciativa de IA.",
      9.3,(0.17,0.22,0.20),104,12.5)
    y-=8
    # palabras clave
    c+=pdf.text(44,y,"Palabras clave: gobernanza de IA, subsidiariedad, gestion de riesgo, matriz RACI, soberania del",9,(0.40,0.33,0.14),"F1"); y-=12
    c+=pdf.text(44,y,"conocimiento, multicampus, Green AI.",9,(0.40,0.33,0.14),"F1"); y-=18
    # objetivos + aportes en dos columnas
    colY=y
    c+=pdf.text(40,colY,"Objetivos",12,VERDE,"F2")
    obj=["Presentar el modelo y mostrar como opera.",
         "Contrastar modelos internacionales + Sur Global.",
         "Explicar la arquitectura multinivel y 4 nucleos.",
         "Mostrar plenaria, quorum y votacion.",
         "Demostrar en vivo el ciclo de una iniciativa."]
    oy=colY-16
    for o in obj:
        for j,ln in enumerate(textwrap.wrap(o,46)):
            c+=pdf.text(44 if j else 44,oy,(chr(149)+" " if j==0 else "   ")+ln,9,(0.17,0.22,0.20),"F1"); oy-=12
    c+=pdf.text(PW/2+6,colY,"Aportes / mensajes clave",12,VERDE,"F2")
    ap=["Gobernanza = derechos, mision y sostenibilidad.",
        "Subsidiariedad: decidir cerca del riesgo/valor.",
        "Veto graduado por riesgo a derechos.",
        "Plenaria + quorum hace trazable quien decide.",
        "Integra referentes LatAm y revision academica."]
    ay=colY-16
    for a in ap:
        for j,ln in enumerate(textwrap.wrap(a,46)):
            c+=pdf.text(PW/2+10,ay,(chr(149)+" " if j==0 else "   ")+ln,9,(0.17,0.22,0.20),"F1"); ay-=12
    # requerimientos
    ry=min(oy,ay)-8
    c+=pdf.rect(40,ry-40,PW-80,40,fill=(0.96,0.93,0.86),stroke=DORADO,lw=0.7)
    c+=pdf.text(48,ry-14,"Requerimientos:",9.5,(0.40,0.33,0.14),"F2")
    c+=pdf.text(140,ry-14,"proyector + audio; navegador moderno (Three.js/Chart.js via CDN); plan B con capturas;",9,(0.40,0.33,0.14),"F1")
    c+=pdf.text(140,ry-27,"opcional 2a pantalla para el modo presentador con cronometro.",9,(0.40,0.33,0.14),"F1")
    c+=pdf.rect(0,0,PW,22,fill=VERDE)
    c+=pdf.text(40,8,"Brief de la ponencia  -  Mesa Tecnica de IA  -  DNED-DIVULAB",8.5,CREMA)
    pdf.add_page(c)

# ---- construir ----
page_cover()
page_brief()
# dividir bloques en páginas (3 por página para que quepan)
groups=[BLOCKS[0:3],BLOCKS[3:5],BLOCKS[5:7],BLOCKS[7:9]]
labels=["Bloques 0-2","Bloques 3-4","Bloques 5-6","Bloques 7-8"]
for g,l in zip(groups,labels):
    page_blocks(g,l)
page_demo_and_plan()

data=pdf.build()
with open("roadmap-exposicion-gobernanza-ia-unal.pdf","wb") as f:
    f.write(data)
print("PDF generado:", len(data), "bytes,", len(pdf.pages), "paginas")
