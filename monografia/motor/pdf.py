"""Generación del PDF (reportlab) con portada, índice real, secciones y numeración."""
import io
from dataclasses import dataclass
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate, Frame, KeepInFrame, KeepTogether, NextPageTemplate, PageBreak,
    PageTemplate, Paragraph, Spacer, Table, TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents

from .texto import es_pie, limpiar, para_pdf, sin_marcas, tokens

MARGEN_IZQ, MARGEN_DER, MARGEN_SUP, MARGEN_INF = 3 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm
FACTOR_WORD = 1.15  # una "línea sencilla" de Word equivale a ~1,15 × el tamaño de letra


@dataclass(frozen=True)
class Perfil:
    tam: float = 12.0      # tamaño de letra (pt)
    esp: float = 1.5       # interlineado (múltiplo, como en Word)
    despues: float = 6.0   # espacio después de cada párrafo (pt)

    @property
    def costo(self) -> float:
        """Aproximación de cuánto espacio ocupa el texto (para ordenar perfiles)."""
        return self.tam * self.tam * self.esp + 0.5 * self.despues


PERFIL_ESTANDAR = Perfil(12.0, 1.5, 6.0)


def inline(texto: str) -> str:
    partes = []
    for t, negrita, cursiva, codigo in tokens(texto):
        t = escape(t)
        if codigo:
            t = f'<font face="Courier">{t}</font>'
        if negrita:
            t = f"<b>{t}</b>"
        if cursiva:
            t = f"<i>{t}</i>"
        partes.append(t)
    return "".join(partes)


def _estilos(p: Perfil) -> dict:
    lead = p.tam * FACTOR_WORD * p.esp
    st = {}
    st["Cuerpo"] = ParagraphStyle(
        "Cuerpo", fontName="Times-Roman", fontSize=p.tam, leading=lead,
        alignment=TA_JUSTIFY, firstLineIndent=1.25 * cm, spaceAfter=p.despues,
    )
    st["Lista"] = ParagraphStyle(
        "Lista", parent=st["Cuerpo"], firstLineIndent=0, leftIndent=1.25 * cm,
        bulletIndent=0.5 * cm, spaceAfter=max(2.0, p.despues / 2),
    )
    st["Biblio"] = ParagraphStyle(
        "Biblio", parent=st["Cuerpo"], firstLineIndent=-1.25 * cm, leftIndent=1.25 * cm,
        alignment=TA_LEFT,
    )
    st["H1"] = ParagraphStyle(
        "H1", fontName="Times-Bold", fontSize=p.tam + 6, leading=(p.tam + 6) * 1.2,
        spaceAfter=18, alignment=TA_LEFT, keepWithNext=1,
    )
    st["H2"] = ParagraphStyle(
        "H2", fontName="Times-Bold", fontSize=p.tam + 2, leading=(p.tam + 2) * 1.2,
        spaceBefore=12, spaceAfter=6, alignment=TA_LEFT, keepWithNext=1,
    )
    st["H3"] = ParagraphStyle(
        "H3", fontName="Times-BoldItalic", fontSize=p.tam + 1, leading=(p.tam + 1) * 1.2,
        spaceBefore=8, spaceAfter=4, alignment=TA_LEFT, keepWithNext=1,
    )
    st["Pie"] = ParagraphStyle(
        "Pie", parent=st["Cuerpo"], fontSize=max(8.0, p.tam - 1.5), leading=max(8.0, p.tam - 1.5) * 1.25,
        firstLineIndent=0, alignment=TA_CENTER, spaceBefore=0, spaceAfter=p.despues + 4,
    )
    tc = max(8.0, p.tam - 2)
    st["Celda"] = ParagraphStyle("Celda", fontName="Times-Roman", fontSize=tc, leading=tc * 1.2)
    # Portada e índice: tamaños fijos
    st["PUni"] = ParagraphStyle("PUni", fontName="Times-Bold", fontSize=15, leading=19, alignment=TA_CENTER)
    st["PTxt"] = ParagraphStyle("PTxt", fontName="Times-Roman", fontSize=12, leading=16, alignment=TA_CENTER)
    st["PTit"] = ParagraphStyle("PTit", fontName="Times-Bold", fontSize=22, leading=28, alignment=TA_CENTER)
    st["PAut"] = ParagraphStyle("PAut", fontName="Times-Roman", fontSize=12, leading=17, alignment=TA_CENTER)
    st["TOCT"] = ParagraphStyle(
        "TOCT", fontName="Times-Bold", fontSize=18, leading=22, spaceAfter=18, alignment=TA_LEFT
    )
    return st


def _toc() -> TableOfContents:
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle("T0", fontName="Times-Bold", fontSize=12, leading=18, leftIndent=0, firstLineIndent=0),
        ParagraphStyle("T1", fontName="Times-Roman", fontSize=11, leading=15, leftIndent=0.8 * cm, firstLineIndent=0),
        ParagraphStyle("T2", fontName="Times-Roman", fontSize=10.5, leading=14, leftIndent=1.6 * cm, firstLineIndent=0),
    ]
    return toc


def _portada(datos: dict, st: dict, ancho: float, alto: float):
    g = lambda k: escape(str(datos.get(k, "") or ""))
    f = [
        Paragraph(g("universidad").upper(), st["PUni"]),
        Spacer(1, 6),
        Paragraph(g("facultad"), st["PTxt"]),
        Paragraph(g("escuela"), st["PTxt"]),
        Spacer(1, 2.6 * cm),
        Paragraph(g("titulo"), st["PTit"]),
        Spacer(1, 1.6 * cm),
        Paragraph("Monografía presentada para el curso de", st["PTxt"]),
        Paragraph(f"<i>{g('curso')}</i>", st["PTxt"]),
        Spacer(1, 2.6 * cm),
        Paragraph(f"<b>Autor:</b> {g('autor')}", st["PAut"]),
        Paragraph(g("codigo"), st["PAut"]),
        Spacer(1, 8),
        Paragraph(f"<b>Docente:</b> {g('docente')}", st["PAut"]),
        Spacer(1, 8),
        Paragraph(f"<b>Ciclo:</b> {g('ciclo')}", st["PAut"]),
    ]
    return KeepInFrame(ancho, alto, f, mode="shrink")


def _tabla(filas, st, ancho):
    ncol = max(len(f) for f in filas)
    datos = []
    for k, fila in enumerate(filas):
        fila = fila + [""] * (ncol - len(fila))
        datos.append([
            Paragraph(f"<b>{inline(c)}</b>" if k == 0 else inline(c), st["Celda"]) for c in fila
        ])
    t = Table(datos, colWidths=[ancho / ncol] * ncol, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


class _Plantilla(BaseDocTemplate):
    def __init__(self, *a, toc_sub=True, **k):
        super().__init__(*a, **k)
        self.toc_sub = toc_sub

    def afterFlowable(self, fl):
        if isinstance(fl, Paragraph):
            nombre = fl.style.name
            if nombre in ("H1", "H2", "H3"):
                nivel = int(nombre[1]) - 1
                if nivel == 0 or self.toc_sub:
                    self.notify("TOCEntry", (nivel, fl.getPlainText(), self.page))


def _sanear(prep: list) -> list:
    res = []
    for sec in prep:
        nb = []
        for b in sec["bloques"]:
            if b[0] == "h":
                nb.append(("h", b[1], para_pdf(b[2])))
            elif b[0] == "p":
                nb.append(("p", para_pdf(b[1])))
            elif b[0] in ("ul", "ol"):
                nb.append((b[0], [para_pdf(x) for x in b[1]]))
            else:
                nb.append(("tabla", [[para_pdf(c) for c in f] for f in b[1]]))
        res.append({**sec, "bloques": nb, "titulo": para_pdf(sec["titulo"])})
    return res


def construir_pdf(datos: dict, prep: list, perfil: Perfil, opciones: dict):
    """Devuelve (bytes_pdf, numero_de_paginas)."""
    prep = _sanear(prep)
    datos = {k: para_pdf(str(v or "")) for k, v in datos.items()}
    W, H = A4
    ancho = W - MARGEN_IZQ - MARGEN_DER
    alto = H - MARGEN_SUP - MARGEN_INF
    st = _estilos(perfil)
    romper_cap = bool(opciones.get("capitulo_nueva_pagina", True))

    def marco(id_):
        return Frame(MARGEN_IZQ, MARGEN_INF, ancho, alto, id=id_,
                     leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

    def en_portada(canvas, doc):
        canvas.saveState()
        canvas.setFont("Times-Roman", 12)
        canvas.drawCentredString(W / 2, MARGEN_INF + 1.5 * cm, limpiar(str(datos.get("ciudad", ""))))
        canvas.drawCentredString(W / 2, MARGEN_INF + 0.6 * cm, limpiar(str(datos.get("anio", ""))))
        canvas.restoreState()

    def en_cuerpo(canvas, doc):
        canvas.saveState()
        canvas.setFont("Times-Roman", 10)
        canvas.drawCentredString(W / 2, 1.3 * cm, str(doc.page))
        canvas.restoreState()

    buf = io.BytesIO()
    doc = _Plantilla(
        buf, pagesize=A4, title=limpiar(str(datos.get("titulo", ""))),
        author=limpiar(str(datos.get("autor", ""))), toc_sub=bool(opciones.get("toc_subtitulos", True)),
        leftMargin=MARGEN_IZQ, rightMargin=MARGEN_DER, topMargin=MARGEN_SUP, bottomMargin=MARGEN_INF,
    )
    doc.addPageTemplates([
        PageTemplate(id="portada", frames=[marco("p")], onPage=en_portada),
        PageTemplate(id="cuerpo", frames=[marco("c")], onPage=en_cuerpo),
    ])

    s = [_portada({k: limpiar(str(v)) for k, v in datos.items()}, st, ancho, alto - 3 * cm),
         NextPageTemplate("cuerpo"), PageBreak(),
         Paragraph("Índice general", st["TOCT"]), _toc()]

    for sec in prep:
        primero = True
        es_biblio = sec["clave"] == "bibliografia"
        bl = sec["bloques"]
        saltar = set()
        for idx, b in enumerate(bl):
            if idx in saltar:
                continue
            tipo = b[0]
            if tipo == "h":
                nivel, texto = b[1], escape(sin_marcas(b[2]))
                if nivel == 1 and (primero or romper_cap):
                    s.append(PageBreak())
                s.append(Paragraph(texto, st[f"H{nivel}"]))
                if nivel == 1:
                    primero = False
            elif tipo == "p":
                if es_pie(b[1]):
                    s.append(Paragraph(inline(b[1]), st["Pie"]))
                else:
                    s.append(Paragraph(inline(b[1]), st["Biblio" if es_biblio else "Cuerpo"]))
            elif tipo == "ul":
                for it in b[1]:
                    s.append(Paragraph(inline(it), st["Lista"], bulletText="•"))
                s.append(Spacer(1, perfil.despues / 2))
            elif tipo == "ol":
                for k, it in enumerate(b[1], 1):
                    s.append(Paragraph(inline(it), st["Lista"], bulletText=f"{k}."))
                s.append(Spacer(1, perfil.despues / 2))
            elif tipo == "tabla":
                grupo = [_tabla(b[1], st, ancho), Spacer(1, 4)]
                if idx + 1 < len(bl) and bl[idx + 1][0] == "p" and es_pie(bl[idx + 1][1]):
                    grupo.append(Paragraph(inline(bl[idx + 1][1]), st["Pie"]))
                    saltar.add(idx + 1)
                else:
                    grupo.append(Spacer(1, perfil.despues + 6))
                s.append(KeepTogether(grupo))

    doc.multiBuild(s)
    return buf.getvalue(), doc.page
