"""Generación del Word (.docx) editable, con el mismo formato que el PDF."""
import io

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .pdf import Perfil
from .texto import es_pie, limpiar, sin_marcas, tokens

FUENTE = "Times New Roman"


def _fuente(estilo, tam, negrita=None, cursiva=None):
    f = estilo.font
    f.name = FUENTE
    f.size = Pt(tam)
    f.color.rgb = RGBColor(0, 0, 0)
    if negrita is not None:
        f.bold = negrita
    if cursiva is not None:
        f.italic = cursiva
    rpr = estilo.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FUENTE)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]


def _campo(par, instruccion, texto_previo=""):
    """Inserta un campo de Word (PAGE, TOC…) con un resultado provisional."""
    def run_con(el):
        r = par.add_run()
        r._r.append(el)
        return r

    ini = OxmlElement("w:fldChar"); ini.set(qn("w:fldCharType"), "begin")
    ins = OxmlElement("w:instrText"); ins.set(qn("xml:space"), "preserve"); ins.text = f" {instruccion} "
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    fin = OxmlElement("w:fldChar"); fin.set(qn("w:fldCharType"), "end")
    run_con(ini); run_con(ins); run_con(sep)
    if texto_previo:
        par.add_run(texto_previo)
    run_con(fin)


def _run_fldchar(tipo):
    r = OxmlElement("w:r")
    f = OxmlElement("w:fldChar")
    f.set(qn("w:fldCharType"), tipo)
    r.append(f)
    return r


def _elementos_campo_inicio(instruccion):
    ini = _run_fldchar("begin")
    r = OxmlElement("w:r")
    ins = OxmlElement("w:instrText")
    ins.set(qn("xml:space"), "preserve")
    ins.text = f" {instruccion} "
    r.append(ins)
    sep = _run_fldchar("separate")
    return [ini, r, sep]


def _inline(par, texto, tam=None):
    for t, negrita, cursiva, codigo in tokens(texto):
        r = par.add_run(t)
        r.bold = True if negrita else None
        r.italic = True if cursiva else None
        if codigo:
            r.font.name = "Courier New"
        if tam:
            r.font.size = Pt(tam)


def _sombrear(celda, color="EEEEEE"):
    tcpr = celda._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), color)
    tcpr.append(sh)


def _par_vacio(doc, antes=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(antes)
    p.paragraph_format.first_line_indent = Cm(0)
    return p


def _centrado(doc, texto, tam, negrita=False, cursiva=False, antes=0, despues=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.15
    pf.space_before, pf.space_after = Pt(antes), Pt(despues)
    r = p.add_run(texto)
    r.font.size = Pt(tam)
    r.bold, r.italic = negrita, cursiva
    return p


def construir_docx(datos: dict, prep: list, perfil: Perfil, opciones: dict) -> bytes:
    doc = Document()
    d = {k: limpiar(str(v or "")) for k, v in datos.items()}
    romper_cap = bool(opciones.get("capitulo_nueva_pagina", True))
    toc_sub = bool(opciones.get("toc_subtitulos", True))

    # --- Metadatos ---
    doc.core_properties.title = d.get("titulo", "")
    doc.core_properties.author = d.get("autor", "")
    doc.core_properties.language = "es-PE"

    # --- Página A4 y márgenes ---
    sec0 = doc.sections[0]
    sec0.page_width, sec0.page_height = Cm(21.0), Cm(29.7)
    sec0.left_margin, sec0.right_margin = Cm(3.0), Cm(2.5)
    sec0.top_margin = sec0.bottom_margin = Cm(2.5)

    # --- Estilos ---
    est = doc.styles
    normal = est["Normal"]
    _fuente(normal, perfil.tam)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(1.25)
    pf.line_spacing = perfil.esp
    pf.space_before, pf.space_after = Pt(0), Pt(perfil.despues)

    for nivel, tam, neg, cur, antes, despues in [
        (1, perfil.tam + 6, True, False, 0, 18),
        (2, perfil.tam + 2, True, False, 12, 6),
        (3, perfil.tam + 1, True, True, 8, 4),
    ]:
        h = est[f"Heading {nivel}"]
        _fuente(h, tam, neg, cur)
        hf = h.paragraph_format
        hf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        hf.first_line_indent = Cm(0)
        hf.line_spacing = 1.15
        hf.space_before, hf.space_after = Pt(antes), Pt(despues)
        hf.keep_with_next = True

    for nombre in ("List Bullet",):
        e = est[nombre]
        _fuente(e, perfil.tam)
        e.paragraph_format.first_line_indent = None
        e.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        e.paragraph_format.line_spacing = perfil.esp
        e.paragraph_format.space_after = Pt(max(2.0, perfil.despues / 2))

    # Estilos de las líneas del índice
    for nivel, tam, neg, sangria in [(1, 12, True, 0.0), (2, 11, False, 0.8), (3, 10.5, False, 1.6)]:
        try:
            e = est[f"TOC {nivel}"]
        except KeyError:
            e = est.add_style(f"TOC {nivel}", 1)
            e.base_style = normal
        _fuente(e, tam, neg)
        e.paragraph_format.first_line_indent = Cm(0)
        e.paragraph_format.left_indent = Cm(sangria)
        e.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        e.paragraph_format.line_spacing = 1.15
        e.paragraph_format.space_after = Pt(4)

    # Estilo del título de portada (permite reconocerlo al volver a subir el archivo)
    et = est["Title"]
    _fuente(et, 22, True, False)
    et.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    et.paragraph_format.first_line_indent = Cm(0)
    et.paragraph_format.line_spacing = 1.15
    et.paragraph_format.space_before, et.paragraph_format.space_after = Pt(90), Pt(40)
    ppr = et.element.get_or_add_pPr()
    for borde in ppr.findall(qn("w:pBdr")):
        ppr.remove(borde)

    eh = est["TOC Heading"]
    _fuente(eh, 18, True, False)
    eh.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    eh.paragraph_format.first_line_indent = Cm(0)
    eh.paragraph_format.space_before, eh.paragraph_format.space_after = Pt(0), Pt(18)
    eh.paragraph_format.keep_with_next = True

    # --- Portada ---
    _centrado(doc, d.get("universidad", "").upper(), 15, negrita=True, despues=6)
    _centrado(doc, d.get("facultad", ""), 12)
    _centrado(doc, d.get("escuela", ""), 12)
    doc.add_paragraph(d.get("titulo", ""), style="Title")
    _centrado(doc, "Monografía presentada para el curso de", 12, antes=20)
    _centrado(doc, d.get("curso", ""), 12, cursiva=True, despues=60)
    for etiqueta, clave in (("Autor", "autor"), ("", "codigo"), ("Docente", "docente"), ("Ciclo", "ciclo")):
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_after = Pt(6)
        if etiqueta:
            r = p.add_run(f"{etiqueta}: ")
            r.bold = True
        p.add_run(d.get(clave, ""))
    _centrado(doc, d.get("ciudad", ""), 12, antes=70)
    _centrado(doc, d.get("anio", ""), 12)

    # --- Nueva sección (sin número en la portada, con número después) ---
    sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec1.footer.is_linked_to_previous = False
    fp = sec1.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.first_line_indent = Cm(0)
    _campo(fp, "PAGE", "1")
    for r in fp.runs:
        r.font.size = Pt(10)

    # --- Índice general (campo TOC real; Word lo actualiza al abrir) ---
    doc.add_paragraph("Índice general", style="TOC Heading")

    titulos_indice = []
    for sec in prep:
        for b in sec["bloques"]:
            if b[0] == "h" and (b[1] == 1 or toc_sub):
                titulos_indice.append((b[1], sin_marcas(b[2])))
    nivel_toc = "1-3" if toc_sub else "1-1"
    instr = f'TOC \\o "{nivel_toc}" \\h \\z \\u'
    entradas = []
    for nivel, texto in titulos_indice:
        q = doc.add_paragraph(style=f"TOC {min(nivel, 3)}")
        q.add_run(texto)
        entradas.append(q)
    # El campo envuelve las líneas provisionales: al actualizarse, Word las reemplaza (sin duplicar).
    if entradas:
        primero_p, ultimo_p = entradas[0], entradas[-1]
        r0 = primero_p.runs[0]._r
        for el in _elementos_campo_inicio(instr):
            r0.addprevious(el)
        ultimo_p._p.append(_run_fldchar("end"))

    # Pedir a Word que actualice campos al abrir
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    settings.append(upd)

    # --- Secciones ---
    for sec in prep:
        primero = True
        es_biblio = sec["clave"] == "bibliografia"
        bl = sec["bloques"]
        for idx, b in enumerate(bl):
            tipo = b[0]
            if tipo == "h":
                nivel, texto = b[1], sin_marcas(b[2])
                p = doc.add_heading(texto, level=nivel)
                if nivel == 1:
                    if primero or romper_cap:
                        p.paragraph_format.page_break_before = True
                    primero = False
            elif tipo == "p":
                p = doc.add_paragraph()
                if es_pie(b[1]):
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.first_line_indent = Cm(0)
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(perfil.despues + 4)
                    _inline(p, b[1], tam=max(8.0, perfil.tam - 1.5))
                    continue
                if es_biblio:
                    p.paragraph_format.left_indent = Cm(1.25)
                    p.paragraph_format.first_line_indent = Cm(-1.25)
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                _inline(p, b[1])
            elif tipo == "ul":
                for it in b[1]:
                    p = doc.add_paragraph(style="List Bullet")
                    _inline(p, it)
            elif tipo == "ol":
                for k, it in enumerate(b[1], 1):
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Cm(1.25)
                    p.paragraph_format.first_line_indent = Cm(-0.75)
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.space_after = Pt(max(2.0, perfil.despues / 2))
                    p.add_run(f"{k}.\t")
                    _inline(p, it)
            elif tipo == "tabla":
                filas = b[1]
                ncol = max(len(f) for f in filas)
                t = doc.add_table(rows=len(filas), cols=ncol)
                t.style = "Table Grid"
                t.alignment = WD_TABLE_ALIGNMENT.CENTER
                for i, fila in enumerate(filas):
                    fila = fila + [""] * (ncol - len(fila))
                    for j, c in enumerate(fila):
                        celda = t.cell(i, j)
                        par = celda.paragraphs[0]
                        par.paragraph_format.first_line_indent = Cm(0)
                        par.paragraph_format.line_spacing = 1.0
                        par.paragraph_format.space_after = Pt(2)
                        par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        par.paragraph_format.keep_with_next = True   # la tabla no se parte entre páginas
                        _inline(par, f"**{c}**" if (i == 0 and c) else c, tam=max(8.0, perfil.tam - 2))
                        if i == 0:
                            _sombrear(celda)
                for fila_tabla in t.rows:
                    trpr = fila_tabla._tr.get_or_add_trPr()
                    trpr.append(OxmlElement("w:cantSplit"))
                if not (idx + 1 < len(bl) and bl[idx + 1][0] == "p" and es_pie(bl[idx + 1][1])):
                    _par_vacio(doc)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
