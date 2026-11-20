"""Importar un archivo (.docx, .txt, .md) y repartirlo en las secciones de la monografía."""
import io
import re

from .texto import SECCIONES, TITULOS_DEFECTO, normalizar

MARCA_TITULO = "%%TITULO%%"

# clave → prefijos (ya normalizados) que identifican el encabezado de esa sección
_CLAVES = [
    ("prologo", ("prologo",)),
    ("introduccion", ("introduccion",)),
    ("cuerpo", ("cuerpo del trabajo",)),
    ("conclusiones", ("conclusion",)),
    ("anexos", ("anexo",)),
    ("bibliografia", ("bibliografia", "referencias", "fuentes consultadas")),
    ("ignorar", ("indice", "tabla de contenido", "contenido", "portada", "caratula")),
]
_RE_H = re.compile(r"^(#{1,6})\s+(.*\S)\s*$")


# ---------------------------------------------------------------------------
#  Lectura de archivos
# ---------------------------------------------------------------------------
def _fmt_runs(p) -> str:
    """Texto de un párrafo de Word con **negrita** y *cursiva* en Markdown."""
    grupos = []  # [texto, negrita, cursiva]
    for r in p.runs:
        t = r.text
        if not t:
            continue
        b, i = bool(r.bold), bool(r.italic)
        if grupos and grupos[-1][1] == b and grupos[-1][2] == i:
            grupos[-1][0] += t
        else:
            grupos.append([t, b, i])
    partes = []
    for t, b, i in grupos:
        nucleo = t.strip()
        if not nucleo:
            partes.append(t)
            continue
        ini = t[: len(t) - len(t.lstrip())]
        fin = t[len(t.rstrip()):]
        if b:
            nucleo = f"**{nucleo}**"
        elif i:
            nucleo = f"*{nucleo}*"
        partes.append(ini + nucleo + fin)
    texto = "".join(partes).strip()
    plano = re.sub(r"[*]", "", texto)
    # Si hay hipervínculos u otros elementos que .runs no incluye, usar el texto plano.
    if len(plano) < len(p.text.strip()):
        return p.text.strip()
    return texto


def docx_a_markdown(archivo) -> tuple[str, list[str]]:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    avisos = []
    doc = Document(archivo)
    if len(doc.inline_shapes) > 0:
        avisos.append(
            f"El archivo contiene {len(doc.inline_shapes)} imagen(es); "
            "no se importan (solo texto y tablas)."
        )
    salida = []
    for hijo in doc.element.body.iterchildren():
        etiqueta = hijo.tag.rsplit("}", 1)[-1]
        if etiqueta == "p":
            p = Paragraph(hijo, doc)
            texto = _fmt_runs(p)
            if not texto:
                continue
            estilo = (p.style.name if p.style is not None else "") or ""
            est = estilo.lower()
            if est.startswith(("toc", "tdc", "índice", "indice", "table of contents")):
                continue  # líneas del índice generado por Word
            m = re.match(r"^(heading|t[ií]tulo|encabezado)\s*(\d)", est)
            if m:
                nivel = max(1, min(int(m.group(2)), 6))
                salida.append("#" * nivel + " " + re.sub(r"[*]", "", texto))
            elif est in ("title", "título", "titulo"):
                salida.append(f"{MARCA_TITULO} {re.sub(r'[*]', '', texto)}")
            elif any(k in est for k in ("list bullet", "viñeta", "vineta")):
                salida.append(f"- {texto}")
            elif any(k in est for k in ("list number", "numerad")):
                salida.append(f"1. {texto}")
            elif hijo.pPr is not None and hijo.pPr.numPr is not None:
                salida.append(f"- {texto}")
            else:
                salida.append(texto)
            salida.append("")
        elif etiqueta == "tbl":
            t = Table(hijo, doc)
            filas = []
            for fila in t.rows:
                celdas = []
                vistos = set()
                for c in fila.cells:
                    if id(c._tc) in vistos:
                        continue
                    vistos.add(id(c._tc))
                    celdas.append(c.text.replace("\n", " ").replace("|", "/").strip())
                filas.append(celdas)
            if filas:
                ancho = max(len(f) for f in filas)
                for k, f in enumerate(filas):
                    f = f + [""] * (ancho - len(f))
                    salida.append("| " + " | ".join(f) + " |")
                    if k == 0:
                        salida.append("|" + "---|" * ancho)
                salida.append("")
    texto = "\n".join(salida)
    # Los elementos consecutivos de una lista forman una sola lista (sin líneas en blanco entre ellos)
    texto = re.sub(r"(?m)^(- .*)\n\n(?=- )", r"\1\n", texto)
    texto = re.sub(r"(?m)^(1\. .*)\n\n(?=1\. )", r"\1\n", texto)
    return texto, avisos


def _decodificar(datos: bytes) -> str:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return datos.decode(enc)
        except UnicodeDecodeError:
            continue
    return datos.decode("utf-8", errors="replace")


def _encabezados_en_texto_plano(texto: str) -> str:
    """Para .txt sin '#': detecta títulos evidentes (INTRODUCCIÓN, Capítulo 1, 1.1 ...)."""
    if re.search(r"^#{1,6}\s", texto, flags=re.M):
        return texto
    lineas = texto.replace("\r", "").split("\n")
    res = []
    for k, ln in enumerate(lineas):
        s = ln.strip()
        antes_vacio = k == 0 or not lineas[k - 1].strip()
        if s and antes_vacio and len(s) <= 100 and not s.endswith((".", ",", ";")):
            n = normalizar(s)
            if any(n.startswith(p) for _, ps in _CLAVES for p in ps if _ != "ignorar") and len(s) <= 60:
                res.append("# " + s)
                continue
            if re.match(r"^cap[ií]tulo\s+[\divxlc]+", s, flags=re.I):
                res.append("# " + s)
                continue
            m = re.match(r"^(\d+)\.(\d+)(\.\d+)?\s+\S", s)
            if m:
                res.append(("### " if m.group(3) else "## ") + s)
                continue
        res.append(ln)
    return "\n".join(res)


def leer_archivo(nombre: str, datos: bytes) -> tuple[str, list[str]]:
    """Devuelve (markdown, avisos) a partir de un archivo subido."""
    ext = nombre.lower().rsplit(".", 1)[-1] if "." in nombre else ""
    if ext == "docx":
        return docx_a_markdown(io.BytesIO(datos))
    if ext in ("txt", "md", "markdown"):
        texto = _decodificar(datos)
        if ext == "txt":
            texto = _encabezados_en_texto_plano(texto)
        return texto, []
    raise ValueError("Formato no soportado. Suba un archivo .docx, .txt o .md.")


# ---------------------------------------------------------------------------
#  Reparto en secciones
# ---------------------------------------------------------------------------
def _clave_de(titulo: str):
    n = normalizar(titulo)
    n = re.sub(r"^([\divxlc]+[.)])\s*", "", n)  # "1. Introducción", "II) Conclusiones"
    for clave, prefijos in _CLAVES:
        if any(n.startswith(p) for p in prefijos):
            return clave
    return None


def separar_secciones(markdown: str):
    """Reparte el texto en las secciones.

    Devuelve (secciones: dict clave→texto, titulos: dict clave→título,
    titulo_sugerido: str|None, avisos: list[str]).
    """
    lineas = markdown.replace("\r", "").split("\n")
    titulo_marca = None
    cuerpo_lineas = []
    for ln in lineas:
        if ln.startswith(MARCA_TITULO):
            titulo_marca = titulo_marca or ln[len(MARCA_TITULO):].strip()
        else:
            cuerpo_lineas.append(ln)
    lineas = cuerpo_lineas

    # ¿Hay algún encabezado que identifique una sección?
    hay_claves = False
    en_codigo = False
    for ln in lineas:
        if ln.strip().startswith("```"):
            en_codigo = not en_codigo
        m = _RE_H.match(ln.strip())
        if m and not en_codigo and len(m.group(1)) <= 2:
            k = _clave_de(m.group(2))
            if k and k != "ignorar":
                hay_claves = True
                break

    secs = {k: [] for k, _ in SECCIONES}
    titulos = {}
    avisos = []

    if not hay_claves:
        secs["cuerpo"] = lineas[:]
        if markdown.strip():
            avisos.append(
                "No se detectaron secciones (Prólogo, Introducción, Conclusiones…). "
                "Todo el texto se colocó en el cuerpo del trabajo; puede repartirlo en la pestaña «Editar»."
            )
    else:
        previo = []
        actual = None          # None = antes de la primera sección
        nivel_actual = 0
        en_codigo = False
        for ln in lineas:
            if ln.strip().startswith("```"):
                en_codigo = not en_codigo
            m = None if en_codigo else _RE_H.match(ln.strip())
            if m:
                nivel, titulo = len(m.group(1)), m.group(2).strip()
                k = _clave_de(titulo) if nivel <= 2 else None
                if k == "ignorar":
                    actual, nivel_actual = "ignorar", nivel
                    continue
                if k:
                    if k == actual:
                        secs[k].append(ln)           # p. ej. «Anexo B» dentro de anexos
                    else:
                        actual, nivel_actual = k, nivel
                        generico = normalizar(titulo) in (normalizar(TITULOS_DEFECTO[k]), "anexo")
                        if k == "anexos" and not generico:
                            secs[k].append("## " + titulo)
                        elif k not in titulos:
                            titulos[k] = titulo
                    continue
                # Encabezado que no es de sección
                if actual in (None, "ignorar"):
                    actual, nivel_actual = ("cuerpo" if actual == "ignorar" else None), nivel
                    if actual is None:
                        previo.append(ln)
                        continue
                elif actual != "cuerpo" and nivel <= nivel_actual:
                    actual, nivel_actual = "cuerpo", nivel
                if actual is None:
                    previo.append(ln)
                else:
                    secs[actual].append(ln)
                continue
            if actual is None:
                previo.append(ln)
            elif actual == "ignorar":
                continue
            else:
                secs[actual].append(ln)

        # Texto anterior a la primera sección: la portada son líneas cortas (se descarta);
        # si hay párrafos largos de prosa, se conservan en el cuerpo del trabajo.
        parrafos_previo = [x.strip() for x in "\n".join(previo).split("\n\n") if x.strip()]
        if any(len(x) >= 250 for x in parrafos_previo):
            secs["cuerpo"] = previo + [""] + secs["cuerpo"]
            avisos.append("El texto anterior a la primera sección se agregó al cuerpo del trabajo.")

    secciones = {k: "\n".join(v).strip() for k, v in secs.items()}
    vacias = [TITULOS_DEFECTO[k] for k, _ in SECCIONES if not secciones[k]]
    if vacias and hay_claves:
        avisos.append("Secciones sin contenido en el archivo: " + ", ".join(vacias) + ".")
    return secciones, titulos, titulo_marca, avisos


def exportar_markdown(titulo: str, secciones: dict, titulos: dict) -> str:
    """Markdown que puede volver a subirse a la aplicación."""
    def bajar_a_nivel_2(texto: str) -> str:
        """Los títulos internos de una sección deben ser nivel 2 o más (el nivel 1 es de la sección)."""
        niveles, en_codigo = [], False
        for ln in texto.split("\n"):
            if ln.strip().startswith("```"):
                en_codigo = not en_codigo
            m = None if en_codigo else _RE_H.match(ln.strip())
            if m:
                niveles.append(len(m.group(1)))
        if not niveles or min(niveles) >= 2:
            return texto
        d, res, en_codigo = 2 - min(niveles), [], False
        for ln in texto.split("\n"):
            if ln.strip().startswith("```"):
                en_codigo = not en_codigo
            m = None if en_codigo else _RE_H.match(ln.strip())
            res.append("#" * min(6, len(m.group(1)) + d) + " " + m.group(2) if m else ln)
        return "\n".join(res)

    partes = []
    if titulo:
        partes.append(f"{MARCA_TITULO} {titulo}\n")
    for clave, defecto in SECCIONES:
        t = (titulos.get(clave) or defecto).strip()
        texto = (secciones.get(clave) or "").strip()
        if clave != "cuerpo":
            texto = bajar_a_nivel_2(texto)
        partes.append(f"# {t}\n\n{texto}\n")
    return "\n".join(partes).strip() + "\n"
