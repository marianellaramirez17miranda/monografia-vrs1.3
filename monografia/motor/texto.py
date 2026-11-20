"""Utilidades de texto: Markdown simplificado → bloques, y preparación del documento."""
import re
import unicodedata

# Secciones del trabajo, en el orden en que aparecen (clave, título por defecto)
SECCIONES = [
    ("prologo", "Prólogo"),
    ("introduccion", "Introducción"),
    ("cuerpo", "Cuerpo del trabajo"),
    ("conclusiones", "Conclusiones"),
    ("anexos", "Anexos"),
    ("bibliografia", "Bibliografía"),
]
TITULOS_DEFECTO = dict(SECCIONES)

PENDIENTE = "*(Sección pendiente de redactar.)*"

_SUSTITUCIONES = {
    "→": "->", "←": "<-", "≤": "<=", "≥": ">=", "≈": "~",
    "−": "-", "‑": "-", "​": "", "­": "", "✓": "", "✔": "",
    "•": "•", "﻿": "",
}


def limpiar(texto: str) -> str:
    """Reemplaza caracteres que las fuentes estándar del PDF no pueden dibujar."""
    for a, b in _SUSTITUCIONES.items():
        texto = texto.replace(a, b)
    return texto


def _no_cp1252(ch: str) -> bool:
    if ch.isspace():
        return False
    try:
        ch.encode("cp1252")
        return False
    except UnicodeEncodeError:
        return True


def no_soportados(*textos) -> list:
    """Caracteres que el PDF no puede dibujar (emoji, alfabeto griego, etc.)."""
    vistos = []
    for t in textos:
        for ch in limpiar(str(t or "")):
            if _no_cp1252(ch) and ch not in vistos:
                vistos.append(ch)
    return vistos


def para_pdf(texto: str) -> str:
    """Sustituye por «?» lo que las fuentes estándar del PDF no pueden dibujar."""
    return "".join("?" if _no_cp1252(ch) else ch for ch in limpiar(texto))


def normalizar(texto: str) -> str:
    """Minúsculas y sin acentos, para comparar títulos."""
    t = unicodedata.normalize("NFD", texto.lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t).strip(" .:-\t")


def palabras(texto: str) -> int:
    return len(re.findall(r"\w+", texto))


# ---------------------------------------------------------------------------
#  Inline: **negrita**, *cursiva*, `código`
# ---------------------------------------------------------------------------
_TOK = re.compile(r"(\*\*.+?\*\*|`[^`\n]+`|\*[^*\n]+?\*)")


def tokens(texto: str):
    """Devuelve (texto, negrita, cursiva, codigo) para cada fragmento."""
    for parte in _TOK.split(texto):
        if not parte:
            continue
        if parte.startswith("**") and parte.endswith("**") and len(parte) > 4:
            yield parte[2:-2], True, False, False
        elif parte.startswith("`") and parte.endswith("`") and len(parte) > 2:
            yield parte[1:-1], False, False, True
        elif parte.startswith("*") and parte.endswith("*") and len(parte) > 2:
            yield parte[1:-1], False, True, False
        else:
            yield parte, False, False, False


_RE_PIE = re.compile(r"^\*(Tabla|Figura|Gr[aá]fico|Cuadro|Fuente|Nota)\b[^*]*\*$", re.I)


def es_pie(texto: str) -> bool:
    """¿Es una nota de tabla/figura tipo *Tabla 1.1. Título. Fuente: ...*?"""
    return bool(_RE_PIE.match(texto.strip()))


def sin_marcas(texto: str) -> str:
    return "".join(t for t, *_ in tokens(texto))


# ---------------------------------------------------------------------------
#  Markdown → bloques
#    ('h', nivel, texto) · ('p', texto) · ('ul', [items]) · ('ol', [items])
#    ('tabla', [[celdas]])
# ---------------------------------------------------------------------------
_RE_H = re.compile(r"^(#{1,6})\s+(.*\S)\s*$")
_RE_UL = re.compile(r"^[-*•]\s+(.*)$")
_RE_OL = re.compile(r"^(\d+)[.)]\s+(.*)$")
_RE_SEP = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$")


def _celdas(linea: str):
    s = linea.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def bloques(texto: str):
    lineas = limpiar(texto).replace("\r", "").split("\n")
    out, par = [], []
    i, n = 0, len(lineas)

    def cerrar():
        if par:
            out.append(("p", " ".join(par)))
            par.clear()

    while i < n:
        crudo = lineas[i]
        s = crudo.strip()
        if not s:
            cerrar()
            i += 1
            continue
        m = _RE_H.match(s)
        if m:
            cerrar()
            out.append(("h", min(len(m.group(1)), 3), m.group(2).strip()))
            i += 1
            continue
        if s.startswith("|"):
            cerrar()
            filas = []
            while i < n and lineas[i].strip().startswith("|"):
                if not _RE_SEP.match(lineas[i].strip()):
                    filas.append(_celdas(lineas[i]))
                i += 1
            if filas:
                out.append(("tabla", filas))
            continue
        m = _RE_UL.match(s)
        if m:
            cerrar()
            items = []
            while i < n:
                t = lineas[i]
                mm = _RE_UL.match(t.strip())
                if mm:
                    items.append(mm.group(1))
                elif t.strip() and t.startswith((" ", "\t")) and items:
                    items[-1] += " " + t.strip()
                else:
                    break
                i += 1
            out.append(("ul", items))
            continue
        m = _RE_OL.match(s)
        if m:
            cerrar()
            items = []
            while i < n:
                t = lineas[i]
                mm = _RE_OL.match(t.strip())
                if mm:
                    items.append(mm.group(2))
                elif t.strip() and t.startswith((" ", "\t")) and items:
                    items[-1] += " " + t.strip()
                else:
                    break
                i += 1
            out.append(("ol", items))
            continue
        par.append(s)
        i += 1
    cerrar()
    return out


def _desplazar(bl, minimo_deseado: int):
    """Mueve los niveles de encabezado para que el menor sea `minimo_deseado`."""
    niveles = [b[1] for b in bl if b[0] == "h"]
    if not niveles:
        return bl
    d = minimo_deseado - min(niveles)
    res = []
    for b in bl:
        if b[0] == "h":
            res.append(("h", max(1, min(3, b[1] + d)), b[2]))
        else:
            res.append(b)
    return res


def preparar(secciones: dict, titulos: dict):
    """Devuelve la lista ordenada de secciones listas para dibujar.

    Cada elemento: {'clave', 'titulo', 'bloques'}. Los encabezados ya vienen
    normalizados: el título de cada sección es nivel 1 y los internos nivel 2-3;
    en el cuerpo, los capítulos son nivel 1.
    """
    resultado = []
    for clave, defecto in SECCIONES:
        titulo = (titulos.get(clave) or defecto).strip() or defecto
        texto = (secciones.get(clave) or "").strip()
        bl = bloques(texto) if texto else [("p", PENDIENTE)]
        if clave == "cuerpo":
            bl = _desplazar(bl, 1)
            if not any(b[0] == "h" and b[1] == 1 for b in bl):
                bl = [("h", 1, titulo)] + bl
        else:
            bl = [("h", 1, titulo)] + _desplazar(bl, 2)
        resultado.append({"clave": clave, "titulo": titulo, "bloques": bl})
    return resultado


def total_palabras(secciones: dict) -> int:
    return sum(palabras(t or "") for t in secciones.values())
