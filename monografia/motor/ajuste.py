"""Ajuste del formato para acercarse al número de hojas elegido."""
from .pdf import PERFIL_ESTANDAR, Perfil, construir_pdf
from .texto import total_palabras

TAMANOS = [11.0, 11.5, 12.0, 12.5, 13.0]
INTERLINEADOS = [1.0, 1.15, 1.3, 1.5, 1.75, 2.0]
ESPACIOS = [0.0, 3.0, 6.0, 9.0, 12.0]


def candidatos(tam=None, esp=None, despues=None):
    tams = [tam] if tam else TAMANOS
    esps = [esp] if esp else INTERLINEADOS
    desps = [despues] if despues is not None else ESPACIOS
    lista = [Perfil(float(t), float(e), float(d)) for t in tams for e in esps for d in desps]
    lista.sort(key=lambda p: p.costo)
    return lista


def _buscar(datos, prep, objetivo, opciones, perfiles):
    """Busca en `perfiles` (ordenados de más compacto a más amplio)."""
    cache = {}

    def probar(i):
        if i not in cache:
            cache[i] = construir_pdf(datos, prep, perfiles[i], opciones)
        return cache[i]

    ultimo = len(perfiles) - 1
    pag_min, pag_max = probar(0)[1], probar(ultimo)[1]
    estandar = PERFIL_ESTANDAR.costo

    def clave(i):
        # Más cerca del objetivo; en empate, más cerca del formato académico estándar.
        return (abs(cache[i][1] - objetivo), abs(perfiles[i].costo - estandar))

    if objetivo <= pag_min:
        elegido = 0
    elif objetivo >= pag_max:
        elegido = ultimo
    else:
        lo, hi = 0, ultimo  # páginas(lo) < objetivo <= páginas(hi)
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if probar(mid)[1] < objetivo:
                lo = mid
            else:
                hi = mid
        # Explorar los formatos vecinos (el total no siempre es perfectamente monótono)
        for j in range(max(0, lo - 3), min(ultimo, hi + 5) + 1):
            probar(j)
        elegido = min(cache, key=clave)

    pdf, paginas = cache[elegido]
    return {
        "pdf": pdf, "paginas": paginas, "perfil": perfiles[elegido],
        "minimo": pag_min, "maximo": pag_max, "compilaciones": len(cache),
    }


def ajustar(datos, prep, secciones, objetivo, opciones, tam=None, esp=None, despues=None):
    """Genera el PDF con el formato que más se acerque a `objetivo` páginas.

    Si no logra el número exacto y los capítulos empiezan en página nueva, vuelve a
    intentarlo con los capítulos seguidos (granularidad de una página).

    Devuelve un dict con: pdf, paginas, perfil, estado ('exacto' | 'cercano' | 'corto' | 'largo'),
    objetivo, minimo, maximo, palabras, palabras_ajuste, compilaciones, opciones.
    """
    perfiles = candidatos(tam, esp, despues)
    opciones = dict(opciones)
    r = _buscar(datos, prep, objetivo, opciones, perfiles)
    compilaciones = r["compilaciones"]

    if r["paginas"] != objetivo and opciones.get("capitulo_nueva_pagina", True):
        alt = dict(opciones, capitulo_nueva_pagina=False)
        r2 = _buscar(datos, prep, objetivo, alt, perfiles)
        compilaciones += r2["compilaciones"]
        if abs(r2["paginas"] - objetivo) < abs(r["paginas"] - objetivo):
            r, opciones = r2, alt
            r["capitulos_unidos"] = True
    r.setdefault("capitulos_unidos", False)

    paginas = r["paginas"]
    if paginas == objetivo:
        estado = "exacto"
    elif objetivo < r["minimo"] and paginas == r["minimo"]:
        estado = "largo"       # aun con el formato más compacto sobran páginas
    elif objetivo > r["maximo"] and paginas == r["maximo"]:
        estado = "corto"       # aun con el formato más amplio faltan páginas
    else:
        estado = "cercano"

    palabras = total_palabras(secciones)
    por_pagina = palabras / max(1, paginas - 2)  # sin portada ni índice
    r.update(
        estado=estado, objetivo=objetivo, palabras=palabras, compilaciones=compilaciones,
        opciones=opciones,
        palabras_ajuste=int(abs(objetivo - paginas) * por_pagina) if palabras else 0,
    )
    return r
