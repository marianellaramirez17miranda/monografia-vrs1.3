"""Generador de monografías académicas — Streamlit.

Flujo: 1) subir un archivo (.docx, .txt, .md) o escribir  →  2) editar  →
       3) elegir el número de hojas  →  4) descargar PDF y Word editable.

Ejecutar en local:  streamlit run app.py
"""
import hashlib
import io
import json
from pathlib import Path

import streamlit as st

from motor.ajuste import ajustar
from motor.importar import exportar_markdown, leer_archivo, separar_secciones
from motor.pdf import PERFIL_ESTANDAR, Perfil, construir_pdf
from motor.texto import SECCIONES, TITULOS_DEFECTO, no_soportados, palabras, preparar, total_palabras
from motor.word import construir_docx

BASE = Path(__file__).parent
CONTENIDO = BASE / "contenido"

CAMPOS = [
    ("titulo", "Título de la monografía"),
    ("autor", "Autor"),
    ("codigo", "Código / datos del autor"),
    ("universidad", "Universidad"),
    ("facultad", "Facultad"),
    ("escuela", "Escuela / carrera"),
    ("curso", "Curso"),
    ("docente", "Docente"),
    ("ciclo", "Ciclo / semestre"),
    ("ciudad", "Ciudad, país"),
    ("anio", "Año"),
]
ARCHIVOS_EJEMPLO = {
    "prologo": "01_prologo.md",
    "introduccion": "02_introduccion.md",
    "cuerpo": "03_cuerpo.md",
    "conclusiones": "04_conclusiones.md",
    "anexos": "05_anexos.md",
    "bibliografia": "06_bibliografia.md",
}

st.set_page_config(page_title="Generador de monografías", page_icon="📄", layout="wide")


# ---------------------------------------------------------------------------
#  Estado
# ---------------------------------------------------------------------------
def _sin_titulo(texto: str) -> str:
    lineas = texto.strip().split("\n")
    if lineas and lineas[0].startswith("# "):
        lineas = lineas[1:]
    return "\n".join(lineas).strip()


def _poner_ejemplo() -> None:
    try:
        datos = json.loads((CONTENIDO / "datos.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        datos = {}
    for k, _ in CAMPOS:
        st.session_state[f"dat_{k}"] = datos.get(k, "")
    for clave, nombre in ARCHIVOS_EJEMPLO.items():
        ruta = CONTENIDO / nombre
        texto = ruta.read_text(encoding="utf-8") if ruta.exists() else ""
        st.session_state[f"sec_{clave}"] = _sin_titulo(texto)
        st.session_state[f"tit_{clave}"] = TITULOS_DEFECTO[clave]


def _iniciar() -> None:
    if "iniciado" not in st.session_state:
        _poner_ejemplo()
        st.session_state["iniciado"] = True
        st.session_state["avisos"] = []


def cb_ejemplo() -> None:
    _poner_ejemplo()
    st.session_state["avisos"] = [("info", "Se cargó el contenido de ejemplo.")]


def cb_vaciar() -> None:
    for k, _ in CAMPOS:
        st.session_state[f"dat_{k}"] = ""
    for clave, titulo in SECCIONES:
        st.session_state[f"sec_{clave}"] = ""
        st.session_state[f"tit_{clave}"] = titulo
    st.session_state["avisos"] = [("info", "Documento en blanco. Escriba en «Editar contenido» o suba un archivo.")]


def cb_importar() -> None:
    f = st.session_state.get("archivo")
    if f is None:
        st.session_state["avisos"] = [("warning", "Primero seleccione un archivo.")]
        return
    try:
        md, avisos = leer_archivo(f.name, f.getvalue())
        secciones, titulos, titulo, av2 = separar_secciones(md)
    except Exception as e:  # archivo dañado, formato no soportado…
        st.session_state["avisos"] = [("error", f"No se pudo leer el archivo: {e}")]
        return
    if not md.strip():
        st.session_state["avisos"] = [("warning", "El archivo no contiene texto.")]
        return
    conservar = st.session_state.get("conservar", False)
    for clave, defecto in SECCIONES:
        if secciones.get(clave) or not conservar:
            st.session_state[f"sec_{clave}"] = secciones.get(clave, "")
            st.session_state[f"tit_{clave}"] = titulos.get(clave, defecto)
    if titulo:
        st.session_state["dat_titulo"] = titulo
    cargadas = [f"{TITULOS_DEFECTO[k]} ({palabras(secciones[k])} palabras)" for k, _ in SECCIONES if secciones.get(k)]
    msgs = [("success", f"Archivo «{f.name}» importado. Secciones detectadas: " + ", ".join(cargadas) + ".")]
    msgs += [("warning", a) for a in avisos + av2]
    st.session_state["avisos"] = msgs


def _mostrar_avisos() -> None:
    for tipo, texto in st.session_state.get("avisos", []):
        getattr(st, tipo)(texto)


def _datos() -> dict:
    return {k: st.session_state.get(f"dat_{k}", "") for k, _ in CAMPOS}


def _secciones() -> tuple[dict, dict]:
    secs = {c: st.session_state.get(f"sec_{c}", "") for c, _ in SECCIONES}
    tits = {c: st.session_state.get(f"tit_{c}", t) for c, t in SECCIONES}
    return secs, tits


def _firma(extra) -> str:
    secs, tits = _secciones()
    crudo = json.dumps([_datos(), secs, tits, extra], sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.md5(crudo.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
#  Vista previa del PDF
# ---------------------------------------------------------------------------
def _vista_previa(pdf: bytes, paginas: int) -> None:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        st.caption("Vista previa no disponible; descargue el PDF para verlo.")
        return
    pagina = st.number_input("Página a visualizar", 1, max(1, paginas), 1, key=f"pag_prev_{paginas}")
    try:
        doc = pdfium.PdfDocument(pdf)
        img = doc[int(pagina) - 1].render(scale=1.4).to_pil()
        st.image(img)
    except Exception:
        st.caption("No se pudo dibujar la vista previa; descargue el PDF para verlo.")


# ---------------------------------------------------------------------------
#  Interfaz
# ---------------------------------------------------------------------------
_iniciar()

st.title("📄 Generador de monografías académicas")
st.caption(
    "Suba su texto, edítelo, elija cuántas hojas debe tener y descargue la monografía en PDF y en Word editable."
)

with st.sidebar:
    st.header("Estructura")
    st.markdown(
        "1. Portada\n2. Índice general\n3. Prólogo\n4. Introducción\n5. Cuerpo del trabajo\n"
        "6. Conclusiones\n7. Anexos\n8. Bibliografía"
    )
    st.divider()
    st.subheader("Cómo escribir el texto")
    st.markdown(
        "- `# Capítulo 1. Título` → capítulo\n"
        "- `## 1.1 Subtítulo` → sección\n"
        "- `### 1.1.1 Apartado`\n"
        "- `**negrita**` y `*cursiva*`\n"
        "- `- elemento` → lista con viñetas\n"
        "- `1. elemento` → lista numerada\n"
        "- Tablas: `| A | B |` con una línea `|---|---|` bajo el encabezado\n"
        "- Un párrafo se separa del siguiente con una línea en blanco\n"
        "- Una nota bajo una tabla: `*Tabla 1.1. Título. Fuente: ...*`\n\n"
        "**No incluye** fórmulas ni imágenes: agréguelas en el Word descargado."
    )

tab_subir, tab_editar, tab_generar = st.tabs(
    ["📤 1 · Subir archivo", "✏️ 2 · Editar contenido", "📑 3 · Hojas y descarga"]
)

# ---------------------- 1 · Subir ----------------------
with tab_subir:
    st.subheader("Suba su documento")
    st.write(
        "Acepta **Word (.docx)**, **texto (.txt)** y **Markdown (.md)**. La aplicación reparte el texto en "
        "Prólogo, Introducción, Cuerpo, Conclusiones, Anexos y Bibliografía según los títulos que encuentre "
        "(use estilos de *Título 1/2/3* en Word). Lo que no se reconozca va al cuerpo del trabajo."
    )
    st.file_uploader("Archivo", type=["docx", "txt", "md"], key="archivo", label_visibility="collapsed")
    st.checkbox("Conservar mis secciones actuales cuando el archivo no traiga esa sección", key="conservar")
    c1, c2, c3 = st.columns(3)
    c1.button("📥 Importar archivo", on_click=cb_importar, type="primary")
    c2.button("🧪 Cargar ejemplo", on_click=cb_ejemplo)
    c3.button("🗑️ Empezar en blanco", on_click=cb_vaciar)
    _mostrar_avisos()

    secs, _ = _secciones()
    st.markdown("**Contenido actual**")
    st.table(
        {
            "Sección": [TITULOS_DEFECTO[c] for c, _ in SECCIONES],
            "Palabras": [palabras(secs[c]) for c, _ in SECCIONES],
        }
    )
    st.caption(
        "Las imágenes de un .docx no se importan (solo texto y tablas). "
        "Tras importar, revise y corrija en la pestaña «Editar contenido»."
    )

# ---------------------- 2 · Editar ----------------------
with tab_editar:
    with st.expander("Datos de la portada", expanded=False):
        izq, der = st.columns(2)
        for i, (clave, etiqueta) in enumerate(CAMPOS):
            (izq if i % 2 == 0 else der).text_input(etiqueta, key=f"dat_{clave}")

    st.write("Edite cada sección. El índice general se genera solo a partir de los títulos.")
    for clave, titulo in SECCIONES:
        with st.expander(titulo, expanded=(clave == "introduccion")):
            st.text_input("Título de la sección", key=f"tit_{clave}")
            st.text_area(
                "Texto", key=f"sec_{clave}", height=340, label_visibility="collapsed",
                help="Use # para capítulos, ## para subtítulos. Vea la guía en el panel lateral.",
            )
            st.caption(f"{palabras(st.session_state.get(f'sec_{clave}', ''))} palabras")

# ---------------------- 3 · Generar ----------------------
with tab_generar:
    secs, tits = _secciones()
    datos = _datos()
    st.write(f"Texto total: **{total_palabras(secs)} palabras**.")
    raros = no_soportados(*secs.values(), *tits.values(), *datos.values())
    if raros:
        st.warning(
            "Estos caracteres no se pueden dibujar en el PDF y aparecerán como «?»: "
            + " ".join(raros[:20]) + ". El archivo Word sí los conserva."
        )

    modo = st.radio(
        "Formato",
        ["Ajustar al número de hojas", "Elegir formato manualmente"],
        horizontal=True, key="modo",
    )
    hojas = st.number_input(
        "Número de hojas (páginas) de la monografía",
        min_value=3, max_value=500, value=30, step=1, key="hojas",
        help="Cuenta la portada y el índice. La aplicación ajusta tamaño de letra, interlineado y espaciado.",
    )

    perfil_manual = None
    if modo == "Elegir formato manualmente":
        a, b, c = st.columns(3)
        tam = a.selectbox("Tamaño de letra (pt)", [10.0, 11.0, 12.0, 13.0, 14.0], index=2)
        esp = b.selectbox("Interlineado", [1.0, 1.15, 1.5, 2.0, 2.5], index=2)
        desp = c.selectbox("Espacio entre párrafos (pt)", [0.0, 3.0, 6.0, 9.0, 12.0], index=2)
        perfil_manual = Perfil(float(tam), float(esp), float(desp))

    with st.expander("Opciones avanzadas"):
        nueva_pag = st.checkbox("Cada capítulo empieza en página nueva", value=True, key="op_cap")
        toc_sub = st.checkbox("Mostrar subtítulos en el índice general", value=True, key="op_toc")
        fijar = None
        if modo == "Ajustar al número de hojas":
            fijar = st.selectbox("Tamaño de letra", ["Automático", 11.0, 12.0, 13.0], key="op_tam")

    opciones = {"capitulo_nueva_pagina": nueva_pag, "toc_subtitulos": toc_sub}
    extra = [modo, int(hojas), opciones, str(perfil_manual), str(fijar)]

    if st.button("⚙️ Generar monografía", type="primary"):
        prep = preparar(secs, tits)
        with st.spinner("Generando…"):
            if modo == "Ajustar al número de hojas":
                tam_fijo = None if fijar in (None, "Automático") else float(fijar)
                r = ajustar(datos, prep, secs, int(hojas), opciones, tam=tam_fijo)
            else:
                pdf, n = construir_pdf(datos, prep, perfil_manual, opciones)
                r = {
                    "pdf": pdf, "paginas": n, "perfil": perfil_manual, "opciones": opciones,
                    "estado": "exacto" if n == int(hojas) else "manual", "objetivo": int(hojas),
                    "palabras": total_palabras(secs),
                    "palabras_ajuste": 0, "capitulos_unidos": False,
                }
                if r["palabras"] and r["estado"] == "manual":
                    por_pag = r["palabras"] / max(1, n - 2)
                    r["palabras_ajuste"] = int(abs(int(hojas) - n) * por_pag)
            r["docx"] = construir_docx(datos, prep, r["perfil"], r["opciones"])
            r["md"] = exportar_markdown(datos.get("titulo", ""), secs, tits).encode("utf-8")
            r["firma"] = _firma(extra)
            st.session_state["res"] = r

    r = st.session_state.get("res")
    if r:
        if r["firma"] != _firma(extra):
            st.warning("El contenido o las opciones cambiaron desde la última generación. Pulse «Generar» otra vez.")
        p = r["perfil"]
        m1, m2, m3 = st.columns(3)
        m1.metric("Hojas obtenidas", r["paginas"])
        m2.metric("Hojas pedidas", r["objetivo"])
        m3.metric("Formato", f"{p.tam:g} pt · int. {p.esp:g}")

        est, falta = r["estado"], r["palabras_ajuste"]
        if r["palabras"] == 0:
            st.warning(
                "El documento todavía no tiene texto: las hojas generadas son solo portada, índice y títulos. "
                "Suba un archivo o escriba en «Editar contenido»."
            )
        elif est == "exacto":
            st.success(f"La monografía tiene exactamente {r['paginas']} hojas.")
        elif est == "cercano":
            st.info(
                f"No se logró exactamente {r['objetivo']} hojas con un formato razonable; "
                f"se generó la versión más cercana: {r['paginas']} hojas."
            )
        elif est == "corto":
            st.warning(
                f"Aun con el formato más amplio, el texto llega a {r['paginas']} hojas. "
                f"Para alcanzar {r['objetivo']} faltan unas {falta} palabras de contenido propio."
            )
        elif est == "largo":
            st.warning(
                f"Aun con el formato más compacto, el texto ocupa {r['paginas']} hojas. "
                f"Para llegar a {r['objetivo']} habría que recortar unas {falta} palabras."
            )
        else:
            st.info(f"Con este formato la monografía tiene {r['paginas']} hojas (pedidas: {r['objetivo']}).")
        if r.get("capitulos_unidos"):
            st.caption("Para acercarse al número de hojas, los capítulos se colocaron seguidos (sin salto de página).")

        d1, d2, d3 = st.columns(3)
        d1.download_button("⬇️ Descargar PDF", r["pdf"], "monografia.pdf", "application/pdf")
        d2.download_button(
            "⬇️ Descargar Word (editable)", r["docx"], "monografia.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        d3.download_button("⬇️ Descargar texto (.md)", r["md"], "monografia.md", "text/markdown")
        st.caption(
            "El PDF tiene exactamente el número de hojas indicado. En Word el número puede variar "
            "ligeramente según la versión; al abrirlo, acepte actualizar los campos para ver el índice "
            "con sus páginas. El archivo .md puede volver a subirse en la pestaña 1 para seguir editando."
        )
        _vista_previa(r["pdf"], r["paginas"])
