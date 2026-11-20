# Generador de monografías académicas

Aplicación web (Streamlit) que **crea una monografía a partir de un archivo que usted sube**, le permite
**editarla** y la genera en el **número de hojas que usted elija**, en **PDF** y en **Word editable**.

Estructura de la monografía: Portada · Índice general · Prólogo · Introducción · Cuerpo del trabajo ·
Conclusiones · Anexos · Bibliografía.

## Cómo se usa

1. **Subir archivo** – suba un `.docx`, `.txt` o `.md`. La aplicación reparte el texto en las secciones según
   los títulos que encuentre (en Word, use los estilos *Título 1 / Título 2 / Título 3*).
   También puede cargar el ejemplo o empezar en blanco.
2. **Editar contenido** – corrija los datos de la portada y el texto de cada sección.
3. **Hojas y descarga** – elija el número de hojas. La aplicación ajusta tamaño de letra, interlineado y espaciado
   para acercarse a ese número, y descarga el **PDF**, el **Word** y el **texto (.md)**.
   El `.md` puede volver a subirse para seguir editando otro día.

### Escritura del texto

| Se escribe | Resultado |
|---|---|
| `# Capítulo 1. Título` | Capítulo (aparece en el índice) |
| `## 1.1 Subtítulo` / `### 1.1.1 Apartado` | Secciones |
| `**negrita**`, `*cursiva*` | Énfasis |
| `- elemento` / `1. elemento` | Listas |
| `\| A \| B \|` + línea `\|---\|---\|` | Tablas |
| `*Tabla 1.1. Título. Fuente: ...*` | Nota bajo la tabla |

## Qué hay que saber

- **El PDF cumple exactamente el número de hojas** cuando el texto lo permite. Si el texto es muy corto o muy largo
  para ese número, la aplicación lo dice y calcula cuántas palabras faltan o sobran. No rellena con texto inventado.
- **El Word** usa el mismo formato, pero el número de hojas puede variar un poco según la versión de Word. Al abrirlo,
  acepte actualizar los campos para ver el índice con sus páginas.
- No incluye **imágenes ni fórmulas** (las imágenes de un `.docx` no se importan): agréguelas en el Word descargado.
- Los emoji y alfabetos no latinos no se dibujan en el PDF (salen como `?`); el Word sí los conserva.
- La aplicación procesa los archivos en memoria y no guarda nada en el servidor: **descargue su trabajo** antes de cerrar.
  Al recargar la página se pierde lo no descargado.

## Publicar en Streamlit Community Cloud (cuenta de GitHub)

1. Suba el **contenido** de esta carpeta a la **raíz** de un repositorio de GitHub (al abrirlo deben verse
   `app.py`, `requirements.txt`, `motor/` y `contenido/`).
2. Entre a <https://share.streamlit.io> con su cuenta de GitHub → **Create app**.
3. Elija el repositorio, la rama `main` y en **Main file path** escriba `app.py`
   (si subió la carpeta entera, `monografia/app.py`).
4. **Deploy**. En 1–2 minutos tendrá una URL `https://…streamlit.app`.

Probar en su computadora:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura del proyecto

```
app.py                 Interfaz (subir · editar · generar)
motor/
  importar.py          Lee .docx/.txt/.md y reparte el texto en secciones
  texto.py             Markdown simplificado → bloques
  pdf.py               Generación del PDF (portada, índice real, numeración)
  word.py              Generación del Word editable
  ajuste.py            Búsqueda del formato que da el número de hojas pedido
contenido/             Texto y datos de ejemplo
requirements.txt       Dependencias
.streamlit/config.toml Tema
```

## Versión LaTeX (opcional)

Además se incluye la plantilla LaTeX original (`main.tex`, `capitulos/`, `config/`, `bibliografia/`, `Makefile`)
para quien prefiera compilar con TeX. `make` genera `main.pdf`, y GitHub Actions
(`.github/workflows/compilar.yml`) lo compila en cada `push`.
