"""
foto_a_docx_local.py
Convierte fotos de apuntes en un documento Word (.docx) usando un modelo de IA
que corre EN TU PC con Ollama: gratis, sin límites y sin internet.

TRES MODOS (opción -m):
    transcribir -> pasa el texto de cada foto tal cual al Word (página por página)
    resumen     -> transcribe y arma un resumen de estudio (en el Word va solo el resumen)
    ambos       -> resumen al principio + transcripción completa (por defecto)

USO (desde la terminal, en la carpeta del script):
    python foto_a_docx_local.py apunte.jpg -m transcribir
    python foto_a_docx_local.py fotos_fisica -t "Física I - Unidad 3" -o fisica_u3.docx --md

Requisitos:
    1) Ollama instalado y abierto (ollama.com)
    2) El modelo descargado:   ollama pull qwen2.5vl:7b
    3) pip install ollama python-docx pillow
"""

import argparse
import io
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import ollama
from docx import Document
from docx.shared import Pt
from PIL import Image, ImageOps

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
MODELO = "qwen2.5vl:7b"          # Modelo de Ollama que lee imágenes
LADO_MAX = 1280                  # Lado máximo de la foto: más grande = más lento
CONTEXTO_FOTO = 8192             # Memoria de trabajo del modelo para cada foto (tokens)
CONTEXTO_RESUMEN = 16384         # Memoria de trabajo para el resumen (todo el texto junto)
EXTENSIONES = {".jpg", ".jpeg", ".png", ".webp"}

FORMATO = """Reglas de formato (respetalas exactamente, porque un programa convierte tu respuesta a Word):
- Títulos con "## " y subtítulos con "### ".
- Viñetas con "- ". Sub-viñetas con dos espacios y "- ".
- Listas numeradas con "1. ", "2. ", etc.
- Negrita SOLO con **así**.
- NO uses tablas, ni bloques de código, ni otros símbolos de markdown.
- Fórmulas en texto plano, por ejemplo: F = m * a, x^2 + 3x - 4 = 0, raiz(2)."""

PROMPT_TRANSCRIBIR = f"""Esta es una foto de apuntes, una hoja de libro o un pizarrón.
TRANSCRIBÍ su contenido de forma fiel y completa, en el mismo idioma y en el mismo orden.

- No resumas, no agregues explicaciones ni información que no esté en la foto.
- Mantené la estructura: si hay títulos, listas o pasos numerados, respetalos.
- Si algo no se lee, escribí [ilegible] en lugar de adivinar.
- Si hay un dibujo, gráfico, diagrama o circuito, describilo en una línea así:
  [Diagrama: descripción breve de lo que muestra, con sus valores y rótulos]
- Si hay una tabla, pasala a viñetas (una viñeta por fila).
- Si la foto no tiene texto, respondé solo: [Sin texto]

{FORMATO}

Respondé SOLO con la transcripción, sin comentarios antes ni después."""

PROMPT_RESUMEN = f"""Con el material que te paso (apuntes, hojas de libro o pizarrón, en orden),
armá un RESUMEN DE ESTUDIO en español, claro y bien organizado.

- Primera línea: "# " + un título para el resumen.
- Conservá definiciones, fórmulas y ejemplos importantes.
- Si algo no se lee bien, escribí [ilegible] en lugar de inventar.
- Al final agregá una sección "## Preguntas de repaso" con 5 preguntas sobre el tema.

{FORMATO}

Respondé SOLO con el resumen, sin comentarios antes ni después."""


# ---------------------------------------------------------------------------
# 1) Juntar las imágenes
# ---------------------------------------------------------------------------
def juntar_imagenes(entradas):
    """Recibe archivos y/o carpetas y devuelve la lista de imágenes ordenada por nombre."""
    imagenes = []
    for entrada in entradas:
        ruta = Path(entrada)
        if ruta.is_dir():
            imagenes += sorted(p for p in ruta.iterdir() if p.suffix.lower() in EXTENSIONES)
        elif ruta.is_file() and ruta.suffix.lower() in EXTENSIONES:
            imagenes.append(ruta)
        else:
            print(f"  Aviso: se ignora '{entrada}' (no existe o no es jpg/png/webp)")
    return imagenes


# ---------------------------------------------------------------------------
# 2) Preparar cada imagen
# ---------------------------------------------------------------------------
def preparar_imagen(ruta):
    """Abre la imagen, la endereza, la achica y la devuelve como bytes JPEG."""
    img = Image.open(ruta)
    img = ImageOps.exif_transpose(img)        # corrige la rotación de fotos del celu
    img = img.convert("RGB")
    img.thumbnail((LADO_MAX, LADO_MAX))        # mantiene la proporción
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=90)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# 3) Hablar con el modelo local
# ---------------------------------------------------------------------------
def preguntar(texto, imagen=None, contexto=CONTEXTO_FOTO):
    """Manda un mensaje al modelo de Ollama y devuelve el texto de la respuesta."""
    mensaje = {"role": "user", "content": texto}
    if imagen is not None:
        mensaje["images"] = [imagen]
    respuesta = ollama.chat(
        model=MODELO,
        messages=[mensaje],
        options={"num_ctx": contexto, "temperature": 0.1},  # baja temperatura = más fiel
    )
    if respuesta.done_reason == "length":
        print("    Aviso: la respuesta se cortó por largo.")
    return respuesta.message.content.strip()


def transcribir(imagenes):
    """Transcribe las fotos UNA POR UNA. Devuelve una lista de (nombre, texto)."""
    paginas = []
    for i, ruta in enumerate(imagenes, start=1):
        inicio = time.time()
        print(f"  Transcribiendo {i}/{len(imagenes)}: {ruta.name} ...", end="", flush=True)
        texto = preguntar(PROMPT_TRANSCRIBIR, imagen=preparar_imagen(ruta))
        print(f" {time.time() - inicio:.0f} s")
        paginas.append((ruta.name, texto))
    return paginas


def resumir(paginas):
    """Arma el resumen a partir del texto ya transcripto."""
    print("  Armando el resumen ...", end="", flush=True)
    inicio = time.time()
    material = "\n\n".join(f"--- Página {i}: {nombre} ---\n{texto}"
                           for i, (nombre, texto) in enumerate(paginas, start=1))
    if len(material) > CONTEXTO_RESUMEN * 3:
        print("\n    Aviso: es mucho texto; el modelo puede no leerlo todo. "
              "Probá con menos fotos por resumen.")
    resumen = preguntar(material + "\n\n" + PROMPT_RESUMEN, contexto=CONTEXTO_RESUMEN)
    print(f" {time.time() - inicio:.0f} s")
    return resumen


# ---------------------------------------------------------------------------
# 4) Convertir el texto en Word
# ---------------------------------------------------------------------------
def agregar_texto(parrafo, texto):
    """Agrega texto a un párrafo respetando las partes en **negrita**."""
    for parte in re.split(r"(\*\*[^*]+\*\*)", texto):
        if parte.startswith("**") and parte.endswith("**") and len(parte) > 4:
            parrafo.add_run(parte[2:-2]).bold = True
        elif parte:
            parrafo.add_run(parte)


def volcar(doc, texto, corrimiento=0):
    """Pasa el texto con formato (#, -, 1.) a estilos reales de Word.
    'corrimiento' baja un nivel los títulos para meterlos dentro de otra sección."""
    for linea in texto.splitlines():
        if not linea.strip():
            continue
        limpia = linea.lstrip()
        sangria = len(linea) - len(limpia)

        titulo = re.match(r"^(#{1,6}) (.*)", limpia)
        if titulo:
            nivel = min(len(titulo.group(1)) + corrimiento, 9)
            agregar_texto(doc.add_heading("", level=nivel), titulo.group(2))
        elif limpia.startswith(("- ", "* ")):
            estilo = "List Bullet 2" if sangria >= 2 else "List Bullet"
            agregar_texto(doc.add_paragraph(style=estilo), limpia[2:])
        elif re.match(r"^\d+[.)] ", limpia):
            agregar_texto(doc.add_paragraph(style="List Number"),
                          re.sub(r"^\d+[.)] ", "", limpia))
        else:
            agregar_texto(doc.add_paragraph(), limpia)


def armar_docx(salida, titulo, resumen=None, paginas=None):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    doc.add_heading(titulo, level=0)

    if resumen and paginas:                       # modo "ambos"
        doc.add_heading("Resumen", level=1)
        volcar(doc, resumen, corrimiento=1)
        doc.add_page_break()
        doc.add_heading("Transcripción completa", level=1)
    elif resumen:                                 # modo "resumen"
        volcar(doc, resumen)
    if paginas:
        nivel_pagina = 2 if resumen else 1
        for i, (nombre, texto) in enumerate(paginas, start=1):
            doc.add_heading(f"Página {i} ({nombre})", level=nivel_pagina)
            volcar(doc, texto, corrimiento=nivel_pagina - 1)

    doc.save(salida)


def armar_md(salida, titulo, resumen=None, paginas=None):
    """Guarda lo mismo en texto plano (.md): ideal para pegarlo o subirlo a un chat."""
    partes = [f"# {titulo}"]
    if resumen:
        partes.append(resumen)
    for i, (nombre, texto) in enumerate(paginas or [], start=1):
        partes.append(f"## Página {i} ({nombre})\n\n{texto}")
    salida.write_text("\n\n".join(partes), encoding="utf-8")


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Fotos de apuntes -> documento Word (IA local con Ollama)")
    parser.add_argument("entradas", nargs="+", help="Fotos y/o carpetas con fotos")
    parser.add_argument("-m", "--modo", choices=["transcribir", "resumen", "ambos"],
                        default="ambos", help="Qué hacer con las fotos (por defecto: ambos)")
    parser.add_argument("-o", "--salida", help="Nombre del .docx de salida")
    parser.add_argument("-t", "--titulo", help="Título del documento")
    parser.add_argument("--md", action="store_true", help="Guardar también un .md con el texto")
    args = parser.parse_args()

    imagenes = juntar_imagenes(args.entradas)
    if not imagenes:
        sys.exit("No encontré imágenes válidas (jpg, png o webp).")

    print(f"Modo: {args.modo} | {len(imagenes)} imagen(es)")

    if args.salida:
        salida = Path(args.salida)
    else:
        sello = datetime.now().strftime("%Y%m%d_%H%M")
        salida = imagenes[0].parent / f"{args.modo}_{sello}.docx"
    titulo = args.titulo or "Apuntes"

    try:
        paginas = transcribir(imagenes)
        resumen = resumir(paginas) if args.modo in ("resumen", "ambos") else None
    except ConnectionError:
        sys.exit("\nError: no me pude conectar con Ollama. ¿Está abierto? (buscá su ícono "
                 "al lado del reloj, o corré 'ollama serve' en otra terminal)")
    except ollama.ResponseError as e:
        if "not found" in str(e).lower():
            sys.exit(f"\nError: falta el modelo. Descargalo con:  ollama pull {MODELO}")
        sys.exit(f"\nError de Ollama: {e}")

    if args.modo == "resumen":
        paginas = None                            # en el Word va solo el resumen

    armar_docx(salida, titulo, resumen, paginas)
    print(f"Listo: {salida.resolve()}")

    if args.md:
        salida_md = salida.with_suffix(".md")
        armar_md(salida_md, titulo, resumen, paginas)
        print(f"Texto plano: {salida_md.resolve()}")


if __name__ == "__main__":
    main()
