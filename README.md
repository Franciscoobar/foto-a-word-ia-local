# Foto a Word con IA local

Convierte fotos de apuntes en un documento Word (`.docx`) usando un modelo de visión que corre **en tu propia PC** con [Ollama]. Es gratis, no tiene límites de uso y las fotos no salen de tu computadora.

## Qué hace

| Modo | Resultado |
|---|---|
| `transcribir` | Pasa el texto de cada foto al Word, página por página |
| `resumen` | Arma un resumen de estudio con preguntas de repaso |
| `ambos` (por defecto) | Pone el resumen arriba y la transcripción completa abajo |

- Corrige solo la rotación de las fotos del celular.
- Convierte títulos, viñetas, listas numeradas y negritas en estilos reales de Word.
- Con `--md` guarda además una copia en texto plano (Markdown).

## Requisitos

- Python 3.10 o más nuevo
- [Ollama] instalado y abierto
- Una GPU NVIDIA con unos 8 GB de VRAM (recomendado; sin GPU funciona, pero mucho más lento)

## Instalación

```bash
ollama pull qwen2.5vl:7b
py -m pip install -r requirements.txt
```

## Uso

```bash
# Una foto
py foto_a_docx_local.py apunte.jpg -m transcribir

# Una carpeta entera (las fotos se ordenan por nombre: 01.jpg, 02.jpg...)
py foto_a_docx_local.py fotos_fisica -t "Física I - Unidad 3" -o fisica_u3.docx --md
```

| Opción | Para qué sirve |
|---|---|
| `-m` | Modo: `transcribir`, `resumen` o `ambos` |
| `-t` | Título del documento |
| `-o` | Nombre del archivo `.docx` de salida |
| `--md` | Guardar también una copia en `.md` |

## Configuración

Al principio del script podés cambiar:
- `MODELO`: el modelo de Ollama que se usa (por defecto `qwen2.5vl:7b`).
- `LADO_MAX`: el tamaño máximo de las fotos. Con fotos más chicas va más rápido, pero lee peor.
- `PROMPT_TRANSCRIBIR` / `PROMPT_RESUMEN`: las instrucciones que recibe el modelo.

## Limitaciones

- Con letra a mano desprolija puede equivocarse. Cuando no entiende algo, escribe `[ilegible]` en vez de inventar.
- Los dibujos y diagramas no se copian: los describe en una línea.
- Las fórmulas quedan en texto plano (`F = m * a`).

[Ollama]: https://ollama.com
