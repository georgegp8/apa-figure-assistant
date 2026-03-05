# Asistente de Figuras APA 7

Herramienta de escritorio para detectar imágenes en documentos Word y generar automáticamente el formato de figura APA 7.

---

## Requisitos

- Python 3.10 o superior
- Windows (probado en Windows 10/11)

---

## Instalación

```bat
# 1. Clonar o descargar el proyecto
cd APA-Figure-Numbering

# 2. Crear entorno virtual
python -m venv .venv

# 3. Activar entorno virtual
.venv\Scripts\activate

# 4. Instalar dependencias
pip install -r requirements.txt
```

---

## Uso

**Opción A — desde terminal (con venv activo):**
```bat
.venv\Scripts\activate
python main.py
```

**Opción B — doble clic:**
```
run.bat
```

---

## Flujo de trabajo

```
1. Abrir documento .docx
        ↓
2. Detección automática de imágenes
        ↓
3. Asistente: título y nota por cada figura
        ↓
4. Generación del documento con formato APA 7
        ↓
5. Índice de figuras (opcional)
        ↓
6. Archivo guardado como nombrearchivo_APA.docx
```

---

## Formato APA 7 generado

```
Figura 1                          ← negrita
Título de la figura en cursiva    ← cursiva

        [imagen centrada]

Nota. Descripción de la figura.   ← "Nota." cursiva, resto normal
```

---

## Estructura del proyecto

```
APA-Figure-Numbering/
├── main.py                    # Punto de entrada
├── run.bat                    # Lanzador rápido (Windows)
├── requirements.txt
├── .vscode/
│   └── settings.json          # Configuración del intérprete para VS Code
├── models/
│   └── figure.py              # Dataclasses: Figure, FigureStyle
├── core/
│   ├── document_manager.py    # Abrir y guardar .docx
│   ├── image_detector.py      # Detectar imágenes en el documento
│   ├── apa_generator.py       # Aplicar formato APA 7 vía XML
│   ├── style_manager.py       # Persistir configuración de estilo
│   └── index_generator.py     # Generar índice de figuras
└── ui/
    ├── main_window.py         # Ventana principal
    ├── figure_wizard.py       # Asistente paso a paso
    └── style_config_dialog.py # Diálogo de configuración de estilo
```

---

## Dependencias

| Paquete | Versión mínima | Uso |
|---|---|---|
| `python-docx` | 1.1.0 | Manipulación de archivos .docx |
| `customtkinter` | 5.2.0 | Interfaz gráfica moderna |
| `Pillow` | 10.0.0 | Vista previa de imágenes en el asistente |
| `lxml` | 5.0.0 | Procesamiento XML del documento |

---

## Notas

- El documento original **no se modifica**; se genera una copia con sufijo `_APA`.
- Se procesan imágenes inline en el cuerpo principal del documento (no dentro de tablas).
- La configuración de estilo (fuente y tamaño) se guarda en `~/.apa_figure_config.json`.
