# Asistente de Figuras APA 7

Herramienta de escritorio para Windows que detecta imágenes en documentos Word (`.docx`) y aplica automáticamente el formato de figura según las normas **APA 7.ª edición**, incluyendo rótulo numerado, título, nota opcional e índice de figuras.

---

## Características

- **Detección automática** de todas las imágenes inline en el documento.
- **Formato APA 7** generado por párrafo:
  - `Figura X` en negrita (campo `SEQ Figura` nativo de Word).
  - Título en cursiva (estilo `Caption` reconocido por Word).
  - Imagen centrada con dimensiones personalizables.
  - `Nota.` opcional (cursiva) seguida del texto descriptivo.
- **Índice de figuras** nativo (`TOC \c "Figura"`) insertado al inicio del documento, actualizable con clic derecho → *Actualizar campos*.
- **Vista previa** de cada imagen con ventana de zoom.
- **Configuración de estilo** persistente: fuente, tamaño e interlineado.
- **Guarda sin sobrescribir** el original (sufijo `_APA`) o permite sobreescritura explícita.

---

## Requisitos

| Requisito | Versión mínima |
|---|---|
| Python | 3.10 |
| Sistema operativo | Windows 10 / 11 |

---

## Instalación

```bat
:: 1. Clonar o descargar el proyecto
cd APA-Figure-Numbering

:: 2. Crear entorno virtual
python -m venv .venv

:: 3. Activar entorno virtual
.venv\Scripts\activate

:: 4. Instalar dependencias
pip install -r requirements.txt
```

---

## Ejecución

**Opción A — doble clic (recomendado):**
```
run.bat
```

**Opción B — desde terminal con entorno virtual activo:**
```bat
.venv\Scripts\activate
python main.py
```

---

## Flujo de trabajo

```
1. Abrir documento .docx
        ↓
2. Detección automática de imágenes
        ↓
3. Asistente figura por figura
   · Título
   · Nota (opcional)
   · Dimensiones personalizadas (opcional)
        ↓
4. Seleccionar si incluir índice de figuras
        ↓
5. Guardar como <nombre>_APA.docx  (o sobrescribir)
```

---

## Formato APA 7 aplicado

```
Figura 1                               ← negrita, campo SEQ de Word
Título de la figura en cursiva         ← cursiva, estilo Caption

        [ imagen centrada ]

Nota. Descripción adicional de la figura.   ← "Nota." cursiva, texto normal
```

El estilo `Caption` y el campo `SEQ Figura` permiten que Word genere automáticamente una *Tabla de ilustraciones* (`Referencias → Insertar tabla de ilustraciones`) y que los números se actualicen al editar el documento.

---

## Índice de figuras

El índice se inserta como un campo Word nativo `{ TOC \h \z \c "Figura" }` al inicio del documento. Se pre-populan las entradas para que sean visibles sin necesidad de actualizar; el clic derecho → *Actualizar campos* sincroniza números de página y aplica el estilo `TableOfFigures` de Word.

> **Nota:** No se activa `w:updateFields` en la configuración del documento, por lo que Word no muestra el diálogo de seguridad al abrir el archivo.

---

## Estructura del proyecto

```
APA-Figure-Numbering/
├── main.py                      # Punto de entrada
├── run.bat                      # Lanzador rápido (Windows)
├── requirements.txt             # Dependencias del proyecto
├── .vscode/
│   └── settings.json            # Intérprete Python para VS Code / Pylance
├── models/
│   └── figure.py                # Dataclasses: Figure, FigureStyle
├── core/
│   ├── document_manager.py      # Abrir y guardar archivos .docx
│   ├── image_detector.py        # Detección de imágenes inline
│   ├── apa_generator.py         # Aplicación del formato APA 7 vía XML
│   ├── style_manager.py         # Persistencia de configuración de estilo
│   └── index_generator.py       # Generación del índice de figuras
└── ui/
    ├── main_window.py           # Ventana principal
    ├── figure_wizard.py         # Asistente paso a paso por figura
    └── style_config_dialog.py   # Diálogo de configuración de estilo
```

---

## Dependencias

| Paquete | Versión mínima | Uso |
|---|---|---|
| `python-docx` | 1.1.0 | Lectura y escritura de archivos `.docx` |
| `customtkinter` | 5.2.0 | Interfaz gráfica moderna |
| `Pillow` | 10.0.0 | Vista previa y zoom de imágenes |
| `lxml` | 5.0.0 | Manipulación XML del documento OOXML |

---

## Limitaciones conocidas

- Solo procesa imágenes **inline** en el cuerpo principal del documento (no imágenes flotantes ni imágenes dentro de tablas).
- Requiere Windows; no probado en macOS ni Linux.
- El índice de figuras se muestra sin puntos guía (`......`) hasta que se ejecuta *Actualizar campos* en Word.

---

## Configuración de estilo

La configuración de fuente, tamaño e interlineado se persiste en:

```
%USERPROFILE%\.apa_figure_config.json
```

Puede editarse desde *Configuración de Estilo* dentro de la aplicación.
