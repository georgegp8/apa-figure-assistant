import copy

import customtkinter as ctk

from models.figure import FigureStyle

COMMON_FONTS = [
    "Times New Roman",
    "Arial",
    "Calibri",
    "Cambria",
    "Georgia",
    "Garamond",
    "Palatino Linotype",
    "Book Antiqua",
]
FONT_SIZES = ["10", "11", "12", "14", "16", "18"]

# Match Word's standard "Interlineado" dropdown values
LINE_SPACING_OPTIONS = ["1.0", "1.15", "1.5", "2.0", "2.5", "3.0"]


class StyleConfigDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_style: FigureStyle):
        super().__init__(parent)
        self.result: FigureStyle | None = None
        self._style = copy.copy(current_style)

        self.title("Configuracion de Estilo")
        self.geometry("480x380")
        self.resizable(False, False)
        self.grab_set()
        self.lift()

        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self,
            text="Configuracion de Estilo APA",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(pady=(20, 15))

        form = ctk.CTkFrame(self)
        form.pack(fill="x", padx=20)
        form.grid_columnconfigure(1, weight=1)

        # --- Font ---
        ctk.CTkLabel(form, text="Fuente:", anchor="w").grid(
            row=0, column=0, sticky="w", padx=15, pady=10
        )
        self.font_combo = ctk.CTkComboBox(form, values=COMMON_FONTS, width=230)
        self.font_combo.set(self._style.font_name)
        self.font_combo.grid(row=0, column=1, padx=15, pady=10, sticky="w")

        # --- Font size ---
        ctk.CTkLabel(form, text="Tamano:", anchor="w").grid(
            row=1, column=0, sticky="w", padx=15, pady=10
        )
        self.size_combo = ctk.CTkComboBox(form, values=FONT_SIZES, width=100)
        self.size_combo.set(str(self._style.font_size))
        self.size_combo.grid(row=1, column=1, padx=15, pady=10, sticky="w")

        # --- Interlineado (Word-standard multipliers) ---
        ctk.CTkLabel(form, text="Interlineado:", anchor="w").grid(
            row=2, column=0, sticky="w", padx=15, pady=10
        )
        spacing_col = ctk.CTkFrame(form, fg_color="transparent")
        spacing_col.grid(row=2, column=1, padx=15, pady=10, sticky="w")

        self.spacing_combo = ctk.CTkComboBox(
            spacing_col, values=LINE_SPACING_OPTIONS, width=100
        )
        # Show the closest stored value
        stored = str(self._style.line_spacing)
        self.spacing_combo.set(stored if stored in LINE_SPACING_OPTIONS else "1.0")
        self.spacing_combo.pack(side="left")

        ctk.CTkLabel(
            spacing_col,
            text="  (se aplica al bloque Figura X / titulo)",
            text_color="gray",
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        # --- Info note ---
        ctk.CTkLabel(
            form,
            text="Numero de figura: Negrita  |  Titulo: Cursiva  |  Nota.: Cursiva",
            text_color="gray",
            font=ctk.CTkFont(size=11),
        ).grid(row=3, column=0, columnspan=2, padx=15, pady=(0, 12))

        # --- Caption style tip ---
        ctk.CTkLabel(
            form,
            text="El titulo usa el estilo 'Caption' de Word para que el\nindice de figuras lo reconozca automaticamente.",
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=11),
            justify="left",
        ).grid(row=4, column=0, columnspan=2, padx=15, pady=(0, 10))

        # --- Buttons ---
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)

        ctk.CTkButton(
            btn_frame,
            text="Cancelar",
            command=self.destroy,
            fg_color="gray",
            width=110,
        ).pack(side="left")
        ctk.CTkButton(
            btn_frame, text="Guardar", command=self._save, width=110
        ).pack(side="right")

    def _save(self):
        try:
            self._style.font_name = self.font_combo.get()
            self._style.font_size = int(self.size_combo.get())
            self._style.line_spacing = float(self.spacing_combo.get())
            self.result = self._style
        except ValueError:
            pass
        self.destroy()
