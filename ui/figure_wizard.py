import io
from typing import List

import customtkinter as ctk
from PIL import Image, UnidentifiedImageError
from tkinter import messagebox

from models.figure import Figure


class FigureWizard(ctk.CTkToplevel):
    """Step-by-step wizard for entering title and note for each figure."""

    def __init__(self, parent, figures: List[Figure], style_config):
        super().__init__(parent)
        self.figures = figures
        self.style_config = style_config
        self.current_index = 0
        self.completed = False
        self._img_ref = None  # prevent GC of CTkImage

        self.title("Asistente de Figuras APA 7")
        self.geometry("580x680")
        self.minsize(500, 580)
        self.grab_set()
        self.lift()

        self._build_ui()
        self._load(0)

    # ------------------------------------------------------------------
    def _build_ui(self):
        # ---- Header bar ----
        header = ctk.CTkFrame(self, corner_radius=0)
        header.pack(fill="x")

        self.progress_label = ctk.CTkLabel(
            header, text="", font=ctk.CTkFont(size=15, weight="bold")
        )
        self.progress_label.pack(side="left", padx=20, pady=12)

        self.step_label = ctk.CTkLabel(header, text="", text_color="gray")
        self.step_label.pack(side="right", padx=20)

        # ---- Progress bar ----
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(fill="x", padx=0, pady=0)
        self.progress_bar.set(0)

        # ---- Scrollable content ----
        content = ctk.CTkScrollableFrame(self)
        content.pack(fill="both", expand=True)

        # Image preview
        preview = ctk.CTkFrame(content)
        preview.pack(fill="x", padx=15, pady=(15, 8))

        ctk.CTkLabel(
            preview, text="Vista previa", font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=12, pady=(10, 5))

        self.img_frame = ctk.CTkFrame(
            preview, height=220, fg_color=("gray88", "gray20")
        )
        self.img_frame.pack(fill="x", padx=12, pady=(0, 12))
        self.img_frame.pack_propagate(False)

        self.img_label = ctk.CTkLabel(
            self.img_frame, text="Cargando...", text_color="gray"
        )
        self.img_label.place(relx=0.5, rely=0.5, anchor="center")

        # Title field
        title_frame = ctk.CTkFrame(content)
        title_frame.pack(fill="x", padx=15, pady=(0, 8))

        ctk.CTkLabel(
            title_frame,
            text="Titulo de la figura *",
            font=ctk.CTkFont(weight="bold"),
        ).pack(anchor="w", padx=12, pady=(12, 4))

        ctk.CTkLabel(
            title_frame,
            text="Se aplicara cursiva automaticamente al generar el documento.",
            text_color="gray",
            font=ctk.CTkFont(size=11),
        ).pack(anchor="w", padx=12, pady=(0, 4))

        self.title_entry = ctk.CTkEntry(
            title_frame,
            placeholder_text="Ej: Retorno real de acciones americanas de 1802 a 2012",
        )
        self.title_entry.pack(fill="x", padx=12, pady=(0, 12))

        # Note section
        note_frame = ctk.CTkFrame(content)
        note_frame.pack(fill="x", padx=15, pady=(0, 8))

        self.note_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            note_frame,
            text="Incluir nota",
            variable=self.note_var,
            command=self._toggle_note,
            font=ctk.CTkFont(weight="bold"),
        ).pack(anchor="w", padx=12, pady=(12, 4))

        ctk.CTkLabel(
            note_frame,
            text='"Nota." se agregara en cursiva automaticamente al inicio.',
            text_color="gray",
            font=ctk.CTkFont(size=11),
        ).pack(anchor="w", padx=12, pady=(0, 4))

        self.note_entry = ctk.CTkTextbox(note_frame, height=80, state="disabled")
        self.note_entry.pack(fill="x", padx=12, pady=(0, 12))

        # Validation
        self.warn_label = ctk.CTkLabel(
            content, text="", text_color=("#E07000", "#FFA040")
        )
        self.warn_label.pack(anchor="w", padx=15, pady=(0, 4))

        # ---- Navigation bar ----
        nav = ctk.CTkFrame(self, corner_radius=0)
        nav.pack(fill="x", side="bottom")

        self.prev_btn = ctk.CTkButton(
            nav,
            text="< Anterior",
            command=self._go_prev,
            width=130,
            state="disabled",
            fg_color="gray",
        )
        self.prev_btn.pack(side="left", padx=15, pady=12)

        self.next_btn = ctk.CTkButton(
            nav, text="Siguiente >", command=self._go_next, width=130
        )
        self.next_btn.pack(side="right", padx=15, pady=12)

    # ------------------------------------------------------------------
    def _toggle_note(self):
        self.note_entry.configure(
            state="normal" if self.note_var.get() else "disabled"
        )

    def _show_image(self, data: bytes, fmt: str):
        try:
            img = Image.open(io.BytesIO(data))
            img.thumbnail((520, 200), Image.Resampling.LANCZOS)
            ctk_img = ctk.CTkImage(
                light_image=img, dark_image=img, size=(img.width, img.height)
            )
            self._img_ref = ctk_img
            self.img_label.configure(image=ctk_img, text="")
        except (UnidentifiedImageError, Exception):
            self._img_ref = None
            self.img_label.configure(
                image=None,
                text=f"[Vista previa no disponible — formato {fmt.upper()}]",
            )

    def _load(self, index: int):
        fig = self.figures[index]
        total = len(self.figures)

        self.progress_label.configure(text=f"Figura {fig.number}")
        self.step_label.configure(text=f"{index + 1} / {total}")
        self.progress_bar.set((index + 1) / total)

        self._show_image(fig.image_data, fig.image_format)

        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, fig.title)

        self.note_var.set(fig.has_note)
        self.note_entry.configure(state="normal")
        self.note_entry.delete("1.0", "end")
        self.note_entry.insert("1.0", fig.note)
        if not fig.has_note:
            self.note_entry.configure(state="disabled")

        self.prev_btn.configure(state="normal" if index > 0 else "disabled")
        self.next_btn.configure(
            text="Finalizar" if index == total - 1 else "Siguiente >"
        )
        self.warn_label.configure(text="")

    def _save_current(self):
        fig = self.figures[self.current_index]
        fig.title = self.title_entry.get().strip()
        fig.has_note = self.note_var.get()
        fig.note = (
            self.note_entry.get("1.0", "end-1c").strip() if fig.has_note else ""
        )

    def _go_prev(self):
        self._save_current()
        self.current_index -= 1
        self._load(self.current_index)

    def _go_next(self):
        self._save_current()
        is_last = self.current_index == len(self.figures) - 1

        if is_last:
            untitled = [f for f in self.figures if not f.title]
            if untitled:
                nums = ", ".join(str(f.number) for f in untitled)
                if not messagebox.askyesno(
                    "Figuras sin titulo",
                    f"Las figuras {nums} no tienen titulo.\n"
                    "Se usara 'Titulo de la figura' como marcador.\n\n"
                    "Desea continuar?",
                    parent=self,
                ):
                    return
            self.completed = True
            self.destroy()
        else:
            self.current_index += 1
            self._load(self.current_index)
