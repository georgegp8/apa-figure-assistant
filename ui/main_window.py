from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from core.apa_generator import APAGenerator
from core.document_manager import DocumentManager
from core.image_detector import ImageDetector
from core.index_generator import IndexGenerator
from core.style_manager import StyleManager
from models.figure import FigureStyle


class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Asistente de Figuras APA 7")
        self.geometry("720x600")
        self.minsize(600, 500)

        self.doc_manager = DocumentManager()
        self.figures = []
        self.style_config: FigureStyle = StyleManager.load()

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        # ---- Top bar ----
        top = ctk.CTkFrame(self, corner_radius=0)
        top.pack(fill="x")

        ctk.CTkLabel(
            top,
            text="  Asistente de Figuras APA 7",
            font=ctk.CTkFont(size=20, weight="bold"),
            anchor="w",
        ).pack(side="left", padx=20, pady=14)

        ctk.CTkButton(
            top,
            text="Estilo",
            command=self._open_style,
            width=90,
            fg_color="gray40",
            hover_color="gray30",
        ).pack(side="right", padx=15, pady=10)

        # ---- Body ----
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=15)

        # Step 1 — open file
        step1 = ctk.CTkFrame(body)
        step1.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            step1, text="1. Abrir documento Word (.docx)", font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=15, pady=(12, 6))

        row1 = ctk.CTkFrame(step1, fg_color="transparent")
        row1.pack(fill="x", padx=15, pady=(0, 12))

        self.file_var = ctk.StringVar(value="Ningun documento seleccionado")
        ctk.CTkLabel(
            row1,
            textvariable=self.file_var,
            text_color="gray",
            anchor="w",
            wraplength=460,
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            row1, text="Abrir .docx", command=self._open_document, width=130
        ).pack(side="right")

        # Step 2 — images list
        step2 = ctk.CTkFrame(body)
        step2.pack(fill="both", expand=True, pady=(0, 12))

        hdr2 = ctk.CTkFrame(step2, fg_color="transparent")
        hdr2.pack(fill="x", padx=15, pady=(12, 6))

        ctk.CTkLabel(
            hdr2, text="2. Imagenes detectadas", font=ctk.CTkFont(weight="bold")
        ).pack(side="left")

        self.count_lbl = ctk.CTkLabel(
            hdr2,
            text="",
            fg_color="gray40",
            corner_radius=8,
            width=110,
            font=ctk.CTkFont(size=12),
        )
        self.count_lbl.pack(side="left", padx=10)

        self.img_list = ctk.CTkScrollableFrame(step2, height=160)
        self.img_list.pack(fill="both", expand=True, padx=15, pady=(0, 12))

        ctk.CTkLabel(
            self.img_list,
            text="Abre un documento para detectar imagenes.",
            text_color="gray",
        ).pack(pady=20)

        # Step 3 — generate
        step3 = ctk.CTkFrame(body)
        step3.pack(fill="x")

        ctk.CTkLabel(
            step3, text="3. Generar documento con formato APA 7", font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=15, pady=(12, 6))

        row3 = ctk.CTkFrame(step3, fg_color="transparent")
        row3.pack(fill="x", padx=15, pady=(0, 12))

        ctk.CTkLabel(
            row3,
            text="El asistente te guiara para agregar titulo y nota a cada figura.",
            text_color="gray",
            font=ctk.CTkFont(size=12),
        ).pack(side="left")

        self.wizard_btn = ctk.CTkButton(
            row3,
            text="Iniciar Asistente",
            command=self._start_wizard,
            state="disabled",
            width=160,
        )
        self.wizard_btn.pack(side="right")

        # ---- Status bar ----
        self.status_var = ctk.StringVar(value="Listo. Abre un documento .docx para comenzar.")
        ctk.CTkLabel(
            self,
            textvariable=self.status_var,
            text_color="gray",
            font=ctk.CTkFont(size=11),
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 8))

    # ------------------------------------------------------------------
    def _open_document(self):
        path = filedialog.askopenfilename(
            title="Seleccionar documento Word",
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")],
        )
        if not path:
            return
        try:
            self._set_status("Abriendo documento...")
            self.doc_manager.open_document(path)
            self.file_var.set(Path(path).name)

            detector = ImageDetector()
            self.figures = detector.detect_images(self.doc_manager.document)
            self._refresh_list()
            self._set_status(
                f"Documento cargado. {len(self.figures)} imagen(es) encontrada(s)."
            )
        except Exception as exc:
            messagebox.showerror("Error al abrir", str(exc), parent=self)
            self._set_status("Error al abrir el documento.")

    def _refresh_list(self):
        for w in self.img_list.winfo_children():
            w.destroy()

        n = len(self.figures)
        if n == 0:
            self.count_lbl.configure(text=" 0 imagenes ", text_color="orange")
            self.wizard_btn.configure(state="disabled")
            ctk.CTkLabel(
                self.img_list,
                text="No se encontraron imagenes en el documento.",
                text_color="orange",
            ).pack(pady=20)
        else:
            self.count_lbl.configure(
                text=f" {n} imagen(es) ", text_color="lightgreen"
            )
            self.wizard_btn.configure(state="normal")
            for fig in self.figures:
                row = ctk.CTkFrame(self.img_list, fg_color="transparent")
                row.pack(fill="x", pady=2)
                ctk.CTkLabel(
                    row, text=f"  Imagen {fig.number}", anchor="w"
                ).pack(side="left")
                self._pending_badge(row)

    @staticmethod
    def _pending_badge(parent):
        ctk.CTkLabel(
            parent,
            text="pendiente",
            text_color="gray",
            font=ctk.CTkFont(size=11),
        ).pack(side="right", padx=6)

    # ------------------------------------------------------------------
    def _open_style(self):
        from ui.style_config_dialog import StyleConfigDialog

        dlg = StyleConfigDialog(self, self.style_config)
        self.wait_window(dlg)
        if dlg.result:
            self.style_config = dlg.result
            StyleManager.save(self.style_config)

    def _start_wizard(self):
        if not self.figures:
            return
        from ui.figure_wizard import FigureWizard

        wiz = FigureWizard(self, self.figures, self.style_config)
        self.wait_window(wiz)
        if wiz.completed:
            self._apply_and_save()

    # ------------------------------------------------------------------
    def _apply_and_save(self):
        include_index = messagebox.askyesno(
            "Indice de figuras",
            "Desea agregar un indice de figuras al final del documento?",
            parent=self,
        )

        # --- Ask overwrite or save as new file ---
        overwrite = messagebox.askyesnocancel(
            "Guardar documento",
            "Como desea guardar el documento?\n\n"
            "Si  → Sobreescribir el archivo original\n"
            "No  → Guardar como nuevo archivo (sufijo _APA)\n"
            "Cancelar → No guardar",
            parent=self,
        )
        if overwrite is None:
            return  # user cancelled

        output_path = (
            str(self.doc_manager.file_path) if overwrite else None
        )

        try:
            self._set_status("Aplicando formato APA 7...")
            self.update()

            gen = APAGenerator(self.style_config, document=self.doc_manager.document)
            for fig in self.figures:
                gen.apply_apa_format(fig)

            if include_index:
                IndexGenerator(self.style_config).generate_index(
                    self.doc_manager.document, self.figures
                )

            out = self.doc_manager.save_document(output_path)
            self._set_status(f"Guardado: {Path(out).name}")
            self._mark_done()
            messagebox.showinfo(
                "Documento guardado",
                f"El documento ha sido generado exitosamente:\n\n{out}",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Error al generar", str(exc), parent=self)
            self._set_status("Error al generar el documento.")

    def _mark_done(self):
        for w in self.img_list.winfo_children():
            w.destroy()
        for fig in self.figures:
            row = ctk.CTkFrame(self.img_list, fg_color="transparent")
            row.pack(fill="x", pady=2)
            title_preview = (
                fig.title[:45] + "..." if len(fig.title) > 45 else fig.title
            ) or "(sin titulo)"
            ctk.CTkLabel(
                row,
                text=f"  Figura {fig.number}: {title_preview}",
                anchor="w",
            ).pack(side="left")
            ctk.CTkLabel(
                row,
                text="listo",
                text_color="lightgreen",
                font=ctk.CTkFont(size=11),
            ).pack(side="right", padx=6)

    def _set_status(self, msg: str):
        self.status_var.set(msg)
        self.update_idletasks()
