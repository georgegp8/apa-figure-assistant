"""
Asistente de Figuras APA 7
--------------------------
Herramienta de escritorio para generar figuras con formato APA 7
en documentos Word (.docx).

Uso:
    python main.py
"""

import customtkinter as ctk


def main():
    ctk.set_appearance_mode("System")   # "Light" | "Dark" | "System"
    ctk.set_default_color_theme("blue")

    from ui.main_window import MainWindow

    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
