from typing import List

from docx.enum.text import WD_BREAK
from docx.shared import Pt

from models.figure import Figure, FigureStyle


class IndexGenerator:
    """Appends a figure index page at the end of the document."""

    def __init__(self, style: FigureStyle = None):
        self.style = style or FigureStyle()

    def generate_index(self, document, figures: List[Figure]):
        # Page break
        pb = document.add_paragraph()
        pb.add_run().add_break(WD_BREAK.PAGE)

        # Title
        title_para = document.add_paragraph()
        run = title_para.add_run("Indice de Figuras")
        run.bold = True
        run.font.name = self.style.font_name
        run.font.size = Pt(self.style.font_size + 2)
        title_para.paragraph_format.space_after = Pt(12)

        # One entry per figure
        for fig in figures:
            entry = document.add_paragraph()

            r1 = entry.add_run(f"Figura {fig.number}")
            r1.bold = True
            r1.font.name = self.style.font_name
            r1.font.size = Pt(self.style.font_size)

            label = fig.title if fig.title else "(sin titulo)"
            r2 = entry.add_run(f"\t{label}")
            r2.font.name = self.style.font_name
            r2.font.size = Pt(self.style.font_size)

            entry.paragraph_format.space_after = Pt(6)
