from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from models.figure import Figure, FigureStyle


class APAGenerator:
    """
    Inserts APA 7 formatted paragraphs around each image paragraph.

    Final structure per figure:
        Figura X          <- bold
        Title             <- italic
        [image paragraph] <- centered
        Nota. text        <- 'Nota.' italic, rest normal (optional)
    """

    def __init__(self, style: FigureStyle = None):
        self.style = style or FigureStyle()

    def apply_apa_format(self, figure: Figure):
        para = figure.paragraph_element

        # 1. Center the image paragraph
        self._set_alignment(para, "center")

        # 2. Insert title BEFORE image (italic)
        title_text = figure.title if figure.title else "Titulo de la figura"
        title_para = self._make_para(title_text, italic=True)
        para.addprevious(title_para)

        # 3. Insert "Figura X" BEFORE title (bold)
        label_para = self._make_para(f"Figura {figure.number}", bold=True)
        title_para.addprevious(label_para)

        # 4. Insert note AFTER image (optional)
        if figure.has_note and figure.note:
            note_para = self._make_note_para(figure.note)
            para.addnext(note_para)

    # ------------------------------------------------------------------
    def _make_para(self, text: str, bold: bool = False, italic: bool = False):
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        r.append(self._rPr(bold=bold, italic=italic))
        t = OxmlElement("w:t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        p.append(r)
        return p

    def _make_note_para(self, note_text: str):
        """Paragraph with 'Nota.' in italic followed by normal text."""
        p = OxmlElement("w:p")

        # "Nota." — italic
        r1 = OxmlElement("w:r")
        r1.append(self._rPr(italic=True))
        t1 = OxmlElement("w:t")
        t1.text = "Nota."
        r1.append(t1)
        p.append(r1)

        # " <note text>" — normal weight
        r2 = OxmlElement("w:r")
        r2.append(self._rPr())
        t2 = OxmlElement("w:t")
        t2.text = " " + note_text
        t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r2.append(t2)
        p.append(r2)

        return p

    def _rPr(self, bold: bool = False, italic: bool = False):
        rPr = OxmlElement("w:rPr")
        if bold:
            rPr.append(OxmlElement("w:b"))
        if italic:
            rPr.append(OxmlElement("w:i"))

        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), self.style.font_name)
        rFonts.set(qn("w:hAnsi"), self.style.font_name)
        rPr.append(rFonts)

        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(self.style.font_size * 2))  # half-points
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(self.style.font_size * 2))
        rPr.append(sz)
        rPr.append(szCs)
        return rPr

    @staticmethod
    def _set_alignment(para_element, alignment: str):
        pPr = para_element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            para_element.insert(0, pPr)
        jc = pPr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            pPr.append(jc)
        jc.set(qn("w:val"), alignment)
