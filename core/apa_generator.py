from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from models.figure import Figure, FigureStyle

# EMU (English Metric Units) per centimetre
_EMU_PER_CM = 360_000


class APAGenerator:
    """
    Inserts APA 7 formatted paragraphs around each image paragraph.

    Final structure per figure:
        Figura X          <- bold, no special style
        Title             <- italic, Word 'Caption' style (recognised by TOF)
        [image paragraph] <- centred, resized if requested
        Nota. text        <- 'Nota.' italic + normal text (optional)
    """

    def __init__(self, style: FigureStyle = None, document=None):
        self.style = style or FigureStyle()
        # Optional python-docx Document used to guarantee the Caption style exists
        self._document = document

    # ------------------------------------------------------------------
    def apply_apa_format(self, figure: Figure):
        para = figure.paragraph_element

        # 1. Centre the image paragraph
        self._set_alignment(para, "center")

        # 2. Resize image if the user supplied custom dimensions
        if figure.image_width_cm > 0 or figure.image_height_cm > 0:
            self._resize_image(para, figure.image_width_cm, figure.image_height_cm)

        # 3. Insert title BEFORE image — italic + Caption style so Word TOF picks it up
        title_text = figure.title if figure.title else "Titulo de la figura"
        title_para = self._make_para(title_text, italic=True)
        self._apply_caption_style(title_para)
        self._apply_space_after(title_para, self.style.space_after_pt)
        para.addprevious(title_para)

        # 4. Insert "Figura X" BEFORE title — bold
        label_para = self._make_para(f"Figura {figure.number}", bold=True)
        self._apply_space_after(label_para, self.style.space_after_pt)
        title_para.addprevious(label_para)

        # 5. Insert note AFTER image (optional)
        if figure.has_note and figure.note:
            note_para = self._make_note_para(figure.note)
            para.addnext(note_para)

    # ------------------------------------------------------------------
    # Paragraph / run builders
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
        """Paragraph: 'Nota.' (italic) followed by normal text."""
        p = OxmlElement("w:p")

        r1 = OxmlElement("w:r")
        r1.append(self._rPr(italic=True))
        t1 = OxmlElement("w:t")
        t1.text = "Nota."
        r1.append(t1)
        p.append(r1)

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
        sz.set(qn("w:val"), str(self.style.font_size * 2))
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(self.style.font_size * 2))
        rPr.append(sz)
        rPr.append(szCs)
        return rPr

    # ------------------------------------------------------------------
    # Paragraph property helpers
    # ------------------------------------------------------------------
    def _apply_caption_style(self, para_el):
        """
        Set w:pStyle to 'Caption' so Word's Table of Figures recognises the title.
        If the document object is available we ensure the style exists first.
        """
        if self._document is not None:
            try:
                self._document.styles["Caption"]
            except KeyError:
                from docx.enum.style import WD_STYLE_TYPE
                from docx.shared import Pt

                s = self._document.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
                s.font.name = self.style.font_name
                s.font.size = Pt(self.style.font_size)
                s.font.italic = True

        pPr = self._get_or_create_pPr(para_el)
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None:
            pStyle = OxmlElement("w:pStyle")
            pPr.insert(0, pStyle)
        pStyle.set(qn("w:val"), "Caption")

    def _apply_space_after(self, para_el, pt: float):
        """Set space-after (in points) on a paragraph element."""
        if pt <= 0:
            return
        pPr = self._get_or_create_pPr(para_el)
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:after"), str(int(pt * 20)))  # twips (1 pt = 20 twips)

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

    @staticmethod
    def _get_or_create_pPr(para_el):
        pPr = para_el.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            para_el.insert(0, pPr)
        return pPr

    # ------------------------------------------------------------------
    # Image resize
    # ------------------------------------------------------------------
    @staticmethod
    def _resize_image(para_element, width_cm: float, height_cm: float):
        """
        Update the EMU dimensions of every inline/anchored drawing in the paragraph.
        Only the axes with a value > 0 are modified (allows width-only or height-only).
        """
        cx = str(int(width_cm * _EMU_PER_CM)) if width_cm > 0 else None
        cy = str(int(height_cm * _EMU_PER_CM)) if height_cm > 0 else None
        if not cx and not cy:
            return

        for drawing in para_element.findall(".//" + qn("w:drawing")):
            # wp:extent — controls the space reserved in the document layout
            extent = drawing.find(".//" + qn("wp:extent"))
            if extent is not None:
                if cx:
                    extent.set("cx", cx)
                if cy:
                    extent.set("cy", cy)

            # a:ext inside a:xfrm — controls the actual picture transform size
            for xfrm in drawing.findall(".//" + qn("a:xfrm")):
                ext = xfrm.find(qn("a:ext"))
                if ext is not None:
                    if cx:
                        ext.set("cx", cx)
                    if cy:
                        ext.set("cy", cy)
