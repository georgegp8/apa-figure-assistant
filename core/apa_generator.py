import copy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from models.figure import Figure, FigureStyle

# 240 twips = 1.0 line (OOXML auto line-rule base)
_TWIPS_BASE = 240
# EMU per centimetre
_EMU_PER_CM = 360_000


class APAGenerator:
    """
    Inserts APA 7 formatted content around each image paragraph.

    Final structure per figure:
        ┌─ ONE paragraph (Caption style) ─────────────────────────┐
        │  Figura X   ← bold run                                  │
        │  <line break>                                            │
        │  Title text ← italic run                                │
        └──────────────────────────────────────────────────────────┘
        [image paragraph]  ← centred, resized if requested
        Nota. text         ← 'Nota.' italic + normal (optional)

    Using a SINGLE Caption-style paragraph for the label+title ensures
    Word's built-in "Table of Figures / Tabla de ilustraciones" can
    find all figure captions regardless of the document locale.
    """

    def __init__(self, style: FigureStyle = None, document=None):
        self.style = style or FigureStyle()
        self._document = document  # python-docx Document (optional but recommended)

    # ------------------------------------------------------------------
    def apply_apa_format(self, figure: Figure):
        para = figure.paragraph_element

        # 1. Centre the image paragraph
        self._set_alignment(para, "center")

        # 2. Resize image if the user supplied custom dimensions
        if figure.image_width_cm > 0 or figure.image_height_cm > 0:
            self._resize_image(para, figure.image_width_cm, figure.image_height_cm)

        # 3. Insert combined "Figura X + title" paragraph BEFORE the image.
        #    ONE paragraph = one TOF entry, Caption style = Word recognises it.
        caption_el = self._make_caption_element(figure)
        para.addprevious(caption_el)

        # 4. Insert note AFTER image (optional)
        if figure.has_note and figure.note:
            note_para = self._make_note_para(figure.note)
            para.addnext(note_para)

    # ------------------------------------------------------------------
    # Caption paragraph builder
    # ------------------------------------------------------------------
    def _make_caption_element(self, figure: Figure):
        """
        Build ONE <w:p> element identical to what Word produces via
        Referencias → Insertar título:

          "Figura " (bold text)
          + { SEQ Figura \\* ARABIC }  ← complex field, cached value = figure.number
          + w:br  (visual line break, keeps one paragraph)
          + title  (italic)

        The Caption style + SEQ field make Word show "Actualizar campos" on
        right-click and allow generating a proper "Tabla de ilustraciones"
        with the label "Figura".
        """
        title_text = figure.title if figure.title else "Titulo de la figura"

        p = OxmlElement("w:p")

        # --- Paragraph properties ---
        pPr = OxmlElement("w:pPr")

        style_id = self._caption_style_id()
        if style_id:
            pStyle = OxmlElement("w:pStyle")
            pStyle.set(qn("w:val"), style_id)
            pPr.append(pStyle)

        spacing = OxmlElement("w:spacing")
        line_val = str(int(_TWIPS_BASE * max(self.style.line_spacing, 1.0)))
        spacing.set(qn("w:line"), line_val)
        spacing.set(qn("w:lineRule"), "auto")
        pPr.append(spacing)

        p.append(pPr)

        bold_rPr = self._rPr(bold=True)

        # --- "Figura " text (bold, with trailing space) ---
        r_text = OxmlElement("w:r")
        r_text.append(bold_rPr)
        t_text = OxmlElement("w:t")
        t_text.text = "Figura "
        t_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r_text.append(t_text)
        p.append(r_text)

        # --- SEQ field: { SEQ Figura \* ARABIC } ---
        # begin
        r_begin = OxmlElement("w:r")
        r_begin.append(copy.deepcopy(bold_rPr))
        fc_begin = OxmlElement("w:fldChar")
        fc_begin.set(qn("w:fldCharType"), "begin")
        r_begin.append(fc_begin)
        p.append(r_begin)

        # instruction
        r_instr = OxmlElement("w:r")
        r_instr.append(copy.deepcopy(bold_rPr))
        instr = OxmlElement("w:instrText")
        instr.text = " SEQ Figura \\* ARABIC "
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r_instr.append(instr)
        p.append(r_instr)

        # separate
        r_sep = OxmlElement("w:r")
        r_sep.append(copy.deepcopy(bold_rPr))
        fc_sep = OxmlElement("w:fldChar")
        fc_sep.set(qn("w:fldCharType"), "separate")
        r_sep.append(fc_sep)
        p.append(r_sep)

        # cached value (what is shown before first field update)
        r_val = OxmlElement("w:r")
        r_val.append(copy.deepcopy(bold_rPr))
        t_val = OxmlElement("w:t")
        t_val.text = str(figure.number)
        r_val.append(t_val)
        p.append(r_val)

        # end
        r_end = OxmlElement("w:r")
        r_end.append(copy.deepcopy(bold_rPr))
        fc_end = OxmlElement("w:fldChar")
        fc_end.set(qn("w:fldCharType"), "end")
        r_end.append(fc_end)
        p.append(r_end)

        # --- Line break ---
        r_br = OxmlElement("w:r")
        r_br.append(OxmlElement("w:br"))
        p.append(r_br)

        # --- Title (italic) ---
        r_title = OxmlElement("w:r")
        r_title.append(self._rPr(italic=True))
        t_title = OxmlElement("w:t")
        t_title.text = title_text
        t_title.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r_title.append(t_title)
        p.append(r_title)

        return p

    def _caption_style_id(self) -> str | None:
        """
        Return the XML styleId of the Caption built-in style.

        Per OOXML spec the styleId is always 'Caption' regardless of locale.
        We verify it actually exists in the document before trusting it.
        Falls back to 'Caption' (Word will create it if absent).
        """
        if self._document is None:
            return "Caption"

        # Check by XML styleId attribute (locale-independent)
        for style in self._document.styles:
            sid = style.element.get(qn("w:styleId"), "")
            if sid == "Caption":
                return "Caption"

        # Not found — ensure it is created and return its id
        try:
            from docx.enum.style import WD_STYLE_TYPE

            s = self._document.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
            s.font.name = self.style.font_name
            s.font.size = Pt(self.style.font_size)
            s.font.italic = True
            return "Caption"
        except Exception:
            return "Caption"  # let Word resolve it

    # ------------------------------------------------------------------
    # Note paragraph
    # ------------------------------------------------------------------
    def _make_note_para(self, note_text: str):
        """'Nota.' (italic) followed by normal text."""
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

    # ------------------------------------------------------------------
    # Run properties
    # ------------------------------------------------------------------
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
    # Alignment
    # ------------------------------------------------------------------
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

    # ------------------------------------------------------------------
    # Image resize
    # ------------------------------------------------------------------
    @staticmethod
    def _resize_image(para_element, width_cm: float, height_cm: float):
        """
        Update EMU dimensions in wp:extent and a:xfrm/a:ext.
        Only axes with value > 0 are modified.
        """
        cx = str(int(width_cm * _EMU_PER_CM)) if width_cm > 0 else None
        cy = str(int(height_cm * _EMU_PER_CM)) if height_cm > 0 else None
        if not cx and not cy:
            return

        for drawing in para_element.findall(".//" + qn("w:drawing")):
            extent = drawing.find(".//" + qn("wp:extent"))
            if extent is not None:
                if cx:
                    extent.set("cx", cx)
                if cy:
                    extent.set("cy", cy)

            for xfrm in drawing.findall(".//" + qn("a:xfrm")):
                ext = xfrm.find(qn("a:ext"))
                if ext is not None:
                    if cx:
                        ext.set("cx", cx)
                    if cy:
                        ext.set("cy", cy)
