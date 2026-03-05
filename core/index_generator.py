from typing import List

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from models.figure import Figure, FigureStyle


class IndexGenerator:
    """
    Inserts a Word-native Table of Figures field at the BEGINNING of the document.

    Structure inserted (before the original first paragraph):

        "Indice de Figuras"   ← heading paragraph
        { TOC \\h \\z \\c "Figura" }  ← TOF field, pre-populated with entries
        <page break paragraph>  ← separates index from document body

    The TOF is pre-populated so it displays immediately without requiring
    "Actualizar campos" and WITHOUT triggering the Word security dialog
    (no w:updateFields in the document settings).
    Users can still right-click → "Actualizar campos" at any time to refresh
    page numbers after editing the document.
    """

    def __init__(self, style: FigureStyle = None):
        self.style = style or FigureStyle()

    # ------------------------------------------------------------------
    def generate_index(self, document, figures: List[Figure]):
        body = document.element.body

        # Safeguard: the very first body child to use as insertion anchor
        original_first = body[0]

        # 1. Heading paragraph  ─────────────────────────────────────
        heading_el = self._make_heading_el(document)
        original_first.addprevious(heading_el)

        # 2. TOF field (pre-populated, multi-paragraph) ─────────────
        #    Insert each paragraph of the field after the heading.
        tof_paragraphs = self._build_tof_paragraphs(figures)
        prev = heading_el
        for p_el in tof_paragraphs:
            prev.addnext(p_el)
            prev = p_el

        # 3. Page break — separates index from document body ────────
        pb_el = self._make_page_break_el()
        prev.addnext(pb_el)

    # ------------------------------------------------------------------
    # Element builders
    # ------------------------------------------------------------------
    def _make_heading_el(self, document):
        """Heading 'Indice de Figuras' using TOC Heading style if available."""
        p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")

        style_id = self._toc_heading_style_id(document)
        if style_id:
            pStyle = OxmlElement("w:pStyle")
            pStyle.set(qn("w:val"), style_id)
            pPr.append(pStyle)

        p.append(pPr)

        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        b = OxmlElement("w:b")
        rPr.append(b)
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), self.style.font_name)
        rFonts.set(qn("w:hAnsi"), self.style.font_name)
        rPr.append(rFonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str((self.style.font_size + 2) * 2))
        rPr.append(sz)
        r.append(rPr)

        t = OxmlElement("w:t")
        t.text = "Indice de Figuras"
        r.append(t)
        p.append(r)
        return p

    def _build_tof_paragraphs(self, figures: List[Figure]) -> list:
        """
        Build a multi-paragraph pre-populated TOF field:

            <w:p>  begin + instrText + separate + FIRST ENTRY  </w:p>
            <w:p>  ENTRY 2  </w:p>
            ...
            <w:p>  end  </w:p>

        Each entry shows: "Figura N  Title"
        Page numbers are omitted (no bookmarks available at generation time).
        Right-click → Actualizar campos populates them.
        """
        if not figures:
            # Empty field: begin ... separate ... end in one paragraph
            p = OxmlElement("w:p")
            self._append_field_begin(p)
            self._append_field_separate(p)
            self._append_field_end(p)
            return [p]

        paragraphs = []

        for idx, fig in enumerate(figures):
            p = OxmlElement("w:p")

            # Apply TableOfFigures style to entry paragraphs
            pPr = OxmlElement("w:pPr")
            pStyle = OxmlElement("w:pStyle")
            pStyle.set(qn("w:val"), "TableOfFigures")
            pPr.append(pStyle)
            p.append(pPr)

            if idx == 0:
                # First paragraph carries begin + instrText + separate
                self._append_field_begin(p)
                self._append_field_separate(p)

            # Entry text: "Figura N  Title"
            title = fig.title if fig.title else "(sin titulo)"
            r = OxmlElement("w:r")
            rPr = self._entry_rPr()
            r.append(rPr)
            t = OxmlElement("w:t")
            t.text = f"Figura {fig.number}"
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(t)
            p.append(r)

            # Tab separator
            r_tab = OxmlElement("w:r")
            r_tab.append(OxmlElement("w:tab"))
            p.append(r_tab)

            # Title
            r_title = OxmlElement("w:r")
            r_title.append(self._entry_rPr())
            t_title = OxmlElement("w:t")
            t_title.text = title
            t_title.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r_title.append(t_title)
            p.append(r_title)

            paragraphs.append(p)

        # Last paragraph carries field end
        p_end = OxmlElement("w:p")
        self._append_field_end(p_end)
        paragraphs.append(p_end)

        return paragraphs

    @staticmethod
    def _make_page_break_el():
        """Paragraph with a page break to separate the index from the body."""
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        noProof = OxmlElement("w:noProof")
        rPr.append(noProof)
        r.append(rPr)
        pb = OxmlElement("w:br")
        pb.set(qn("w:type"), "page")
        r.append(pb)
        p.append(r)
        return p

    # ------------------------------------------------------------------
    # Field character helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _append_field_begin(para_el):
        r = OxmlElement("w:r")
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), "begin")
        fc.set(qn("w:dirty"), "1")   # marks as stale → user can update page numbers
        r.append(fc)
        para_el.append(r)

        r2 = OxmlElement("w:r")
        instr = OxmlElement("w:instrText")
        instr.text = ' TOC \\h \\z \\c "Figura" '
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r2.append(instr)
        para_el.append(r2)

    @staticmethod
    def _append_field_separate(para_el):
        r = OxmlElement("w:r")
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), "separate")
        r.append(fc)
        para_el.append(r)

    @staticmethod
    def _append_field_end(para_el):
        r = OxmlElement("w:r")
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), "end")
        r.append(fc)
        para_el.append(r)

    # ------------------------------------------------------------------
    def _entry_rPr(self):
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), self.style.font_name)
        rFonts.set(qn("w:hAnsi"), self.style.font_name)
        rPr.append(rFonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(self.style.font_size * 2))
        rPr.append(sz)
        return rPr

    # ------------------------------------------------------------------
    @staticmethod
    def _toc_heading_style_id(document) -> str | None:
        """Find TOC Heading style by XML ID (locale-independent)."""
        for candidate in ("TOCHeading", "TOC Heading"):
            for style in document.styles:
                if style.element.get(qn("w:styleId"), "") == candidate:
                    return candidate
        return None
