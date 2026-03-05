from typing import List

from docx.oxml.ns import qn

from models.figure import Figure


class ImageDetector:
    """Detects inline images in the main body of a .docx document."""

    def detect_images(self, document) -> List[Figure]:
        figures: List[Figure] = []
        number = 1

        # Iterate direct children of <w:body> to preserve document order.
        # This covers paragraphs in the main flow (not inside tables).
        body = document.element.body
        for child in body:
            if child.tag == qn("w:p"):
                info = self._extract_image(child, document)
                if info:
                    figures.append(
                        Figure(
                            number=number,
                            image_data=info["data"],
                            image_format=info["format"],
                            paragraph_element=child,
                        )
                    )
                    number += 1

        return figures

    # ------------------------------------------------------------------
    def _extract_image(self, para_element, document) -> dict | None:
        """Return image bytes and format for the first image found in a paragraph."""
        drawings = para_element.findall(".//" + qn("w:drawing"))
        if not drawings:
            return None

        for drawing in drawings:
            for blip in drawing.findall(".//" + qn("a:blip")):
                r_id = blip.get(qn("r:embed"))
                if r_id and r_id in document.part.related_parts:
                    part = document.part.related_parts[r_id]
                    ct = part.content_type.lower()
                    fmt = (
                        "png"
                        if "png" in ct
                        else "jpeg"
                        if "jpeg" in ct or "jpg" in ct
                        else "gif"
                        if "gif" in ct
                        else "bmp"
                        if "bmp" in ct
                        else "tiff"
                        if "tiff" in ct
                        else "png"
                    )
                    return {"data": part.blob, "format": fmt}
        return None
