from dataclasses import dataclass


@dataclass
class FigureStyle:
    font_name: str = "Times New Roman"
    font_size: int = 12
    # Points of space inserted after the label and title paragraphs (0 = Word default)
    space_after_pt: float = 0.0


@dataclass
class Figure:
    number: int
    image_data: bytes
    image_format: str
    paragraph_element: object  # lxml <w:p> element reference
    title: str = ""
    note: str = ""
    has_note: bool = False
    # Custom image dimensions in cm (0.0 = keep original)
    image_width_cm: float = 0.0
    image_height_cm: float = 0.0
