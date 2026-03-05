from dataclasses import dataclass


@dataclass
class FigureStyle:
    font_name: str = "Times New Roman"
    font_size: int = 12
    # Word-standard line-spacing multiplier (1.0, 1.15, 1.5, 2.0, 2.5, 3.0)
    line_spacing: float = 1.0


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
