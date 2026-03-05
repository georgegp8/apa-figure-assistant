from dataclasses import dataclass


@dataclass
class FigureStyle:
    font_name: str = "Times New Roman"
    font_size: int = 12


@dataclass
class Figure:
    number: int
    image_data: bytes
    image_format: str
    paragraph_element: object  # lxml <w:p> element reference
    title: str = ""
    note: str = ""
    has_note: bool = False
