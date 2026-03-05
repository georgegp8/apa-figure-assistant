import json
from pathlib import Path

from models.figure import FigureStyle

_CONFIG_FILE = Path.home() / ".apa_figure_config.json"


class StyleManager:
    @staticmethod
    def load() -> FigureStyle:
        try:
            if _CONFIG_FILE.exists():
                with open(_CONFIG_FILE, encoding="utf-8") as f:
                    data = json.load(f)
                return FigureStyle(
                    font_name=data.get("font_name", "Times New Roman"),
                    font_size=int(data.get("font_size", 12)),
                    space_after_pt=float(data.get("space_after_pt", 0.0)),
                )
        except Exception:
            pass
        return FigureStyle()

    @staticmethod
    def save(style: FigureStyle):
        try:
            with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "font_name": style.font_name,
                        "font_size": style.font_size,
                        "space_after_pt": style.space_after_pt,
                    },
                    f,
                    indent=2,
                )
        except Exception:
            pass
