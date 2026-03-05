from pathlib import Path

from docx import Document


class DocumentManager:
    def __init__(self):
        self.document = None
        self.file_path = None

    def open_document(self, file_path: str):
        self.file_path = Path(file_path)
        self.document = Document(file_path)
        return self.document

    def save_document(self, output_path: str = None) -> str:
        if output_path is None:
            stem = self.file_path.stem
            output_path = str(
                self.file_path.parent / f"{stem}_APA{self.file_path.suffix}"
            )
        self.document.save(output_path)
        return output_path
