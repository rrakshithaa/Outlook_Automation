from io import BytesIO
import pandas as pd
from PyPDF2 import PdfReader
from docx import Document

class FileExtractor:
    @staticmethod
    def extract_details(file_content, file_type):
        if file_type == "pdf":
            return FileExtractor._extract_from_pdf(file_content)
        elif file_type == "docx":
            return FileExtractor._extract_from_docx(file_content)
        elif file_type == "xlsx":
            return FileExtractor._extract_from_excel(file_content)
        return {}

    @staticmethod
    def _extract_from_pdf(file_content):
        try:
            reader = PdfReader(BytesIO(file_content))
            text = "".join(page.extract_text() for page in reader.pages)
            return text
        except Exception as e:
            print(f"Error extracting PDF: {e}")
            return {}
