# src/fileloader.py

import abc


class FileLoader(abc.ABC):
    @abc.abstractmethod
    def validate_file(self, filepath: str) -> bool:
        """Check if the file is valid for this loader."""
        pass

    @abc.abstractmethod
    def load_file(self, filepath: str):
        """Load the file and return its content or an object representation."""
        pass


# PDF Loader using PyMuPDF (fitz)
import os
import fitz  # PyMuPDF


class PDFLoader(FileLoader):
    def validate_file(self, filepath: str) -> bool:
        return filepath.lower().endswith(".pdf")

    def load_file(self, filepath: str):
        if not self.validate_file(filepath):
            raise ValueError("Invalid file type for PDFLoader")
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        # Open the PDF using PyMuPDF to capture rich metadata.
        document = fitz.open(filepath)
        return document


# DOCXLoader
from docx import Document


class DOCXLoader(FileLoader):
    def validate_file(self, filepath: str) -> bool:
        return filepath.lower().endswith(".docx")

    def load_file(self, filepath: str):
        if not self.validate_file(filepath):
            raise ValueError("Invalid file type for DOCXLoader")
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        document = Document(filepath)
        return document


# PPTLoader
from pptx import Presentation


class PPTLoader(FileLoader):
    def validate_file(self, filepath: str) -> bool:
        return filepath.lower().endswith((".ppt", ".pptx"))

    def load_file(self, filepath: str):
        if not self.validate_file(filepath):
            raise ValueError("Invalid file type for PPTLoader")
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        presentation = Presentation(filepath)
        return presentation
