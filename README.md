# Modular Document Extraction and Storage

## Overview

This project extracts various types of content and their associated metadata from PDF, DOCX, and PPT/PPTX files. It is structured into three main components:

1. **File Loading:**  
   Abstracts the file reading logic for each format.
2. **Data Extraction:**  
   Extracts text (with font, size, and formatting metadata), hyperlinks, images, and tables from the documents.
3. **Data Storage:**  
   Saves the extracted data either as files (JSON for text, links, images and CSV for tables) or into an SQLite database.

## Architecture and Design

### File Loading

- **Abstract Class: `FileLoader`**  
  Provides the blueprint for:
  - `validate_file(filepath: str) -> bool`
  - `load_file(filepath: str)`

- **Concrete Classes:**
  - **PDFLoader:** Uses PyPDF2 (and optionally PyMuPDF for rich metadata) to load PDF files.
  - **DOCXLoader:** Uses python-docx to load DOCX files.
  - **PPTLoader:** Uses python-pptx to load PPT/PPTX files.

### Data Extraction

- **DataExtractor Class**  
  Takes a `FileLoader` instance as input and provides methods:
  - `extract_text(filepath: str)`  
    Extracts text with metadata (e.g., page/slide number, rotation, dimensions, font details) using PyMuPDF (for PDFs) or the appropriate library for DOCX/PPT.
  - `extract_links(filepath: str)`  
    Extracts hyperlinks along with metadata (such as their location/rectangle and type).
  - `extract_images(filepath: str)`  
    Extracts image data and encodes it as Base64. For PPTX files, images are recursively extracted from shapes and, as a fallback, via relationships.
  - `extract_tables(filepath: str)`  
    Extracts tables and their metadata (dimensions, styles, and cell content) from DOCX, PPT, and PDF files.

### Data Storage

- **Abstract Class: `Storage`**  
  Defines the method:
  - `store_data(data, data_type: str)`

- **Concrete Implementations:**
  - **FileStorage:**  
    Saves extracted text, links, and images as JSON files and tables as CSV files.
  - **SQLStorage:**  
    Inserts extracted data into an SQLite database (with separate tables for text, links, images, and tables).

## Usage

The main entry point is `src/main.py`. To process a document, specify the file and storage method:
```bash
python src/main.py
```
By default, it uses file-based storage (outputting JSON and CSV files). To switch to SQL storage, change the parameter in the `process_file` function.

## Testing

A comprehensive test suite is provided in the `tests` folder. The tests cover extraction and storage for all file types and edge cases. To run the tests:
```bash
python -m unittest -v tests/test_suite.py
```
This will run unit tests for:
- PDF extraction (text, links, images)
- DOCX extraction
- PPT/PPTX extraction (text, links, tables, images)
- Storage (both file and SQL)

## Viewing SQL Data

If you choose SQL storage, the data is stored in an SQLite database (e.g., `extracted_data.db`). You can view the database using:
- **Command line:**
  ```bash
  sqlite3 extracted_data.db
  ```
  Then execute SQL queries like:
  ```sql
  SELECT * FROM text_data;
  ```

## Output Screenshots

DOCX Output (extracted text)
- ![image](https://github.com/user-attachments/assets/e32be2d0-2a8a-498f-8c8f-de1e7520a257)

DOCX Output (extracted images)
- ![image](https://github.com/user-attachments/assets/5c2d3393-ec7e-4ec5-b5dd-5b722316b3dc)

PDF Output (extracted text)
- ![image](https://github.com/user-attachments/assets/572658b4-a57c-4d8b-ae1d-02eee40bbbe3)

PDF Output (extracted table)
- ![image](https://github.com/user-attachments/assets/9da2285f-e042-4f1e-8266-259544de5713)

PPTX Output (extracted text)
- ![image](https://github.com/user-attachments/assets/2480252e-50a3-4490-aa62-bc0823563155)

PPTX Output (extracted links)
- ![image](https://github.com/user-attachments/assets/e3a6670e-1219-438b-84d0-c58eaaa973cf)

## Learnings and Reflections

This project provided the opportunity to learn and integrate several concepts and libraries:
- **Modular Design:**  
  Creating abstract classes for file loading and storage allowed a flexible and extendable architecture.
- **Rich Metadata Extraction:**  
  Extracting detailed metadata (font styles, dimensions, etc.) from PDFs using PyMuPDF improved the granularity of the data.
- **Error Handling and Fallbacks:**  
  Implementing fallback mechanisms (using pdfplumber when PyMuPDF fails) ensured broader compatibility across file formats.
- **Testing:**  
  Writing unit tests using fake loader classes helped validate functionality without relying on external files.
- **Version Control:**  
  Meaningful commit messages and incremental commits facilitated tracking changes and understanding the evolution of the project.
