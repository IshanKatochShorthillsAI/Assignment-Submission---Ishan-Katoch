from fileloader import PDFLoader, DOCXLoader, PPTLoader
from data_extractor import DataExtractor
from storage import FileStorage, SQLStorage


def process_file(filepath: str, storage_method: str = "file"):
    if filepath.lower().endswith(".pdf"):
        loader = PDFLoader()
    elif filepath.lower().endswith(".docx"):
        loader = DOCXLoader()
    elif filepath.lower().endswith((".ppt", ".pptx")):
        loader = PPTLoader()
    else:
        raise ValueError("Unsupported file type.")

    extractor = DataExtractor(loader)

    # Extract data
    text_data = extractor.extract_text(filepath)
    links = extractor.extract_links(filepath)
    images = extractor.extract_images(filepath)
    tables = extractor.extract_tables(filepath)

    # Choose storage method
    if storage_method == "file":
        storage = FileStorage(output_dir="output")
    elif storage_method == "sql":
        storage = SQLStorage(db_path="extracted_data.db")
    else:
        raise ValueError("Unsupported storage method.")

    # Store the extracted data
    storage.store_data(text_data, "text")
    storage.store_data(links, "links")
    storage.store_data(images, "images")
    storage.store_data(tables, "tables")


if __name__ == "__main__":
    process_file("Test Slides.pptx", storage_method="file")
