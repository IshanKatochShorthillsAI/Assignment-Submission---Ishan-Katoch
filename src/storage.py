# src/storage.py

import abc
import os
import csv
import json
import sqlite3


class Storage(abc.ABC):
    @abc.abstractmethod
    def store_data(self, data, data_type: str):
        """
        Store the provided data.
        :param data: The data to store.
        :param data_type: A string indicating the type of data (e.g., 'text', 'links', 'images', 'tables').
        """
        pass


class FileStorage(Storage):
    def __init__(self, output_dir: str):
        self.output_dir = output_dir
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def store_data(self, data, data_type: str):
        if data_type == "text":
            # Write full metadata as JSON.
            filename = os.path.join(self.output_dir, "extracted_text.json")
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
            print(f"Text data saved to {filename}")

        elif data_type == "links":
            filename = os.path.join(self.output_dir, "extracted_links.json")
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
            print(f"Link data saved to {filename}")

        elif data_type == "images":
            filename = os.path.join(self.output_dir, "extracted_images.json")
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
            print(f"Image data saved to {filename}")

        elif data_type == "tables":
            # For each table, create a CSV file that includes metadata in header rows.
            for idx, table in enumerate(data):
                filename = os.path.join(self.output_dir, f"table_{idx+1}.csv")
                with open(filename, "w", newline="", encoding="utf-8") as csvfile:
                    writer = csv.writer(csvfile)
                    # Write metadata header row.
                    writer.writerow(["Table Index", "Style", "Rows", "Columns"])
                    writer.writerow(
                        [
                            table.get("table_index", ""),
                            table.get("style", ""),
                            table.get("rows", ""),
                            table.get("columns", ""),
                        ]
                    )
                    # Write a blank row to separate metadata from table data.
                    writer.writerow([])
                    # Write the table cell data.
                    if "data" in table:
                        writer.writerows(table["data"])
                print(f"Table {idx+1} data saved to {filename}")
        else:
            raise ValueError("Unsupported data type for file storage.")


class SQLStorage(Storage):
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.conn = sqlite3.connect(self.db_path)
        self._create_tables()

    def _create_tables(self):
        cursor = self.conn.cursor()
        # Create table for text data
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS text_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                page_or_slide INTEGER,
                content TEXT
            )
            """
        )
        # Create table for links
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS links (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                page_or_slide INTEGER,
                text TEXT,
                url TEXT
            )
            """
        )
        # Create table for images metadata
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                page_or_slide INTEGER,
                name TEXT,
                format TEXT,
                width INTEGER,
                height INTEGER
            )
            """
        )
        # Create table for tables (store table data as JSON string)
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS tables (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                table_index INTEGER,
                data TEXT,
                rows INTEGER,
                columns INTEGER
            )
            """
        )
        self.conn.commit()

    def store_data(self, data, data_type: str):
        cursor = self.conn.cursor()
        if data_type == "text":
            for item in data:
                page_or_slide = item.get("page") or item.get("slide") or None
                content = json.dumps(item)
                cursor.execute(
                    "INSERT INTO text_data (page_or_slide, content) VALUES (?, ?)",
                    (page_or_slide, content),
                )
        elif data_type == "links":
            for link in data:
                page_or_slide = link.get("page") or link.get("slide") or None
                link_text = link.get("text", "") or link.get("contents", "")
                url = link.get("url") or link.get("uri")
                cursor.execute(
                    "INSERT INTO links (page_or_slide, text, url) VALUES (?, ?, ?)",
                    (page_or_slide, link_text, url),
                )
        elif data_type == "images":
            for image in data:
                page_or_slide = image.get("page") or image.get("slide") or None
                name = image.get("name") or image.get("image_id") or None
                fmt = image.get("format")
                width = image.get("width")
                height = image.get("height")
                cursor.execute(
                    "INSERT INTO images (page_or_slide, name, format, width, height) VALUES (?, ?, ?, ?, ?)",
                    (page_or_slide, name, fmt, width, height),
                )
        elif data_type == "tables":
            for table in data:
                table_index = table.get("table_index") or table.get("slide")
                table_data_json = json.dumps(table.get("data", []))
                rows = table.get("rows")
                columns = table.get("columns")
                cursor.execute(
                    "INSERT INTO tables (table_index, data, rows, columns) VALUES (?, ?, ?, ?)",
                    (table_index, table_data_json, rows, columns),
                )
        else:
            raise ValueError("Unsupported data type for SQL storage.")

        self.conn.commit()
