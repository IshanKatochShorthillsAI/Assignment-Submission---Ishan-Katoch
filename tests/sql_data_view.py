import sqlite3

conn = sqlite3.connect("extracted_data.db")
cursor = conn.cursor()
cursor.execute("SELECT * FROM text_data")
rows = cursor.fetchall()
for row in rows:
    print(row)
conn.close()
