import sqlite3
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent


def main():
    connection = sqlite3.connect(BASE_DIR / "rbx_fakt_manag.db")
    cursor = connection.cursor()

    with open(BASE_DIR / "schema.sql", "r", encoding="utf-8") as schema_file:
        schema = schema_file.read()

    cursor.executescript(schema)
    connection.commit()
    connection.close()
    print("Database schema initialized.")


if __name__ == "__main__":
    main()