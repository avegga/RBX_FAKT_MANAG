import sqlite3

# Initialize the database
connection = sqlite3.connect('rbx_fakt_manag.db')
cursor = connection.cursor()

# Read schema from file
with open('schema.sql', 'r', encoding='utf-8') as schema_file:
    schema_lines = schema_file.readlines()

# Filter out comments
schema = ''.join(line for line in schema_lines if not line.strip().startswith('#'))

# Execute schema
cursor.executescript(schema)

# Commit and close
connection.commit()
connection.close()