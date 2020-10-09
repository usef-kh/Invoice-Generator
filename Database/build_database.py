import sqlite3

def make_connection(db='Database.db'):
    conn = sqlite3.connect(db)
    return conn


def create_table(command):
    conn = make_connection()
    c = conn.cursor()

    c.execute(command)

    conn.commit()
    conn.close()


morespace_spaces = """
CREATE TABLE IF NOT EXISTS morespace_spaces (
code        text,
decription  text,
rate        real,
unit        text
)
"""

morespace_items = """
CREATE TABLE IF NOT EXISTS morespace_items (
item        text,
rate        real,
unit        text
)
"""

morespace_extras = """
CREATE TABLE IF NOT EXISTS morespace_extras (
item        text,
rate        real,
unit        text
)
"""

mediaspace_spaces = """
CREATE TABLE IF NOT EXISTS mediaspace_spaces (
space           text,
weekday_rate    real,
weekend_rate    real,
unit            text
)
"""

mediaspace_items = """
CREATE TABLE IF NOT EXISTS mediaspace_items (
item        text,
rate        real,
unit        text
)
"""

mediaspace_extras = """
CREATE TABLE IF NOT EXISTS mediaspace_extras (
item        text,
rate        real,
unit        text
)
"""

tables = [morespace_spaces,
          morespace_items,
          morespace_extras,
          mediaspace_spaces,
          mediaspace_items,
          mediaspace_extras]

for table in tables:
    create_table(table)