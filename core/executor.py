import sqlite3
import config


def execute(sql):
    conn = sqlite3.connect(f"file:{config.DB_PATH}?mode=ro", uri=True)  # read-only
    conn.execute("PRAGMA query_only = ON")
    try:
        cur = conn.execute(sql)
        cols = [d[0] for d in cur.description]
        rows = cur.fetchmany(config.MAX_ROWS)
        return cols, rows
    finally:
        conn.close()
