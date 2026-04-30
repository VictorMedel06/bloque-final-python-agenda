from __future__ import annotations

"""
Punto de entrada de la aplicación.
"""

import tkinter as tk
from pathlib import Path

from contact_app.db import connect, init_db
from contact_app.ui import ContactApp


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    db_path = base_dir / "contacts.sqlite3"
    conn = connect(db_path)
    init_db(conn)

    root = tk.Tk()

    def on_close() -> None:
        try:
            conn.close()
        finally:
            root.destroy()

    app = ContactApp(root, conn, on_close=on_close)
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
