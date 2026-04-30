from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

from contact_app.models import Contact


def connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def add_contact(conn: sqlite3.Connection, contact: Contact) -> int:
    now = datetime.utcnow().isoformat(timespec="seconds")
    cur = conn.execute(
        "INSERT INTO contacts (name, phone, email, created_at) VALUES (?, ?, ?, ?)",
        (contact.name, contact.phone, contact.email, now),
    )
    conn.commit()
    return int(cur.lastrowid)


def update_contact(conn: sqlite3.Connection, contact_id: int, contact: Contact) -> None:
    conn.execute(
        "UPDATE contacts SET name = ?, phone = ?, email = ? WHERE id = ?",
        (contact.name, contact.phone, contact.email, contact_id),
    )
    conn.commit()


def delete_contact(conn: sqlite3.Connection, contact_id: int) -> None:
    conn.execute("DELETE FROM contacts WHERE id = ?", (contact_id,))
    conn.commit()


def get_contact(conn: sqlite3.Connection, contact_id: int) -> Optional[Contact]:
    row = conn.execute("SELECT * FROM contacts WHERE id = ?", (contact_id,)).fetchone()
    if row is None:
        return None
    return _row_to_contact(row)


def list_contacts(conn: sqlite3.Connection, search: Optional[str] = None) -> list[Contact]:
    if search:
        like = f"%{search.strip()}%"
        rows = conn.execute(
            """
            SELECT * FROM contacts
            WHERE name LIKE ? OR phone LIKE ? OR email LIKE ?
            ORDER BY name COLLATE NOCASE ASC
            """,
            (like, like, like),
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM contacts ORDER BY name COLLATE NOCASE ASC"
        ).fetchall()
    return [_row_to_contact(r) for r in rows]


def count_contacts(conn: sqlite3.Connection, search: Optional[str] = None) -> int:
    if search:
        like = f"%{search.strip()}%"
        row = conn.execute(
            """
            SELECT COUNT(*) AS total
            FROM contacts
            WHERE name LIKE ? OR phone LIKE ? OR email LIKE ?
            """,
            (like, like, like),
        ).fetchone()
    else:
        row = conn.execute("SELECT COUNT(*) AS total FROM contacts").fetchone()
    if row is None:
        return 0
    return int(row["total"])


def _row_to_contact(row: sqlite3.Row) -> Contact:
    created_at = None
    if row["created_at"]:
        try:
            created_at = datetime.fromisoformat(row["created_at"])
        except ValueError:
            created_at = None
    return Contact(
        id=int(row["id"]),
        name=str(row["name"]),
        phone=str(row["phone"]),
        email=str(row["email"]),
        created_at=created_at,
    )
