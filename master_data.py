import sqlite3
from typing import List, Dict, Optional

class MasterDataDatabase:
    """Database for managing master data (Names and Baustellen)."""

    def __init__(self, db_file="master_data.db"):
        self.db_file = db_file
        self.init_database()

    def init_database(self):
        """Create tables if they don't exist."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        # Names table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS names (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Baustellen table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS baustellen (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nummer TEXT NOT NULL,
                name TEXT NOT NULL,
                verpflegungsgeld REAL DEFAULT 0.0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(nummer, name)
            )
        ''')

        conn.commit()
        conn.close()

    # --- NAMES Methods ---
    def add_name(self, name: str) -> Optional[int]:
        """Add a new name. Returns ID or None if already exists."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute('INSERT INTO names (name) VALUES (?)', (name,))
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            # Name already exists
            return None
        finally:
            conn.close()

    def get_all_names(self) -> List[Dict]:
        """Get all names."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM names ORDER BY name ASC')
        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]

    def update_name(self, name_id: int, new_name: str) -> bool:
        """Update a name. Returns True if successful."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute('UPDATE names SET name = ? WHERE id = ?', (new_name, name_id))
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def delete_name(self, name_id: int) -> bool:
        """Delete a name. Returns True if successful."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute('DELETE FROM names WHERE id = ?', (name_id,))
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error:
            return False
        finally:
            conn.close()

    # --- BAUSTELLEN Methods ---
    def add_baustelle(self, nummer: str, name: str, verpflegungsgeld: float = 0.0) -> Optional[int]:
        """Add a new baustelle. Returns ID or None if already exists."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute(
                'INSERT INTO baustellen (nummer, name, verpflegungsgeld) VALUES (?, ?, ?)',
                (nummer, name, verpflegungsgeld)
            )
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()

    def get_all_baustellen(self) -> List[Dict]:
        """Get all baustellen."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM baustellen ORDER BY nummer ASC, name ASC')
        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]

    def update_baustelle(self, baustelle_id: int, nummer: str, name: str, verpflegungsgeld: float) -> bool:
        """Update a baustelle. Returns True if successful."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute(
                'UPDATE baustellen SET nummer = ?, name = ?, verpflegungsgeld = ? WHERE id = ?',
                (nummer, name, verpflegungsgeld, baustelle_id)
            )
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def delete_baustelle(self, baustelle_id: int) -> bool:
        """Delete a baustelle. Returns True if successful."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute('DELETE FROM baustellen WHERE id = ?', (baustelle_id,))
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error:
            return False
        finally:
            conn.close()
