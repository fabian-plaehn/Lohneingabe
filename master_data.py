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

        # SKUG settings table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS skug_settings (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                winter_monday REAL DEFAULT 8.0,
                winter_tuesday REAL DEFAULT 8.0,
                winter_wednesday REAL DEFAULT 8.0,
                winter_thursday REAL DEFAULT 8.0,
                winter_friday REAL DEFAULT 6.0,
                summer_monday REAL DEFAULT 8.5,
                summer_tuesday REAL DEFAULT 8.5,
                summer_wednesday REAL DEFAULT 8.5,
                summer_thursday REAL DEFAULT 8.5,
                summer_friday REAL DEFAULT 7.0,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Insert default SKUG settings if not exists
        cursor.execute('SELECT COUNT(*) FROM skug_settings WHERE id = 1')
        if cursor.fetchone()[0] == 0:
            cursor.execute('''
                INSERT INTO skug_settings (id, winter_monday, winter_tuesday, winter_wednesday,
                                          winter_thursday, winter_friday, summer_monday,
                                          summer_tuesday, summer_wednesday, summer_thursday, summer_friday)
                VALUES (1, 8.0, 8.0, 8.0, 8.0, 6.0, 8.5, 8.5, 8.5, 8.5, 7.0)
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

    # --- SKUG SETTINGS Methods ---
    def get_skug_settings(self) -> Dict:
        """Get SKUG settings."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM skug_settings WHERE id = 1')
        row = cursor.fetchone()
        conn.close()

        return dict(row) if row else {}

    def update_skug_settings(self, settings: Dict) -> bool:
        """Update SKUG settings. Returns True if successful."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute('''
                UPDATE skug_settings SET
                    winter_monday = ?,
                    winter_tuesday = ?,
                    winter_wednesday = ?,
                    winter_thursday = ?,
                    winter_friday = ?,
                    summer_monday = ?,
                    summer_tuesday = ?,
                    summer_wednesday = ?,
                    summer_thursday = ?,
                    summer_friday = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = 1
            ''', (
                settings.get('winter_monday', 8.0),
                settings.get('winter_tuesday', 8.0),
                settings.get('winter_wednesday', 8.0),
                settings.get('winter_thursday', 8.0),
                settings.get('winter_friday', 6.0),
                settings.get('summer_monday', 8.5),
                settings.get('summer_tuesday', 8.5),
                settings.get('summer_wednesday', 8.5),
                settings.get('summer_thursday', 8.5),
                settings.get('summer_friday', 7.0)
            ))
            conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error:
            return False
        finally:
            conn.close()
