import sqlite3
from typing import List, Dict, Optional

class MasterDataDatabase:
    """Database for managing master data (Names and Baustellen)."""

    SCHEMA_VERSION = 2  # Current database schema version

    def __init__(self, db_file="master_data.db"):
        self.db_file = db_file
        self.init_database()

    def init_database(self):
        """Create tables if they don't exist."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        # Schema version table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS schema_version (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                version INTEGER NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Check and set schema version
        cursor.execute('SELECT version FROM schema_version WHERE id = 1')
        row = cursor.fetchone()
        current_version = 0
        if row:
            current_version = row[0]
        else:
            cursor.execute('INSERT INTO schema_version (id, version) VALUES (1, ?)', (self.SCHEMA_VERSION,))
            current_version = self.SCHEMA_VERSION

        # Names table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS names (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                worker_type TEXT DEFAULT 'Fest',
                kein_verpflegungsgeld INTEGER DEFAULT 0,
                keine_feiertagssstunden INTEGER DEFAULT 0,
                weekly_hours REAL DEFAULT 0.0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Migrations
        if current_version < 2:
            try:
                cursor.execute('ALTER TABLE names ADD COLUMN kein_verpflegungsgeld INTEGER DEFAULT 0')
            except sqlite3.OperationalError:
                pass # Column might already exist if partial migration happened
            try:
                cursor.execute('ALTER TABLE names ADD COLUMN keine_feiertagssstunden INTEGER DEFAULT 0')
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute('ALTER TABLE names ADD COLUMN weekly_hours REAL DEFAULT 0.0')
            except sqlite3.OperationalError:
                pass
            
            cursor.execute('UPDATE schema_version SET version = 2 WHERE id = 1')

        # Baustellen table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS baustellen (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nummer TEXT NOT NULL,
                name TEXT NOT NULL,
                verpflegungsgeld REAL DEFAULT 0.0,
                fahrzeit REAL DEFAULT 0.0,
                distance_km REAL DEFAULT 0.0,
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
    def add_name(self, name: str, worker_type: str = 'Fest', kein_verpflegungsgeld: bool = False, 
                 keine_feiertagssstunden: bool = False, weekly_hours: float = 0.0) -> Optional[int]:
        """Add a new name. Returns ID or None if already exists."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute('''
                INSERT INTO names (name, worker_type, kein_verpflegungsgeld, keine_feiertagssstunden, weekly_hours) 
                VALUES (?, ?, ?, ?, ?)
            ''', (name, worker_type, 1 if kein_verpflegungsgeld else 0, 
                  1 if keine_feiertagssstunden else 0, weekly_hours))
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
    
    def get_all_names_list(self) -> List[str]:
        """Get all names as a list of strings."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        cursor.execute('SELECT name FROM names ORDER BY name ASC')
        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows]
    
    def get_worker_type_by_name(self, name: str) -> Optional[str]:
        """Get the worker_type of a name."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT worker_type FROM names WHERE name = ?', (name,))
        row = cursor.fetchone()
        conn.close()

        return row[0] if row else None


    def update_name(self, name_id: int, new_name: str, worker_type: str = None, 
                    kein_verpflegungsgeld: bool = None, keine_feiertagssstunden: bool = None, 
                    weekly_hours: float = None) -> bool:
        """Update a name. Returns True if successful."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            # Build query dynamically based on provided arguments
            updates = ['name = ?']
            params = [new_name]
            
            if worker_type is not None:
                updates.append('worker_type = ?')
                params.append(worker_type)
            
            if kein_verpflegungsgeld is not None:
                updates.append('kein_verpflegungsgeld = ?')
                params.append(1 if kein_verpflegungsgeld else 0)
                
            if keine_feiertagssstunden is not None:
                updates.append('keine_feiertagssstunden = ?')
                params.append(1 if keine_feiertagssstunden else 0)
                
            if weekly_hours is not None:
                updates.append('weekly_hours = ?')
                params.append(weekly_hours)
                
            params.append(name_id)
            
            query = f'UPDATE names SET {", ".join(updates)} WHERE id = ?'
            cursor.execute(query, params)
            
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
    def add_baustelle(self, nummer: str, name: str, verpflegungsgeld: float = 0.0, fahrzeit: float = 0.0, distance_km: float = 0.0) -> Optional[int]:
        """Add a new baustelle. Returns ID or None if already exists."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute(
                'INSERT INTO baustellen (nummer, name, verpflegungsgeld, fahrzeit, distance_km) VALUES (?, ?, ?, ?, ?)',
                (nummer, name, verpflegungsgeld, fahrzeit, distance_km)
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
    
    def get_baustelle_by_nummer(self, baustelle_id: int) -> Optional[Dict]:
        """Get a baustelle by nummer."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM baustellen WHERE nummer = ?', (baustelle_id,))
        row = cursor.fetchone()
        conn.close()

        return dict(row) if row else None

    def update_baustelle(self, baustelle_id: int, nummer: str, name: str, verpflegungsgeld: float, fahrzeit: float = 0.0, distance_km: float = 0.0) -> bool:
        """Update a baustelle. Returns True if successful."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute(
                'UPDATE baustellen SET nummer = ?, name = ?, verpflegungsgeld = ?, fahrzeit = ?, distance_km = ? WHERE id = ?',
                (nummer, name, verpflegungsgeld, fahrzeit, distance_km, baustelle_id)
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
