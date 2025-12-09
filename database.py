import sqlite3
from typing import List, Dict, Optional
import pandas as pd

class Database:
    SCHEMA_VERSION = 4

    def __init__(self, db_file="stundenliste.db"):
        self.db_file = db_file
        self.init_database()
    
    def init_database(self):
        """Create table if it doesn't exist."""
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
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS stunden_eintraege (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                jahr INTEGER NOT NULL,
                monat INTEGER NOT NULL,
                tag INTEGER NOT NULL,
                name TEXT NOT NULL,
                wochentag TEXT,
                stunden REAL NOT NULL,
                urlaub TEXT,
                krank TEXT,
                unter_8h BOOLEAN,
                skug TEXT,
                baustelle TEXT,
                travel_status TEXT,
                fruehstueck BOOLEAN,
                mittag BOOLEAN,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(jahr, monat, tag, name)
            )
        ''')
        if current_version < 2:
            try:
                cursor.execute('ALTER TABLE stunden_eintraege RENAME COLUMN unter_8h TO kg_8h')
            except sqlite3.OperationalError:
                pass  # Column already renamed
            cursor.execute('UPDATE schema_version SET version = 2 WHERE id = 1')
            current_version = 2
        
        if current_version < 3:
            try:
                cursor.execute('ALTER TABLE stunden_eintraege ADD COLUMN fruehstueck BOOLEAN')
                cursor.execute('ALTER TABLE stunden_eintraege ADD COLUMN mittag BOOLEAN')
            except sqlite3.OperationalError:
                pass # Columns likely already exist
            cursor.execute('UPDATE schema_version SET version = 3 WHERE id = 1')
            current_version = 3

        if current_version < 4:
            # Remove UNIQUE constraint by recreating table
            cursor.execute('CREATE TABLE stunden_eintraege_new ('
                           'id INTEGER PRIMARY KEY AUTOINCREMENT,'
                           'jahr INTEGER NOT NULL,'
                           'monat INTEGER NOT NULL,'
                           'tag INTEGER NOT NULL,'
                           'name TEXT NOT NULL,'
                           'wochentag TEXT,'
                           'stunden REAL NOT NULL,'
                           'urlaub TEXT,'
                           'krank TEXT,'
                           'kg_8h BOOLEAN,'
                           'skug TEXT,'
                           'baustelle TEXT,'
                           'travel_status TEXT,'
                           'fruehstueck BOOLEAN,'
                           'mittag BOOLEAN,'
                           'created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,'
                           'updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP'
                           ')')
            
            # Copy data
            cursor.execute('INSERT INTO stunden_eintraege_new (id, jahr, monat, tag, name, wochentag, stunden, urlaub, krank, kg_8h, skug, baustelle, travel_status, fruehstueck, mittag, created_at, updated_at) '
                           'SELECT id, jahr, monat, tag, name, wochentag, stunden, urlaub, krank, kg_8h, skug, baustelle, travel_status, fruehstueck, mittag, created_at, updated_at FROM stunden_eintraege')
            
            cursor.execute('DROP TABLE stunden_eintraege')
            cursor.execute('ALTER TABLE stunden_eintraege_new RENAME TO stunden_eintraege')
            
            cursor.execute('UPDATE schema_version SET version = 4 WHERE id = 1')
            current_version = 4

        conn.commit()
        conn.close()
    
    def add_or_update_entry(self, data: Dict) -> tuple[int, bool]:
        """
        Add a new entry or update if the combination of jahr, monat, tag, name exists.
        Returns (row_id, was_updated) where was_updated is True if existing row was updated.
        """
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            # Check if entry exists
            cursor.execute('''
                SELECT id FROM stunden_eintraege 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            ''', (
                data.get('Jahr'),
                data.get('Monat'),
                data.get('Tag'),
                data.get('Name')
            ))
            
            existing = cursor.fetchone()
            
            if existing:
                # Update existing entry - Dynamic Update
                entry_id = existing[0]
                
                updates = []
                params = []
                
                # List of possible fields to update
                fields_map = {
                    'Wochentag': 'wochentag',
                    'Stunden': 'stunden',
                    'Urlaub': 'urlaub',
                    'Krank': 'krank',
                    'kg_8h': 'kg_8h',
                    'SKUG': 'skug',
                    'Baustelle': 'baustelle',
                    'travel_status': 'travel_status',
                    'fruehstueck': 'fruehstueck',
                    'mittag': 'mittag'
                }
                
                for data_key, db_col in fields_map.items():
                    if data_key in data:
                        updates.append(f"{db_col}=?")
                        params.append(data[data_key])
                
                updates.append("updated_at=CURRENT_TIMESTAMP")
                
                if updates:
                    query = f"UPDATE stunden_eintraege SET {', '.join(updates)} WHERE id = ?"
                    params.append(entry_id)
                    cursor.execute(query, params)
                    conn.commit()
                
                return (entry_id, True)
            else:
                # Insert new entry
                print("Inserting new entry:", data)
                
                # Default values for required fields if missing
                if 'Stunden' not in data:
                    data['Stunden'] = 0.0
                    
                cursor.execute('''
                    INSERT INTO stunden_eintraege
                    (jahr, monat, tag, name, wochentag, stunden, urlaub, krank, kg_8h, skug, baustelle, travel_status, fruehstueck, mittag)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    data.get('Jahr'),
                    data.get('Monat'),
                    data.get('Tag'),
                    data.get('Name'),
                    data.get('Wochentag'),
                    data.get('Stunden'),
                    data.get('Urlaub'),
                    data.get('Krank'),
                    data.get('kg_8h'),
                    data.get('SKUG'),
                    data.get('Baustelle'),
                    data.get('travel_status'),
                    data.get('fruehstueck'),
                    data.get('mittag')
                ))
                
                conn.commit()
                row_id = cursor.lastrowid
                return (row_id, False)
        
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            conn.rollback()
            raise
        
        finally:
            conn.close()

    def update_entry(self, entry_id: int, data: Dict) -> bool:
        """Update an existing entry by ID."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            updates = []
            params = []
            
            # List of possible fields to update
            fields_map = {
                'Wochentag': 'wochentag',
                'Stunden': 'stunden',
                'Urlaub': 'urlaub',
                'Krank': 'krank',
                'kg_8h': 'kg_8h',
                'SKUG': 'skug',
                'Baustelle': 'baustelle',
                'travel_status': 'travel_status',
                'fruehstueck': 'fruehstueck',
                'mittag': 'mittag'
            }
            
            for data_key, db_col in fields_map.items():
                if data_key in data:
                    updates.append(f"{db_col}=?")
                    params.append(data[data_key])
            
            updates.append("updated_at=CURRENT_TIMESTAMP")
            
            if updates:
                query = f"UPDATE stunden_eintraege SET {', '.join(updates)} WHERE id = ?"
                params.append(entry_id)
                cursor.execute(query, params)
                conn.commit()
                return True
            return False
        
        except sqlite3.Error as e:
            print(f"Database error updating entry: {e}")
            conn.rollback()
            return False
        finally:
            conn.close()

    def insert_entry(self, data: Dict) -> int:
        """Insert a new entry."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            # Default values for required fields if missing
            if 'Stunden' not in data:
                data['Stunden'] = 0.0
                
            cursor.execute('''
                INSERT INTO stunden_eintraege
                (jahr, monat, tag, name, wochentag, stunden, urlaub, krank, kg_8h, skug, baustelle, travel_status, fruehstueck, mittag)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('Jahr'),
                data.get('Monat'),
                data.get('Tag'),
                data.get('Name'),
                data.get('Wochentag'),
                data.get('Stunden'),
                data.get('Urlaub'),
                data.get('Krank'),
                data.get('kg_8h'),
                data.get('SKUG'),
                data.get('Baustelle'),
                data.get('travel_status'),
                data.get('fruehstueck'),
                data.get('mittag')
            ))
            
            conn.commit()
            return cursor.lastrowid
        
        except sqlite3.Error as e:
            print(f"Database error inserting entry: {e}")
            conn.rollback()
            raise
        finally:
            conn.close()
    
    def get_all_entries(self) -> List[Dict]:
        """Retrieve all entries from the database."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM stunden_eintraege 
            ORDER BY jahr DESC, monat DESC, tag DESC
        ''')
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_entries_by_month_and_name(self, year: int, month: int, name: str) -> List[Dict]:
        """Get all entries for a specific person in a specific month."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM stunden_eintraege 
            WHERE jahr = ? AND monat = ? AND name = ?
            ORDER BY tag ASC
        ''', (year, month, name))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_entry(self, year: int, month: int, day: int, name: str) -> Optional[Dict]:
        """Get a single entry for a specific person on a specific date. (First found)"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM stunden_eintraege 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        ''', (year, month, day, name))
        
        row = cursor.fetchone()
        conn.close()
        
        return dict(row) if row else None
    
    def clear_entries_for_day(self, year: int, month: int, day: int, name: str) -> int:
        existing_entries = self.get_entries_for_day(year, month, day, name)
        for entry in existing_entries:
            self.delete_entry(entry['id'])
        return len(existing_entries)

    def get_entries_for_day(self, year: int, month: int, day: int, name: str) -> List[Dict]:
        """Get all entries for a specific person on a specific date."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM stunden_eintraege 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        ''', (year, month, day, name))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_entries_by_date_and_baustelle(self, year: int, month: int, day: int, baustelle: str) -> List[Dict]:
        """Get all entries for a specific construction site on a specific date."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Try exact match first, then partial match (for "Nummer - Name" format)
        cursor.execute('''
            SELECT * FROM stunden_eintraege
            WHERE jahr = ? AND monat = ? AND tag = ? AND (baustelle = ? OR baustelle LIKE ?)
            ORDER BY name ASC
        ''', (year, month, day, baustelle, f"{baustelle}%"))

        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]
    
    def get_entries_by_date(self, year: int, month: int) -> List[Dict]:
        """Get entries for a specific month."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM stunden_eintraege 
            WHERE jahr = ? AND monat = ?
            ORDER BY tag ASC
        ''', (year, month))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def delete_entry(self, entry_id: int) -> bool:
        """Delete an entry by ID."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute('DELETE FROM stunden_eintraege WHERE id = ?', (entry_id,))
            conn.commit()
            success = cursor.rowcount > 0
            return success
        
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            conn.rollback()
            return False
        
        finally:
            conn.close()
    
    def export_to_excel(self, filename="stundenliste.xlsx"):
        """Export all data to Excel."""
        entries = self.get_all_entries()

        if not entries:
            print("No data to export")
            return False

        df = pd.DataFrame(entries)
        # Reorder columns for better readability
        columns_order = ['id', 'jahr', 'monat', 'tag', 'name', 'wochentag',
                        'stunden', 'checkbox1', 'checkbox2', 'skug', 'baustelle',
                        'created_at', 'updated_at']
        df = df[columns_order]

        df.to_excel(filename, index=False)
        return True

