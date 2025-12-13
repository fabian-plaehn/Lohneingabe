import sqlite3
from typing import List, Dict, Optional
import pandas as pd

class Database:
    SCHEMA_VERSION = 5

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

        if current_version < 5:
            import shutil
            import os
            from datetime import datetime as dt
            
            # Create backup folder with timestamp
            backup_folder = os.path.join(os.path.dirname(self.db_file), f"Backup_{current_version}")
            os.makedirs(backup_folder, exist_ok=True)
            
            timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_folder, f"stundenliste_backup_{timestamp}.db")
            
            # Close connection before backup
            conn.close()
            
            # Backup the database
            shutil.copy2(self.db_file, backup_file)
            print(f"Database backed up to: {backup_file}")
            
            # Reconnect
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # Create new tages_metadaten table (unique per day/worker)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS tages_metadaten (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    jahr INTEGER NOT NULL,
                    monat INTEGER NOT NULL,
                    tag INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    wochentag TEXT,
                    skug TEXT,
                    kg_8h BOOLEAN,
                    travel_status TEXT,
                    fruehstueck BOOLEAN,
                    mittag BOOLEAN,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(jahr, monat, tag, name)
                )
            ''')
            
            # Create new arbeitsstunden table (multiple entries per day/worker allowed)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS arbeitsstunden (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    jahr INTEGER NOT NULL,
                    monat INTEGER NOT NULL,
                    tag INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    wochentag TEXT,
                    kostenstelle TEXT,
                    stunden REAL NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Migrate data from old table to new tables
            cursor.execute('SELECT * FROM stunden_eintraege')
            old_entries = cursor.fetchall()
            
            for entry in old_entries:
                # entry structure: id, jahr, monat, tag, name, wochentag, stunden, urlaub, krank, 
                #                  kg_8h, skug, baustelle, travel_status, fruehstueck, mittag, created_at, updated_at
                entry_id, jahr, monat, tag, name, wochentag, stunden, urlaub, krank, kg_8h, skug, baustelle, travel_status, fruehstueck, mittag, created_at, updated_at = entry
                
                # Insert into tages_metadaten (one entry per day/worker)
                try:
                    cursor.execute('''
                        INSERT OR IGNORE INTO tages_metadaten 
                        (jahr, monat, tag, name, wochentag, skug, kg_8h, travel_status, fruehstueck, mittag, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (jahr, monat, tag, name, wochentag, skug, kg_8h, travel_status, fruehstueck, mittag, created_at, updated_at))
                except sqlite3.IntegrityError:
                    # Entry already exists, skip
                    pass
                
                # Insert into arbeitsstunden
                # Handle Urlaub/Krank as kostenstelle entries
                if urlaub:
                    cursor.execute('''
                        INSERT INTO arbeitsstunden 
                        (jahr, monat, tag, name, wochentag, kostenstelle, stunden, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (jahr, monat, tag, name, wochentag, 'Urlaub', float(urlaub) if urlaub else 0.0, created_at, updated_at))
                
                if krank:
                    cursor.execute('''
                        INSERT INTO arbeitsstunden 
                        (jahr, monat, tag, name, wochentag, kostenstelle, stunden, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (jahr, monat, tag, name, wochentag, 'Krank', float(krank) if krank else 0.0, created_at, updated_at))
                
                # Insert regular work hours with baustelle as kostenstelle
                if stunden and stunden > 0 and baustelle:
                    cursor.execute('''
                        INSERT INTO arbeitsstunden 
                        (jahr, monat, tag, name, wochentag, kostenstelle, stunden, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (jahr, monat, tag, name, wochentag, baustelle, stunden, created_at, updated_at))
            
            # Drop old table
            cursor.execute('DROP TABLE IF EXISTS stunden_eintraege')
            
            # Update schema version
            cursor.execute('UPDATE schema_version SET version = 5 WHERE id = 1')
            current_version = 5
            print("Migration to schema version 5 completed successfully!")


        conn.commit()
        conn.close()

    def add_arbeitsstunden(self, data: Dict) -> int:
        """
        Add a work hours entry to arbeitsstunden table.
        Multiple entries per day/worker are allowed.
        
        Args:
            data: Dictionary with keys: Jahr, Monat, Tag, Name, Wochentag, Kostenstelle, Stunden
        
        Returns:
            ID of the inserted row
        """
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT INTO arbeitsstunden
                (jahr, monat, tag, name, wochentag, kostenstelle, stunden)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('jahr'),
                data.get('monat'),
                data.get('tag'),
                data.get('name'),
                data.get('wochentag'),
                data.get('kostenstelle'),
                data.get('stunden', 0.0)
            ))
            
            conn.commit()
            return cursor.lastrowid
        
        except sqlite3.Error as e:
            print(f"Database error adding arbeitsstunden: {e}")
            conn.rollback()
            raise
        finally:
            conn.close()
    
    def get_arbeitsstunden_for_day(self, year: int, month: int, day: int, name: str) -> List[Dict]:
        """Get all arbeitsstunden entries for a specific person on a specific date."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM arbeitsstunden 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        ''', (year, month, day, name))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_arbeitsstunden_for_month(self, year: int, month: int, name: str) -> List[Dict]:
        """Get all arbeitsstunden entries for a specific person in a specific month."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM arbeitsstunden 
            WHERE jahr = ? AND monat = ? AND name = ?
        ''', (year, month, name))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]

    def update_arbeitsstunden(self, entry_id: int, data: Dict) -> bool:
        """Update an arbeitsstunden entry by ID."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            updates = []
            params = []
            
            fields_map = {
                'Kostenstelle': 'kostenstelle',
                'Stunden': 'stunden'
            }
            
            for data_key, db_col in fields_map.items():
                if data_key in data:
                    updates.append(f"{db_col}=?")
                    params.append(data[data_key])
            
            updates.append("updated_at=CURRENT_TIMESTAMP")
            
            if updates:
                query = f"UPDATE arbeitsstunden SET {', '.join(updates)} WHERE id = ?"
                params.append(entry_id)
                cursor.execute(query, params)
                conn.commit()
                return True
            return False
        
        except sqlite3.Error as e:
            print(f"Database error updating arbeitsstunden: {e}")
            conn.rollback()
            return False
        finally:
            conn.close()
    
    def delete_arbeitsstunden(self, entry_id: int) -> bool:
        """Delete an arbeitsstunden entry by ID."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute('DELETE FROM arbeitsstunden WHERE id = ?', (entry_id,))
            conn.commit()
            return cursor.rowcount > 0
        
        except sqlite3.Error as e:
            print(f"Database error deleting arbeitsstunden: {e}")
            conn.rollback()
            return False
        finally:
            conn.close()

    def add_or_update_metadata(self, data: Dict) -> tuple[int, bool]:
        """
        Add a new entry or update if the combination of jahr, monat, tag, name exists.
        Returns (row_id, was_updated) where was_updated is True if existing row was updated.
        """
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            jahr = data.get('jahr')
            monat = data.get('monat')
            tag = data.get('tag')
            name = data.get('name')
            wochentag = data.get('wochentag')
            
            # Handle tages_metadaten (unique per day/worker)
            metadata_fields = {
                'skug': data.get('skug'),
                'kg_8h': data.get('kg_8h'),
                'travel_status': data.get('travel_status'),
                'fruehstueck': data.get('fruehstueck'),
                'mittag': data.get('mittag')
            }
            
            # Check if metadata entry exists
            cursor.execute('''
                SELECT id FROM tages_metadaten 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            ''', (jahr, monat, tag, name))
            
            existing_metadata = cursor.fetchone()
            
            if existing_metadata:
                # Update metadata
                metadata_id = existing_metadata[0]
                updates = []
                params = []
                
                for field, value in metadata_fields.items():
                    if field in ['skug', 'kg_8h', 'travel_status', 'fruehstueck', 'mittag']:
                        updates.append(f"{field}=?")
                        params.append(value)
                
                if updates:
                    updates.append("updated_at=CURRENT_TIMESTAMP")
                    query = f"UPDATE tages_metadaten SET {', '.join(updates)} WHERE id = ?"
                    params.append(metadata_id)
                    cursor.execute(query, params)
            else:
                # Insert new metadata
                cursor.execute('''
                    INSERT INTO tages_metadaten
                    (jahr, monat, tag, name, wochentag, skug, kg_8h, travel_status, fruehstueck, mittag)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (jahr, monat, tag, name, wochentag, 
                      metadata_fields['skug'], metadata_fields['kg_8h'], 
                      metadata_fields['travel_status'], metadata_fields['fruehstueck'], 
                      metadata_fields['mittag']))
                metadata_id = cursor.lastrowid
            conn.commit()
            return (metadata_id, existing_metadata is not None)
        
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            conn.rollback()
            raise
        
        finally:
            conn.close()

    def get_metadata_by_date(self, year:int, month:int, day:int, name:str) -> Dict:
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                SELECT * FROM tages_metadaten 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            ''', (year, month, day, name))
            
            metadata = cursor.fetchone()
            
            if metadata:
                return {
                    'id': metadata[0],
                    'jahr': metadata[1],
                    'monat': metadata[2],
                    'tag': metadata[3],
                    'name': metadata[4],
                    'wochentag': metadata[5],
                    'skug': metadata[6],
                    'kg_8h': metadata[7],
                    'travel_status': metadata[8],
                    'fruehstueck': metadata[9],
                    'mittag': metadata[10]
                }
            return None
        
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            raise
        
        finally:
            conn.close()

    def get_metadata_for_month(self, year:int, month:int, name:str) -> List[Dict]:
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                SELECT * FROM tages_metadaten 
                WHERE jahr = ? AND monat = ? AND name = ?
            ''', (year, month, name))
            
            metadata = cursor.fetchall()
            
            if metadata:
                return [
                    {
                        'id': metadata[0],
                        'jahr': metadata[1],
                        'monat': metadata[2],
                        'tag': metadata[3],
                        'name': metadata[4],
                        'wochentag': metadata[5],
                        'skug': metadata[6],
                        'kg_8h': metadata[7],
                        'travel_status': metadata[8],
                        'fruehstueck': metadata[9],
                        'mittag': metadata[10]
                    }
                ]
            return None
        
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            raise
        
        finally:
            conn.close()

    def update_entry_metadata(self, entry_id: int, data: Dict) -> bool:
        """Update an existing metadata entry by ID (tages_metadaten table)."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            updates = []
            params = []
            
            # List of possible fields to update in tages_metadaten
            fields_map = {
                'SKUG': 'skug',
                'kg_8h': 'kg_8h',
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
                query = f"UPDATE tages_metadaten SET {', '.join(updates)} WHERE id = ?"
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
    
    def get_all_entries(self) -> List[Dict]:
        """Retrieve all entries from the database (joins both tables)."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Join both tables to get complete data
        cursor.execute('''
            SELECT 
                tm.id, tm.jahr, tm.monat, tm.tag, tm.name, tm.wochentag,
                tm.skug, tm.kg_8h, tm.travel_status, tm.fruehstueck, tm.mittag,
                tm.created_at, tm.updated_at,
                GROUP_CONCAT(a.kostenstelle || ':' || a.stunden, ';') as arbeitsstunden_data
            FROM tages_metadaten tm
            LEFT JOIN arbeitsstunden a ON 
                tm.jahr = a.jahr AND tm.monat = a.monat AND 
                tm.tag = a.tag AND tm.name = a.name
            GROUP BY tm.id, tm.jahr, tm.monat, tm.tag, tm.name
            ORDER BY tm.jahr DESC, tm.monat DESC, tm.tag DESC
        ''')
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_entries_by_month_and_name(self, year: int, month: int, name: str) -> List[Dict]:
        """Get all entries for a specific person in a specific month (joins both tables)."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        # {arbeitsstunden_data, "metadata":{metadata}}

        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_entry(self, year: int, month: int, day: int, name: str) -> Optional[Dict]:
        """Get a single entry for a specific person on a specific date (joins both tables)."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Get metadata
        cursor.execute('''
            SELECT * FROM tages_metadaten 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        ''', (year, month, day, name))
        
        metadata = cursor.fetchone()
        
        if not metadata:
            conn.close()
            return None
        
        result = dict(metadata)
        
        # Get arbeitsstunden
        cursor.execute('''
            SELECT * FROM arbeitsstunden 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        ''', (year, month, day, name))
        
        arbeitsstunden = cursor.fetchall()
        result['arbeitsstunden'] = [dict(row) for row in arbeitsstunden]
        
        conn.close()
        return result
    
    def clear_entries_for_day(self, year: int, month: int, day: int, name: str) -> int:
        """Clear all entries for a specific day (both metadata and arbeitsstunden)."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        print("Clear entries for day:", year, month, day, name)
        try:
            # Delete from arbeitsstunden first
            cursor.execute('''
                DELETE FROM arbeitsstunden 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            ''', (year, month, day, name))
            arbeitsstunden_count = cursor.rowcount
            
            # Delete from tages_metadaten
            cursor.execute('''
                DELETE FROM tages_metadaten 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            ''', (year, month, day, name))
            metadata_count = cursor.rowcount
            
            conn.commit()
            return arbeitsstunden_count + metadata_count
        
        except sqlite3.Error as e:
            print(f"Database error clearing entries: {e}")
            conn.rollback()
            return 0
        finally:
            conn.close()

    def get_entries_for_day(self, year: int, month: int, day: int, name: str) -> List[Dict]:
        """Get all entries for a specific person on a specific date (returns combined data)."""
        # This is similar to get_entry but returns a list for compatibility
        entry = self.get_entry(year, month, day, name)
        return [entry] if entry else []
    
    def get_entries_by_date_and_baustelle(self, year: int, month: int, day: int, kostenstelle: str) -> List[Dict]:
        """Get all entries for a specific construction site (kostenstelle) on a specific date."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Query arbeitsstunden table with kostenstelle
        # Try exact match first, then partial match (for "Nummer - Name" format)
        cursor.execute('''
            SELECT a.*, tm.skug, tm.kg_8h, tm.travel_status, tm.fruehstueck, tm.mittag
            FROM arbeitsstunden a
            LEFT JOIN tages_metadaten tm ON 
                a.jahr = tm.jahr AND a.monat = tm.monat AND 
                a.tag = tm.tag AND a.name = tm.name
            WHERE a.jahr = ? AND a.monat = ? AND a.tag = ? 
                AND (a.kostenstelle = ? OR a.kostenstelle LIKE ?)
            ORDER BY a.name ASC
        ''', (year, month, day, kostenstelle, f"{kostenstelle}%"))

        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]
    
    def get_entries_by_date(self, year: int, month: int) -> List[Dict]:
        """Get entries for a specific month (joins both tables)."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT 
                tm.id, tm.jahr, tm.monat, tm.tag, tm.name, tm.wochentag,
                tm.skug, tm.kg_8h, tm.travel_status, tm.fruehstueck, tm.mittag,
                tm.created_at, tm.updated_at,
                GROUP_CONCAT(a.kostenstelle || ':' || a.stunden, ';') as arbeitsstunden_data
            FROM tages_metadaten tm
            LEFT JOIN arbeitsstunden a ON 
                tm.jahr = a.jahr AND tm.monat = a.monat AND 
                tm.tag = a.tag AND tm.name = a.name
            WHERE tm.jahr = ? AND tm.monat = ?
            GROUP BY tm.id, tm.jahr, tm.monat, tm.tag, tm.name
            ORDER BY tm.tag ASC
        ''', (year, month))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def delete_entry_metadata(self, entry_id: int) -> bool:
        """Delete a metadata entry by ID (tages_metadaten table)."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        try:
            cursor.execute('DELETE FROM tages_metadaten WHERE id = ?', (entry_id,))
            conn.commit()
            success = cursor.rowcount > 0
            return success
        
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            conn.rollback()
            return False
        
        finally:
            conn.close()
        
