import sqlite3
from datetime import datetime
from typing import List, Dict, Optional
import pandas as pd

class Database:
    def __init__(self, db_file="stundenliste.db"):
        self.db_file = db_file
        self.init_database()
    
    def init_database(self):
        """Create table if it doesn't exist."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS stunden_eintraege (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                jahr INTEGER NOT NULL,
                monat INTEGER NOT NULL,
                tag INTEGER NOT NULL,
                name TEXT NOT NULL,
                wochentag TEXT,
                stunden REAL NOT NULL,
                checkbox1 BOOLEAN,
                checkbox2 BOOLEAN,
                skug TEXT,
                baustelle TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(jahr, monat, tag, name)
            )
        ''')
        
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
                # Update existing entry
                entry_id = existing[0]
                cursor.execute('''
                    UPDATE stunden_eintraege 
                    SET wochentag=?, stunden=?, checkbox1=?, checkbox2=?, 
                        skug=?, baustelle=?, updated_at=CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (
                    data.get('Wochentag'),
                    data.get('Stunden'),
                    data.get('CheckBox1'),
                    data.get('CheckBox2'),
                    data.get('SKUG'),
                    data.get('Baustelle'),
                    entry_id
                ))
                conn.commit()
                return (entry_id, True)
            else:
                # Insert new entry
                cursor.execute('''
                    INSERT INTO stunden_eintraege 
                    (jahr, monat, tag, name, wochentag, stunden, checkbox1, checkbox2, skug, baustelle)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    data.get('Jahr'),
                    data.get('Monat'),
                    data.get('Tag'),
                    data.get('Name'),
                    data.get('Wochentag'),
                    data.get('Stunden'),
                    data.get('CheckBox1'),
                    data.get('CheckBox2'),
                    data.get('SKUG'),
                    data.get('Baustelle')
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
    
    def get_entries_by_date_and_baustelle(self, year: int, month: int, day: int, baustelle: str) -> List[Dict]:
        """Get all entries for a specific construction site on a specific date."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM stunden_eintraege 
            WHERE jahr = ? AND monat = ? AND tag = ? AND baustelle = ?
            ORDER BY name ASC
        ''', (year, month, day, baustelle))
        
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