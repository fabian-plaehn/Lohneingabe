import os
import shutil
import sqlite3
from datetime import datetime as dt
from typing import List, Dict, Optional

import pandas as pd


class Database:
    SCHEMA_VERSION = 9

    def __init__(self, db_file="stundenliste.db", master_db=None):
        self.db_file = db_file
        self.master_db = master_db
        self.init_database()

    def set_master_db(self, master_db):
        self.master_db = master_db

    def _build_metadata_base(self, year: int, month: int, day: int, name: str) -> Dict:
        return {
            "jahr": year,
            "monat": month,
            "tag": day,
            "name": name,
            "wochentag": None,
            "skug": None,
            "no_skug": False,
            "kg_8h": None,
            "travel_status": None,
            "fruehstueck": False,
            "mittag": False,
            "urlaub": None,
            "krank": None,
        }

    def _calculate_skug_value(
        self, year: int, month: int, day: int, metadata: Dict
    ) -> Optional[float]:
        if month not in [12, 1, 2, 3] or metadata.get("no_skug", False):
            return None
        if self.master_db is None:
            raise Exception("Master database not set for kg_8h calculation")
        
        name_data = self.master_db.get_name_by_name(metadata["name"])
        if name_data.get("kein_fzk", False):
            return None

        try:
            from utils import calculate_skug

            total_hours = sum(
                float(entry.get("stunden") or 0.0)
                for entry in self.get_arbeitsstunden_for_day(
                    year, month, day, metadata["name"]
                )
            )
            skug_settings = self.master_db.get_skug_settings()
            skug = calculate_skug(year, month, day, total_hours, skug_settings)
            if skug is None:
                return None
            return skug if skug >= 1 else 0
        except Exception:
            return metadata.get("skug")

    def _calculate_kg_8h_value(self, year: int, month: int, day: int, metadata: Dict):
        if self.master_db is None:
            raise Exception("Master database not set for kg_8h calculation")
        if metadata.get("krank") or metadata.get("urlaub"):
            return None
        if metadata.get("travel_status"):
            return None
        try:
            from utils import get_effective_fahrzeit

            total_hours = 0.0
            highest_fahrzeit = 0.0
            worker_id = self.master_db.get_worker_id_by_name(metadata["name"])

            for entry in self.get_arbeitsstunden_for_day(
                year, month, day, metadata["name"]
            ):
                total_hours += float(entry.get("stunden") or 0.0)
                kostenstelle = entry.get("kostenstelle")
                if not kostenstelle:
                    continue

                bst_nummer = (
                    kostenstelle.split("-")[0].strip()
                    if "-" in kostenstelle
                    else str(kostenstelle).strip()
                )
                bst_data = self.master_db.get_baustelle_by_nummer(bst_nummer)
                if not bst_data:
                    continue

                fahrzeit = get_effective_fahrzeit(
                    self.master_db,
                    worker_id,
                    bst_data["id"],
                    bst_data.get("fahrzeit", 0.0),
                )
                highest_fahrzeit = max(highest_fahrzeit, float(fahrzeit or 0.0))

            total_hours += highest_fahrzeit
            if metadata.get("fruehstueck"):
                total_hours += 0.25
            if metadata.get("mittag"):
                total_hours += 0.5
            return total_hours <= 8.0
        except Exception:
            return metadata.get("kg_8h")

    def _resolve_metadata_entry(
        self, metadata: Optional[Dict], year: int, month: int, day: int, name: str
    ) -> Dict:
        resolved = self._build_metadata_base(year, month, day, name)
        if metadata:
            resolved.update(metadata)
        resolved["jahr"] = year
        resolved["monat"] = month
        resolved["tag"] = day
        resolved["name"] = name
        resolved["skug"] = self._calculate_skug_value(year, month, day, resolved)
        resolved["kg_8h"] = self._calculate_kg_8h_value(year, month, day, resolved)
        return resolved

    def get_stored_metadata_by_date(
        self, year: int, month: int, day: int, name: str
    ) -> Optional[Dict]:
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        try:
            cursor.execute(
                """
                SELECT * FROM tages_metadaten
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            """,
                (year, month, day, name),
            )

            metadata = cursor.fetchone()
            return dict(metadata) if metadata else None

        except sqlite3.Error as e:
            print(f"Database error: {e}")
            raise

        finally:
            conn.close()

    def _backup_database(self, from_version: int, to_version: int):
        backup_folder = os.path.join(
            os.path.dirname(self.db_file), f"Backup_{from_version}_to_{to_version}"
        )
        os.makedirs(backup_folder, exist_ok=True)

        timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(
            backup_folder,
            f"stundenliste_backup_v{from_version}_to_v{to_version}_{timestamp}.db",
        )

        shutil.copy2(self.db_file, backup_file)
        print(f"Database backed up to: {backup_file}")

    def _backup_and_reconnect(self, conn, from_version: int, to_version: int):
        conn.commit()
        conn.close()
        self._backup_database(from_version, to_version)
        new_conn = sqlite3.connect(self.db_file)
        return new_conn, new_conn.cursor()

    def init_database(self):
        """Create table if it doesn't exist."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        # Schema version table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS schema_version (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                version INTEGER NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        # Check and set schema version
        cursor.execute("SELECT version FROM schema_version WHERE id = 1")
        row = cursor.fetchone()
        current_version = 0
        if row:
            current_version = row[0]
        else:
            cursor.execute(
                "INSERT INTO schema_version (id, version) VALUES (1, ?)",
                (self.SCHEMA_VERSION,),
            )
            current_version = self.SCHEMA_VERSION

        cursor.execute("""
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
        """)
        if current_version < 2:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 2)
            try:
                cursor.execute(
                    "ALTER TABLE stunden_eintraege RENAME COLUMN unter_8h TO kg_8h"
                )
            except sqlite3.OperationalError:
                pass  # Column already renamed
            cursor.execute("UPDATE schema_version SET version = 2 WHERE id = 1")
            current_version = 2

        if current_version < 3:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 3)
            try:
                cursor.execute(
                    "ALTER TABLE stunden_eintraege ADD COLUMN fruehstueck BOOLEAN"
                )
                cursor.execute(
                    "ALTER TABLE stunden_eintraege ADD COLUMN mittag BOOLEAN"
                )
            except sqlite3.OperationalError:
                pass  # Columns likely already exist
            cursor.execute("UPDATE schema_version SET version = 3 WHERE id = 1")
            current_version = 3

        if current_version < 4:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 4)
            # Remove UNIQUE constraint by recreating table
            cursor.execute(
                "CREATE TABLE stunden_eintraege_new ("
                "id INTEGER PRIMARY KEY AUTOINCREMENT,"
                "jahr INTEGER NOT NULL,"
                "monat INTEGER NOT NULL,"
                "tag INTEGER NOT NULL,"
                "name TEXT NOT NULL,"
                "wochentag TEXT,"
                "stunden REAL NOT NULL,"
                "urlaub TEXT,"
                "krank TEXT,"
                "kg_8h BOOLEAN,"
                "skug TEXT,"
                "baustelle TEXT,"
                "travel_status TEXT,"
                "fruehstueck BOOLEAN,"
                "mittag BOOLEAN,"
                "created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,"
                "updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP"
                ")"
            )

            # Copy data
            cursor.execute(
                "INSERT INTO stunden_eintraege_new (id, jahr, monat, tag, name, wochentag, stunden, urlaub, krank, kg_8h, skug, baustelle, travel_status, fruehstueck, mittag, created_at, updated_at) "
                "SELECT id, jahr, monat, tag, name, wochentag, stunden, urlaub, krank, kg_8h, skug, baustelle, travel_status, fruehstueck, mittag, created_at, updated_at FROM stunden_eintraege"
            )

            cursor.execute("DROP TABLE stunden_eintraege")
            cursor.execute(
                "ALTER TABLE stunden_eintraege_new RENAME TO stunden_eintraege"
            )

            cursor.execute("UPDATE schema_version SET version = 4 WHERE id = 1")
            current_version = 4

        if current_version < 5:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 5)

            # Create new tages_metadaten table (unique per day/worker)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS tages_metadaten (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    jahr INTEGER NOT NULL,
                    monat INTEGER NOT NULL,
                    tag INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    wochentag TEXT,
                    skug TEXT,
                    no_skug BOOLEAN DEFAULT 0,
                    kg_8h BOOLEAN,
                    travel_status TEXT,
                    fruehstueck BOOLEAN,
                    mittag BOOLEAN,
                    urlaub TEXT,
                    krank TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(jahr, monat, tag, name)
                )
            """)

            # Create new arbeitsstunden table (multiple entries per day/worker allowed)
            cursor.execute("""
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
            """)

            # Migrate data from old table to new tables
            cursor.execute("SELECT * FROM stunden_eintraege")
            old_entries = cursor.fetchall()

            for entry in old_entries:
                # entry structure: id, jahr, monat, tag, name, wochentag, stunden, urlaub, krank,
                #                  kg_8h, skug, baustelle, travel_status, fruehstueck, mittag, created_at, updated_at
                (
                    entry_id,
                    jahr,
                    monat,
                    tag,
                    name,
                    wochentag,
                    stunden,
                    urlaub,
                    krank,
                    kg_8h,
                    skug,
                    baustelle,
                    travel_status,
                    fruehstueck,
                    mittag,
                    created_at,
                    updated_at,
                ) = entry

                # Insert into tages_metadaten (one entry per day/worker)
                try:
                    cursor.execute(
                        """
                        INSERT OR IGNORE INTO tages_metadaten 
                        (jahr, monat, tag, name, wochentag, skug, kg_8h, travel_status, fruehstueck, mittag, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            jahr,
                            monat,
                            tag,
                            name,
                            wochentag,
                            skug,
                            kg_8h,
                            travel_status,
                            fruehstueck,
                            mittag,
                            created_at,
                            updated_at,
                        ),
                    )
                except sqlite3.IntegrityError:
                    # Entry already exists, skip
                    pass

                # Insert into arbeitsstunden
                # Handle Urlaub/Krank as kostenstelle entries
                if urlaub:
                    cursor.execute(
                        """
                        INSERT INTO arbeitsstunden 
                        (jahr, monat, tag, name, wochentag, kostenstelle, stunden, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            jahr,
                            monat,
                            tag,
                            name,
                            wochentag,
                            "Urlaub",
                            float(urlaub) if urlaub else 0.0,
                            created_at,
                            updated_at,
                        ),
                    )

                if krank:
                    cursor.execute(
                        """
                        INSERT INTO arbeitsstunden 
                        (jahr, monat, tag, name, wochentag, kostenstelle, stunden, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            jahr,
                            monat,
                            tag,
                            name,
                            wochentag,
                            "Krank",
                            float(krank) if krank else 0.0,
                            created_at,
                            updated_at,
                        ),
                    )

                # Insert regular work hours with baustelle as kostenstelle
                if stunden and stunden > 0 and baustelle:
                    cursor.execute(
                        """
                        INSERT INTO arbeitsstunden 
                        (jahr, monat, tag, name, wochentag, kostenstelle, stunden, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        (
                            jahr,
                            monat,
                            tag,
                            name,
                            wochentag,
                            baustelle,
                            stunden,
                            created_at,
                            updated_at,
                        ),
                    )

            # Drop old table
            cursor.execute("DROP TABLE IF EXISTS stunden_eintraege")

            # Update schema version
            cursor.execute("UPDATE schema_version SET version = 5 WHERE id = 1")
            current_version = 5
            print("Migration to schema version 5 completed successfully!")

        if current_version < 6:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 6)
            try:
                cursor.execute("ALTER TABLE tages_metadaten ADD COLUMN urlaub TEXT")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE tages_metadaten ADD COLUMN krank TEXT")
            except sqlite3.OperationalError:
                pass
            cursor.execute("UPDATE schema_version SET version = 6 WHERE id = 1")
            current_version = 6

        if current_version < 7:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 7)
            try:
                cursor.execute(
                    "ALTER TABLE tages_metadaten ADD COLUMN no_skug BOOLEAN DEFAULT 0"
                )
            except sqlite3.OperationalError:
                pass
            cursor.execute("UPDATE schema_version SET version = 7 WHERE id = 1")
            current_version = 7

        if current_version < 8:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 8)
            try:
                cursor.execute("ALTER TABLE tages_metadaten ADD COLUMN urlaub TEXT")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE tages_metadaten ADD COLUMN krank TEXT")
            except sqlite3.OperationalError:
                pass

            cursor.execute(
                """
                INSERT OR IGNORE INTO tages_metadaten (jahr, monat, tag, name, wochentag, urlaub, krank)
                SELECT
                    jahr,
                    monat,
                    tag,
                    name,
                    MIN(wochentag) AS wochentag,
                    CASE
                        WHEN SUM(CASE WHEN kostenstelle = '940' THEN COALESCE(stunden, 0) ELSE 0 END) > 0
                        THEN CAST(SUM(CASE WHEN kostenstelle = '940' THEN COALESCE(stunden, 0) ELSE 0 END) AS TEXT)
                    END AS urlaub,
                    CASE
                        WHEN SUM(CASE WHEN kostenstelle = 'Krank' THEN COALESCE(stunden, 0) ELSE 0 END) > 0
                        THEN CAST(SUM(CASE WHEN kostenstelle = 'Krank' THEN COALESCE(stunden, 0) ELSE 0 END) AS TEXT)
                    END AS krank
                FROM arbeitsstunden
                WHERE kostenstelle IN ('940', 'Krank')
                GROUP BY jahr, monat, tag, name
                """
            )

            cursor.execute(
                """
                UPDATE tages_metadaten
                SET urlaub = CASE
                        WHEN urlaub IS NULL OR urlaub = '' THEN (
                            SELECT CAST(SUM(COALESCE(a.stunden, 0)) AS TEXT)
                            FROM arbeitsstunden a
                            WHERE a.jahr = tages_metadaten.jahr
                              AND a.monat = tages_metadaten.monat
                              AND a.tag = tages_metadaten.tag
                              AND a.name = tages_metadaten.name
                              AND a.kostenstelle = '940'
                            GROUP BY a.jahr, a.monat, a.tag, a.name
                        )
                        ELSE urlaub
                    END,
                    krank = CASE
                        WHEN krank IS NULL OR krank = '' THEN (
                            SELECT CAST(SUM(COALESCE(a.stunden, 0)) AS TEXT)
                            FROM arbeitsstunden a
                            WHERE a.jahr = tages_metadaten.jahr
                              AND a.monat = tages_metadaten.monat
                              AND a.tag = tages_metadaten.tag
                              AND a.name = tages_metadaten.name
                              AND a.kostenstelle = 'Krank'
                            GROUP BY a.jahr, a.monat, a.tag, a.name
                        )
                        ELSE krank
                    END,
                    updated_at = CURRENT_TIMESTAMP
                WHERE EXISTS (
                    SELECT 1
                    FROM arbeitsstunden a
                    WHERE a.jahr = tages_metadaten.jahr
                      AND a.monat = tages_metadaten.monat
                      AND a.tag = tages_metadaten.tag
                      AND a.name = tages_metadaten.name
                      AND a.kostenstelle IN ('940', 'Krank')
                )
                """
            )

            cursor.execute("UPDATE schema_version SET version = 8 WHERE id = 1")
            current_version = 8

        if current_version < 9:
            conn, cursor = self._backup_and_reconnect(conn, current_version, 9)
            cursor.execute(
                """
                DELETE FROM tages_metadaten
                WHERE NOT EXISTS (
                    SELECT 1
                    FROM arbeitsstunden
                    WHERE arbeitsstunden.jahr = tages_metadaten.jahr
                      AND arbeitsstunden.monat = tages_metadaten.monat
                      AND arbeitsstunden.tag = tages_metadaten.tag
                      AND arbeitsstunden.name = tages_metadaten.name
                )
                """
            )
            deleted_count = cursor.rowcount
            print(
                f"Migration to schema version 9 removed {deleted_count} orphaned metadata entries."
            )
            cursor.execute("UPDATE schema_version SET version = 9 WHERE id = 1")
            current_version = 9

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
            cursor.execute(
                """
                INSERT INTO arbeitsstunden
                (jahr, monat, tag, name, wochentag, kostenstelle, stunden)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    data.get("jahr"),
                    data.get("monat"),
                    data.get("tag"),
                    data.get("name"),
                    data.get("wochentag"),
                    data.get("kostenstelle"),
                    data.get("stunden", 0.0),
                ),
            )

            conn.commit()
            return cursor.lastrowid

        except sqlite3.Error as e:
            print(f"Database error adding arbeitsstunden: {e}")
            conn.rollback()
            raise
        finally:
            conn.close()

    def get_arbeitsstunden_for_day(
        self, year: int, month: int, day: int, name: str
    ) -> List[Dict]:
        """Get all arbeitsstunden entries for a specific person on a specific date."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT * FROM arbeitsstunden 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        """,
            (year, month, day, name),
        )

        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]

    def get_arbeitsstunden_for_month(
        self, year: int, month: int, name: str
    ) -> List[Dict]:
        """Get all arbeitsstunden entries for a specific person in a specific month."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT * FROM arbeitsstunden 
            WHERE jahr = ? AND monat = ? AND name = ?
        """,
            (year, month, name),
        )

        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]

    def get_used_baustellen_numbers_for_year(self, year: int) -> List[str]:
        """Get distinct baustellen numbers used in arbeitsstunden for a year."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT DISTINCT
                CASE
                    WHEN instr(kostenstelle, ' - ') > 0
                    THEN trim(substr(kostenstelle, 1, instr(kostenstelle, ' - ') - 1))
                    ELSE trim(kostenstelle)
                END AS baustelle_nummer
            FROM arbeitsstunden
            WHERE jahr = ?
              AND trim(COALESCE(kostenstelle, '')) != ''
              AND trim(kostenstelle) NOT IN ('Krank', '900', '940')
            ORDER BY baustelle_nummer ASC
        """,
            (year,),
        )

        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows if row[0]]

    def update_arbeitsstunden(self, entry_id: int, data: Dict) -> bool:
        """Update an arbeitsstunden entry by ID."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            updates = []
            params = []

            fields_map = {"Kostenstelle": "kostenstelle", "Stunden": "stunden"}

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

    def get_arbeitsstunden_by_id(self, entry_id: int) -> Optional[Dict]:
        """Get a single arbeitsstunden entry by ID."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT * FROM arbeitsstunden
            WHERE id = ?
        """,
            (entry_id,),
        )

        row = cursor.fetchone()
        conn.close()

        return dict(row) if row else None

    def delete_arbeitsstunden(self, entry_id: int) -> bool:
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute(
                "SELECT jahr, monat, tag, name FROM arbeitsstunden WHERE id = ?",
                (entry_id,),
            )
            row = cursor.fetchone()
            if not row:
                return False

            jahr, monat, tag, name = row

            cursor.execute("DELETE FROM arbeitsstunden WHERE id = ?", (entry_id,))
            if cursor.rowcount == 0:
                conn.rollback()
                return False

            cursor.execute(
                """
                SELECT COUNT(*) FROM arbeitsstunden
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
                """,
                (jahr, monat, tag, name),
            )
            remaining = cursor.fetchone()[0]

            if remaining == 0:
                cursor.execute(
                    """
                    DELETE FROM tages_metadaten
                    WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
                    """,
                    (jahr, monat, tag, name),
                )

            conn.commit()
            return True
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
            jahr = data.get("jahr")
            monat = data.get("monat")
            tag = data.get("tag")
            name = data.get("name")
            wochentag = data.get("wochentag")

            # Handle tages_metadaten (unique per day/worker)
            metadata_fields = {
                "no_skug": data.get("no_skug"),
                "travel_status": data.get("travel_status"),
                "fruehstueck": data.get("fruehstueck"),
                "mittag": data.get("mittag"),
                "urlaub": data.get("urlaub"),
                "krank": data.get("krank"),
            }

            # Check if metadata entry exists
            cursor.execute(
                """
                SELECT id FROM tages_metadaten 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            """,
                (jahr, monat, tag, name),
            )

            existing_metadata = cursor.fetchone()

            if existing_metadata:
                # Update metadata
                metadata_id = existing_metadata[0]
                updates = []
                params = []

                for field, value in metadata_fields.items():
                    updates.append(f"{field}=?")
                    params.append(value)

                if updates:
                    updates.append("updated_at=CURRENT_TIMESTAMP")
                    query = (
                        f"UPDATE tages_metadaten SET {', '.join(updates)} WHERE id = ?"
                    )
                    params.append(metadata_id)
                    cursor.execute(query, params)
            else:
                # Insert new metadata
                cursor.execute(
                    """
                    INSERT INTO tages_metadaten
                    (jahr, monat, tag, name, wochentag, no_skug, travel_status, fruehstueck, mittag, urlaub, krank)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        jahr,
                        monat,
                        tag,
                        name,
                        wochentag,
                        metadata_fields["no_skug"],
                        metadata_fields["travel_status"],
                        metadata_fields["fruehstueck"],
                        metadata_fields["mittag"],
                        metadata_fields["urlaub"],
                        metadata_fields["krank"],
                    ),
                )
                metadata_id = cursor.lastrowid
            conn.commit()
            return (metadata_id, existing_metadata is not None)

        except sqlite3.Error as e:
            print(f"Database error: {e}")
            conn.rollback()
            raise

        finally:
            conn.close()

    def get_metadata_by_date(
        self, year: int, month: int, day: int, name: str
    ) -> Optional[Dict]:
        metadata = self.get_stored_metadata_by_date(year, month, day, name)
        if metadata is None:
            day_entries = self.get_arbeitsstunden_for_day(year, month, day, name)
            if not day_entries:
                return None
        return self._resolve_metadata_entry(metadata, year, month, day, name)

    def get_metadata_for_month(self, year: int, month: int, name: str) -> List[Dict]:
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        try:
            cursor.execute(
                """
                SELECT * FROM tages_metadaten 
                WHERE jahr = ? AND monat = ? AND name = ?
            """,
                (year, month, name),
            )

            rows = cursor.fetchall()
            conn.close()
            return [
                self._resolve_metadata_entry(dict(row), year, month, row["tag"], name)
                for row in rows
            ]

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
                "SKUG": "skug",
                "no_skug": "no_skug",
                "kg_8h": "kg_8h",
                "travel_status": "travel_status",
                "fruehstueck": "fruehstueck",
                "mittag": "mittag",
                "urlaub": "urlaub",
                "krank": "krank",
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
        cursor.execute("""
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
        """)

        rows = cursor.fetchall()
        conn.close()

        resolved_rows = []
        for row in rows:
            entry = dict(row)
            resolved_rows.append(
                self._resolve_metadata_entry(
                    entry,
                    entry["jahr"],
                    entry["monat"],
                    entry["tag"],
                    entry["name"],
                )
            )

        return resolved_rows

    def get_entries_by_month_and_name(
        self, year: int, month: int, name: str
    ) -> List[Dict]:
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
        cursor.execute(
            """
            SELECT * FROM tages_metadaten 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        """,
            (year, month, day, name),
        )

        metadata = cursor.fetchone()

        if not metadata:
            conn.close()
            return None

        result = self._resolve_metadata_entry(dict(metadata), year, month, day, name)

        # Get arbeitsstunden
        cursor.execute(
            """
            SELECT * FROM arbeitsstunden 
            WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
        """,
            (year, month, day, name),
        )

        arbeitsstunden = cursor.fetchall()
        result["arbeitsstunden"] = [dict(row) for row in arbeitsstunden]

        conn.close()
        return result

    def clear_entries_for_day(self, year: int, month: int, day: int, name: str) -> int:
        """Clear all entries for a specific day (both metadata and arbeitsstunden)."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        print("Clear entries for day:", year, month, day, name)
        try:
            # Delete from arbeitsstunden first
            cursor.execute(
                """
                DELETE FROM arbeitsstunden 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            """,
                (year, month, day, name),
            )
            arbeitsstunden_count = cursor.rowcount

            # Delete from tages_metadaten
            cursor.execute(
                """
                DELETE FROM tages_metadaten 
                WHERE jahr = ? AND monat = ? AND tag = ? AND name = ?
            """,
                (year, month, day, name),
            )
            metadata_count = cursor.rowcount

            conn.commit()
            return arbeitsstunden_count + metadata_count

        except sqlite3.Error as e:
            print(f"Database error clearing entries: {e}")
            conn.rollback()
            return 0
        finally:
            conn.close()

    def get_entries_for_day(
        self, year: int, month: int, day: int, name: str
    ) -> List[Dict]:
        """Get all entries for a specific person on a specific date (returns combined data)."""
        # This is similar to get_entry but returns a list for compatibility
        entry = self.get_entry(year, month, day, name)
        return [entry] if entry else []

    def get_entries_by_date_and_baustelle(
        self, year: int, month: int, day: int, kostenstelle: str
    ) -> List[Dict]:
        """Get all entries for a specific construction site (kostenstelle) on a specific date."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Query arbeitsstunden table with kostenstelle
        # Try exact match first, then partial match (for "Nummer - Name" format)
        cursor.execute(
            """
            SELECT a.*, tm.skug, tm.kg_8h, tm.travel_status, tm.fruehstueck, tm.mittag,
                   tm.no_skug, tm.urlaub, tm.krank, tm.id as metadata_id, tm.wochentag as metadata_wochentag
            FROM arbeitsstunden a
            LEFT JOIN tages_metadaten tm ON 
                a.jahr = tm.jahr AND a.monat = tm.monat AND 
                a.tag = tm.tag AND a.name = tm.name
            WHERE a.jahr = ? AND a.monat = ? AND a.tag = ? 
                AND (a.kostenstelle = ? OR a.kostenstelle LIKE ?)
            ORDER BY a.name ASC
        """,
            (year, month, day, kostenstelle, f"{kostenstelle}%"),
        )

        rows = cursor.fetchall()
        conn.close()

        resolved_rows = []
        for row in rows:
            entry = dict(row)
            resolved_metadata = self._resolve_metadata_entry(
                {
                    "id": entry.get("metadata_id"),
                    "wochentag": entry.get("metadata_wochentag"),
                    "skug": entry.get("skug"),
                    "kg_8h": entry.get("kg_8h"),
                    "travel_status": entry.get("travel_status"),
                    "fruehstueck": entry.get("fruehstueck"),
                    "mittag": entry.get("mittag"),
                    "no_skug": entry.get("no_skug"),
                    "urlaub": entry.get("urlaub"),
                    "krank": entry.get("krank"),
                },
                entry["jahr"],
                entry["monat"],
                entry["tag"],
                entry["name"],
            )
            entry["skug"] = resolved_metadata.get("skug")
            entry["kg_8h"] = resolved_metadata.get("kg_8h")
            entry["travel_status"] = resolved_metadata.get("travel_status")
            entry["fruehstueck"] = resolved_metadata.get("fruehstueck")
            entry["mittag"] = resolved_metadata.get("mittag")
            entry["no_skug"] = resolved_metadata.get("no_skug")
            entry["urlaub"] = resolved_metadata.get("urlaub")
            entry["krank"] = resolved_metadata.get("krank")
            entry["wochentag"] = resolved_metadata.get("wochentag")
            entry.pop("metadata_id", None)
            entry.pop("metadata_wochentag", None)
            resolved_rows.append(entry)

        return resolved_rows

    def get_entries_by_date(self, year: int, month: int) -> List[Dict]:
        """Get entries for a specific month (joins both tables)."""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute(
            """
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
        """,
            (year, month),
        )

        rows = cursor.fetchall()
        conn.close()

        resolved_rows = []
        for row in rows:
            entry = dict(row)
            resolved_rows.append(
                self._resolve_metadata_entry(
                    entry,
                    entry["jahr"],
                    entry["monat"],
                    entry["tag"],
                    entry["name"],
                )
            )

        return resolved_rows

    def delete_entry_metadata(self, entry_id: int) -> bool:
        """Delete a metadata entry by ID (tages_metadaten table)."""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        try:
            cursor.execute("DELETE FROM tages_metadaten WHERE id = ?", (entry_id,))
            conn.commit()
            success = cursor.rowcount > 0
            return success

        except sqlite3.Error as e:
            print(f"Database error: {e}")
            conn.rollback()
            return False

        finally:
            conn.close()
