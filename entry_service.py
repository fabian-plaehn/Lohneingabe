from utils import (
    calculate_skug,
    determine_kg_8h_flag,
    get_weekday_abbr,
    handle_krank_urlaub,
)


class EntryService:
    def __init__(self, db, master_db):
        self.db = db
        self.master_db = master_db

    def parse_hours_input(self, raw_value):
        if raw_value is None:
            return None
        raw_text = str(raw_value).strip()
        if not raw_text:
            return None
        raw_text = raw_text.replace(",", ".")
        try:
            hours = float(raw_text)
        except ValueError:
            return None
        if not (0 <= hours <= 24):
            return None
        return hours

    def is_valid_kostenstelle(self, kostenstelle_input):
        if not kostenstelle_input:
            return False
        if kostenstelle_input in ["Krank", "900", "940"]:
            return True
        return any(
            f"{b['nummer']} - {b['name']}" == kostenstelle_input
            for b in self.master_db.get_all_baustellen()
        )

    def normalize_kostenstelle_input(self, raw_value):
        if raw_value is None:
            return None
        text = str(raw_value).strip()
        if not text:
            return None
        if text in ["Krank", "900", "940"]:
            return text
        if text.isdigit():
            bst = self.master_db.get_baustelle_by_nummer(int(text))
            if bst:
                return f"{bst['nummer']} - {bst['name']}"
        if self.is_valid_kostenstelle(text):
            return text
        return None

    def create_entry_from_preview(self, year, month, day, name, stunden, kostenstelle):
        if kostenstelle in ["Krank", "900", "940"]:
            skug_settings = self.master_db.get_skug_settings()
            handle_krank_urlaub(
                year,
                month,
                day,
                name,
                self.db,
                self.master_db,
                kostenstelle == "Krank",
                kostenstelle in ["900", "940"],
                skug_settings,
            )
            return True
        wochentag = get_weekday_abbr(year, month, str(day)) or ""
        entry_data = {
            "jahr": year,
            "monat": month,
            "tag": str(day),
            "name": name,
            "wochentag": wochentag,
            "stunden": stunden,
            "kostenstelle": kostenstelle,
        }
        self.db.add_arbeitsstunden(entry_data)
        return False

    def recompute_metadata_for_day(self, year, month, day, name):
        metadata_entry = self.db.get_metadata_by_date(year, month, day, name)
        if metadata_entry is None:
            metadata_entry = {}
        metadata_entry.update(
            {
                "jahr": year,
                "monat": month,
                "tag": str(day),
                "name": name,
                "wochentag": get_weekday_abbr(year, month, str(day)) or "",
            }
        )

        is_winter = month in [12, 1, 2, 3]
        if is_winter and not metadata_entry.get("no_skug", False):
            skug_settings = self.master_db.get_skug_settings()
            total_hours = sum(
                e.get("stunden", 0) or 0
                for e in self.db.get_arbeitsstunden_for_day(year, month, day, name)
            )
            metadata_entry["skug"] = calculate_skug(
                year, month, day, total_hours, skug_settings
            )
        else:
            metadata_entry["skug"] = None

        metadata_entry["kg_8h"] = determine_kg_8h_flag(
            self.db, self.master_db, year, month, day, name
        )
        self.db.add_or_update_metadata(metadata_entry)

    def apply_preview_changes(self, pending_edits, pending_flags):
        errors = []
        edit_ops = {}

        krank_urlaub_keys = {
            key
            for key, flags in pending_flags.items()
            if flags.get("krank") or flags.get("urlaub")
        }

        for key, data in pending_edits.items():
            cell_info = data.get("cell_info", {})
            value = data.get("value")
            field = cell_info.get("field")
            year = cell_info.get("year")
            month = cell_info.get("month")
            day = cell_info.get("day")
            name = cell_info.get("name")
            entry_id = cell_info.get("entry_id")
            wb_row = key[0] + 1

            if not all([year, month, day, name]):
                errors.append("Ungültige Zellzuordnung.")
                continue

            if (year, month, day, name) in krank_urlaub_keys:
                continue

            op_key = (year, month, day, name, entry_id, wb_row)
            if op_key not in edit_ops:
                edit_ops[op_key] = {}

            if field == "Stunden":
                stunden = self.parse_hours_input(value)
                if stunden is None:
                    errors.append(
                        f"Ungültige Stunden bei {name} am {day}.{month}.{year}."
                    )
                    continue
                edit_ops[op_key]["stunden"] = stunden
            elif field == "Kostenstelle":
                bst_value = self.normalize_kostenstelle_input(value)
                if not bst_value:
                    errors.append(
                        f"Kostenstelle fehlt/ungültig bei {name} am {day}.{month}.{year}."
                    )
                    continue
                if bst_value in ["Krank", "900", "940"]:
                    errors.append(
                        f"Krank/Urlaub bitte per Rechtsklick setzen: {name} am {day}.{month}.{year}."
                    )
                    continue
                edit_ops[op_key]["kostenstelle"] = bst_value
            else:
                errors.append("Nicht editierbare Zelle geändert.")

        for key, flags in pending_flags.items():
            year, month, day, name = key
            if flags.get("krank") and flags.get("urlaub"):
                errors.append(
                    f"Krank und Urlaub gleichzeitig bei {name} am {day}.{month}.{year}."
                )

        for op_key, values in edit_ops.items():
            year, month, day, name, entry_id, wb_row = op_key
            if entry_id is None and (
                "stunden" not in values or "kostenstelle" not in values
            ):
                errors.append(
                    f"Stunden und Kostenstelle erforderlich bei {name} am {day}.{month}.{year}."
                )

        if errors:
            return errors, set()

        affected_days = set()
        for key, flags in pending_flags.items():
            year, month, day, name = key
            if flags.get("krank") or flags.get("urlaub"):
                skug_settings = self.master_db.get_skug_settings()
                handle_krank_urlaub(
                    year,
                    month,
                    day,
                    name,
                    self.db,
                    self.master_db,
                    bool(flags.get("krank")),
                    bool(flags.get("urlaub")),
                    skug_settings,
                )
                affected_days.add((year, month, day, name))

        for op_key, values in edit_ops.items():
            year, month, day, name, entry_id, wb_row = op_key
            if (year, month, day, name) in affected_days:
                continue

            if entry_id is None:
                self.create_entry_from_preview(
                    year, month, day, name, values["stunden"], values["kostenstelle"]
                )
            else:
                update_data = {}
                if "stunden" in values:
                    update_data["Stunden"] = values["stunden"]
                if "kostenstelle" in values:
                    update_data["Kostenstelle"] = values["kostenstelle"]
                if update_data:
                    if not self.db.update_arbeitsstunden(entry_id, update_data):
                        return ["Änderung konnte nicht gespeichert werden."], set()

            affected_days.add((year, month, day, name))

        for key, flags in pending_flags.items():
            year, month, day, name = key
            if flags.get("krank") or flags.get("urlaub"):
                continue
            metadata_entry = self.db.get_metadata_by_date(year, month, day, name)
            if metadata_entry is None:
                metadata_entry = {}
            metadata_entry.update(
                {
                    "jahr": year,
                    "monat": month,
                    "tag": str(day),
                    "name": name,
                    "wochentag": get_weekday_abbr(year, month, str(day)) or "",
                }
            )
            if "fruehstueck" in flags:
                metadata_entry["fruehstueck"] = flags["fruehstueck"]
            if "mittag" in flags:
                metadata_entry["mittag"] = flags["mittag"]
            if "no_skug" in flags:
                metadata_entry["no_skug"] = flags["no_skug"]
            if "travel_status" in flags:
                metadata_entry["travel_status"] = flags["travel_status"]
            self.db.add_or_update_metadata(metadata_entry)
            affected_days.add((year, month, day, name))

        for year, month, day, name in affected_days:
            if (year, month, day, name) in pending_flags and (
                pending_flags[(year, month, day, name)].get("krank")
                or pending_flags[(year, month, day, name)].get("urlaub")
            ):
                continue
            self.recompute_metadata_for_day(year, month, day, name)

        return errors, affected_days
