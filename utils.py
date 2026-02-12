from datetime import datetime, timedelta
import locale
from datetime import datetime
import holidays
import calendar

from database import Database
from master_data import MasterDataDatabase
from datatypes import TravelStatus
from datatypes import WorkerTypes


locale.setlocale(locale.LC_TIME, "de_DE")

# Initialize German holidays
german_holidays = holidays.country_holidays("DE", subdiv="SH")

AN_ODER_ABREISE_VERPFLEGUNG = 14
AWAY_24H_VERPFLEGUNG = 28


def get_weekday_abbr(year, month, day):
    """Returns abbreviated weekday name or None if invalid date."""
    try:
        d = datetime(int(year), int(month), int(day))
        return d.strftime("%a")
    except (ValueError, TypeError):
        return None


def validate_date(year, month, day):
    """Validates if the given date is valid."""
    try:
        datetime(int(year), int(month), int(day))
        return True
    except (ValueError, TypeError):
        return False


def is_holiday(year, month, day):
    """Check if a given date is a German holiday."""
    try:
        date = datetime(int(year), int(month), int(day))
        return date in german_holidays
    except (ValueError, TypeError):
        return False


def is_weekend(year, month, day):
    """Check if a given date is a weekend (Saturday or Sunday)."""
    try:
        date = datetime(int(year), int(month), int(day))
        # weekday() returns 5 for Saturday, 6 for Sunday
        return date.weekday() >= 5
    except (ValueError, TypeError):
        return False


def parse_date_range(
    date_string, year=None, month=None, skip_weekends=False, skip_holidays=False
):
    """
    Parse a date range string like "3-7,9,11-13" into a list of day numbers.
    Optionally filters out weekends and holidays.

    Args:
        date_string: String with day numbers and ranges (e.g., "3-7,9,11-13")
        year: Year as integer (required if skip_weekends or skip_holidays is True)
        month: Month as integer (required if skip_weekends or skip_holidays is True)
        skip_weekends: If True, exclude weekends from the result
        skip_holidays: If True, exclude German holidays from the result

    Returns:
        List of integers representing days, or None if invalid

    Examples:
        "3-7" -> [3, 4, 5, 6, 7]
        "1,3,5" -> [1, 3, 5]
        "3-7,9,11-13" -> [3, 4, 5, 6, 7, 9, 11, 12, 13]
        "1-7" with skip_weekends=True -> [1, 2, 3, 4, 5] (excludes Sat/Sun)
    """
    if not date_string or not date_string.strip():
        return None

    # Validate year/month if filtering is requested
    if (skip_weekends or skip_holidays) and (year is None or month is None):
        raise ValueError(
            "Year and month are required when skip_weekends or skip_holidays is True"
        )

    days = set()

    try:
        # Split by comma
        parts = date_string.split(",")

        for part in parts:
            part = part.strip()

            if "-" in part:
                # Range like "3-7"
                range_parts = part.split("-")
                if len(range_parts) != 2:
                    return None

                start = int(range_parts[0].strip())
                end = int(range_parts[1].strip())

                if start > end or start < 1 or end > 31:
                    return None

                days.update(range(start, end + 1))
            else:
                # Single day
                day = int(part)
                if day < 1 or day > 31:
                    return None
                days.add(day)

        # Filter out weekends and holidays if requested
        if skip_weekends or skip_holidays:
            filtered_days = set()
            for day in days:
                # Check if date is valid for the given month
                if not validate_date(year, month, day):
                    continue

                # Skip weekends if requested
                if skip_weekends and is_weekend(year, month, day):
                    continue

                # Skip holidays if requested
                if skip_holidays and is_holiday(year, month, day):
                    continue

                filtered_days.add(day)

            days = filtered_days

        return sorted(list(days))

    except (ValueError, AttributeError):
        return None


def validate_days_in_month(year, month, days):
    """
    Validate that all days in the list are valid for the given year and month.

    Args:
        year: Year as integer or string
        month: Month as integer or string (1-12)
        days: List of day integers

    Returns:
        Tuple of (is_valid: bool, invalid_days: list)
        - is_valid: True if all days are valid for the month
        - invalid_days: List of days that are invalid for the given month

    Examples:
        validate_days_in_month(2024, 2, [28, 29, 30]) -> (False, [30])
        validate_days_in_month(2024, 2, [28, 29]) -> (True, [])
    """
    invalid_days = []

    try:
        year_int = int(year)
        month_int = int(month)
    except (ValueError, TypeError):
        return (False, days)

    for day in days:
        if not validate_date(year_int, month_int, day):
            invalid_days.append(day)

    return (len(invalid_days) == 0, invalid_days)


def parse_multiple_names(names_string):
    """
    Parse a comma-separated list of names.

    Args:
        names_string: String with comma-separated names (e.g., "Max, Anna, Peter")

    Returns:
        List of strings with trimmed names

    Examples:
        "Max, Anna" -> ["Max", "Anna"]
        "Peter" -> ["Peter"]
    """
    if not names_string or not names_string.strip():
        return []

    # Split by comma and trim each name
    names = [name.strip() for name in names_string.split(",")]

    # Filter out empty strings
    names = [name for name in names if name]

    return names


def calculate_skug(year, month, day, hours_worked, skug_settings):
    """
    Calculate SKUG value based on date, hours worked, and settings.

    Args:
        year: Year as integer
        month: Month as integer (1-12)
        day: Day as integer
        hours_worked: Hours worked as float
        skug_settings: Dictionary with SKUG settings from database

    Returns:
        Float representing SKUG value, or 0.0 if invalid

    Logic:
        - Winter (December-March): Use winter settings
        - Summer (April-November): Use summer settings
        - SKUG = target_hours - hours_worked
    """
    try:
        # Get weekday (0=Monday, 6=Sunday)
        date = datetime(int(year), int(month), int(day))
        weekday = date.weekday()

        # Only calculate for Monday-Friday (0-4)
        if weekday > 4:
            return 0.0

        # Determine season (winter: Dec-Mar, summer: Apr-Nov)
        is_winter = month in [12, 1, 2, 3]

        # Map weekday to setting key
        weekday_map = {
            0: "monday",
            1: "tuesday",
            2: "wednesday",
            3: "thursday",
            4: "friday",
        }

        day_name = weekday_map[weekday]
        season = "winter" if is_winter else "summer"
        setting_key = f"{season}_{day_name}"

        # Get target hours for this day
        target_hours = float(skug_settings.get(setting_key, 8.0))

        # Calculate SKUG
        skug = target_hours - float(hours_worked)
        # Return SKUG, can be negative if worked more than target
        return round(skug, 2)

    except (ValueError, TypeError, KeyError):
        return 0.0


def get_effective_fahrzeit(
    master_db: MasterDataDatabase,
    worker_id: int,
    baustelle_id: int,
    default_fahrzeit: float,
) -> float:
    """
    Get the effective Fahrzeit for a worker at a baustelle, considering overrides.
    """
    if worker_id:
        override = master_db.get_override(worker_id, baustelle_id)
        if override and override["fahrzeit"] is not None:
            return float(override["fahrzeit"])
    return float(default_fahrzeit)


def get_effective_verpflegungsgeld(
    master_db: MasterDataDatabase,
    worker_id: int,
    baustelle_id: int,
    default_verpflegungsgeld: float,
) -> float:
    """
    Get the effective Verpflegungsgeld for a worker at a baustelle, considering overrides.
    """
    if worker_id:
        override = master_db.get_override(worker_id, baustelle_id)
        if override and override["verpflegungsgeld"] is not None:
            return float(override["verpflegungsgeld"])
    return float(default_verpflegungsgeld)


def get_fahrstunden_for_name(
    name, month, year, master_db: MasterDataDatabase, db: Database
):
    """
    Get the total Fahrstunden for a given name in a specific month and year.

    Args:
        name: Name of the person as string
        month: Month as integer (1-12)
        year: Year as integer

    Returns:
        Float representing total Fahrstunden for that person in the month
    """

    print("Getting Fahrstunden for name:", name, "month:", month, "year:", year)
    # Get all arbeitsstunden entries for the person in the specified month and year
    worker_id = master_db.get_worker_id_by_name(name)

    # Query arbeitsstunden table
    conn = db.db_file
    import sqlite3

    connection = sqlite3.connect(conn)
    connection.row_factory = sqlite3.Row
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT kostenstelle FROM arbeitsstunden
        WHERE jahr = ? AND monat = ? AND name = ?
    """,
        (year, month, name),
    )

    entries = cursor.fetchall()
    connection.close()

    total_fahrstunden = 0.0
    for entry in entries:
        kostenstelle = entry["kostenstelle"]
        if kostenstelle and kostenstelle not in ["Krank", "900", "940"]:
            # Extract baustelle number (format: "number - name")
            baustelle_nummer = (
                kostenstelle.split("-")[0].strip()
                if "-" in kostenstelle
                else kostenstelle
            )
            baustelle_data = master_db.get_baustelle_by_nummer(baustelle_nummer)
            if baustelle_data:
                fahrzeit = get_effective_fahrzeit(
                    master_db,
                    worker_id,
                    baustelle_data["id"],
                    baustelle_data.get("fahrzeit", 0.0),
                )
                total_fahrstunden += fahrzeit * 2  # round trip

    return round(total_fahrstunden, 2)


def get_skug_hours_for_name(name, month, year, db: Database):
    # Get all entries for the person in the specified month and year
    metadata = db.get_metadata_for_month(year, month, name)
    if not metadata:
        return 0.0
    skug = 0.0
    for entry in metadata:
        try:
            if float(entry["skug"]) < 1:
                continue
            skug += float(entry["skug"])
        except:
            pass
    return skug


def get_verpflegungsgeld_for_name(
    name, month, year, master_db: MasterDataDatabase, db: Database
):
    # Get all entries for the person in the specified month and year
    metadata = db.get_metadata_for_month(year, month, name)
    if not metadata:
        return 0.0

    worker_id = master_db.get_worker_id_by_name(name)

    total_verpflegungsgeld = 0.0

    for m_entry in metadata:
        travel_status = m_entry.get("travel_status")

        if travel_status:
            if travel_status == TravelStatus.Away24h:
                total_verpflegungsgeld += AWAY_24H_VERPFLEGUNG
            else:
                total_verpflegungsgeld += AN_ODER_ABREISE_VERPFLEGUNG
        elif m_entry.get("kg_8h"):
            continue
        else:
            arbeitsstunden_data = db.get_arbeitsstunden_for_day(
                year, month, m_entry.get("tag"), name
            )
            highest_verpflegungsgeld = 0.0
            for arbeits_entry in arbeitsstunden_data:
                kostenstelle = arbeits_entry.get(
                    "kostenstelle"
                )  # stunden not needed here
                if kostenstelle and kostenstelle not in ["940", "900", "Krank"]:
                    baustelle_id_str = kostenstelle.split("-")[0].strip()
                    baustelle = master_db.get_baustelle_by_nummer(baustelle_id_str)
                    if baustelle:
                        verpflegungsgeld = get_effective_verpflegungsgeld(
                            master_db,
                            worker_id,
                            baustelle["id"],
                            baustelle.get("verpflegungsgeld", 0.0),
                        )
                        if highest_verpflegungsgeld < verpflegungsgeld:
                            highest_verpflegungsgeld = verpflegungsgeld
            total_verpflegungsgeld += highest_verpflegungsgeld
    return round(total_verpflegungsgeld, 2)


def get_normal_hours_per_month(year, month, master_db: MasterDataDatabase):
    """
    Calculate the normal working hours for a given month based on SKUG settings.

    Args:
        year: Year as integer
        month: Month as integer (1-12)
        master_db: Instance of MasterDataDatabase to fetch SKUG settings
    Returns:
        Float representing total normal working hours for the month
    """

    skug_settings = master_db.get_skug_settings()
    total_hours = 0.0

    # Get number of days in the month
    if month == 12:
        next_month = 1
        next_year = year + 1
    else:
        next_month = month + 1
        next_year = year

    first_day = datetime(year, month, 1)
    if next_month == 1:
        first_day_next_month = datetime(next_year, next_month, 1)
    else:
        first_day_next_month = datetime(year, next_month, 1)

    num_days = (first_day_next_month - first_day).days

    for day in range(1, num_days + 1):
        date = datetime(year, month, day)
        weekday = date.weekday()

        # Only consider Monday to Friday
        if weekday <= 4:
            is_winter = month in [12, 1, 2, 3]
            weekday_map = {
                0: "monday",
                1: "tuesday",
                2: "wednesday",
                3: "thursday",
                4: "friday",
            }
            day_name = weekday_map[weekday]
            season = "winter" if is_winter else "summer"
            setting_key = f"{season}_{day_name}"

            target_hours = float(skug_settings.get(setting_key, 8.0))
            total_hours += target_hours

    return round(total_hours, 2)


def get_days_of_urlaub(name, month, year, db: Database):
    """Get the number of Urlaub days for a person in a specific month."""
    # Query arbeitsstunden table for entries with kostenstelle = '900' or '940'
    import sqlite3

    conn = sqlite3.connect(db.db_file)
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT COUNT(DISTINCT tag) FROM arbeitsstunden
        WHERE jahr = ? AND monat = ? AND name = ? AND kostenstelle IN ('900', '940')
    """,
        (year, month, name),
    )

    result = cursor.fetchone()
    conn.close()

    urlaub_days = result[0] if result else 0
    return urlaub_days


def get_hours_of_urlaub(name, month, year, db: Database):
    """Get the number of hours for a person in a specific month."""
    # Query arbeitsstunden table for entries with kostenstelle = '900' or '940'
    import sqlite3

    conn = sqlite3.connect(db.db_file)
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT SUM(stunden) FROM arbeitsstunden
        WHERE jahr = ? AND monat = ? AND name = ? AND kostenstelle IN ('900', '940')
    """,
        (year, month, name),
    )

    result = cursor.fetchone()
    conn.close()

    urlaub_hours = result[0] if result else 0
    if urlaub_hours == None:
        urlaub_hours = 0
    return urlaub_hours


def get_days_of_krank(name, month, year, db: Database):
    """Get the number of Krank days for a person in a specific month."""
    # Query arbeitsstunden table for entries with kostenstelle = 'Krank'
    import sqlite3

    conn = sqlite3.connect(db.db_file)
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT COUNT(DISTINCT tag) FROM arbeitsstunden
        WHERE jahr = ? AND monat = ? AND name = ? AND kostenstelle = 'Krank'
    """,
        (year, month, name),
    )

    result = cursor.fetchone()
    conn.close()

    krank_days = result[0] if result else 0
    return krank_days


def get_hours_of_krank(name, month, year, db: Database):
    """Get the number of hours for a person in a specific month."""
    # Query arbeitsstunden table for entries with kostenstelle = 'Krank'
    import sqlite3

    conn = sqlite3.connect(db.db_file)
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT SUM(stunden) FROM arbeitsstunden
        WHERE jahr = ? AND monat = ? AND name = ? AND kostenstelle = 'Krank'
    """,
        (year, month, name),
    )

    result = cursor.fetchone()
    conn.close()

    krank_hours = result[0] if result else 0
    if krank_hours == None:
        krank_hours = 0
    return krank_hours


def get_days_of_feiertag(name, month, year):
    num_days = calendar.monthrange(year, month)[1]
    days = 0
    for day in range(1, num_days + 1):
        if is_holiday(year, month, day) and not is_weekend(year, month, day):
            days += 1
    return days


def get_hours_of_feiertag(name, month, year, skug_settings, person_data):
    num_days = calendar.monthrange(year, month)[1]
    hours = 0
    weekly_hours = person_data.get("weekly_hours", 0.0)
    worker_type = person_data.get("worker_type", WorkerTypes.Fest)
    b_keine_feiertagsstunden = person_data.get("keine_feiertagssstunden", False)
    for day in range(1, num_days + 1):
        if is_holiday(year, month, day) and not is_weekend(year, month, day):
            if b_keine_feiertagsstunden:
                continue
            if worker_type == WorkerTypes.Fest:
                hours += weekly_hours / 5.0
            else:
                hours += calculate_skug(year, month, day, weekly_hours, skug_settings)

    return hours


def get_next_day(year, month, day):
    try:
        current_date = datetime(int(year), int(month), int(day))
        next_date = current_date + timedelta(days=1)
        return (next_date.year, next_date.month, next_date.day)
    except (ValueError, TypeError):
        # If invalid date, just increment day by 1
        return (year, month, day + 1)


def get_next_day_skip_weekend(year, month, day):
    try:
        current_date = datetime(int(year), int(month), int(day))
        next_date = current_date + timedelta(days=1)

        while next_date.weekday() >= 5:
            next_date += timedelta(days=1)

        return (next_date.year, next_date.month, next_date.day)
    except (ValueError, TypeError):
        return (year, month, day + 1)


def validate_required_fields(jahr, monat, name, stunden) -> tuple[bool, str]:
    if not jahr:
        return (False, "Jahr ist erforderlich!")

    if not monat:
        return (False, "Monat ist erforderlich!")

    if not name:
        return (False, "Name ist erforderlich!")
    try:
        jahr_int = int(jahr)
        monat_int = int(monat)

        if stunden:
            stunden_float = float(stunden)
            if not (0 <= stunden_float <= 24):
                return (False, "Stunden müssen zwischen 0 und 24 liegen!")

        if not (1900 <= jahr_int <= 2100):
            return (False, "Jahr muss zwischen 1900 und 2100 liegen!")

        if not (1 <= monat_int <= 12):
            return (False, "Monat muss zwischen 1 und 12 liegen!")

    except ValueError:
        return (False, "Jahr, Monat, Tag und Stunden müssen Zahlen sein!")

    return (True, "")


def handle_krank_urlaub(
    jahr_int,
    monat_int,
    day,
    name,
    db: Database,
    master_db: MasterDataDatabase,
    input_krank,
    input_urlaub,
    skug_settings,
):
    db.clear_entries_for_day(jahr_int, monat_int, day, name)
    wochentag = get_weekday_abbr(jahr_int, monat_int, str(day)) or ""
    final_urlaub_val = ""
    final_krank_val = ""
    kostenstelle = ""

    if input_krank:
        krank_value = calculate_skug(jahr_int, monat_int, day, 0, skug_settings)
        final_krank_val = str(krank_value) if krank_value != 0.0 else ""
        kostenstelle = "Krank"
    elif input_urlaub:
        urlaub_value = calculate_skug(jahr_int, monat_int, day, 0, skug_settings)
        final_urlaub_val = str(urlaub_value) if urlaub_value != 0.0 else ""
        worker_type = master_db.get_worker_type_by_name(name) or WorkerTypes.Fest
        kostenstelle = "900" if worker_type == WorkerTypes.Fest else "940"

    data = {
        "jahr": jahr_int,
        "monat": monat_int,
        "tag": str(day),
        "name": name,
        "wochentag": wochentag,
        "stunden": calculate_skug(jahr_int, monat_int, day, 0, skug_settings),
        "urlaub": final_urlaub_val,
        "krank": final_krank_val,
        "kg_8h": None,
        "skug": "",
        "kostenstelle": kostenstelle,
        "fruehstueck": False,
        "mittag": False,
        "travel_status": None,
        "urlaub": final_urlaub_val,
        "krank": final_krank_val,
    }

    db.add_arbeitsstunden(data)
    db.add_or_update_metadata(data)


def check_arbeitsstunden(entry_data):
    stunden = entry_data.get("stunden", None)
    baustelle = entry_data.get("kostenstelle", None)
    print("stunden:", stunden, " baustelle:", baustelle)
    if stunden is None or baustelle is None or not baustelle:
        return False
    return True


def try_load_existing_entry(
    jahr_int, monat_int, day, name, baustelle_input, db: Database
):
    existing_entries = db.get_arbeitsstunden_for_day(jahr_int, monat_int, day, name)
    target_entry_id = None
    entry_data = {}
    errors = []
    if baustelle_input:
        match = next(
            (
                e
                for e in existing_entries
                if ("kostenstelle" in e and e["kostenstelle"] == baustelle_input)
            ),
            None,
        )
        if match:
            target_entry_id = match["id"]
            entry_data = dict(match)
        else:
            for e in existing_entries:
                kostenstelle = e.get("kostenstelle", None)
                if kostenstelle in ["Krank", "940", "900"]:
                    db.delete_arbeitsstunden(e["id"])
            target_entry_id = None
            entry_data = {}
    else:
        if len(existing_entries) == 1:
            target_entry_id = existing_entries[0]["id"]
            entry_data = dict(existing_entries[0])
    return target_entry_id, entry_data, errors


def determine_kg_8h_flag(
    db: Database, master_db: MasterDataDatabase, jahr_int, monat_int, day, name
):
    day_entries = db.get_arbeitsstunden_for_day(jahr_int, monat_int, day, name)
    metadata_entry = db.get_metadata_by_date(jahr_int, monat_int, day, name)
    total_hours = 0.0
    highest_fahrzeit = 0.0
    for e in day_entries:
        h = float(e.get("stunden") or 0.0)
        bst_name = e.get("kostenstelle")
        if bst_name:
            bst_nummer = bst_name.split("-")[0].strip() if "-" in bst_name else bst_name
            bst_data = master_db.get_baustelle_by_nummer(bst_nummer)
            if bst_data:
                worker_id = master_db.get_worker_id_by_name(name)
                fahrzeit = get_effective_fahrzeit(
                    master_db, worker_id, bst_data["id"], bst_data.get("fahrzeit", 0.0)
                )
                if fahrzeit > highest_fahrzeit:
                    highest_fahrzeit = fahrzeit

        total_hours += h
    total_hours += highest_fahrzeit
    if metadata_entry.get("fruehstueck"):
        total_hours += 0.25
    if metadata_entry.get("mittag"):
        total_hours += 0.5
    is_unter_8h = total_hours <= 8.0
    if metadata_entry["travel_status"]:
        is_unter_8h = None
    return is_unter_8h
