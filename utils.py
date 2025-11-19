from datetime import datetime
import locale
from database import Database
from master_data import MasterDataDatabase
import holidays

locale.setlocale(locale.LC_TIME, 'de_DE')

# Initialize German holidays
german_holidays = holidays.Germany(subdiv='SH')

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


def parse_date_range(date_string, year=None, month=None, skip_weekends=False, skip_holidays=False):
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
        raise ValueError("Year and month are required when skip_weekends or skip_holidays is True")

    days = set()

    try:
        # Split by comma
        parts = date_string.split(',')

        for part in parts:
            part = part.strip()

            if '-' in part:
                # Range like "3-7"
                range_parts = part.split('-')
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
    names = [name.strip() for name in names_string.split(',')]

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
            0: 'monday',
            1: 'tuesday',
            2: 'wednesday',
            3: 'thursday',
            4: 'friday'
        }

        day_name = weekday_map[weekday]
        season = 'winter' if is_winter else 'summer'
        setting_key = f"{season}_{day_name}"

        # Get target hours for this day
        target_hours = float(skug_settings.get(setting_key, 8.0))

        # Calculate SKUG
        skug = target_hours - float(hours_worked)

        # Return SKUG, can be negative if worked more than target
        return round(skug, 2)

    except (ValueError, TypeError, KeyError):
        return 0.0
    
def get_verpflegungsgeld_for_name(name, month, year, master_db:MasterDataDatabase, db:Database):
    """
    Get the total Verpflegungsgeld for a given name in a specific month and year.

    Args:
        name: Name of the person as string
        month: Month as integer (1-12)
        year: Year as integer

    Returns:
        Float representing total Verpflegungsgeld for that person in the month
    """

    print("Getting Verpflegungsgeld for name:", name, "month:", month, "year:", year)
    # Get all entries for the person in the specified month and year
    entries = db.get_entries_by_month_and_name(year, month, name)

    total_verpflegungsgeld = 0.0

    for entry in entries:
        baustelle_id = entry.get('baustelle').split('-')[0].strip() if entry.get('baustelle') else None
        print("Entry Baustelle ID:", baustelle_id)
        if baustelle_id:
            baustelle = master_db.get_baustelle_by_nummer(baustelle_id)
            print("Baustelle data:", baustelle)
            if baustelle:
                verpflegungsgeld = baustelle.get('verpflegungsgeld', 0.0)
                total_verpflegungsgeld += float(verpflegungsgeld)

    return round(total_verpflegungsgeld, 2)

def get_normal_hours_per_month(year, month, master_db:MasterDataDatabase):
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
                0: 'monday',
                1: 'tuesday',
                2: 'wednesday',
                3: 'thursday',
                4: 'friday'
            }
            day_name = weekday_map[weekday]
            season = 'winter' if is_winter else 'summer'
            setting_key = f"{season}_{day_name}"

            target_hours = float(skug_settings.get(setting_key, 8.0))
            total_hours += target_hours

    return round(total_hours, 2)