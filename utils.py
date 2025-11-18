from datetime import datetime
import locale

locale.setlocale(locale.LC_TIME, 'de_DE')

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


def parse_date_range(date_string):
    """
    Parse a date range string like "3-7,9,11-13" into a list of day numbers.

    Args:
        date_string: String with day numbers and ranges (e.g., "3-7,9,11-13")

    Returns:
        List of integers representing days, or None if invalid

    Examples:
        "3-7" -> [3, 4, 5, 6, 7]
        "1,3,5" -> [1, 3, 5]
        "3-7,9,11-13" -> [3, 4, 5, 6, 7, 9, 11, 12, 13]
    """
    if not date_string or not date_string.strip():
        return None

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