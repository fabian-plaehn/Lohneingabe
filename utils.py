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