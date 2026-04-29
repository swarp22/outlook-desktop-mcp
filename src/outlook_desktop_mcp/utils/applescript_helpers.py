"""Helpers for building and parsing AppleScript safely."""
import re
from datetime import datetime


def escape(text: str) -> str:
    """Escape a string for safe embedding inside AppleScript double quotes.

    Handles backslashes, double quotes, and other special characters.
    """
    text = text.replace("\\", "\\\\")
    text = text.replace('"', '\\"')
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "\\r")
    text = text.replace("\t", "\\t")
    return text


def format_date(dt: datetime) -> str:
    """Convert a Python datetime to an AppleScript date string.

    Returns a string like: date "Sunday, March 22, 2026 at 2:00:00 PM"
    AppleScript parses dates based on the system locale, so we use a
    locale-friendly format that osascript can interpret.
    """
    return f'date "{dt.strftime("%Y-%m-%d %H:%M:%S")}"'


def applescript_date_var(var_name: str, dt: datetime, indent: str = "    ") -> str:
    """Build an AppleScript snippet that assigns ``dt`` to ``var_name``.

    Constructs the date programmatically (numeric components) instead of via
    a ``date "..."`` literal. AppleScript parses date literals through the
    system locale, which mis-parses ISO strings like ``"2026-04-27 00:00:00"``
    in non-English locales (e.g. de_DE turns it into 16. Oktober 2032).
    Numeric assignment is locale-independent.

    Sets ``day`` to 1 first so a later month change cannot overflow (e.g. it
    avoids March when assigning month=2 to a date currently on the 31st).
    """
    return (
        f"{indent}set {var_name} to current date\n"
        f"{indent}set day of {var_name} to 1\n"
        f"{indent}set year of {var_name} to {dt.year}\n"
        f"{indent}set month of {var_name} to {dt.month}\n"
        f"{indent}set day of {var_name} to {dt.day}\n"
        f"{indent}set hours of {var_name} to {dt.hour}\n"
        f"{indent}set minutes of {var_name} to {dt.minute}\n"
        f"{indent}set seconds of {var_name} to {dt.second}\n"
    )


def parse_date(text: str) -> str:
    """Parse an AppleScript date string to ISO 8601 format.

    AppleScript dates look like: "Sunday, March 22, 2026 at 2:00:00 PM"
    or various locale-specific formats. We attempt several common patterns.
    """
    text = text.strip()
    # Remove day name prefix if present (e.g., "Sunday, ")
    text = re.sub(r"^\w+day,\s*", "", text)
    # Remove " at " between date and time
    text = text.replace(" at ", " ")
    # Try common formats
    for fmt in (
        "%B %d, %Y %I:%M:%S %p",   # March 22, 2026 2:00:00 PM
        "%d. %B %Y %H:%M:%S",       # 22. mars 2026 14:00:00 (Norwegian)
        "%Y-%m-%d %H:%M:%S",        # 2026-03-22 14:00:00
        "%d/%m/%Y %H:%M:%S",        # 22/03/2026 14:00:00
        "%m/%d/%Y %H:%M:%S",        # 03/22/2026 14:00:00
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.isoformat()
        except ValueError:
            continue
    # Fallback: return as-is
    return text


# Locale-independent AppleScript folder keywords
FOLDER_MAP = {
    "inbox": "inbox",
    "sent": "sent items",
    "sentmail": "sent items",
    "sent items": "sent items",
    "drafts": "drafts",
    "deleted": "deleted items",
    "deleted items": "deleted items",
    "trash": "deleted items",
    "junk": "junk mail",
    "spam": "junk mail",
    "outbox": "outbox",
}


def resolve_folder_ref(folder_name: str) -> str:
    """Map a user-facing folder name to an AppleScript folder reference.

    Returns an AppleScript expression like 'inbox' or 'mail folder "Archive"'.
    Built-in folders use locale-independent keywords; custom folders use name lookup.
    """
    key = folder_name.lower().strip()
    if key in FOLDER_MAP:
        return FOLDER_MAP[key]
    # Custom folder — search by name
    return f'mail folder "{escape(folder_name)}"'


# Delimiter used for structured AppleScript output
DELIM = "|||"
RECORD_DELIM = "==="
