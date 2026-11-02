"""Compare the two known SharedLogger version schemes without guessing across them."""
from __future__ import annotations

import re
from datetime import date

_DATED = re.compile(r"^V\.(\d{2})(\d{2})(\d{2})$", re.IGNORECASE)
_SEMANTIC = re.compile(r"^V?(\d+(?:\.\d+){0,7})(?:[-+].*)?$", re.IGNORECASE)


def version_sort_key(text: str) -> tuple[str, object]:
    """Dated V.DDMMYY and numeric Vx.y.z compare only within own family."""
    match = _DATED.fullmatch(text or '')
    if match:
        day, month, year = map(int, match.groups())
        try:
            return ('date', date(2000 + year, month, day).toordinal())
        except ValueError:
            return ('unknown', text)
    match = _SEMANTIC.fullmatch(text or '')
    if match:
        return ('semantic', tuple(int(x) for x in match.group(1).split('.')) + (0,) * (8 - len(match.group(1).split('.'))))
    return ('unknown', text)
