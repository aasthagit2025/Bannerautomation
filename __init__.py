"""Generate WinCross banner files from a banner specification sheet."""

from .excel_spec import read_points, parse_width_overrides, SheetError
from .generator import emit, DEFAULTS, normalise_logic, consistency_checks
from .width_spacing import validate, width_report, SpecError

__version__ = "0.1.0"
__all__ = [
    "read_points", "parse_width_overrides", "SheetError",
    "emit", "DEFAULTS", "normalise_logic", "consistency_checks",
    "validate", "width_report", "SpecError",
]
