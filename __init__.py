"""Generate WinCross banner files from a banner specification or banner plan."""

from .excel_spec import read_points, parse_width_overrides, SheetError
from .plan_reader import read_plan, looks_like_plan, summarise
from .logic_translate import translate, translate_all, Translation
from .reader import read_any, detect
from .generator import emit, DEFAULTS, normalise_logic, consistency_checks
from .width_spacing import validate, width_report, SpecError

__version__ = "0.2.0"
__all__ = [
    "read_points", "parse_width_overrides", "SheetError",
    "read_plan", "looks_like_plan", "summarise",
    "translate", "translate_all", "Translation",
    "read_any", "detect",
    "emit", "DEFAULTS", "normalise_logic", "consistency_checks",
    "validate", "width_report", "SpecError",
]
