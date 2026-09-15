"""Assemble a WinCross banner file from banner points."""

from . import header_block as hb
from . import width_spacing as ws


DEFAULTS = {
    "banner_id": 1,
    "point_width": 70,
    "spaces_before": 1,
    "stub_width": 1,
    "column_divider": "",
    "banner_title": "",
    "banner_filter": "",
    "weights": "",
    "comparison_groups": "0,0",
    "stat_test": "^  ,0",
    "options": "1,SB,HD,W200",
    "extra": "",
    "emit_header_block": True,
    "normalise_logic": True,
}


def stat_letters(n):
    """A..Z, then A1..Z1, A2.. - the convention used in WinCross banner files."""
    out = []
    for i in range(n):
        letter = chr(ord("A") + i % 26)
        cycle = i // 26
        out.append(letter if cycle == 0 else f"{letter}{cycle}")
    return out


def normalise_logic(expr):
    """Tidy spacing without changing meaning: 'Q9 (2,7)' -> 'Q9(2,7)'."""
    out = " ".join(str(expr).split())
    for _ in range(3):
        out = out.replace(" (", "(").replace("( ", "(").replace(" )", ")")
    out = out.replace("{ ", "{").replace(" }", "}")
    # keep a space around the boolean operators
    for op in ("AND", "OR", "NOT"):
        out = out.replace(f"){op}", f") {op}").replace(f"{op}(", f"{op} (")
    return " ".join(out.split())


def consistency_checks(points):
    """Study-agnostic checks worth surfacing before a banner ships."""
    notes = []

    seen = {}
    for p in points:
        key = normalise_logic(p["logic"]).upper()
        seen.setdefault(key, []).append(p["column"])
    for logic, cols in seen.items():
        if len(cols) > 1:
            notes.append(
                f"columns {cols} share identical logic - they will always "
                f"report the same base and can never test significant against "
                f"each other")

    for p in points:
        if not p["logic"]:
            notes.append(f"column {p['column']} ({p['label']!r}) has no logic")
        if not p["label"]:
            notes.append(f"column {p['column']} has no label")

    return notes


def emit(points, settings=None):
    """Return (text, warnings, errors, stats)."""
    cfg = dict(DEFAULTS)
    cfg.update(settings or {})

    if cfg["normalise_logic"]:
        for p in points:
            p["logic"] = normalise_logic(p["logic"])

    spaces = int(cfg["spaces_before"])
    errors, warnings, stats = ws.validate(points, spaces, cfg["column_divider"])
    warnings = list(warnings) + consistency_checks(points)
    if errors:
        return "", warnings, errors, stats

    n = len(points)
    w = cfg["point_width"]
    lines = [
        "*Banner",
        f" ID:{cfg['banner_id']}",
        f" SW:{ws.sw_directive(points, spaces)}",
        " HP:" + ",".join(["1"] * n),
        f" CP:{cfg['comparison_groups']}",
        " SL:" + ",".join(stat_letters(n)),
        f" ST:{cfg['stat_test']}",
        f" WT:{cfg['weights']}",
        f" OP:{cfg['options']}",
        f" BT:{cfg['banner_title']}",
        f" BF:{cfg['banner_filter']}",
        f" XL:{cfg['extra']}",
        f" PT:{n},1",
    ]
    lines += [f" {p['logic']}^W{w}" for p in points]

    if cfg["emit_header_block"]:
        lines += hb.render(points, spaces, int(cfg["stub_width"]),
                           cfg.get("justification"))

    return "\n".join(lines) + "\n", warnings, [], stats
