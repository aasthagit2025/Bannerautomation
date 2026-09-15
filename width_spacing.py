#!/usr/bin/env python3
"""
Width and spacing handling for generated WinCross banners.

Encodes the constraints from Setup|Banners|Edit Banner > Width and Spacing:

  - "Spaces before each column" is a SINGLE value applied to ALL columns.
    The job file physically allows a different value per column, but editing
    the banner in the GUI afterwards will raise a warning. So we validate it.
  - Maximum spaces before each column is 5.
  - Default column width is 10 characters; width is set PER COLUMN.
  - Column divider characters are limited to the number of spaces before
    each column (1 space -> 1 divider char, 2 spaces -> 2 chars, etc).
  - Hidden columns do not count toward overall report width.
"""

MAX_SPACES = 5
DEFAULT_WIDTH = 10


class SpecError(Exception):
    pass


def label_lines(label):
    """Split a label into its rendered header lines.

    Two wrap conventions are supported, because both appear in practice:
      - a newline, which is how line breaks are stored in the Excel sheets
      - a double space, the convention used in plain-text banner files
    """
    if label is None:
        return []
    text = str(label)
    parts = text.split("\n") if "\n" in text else text.split("  ")
    return [seg.strip() for seg in parts if seg.strip()]


def longest_token(label):
    """Longest unbreakable run of characters in a label."""
    return max((len(tok) for line in label_lines(label) for tok in line.split()),
               default=0)


def validate(points, spaces_before, divider=""):
    """Return (errors, warnings, stats). Errors block generation."""
    errors, warnings = [], []

    if not isinstance(spaces_before, int) or spaces_before < 1:
        errors.append(f"spaces_before must be an integer >= 1 (got {spaces_before!r})")
    elif spaces_before > MAX_SPACES:
        errors.append(
            f"spaces_before is {spaces_before}; WinCross allows a maximum of {MAX_SPACES}")

    if divider and len(divider) > spaces_before:
        errors.append(
            f"column divider {divider!r} is {len(divider)} character(s) but "
            f"spaces_before is {spaces_before}; divider may not exceed it")

    for i, p in enumerate(points, start=1):
        w = p["width"]
        if w < 1:
            errors.append(f"col {i} ({p['label']!r}): width {w} is not valid")
        tok = longest_token(p["label"])
        if tok > w:
            warnings.append(
                f"col {i} ({p['label']!r}): longest word is {tok} chars but "
                f"column width is {w} - label will truncate")

    # Widths that differ between identical labels make parallel blocks misalign
    by_label = {}
    for i, p in enumerate(points, start=1):
        by_label.setdefault(p["label"], []).append((i, p["width"]))
    for lab, occ in by_label.items():
        if len({w for _, w in occ}) > 1:
            warnings.append(
                f"label {lab!r} appears with different widths {occ} - "
                f"parallel blocks will not align")

    visible = [p for p in points if p.get("visible", True)]
    hidden = len(points) - len(visible)
    stats = {
        "columns": len(points),
        "hidden": hidden,
        "report_width": sum(p["width"] + spaces_before for p in visible),
        "widths_used": sorted({p["width"] for p in points}),
    }
    return errors, warnings, stats


def sw_directive(points, spaces_before):
    """Build the SW string: one (spaces, width) pair per column."""
    return ",".join(f"{spaces_before},{p['width']}" for p in points)


def width_report(points, spaces_before):
    """Human-readable width map, for QC against the client's banner layout."""
    rows = ["  col  width  lines  label", "  ---  -----  -----  -----"]
    pos = 1
    for i, p in enumerate(points, start=1):
        n_lines = len(label_lines(p["label"]))
        rows.append(f"  {i:>3}  {p['width']:>5}  {n_lines:>5}  {p['label']}")
        pos += spaces_before + p["width"]
    rows.append(f"\n  total report width: {pos - 1} characters")
    return "\n".join(rows)
