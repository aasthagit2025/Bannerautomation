#!/usr/bin/env python3
"""
Render the WinCross banner header block: the tiered rule lines and centred
label rows that sit beneath the logic lines in the banner file.

Derived from the structure of a known-good file and verified against its
rule spans. The layout is:

    <rule line>        dots spanning each super-header group
    <label line>       super-header labels, centred in their span
    <rule line>        dots spanning each column group
    <label line>       group labels, centred in their span
    <blank>
    <rule line>        dots spanning each individual column
    <label line>       column labels, centred
    <label line>...    continuation rows for wrapped labels

A rule spanning columns a..b is (sum of widths) + (number of internal
spacers) characters wide. Labels wrap on double-space, matching the
convention in the client's banner structure sheets.
"""

from .width_spacing import label_lines

RULE_CHAR = "."


def justify(text, width, how="center"):
    """Place text in a field of `width` characters.

    WinCross enters banner text LEFT-justified by default; centre and right
    are explicit choices (Cells|Center Justify / Right Justify). We default
    to centre because both client files use centred text throughout, but the
    spec can set it per tier or per column.
    """
    if len(text) >= width:
        return text[:width]
    pad = width - len(text)
    if how == "left":
        return text + " " * pad
    if how == "right":
        return " " * pad + text
    left = pad // 2                      # centre: bias left when uneven
    return " " * left + text + " " * (pad - left)


def centre(text, width):
    return justify(text, width, "center")


def spans(points, spaces_before, key):
    """Group consecutive columns sharing the same value of `key`.

    Returns [(label, start_index, end_index, span_width), ...]. Columns with
    an empty value form their own unlabelled span rather than merging into a
    neighbour - this matches how a standalone total column renders.
    """
    out = []
    i = 0
    while i < len(points):
        val = points[i].get(key, "") or ""
        j = i
        if val:
            while j + 1 < len(points) and (points[j + 1].get(key, "") or "") == val:
                j += 1
        width = sum(p["width"] for p in points[i:j + 1]) + (j - i) * spaces_before
        out.append((val, i, j, width))
        i = j + 1
    return out


def tier(points, spaces_before, key, stub, how="center"):
    """Return (rule_line, [label_lines]) for one header tier."""
    segs = spans(points, spaces_before, key)

    rule, labels = " " * stub, " " * stub
    depth = max((len(label_lines(s[0])) for s in segs), default=1)
    rows = [" " * stub for _ in range(depth)]

    for idx, (lab, _, _, width) in enumerate(segs):
        gap = " " * spaces_before if idx else ""
        rule += gap + RULE_CHAR * width
        wrapped = label_lines(lab)
        for r in range(depth):
            piece = wrapped[r] if r < len(wrapped) else ""
            rows[r] += gap + justify(piece, width, how)

    return rule, rows


def render(points, spaces_before=1, stub=1, justification=None):
    """Build the full header block as a list of text lines.

    `justification` maps tier name -> "left" | "center" | "right", e.g.
    {"super": "center", "group": "center", "column": "center"}.
    Anything unset falls back to centre.
    """
    just = {"super": "center", "group": "center", "column": "center"}
    just.update(justification or {})
    lines = []

    # Render one tier per heading row present in the sheet, outermost first.
    depth = max((len(p.get("tiers", [])) for p in points), default=0)
    for idx in range(depth):
        for p in points:
            tiers = p.get("tiers", [])
            p["_tier"] = tiers[idx] if idx < len(tiers) else ""
        name = "super" if idx == 0 else "group"
        rule, labels = tier(points, spaces_before, "_tier", stub,
                            just.get(name, "center"))
        lines.append(rule)
        lines.extend(labels)
    for p in points:
        p.pop("_tier", None)

    if depth:
        lines.append("")

    # Column tier: every column is its own span, so key on a unique index
    for i, p in enumerate(points):
        p["_col"] = f"{p['label']}\x00{i}"
    rule, _ = tier(points, spaces_before, "_col", stub)
    lines.append(rule)

    depth = max(len(label_lines(p["label"])) for p in points)
    for r in range(depth):
        row = " " * stub
        for idx, p in enumerate(points):
            wrapped = label_lines(p["label"])
            piece = wrapped[r] if r < len(wrapped) else ""
            how = p.get("justify", just["column"])
            row += (" " * spaces_before if idx else "") + justify(piece, p["width"], how)
        lines.append(row)

    for p in points:
        p.pop("_col", None)
    return lines
