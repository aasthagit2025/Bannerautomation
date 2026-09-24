"""
WinCross Banner Generator - Streamlit app.

Upload a banner specification workbook, review the validation output,
and download a WinCross banner file.

Run locally:   streamlit run app.py
"""

import io
import re

import openpyxl
import streamlit as st


# ====================================================================
# width_spacing.py
# ====================================================================
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
        for line in label_lines(p["label"]):
            if len(line) > w:
                shown = line[:w]
                warnings.append(
                    f"col {i}: header line {line!r} is {len(line)} chars but "
                    f"column width is {w} - will render as {shown!r}")

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


# ====================================================================
# header_block.py
# ====================================================================
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


# ====================================================================
# logic_translate.py
# ====================================================================
"""
Translate banner-plan conditions into WinCross logic expressions.

Banner plans are written for humans: `S0=1`, `S8=1 OR 3`, `S4>14`. WinCross
wants `S0(1)`, `S8(1,3)`, `S4(15-9999)`. This module does that translation
and, just as importantly, refuses to guess when it cannot.

Every result carries a status:

    ok        translated with no assumptions
    assumed   translated, but a range bound had to be supplied
    blocked   cannot be translated; needs a human

`blocked` is the point of the module. A banner plan that still contains
`S5r4>XX` has an unfilled placeholder in it, and silently emitting something
plausible would put a wrong column into a deliverable.
"""

import re

# Bounds used when a comparison is open-ended. These are assumptions and are
# always reported as such.
DEFAULT_MIN = 0
DEFAULT_MAX = 9999

PLACEHOLDER = re.compile(r"\b(X{2,}|\?{2,}|TBD|TBC)\b", re.I)

# Text that describes a base rather than a condition
WHOLE_SAMPLE = {"all respondents", "all", "total", "everyone", "base"}


class Translation:
    def __init__(self, source, expr="", status="blocked", note=""):
        self.source = source
        self.expr = expr
        self.status = status
        self.note = note

    def __repr__(self):
        return f"<{self.status}: {self.source!r} -> {self.expr!r}>"


def _codes_from_equality(rhs):
    """'2,3,4 OR 5' -> [2, 3, 4, 5]. Returns None if anything is not numeric."""
    parts = re.split(r"\s*(?:,|\bOR\b|\bor\b|/)\s*", rhs)
    codes = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if not re.fullmatch(r"-?\d+", part):
            return None
        codes.append(int(part))
    return codes or None


def translate(condition, var_hint="", lo=DEFAULT_MIN, hi=DEFAULT_MAX):
    """Translate one condition string into a WinCross expression."""
    raw = "" if condition is None else str(condition).strip()
    if not raw:
        return Translation(raw, "", "blocked", "no condition given")

    # Normalise the unicode comparison operators that come out of Word/Excel
    text = (raw.replace("\u2264", "<=").replace("\u2265", ">=")
               .replace("\u2260", "<>").replace("\u2212", "-")
               .replace("\xa0", " "))
    text = " ".join(text.split())

    if text.lower() in WHOLE_SAMPLE:
        return Translation(
            raw, "", "blocked",
            "describes the whole sample rather than a condition - a total "
            "column needs its base defined explicitly")

    if PLACEHOLDER.search(text):
        return Translation(
            raw, "", "blocked",
            "contains an unfilled placeholder - the cut point has not been "
            "decided yet")

    # Already in WinCross form, e.g. S8r5(1) or Q1 (1) AND Q2(3)
    if re.search(r"\w\s*\([\d,\s\-]+\)", text):
        return Translation(raw, " ".join(text.split()), "ok",
                           "already in WinCross syntax")

    # VAR >= n / VAR > n / VAR <= n / VAR < n
    m = re.fullmatch(r"([A-Za-z_]\w*)\s*(<=|>=|<>|<|>)\s*(-?\d+)", text)
    if m:
        var, op, n = m.group(1), m.group(2), int(m.group(3))
        if op == ">":
            return Translation(raw, f"{var}({n+1}-{hi})", "assumed",
                               f"upper bound {hi} assumed for '{op}{n}'")
        if op == ">=":
            return Translation(raw, f"{var}({n}-{hi})", "assumed",
                               f"upper bound {hi} assumed for '{op}{n}'")
        if op == "<":
            return Translation(raw, f"{var}({lo}-{n-1})", "assumed",
                               f"lower bound {lo} assumed for '{op}{n}'")
        if op == "<=":
            return Translation(raw, f"{var}({lo}-{n})", "assumed",
                               f"lower bound {lo} assumed for '{op}{n}'")
        return Translation(raw, "", "blocked",
                           f"'{op}' has no direct WinCross equivalent")

    # A compound condition joins two comparisons: 'S0=1 AND S4>14'. This has
    # to be detected before the equality rule, which would otherwise swallow
    # the whole right-hand side. It is distinguished from a code list like
    # 'S0=2,3,4 OR 5' by checking that every part carries its own operator.
    parts = re.split(r"\s+(AND|OR)\s+", text, flags=re.I)
    if len(parts) > 1:
        operands = [p for i, p in enumerate(parts) if i % 2 == 0]
        if all(re.search(r"[<>=]", p) for p in operands):
            pieces, blocked, notes = [], [], []
            for i, part in enumerate(parts):
                if i % 2:
                    pieces.append(part.upper())
                    continue
                sub = translate(part.strip(), var_hint, lo, hi)
                if sub.status == "blocked":
                    blocked.append(sub.note)
                else:
                    if sub.status == "assumed":
                        notes.append(sub.note)
                    pieces.append(sub.expr)
            if blocked:
                return Translation(raw, "", "blocked", "; ".join(blocked))
            return Translation(raw, " ".join(pieces),
                               "assumed" if notes else "ok", "; ".join(notes))

    # VAR = codes, with commas and/or OR
    m = re.fullmatch(r"([A-Za-z_]\w*)\s*=\s*(.+)", text)
    if m:
        var, rhs = m.group(1), m.group(2)
        codes = _codes_from_equality(rhs)
        if codes is not None:
            body = ",".join(str(c) for c in codes)
            return Translation(raw, f"{var}({body})", "ok")
        return Translation(raw, "", "blocked",
                           f"right-hand side {rhs!r} is not a list of codes")

    return Translation(raw, "", "blocked", "condition syntax not recognised")


def translate_all(conditions, lo=DEFAULT_MIN, hi=DEFAULT_MAX):
    return [translate(c, lo=lo, hi=hi) for c in conditions]


# ====================================================================
# excel_spec.py
# ====================================================================
"""
Read a banner structure / banner spec workbook into a list of banner points.

The sheets used in practice share one grid layout:

      col A      | col B      | col C      | ...
    ------------------------------------------------
    (blank)      |                                     <- row 1, usually empty
    (blank)      | super-header, merged across columns  <- row 2
    (blank)      | group heading, merged across columns <- row 3
    (blank)      | column label                         <- row 4
    (blank)      | (sometimes blank)                    <- row 5
    (blank)      | banner logic                         <- last used row

Merged cell ranges carry the spans directly, so headings are read rather
than inferred. The logic row is located as the last row containing data,
which handles sheets with and without the blank spacer row.
"""

import openpyxl

DEFAULT_WIDTH = 10


class SheetError(Exception):
    pass


def _merge_map(ws, row):
    """Map column index -> the value governing it on `row`, following merges."""
    governing = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row:
            value = ws.cell(rng.min_row, rng.min_col).value
            for col in range(rng.min_col, rng.max_col + 1):
                governing[col] = value
    return governing


def _cell(ws, row, col, merges):
    value = merges.get(col, ws.cell(row, col).value)
    if value is None:
        return ""
    return str(value).strip()


def find_rows(ws):
    """Locate the logic row, the label row, and any heading rows above them.

    The logic row is the last populated row; the label row is the one above
    it. Everything populated above that is a heading tier, outermost first.
    Banners vary in how many tiers they carry - some have a super-header and
    a group heading, some only one heading, some none - so the count is read
    from the sheet rather than assumed.
    """
    populated = [
        r for r in range(1, ws.max_row + 1)
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1))
    ]
    if len(populated) < 2:
        raise SheetError(
            f"expected at least 2 populated rows (labels and logic); "
            f"found {len(populated)}")
    logic_row = populated[-1]
    label_row = populated[-2]
    tier_rows = populated[:-2]
    return tier_rows, label_row, logic_row


def read_points(path_or_buffer, sheet=None, default_width=DEFAULT_WIDTH,
                width_overrides=None):
    """Return (points, meta). Each point: super, group, label, logic, width."""
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]

    tier_rows, label_row, logic_row = find_rows(ws)
    merges = [_merge_map(ws, r) for r in tier_rows]

    width_overrides = width_overrides or {}
    points = []
    for col in range(1, ws.max_column + 1):
        label = _cell(ws, label_row, col, {})
        logic = _cell(ws, logic_row, col, {})
        if not label and not logic:
            continue                       # blank stub column on the left
        n = len(points) + 1
        point = {
            "label": label,
            "logic": " ".join(logic.split()),   # normalise internal whitespace
            "width": int(width_overrides.get(n, default_width)),
            "column": n,
            "tiers": [_cell(ws, r, col, m)
                      for r, m in zip(tier_rows, merges)],
        }
        # keep the two-tier names available for display and older callers
        point["super"] = point["tiers"][0] if len(point["tiers"]) > 0 else ""
        point["group"] = point["tiers"][1] if len(point["tiers"]) > 1 else ""
        points.append(point)

    if not points:
        raise SheetError("no banner columns found - check the sheet layout")

    meta = {
        "sheet": ws.title,
        "rows": {"headings": tier_rows, "label": label_row, "logic": logic_row},
        "tiers": len(tier_rows),
        "columns": len(points),
    }
    return points, meta


def parse_width_overrides(text):
    """Parse '1:20, 2:20, 6:20' into {1: 20, 2: 20, 6: 20}."""
    out = {}
    for chunk in (text or "").replace(";", ",").split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if ":" not in chunk:
            raise ValueError(f"expected 'column:width', got {chunk!r}")
        col, width = chunk.split(":", 1)
        out[int(col.strip())] = int(width.strip())
    return out


# ====================================================================
# plan_reader.py
# ====================================================================
"""
Read a banner plan laid out vertically, one row per banner column.

This is the second of the two shapes seen in practice. Where the banner
structure sheet runs horizontally with one column per banner point, a
banner plan runs downwards under named headings:

    Column | Variable | Group | Label | Response | Condition | N | ...
      1    | Total    |   1   |       | Total    | All resp. | 135
      2    | S0       |   2   | Geog. | US       | S0=1      |  75
      3    |          |       |       | EUR      | S0=2,3... |  60

One sheet can hold several banners, each introduced by a title row such as
"Banner 2: US (S0=1)" followed by its own heading row. Group headings come
from the Label column, which is merged down the rows it covers.

Conditions are written in plan syntax, not WinCross syntax, so they are put
through the translator and their status is carried on each point.
"""

import re


DEFAULT_WIDTH = 10

# Heading names, lowercased, mapped to the field they populate
HEADINGS = {
    "column": "column", "col": "column", "#": "column",
    "variable": "variable", "var": "variable",
    "group": "group_no", "grp": "group_no",
    "label": "group", "group label": "group", "heading": "group",
    "response": "label", "banner point": "label", "text": "label",
    "condition": "condition", "logic": "condition", "definition": "condition",
    "n": "n", "base": "n", "base size": "n",
}

TITLE = re.compile(r"^\s*banner\s*(\d+)?\s*[:\-]?\s*(.*)$", re.I)


def looks_like_plan(ws):
    """True when the sheet carries a banner-plan heading row."""
    for r in range(1, min(ws.max_row, 40) + 1):
        seen = {
            str(ws.cell(r, c).value).strip().lower()
            for c in range(1, min(ws.max_column, 15) + 1)
            if ws.cell(r, c).value not in (None, "")
        }
        if {"condition", "response"} <= seen or {"condition", "label"} <= seen:
            return True
    return False


def _resolved(ws, row, col):
    """Cell value, following a merged range to its top-left anchor."""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            v = ws.cell(rng.min_row, rng.min_col).value
            return "" if v is None else str(v).strip()
    v = ws.cell(row, col).value
    return "" if v is None else str(v).strip()


def _heading_row(ws, row):
    """Map column index -> field name, if `row` is a heading row."""
    mapping = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=row, column=c).value
        if v in (None, ""):
            continue
        key = str(v).strip().lower()
        if key in HEADINGS:
            mapping[c] = HEADINGS[key]
    return mapping if {"condition"} <= set(mapping.values()) else None


def read_plan(path_or_buffer, sheet=None, default_width=DEFAULT_WIDTH,
              lo=0, hi=9999):
    """Return {banner_name: [points]} plus a translation report."""
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True)
    names = [sheet] if sheet else wb.sheetnames
    banners, report = {}, []

    for name in names:
        ws = wb[name]
        if not looks_like_plan(ws):
            continue

        current, cols, title = None, None, None
        for r in range(1, ws.max_row + 1):
            first = _resolved(ws, r, 1)

            heading = _heading_row(ws, r)
            if heading:
                cols = heading
                if current is None:
                    title = title or f"{name}"
                    banners.setdefault(title, [])
                    current = title
                continue

            # A title row: text in column 1, nothing that looks like data
            if first and not cols:
                m = TITLE.match(first)
                if m:
                    title = first.strip()
                continue
            if first and cols and not re.fullmatch(r"\d+", first):
                m = TITLE.match(first)
                if m and "banner" in first.lower():
                    title = first.strip()
                    cols = None
                    current = None
                    continue

            if not cols:
                continue

            row_vals = {field: _resolved(ws, r, c) for c, field in cols.items()}
            if not row_vals.get("condition") and not row_vals.get("label"):
                continue

            if current is None:
                title = title or name
                banners.setdefault(title, [])
                current = title

            t = translate(row_vals.get("condition", ""), lo=lo, hi=hi)
            n = len(banners[current]) + 1
            point = {
                "column": n,
                "label": row_vals.get("label", ""),
                "logic": t.expr,
                "width": default_width,
                "tiers": [row_vals.get("group", "")],
                "super": row_vals.get("group", ""),
                "group": "",
                "variable": row_vals.get("variable", ""),
                "base_n": row_vals.get("n", ""),
                "condition": t.source,
                "status": t.status,
                "note": t.note,
            }
            banners[current].append(point)
            report.append({
                "banner": current, "column": n,
                "label": point["label"], "condition": t.source,
                "expression": t.expr, "status": t.status, "note": t.note,
            })

    if not banners:
        raise ValueError("no banner-plan tables found in this workbook")
    return banners, report


def summarise(report):
    counts = {"ok": 0, "assumed": 0, "blocked": 0}
    for row in report:
        counts[row["status"]] = counts.get(row["status"], 0) + 1
    return counts


# ====================================================================
# generator.py
# ====================================================================
"""Assemble a WinCross banner file from banner points."""



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
    errors, warnings, stats = validate(points, spaces, cfg["column_divider"])
    warnings = list(warnings) + consistency_checks(points)
    if errors:
        return "", warnings, errors, stats

    n = len(points)
    w = cfg["point_width"]
    lines = [
        "*Banner",
        f" ID:{cfg['banner_id']}",
        f" SW:{sw_directive(points, spaces)}",
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
        lines += render(points, spaces, int(cfg["stub_width"]),
                           cfg.get("justification"))

    return "\n".join(lines) + "\n", warnings, [], stats


# ====================================================================
# reader.py - format detection
# ====================================================================
def detect(path_or_buffer):
    """Return 'plan' or 'grid'."""
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True)
    for name in wb.sheetnames:
        if looks_like_plan(wb[name]):
            return "plan"
    return "grid"


def read_any(path_or_buffer, default_width=10, width_overrides=None,
             lo=0, hi=9999):
    """Read either banner format. Returns (banners, report, fmt)."""
    fmt = detect(path_or_buffer)
    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)

    if fmt == "plan":
        banners, report = read_plan(path_or_buffer,
                                    default_width=default_width, lo=lo, hi=hi)
        if width_overrides:
            for points in banners.values():
                for p in points:
                    if p["column"] in width_overrides:
                        p["width"] = int(width_overrides[p["column"]])
        return banners, report, fmt

    points, meta = read_points(path_or_buffer, default_width=default_width,
                               width_overrides=width_overrides)
    return {meta.get("sheet", "Banner"): points}, [], fmt



# ====================================================================
# Streamlit interface
# ====================================================================

st.set_page_config(page_title="WinCross Banner Generator",
                   page_icon="|", layout="wide")

st.title("WinCross Banner Generator")
st.caption(
    "Turn a banner specification sheet into a WinCross banner file: "
    "directives, logic lines and the header text block."
)

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Banner settings")

    banner_id = st.number_input("Banner ID", min_value=1, value=1, step=1)
    banner_title = st.text_input("Banner title (BT)", value="")
    banner_filter = st.text_input("Filter logic (BF)", value="")

    st.subheader("Width and spacing")
    st.caption(
        "WinCross requires the same 'spaces before' on every column. "
        "The maximum is 5. Default column width is 10."
    )
    spaces_before = st.slider("Spaces before each column", 1, 5, 1)
    default_width = st.number_input("Default column width", 1, 80, 10)
    width_text = st.text_input(
        "Width overrides", value="1:20, 2:20, 6:20",
        help="Column:width pairs, e.g. 1:20, 2:20, 6:20",
    )
    divider = st.text_input(
        "Column divider character(s)", value="", max_chars=5,
        help="Plain-text reports only. Cannot exceed 'spaces before each column'.",
    )

    st.subheader("Header block")
    emit_header = st.checkbox("Generate header text block", value=True)
    stub_width = st.number_input("Left stub width", 0, 20, 1)
    just = st.selectbox("Text justification", ["center", "left", "right"], index=0,
                        help="WinCross defaults to left; both sample banners are centred.")

    st.subheader("Open-ended ranges")
    st.caption(
        "Banner plans write conditions like 'S4>14'. WinCross needs both "
        "ends of a range, so the missing bound is supplied here."
    )
    range_lo = st.number_input("Assumed lower bound", -9999, 9999, 0)
    range_hi = st.number_input("Assumed upper bound", 1, 999999, 9999)

    st.subheader("Advanced")
    stat_test = st.text_input("Statistical testing (ST)", value="^  ,0")
    comparison = st.text_input("Comparison groups (CP)", value="0,0")
    weights = st.text_input("Weights (WT)", value="")
    options = st.text_input("Options (OP)", value="1,SB,HD,W200")
    point_width = st.number_input("Banner point width (^W)", 1, 200, 70)
    normalise = st.checkbox("Normalise logic spacing", value=True)


# ------------------------------------------------------------------ input
uploaded = st.file_uploader(
    "Banner specification workbook (.xlsx)", type=["xlsx", "xlsm"]
)

with st.expander("Expected sheet layout"):
    st.markdown(
        """
The reader finds the last four populated rows and treats them as:

| Row | Contents |
|---|---|
| super-header | merged across the columns it spans |
| group heading | merged across the columns it spans |
| column label | one per column |
| banner logic | one per column, in WinCross syntax |

Merged cell ranges define the header spans, so headings are read from the
sheet rather than guessed. Line breaks inside a label become separate
header lines. The leftmost blank column is ignored.
        """
    )

if not uploaded:
    st.info("Upload a specification workbook to begin.")
    st.stop()

try:
    overrides = parse_width_overrides(width_text)
except ValueError as exc:
    st.error(f"Could not read width overrides: {exc}")
    st.stop()

try:
    banners, report, fmt = read_any(
        uploaded, default_width=int(default_width), width_overrides=overrides,
        lo=int(range_lo), hi=int(range_hi),
    )
except SheetError as exc:
    st.error(f"Could not read the sheet: {exc}")
    st.stop()
except Exception as exc:  # noqa: BLE001 - surface any reader failure to the user
    st.error(f"Unexpected problem reading the workbook: {exc}")
    st.stop()

label = {"grid": "banner structure grid (one spreadsheet column per banner column)",
         "plan": "banner plan (one spreadsheet row per banner column)"}[fmt]
st.success(f"Detected a {label}. Found {len(banners)} banner(s).")

if len(banners) > 1:
    chosen = st.selectbox("Which banner?", list(banners))
else:
    chosen = list(banners)[0]
points = banners[chosen]

if report:
    rows = [r for r in report if r["banner"] == chosen]
    blocked = [r for r in rows if r["status"] == "blocked"]
    assumed = [r for r in rows if r["status"] == "assumed"]
    if blocked:
        st.error(
            f"{len(blocked)} column(s) could not be translated and are "
            f"excluded from the output. See the Translation tab.")
    if assumed:
        st.warning(
            f"{len(assumed)} column(s) needed a range bound to be assumed. "
            f"Check them in the Translation tab.")

skipped = [p for p in points if not p.get("logic")]
points = [p for p in points if p.get("logic")]
for i, p in enumerate(points, start=1):
    p["column"] = i
if skipped:
    st.info(
        f"{len(skipped)} column(s) have no usable logic and were left out: "
        + ", ".join(repr(p["label"]) for p in skipped[:6])
        + (" ..." if len(skipped) > 6 else "")
    )
if not points:
    st.error("No columns have usable logic - nothing to generate.")
    st.stop()

settings = {
    "banner_id": int(banner_id),
    "point_width": int(point_width),
    "spaces_before": int(spaces_before),
    "stub_width": int(stub_width),
    "column_divider": divider,
    "banner_title": banner_title,
    "banner_filter": banner_filter,
    "weights": weights,
    "comparison_groups": comparison,
    "stat_test": stat_test,
    "options": options,
    "emit_header_block": emit_header,
    "normalise_logic": normalise,
    "justification": {"super": just, "group": just, "column": just},
}

text, warnings, errors, stats = emit(points, settings)

# ----------------------------------------------------------------- output
c1, c2, c3, c4 = st.columns(4)
c1.metric("Columns", stats["columns"])
c2.metric("Report width", f"{stats['report_width']} chars")
c3.metric("Widths used", ", ".join(str(w) for w in stats["widths_used"]))
c4.metric("Warnings", len(warnings))

if errors:
    st.error("Generation blocked - fix these first:")
    for msg in errors:
        st.write(f"- {msg}")
    st.stop()

if warnings:
    with st.expander(f"{len(warnings)} warning(s) - review before shipping", expanded=True):
        for msg in warnings:
            st.warning(msg)

names = ["Banner file", "Column map", "Header preview"]
if report:
    names.append("Translation")
tabs = st.tabs(names)
tab_file, tab_cols, tab_header = tabs[0], tabs[1], tabs[2]
tab_trans = tabs[3] if report else None

with tab_file:
    st.download_button(
        "Download banner file", data=text.encode("utf-8"),
        file_name=f"{chosen.replace(':', '').replace(' ', '_')}.txt",
        mime="text/plain",
    )
    st.code(text, language="text")

with tab_cols:
    st.dataframe(
        [
            {
                "Col": p["column"],
                "Width": p["width"],
                "Label": p["label"].replace("\n", " / "),
                "Group": p["group"].replace("\n", " "),
                "Super": p["super"].replace("\n", " "),
                "Logic": p["logic"],
            }
            for p in points
        ],
        use_container_width=True, hide_index=True,
    )

with tab_header:
    if not emit_header:
        st.info("Header block generation is switched off in the sidebar.")
    else:
        block = text.rstrip("\n").split("\n")[13 + len(points):]
        st.caption(
            "Rule lines span each merged heading: the summed column widths "
            "plus the spacers between them. Scroll horizontally to inspect."
        )
        st.code("\n".join(block), language="text")


if tab_trans is not None:
    with tab_trans:
        st.caption(
            "How each plan condition was turned into WinCross logic. "
            "'blocked' rows are excluded from the generated file."
        )
        st.dataframe(
            [
                {
                    "Col": r["column"],
                    "Status": r["status"],
                    "Label": r["label"],
                    "Condition": r["condition"],
                    "WinCross": r["expression"] or "-",
                    "Note": r["note"],
                }
                for r in report if r["banner"] == chosen
            ],
            use_container_width=True, hide_index=True,
        )
