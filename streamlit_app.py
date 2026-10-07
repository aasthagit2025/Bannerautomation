#
# WinCross Banner Generator - Streamlit app.
#
# Upload a banner specification, pick the client, download a WinCross banner
# file. Everything the tool wants to tell you about the spec is collected in
# one Checks tab rather than stacked above the output.
#
# Run locally:   streamlit run app.py
#

import io
import re
import zipfile
import warnings

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import streamlit as st


# ====================================================================
# safe_load.py
# ====================================================================
#
# Open a workbook that openpyxl would otherwise refuse.
#
# Files written by some tools carry stylesheet attributes openpyxl does not
# recognise, and loading them raises a TypeError before any data is read:
#
#     CellStyle.__init__() got an unexpected keyword argument 'applyColorFormat'
#
# The data is fine; only the styling is unusual. This retries by stripping the
# offending attributes from a copy of the file. Nothing the generator does
# depends on cell styling, so nothing is lost.
#

import io
import re
import zipfile
import warnings

import openpyxl

# Attributes seen in the wild that openpyxl's CellStyle does not accept
BAD_ATTRS = re.compile(
    r'\s(?:applyColorFormat|applyNumberFormat2|applyBorderFormat|'
    r'applyPatternFormat)="[^"]*"'
)


def _sanitise(source):
    """Return a BytesIO of the workbook with unsupported style attrs removed."""
    if hasattr(source, "seek"):
        source.seek(0)
        data = source.read()
        src = io.BytesIO(data)
    else:
        src = source

    out = io.BytesIO()
    with zipfile.ZipFile(src) as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    text = payload.decode("utf-8", errors="replace")
                    payload = BAD_ATTRS.sub("", text).encode("utf-8")
                zout.writestr(item, payload)
    out.seek(0)
    return out


def load(source, **kwargs):
    """Load a workbook, repairing the stylesheet if openpyxl rejects it."""
    kwargs.setdefault("data_only", True)
    try:
        if hasattr(source, "seek"):
            source.seek(0)
        return openpyxl.load_workbook(source, **kwargs)
    except TypeError as exc:
        if "unexpected keyword argument" not in str(exc):
            raise
    except Exception:
        raise

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(_sanitise(source), **kwargs)


# ====================================================================
# width_spacing.py
# ====================================================================
#
# Width and spacing handling for generated WinCross banners.
#
# Encodes the constraints from Setup|Banners|Edit Banner > Width and Spacing:
#
#   - "Spaces before each column" is a SINGLE value applied to ALL columns.
#     The job file physically allows a different value per column, but editing
#     the banner in the GUI afterwards will raise a warning. So we validate it.
#   - Maximum spaces before each column is 5.
#   - Default column width is 10 characters; width is set PER COLUMN.
#   - Column divider characters are limited to the number of spaces before
#     each column (1 space -> 1 divider char, 2 spaces -> 2 chars, etc).
#   - Hidden columns do not count toward overall report width.
#

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


def wrap_to(label, width):
    """Split a label into header lines that fit `width`.

    Explicit breaks in the label are honoured first; each resulting line is
    then word-wrapped so nothing is cut off. A single word longer than the
    column is left alone and reported by validate() - hyphenating it would
    be worse than a visible warning.
    """
    out = []
    for line in label_lines(label):
        words, current = line.split(), ""
        for word in words:
            trial = f"{current} {word}".strip()
            if len(trial) <= width or not current:
                current = trial
            else:
                out.append(current)
                current = word
        if current:
            out.append(current)
    return out or [""]


def longest_token(label):
    """Longest unbreakable run of characters in a label."""
    return max((len(tok) for line in label_lines(label) for tok in line.split()),
               default=0)


def validate(points, spaces_before, divider="", wrap=True):
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
        if wrap:
            tok = longest_token(p["label"])
            if tok > w:
                warnings.append(
                    f"col {i}: the word {tok} characters long in {p['label']!r} "
                    f"cannot fit a {w}-character column and will be cut")
        else:
            for line in label_lines(p["label"]):
                if len(line) > w:
                    warnings.append(
                        f"col {i}: header line {line!r} is {len(line)} chars but "
                        f"column width is {w} - will render as {line[:w]!r}")

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
#
# Render the WinCross banner header block: the tiered rule lines and centred
# label rows that sit beneath the logic lines in the banner file.
#
# Derived from the structure of a known-good file and verified against its
# rule spans. The layout is:
#
#     <rule line>        dots spanning each super-header group
#     <label line>       super-header labels, centred in their span
#     <rule line>        dots spanning each column group
#     <label line>       group labels, centred in their span
#     <blank>
#     <rule line>        dots spanning each individual column
#     <label line>       column labels, centred
#     <label line>...    continuation rows for wrapped labels
#
# A rule spanning columns a..b is (sum of widths) + (number of internal
# spacers) characters wide. Labels wrap on double-space, matching the
# convention in the client's banner structure sheets.
#


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


def tier(points, spaces_before, key, stub, how="center", wrap=True):
    """Return (rule_line, [label_lines]) for one header tier.

    Heading labels wrap within their span exactly as column labels do.
    Without this a heading wider than the columns it covers is simply cut -
    "Q2: Pack Vs. Stick buying" over a 21-character span becomes
    "Q2: Pack Vs. Stick bu", which is how a group loses its name.
    """
    segs = spans(points, spaces_before, key)

    wrapped_all = [wrap_to(lab, width) if wrap else label_lines(lab)
                   for lab, _, _, width in segs]
    rule = " " * stub
    depth = max((len(w) for w in wrapped_all), default=1)
    rows = [" " * stub for _ in range(depth)]

    for idx, (lab, _, _, width) in enumerate(segs):
        gap = " " * spaces_before
        rule += gap + RULE_CHAR * width
        wrapped = wrapped_all[idx]
        for r in range(depth):
            piece = wrapped[r] if r < len(wrapped) else ""
            rows[r] += gap + justify(piece, width, how)

    return rule, rows


def render(points, spaces_before=1, stub=1, justification=None, wrap=True):
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
                            just.get(name, "center"), wrap)
        lines.append(rule)
        lines.extend(labels)
    for p in points:
        p.pop("_tier", None)

    # WinCross writes no blank between the heading tiers and the column
    # labels, and closes the block with one full-width line of spaces.

    # Column tier: every column is its own span, so key on a unique index
    for i, p in enumerate(points):
        p["_col"] = f"{p['label']}\x00{i}"
    rule, _ = tier(points, spaces_before, "_col", stub, wrap=False)
    lines.append(rule)

    wrapped_all = [wrap_to(p["label"], p["width"]) if wrap
                   else label_lines(p["label"]) for p in points]
    depth = max(len(w) for w in wrapped_all)
    for r in range(depth):
        row = " " * stub
        for idx, p in enumerate(points):
            wrapped = wrapped_all[idx]
            piece = wrapped[r] if r < len(wrapped) else ""
            how = p.get("justify", just["column"])
            row += " " * spaces_before + justify(piece, p["width"], how)
        lines.append(row)

    for p in points:
        p.pop("_col", None)

    width = spaces_before * len(points) + sum(p["width"] for p in points) + stub
    lines.append(" " * width)
    return lines


# ====================================================================
# logic_translate.py
# ====================================================================
#
# Translate banner-plan conditions into WinCross logic expressions.
#
# Banner plans are written for humans: `S0=1`, `S8=1 OR 3`, `S4>14`. WinCross
# wants `S0(1)`, `S8(1,3)`, `S4(15-9999)`. This module does that translation
# and, just as importantly, refuses to guess when it cannot.
#
# Every result carries a status:
#
#     ok        translated with no assumptions
#     assumed   translated, but a range bound had to be supplied
#     blocked   cannot be translated; needs a human
#
# `blocked` is the point of the module. A banner plan that still contains
# `S5r4>XX` has an unfilled placeholder in it, and silently emitting something
# plausible would put a wrong column into a deliverable.
#

import re

# Bounds used when a comparison is open-ended. These are assumptions and are
# always reported as such.
DEFAULT_MIN = 0
DEFAULT_MAX = 9999

PLACEHOLDER = re.compile(r"\b(X{2,}|\?{2,}|TBD|TBC)\b", re.I)


def resolve_placeholders(text, values):
    """Substitute decided cut points into a condition.

    `values` maps a variable name to its cut point, e.g. {"S5r4": 10}. A
    condition like 'S5r4>XX' becomes 'S5r4>10'. Matching is on the variable
    at the start of the condition, so the same placeholder can take a
    different value for each variable - which is the usual case, since the
    mitral and tricuspid thresholds are set independently.

    A key of "*" supplies a fallback for any variable not named.
    """
    if not values or not PLACEHOLDER.search(text):
        return text
    m = re.match(r"\s*([A-Za-z_]\w*)", text)
    if not m:
        return text
    var = m.group(1)
    for key in (var, var.lower(), var.upper(), "*"):
        if key in values and values[key] not in (None, ""):
            return PLACEHOLDER.sub(str(values[key]), text)
    return text

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
    """'2,3,4 OR 5' -> '2,3,4,5'; '1-4' -> '1-4'; '1-4, 7' -> '1-4,7'.

    Ranges are kept as ranges because WinCross writes them that way, and a
    spec that says 1-13 means a span rather than two codes.
    """
    parts = re.split(r"\s*(?:,|\bOR\b|\bor\b|/)\s*", rhs)
    out = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if re.fullmatch(r"-?\d+", part):
            out.append(part)
        elif re.fullmatch(r"\d+\s*-\s*\d+", part):
            out.append(re.sub(r"\s*-\s*", "-", part))
        else:
            return None
    return ",".join(out) or None


def translate(condition, var_hint="", lo=DEFAULT_MIN, hi=DEFAULT_MAX,
              placeholders=None):
    """Translate one condition string into a WinCross expression."""
    raw = "" if condition is None else str(condition).strip()
    if not raw:
        return Translation(raw, "", "blocked", "no condition given")

    # Normalise the unicode comparison operators that come out of Word/Excel
    text = (raw.replace("\u2264", "<=").replace("\u2265", ">=")
               .replace("\u2260", "<>").replace("\u2212", "-")
               .replace("\xa0", " "))
    text = " ".join(text.split())

    resolved = resolve_placeholders(text, placeholders or {})
    substituted = resolved != text
    text = resolved

    if text.lower() in WHOLE_SAMPLE:
        return Translation(
            raw, "", "blocked",
            "describes the whole sample rather than a condition - a total "
            "column needs its base defined explicitly")

    # A condition wrapped in a description: 'Total Customer Sample
    # (Sample_Type=2)'. Specs often name the base in words and put the real
    # condition in brackets after it. Take the bracketed part when it parses
    # on its own; the surrounding prose is a label, not logic.
    for inner in re.findall(r"\(([^()]+)\)", text):
        inner = inner.strip()
        if not inner or not re.search(r"[=<>]", inner):
            continue
        sub = translate(inner, var_hint, lo, hi, placeholders)
        if sub.status != "blocked":
            extra = f"{sub.note}; " if sub.note else ""
            return Translation(
                raw, sub.expr, sub.status,
                f"{extra}condition taken from {inner!r} inside the description")

    # A sum across the rows of a battery: 'S4r1 + r2 + r3 <= 10'. The
    # shorthand means S4r1 + S4r2 + S4r3. WinCross has no inline arithmetic,
    # so this needs a constructed variable - but the intent is clear enough
    # to say exactly what is needed.
    if "+" in text and re.search(r"(<=|>=|<|>|=)", text):
        head, op, rhs = re.split(r"(<=|>=|<|>|=)", text, maxsplit=1)[:3]
        terms = [t.strip() for t in head.split("+") if t.strip()]
        if len(terms) > 1:
            stem = re.match(r"([A-Za-z_]+\d*[A-Za-z_]*?)(?:r?\d+)?$", terms[0])
            prefix = stem.group(1) if stem else ""
            full = [terms[0]] + [
                t if re.match(r"[A-Za-z_]{2,}", t) else f"{prefix}{t}"
                for t in terms[1:]
            ]
            return Translation(
                raw, "", "blocked",
                f"this sums {' + '.join(full)} and compares the total to "
                f"{rhs.strip()}. WinCross has no inline arithmetic - build a "
                f"constructed variable for the sum, then condition on that")

    if PLACEHOLDER.search(text):
        return Translation(
            raw, "", "blocked",
            "contains an unfilled placeholder - supply the cut point for this "
            "variable to generate this column")

    # Already in WinCross form, e.g. S8r5(1) or Q1 (1) AND Q2(3)
    if re.search(r"\w\s*\([\d,\s\-]+\)", text):
        return Translation(raw, " ".join(text.split()), "ok",
                           "already in WinCross syntax")

    # VAR >= n / VAR > n / VAR <= n / VAR < n
    m = re.fullmatch(r"([A-Za-z_]\w*)\s*(<=|>=|<>|<|>)\s*(-?\d+)", text)
    if m:
        var, op, n = m.group(1), m.group(2), int(m.group(3))
        pre = "cut point supplied; " if substituted else ""
        if op == ">":
            return Translation(raw, f"{var}({n+1}-{hi})", "assumed",
                               f"{pre}upper bound {hi} assumed for '{op}{n}'")
        if op == ">=":
            return Translation(raw, f"{var}({n}-{hi})", "assumed",
                               f"{pre}upper bound {hi} assumed for '{op}{n}'")
        if op == "<":
            return Translation(raw, f"{var}({lo}-{n-1})", "assumed",
                               f"{pre}lower bound {lo} assumed for '{op}{n}'")
        if op == "<=":
            return Translation(raw, f"{var}({lo}-{n})", "assumed",
                               f"{pre}lower bound {lo} assumed for '{op}{n}'")
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
                sub = translate(part.strip(), var_hint, lo, hi, placeholders)
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
            return Translation(raw, f"{var}({codes})", "ok")
        return Translation(raw, "", "blocked",
                           f"right-hand side {rhs!r} is not a list of codes")

    # A hyphen or colon where an equals sign was meant: 'S0-1' for 'S0=1'.
    # Common enough to be worth naming rather than reporting as unparseable.
    m = re.fullmatch(r"([A-Za-z_]\w*)\s*[-\u2013]\s*(\d+(?:\s*,\s*\d+)*)", text)
    if m:
        var, codes = m.group(1), m.group(2)
        guess = f"{var}({','.join(c.strip() for c in codes.split(','))})"
        return Translation(
            raw, "", "blocked",
            f"looks like a typo - a hyphen where '=' was meant. If it should "
            f"be {var}={codes}, the logic is {guess}")

    # 'D2a: Prefers home depot' names the variable but describes the cut in
    # words. Say which variable it is - that is the useful half of the answer.
    m = re.match(r"([A-Za-z_]\w*)\s*[:\-]\s*(.+)", text)
    if m:
        return Translation(
            raw, "", "blocked",
            f"variable {m.group(1)} is named but the condition is written as "
            f"prose - the codes are needed, e.g. {m.group(1)}(1) or {m.group(1)}=1")

    return Translation(raw, "", "blocked", "condition syntax not recognised")


def translate_all(conditions, lo=DEFAULT_MIN, hi=DEFAULT_MAX, placeholders=None):
    return [translate(c, lo=lo, hi=hi, placeholders=placeholders)
            for c in conditions]


def parse_placeholders(text):
    """Parse 'S5r4:10, S5r6:5' into {'S5r4': 10, 'S5r6': 5}."""
    out = {}
    for chunk in (text or "").replace(";", ",").split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if ":" not in chunk and "=" not in chunk:
            raise ValueError(f"expected 'variable:value', got {chunk!r}")
        sep = ":" if ":" in chunk else "="
        var, val = chunk.split(sep, 1)
        out[var.strip()] = val.strip()
    return out


# ====================================================================
# excel_spec.py
# ====================================================================
#
# Read a banner structure / banner spec workbook into a list of banner points.
#
# The sheets used in practice share one grid layout:
#
#       col A      | col B      | col C      | ...
#     ------------------------------------------------
#     (blank)      |                                     <- row 1, usually empty
#     (blank)      | super-header, merged across columns  <- row 2
#     (blank)      | group heading, merged across columns <- row 3
#     (blank)      | column label                         <- row 4
#     (blank)      | (sometimes blank)                    <- row 5
#     (blank)      | banner logic                         <- last used row
#
# Merged cell ranges carry the spans directly, so headings are read rather
# than inferred. The logic row is located as the last row containing data,
# which handles sheets with and without the blank spacer row.
#

import re

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


LOGIC_HINT = re.compile(r"[=<>]|\w\s*\([\d,\s\-]+\)|\b(AND|OR|NOT)\b", re.I)


def looks_like_logic(ws, row, cols):
    """True when a row carries conditions rather than labels.

    A banner sheet sometimes arrives as layout only - headings and labels,
    no conditions - and the last two populated rows are then both label
    rows. Treating the labels as logic would emit nonsense, so the bottom
    row is checked for the marks of a condition.
    """
    values = [ws.cell(row, c).value for c in cols]
    values = [str(v).strip() for v in values if v not in (None, "")]
    if not values:
        return False
    hits = sum(1 for v in values if LOGIC_HINT.search(v))
    return hits >= max(1, len(values) // 2)


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
    cols = range(1, ws.max_column + 1)
    if not looks_like_logic(ws, populated[-1], cols):
        # Layout only: the bottom row is labels, and there is no logic.
        return populated[:-1], populated[-1], None

    logic_row = populated[-1]
    label_row = populated[-2]
    tier_rows = populated[:-2]
    return tier_rows, label_row, logic_row


def read_points(path_or_buffer, sheet=None, default_width=DEFAULT_WIDTH,
                width_overrides=None):
    """Return (points, meta). Each point: super, group, label, logic, width."""
    wb = load(path_or_buffer)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]

    tier_rows, label_row, logic_row = find_rows(ws)
    merges = [_merge_map(ws, r) for r in tier_rows]

    width_overrides = width_overrides or {}
    points = []
    for col in range(1, ws.max_column + 1):
        label = _cell(ws, label_row, col, {})
        logic = _cell(ws, logic_row, col, {}) if logic_row else ""
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
        "layout_only": logic_row is None,
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
#
# Read a banner plan laid out vertically, one row per banner column.
#
# This is the second of the two shapes seen in practice. Where the banner
# structure sheet runs horizontally with one column per banner point, a
# banner plan runs downwards under named headings:
#
#     Column | Variable | Group | Label | Response | Condition | N | ...
#       1    | Total    |   1   |       | Total    | All resp. | 135
#       2    | S0       |   2   | Geog. | US       | S0=1      |  75
#       3    |          |       |       | EUR      | S0=2,3... |  60
#
# One sheet can hold several banners, each introduced by a title row such as
# "Banner 2: US (S0=1)" followed by its own heading row. Group headings come
# from the Label column, which is merged down the rows it covers.
#
# Conditions are written in plan syntax, not WinCross syntax, so they are put
# through the translator and their status is carried on each point.
#

import re

import openpyxl



DEFAULT_WIDTH = 10

# Heading names, lowercased, mapped to the field they populate
HEADINGS = {
    "column": "column", "col": "column", "#": "column",
    "variable": "variable", "var": "variable",
    "group": "group_no", "grp": "group_no",
    "label": "group", "group label": "group", "heading": "group",
    "response": "label", "banner point": "label", "text": "label",
    "banner points": "label", "point": "label",
    "header": "group", "headers": "group",
    "condition": "condition", "logic": "condition", "definition": "condition",
    "definitions": "condition", "base": "condition",
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
        mapped = {HEADINGS[k] for k in seen if k in HEADINGS}
        if "condition" in mapped and "label" in mapped:
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
              lo=0, hi=9999, placeholders=None, total_logic=""):
    """Return {banner_name: [points]} plus a translation report."""
    wb = load(path_or_buffer)
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

            t = translate(row_vals.get("condition", ""), lo=lo, hi=hi,
                          placeholders=placeholders)
            # A total column is described rather than conditioned. If the house
            # convention for it has been given, use that.
            if t.status == "blocked" and total_logic and \
                    str(row_vals.get("condition", "")).strip().lower() in WHOLE_SAMPLE:
                t = Translation(t.source, total_logic, "assumed",
                                "total column base supplied")
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
# vertical_reader.py
# ====================================================================
#
# Read a vertical banner sheet that carries no header row.
#
# Two of the shapes seen in practice run downwards but have no
# Label/Response/Condition heading to key on:
#
#   separate group column        group headings sit in their own column,
#                                merged down the rows they cover
#
#       Store Preference | Those who prefer home depot | D2a: ...
#                        | Those who prefer competitor | D2a: ...
#       Income           | Less than $500k             | D3 = 1,2,3
#
#   group heading rows           a row with a label but no condition opens a
#                                new group; the rows beneath it are its columns
#
#       Q2: Pack Vs. Stick buying
#       Pack buyers                                    | Q2=2
#       Loose stick buyers                             | Q2=1
#
# Both are handled by locating the condition column first - it is the one
# whose cells look like conditions - and then reading the rest relative to it.
# Blank separator rows are ignored, and a row whose condition is prose rather
# than logic is carried through so the translator can report it.
#

import re


LOGIC_HINT = re.compile(r"[=<>]|\w\s*\([\d,\s\-]+\)|\b(AND|OR|NOT)\b", re.I)


def _text(ws, row, col):
    """Cell value, following a merged range to its anchor."""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            v = ws.cell(rng.min_row, rng.min_col).value
            return "" if v is None else str(v).strip()
    v = ws.cell(row, col).value
    return "" if v is None else str(v).strip()


def column_scores(ws, limit=400):
    """How many cells in each column look like a condition."""
    scores = {}
    for c in range(1, ws.max_column + 1):
        hits = 0
        for r in range(1, min(ws.max_row, limit) + 1):
            v = ws.cell(r, c).value
            if v in (None, ""):
                continue
            if LOGIC_HINT.search(str(v)):
                hits += 1
        scores[c] = hits
    return scores


def row_scores(ws, limit=400):
    """How many cells in each row look like a condition."""
    scores = {}
    for r in range(1, min(ws.max_row, limit) + 1):
        hits = 0
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v in (None, ""):
                continue
            if LOGIC_HINT.search(str(v)):
                hits += 1
        scores[r] = hits
    return scores


def is_vertical(ws):
    """True when conditions run down a column rather than across a row."""
    best_col = max(column_scores(ws).values(), default=0)
    best_row = max(row_scores(ws).values(), default=0)
    return best_col >= 2 and best_col > best_row


def find_columns(ws):
    """Locate the condition column, the label column and any group column."""
    scores = column_scores(ws)
    cond = max(scores, key=lambda c: (scores[c], -c))
    if scores[cond] < 2:
        raise ValueError("no column of conditions found")

    # Rows that carry a condition define where labels must also appear
    cond_rows = [r for r in range(1, ws.max_row + 1)
                 if _text(ws, r, cond) and LOGIC_HINT.search(_text(ws, r, cond))]

    # The label column carries a different value on every row. A merged group
    # column reads as filled on every row too, so counting filled cells picks
    # the wrong one - distinct values separate them.
    profile = {}
    for c in range(1, ws.max_column + 1):
        if c == cond:
            continue
        values = [_text(ws, r, c) for r in cond_rows]
        filled = [v for v in values if v]
        profile[c] = (len(set(filled)), len(filled))
    if not profile:
        raise ValueError("no label column found")

    label_col = max(profile, key=lambda c: (profile[c][0], profile[c][1], -abs(c - cond)))

    # A group column sits left of the labels and repeats across their rows
    group_col = None
    for c in range(1, label_col):
        distinct, filled = profile.get(c, (0, 0))
        if filled and distinct < profile[label_col][0]:
            group_col = c
            break

    return cond, label_col, group_col, cond_rows


def read_headerless(path_or_buffer, sheet=None, default_width=10,
                    lo=0, hi=9999, placeholders=None, total_logic=""):
    """Return ({name: points}, report) for a vertical sheet with no headings."""
    wb = load(path_or_buffer)
    names = [sheet] if sheet else wb.sheetnames
    banners, report = {}, []

    for name in names:
        ws = wb[name]
        if not is_vertical(ws):
            continue
        try:
            cond_col, label_col, group_col, cond_rows = find_columns(ws)
        except ValueError:
            continue

        points, current_group = [], ""
        for r in range(1, ws.max_row + 1):
            label = _text(ws, r, label_col)
            condition = _text(ws, r, cond_col)

            if group_col:
                g = _text(ws, r, group_col)
                if g:
                    current_group = g

            if not label and not condition:
                continue

            # Skip a heading row such as Label | Response | Condition
            if condition.lower() in ("condition", "logic", "definition", "conditions"):
                continue

            # Without a group column, a label on its own opens a new group.
            # With one, a label and no condition is a banner column whose
            # condition is missing - that must be reported, not dropped.
            if not condition:
                if label and not group_col:
                    current_group = label
                    continue
                if not label:
                    continue

            t = translate(condition, lo=lo, hi=hi, placeholders=placeholders)
            if t.status == "blocked" and total_logic and not label:
                continue

            n = len(points) + 1
            points.append({
                "column": n,
                "label": label,
                "logic": t.expr,
                "width": default_width,
                "tiers": [current_group],
                "super": current_group,
                "group": "",
                "condition": t.source,
                "status": t.status,
                "note": t.note,
            })
            report.append({
                "banner": name, "column": n, "label": label,
                "condition": t.source, "expression": t.expr,
                "status": t.status, "note": t.note,
            })

        if points:
            banners[name] = points

    if not banners:
        raise ValueError("no vertical banner table found in this workbook")
    return banners, report


# ====================================================================
# codebook.py
# ====================================================================
#
# Resolve value labels to variable codes using a codebook.
#
# The hardest banner specs are the ones that give only value labels: a column
# headed "50-99 employees" under a group called "Company size", with no
# variable name and no code. Nothing in the banner says that is `D1(3)` -
# that lives in the data. So the data has to be supplied.
#
# Two sources work:
#
#   SPSS .sav      variable and value labels are read directly
#   spreadsheet    three columns: Variable | Value | Label
#
# Given those, a label can be looked up. Two routes:
#
#   resolve_in()    the variable is known - find the code whose label matches
#   infer_group()   the variable is not known - find the variable whose set of
#                   value labels best matches the group's set of column labels
#
# The second is what makes a label-only spec workable. A group with the
# labels {Pack buyers, Loose stick buyers} matches exactly one variable in a
# typical codebook, and once the variable is identified every code follows.
#
# Every result carries a score and the matches are reported, because a label
# that merely resembles a value label is a trap, not an answer.
#

import re
from difflib import SequenceMatcher


# Labels are compared after stripping case, punctuation and filler words
FILLER = {"the", "a", "an", "of", "or", "and", "to", "in", "is", "are",
          "than", "those", "who", "that"}


def normalise(text):
    text = re.sub(r"[^\w\s%+<>-]", " ", str(text or "").lower())
    words = [w for w in text.split() if w not in FILLER]
    return " ".join(words)


def similarity(a, b):
    na, nb = normalise(a), normalise(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    ratio = SequenceMatcher(None, na, nb).ratio()
    ta, tb = set(na.split()), set(nb.split())
    overlap = len(ta & tb) / max(len(ta | tb), 1)
    return max(ratio, overlap)


def from_sav(path):
    """Read variable and value labels from an SPSS file."""
    import pyreadstat
    _, meta = pyreadstat.read_sav(path, metadataonly=True)
    book = {}
    for var, mapping in (meta.variable_value_labels or {}).items():
        book[var] = {code: str(label) for code, label in mapping.items()}
    return book


def from_sheet(path_or_buffer, sheet=None):
    """Read a Variable | Value | Label spreadsheet."""
    wb = load(path_or_buffer)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]

    # Locate the three columns from a heading row, else assume the first three
    cols = {"variable": 1, "value": 2, "label": 3}
    start = 1
    for r in range(1, min(ws.max_row, 20) + 1):
        found = {}
        for c in range(1, min(ws.max_column, 12) + 1):
            v = ws.cell(r, c).value
            if v in (None, ""):
                continue
            key = str(v).strip().lower()
            if key in ("variable", "var", "variable name", "name"):
                found["variable"] = c
            elif key in ("value", "code", "val"):
                found["value"] = c
            elif key in ("label", "value label", "text"):
                found["label"] = c
        if len(found) == 3:
            cols, start = found, r + 1
            break

    book = {}
    current_var = ""
    for r in range(start, ws.max_row + 1):
        var = ws.cell(r, cols["variable"]).value
        val = ws.cell(r, cols["value"]).value
        lab = ws.cell(r, cols["label"]).value
        if var not in (None, ""):
            current_var = str(var).strip()
        if current_var and val not in (None, "") and lab not in (None, ""):
            try:
                code = int(float(val))
            except (TypeError, ValueError):
                continue
            book.setdefault(current_var, {})[code] = str(lab).strip()
    return book


def load_codebook(path_or_buffer, filename=""):
    """Read a codebook from a .sav or a spreadsheet."""
    name = (filename or getattr(path_or_buffer, "name", "") or str(path_or_buffer)).lower()
    if name.endswith(".sav"):
        return from_sav(path_or_buffer)
    return from_sheet(path_or_buffer)


def resolve_in(book, variable, label, threshold=0.80):
    """Find the code in `variable` whose value label matches `label`."""
    mapping = book.get(variable) or book.get(variable.upper()) or \
        book.get(variable.lower()) or {}
    if not mapping:
        return None, 0.0, f"variable {variable} not in the codebook"
    best_code, best_score = None, 0.0
    for code, value_label in mapping.items():
        score = similarity(label, value_label)
        if score > best_score:
            best_code, best_score = code, score
    if best_score < threshold:
        return None, best_score, (
            f"best match in {variable} was {mapping.get(best_code)!r} "
            f"at {best_score:.0%} - below the threshold")
    return best_code, best_score, ""


def infer_group(book, labels, threshold=0.75):
    """Find the variable whose value labels best match a group's labels.

    Returns (variable, mapping of label -> code, score). The score is the
    mean of the per-label best matches, so a variable only wins when every
    column in the group finds a home in it.
    """
    wanted = [l for l in labels if str(l).strip()]
    if not wanted:
        return None, {}, 0.0

    best = (None, {}, 0.0)
    for var, mapping in book.items():
        if len(mapping) < len(wanted):
            continue
        assigned, total, used = {}, 0.0, set()
        for label in wanted:
            top_code, top_score = None, 0.0
            for code, value_label in mapping.items():
                if code in used:
                    continue
                score = similarity(label, value_label)
                if score > top_score:
                    top_code, top_score = code, score
            if top_code is not None:
                assigned[label] = top_code
                used.add(top_code)
            total += top_score
        mean = total / len(wanted)
        if mean > best[2]:
            best = (var, assigned, mean)

    if best[2] < threshold:
        return None, {}, best[2]
    return best


def codebook_summary(book):
    return {
        "variables": len(book),
        "codes": sum(len(m) for m in book.values()),
    }


def resolve_points(points, book, threshold=0.75, only_missing=True):
    """Fill in logic for points that have none, using the codebook.

    Columns are handled a group at a time: the group's labels are matched
    against each variable's value labels, and the variable that fits the
    whole group wins. Resolving a group together rather than a label at a
    time is what makes this reliable - a single label like "Yes" matches
    dozens of variables, while a set of them usually matches one.

    Returns a report, one row per column touched.
    """
    groups = {}
    for p in points:
        if only_missing and p.get("logic"):
            continue
        key = (p.get("tiers") or [""])[0]
        groups.setdefault(key, []).append(p)

    report = []
    for group, members in groups.items():
        labels = [p["label"].replace("\n", " ").strip() for p in members]
        var, mapping, score = infer_group(book, labels, threshold)

        if not var:
            for p in members:
                report.append({
                    "column": p["column"], "group": group, "label": p["label"],
                    "variable": "", "logic": "", "score": score,
                    "status": "unmatched",
                    "note": f"no variable in the codebook fits this group "
                            f"(best {score:.0%})",
                })
            continue

        for p, label in zip(members, labels):
            code = mapping.get(label)
            if code is None:
                report.append({
                    "column": p["column"], "group": group, "label": p["label"],
                    "variable": var, "logic": "", "score": score,
                    "status": "unmatched",
                    "note": f"{var} fits the group but this label found no code",
                })
                continue
            p["logic"] = f"{var}({code})"
            report.append({
                "column": p["column"], "group": group, "label": p["label"],
                "variable": var, "logic": p["logic"], "score": score,
                "status": "matched" if score >= 0.95 else "near",
                "note": "" if score >= 0.95 else
                        f"group matched at {score:.0%} - check this one",
            })
    return report


# ====================================================================
# profiles.py
# ====================================================================
#
# Per-client banner conventions.
#
# Each client wants the same banner built slightly differently: whether a
# total column is prepended, what logic that column uses, how wide the
# columns run, and which header directives are carried. Holding those as
# named profiles means the person generating the banner picks the client
# rather than remembering the settings.
#
# The total column is the one that matters most. A spec sheet lists the
# analytical columns only; the total is added by the job. If it is added
# without a label being reserved for it, every header label sits one column
# to the left of the data it describes - which is a silent error, because
# the numbers underneath are all correct.
#

TOTAL_LOGIC = "TN"   # what a total column uses; confirmed from a live job file

PROFILES = {
    "Other projects (house format)": {
        "key": "other",
        "description": (
            "The team's own fixed top-break. A total column is prepended and "
            "labelled, so the spec's columns start at position 2."
        ),
        "prepend_total": True,
        "total_label": "Total",
        # the heading row carries "Total" above the total column too, rather
        # than leaving that span unlabelled - taken from a live job file
        "total_group": "Total",
        "total_logic": TOTAL_LOGIC,
        # every column runs 30 characters wide in this client's banners
        "total_width": 30,
        "default_width": 30,
        "spaces_before": 1,
        "stat_test": "^  ,0",
        "comparison_groups": "0,0",
        "options": "1,SB,HD,W200",
        "point_width": 70,
        "wrap_labels": True,
    },
    "IN2": {
        "key": "in2",
        "description": (
            "Banner plan with its own Total row, so no total is prepended - "
            "the spec's 'All respondents' becomes TN."
        ),
        "prepend_total": False,
        "total_label": "Total",
        "total_logic": TOTAL_LOGIC,
        "total_width": 20,
        "default_width": 10,
        # taken from a live IN2 job file: total and the two lead
        # segment columns run wide, the rest at the default
        "width_overrides": {1: 20, 2: 20, 5: 20},
        "spaces_before": 1,
        "stat_test": "^  ,0",
        "comparison_groups": "0,0",
        "options": "1,SB,HD,W200",
        "point_width": 70,
        "wrap_labels": True,
    },
    "Harmon Research": {
        "key": "harmon",
        "description": (
            "Banner plan, one row per column, conditions in plan syntax. "
            "Several banners per workbook."
        ),
        "prepend_total": False,
        "total_label": "Total",
        "total_logic": TOTAL_LOGIC,
        "total_width": 20,
        "default_width": 20,
        "spaces_before": 1,
        "stat_test": "^  ,0",
        "comparison_groups": "0,0",
        "options": "1,SB,HD,W200",
        "point_width": 70,
        "wrap_labels": True,
    },
}

DEFAULT_PROFILE = "Other projects (house format)"


def get_profile(name):
    return dict(PROFILES.get(name, PROFILES[DEFAULT_PROFILE]))


def prepend_total(points, profile):
    """Add the total column at position 1, with a label reserved for it.

    Returns the new list. Existing column numbers are renumbered so the
    header block and the logic lines stay in step - the failure this is
    here to prevent is a total column with no label, which shifts every
    header one place left while leaving the numbers correct underneath.
    """
    if not profile.get("prepend_total"):
        return points

    first_group = (points[0].get("tiers") or [""])[0] if points else ""
    depth = max(1, len(points[0].get("tiers") or [""])) if points else 1
    heading = profile.get("total_group", "")
    total = {
        "column": 1,
        "label": profile.get("total_label", "Total"),
        "logic": profile.get("total_logic", TOTAL_LOGIC),
        "width": int(profile.get("total_width", 20)),
        # Its own span in the heading row, so it neither absorbs the first
        # group's heading nor pushes that heading off centre.
        "tiers": [heading] + [""] * (depth - 1),
        "super": heading,
        "group": "",
        "is_total": True,
    }
    out = [total] + list(points)
    for i, p in enumerate(out, start=1):
        p["column"] = i
    return out


def settings_from(profile):
    """The emit() settings a profile implies."""
    return {
        "spaces_before": profile.get("spaces_before", 1),
        "stat_test": profile.get("stat_test", "^  ,0"),
        "comparison_groups": profile.get("comparison_groups", "0,0"),
        "options": profile.get("options", "1,SB,HD,W200"),
        "point_width": profile.get("point_width", 70),
        "wrap_labels": profile.get("wrap_labels", True),
    }


def check_total_alignment(points):
    """Warn when a total column exists without a label reserved for it."""
    notes = []
    for p in points:
        logic = str(p.get("logic", "")).strip().upper()
        if logic in ("TN", "TOTAL") and not str(p.get("label", "")).strip():
            notes.append(
                f"column {p['column']} is a total column with no label - every "
                f"header label will sit one column left of the data it describes")
    return notes


# ====================================================================
# expand.py
# ====================================================================
#
# Expand a spec row that stands for several banner columns.
#
# Some specs compress a whole group into one row, putting the group name and
# its members in the label and a code range in the condition:
#
#     label      CP=Faisalabad, Gujranwala, Lahore, Sialkot
#     condition  QS1=1-4
#
# That is four banner columns, not one: CP is the heading, the four cities are
# the labels, and the codes run 1, 2, 3, 4 in the order the names are listed.
# The deliverable shows them that way, so the banner has to as well.
#
# Expansion only happens when the arithmetic agrees - the number of names
# listed must equal the number of codes in the range. When it does not, the
# row is left exactly as it was and the mismatch is reported, because a guess
# about which name goes with which code would be invisible in the output and
# wrong in the data.
#

import re

# 'CP=Faisalabad, Gujranwala, Lahore, Sialkot'
LISTED = re.compile(r"^\s*([^=]{1,40}?)\s*=\s*(.+)$", re.S)
RANGE = re.compile(r"^([A-Za-z_]\w*)\(\s*(\d+)\s*-\s*(\d+)\s*\)$")


def _names(text):
    parts = [p.strip() for p in re.split(r",|/|\band\b", text) if p.strip()]
    return parts


def expand_points(points, report=None):
    """Return (new_points, notes). Expands rows that stand for a group."""
    out, notes = [], []

    for p in points:
        label = str(p.get("label", "")).replace("\n", " ").strip()
        logic = str(p.get("logic", "")).strip()

        m_lab = LISTED.match(label)
        m_log = RANGE.match(logic)
        if not (m_lab and m_log):
            out.append(p)
            continue

        heading, listed = m_lab.group(1).strip(), m_lab.group(2)
        names = _names(listed)
        var, lo, hi = m_log.group(1), int(m_log.group(2)), int(m_log.group(3))
        codes = list(range(lo, hi + 1))

        if len(names) < 2:
            out.append(p)
            continue

        if len(names) != len(codes):
            notes.append(
                f"column {p['column']} ({label[:40]!r}) lists {len(names)} "
                f"names but the condition covers {len(codes)} codes "
                f"({var} {lo}-{hi}) - left as a single column, since pairing "
                f"them would be guesswork")
            out.append(p)
            continue

        for name, code in zip(names, codes):
            child = dict(p)
            child["label"] = name
            child["logic"] = f"{var}({code})"
            child["tiers"] = [heading]
            child["super"] = heading
            child["group"] = ""
            out.append(child)

        notes.append(
            f"column {p['column']}: expanded into {len(names)} columns under "
            f"{heading!r} - {', '.join(f'{n} = {var}({c})' for n, c in list(zip(names, codes))[:3])}"
            + (" ..." if len(names) > 3 else ""))

    for i, p in enumerate(out, start=1):
        p["column"] = i
    return out, notes


# ====================================================================
# generator.py
# ====================================================================
# Assemble a WinCross banner file from banner points.

import re



# WinCross writes job files in the Windows Western codepage, not UTF-8. An
# en dash in a group heading comes out as a single 0x96 byte there; written
# as UTF-8 it becomes two bytes and every column to its right shifts, which
# breaks the header alignment the whole file depends on.
ENCODING = "cp1252"


def encode(text, encoding=ENCODING):
    """Encode a banner file the way WinCross stores one."""
    return text.encode(encoding, errors="replace")


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
    "wrap_labels": True,
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

    # A banner where most columns carry the same expression is not a banner:
    # every column reports the total and nothing is being crossed. This is
    # what a layout-only spec produces, and it looks fine until the tables run.
    for logic, cols in seen.items():
        if len(cols) >= max(3, len(points) // 2):
            notes.append(
                f"{len(cols)} of {len(points)} columns carry the same logic "
                f"({logic}) - every one of them will report the same base as "
                f"the total, and the banner will not cross anything. The spec "
                f"is almost certainly missing its conditions")
            break
    else:
        for logic, cols in seen.items():
            if len(cols) > 1:
                notes.append(
                    f"columns {cols} share identical logic - they will always "
                    f"report the same base and can never test significant "
                    f"against each other")

    # Two columns of the same group drawing on overlapping codes of the same
    # variable will double-count anyone in the overlap. Within a group the
    # bands are meant to partition, so an overlap is nearly always a typo.
    pattern = re.compile(r"^([A-Za-z_]\w*)\((\d+)\s*-\s*(\d+)\)$")
    by_group = {}
    for p in points:
        key = (p.get("tiers") or [""])[0]
        m = pattern.match(str(p.get("logic", "")).strip())
        if m:
            by_group.setdefault(key, []).append(
                (p["column"], p["label"], m.group(1),
                 int(m.group(2)), int(m.group(3))))
    for group, bands in by_group.items():
        for i in range(len(bands)):
            for j in range(i + 1, len(bands)):
                c1, l1, v1, a1, b1 = bands[i]
                c2, l2, v2, a2, b2 = bands[j]
                if v1 != v2:
                    continue
                lo, hi = max(a1, a2), min(b1, b2)
                if lo <= hi:
                    shared = f"{lo}" if lo == hi else f"{lo}-{hi}"
                    notes.append(
                        f"columns {c1} ({l1!r}) and {c2} ({l2!r}) both include "
                        f"{v1} code {shared} - anyone there is counted twice, "
                        f"and the bands will not sum to the group total")

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
    errors, warnings, stats = validate(points, spaces, cfg["column_divider"],
                                          wrap=cfg.get("wrap_labels", True))
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
                           cfg.get("justification"),
                           wrap=cfg.get("wrap_labels", True))

    return "\n".join(lines) + "\n", warnings, [], stats


# ====================================================================
# template.py
# ====================================================================
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

#
# Turn a layout-only banner sheet into a fill-in template.
#
# Some banners arrive as layout alone: group headings and column labels, no
# conditions. There is nothing to generate from that, because the mapping
# from a label like "50-99 employees" to a code like `D1(3)` lives in the
# questionnaire, not in the banner.
#
# Rather than rejecting the file, this writes it back out with a Condition
# row added, ready to be completed and re-uploaded. It also extracts the
# variable name from each group heading - "Company size (D1)" gives `D1` -
# and offers a sequential draft, clearly marked as unverified, since code
# order frequently does not follow label order.
#

import re


VARIABLE = re.compile(r"\(([A-Za-z]\w*)\)\s*$")

HEAD = Font(name="Arial", size=10, bold=True)
BODY = Font(name="Arial", size=10)
DRAFT = Font(name="Arial", size=10, italic=True, color="9C5700")
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
GROUP_FILL = PatternFill("solid", fgColor="DCE6F1")
LABEL_FILL = PatternFill("solid", fgColor="F2F2F2")
FILLME = PatternFill("solid", fgColor="FFF2CC")
DRAFT_FILL = PatternFill("solid", fgColor="FCE4D6")


def variable_of(heading):
    """'Company size (D1)' -> 'D1'. Returns '' when no variable is named."""
    m = VARIABLE.search((heading or "").replace("\n", " ").strip())
    return m.group(1) if m else ""


def draft_conditions(points):
    """Sequential code guesses per group. Unverified by construction."""
    drafts, counters = [], {}
    for p in points:
        group = (p.get("tiers") or [""])[0]
        var = variable_of(group)
        if not var:
            drafts.append("")
            continue
        counters[var] = counters.get(var, 0) + 1
        drafts.append(f"{var}({counters[var]})")
    return drafts


def layout_checks(points):
    """Problems visible from the layout alone, before any logic exists."""
    notes = []

    groups, order = {}, []
    for p in points:
        g = (p.get("tiers") or [""])[0].replace("\n", " ").strip()
        if not g:
            continue
        if g not in groups:
            groups[g] = []
            order.append(g)
        groups[g].append(p["label"].replace("\n", " ").strip())

    # One variable serving two different groups is nearly always a mistake in
    # the heading: the same variable cannot carry two different code frames.
    by_var = {}
    for g in order:
        v = variable_of(g)
        if v:
            by_var.setdefault(v, []).append(g)
    for var, gs in by_var.items():
        if len(gs) > 1:
            notes.append(
                f"variable {var} is named by {len(gs)} different groups "
                f"({', '.join(repr(g) for g in gs)}) - one heading is probably "
                f"wrong, since a variable has only one code frame")

    # A label that spans other labels in its own group is a net, not a code.
    net = re.compile(r"^\s*(\d+)\s*\+|\ball\b|\bany\b|\bnet\b|/", re.I)
    for g in order:
        labels = groups[g]
        for lab in labels:
            if net.search(lab) and len(labels) > 2:
                notes.append(
                    f"{g!r}: {lab!r} looks like a net across several codes "
                    f"rather than a single code - the draft will be wrong for it")
    return notes


def write_template(points, path, include_drafts=True):
    """Write a workbook with the layout preserved and a Condition row added."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Banner"

    drafts = draft_conditions(points) if include_drafts else [""] * len(points)

    ws.cell(2, 1, "group heading").font = HEAD
    ws.cell(3, 1, "column label").font = HEAD
    ws.cell(4, 1, "CONDITION - fill this in").font = HEAD
    if include_drafts:
        ws.cell(5, 1, "draft - VERIFY before use").font = DRAFT

    for i, p in enumerate(points):
        col = 2 + i
        group = (p.get("tiers") or [""])[0]

        c = ws.cell(2, col, group)
        c.font, c.alignment, c.fill = HEAD, CTR, GROUP_FILL
        c = ws.cell(3, col, p["label"])
        c.font, c.alignment, c.fill = BODY, CTR, LABEL_FILL
        c = ws.cell(4, col, "")
        c.font, c.alignment, c.fill = BODY, CTR, FILLME
        if include_drafts:
            c = ws.cell(5, col, drafts[i])
            c.font, c.alignment, c.fill = DRAFT, CTR, DRAFT_FILL

        ws.column_dimensions[c.column_letter].width = 22

    ws.column_dimensions["A"].width = 26
    for r in (2, 3, 4, 5):
        ws.row_dimensions[r].height = 34

    notes = wb.create_sheet("Read me")
    lines = [
        ("This banner arrived as layout only", True),
        ("", False),
        ("It has group headings and column labels but no conditions, so there", False),
        ("is nothing to generate from yet. Fill in row 4 and upload this file", False),
        ("back to the generator.", False),
        ("", False),
        ("Row 4 is the one to complete. Leave rows 2 and 3 as they are.", False),
        ("", False),
        ("Conditions can be written either way:", True),
        ("  WinCross syntax    D1(1)   D1(1,2)   S5(3)", False),
        ("  Plan syntax        D1=1    D1=1 OR 2  S5=3", False),
        ("", False),
        ("About the draft row", True),
        ("Row 5 numbers the responses within each group in the order they", False),
        ("appear, using the variable named in the heading. It is a starting", False),
        ("point, not an answer: code order often does not follow label order,", False),
        ("nets and overlapping categories never do, and a label such as", False),
        ("'50+ employees' is usually a net across several codes rather than", False),
        ("one of them.", False),
        ("", False),
        ("Check every draft against the questionnaire before copying it into", False),
        ("row 4. Delete row 5 once you are done.", False),
    ]
    for i, (text, bold) in enumerate(lines, start=1):
        notes.cell(i, 1, text).font = Font(name="Arial", size=11, bold=bold)
    notes.column_dimensions["A"].width = 78

    wb.save(path)
    return path


def apply_drafts(points, total_logic="TOTAL"):
    """Fill empty logic with sequential drafts so a banner file can be built.

    This produces a structurally complete WinCross file from a layout-only
    sheet: widths, header block, directives and stat letters are all correct,
    and only the logic lines need checking. Every drafted line is returned so
    it can be listed for verification - the codes are positional guesses, not
    read from the questionnaire.
    """
    drafts = draft_conditions(points)
    drafted = []
    for p, d in zip(points, drafts):
        if p.get("logic"):
            continue
        group = (p.get("tiers") or [""])[0].replace("\n", " ").strip()
        if not d:
            p["logic"] = total_logic
            drafted.append({"column": p["column"], "label": p["label"],
                            "group": group, "logic": total_logic,
                            "why": "no variable in heading - total column base used"})
            continue
        p["logic"] = d
        drafted.append({"column": p["column"], "label": p["label"],
                        "group": group, "logic": d,
                        "why": "position within group; verify against questionnaire"})
    return drafted


# ====================================================================
# excel_report.py
# ====================================================================
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

#
# Write the generated banner out as a workbook.
#
# The banner file itself is fixed-width text, which is right for WinCross and
# awkward for a person checking it. This produces the same content as a
# workbook that can be reviewed, commented on and circulated:
#
#   Columns     one row per banner point - number, stat letter, group, label,
#               width and logic. This is the sheet to check against the spec.
#   Checks      everything the tool raised while reading the spec
#   Banner file the raw text, one row per line, for reference
#
# Nothing here is read back by the tool; it exists so the banner can be
# checked by someone who does not read WinCross syntax.
#


HEAD = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY = Font(name="Calibri", size=11)
MONO = Font(name="Consolas", size=9)
HEAD_FILL = PatternFill("solid", fgColor="2F5597")
BAND = PatternFill("solid", fgColor="F2F6FB")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
ERROR_FILL = PatternFill("solid", fgColor="F8CBAD")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(vertical="top", wrap_text=True)
TOP = Alignment(vertical="top")


def _letters(n):
    out = []
    for i in range(n):
        letter = chr(ord("A") + i % 26)
        cycle = i // 26
        out.append(letter if cycle == 0 else f"{letter}{cycle}")
    return out


def _header(ws, row, titles, widths):
    for i, (title, width) in enumerate(zip(titles, widths), start=1):
        c = ws.cell(row, i, title)
        c.font, c.fill, c.border = HEAD, HEAD_FILL, BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = ws.cell(row + 1, 1)


def write_report(points, banner_text, checks=None, stats=None, meta=None,
                 path_or_buffer="banner.xlsx"):
    """Write the banner, its column map and its checks to a workbook."""
    wb = Workbook()

    # ------------------------------------------------------------ Columns
    ws = wb.active
    ws.title = "Columns"
    letters = _letters(len(points))

    info = []
    if meta:
        info.append(f"Client: {meta.get('profile', '-')}")
        info.append(f"Source: {meta.get('source', '-')}")
        info.append(f"Layout: {meta.get('format', '-')}")
    if stats:
        info.append(f"{stats['columns']} columns")
        info.append(f"report width {stats['report_width']} characters")
    ws.cell(1, 1, "   |   ".join(info)).font = Font(size=10, italic=True,
                                                   color="595959")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)

    _header(ws, 3, ["Col", "Letter", "Group", "Label", "Width", "Logic"],
            [6, 8, 34, 34, 8, 52])

    for i, p in enumerate(points):
        r = 4 + i
        group = (p.get("tiers") or [""])[0].replace("\n", " ").strip()
        values = [p["column"], letters[i], group,
                  p["label"].replace("\n", " ").strip(), p["width"], p["logic"]]
        for col, value in enumerate(values, start=1):
            c = ws.cell(r, col, value)
            c.font = MONO if col == 6 else BODY
            c.border, c.alignment = BOX, WRAP if col in (3, 4, 6) else TOP
            if p.get("is_total"):
                c.fill = TOTAL_FILL
            elif i % 2:
                c.fill = BAND
    ws.auto_filter.ref = f"A3:F{3 + len(points)}"

    # ------------------------------------------------------------- Checks
    cs = wb.create_sheet("Checks")
    _header(cs, 1, ["Level", "What the tool found"], [12, 120])
    fills = {"error": ERROR_FILL, "warn": WARN_FILL, "ok": OK_FILL}
    names = {"error": "Error", "warn": "Warning", "ok": "OK"}
    rows = checks or []
    if not rows:
        rows = [("ok", "No problems raised.")]
    for i, (level, text) in enumerate(rows):
        r = 2 + i
        for col, value in enumerate([names.get(level, level), text], start=1):
            c = cs.cell(r, col, value)
            c.font, c.border, c.alignment = BODY, BOX, WRAP
            c.fill = fills.get(level, BAND)

    # -------------------------------------------------------- Banner file
    bf = wb.create_sheet("Banner file")
    _header(bf, 1, ["Line", "Content"], [7, 170])
    for i, line in enumerate(banner_text.rstrip("\n").split("\n")):
        r = 2 + i
        bf.cell(r, 1, i + 1).font = BODY
        c = bf.cell(r, 2, line)
        c.font, c.alignment = MONO, TOP

    wb.save(path_or_buffer)
    return path_or_buffer


def write_wincross_layout(points, path_or_buffer="banner_layout.xlsx",
                          table_title="Table 1"):
    """Write the banner in the shape WinCross exports a run to Excel.

    One column per banner point, with the heading rows merged across the
    columns they span, the labels one per cell, and the stat letter row
    beneath - the layout the Banner sheet of a WinCross export has.

    The frequency rows are left empty: those come from running the tables
    against the data, which only WinCross can do. This shows the header
    layout the banner will produce, so it can be checked before a run.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Banner"

    letters = _letters(len(points))
    first = 4                                   # banner columns start at D

    ws.cell(1, 2, "Apply Filter").font = BODY
    ws.cell(1, 3, table_title).font = BODY
    ws.cell(3, 2, "Job Title 1").font = BODY
    ws.cell(6, 2, "Table Title 1").font = BODY

    # heading rows, one per tier, merged across the columns each spans
    depth = max((len(p.get("tiers") or [""]) for p in points), default=1)
    row = 9
    for t in range(depth):
        ws.cell(row, 2, "Banner Text").font = BODY
        i = 0
        while i < len(points):
            tiers = points[i].get("tiers") or [""]
            value = (tiers[t] if t < len(tiers) else "") or ""
            j = i
            if value:
                while j + 1 < len(points):
                    nxt = points[j + 1].get("tiers") or [""]
                    if ((nxt[t] if t < len(nxt) else "") or "") != value:
                        break
                    j += 1
            c = ws.cell(row, first + i, value.replace("\n", " ").strip())
            c.font, c.alignment, c.border = HEAD, Alignment(
                horizontal="center", vertical="center", wrap_text=True), BOX
            c.fill = HEAD_FILL
            if j > i:
                ws.merge_cells(start_row=row, start_column=first + i,
                               end_row=row, end_column=first + j)
            i = j + 1
        row += 1

    # column labels, one per cell
    ws.cell(row, 2, "Banner Text").font = BODY
    for i, p in enumerate(points):
        c = ws.cell(row, first + i, p["label"].replace("\n", " ").strip())
        c.font = Font(name="Calibri", size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BOX
        c.fill = TOTAL_FILL if p.get("is_total") else BAND
    label_row = row
    row += 1

    # stat letters
    ws.cell(row, 2, "Banner Text").font = BODY
    for i, letter in enumerate(letters):
        c = ws.cell(row, first + i, f"({letter})")
        c.font = Font(name="Calibri", size=10, color="595959")
        c.alignment = Alignment(horizontal="center")
        c.border = BOX
    row += 1

    # the logic behind each column, so the header can be checked against it
    ws.cell(row, 2, "Banner Logic").font = Font(size=10, italic=True,
                                                color="595959")
    for i, p in enumerate(points):
        c = ws.cell(row, first + i, p["logic"])
        c.font = MONO
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BOX
    row += 1

    ws.cell(row, 2, "Filter Frequency Row").font = BODY
    ws.cell(row, 3, "BASE: Total Respondents").font = BODY
    for i in range(len(points)):
        ws.cell(row, first + i, "").border = BOX

    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 26
    for i in range(len(points)):
        ws.column_dimensions[get_column_letter(first + i)].width = 16
    ws.row_dimensions[label_row].height = 46
    ws.freeze_panes = ws.cell(label_row + 3, first)

    wb.save(path_or_buffer)
    return path_or_buffer


# ====================================================================
# reader.py - format detection
# ====================================================================
def detect(path_or_buffer):
    """Return 'plan', 'vertical' or 'grid'."""
    wb = load(path_or_buffer)
    for name in wb.sheetnames:
        if looks_like_plan(wb[name]):
            return "plan"
    for name in wb.sheetnames:
        if is_vertical(wb[name]):
            return "vertical"
    return "grid"


def read_any(path_or_buffer, default_width=10, width_overrides=None,
             lo=0, hi=9999, placeholders=None, total_logic=""):
    """Read either banner format. Returns (banners, report, fmt)."""
    fmt = detect(path_or_buffer)
    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)

    if fmt == "vertical":
        banners, report = read_headerless(
            path_or_buffer, default_width=default_width, lo=lo, hi=hi,
            placeholders=placeholders, total_logic=total_logic)
        if width_overrides:
            for points in banners.values():
                for p in points:
                    if p["column"] in width_overrides:
                        p["width"] = int(width_overrides[p["column"]])
        return banners, report, fmt

    if fmt == "plan":
        banners, report = read_plan(path_or_buffer, default_width=default_width,
                                    lo=lo, hi=hi, placeholders=placeholders,
                                    total_logic=total_logic)
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

# Messages raised while reading and preparing the spec. They are collected
# rather than printed as they occur, so the page stays readable and
# everything needing attention sits in one place.
CHECKS = []
LEVELS = {"ok": "success", "warn": "warning", "error": "error"}


def note(level, text):
    CHECKS.append((level, text))


LAYOUTS = """
The tool works out which shape a workbook uses. Four are handled.

**Grid** - the banner runs left to right, one spreadsheet column per banner
column: heading rows above the labels, a row of WinCross logic beneath.
Column A must be empty.

**Plan** - one spreadsheet row per banner column, under headings named
Column, Variable, Group, Label, Response, Condition. Several banners can
share one sheet.

**Vertical** - the same, running downwards, but with no heading row. Group
names either sit in their own column or on a row of their own above the
columns they cover.

**Layout only** - headings and labels but no conditions. Nothing can be
generated from this alone: upload a codebook alongside it, or fill in the
conditions first.

---

**Codebook.** When a spec gives value labels but no variable names or codes,
upload an SPSS `.sav` or a sheet of `Variable | Value | Label`. Each group of
columns is matched as a set against every variable, and the variable whose
value labels fit the whole group is used. Matching a group together is what
makes it reliable - "Yes" on its own matches dozens of variables.

**Total column.** Spec sheets list the analytical columns only; the total is
added by the job. If it is added without a label reserved for it, every
header label sits one column left of the data it describes - and the numbers
underneath stay correct, so nothing looks broken.
"""

BUILD = "2026-10-05 / build 12 / conditions inside descriptions"

st.title("WinCross Banner Generator")
st.caption(
    "Turn a banner specification into a WinCross banner file: directives, "
    "logic lines and the header text block."
)
st.caption(f"Version: {BUILD}")

# -------------------------------------------------------- client and input
sel, spec_col, book_col = st.columns([1.1, 1, 1])

with sel:
    profile_name = st.selectbox(
        "Client / project", list(PROFILES),
        index=list(PROFILES).index(DEFAULT_PROFILE),
    )
    profile = get_profile(profile_name)

with spec_col:
    uploaded = st.file_uploader(
        "Banner specification (.xlsx)", type=["xlsx", "xlsm"]
    )

with book_col:
    codebook_file = st.file_uploader(
        "Codebook, optional (.sav / .xlsx)", type=["sav", "xlsx", "xlsm"]
    )

st.caption(profile["description"])

# ------------------------------------------------------------- settings
with st.sidebar:
    st.header("Settings")
    st.caption(f"Defaults come from **{profile_name}**. Change anything here.")

    banner_id = st.number_input("Banner ID", min_value=1, value=1, step=1)
    banner_title = st.text_input("Banner title (BT)", value="")
    banner_filter = st.text_input("Filter logic (BF)", value="")

    st.subheader("Total column")
    add_total = st.checkbox(
        "Prepend a total column", value=bool(profile.get("prepend_total")),
        help="Adds the total as column 1 with a label reserved for it, so "
             "the header stays aligned with the data.",
    )
    total_label = st.text_input("Total column label",
                                value=profile.get("total_label", "Total"))
    total_logic = st.text_input(
        "Total column logic", value=profile.get("total_logic", "TN"),
        help="WinCross uses TN. Also used for any column a plan describes "
             "as 'All respondents'.",
    )

    st.subheader("Width and spacing")
    st.caption("WinCross requires the same 'spaces before' on every column, "
               "maximum 5. Default column width is 10.")
    spaces_before = st.slider("Spaces before each column", 1, 5,
                              int(profile.get("spaces_before", 1)))
    default_width = st.number_input(
        "Default column width", 1, 80, int(profile.get("default_width", 10)))
    width_text = st.text_input(
        "Width overrides",
        value=", ".join(f"{k}:{v}" for k, v in
                        (profile.get("width_overrides") or {}).items()),
        help="Column:width pairs, e.g. 1:20, 2:20, 6:20",
    )
    divider = st.text_input(
        "Column divider character(s)", value="", max_chars=5,
        help="Plain-text reports only. Cannot exceed 'spaces before'.",
    )

    st.subheader("Header block")
    emit_header = st.checkbox("Generate header text block", value=True)
    wrap_labels = st.checkbox(
        "Wrap long labels", value=bool(profile.get("wrap_labels", True)),
        help="Word-wrap across extra header lines instead of cutting a label "
             "off at the column width.",
    )
    stub_width = st.number_input("Left stub width", 0, 20, 1)
    just = st.selectbox("Text justification", ["center", "left", "right"], 0)

    st.subheader("Conditions")
    st.caption("Plans write conditions like 'S4>14'. WinCross needs both ends "
               "of a range, so the open end is supplied here.")
    range_lo = st.number_input("Assumed lower bound", -9999, 9999, 0)
    range_hi = st.number_input("Assumed upper bound", 1, 999999, 9999)
    cutpoints = st.text_input(
        "Cut points", value="",
        help="For placeholders like 'S5r4>XX'. Pairs: S5r4:10, S5r6:5",
    )

    st.subheader("Directives")
    stat_test = st.text_input("Statistical testing (ST)",
                              value=profile.get("stat_test", "^  ,0"))
    comparison = st.text_input("Comparison groups (CP)",
                               value=profile.get("comparison_groups", "0,0"))
    weights = st.text_input("Weights (WT)", value="")
    options = st.text_input("Options (OP)",
                            value=profile.get("options", "1,SB,HD,W200"))
    point_width = st.number_input("Banner point width (^W)", 1, 200,
                                  int(profile.get("point_width", 70)))
    normalise = st.checkbox("Normalise logic spacing", value=True)


if not uploaded:
    st.info("Upload a banner specification to begin.")
    with st.expander("Which layouts are supported?", expanded=True):
        st.markdown(LAYOUTS)
    st.stop()

# --------------------------------------------------------------- read spec
try:
    overrides = parse_width_overrides(width_text)
    cuts = parse_placeholders(cutpoints)
except ValueError as exc:
    st.error(f"Could not read a setting: {exc}")
    st.stop()

try:
    banners, report, fmt = read_any(
        uploaded, default_width=int(default_width), width_overrides=overrides,
        lo=int(range_lo), hi=int(range_hi),
        placeholders=cuts, total_logic=total_logic,
    )
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read the workbook: {exc}")
    st.stop()

shape = {
    "grid": "grid - one spreadsheet column per banner column",
    "plan": "plan - one row per banner column, with headings",
    "vertical": "vertical - one row per banner column, no headings",
}[fmt]
note("ok", f"Detected a **{shape}** layout. {len(banners)} banner(s) found.")

chosen = (st.selectbox("Which banner?", list(banners))
          if len(banners) > 1 else list(banners)[0])
points = banners[chosen]

# --------------------------------------------------- codebook and drafting
book, cb_report = {}, []
if codebook_file is not None:
    try:
        book = load_codebook(codebook_file, codebook_file.name)
        info = codebook_summary(book)
        note("ok", f"Codebook read: {info['variables']} variables, "
                   f"{info['codes']} codes.")
    except Exception as exc:  # noqa: BLE001
        note("error", f"Could not read the codebook: {exc}")

if report:
    rows = [r for r in report if r["banner"] == chosen]
    blocked = [r for r in rows if r["status"] == "blocked"]
    assumed = [r for r in rows if r["status"] == "assumed"]
    if blocked:
        note("error", f"{len(blocked)} condition(s) could not be translated "
                      f"and are excluded. See the Translation tab.")
    if assumed:
        note("warn", f"{len(assumed)} condition(s) needed a range bound to be "
                     f"assumed. See the Translation tab.")

if book and any(not p.get("logic") for p in points):
    cb_report = resolve_points(points, book)
    filled = sum(1 for r in cb_report if r["logic"])
    near = sum(1 for r in cb_report if r["status"] == "near")
    note("ok" if filled else "warn",
         f"Codebook matched {filled} of {len(cb_report)} unresolved column(s)."
         + (f" {near} matched below 95% - check them." if near else ""))

drafted = []
layout_only = fmt == "grid" and not any(p.get("logic") for p in points)
if layout_only:
    st.error(
        "**This sheet has no conditions in it** - only group headings and "
        "column labels. A banner built from it cannot cross anything: every "
        "column would report the same base as the total.\n\n"
        "Two ways forward, either of which gives a working banner:\n\n"
        "- upload the sheet that carries the conditions, if there is one "
        "(often a second tab or a companion file), or\n"
        "- upload a **codebook** alongside this sheet, so the labels can be "
        "matched to real variable codes."
    )
    proceed = st.checkbox(
        "Generate anyway with drafted logic (positional guesses - the result "
        "will not be usable without checking every line)",
        value=False,
    )
    if not proceed:
        st.stop()
    note("error", "Built from drafted logic: positional guesses, not read "
                  "from a questionnaire. Check every line.")
    for n in layout_checks(points):
        note("warn", n)
    drafted = apply_drafts(points, total_logic=total_logic or "TN")

expand_rows = st.checkbox(
    "Expand rows that list several labels", value=True,
    help="A row like 'CP=Faisalabad, Gujranwala, Lahore, Sialkot' against "
         "'QS1=1-4' is four banner columns, not one. Only expands when the "
         "number of names matches the number of codes.",
)
if expand_rows:
    with_logic = [p for p in points if p.get("logic")]
    without = [p for p in points if not p.get("logic")]
    expanded, exp_notes = expand_points(with_logic)
    if len(expanded) != len(with_logic):
        note("ok", f"Expanded {len(with_logic)} spec rows into "
                   f"{len(expanded)} banner columns.")
    for n in exp_notes:
        note("warn" if "left as a single column" in n else "ok", n)
    points = expanded + without
    for i, p in enumerate(points, start=1):
        p["column"] = i

skipped = [p for p in points if not p.get("logic")]
points = [p for p in points if p.get("logic")]
if skipped:
    note("warn", f"{len(skipped)} column(s) have no usable logic and were left "
                 f"out: " + ", ".join(repr(p["label"]) for p in skipped[:5])
                 + (" ..." if len(skipped) > 5 else ""))

if not points:
    st.error("No columns have usable logic - nothing to generate.")
    for lv, msg in CHECKS:
        getattr(st, LEVELS[lv])(msg)
    st.stop()

if add_total:
    points = prepend_total(points, {
        **profile, "prepend_total": True,
        "total_label": total_label, "total_logic": total_logic or "TN",
    })
    note("ok", f"Total column prepended as column 1 ({total_label!r} / "
               f"{total_logic or 'TN'}).")

for n in check_total_alignment(points):
    note("error", n)

# ---------------------------------------------------------------- generate
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
    "wrap_labels": wrap_labels,
    "justification": {"super": just, "group": just, "column": just},
}

text, warnings, errors, stats = emit(points, settings)
for w in warnings:
    note("warn", w)
for e in errors:
    note("error", e)

# ------------------------------------------------------------------ output
problems = sum(1 for lv, _ in CHECKS if lv in ("warn", "error"))
c1, c2, c3, c4 = st.columns(4)
c1.metric("Columns", stats["columns"])
c2.metric("Report width", f"{stats['report_width']} chars")
c3.metric("Widths used", ", ".join(str(w) for w in stats["widths_used"]))
c4.metric("Needs attention", problems)

names = [f"Checks ({problems})" if problems else "Checks",
         "Banner file", "Column map", "Header preview"]
if report:
    names.append("Translation")
if cb_report:
    names.append("Codebook matches")
if drafted:
    names.append("Drafted logic")
tabs = st.tabs(names)

with tabs[0]:
    st.caption(
        "Everything the tool noticed while reading the spec. Errors mean a "
        "column was left out or will be wrong; warnings are worth a look."
    )
    for level in ("error", "warn", "ok"):
        for lv, msg in CHECKS:
            if lv == level:
                getattr(st, LEVELS[lv])(msg)
    with st.expander("Which layouts are supported?"):
        st.markdown(LAYOUTS)

with tabs[1]:
    if errors:
        st.error("Generation blocked - see the Checks tab.")
    else:
        stem = chosen.replace(":", "").replace(" ", "_")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "Download banner file (.txt)", data=encode(text),
                file_name=f"{stem}.txt", mime="text/plain",
            )
        with d2:
            xbuf = io.BytesIO()
            write_report(
                points, text, CHECKS, stats,
                {"profile": profile_name, "source": uploaded.name, "format": fmt},
                xbuf,
            )
            st.download_button(
                "Download review workbook (.xlsx)", data=xbuf.getvalue(),
                file_name=f"{stem}_review.xlsx",
                mime="application/vnd.openxmlformats-officedocument."
                     "spreadsheetml.sheet",
            )
        st.caption(
            "The .txt is the banner for WinCross. The workbook holds the same "
            "content as a column map plus the checks, for review."
        )
        st.code(text, language="text")

with tabs[2]:
    st.dataframe(
        [{"Col": p["column"], "Width": p["width"],
          "Label": p["label"].replace("\n", " "),
          "Group": (p.get("tiers") or [""])[0].replace("\n", " "),
          "Logic": p["logic"]}
         for p in points],
        use_container_width=True, hide_index=True,
    )

with tabs[3]:
    if not emit_header:
        st.info("Header block generation is switched off in the sidebar.")
    else:
        st.caption(
            "Rule lines span each heading: the summed column widths plus the "
            "spacers between them. Every line runs the full report width."
        )
        st.code("\n".join(text.rstrip("\n").split("\n")[13 + len(points):]),
                language="text")

if report:
    with tabs[names.index("Translation")]:
        st.caption("How each plan condition became WinCross logic. "
                   "'blocked' rows are excluded from the file.")
        st.dataframe(
            [{"Col": r["column"], "Status": r["status"], "Label": r["label"],
              "Condition": r["condition"], "WinCross": r["expression"] or "-",
              "Note": r["note"]}
             for r in report if r["banner"] == chosen],
            use_container_width=True, hide_index=True,
        )

if cb_report:
    with tabs[names.index("Codebook matches")]:
        st.caption(
            "Each group was matched as a set against the codebook: the "
            "variable whose value labels fit the whole group wins. Anything "
            "below 95% is worth checking."
        )
        st.dataframe(
            [{"Col": r["column"], "Status": r["status"],
              "Group": r["group"].replace("\n", " "),
              "Label": r["label"].replace("\n", " "),
              "Variable": r["variable"], "Logic": r["logic"] or "-",
              "Match": f"{r['score']:.0%}", "Note": r["note"]}
             for r in cb_report],
            use_container_width=True, hide_index=True,
        )

if drafted:
    with tabs[names.index("Drafted logic")]:
        st.error(
            "Every line below is a positional guess. Check each against the "
            "questionnaire, or upload a codebook to resolve them properly. "
            "Labels that are nets across several codes will be wrong."
        )
        st.dataframe(
            [{"Col": d["column"], "Group": d["group"],
              "Label": d["label"].replace("\n", " "),
              "Drafted logic": d["logic"], "Why": d["why"]}
             for d in drafted],
            use_container_width=True, hide_index=True,
        )
        buf = io.BytesIO()
        write_template(points, buf)
        st.download_button(
            "Download fill-in spreadsheet instead", data=buf.getvalue(),
            file_name="banner_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.sheet",
        )
