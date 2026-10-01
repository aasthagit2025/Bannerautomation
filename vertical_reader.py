"""
Read a vertical banner sheet that carries no header row.

Two of the shapes seen in practice run downwards but have no
Label/Response/Condition heading to key on:

  separate group column        group headings sit in their own column,
                               merged down the rows they cover

      Store Preference | Those who prefer home depot | D2a: ...
                       | Those who prefer competitor | D2a: ...
      Income           | Less than $500k             | D3 = 1,2,3

  group heading rows           a row with a label but no condition opens a
                               new group; the rows beneath it are its columns

      Q2: Pack Vs. Stick buying
      Pack buyers                                    | Q2=2
      Loose stick buyers                             | Q2=1

Both are handled by locating the condition column first - it is the one
whose cells look like conditions - and then reading the rest relative to it.
Blank separator rows are ignored, and a row whose condition is prose rather
than logic is carried through so the translator can report it.
"""

import re

from .safe_load import load as _load
from .logic_translate import translate

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
    wb = _load(path_or_buffer)
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
