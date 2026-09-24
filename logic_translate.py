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
