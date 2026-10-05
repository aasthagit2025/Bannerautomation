#!/usr/bin/env python3
"""Rebuild the deployable app from the wincross package.

Writes the same bundle to two names: app.py and streamlit_app.py. Streamlit
Cloud looks for either, so whichever the deployment is pointed at will work.
The UI source lives in ui_source.py - edit that, never the generated files.

"""
import ast

MODULES = ["safe_load", "width_spacing", "header_block", "logic_translate",
           "excel_spec", "plan_reader", "vertical_reader", "codebook",
           "profiles", "generator", "template"]
ALIASES = [("hb.render(", "render("), ("ws.validate(", "validate("),
           ("ws.sw_directive(", "sw_directive("), ("ws.SpecError", "SpecError"),
           ("_load(", "load(")]
EXPORTS = ["read_points", "parse_width_overrides", "SheetError", "emit",
           "DEFAULTS", "normalise_logic", "consistency_checks", "validate",
           "width_report", "SpecError", "read_any", "detect", "read_plan",
           "translate", "summarise", "parse_placeholders",
           "write_template", "draft_conditions", "layout_checks", "variable_of",
           "read_headerless", "is_vertical", "load_codebook", "resolve_points",
           "infer_group", "resolve_in", "codebook_summary", "PROFILES",
           "DEFAULT_PROFILE", "get_profile", "prepend_total", "settings_from",
           "check_total_alignment", "PROFILES", "DEFAULT_PROFILE",
           "apply_drafts", "layout_checks", "write_template"]

READER = '''

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
'''


def strip_bare_strings(src):
    """Comment out every bare top-level string expression.

    Streamlit's "magic" renders any bare expression at the top level of the
    script. Once the modules are flattened into one file, each module
    docstring becomes exactly that, and Streamlit prints the lot as page
    content, burying the interface. Keeping the text as comments preserves it
    for anyone reading the file without it being rendered.
    """
    tree = ast.parse(src)
    spans = [
        (n.lineno, n.end_lineno) for n in tree.body
        if isinstance(n, ast.Expr) and isinstance(n.value, ast.Constant)
        and isinstance(n.value.value, str)
    ]
    if not spans:
        return src

    lines = src.split("\n")
    out, i = [], 0
    while i < len(lines):
        span = next((s for s in spans if s[0] - 1 == i), None)
        if span:
            for ln in lines[span[0] - 1:span[1]]:
                body = ln.replace('"""', "").rstrip()
                out.append(("# " + body).rstrip() if body.strip() else "#")
            i = span[1]
        else:
            out.append(lines[i])
            i += 1
    return "\n".join(out)


def build():
    parts = []
    for m in MODULES:
        src = strip_bare_strings(open(f"wincross/{m}.py").read())
        body = [l for l in src.split("\n")
                if not l.strip().startswith("from .")
                and not l.strip().startswith("#!/usr")]
        parts.append(f"# {'='*68}\n# {m}.py\n# {'='*68}\n" + "\n".join(body).strip())
    core = "\n\n\n".join(parts)
    for a, b in ALIASES:
        core = core.replace(a, b)
    core += "\n" + READER
    # safe_load pulls these in at module scope; the bundle imports them once at
    # the top, so drop the duplicated imports from the inlined copy
    for dup in ("import io\nimport re\nimport shutil\nimport tempfile\n"
                "import zipfile\n\nimport openpyxl", "import shutil\nimport tempfile\n"):
        core = core.replace(dup, "")

    app = strip_bare_strings(open("ui_source.py").read()).replace("import wincross as wx\n", "")
    for fn in EXPORTS:
        app = app.replace(f"wx.{fn}", fn)
    head, rest = app.split("st.set_page_config", 1)
    head = head.rstrip().replace(
        "import io\n\nimport streamlit as st",
        "import io\nimport re\nimport zipfile\nimport warnings\n\nimport openpyxl\n"
        "from openpyxl import Workbook\n"
        "from openpyxl.styles import Alignment, Font, PatternFill\n"
        "import streamlit as st")

    out = (head + "\n\n\n" + core + "\n\n\n"
           + "# " + "=" * 68 + "\n# Streamlit interface\n# " + "=" * 68 + "\n\n"
           + "st.set_page_config" + rest)
    open("app.py", "w").write(out)
    open("streamlit_app.py", "w").write(out)
    import ast
    ast.parse(out)
    return len(out.split("\n"))


if __name__ == "__main__":
    print(f"app.py and streamlit_app.py rebuilt: {build()} lines")
