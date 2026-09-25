#!/usr/bin/env python3
"""Rebuild streamlit_app.py from the wincross package.

The deployed app is a single file so there is no folder structure to get
wrong. That file is generated, never edited by hand: edit the package, then
run this.
"""
MODULES = ["width_spacing", "header_block", "logic_translate", "excel_spec",
           "plan_reader", "generator"]
ALIASES = [("hb.render(", "render("), ("ws.validate(", "validate("),
           ("ws.sw_directive(", "sw_directive("), ("ws.SpecError", "SpecError")]
EXPORTS = ["read_points", "parse_width_overrides", "SheetError", "emit",
           "DEFAULTS", "normalise_logic", "consistency_checks", "validate",
           "width_report", "SpecError", "read_any", "detect", "read_plan",
           "translate", "summarise", "parse_placeholders"]

READER = '''

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
             lo=0, hi=9999, placeholders=None, total_logic=""):
    """Read either banner format. Returns (banners, report, fmt)."""
    fmt = detect(path_or_buffer)
    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)

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


def build():
    parts = []
    for m in MODULES:
        src = open(f"wincross/{m}.py").read()
        body = [l for l in src.split("\n")
                if not l.strip().startswith("from .")
                and not l.strip().startswith("#!/usr")]
        parts.append(f"# {'='*68}\n# {m}.py\n# {'='*68}\n" + "\n".join(body).strip())
    core = "\n\n\n".join(parts)
    for a, b in ALIASES:
        core = core.replace(a, b)
    core += "\n" + READER

    app = open("app.py").read().replace("import wincross as wx\n", "")
    for fn in EXPORTS:
        app = app.replace(f"wx.{fn}", fn)
    head, rest = app.split("st.set_page_config", 1)
    head = head.rstrip().replace(
        "import io\n\nimport streamlit as st",
        "import io\nimport re\n\nimport openpyxl\nimport streamlit as st")

    out = (head + "\n\n\n" + core + "\n\n\n"
           + "# " + "=" * 68 + "\n# Streamlit interface\n# " + "=" * 68 + "\n\n"
           + "st.set_page_config" + rest)
    open("streamlit_app.py", "w").write(out)
    import ast
    ast.parse(out)
    return len(out.split("\n"))


if __name__ == "__main__":
    print(f"streamlit_app.py rebuilt: {build()} lines")
