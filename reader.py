"""
Detect which banner format a workbook uses and read it accordingly.

Two shapes occur in practice, one per client:

  grid  - the banner runs left to right, one spreadsheet column per banner
          column, with heading rows above the labels and a row of WinCross
          logic beneath. Read by excel_spec.

  plan  - the banner runs top to bottom, one spreadsheet row per banner
          column, under named headings (Column, Variable, Group, Label,
          Response, Condition). Conditions are in plan syntax and need
          translating. One sheet may hold several banners. Read by
          plan_reader.

Detection looks for the plan's heading row. Anything without one is treated
as a grid, which is the older and simpler shape.
"""

import openpyxl

from . import excel_spec, plan_reader


def detect(path_or_buffer):
    """Return 'plan' or 'grid'."""
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True)
    for name in wb.sheetnames:
        if plan_reader.looks_like_plan(wb[name]):
            return "plan"
    return "grid"


def read_any(path_or_buffer, default_width=10, width_overrides=None,
             lo=0, hi=9999):
    """Read either format.

    Returns (banners, report, fmt) where `banners` maps a banner name to its
    list of points. A grid workbook yields a single banner; a plan workbook
    yields one per table found. `report` is the translation report, empty
    for grid workbooks since their logic is already WinCross syntax.
    """
    fmt = detect(path_or_buffer)

    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)

    if fmt == "plan":
        banners, report = plan_reader.read_plan(
            path_or_buffer, default_width=default_width, lo=lo, hi=hi)
        if width_overrides:
            for points in banners.values():
                for p in points:
                    if p["column"] in width_overrides:
                        p["width"] = int(width_overrides[p["column"]])
        return banners, report, fmt

    points, meta = excel_spec.read_points(
        path_or_buffer, default_width=default_width,
        width_overrides=width_overrides)
    name = meta.get("sheet", "Banner")
    return {name: points}, [], fmt
