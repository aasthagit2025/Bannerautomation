#!/usr/bin/env python3
"""Command-line generation, for batch runs and CI checks.

    python cli.py spec.xlsx -o banner.txt --widths 1:20,2:20,6:20
"""
import argparse
import sys

import wincross as wx


def main():
    ap = argparse.ArgumentParser(description="Generate a WinCross banner file.")
    ap.add_argument("workbook", help="banner specification .xlsx")
    ap.add_argument("-o", "--out", default="banner.txt")
    ap.add_argument("--widths", default="", help="overrides, e.g. 1:20,2:20,6:20")
    ap.add_argument("--default-width", type=int, default=10)
    ap.add_argument("--spaces", type=int, default=1)
    ap.add_argument("--no-header", action="store_true",
                    help="omit the header text block")
    ap.add_argument("--strict", action="store_true",
                    help="exit non-zero if any warnings are raised")
    args = ap.parse_args()

    points, meta = wx.read_points(
        args.workbook,
        default_width=args.default_width,
        width_overrides=wx.parse_width_overrides(args.widths),
    )
    text, warnings, errors, stats = wx.emit(points, {
        "spaces_before": args.spaces,
        "emit_header_block": not args.no_header,
    })

    for msg in warnings:
        print(f"WARNING: {msg}", file=sys.stderr)
    for msg in errors:
        print(f"ERROR:   {msg}", file=sys.stderr)
    if errors:
        return 2

    with open(args.out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(f"{stats['columns']} columns -> {args.out} "
          f"(report width {stats['report_width']} chars)", file=sys.stderr)

    return 1 if (args.strict and warnings) else 0


if __name__ == "__main__":
    sys.exit(main())
