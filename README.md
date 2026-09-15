# WinCross Banner Generator

Generate WinCross banner files from a banner specification workbook.

Banner coding is repetitive but unforgiving: the column widths, the stat
letters, the header rule lines and the logic expressions all have to agree
with each other, and a single mismatch shifts everything to its right. This
tool takes the specification sheet that already exists at the start of a
study and produces the banner file from it, so the parts that can be derived
are derived rather than typed.

## What it produces

A complete banner file:

- **Header directives** — `SW`, `HP`, `CP`, `SL`, `ST`, `WT`, `OP`, `BT`,
  `BF`, `XL`, `PT`, sized from the column count
- **Logic lines** — one per banner column, with spacing normalised
- **Header text block** — the tiered rule lines and centred labels

## Input format

The reader takes the last four populated rows of the sheet:

| Row | Contents |
|---|---|
| super-header | merged across the columns it spans |
| group heading | merged across the columns it spans |
| column label | one per column |
| banner logic | one per column, in WinCross syntax |

Merged cell ranges define the header spans, so headings are read from the
sheet rather than inferred. A line break inside a label becomes a separate
header line. The leftmost blank column is ignored.

Two real specification sheets are in `examples/`.

## Usage

### Streamlit

```bash
pip install -r requirements.txt
streamlit run app.py
```

Upload a workbook, set widths and spacing in the sidebar, review the
warnings, download the banner file.

### Command line

```bash
python cli.py examples/banner_spec.xlsx -o banner.txt --widths 1:20,2:20,6:20
```

`--strict` exits non-zero when any warning is raised, which makes it usable
as a pre-delivery check in CI.

## WinCross constraints enforced

From Setup | Banners | Edit Banner, Width and Spacing tab:

| Rule | Behaviour |
|---|---|
| Spaces before column ≤ 5 | Error, blocks generation |
| Same spacing on every column | Enforced by design — held as one value |
| Divider characters ≤ spaces before | Error, blocks generation |
| Default column width 10 | Applied when unspecified |
| Hidden columns | Excluded from report width |

The uniform-spacing rule matters most. The job file format physically allows
a different value per column, but editing the banner in the GUI afterwards
raises a warning. Holding it as a single value makes a non-compliant banner
impossible to express rather than merely caught after the fact.

## Checks raised

- A label whose longest word exceeds its column width (will truncate)
- The same label carrying different widths in parallel blocks (will
  misalign) — this catches a real defect found in a shipped banner
- Two columns with identical logic — they report the same base and can never
  test significant against each other
- Columns missing a label or missing logic

These are warnings, not errors. Some are deliberate: a total column
restated under a second super-header legitimately duplicates logic.

## Header block layout

A rule spanning columns *a* through *b* is the sum of those column widths
plus the spacers between them. Labels are placed within their span and wrap
on line breaks.

WinCross enters banner text left-justified by default; centre and right are
explicit choices under the Cells menu. The generator defaults to centre
because the sample banners are centred throughout, and the sidebar can
change it.

## Known gaps

Two details were derived from sample files rather than from documentation,
and should be confirmed against Preview Banner before the output is trusted
unreviewed:

- **Left stub width**, currently 1 character. If the whole header block sits
  a fixed distance off, this is the setting.
- **Centring bias** when a label cannot centre evenly; the generator biases
  left. If individual labels sit one character off, this is the cause.

`HP` is emitted as all `1`s, matching both sample files. It carries one
value per column and is suspected to encode cell justification, but this is
untested — centre-justifying one column in WinCross and re-reading the
saved file would settle it.

`OP` and `XL` are passed through as constants. If either needs to vary by
study, that is not yet handled.

## Layout

```
app.py                  Streamlit interface
cli.py                  command-line entry point
wincross/
  excel_spec.py         workbook reader
  generator.py          directive and logic assembly
  header_block.py       header text block renderer
  width_spacing.py      width and spacing rules and validation
examples/               real specification workbooks
```

## Deploying

Push to GitHub, then on [share.streamlit.io](https://share.streamlit.io)
point a new app at the repository with `app.py` as the entry point.
Dependencies come from `requirements.txt`. No secrets or credentials are
needed — nothing leaves the process.

Note that uploaded workbooks are processed in memory by the hosted app. If
the specifications are client-confidential, run it locally or deploy
somewhere private instead.
