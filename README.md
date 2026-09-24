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

## Two input formats

The tool detects which shape a workbook uses and reads it accordingly.

### Grid format

The banner runs left to right, one spreadsheet column per banner column.
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

Heading rows are optional and counted from the sheet, so a banner with one
heading tier works as well as one with two.

### Plan format

The banner runs top to bottom, one spreadsheet row per banner column, under
named headings:

| Column | Variable | Group | Label | Response | Condition | N |
|---|---|---|---|---|---|---|
| 1 | Total | 1 | | Total | All respondents | 135 |
| 2 | S0 | 2 | Geography | US | S0=1 | 75 |
| 3 | | | | EUR | S0=2,3,4 OR 5 | 60 |

Group headings come from the Label column, merged down the rows they cover.
One sheet can hold several banners, each introduced by a title row such as
"Banner 2: US (S0=1)"; the app lets you pick which to generate.

Conditions here are written for people rather than for WinCross, so they are
translated. See below.

Real sheets in both formats are in `examples/`.

## Condition translation

Plan conditions are translated into WinCross expressions, and every one is
reported with a status:

| Status | Meaning |
|---|---|
| `ok` | translated with no assumptions |
| `assumed` | translated, but a range bound had to be supplied |
| `blocked` | cannot be translated; excluded from the output |

Examples:

| Condition | WinCross | Status |
|---|---|---|
| `S0=1` | `S0(1)` | ok |
| `S8=1 OR 3` | `S8(1,3)` | ok |
| `S0=2,3,4 OR 5` | `S0(2,3,4,5)` | ok |
| `S4>14` | `S4(15-9999)` | assumed |
| `S5r4>XX` | — | blocked |
| `All respondents` | — | blocked |

`blocked` is the point of it. `XX` is an unfilled placeholder: the client has
not decided the cut point. Emitting a plausible number there would put a
wrong column into a deliverable, so the column is left out and reported
instead.

Open-ended comparisons need both ends of a range in WinCross. The assumed
bounds default to 0 and 9999 and are set in the sidebar; tightening them to
the real limits of the variable is worth doing.

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
streamlit_app.py        single-file build of the whole app, for deployment
wincross/
  reader.py             format detection, dispatches to the right reader
  excel_spec.py         grid-format reader
  plan_reader.py        plan-format reader, handles multiple banners
  logic_translate.py    plan condition -> WinCross expression
  generator.py          directive and logic assembly
  header_block.py       header text block renderer
  width_spacing.py      width and spacing rules and validation
examples/               real specification workbooks
```

## Deploying

Push `streamlit_app.py` and `requirements.txt` to the repository root, then
on [share.streamlit.io](https://share.streamlit.io) point a new app at it.
`streamlit_app.py` is a self-contained build of everything in `wincross/`,
so no other files are needed and there is no folder structure to get wrong.

Rebuild it after changing anything under `wincross/`; the package version is
the one to edit.
Dependencies come from `requirements.txt`. No secrets or credentials are
needed — nothing leaves the process.

Note that uploaded workbooks are processed in memory by the hosted app. If
the specifications are client-confidential, run it locally or deploy
somewhere private instead.
