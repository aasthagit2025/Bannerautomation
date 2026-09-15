"""
WinCross Banner Generator - Streamlit app.

Upload a banner specification workbook, review the validation output,
and download a WinCross banner file.

Run locally:   streamlit run app.py
"""

import io

import streamlit as st

import wincross as wx


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
    overrides = wx.parse_width_overrides(width_text)
except ValueError as exc:
    st.error(f"Could not read width overrides: {exc}")
    st.stop()

try:
    points, meta = wx.read_points(
        uploaded, default_width=int(default_width), width_overrides=overrides
    )
except wx.SheetError as exc:
    st.error(f"Could not read the sheet: {exc}")
    st.stop()
except Exception as exc:  # noqa: BLE001 - surface any reader failure to the user
    st.error(f"Unexpected problem reading the workbook: {exc}")
    st.stop()

st.success(
    f"Read {meta['columns']} banner columns from sheet '{meta['sheet']}' "
    f"(headings on rows {meta['rows']['super']} and {meta['rows']['group']}, "
    f"labels on row {meta['rows']['label']}, logic on row {meta['rows']['logic']})."
)

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

text, warnings, errors, stats = wx.emit(points, settings)

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

tab_file, tab_cols, tab_header = st.tabs(
    ["Banner file", "Column map", "Header preview"]
)

with tab_file:
    st.download_button(
        "Download banner file", data=text.encode("utf-8"),
        file_name="banner.txt", mime="text/plain",
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
