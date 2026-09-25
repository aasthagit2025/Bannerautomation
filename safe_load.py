"""
Open a workbook that openpyxl would otherwise refuse.

Files written by some tools carry stylesheet attributes openpyxl does not
recognise, and loading them raises a TypeError before any data is read:

    CellStyle.__init__() got an unexpected keyword argument 'applyColorFormat'

The data is fine; only the styling is unusual. This retries by stripping the
offending attributes from a copy of the file. Nothing the generator does
depends on cell styling, so nothing is lost.
"""

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
