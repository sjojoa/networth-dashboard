"""Minimal .xlsx reader — standard library only.

An .xlsx is a zip of XML. Reading it directly (instead of via openpyxl) means the
dashboard launcher works with any python3 on the machine, including Apple's
/usr/bin/python3, which has no third-party packages.

Exposes just what the dashboard needs: cell values and the fill color of a cell
(the workbook's legend encodes account ownership as header fill colors).
"""

import datetime
import re
import zipfile
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
EPOCH = datetime.date(1899, 12, 30)  # Excel's day 0 (1900 date system)

# Built-in number formats that mean "this number is a date".
BUILTIN_DATE_FMTS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))


def col_to_index(ref):
    """'C7' -> 3 (1-based column number)."""
    letters = re.match(r"([A-Z]+)", ref).group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def row_of(ref):
    return int(re.search(r"(\d+)", ref).group(1))


class Cell:
    __slots__ = ("value", "fill")

    def __init__(self, value, fill):
        self.value = value
        self.fill = fill  # None, or a hashable key identifying the solid fill


class Sheet:
    def __init__(self, name, cells, max_row, max_col):
        self.name = name
        self._cells = cells  # {(row, col): Cell}
        self.max_row = max_row
        self.max_col = max_col

    def cell(self, row, col):
        return self._cells.get((row, col), Cell(None, None))

    def row(self, row):
        return [self.cell(row, c) for c in range(1, self.max_col + 1)]


def _shared_strings(zf):
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall("m:si", NS):
        # Text may be split across runs; concatenate every <t> under the item.
        out.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    return out


def _styles(zf):
    """Return (fill_key_per_style_index, is_date_per_style_index)."""
    root = ET.fromstring(zf.read("xl/styles.xml"))

    fills = []
    for fill in root.find("m:fills", NS) or []:
        pat = fill.find("m:patternFill", NS)
        key = None
        if pat is not None and pat.get("patternType") == "solid":
            fg = pat.find("m:fgColor", NS)
            if fg is not None:
                if fg.get("theme") is not None:
                    key = ("theme", fg.get("theme"), round(float(fg.get("tint") or 0), 3))
                elif fg.get("rgb") is not None:
                    key = ("rgb", fg.get("rgb"), round(float(fg.get("tint") or 0), 3))
        fills.append(key)

    custom_date_fmts = set()
    numfmts = root.find("m:numFmts", NS)
    if numfmts is not None:
        for nf in numfmts.findall("m:numFmt", NS):
            code = (nf.get("formatCode") or "").lower()
            stripped = re.sub(r"\[[^\]]*\]|\"[^\"]*\"", "", code)
            if re.search(r"[dmyh]", stripped):
                custom_date_fmts.add(int(nf.get("numFmtId")))

    fill_by_style, date_by_style = [], []
    for xf in root.find("m:cellXfs", NS) or []:
        fill_id = int(xf.get("fillId") or 0)
        fill_by_style.append(fills[fill_id] if fill_id < len(fills) else None)
        fmt_id = int(xf.get("numFmtId") or 0)
        date_by_style.append(fmt_id in BUILTIN_DATE_FMTS or fmt_id in custom_date_fmts)
    return fill_by_style, date_by_style


def _sheet_paths(zf):
    """[(sheet name, zip path)] in workbook order."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = {rel.get("Id"): rel.get("Target") for rel in rels.findall("pr:Relationship", NS)}

    out = []
    for sheet in wb.find("m:sheets", NS):
        path = target[sheet.get("{%s}id" % NS["r"])]
        if not path.startswith("/"):
            path = "xl/" + path.lstrip("./")
        out.append((sheet.get("name"), path.lstrip("/")))
    return out


def load(path):
    """Read a workbook and return [Sheet] in workbook order."""
    with zipfile.ZipFile(path) as zf:
        strings = _shared_strings(zf)
        fill_by_style, date_by_style = _styles(zf)

        sheets = []
        for name, sheet_path in _sheet_paths(zf):
            root = ET.fromstring(zf.read(sheet_path))
            cells, max_row, max_col = {}, 0, 0

            for c in root.iter("{%s}c" % NS["m"]):
                ref = c.get("r")
                row, col = row_of(ref), col_to_index(ref)
                style = int(c.get("s") or 0)
                ctype = c.get("t") or "n"

                v = c.find("m:v", NS)
                if ctype == "inlineStr":
                    is_el = c.find("m:is", NS)
                    text = "".join(t.text or "" for t in is_el.iter("{%s}t" % NS["m"])) if is_el is not None else None
                    value = text
                elif v is None or v.text is None:
                    value = None
                elif ctype == "s":
                    value = strings[int(v.text)]
                elif ctype in ("str", "e"):
                    value = v.text
                elif ctype == "b":
                    value = v.text == "1"
                else:
                    value = float(v.text)
                    if value.is_integer() and abs(value) < 1e15:
                        value = int(value)
                    if style < len(date_by_style) and date_by_style[style]:
                        value = EPOCH + datetime.timedelta(days=float(v.text))

                fill = fill_by_style[style] if style < len(fill_by_style) else None
                if value is not None or fill is not None:
                    cells[(row, col)] = Cell(value, fill)
                    max_row, max_col = max(max_row, row), max(max_col, col)

            sheets.append(Sheet(name, cells, max_row, max_col))
    return sheets
