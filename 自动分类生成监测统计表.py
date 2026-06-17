#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动分类生成监测次数统计表。

使用方式：
1. 把每一期成果表放在各期文件夹中，例如 001期、002期、059期。
2. 双击“运行自动分类监测统计.bat”，或运行：
   python 自动分类生成监测统计表.py
3. 程序会自动生成：
   - 监测次数统计表_自动生成.xlsx
   - 监测次数统计表_自动生成明细.json

自动规则：
- 只统计可见工作表；隐藏工作表视为旧表或辅助表，不计入。
- 普通项目按工作表名自动分类：
  “地表沉降”“地表沉降(2)”“地表沉降（3）”都会归为“地表沉降”。
- 普通项目点数：主表 A 列“点号”去重。
- 普通项目次数：本周累计变量区内有实际数据的天数，多张同类表取最大值，不累加。
- 测斜项目：默认“工作表名像点号，且表内 A 列有深度”的工作表为测斜。
  例如 ZQT01、ZQT02、CX01 等会统一归到“测斜”。
- 测斜点数：可见点号工作表去重。
- 测斜次数：测斜表本周累计变量区内有实际数据的天数，多张表取最大值。
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {
    "main": MAIN_NS,
    "rel": REL_NS,
    "officeRel": OFFICE_REL_NS,
}
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", OFFICE_REL_NS)

POINT_HEADER = "点号"
DEPTH_KEY = "深度"
WEEK_CHANGE_KEY = "本周"
INVALID_STRINGS = {"", "/", "\\", "-", "--", "---", "----"}

DEFAULT_OUTPUT = "监测次数统计表_自动生成.xlsx"
DEFAULT_REPORT = "监测次数统计表_自动生成明细.json"

PHASE_RE = re.compile(r"^(\d+)")
CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")
SHEET_SUFFIX_RE = re.compile(r"\s*[\(（]\s*\d+\s*[\)）]\s*$")
DEFAULT_SHEET_RE = re.compile(r"^sheet\d*$", re.I)
POINT_SHEET_RE = re.compile(r"^[A-Za-z]{1,10}[-_]?\d+[A-Za-z0-9_-]*$")

READ_ROW_LIMIT = 200
READ_COL_LIMIT = 40


def col_to_num(col: str) -> int:
    value = 0
    for ch in col.upper():
        value = value * 26 + ord(ch) - 64
    return value


def num_to_col(num: int) -> str:
    chars = []
    while num:
        num, rem = divmod(num - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


def cell_coord(ref: str) -> tuple[int | None, int | None]:
    m = CELL_REF_RE.match(ref)
    if not m:
        return None, None
    return int(m.group(2)), col_to_num(m.group(1))


def valid_value(value: object) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() not in INVALID_STRINGS
    return True


def read_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("main:si", NS):
        strings.append("".join((t.text or "") for t in si.findall(".//main:t", NS)))
    return strings


def read_workbook_sheets(zf: ZipFile) -> list[dict[str, str]]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map: dict[str, str] = {}
    for rel in rels.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if not rid:
            continue
        rel_map[rid] = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")

    sheets: list[dict[str, str]] = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        rid = sheet.attrib.get(f"{{{OFFICE_REL_NS}}}id")
        path = rel_map.get(rid or "")
        if path:
            sheets.append(
                {
                    "name": sheet.attrib.get("name", ""),
                    "path": path,
                    "state": sheet.attrib.get("state", "visible"),
                }
            )
    return sheets


def read_sheet_cells(zf: ZipFile, sheet_path: str, shared: list[str]) -> dict[tuple[int, int], object]:
    cells: dict[tuple[int, int], object] = {}
    try:
        root = ET.fromstring(zf.read(sheet_path))
    except KeyError:
        return cells

    for cell in root.findall(".//main:sheetData/main:row/main:c", NS):
        row, col = cell_coord(cell.attrib.get("r", ""))
        if row is None or row > READ_ROW_LIMIT or col > READ_COL_LIMIT:
            continue

        cell_type = cell.attrib.get("t")
        value: object | None = None
        if cell_type == "s":
            v = cell.find("main:v", NS)
            if v is not None and v.text is not None:
                try:
                    value = shared[int(v.text)]
                except Exception:
                    value = v.text
        elif cell_type == "inlineStr":
            value = "".join((t.text or "") for t in cell.findall(".//main:t", NS))
        else:
            v = cell.find("main:v", NS)
            if v is not None and v.text is not None:
                text = v.text
                if cell_type in {"str", "e"}:
                    value = text
                else:
                    try:
                        num = float(text)
                        value = int(num) if num.is_integer() else num
                    except Exception:
                        value = text
        if value is not None:
            cells[(row, col)] = value
    return cells


def normalize_category_name(sheet_name: str) -> str:
    name = sheet_name.strip()
    while True:
        new_name = SHEET_SUFFIX_RE.sub("", name).strip()
        if new_name == name:
            return name
        name = new_name


def is_default_or_summary_sheet(sheet_name: str) -> bool:
    compact = sheet_name.strip()
    if not compact:
        return True
    if compact == "统计表":
        return True
    if DEFAULT_SHEET_RE.match(compact):
        return True
    return False


def find_point_header_row(cells: dict[tuple[int, int], object]) -> int | None:
    for row in range(1, 20):
        for col in range(1, 5):
            if cells.get((row, col)) == POINT_HEADER:
                return row
    return None


def find_depth_header_row(cells: dict[tuple[int, int], object]) -> int | None:
    for row in range(1, 20):
        value = cells.get((row, 1))
        if isinstance(value, str) and DEPTH_KEY in value:
            return row
    return None


def ordinary_data_columns(cells: dict[tuple[int, int], object], header_row: int) -> range:
    point_col = 1
    start_col = point_col + 3
    stop_col = None
    for col in range(start_col + 1, READ_COL_LIMIT + 1):
        value = cells.get((header_row, col))
        if isinstance(value, str) and WEEK_CHANGE_KEY in value:
            stop_col = col
            break
    if stop_col is None:
        stop_col = start_col + 7
    return range(start_col, stop_col)


def ordinary_sheet_stats(cells: dict[tuple[int, int], object]) -> tuple[set[str], int] | None:
    header_row = find_point_header_row(cells)
    if header_row is None:
        return None

    data_cols = ordinary_data_columns(cells, header_row)
    points: set[str] = set()
    day_cols: set[int] = set()
    for row in range(header_row + 2, READ_ROW_LIMIT + 1):
        point = cells.get((row, 1))
        if not valid_value(point):
            break
        row_has_data = False
        for col in data_cols:
            if valid_value(cells.get((row, col))):
                row_has_data = True
                day_cols.add(col)
        if row_has_data:
            points.add(str(point).strip())
    return points, len(day_cols)


def is_point_named_inclinometer(sheet_name: str, cells: dict[tuple[int, int], object]) -> str | None:
    compact = sheet_name.strip().replace(" ", "").upper()
    if DEFAULT_SHEET_RE.match(compact):
        return None
    if not POINT_SHEET_RE.match(compact):
        return None
    if find_depth_header_row(cells) is None:
        return None
    return compact


def inclinometer_data_columns(cells: dict[tuple[int, int], object], header_row: int) -> range:
    start_col = 3
    stop_col = None
    for col in range(start_col + 1, READ_COL_LIMIT + 1):
        value = cells.get((header_row, col))
        if isinstance(value, str) and WEEK_CHANGE_KEY in value:
            stop_col = col
            break
    if stop_col is None:
        stop_col = start_col + 7
    return range(start_col, stop_col)


def inclinometer_sheet_days(cells: dict[tuple[int, int], object]) -> int:
    header_row = find_depth_header_row(cells)
    if header_row is None:
        return 0
    data_cols = inclinometer_data_columns(cells, header_row)
    day_cols: set[int] = set()
    seen_depth = False
    for row in range(header_row + 2, READ_ROW_LIMIT + 1):
        depth = cells.get((row, 1))
        if not valid_value(depth):
            if seen_depth:
                break
            continue
        seen_depth = True
        for col in data_cols:
            if valid_value(cells.get((row, col))):
                day_cols.add(col)
    return len(day_cols)


def discover_phase_workbooks(root: Path) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    rows: list[dict[str, object]] = []
    warnings: list[dict[str, object]] = []

    phase_dirs = [p for p in root.iterdir() if p.is_dir() and PHASE_RE.match(p.name)]
    if phase_dirs:
        for folder in sorted(phase_dirs, key=lambda p: p.name):
            phase = int(PHASE_RE.match(folder.name).group(1))  # type: ignore[union-attr]
            files = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
            if not files:
                warnings.append({"phase": phase, "folder": folder.name, "warning": "no xlsx"})
                continue
            if len(files) > 1:
                warnings.append(
                    {
                        "phase": phase,
                        "folder": folder.name,
                        "warning": "multiple xlsx; using first",
                        "files": [p.name for p in files],
                    }
                )
            rows.append({"phase": phase, "folder": folder.name, "file": files[0]})
        return rows, warnings

    for file in sorted(p for p in root.glob("*.xlsx") if not p.name.startswith("~$")):
        m = PHASE_RE.match(file.stem)
        if not m:
            continue
        rows.append({"phase": int(m.group(1)), "folder": "", "file": file})
    return rows, warnings


def collect_stats(root: Path) -> dict[str, object]:
    phase_files, warnings = discover_phase_workbooks(root)
    category_order: list[str] = []
    saw_inclinometer = False
    result_rows: list[dict[str, object]] = []

    for item in sorted(phase_files, key=lambda x: int(x["phase"])):
        source = Path(item["file"])  # type: ignore[arg-type]
        phase_stats: dict[str, dict[str, object]] = {}
        hidden_skipped: list[str] = []

        with ZipFile(source) as zf:
            shared = read_shared_strings(zf)
            for sheet in read_workbook_sheets(zf):
                sheet_name = sheet["name"]
                if sheet["state"] != "visible":
                    hidden_skipped.append(sheet_name)
                    continue
                if is_default_or_summary_sheet(sheet_name):
                    continue

                cells = read_sheet_cells(zf, sheet["path"], shared)

                point_name = is_point_named_inclinometer(sheet_name, cells)
                if point_name:
                    saw_inclinometer = True
                    stats = phase_stats.setdefault(
                        "测斜",
                        {"points_set": set(), "times": 0, "sheets": [], "type": "inclinometer"},
                    )
                    stats["points_set"].add(point_name)  # type: ignore[union-attr]
                    stats["times"] = max(int(stats["times"]), inclinometer_sheet_days(cells))
                    stats["sheets"].append(sheet_name)  # type: ignore[union-attr]
                    continue

                ordinary = ordinary_sheet_stats(cells)
                if ordinary is None:
                    continue
                points, times = ordinary
                if not points:
                    continue

                category = normalize_category_name(sheet_name)
                if category not in category_order:
                    category_order.append(category)
                stats = phase_stats.setdefault(
                    category,
                    {"points_set": set(), "times": 0, "sheets": [], "type": "ordinary"},
                )
                stats["points_set"].update(points)  # type: ignore[union-attr]
                stats["times"] = max(int(stats["times"]), times)
                stats["sheets"].append(sheet_name)  # type: ignore[union-attr]

        cleaned_stats = {}
        for category, data in phase_stats.items():
            point_count = len(data["points_set"])  # type: ignore[arg-type]
            cleaned_stats[category] = {
                "points": point_count if point_count else None,
                "times": data["times"] if point_count else None,
                "sheets": data["sheets"],
                "type": data["type"],
            }

        result_rows.append(
            {
                "phase": int(item["phase"]),
                "folder": item["folder"],
                "file": source.name,
                "stats": cleaned_stats,
                "hidden_skipped": hidden_skipped,
            }
        )

    if saw_inclinometer and "测斜" not in category_order:
        category_order.append("测斜")
    elif saw_inclinometer:
        category_order = [c for c in category_order if c != "测斜"] + ["测斜"]

    return {
        "root": str(root),
        "categories": category_order,
        "rows": result_rows,
        "warnings": warnings,
    }


def append_text_cell(row_el: ET.Element, row: int, col: int, text: str, style: int = 0) -> None:
    cell = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", {"r": f"{num_to_col(col)}{row}", "t": "inlineStr"})
    if style:
        cell.attrib["s"] = str(style)
    inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    t = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
    t.text = text


def append_number_cell(row_el: ET.Element, row: int, col: int, value: int | None, style: int = 0) -> None:
    cell = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", {"r": f"{num_to_col(col)}{row}"})
    if style:
        cell.attrib["s"] = str(style)
    if value is not None:
        v = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
        v.text = str(value)


def build_sheet_xml(stats: dict[str, object]) -> bytes:
    categories: list[str] = stats["categories"]  # type: ignore[assignment]
    rows: list[dict[str, object]] = stats["rows"]  # type: ignore[assignment]
    max_col = 1 + len(categories) * 3
    max_row = 2 + len(rows)

    worksheet = ET.Element(f"{{{MAIN_NS}}}worksheet")
    dimension = ET.SubElement(worksheet, f"{{{MAIN_NS}}}dimension", {"ref": f"A1:{num_to_col(max_col)}{max_row}"})
    dimension.text = None

    sheet_views = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetViews")
    sheet_view = ET.SubElement(sheet_views, f"{{{MAIN_NS}}}sheetView", {"workbookViewId": "0"})
    ET.SubElement(
        sheet_view,
        f"{{{MAIN_NS}}}pane",
        {"ySplit": "2", "topLeftCell": "A3", "activePane": "bottomLeft", "state": "frozen"},
    )
    ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetFormatPr", {"defaultRowHeight": "18"})

    cols = ET.SubElement(worksheet, f"{{{MAIN_NS}}}cols")
    ET.SubElement(cols, f"{{{MAIN_NS}}}col", {"min": "1", "max": "1", "width": "12", "customWidth": "1"})
    for idx in range(2, max_col + 1):
        width = "4" if (idx - 2) % 3 == 2 else "12"
        ET.SubElement(cols, f"{{{MAIN_NS}}}col", {"min": str(idx), "max": str(idx), "width": width, "customWidth": "1"})

    sheet_data = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetData")

    row1 = ET.SubElement(sheet_data, f"{{{MAIN_NS}}}row", {"r": "1"})
    append_text_cell(row1, 1, 1, "", 1)
    for index, category in enumerate(categories):
        col = 2 + index * 3
        append_text_cell(row1, 1, col, category, 1)
        append_text_cell(row1, 1, col + 1, "", 1)
        append_text_cell(row1, 1, col + 2, "", 0)

    row2 = ET.SubElement(sheet_data, f"{{{MAIN_NS}}}row", {"r": "2"})
    append_text_cell(row2, 2, 1, "期数", 2)
    for index in range(len(categories)):
        col = 2 + index * 3
        append_text_cell(row2, 2, col, "点数", 2)
        append_text_cell(row2, 2, col + 1, "次数", 2)
        append_text_cell(row2, 2, col + 2, "", 0)

    for row_index, item in enumerate(rows, start=3):
        row_el = ET.SubElement(sheet_data, f"{{{MAIN_NS}}}row", {"r": str(row_index)})
        append_text_cell(row_el, row_index, 1, f"第{item['phase']}期", 3)
        row_stats: dict[str, dict[str, object]] = item["stats"]  # type: ignore[assignment]
        for cat_index, category in enumerate(categories):
            col = 2 + cat_index * 3
            cat_stats = row_stats.get(category, {})
            append_number_cell(row_el, row_index, col, cat_stats.get("points"), 3)  # type: ignore[arg-type]
            append_number_cell(row_el, row_index, col + 1, cat_stats.get("times"), 3)  # type: ignore[arg-type]
            append_text_cell(row_el, row_index, col + 2, "", 0)

    merge_cells = ET.SubElement(worksheet, f"{{{MAIN_NS}}}mergeCells", {"count": str(len(categories))})
    for index in range(len(categories)):
        col = 2 + index * 3
        ET.SubElement(merge_cells, f"{{{MAIN_NS}}}mergeCell", {"ref": f"{num_to_col(col)}1:{num_to_col(col + 1)}1"})

    ET.SubElement(
        worksheet,
        f"{{{MAIN_NS}}}pageMargins",
        {"left": "0.7", "right": "0.7", "top": "0.75", "bottom": "0.75", "header": "0.3", "footer": "0.3"},
    )
    return ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def styles_xml() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"""


def write_xlsx(output_file: Path, stats: dict[str, object]) -> None:
    now = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    sheet_xml = build_sheet_xml(stats)
    with ZipFile(output_file, "w", ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>""",
        )
        zf.writestr(
            "_rels/.rels",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""",
        )
        zf.writestr(
            "docProps/core.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>监测统计工具</dc:creator>
  <cp:lastModifiedBy>监测统计工具</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>""",
        )
        zf.writestr(
            "docProps/app.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>监测统计工具</Application>
</Properties>""",
        )
        zf.writestr(
            "xl/workbook.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{MAIN_NS}" xmlns:r="{OFFICE_REL_NS}">
  <sheets><sheet name="统计表" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        )
        zf.writestr("xl/styles.xml", styles_xml())
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="自动分类生成监测次数统计表")
    parser.add_argument("root", nargs="?", default=None, help="项目根目录，默认使用脚本所在目录")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help=f"输出 xlsx 文件名，默认：{DEFAULT_OUTPUT}")
    parser.add_argument("--report", default=DEFAULT_REPORT, help=f"统计明细 JSON 文件名，默认：{DEFAULT_REPORT}")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root = Path(args.root).resolve() if args.root else Path(__file__).resolve().parent
    if not root.exists():
        print(f"目录不存在：{root}", file=sys.stderr)
        return 1

    stats = collect_stats(root)
    if not stats["rows"]:  # type: ignore[index]
        print("没有找到每期成果表。请确认存在 001期、002期 等文件夹，或根目录内有按期号开头的 xlsx。", file=sys.stderr)
        return 1

    output_file = root / args.output
    report_file = root / args.report
    write_xlsx(output_file, stats)
    report_file.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"完成：处理 {len(stats['rows'])} 期。")
    print("自动分类：", "、".join(stats["categories"]))  # type: ignore[index]
    phase59 = next((row for row in stats["rows"] if row["phase"] == 59), None)  # type: ignore[index]
    if phase59:
        row_stats = phase59["stats"]
        if "地表沉降" in row_stats:
            item = row_stats["地表沉降"]
            print(f"第59期地表沉降：{item['points']} 点 / {item['times']} 次")
        if "测斜" in row_stats:
            item = row_stats["测斜"]
            print(f"第59期测斜：{item['points']} 点 / {item['times']} 次")
    if stats["warnings"]:  # type: ignore[index]
        print(f"注意：发现 {len(stats['warnings'])} 条警告，详见 {report_file.name}")
    print(f"输出文件：{output_file}")
    print(f"明细文件：{report_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
