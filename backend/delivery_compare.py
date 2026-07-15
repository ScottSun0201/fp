#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compare system delivery tracking numbers with courier statements.

This script intentionally uses only the Python standard library, so it can run
on a clean macOS Python install without pandas/openpyxl.
"""

from __future__ import annotations

import argparse
import csv
import posixpath
import re
import sys
import zipfile
from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"a": MAIN_NS, "r": REL_NS}
PKG_NS = {"rel": PKG_REL_NS}

SYSTEM_TRACKING_COL = "logistics_number"
SYSTEM_COMPANY_COL = "logistics_company"
DEFAULT_SYSTEM_DATE_COL = "created_at"

SYSTEM_OUTPUT_COLUMNS = [
    "id",
    "created_at",
    "transaction_time",
    "platform",
    "shop",
    "order_number",
    "system_number",
    "title",
    "product_info",
    "quantity",
    "logistics_number",
    "logistics_company",
    "receiver_name",
    "receiver_mobile",
    "customer_address",
    "receiver_address",
    "order_status",
    "amount",
]

RETURN_TRACKING_COLUMNS = ("logistics_number", "return_tracking_number")

RETURN_OUTPUT_COLUMNS = [
    "id",
    "created_at",
    "updated_at",
    "platform",
    "shop",
    "product_info",
    "order_number",
    "serial_number_mark",
    "logistics_number",
    "logistics_company",
    "quantity",
    "product_status",
    "reason",
    "application_time",
    "return_tracking_number",
    "customer_address",
    "delivery_time",
    "return_logistics_company",
]


@dataclass(frozen=True)
class CourierConfig:
    name: str
    file_name: str
    company_names: Tuple[str, ...]
    start_date: str
    end_date: str
    tracking_headers: Tuple[str, ...]


@dataclass(frozen=True)
class SupplementConfig:
    file_name: str
    tracking_headers: Tuple[str, ...]


COURIERS = (
    CourierConfig(
        name="极兔",
        file_name="26极兔.xlsx",
        company_names=("极兔速递", "极兔快递"),
        start_date="2026-05-01",
        end_date="2026-06-01",
        tracking_headers=("运单编号", "运单号"),
    ),
    CourierConfig(
        name="中通",
        file_name="26淘品中通.xlsx",
        company_names=("中通快递",),
        start_date="2026-01-01",
        end_date="2026-06-01",
        tracking_headers=("运单号", "运单编号"),
    ),
    CourierConfig(
        name="顺丰",
        file_name="26顺丰.xlsx",
        company_names=("顺丰速运",),
        start_date="2026-05-01",
        end_date="2026-06-01",
        tracking_headers=("运单号码", "运单号", "运单编号"),
    ),
)

SUPPLEMENTS = (
    SupplementConfig(file_name="2月发货在线汇总表.xlsx", tracking_headers=("快递单号",)),
    SupplementConfig(file_name="5月骊威发货汇总表.xlsx", tracking_headers=("快递单号",)),
)


def normalize_tracking(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.strip().replace("\u3000", "")
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    text = text.upper()
    match = re.search(r"(JT\d+|SF\d+|\d{10,})", text)
    if match:
        return match.group(1)
    return text.upper()


def looks_like_tracking(value: str) -> bool:
    if not value or not any(char.isdigit() for char in value):
        return False
    if value in {"没货", "无", "暂无", "空"}:
        return False
    return True


def infer_courier_from_tracking(tracking: str) -> str:
    if tracking.startswith("JT"):
        return "极兔"
    if tracking.startswith("SF"):
        return "顺丰"
    if tracking and tracking[0].isdigit():
        return "中通"
    return "无法判断"


def infer_return_courier(row: dict, tracking_column: str, tracking: str) -> str:
    if tracking_column == "return_tracking_number":
        company = (row.get("return_logistics_company") or row.get("logistics_company") or "").strip()
    else:
        company = (row.get("logistics_company") or row.get("return_logistics_company") or "").strip()

    if "中通" in company:
        return "中通"
    if "极兔" in company:
        return "极兔"
    if "顺丰" in company:
        return "顺丰"
    if tracking.startswith("JT"):
        return "极兔"
    if tracking.startswith("SF"):
        return "顺丰"
    return "无法判断"


def parse_system_datetime(value: str) -> datetime | None:
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%d/%m/%Y %H:%M:%S.%f", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return None


def parse_iso_date(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%d")


def excel_serial_to_datetime(value: str) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value
    if number < 1 or number > 80000:
        return value
    # Excel's 1900 date system includes the fake 1900-02-29 day. Using 1899-12-30
    # matches how Excel-compatible tools normally interpret serial dates.
    dt = datetime(1899, 12, 30) + timedelta(days=number)
    if abs(number - int(number)) < 0.0000001:
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def col_to_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 0
    result = 0
    for char in match.group(1):
        result = result * 26 + ord(char) - 64
    return result - 1


def index_to_col(index: int) -> str:
    index += 1
    chars = []
    while index:
        index, rem = divmod(index - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    values = []
    for si in root.findall("a:si", NS):
        texts = [node.text or "" for node in si.findall(".//a:t", NS)]
        values.append("".join(texts))
    return values


def get_sheet_paths(zf: zipfile.ZipFile) -> List[Tuple[str, str]]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("rel:Relationship", PKG_NS)
    }

    sheets = []
    for sheet in workbook.find("a:sheets", NS):
        rid = sheet.attrib[f"{{{REL_NS}}}id"]
        target = rid_to_target[rid]
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = posixpath.normpath(posixpath.join("xl", target))
        sheets.append((sheet.attrib["name"], path))
    return sheets


def get_cell_value(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//a:t", NS))

    value_node = cell.find("a:v", NS)
    if value_node is None:
        return ""
    value = value_node.text or ""
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return value
    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"
    return value


def read_xlsx_rows(path: Path) -> Iterable[Tuple[str, int, List[str]]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = load_shared_strings(zf)
        for sheet_name, sheet_path in get_sheet_paths(zf):
            root = ET.fromstring(zf.read(sheet_path))
            for row in root.findall(".//a:sheetData/a:row", NS):
                row_index = int(float(row.attrib.get("r", "0") or 0))
                cells = []
                max_col = -1
                for cell in row.findall("a:c", NS):
                    index = col_to_index(cell.attrib.get("r", "A1"))
                    value = get_cell_value(cell, shared_strings)
                    cells.append((index, value))
                    max_col = max(max_col, index)
                values = [""] * (max_col + 1)
                for index, value in cells:
                    values[index] = value
                yield sheet_name, row_index, values


def unique_headers(headers: Sequence[str]) -> List[str]:
    counts: Dict[str, int] = defaultdict(int)
    result = []
    for i, header in enumerate(headers, start=1):
        name = (header or "").strip() or f"空列{i}"
        counts[name] += 1
        if counts[name] > 1:
            name = f"{name}_{counts[name]}"
        result.append(name)
    return result


def read_courier_records(path: Path, tracking_headers: Sequence[str]) -> Tuple[List[dict], List[str]]:
    records: List[dict] = []
    all_headers = OrderedDict()
    current_headers_by_sheet: Dict[str, Tuple[int, List[str], str]] = {}

    for sheet_name, row_number, values in read_xlsx_rows(path):
        stripped = [(value or "").strip() for value in values]
        tracking_header = next((h for h in tracking_headers if h in stripped), "")
        if tracking_header:
            headers = unique_headers(stripped)
            current_headers_by_sheet[sheet_name] = (row_number, headers, tracking_header)
            for header in headers:
                all_headers.setdefault(header, None)
            continue

        if sheet_name not in current_headers_by_sheet:
            continue
        header_row, headers, tracking_header = current_headers_by_sheet[sheet_name]
        if row_number <= header_row:
            continue

        data = {header: (stripped[i] if i < len(stripped) else "") for i, header in enumerate(headers)}
        tracking = normalize_tracking(data.get(tracking_header, ""))
        if not tracking:
            continue

        for key in list(data):
            if "时间" in key and data[key]:
                data[key] = excel_serial_to_datetime(data[key])

        data["_sheet"] = sheet_name
        data["_row"] = str(row_number)
        data["_tracking"] = tracking
        records.append(data)

    return records, list(all_headers.keys())


def read_supplement_records(path: Path, tracking_headers: Sequence[str]) -> Tuple[List[dict], List[str]]:
    records: List[dict] = []
    all_headers = OrderedDict()
    current_headers_by_sheet: Dict[str, Tuple[int, List[str], str]] = {}

    for sheet_name, row_number, values in read_xlsx_rows(path):
        stripped = [(value or "").strip() for value in values]
        tracking_header = next((h for h in tracking_headers if h in stripped), "")
        if tracking_header:
            headers = unique_headers(stripped)
            current_headers_by_sheet[sheet_name] = (row_number, headers, tracking_header)
            for header in headers:
                all_headers.setdefault(header, None)
            continue

        if sheet_name not in current_headers_by_sheet:
            continue
        header_row, headers, tracking_header = current_headers_by_sheet[sheet_name]
        if row_number <= header_row:
            continue

        data = {header: (stripped[i] if i < len(stripped) else "") for i, header in enumerate(headers)}
        tracking = normalize_tracking(data.get(tracking_header, ""))
        if not looks_like_tracking(tracking):
            continue

        for key in list(data):
            if ("时间" in key or "日期" in key) and data[key]:
                data[key] = excel_serial_to_datetime(data[key])

        data["_source_file"] = path.name
        data["_sheet"] = sheet_name
        data["_row"] = str(row_number)
        data["_tracking"] = tracking
        records.append(data)

    return records, list(all_headers.keys())


def read_all_supplements(paths: Sequence[Path]) -> Tuple[List[dict], List[str]]:
    records: List[dict] = []
    headers = OrderedDict()
    config_by_name = {config.file_name: config for config in SUPPLEMENTS}
    for path in paths:
        config = config_by_name.get(path.name, SupplementConfig(path.name, ("快递单号",)))
        file_records, file_headers = read_supplement_records(path, config.tracking_headers)
        records.extend(file_records)
        for header in file_headers:
            headers.setdefault(header, None)
    return records, list(headers.keys())


def read_system_records(
    path: Path,
    company_names: Sequence[str],
    start_date: str,
    end_date: str,
    date_column: str,
) -> Tuple[List[dict], List[str]]:
    start = parse_iso_date(start_date)
    end = parse_iso_date(end_date)
    wanted_companies = set(company_names)
    records: List[dict] = []

    with path.open(newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            return [], []
        if SYSTEM_TRACKING_COL not in reader.fieldnames:
            raise ValueError(f"系统表缺少字段：{SYSTEM_TRACKING_COL}")
        if SYSTEM_COMPANY_COL not in reader.fieldnames:
            raise ValueError(f"系统表缺少字段：{SYSTEM_COMPANY_COL}")
        if date_column not in reader.fieldnames:
            raise ValueError(f"系统表缺少日期字段：{date_column}")

        for row_number, row in enumerate(reader, start=2):
            if (row.get(SYSTEM_COMPANY_COL) or "").strip() not in wanted_companies:
                continue
            dt = parse_system_datetime(row.get(date_column, ""))
            if dt is None or not (start <= dt < end):
                continue
            tracking = normalize_tracking(row.get(SYSTEM_TRACKING_COL, ""))
            if not tracking:
                continue
            row["_row"] = str(row_number)
            row["_tracking"] = tracking
            records.append(row)
        return records, reader.fieldnames


def read_return_records(path: Path) -> Tuple[List[dict], List[str]]:
    records: List[dict] = []
    with path.open(newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            return [], []

        for row_number, row in enumerate(reader, start=2):
            for tracking_column in RETURN_TRACKING_COLUMNS:
                tracking = normalize_tracking(row.get(tracking_column, ""))
                if not looks_like_tracking(tracking):
                    continue
                data = dict(row)
                data["_row"] = str(row_number)
                data["_tracking"] = tracking
                data["_tracking_column"] = tracking_column
                data["_courier"] = infer_return_courier(row, tracking_column, tracking)
                records.append(data)
        return records, reader.fieldnames


def group_by_tracking(records: Sequence[dict]) -> Dict[str, List[dict]]:
    grouped: Dict[str, List[dict]] = defaultdict(list)
    for record in records:
        grouped[record["_tracking"]].append(record)
    return grouped


def prefixed_row(prefix: str, record: dict, columns: Sequence[str]) -> Dict[str, str]:
    return {f"{prefix}{col}": record.get(col, "") for col in columns}


def build_system_detail_rows(records: Sequence[dict], columns: Sequence[str]) -> List[dict]:
    rows = []
    for record in records:
        row = {"物流单号": record["_tracking"], "系统CSV行号": record.get("_row", "")}
        for column in columns:
            row[column] = record.get(column, "")
        rows.append(row)
    return rows


def build_courier_detail_rows(records: Sequence[dict], columns: Sequence[str]) -> List[dict]:
    rows = []
    for record in records:
        row = {
            "物流单号": record["_tracking"],
            "快递Sheet": record.get("_sheet", ""),
            "快递Excel行号": record.get("_row", ""),
        }
        for column in columns:
            row[column] = record.get(column, "")
        rows.append(row)
    return rows


def build_supplement_detail_rows(records: Sequence[dict], columns: Sequence[str]) -> List[dict]:
    rows = []
    for record in records:
        row = {
            "物流单号": record["_tracking"],
            "补充文件": record.get("_source_file", ""),
            "补充Sheet": record.get("_sheet", ""),
            "补充Excel行号": record.get("_row", ""),
        }
        for column in columns:
            row[column] = record.get(column, "")
        rows.append(row)
    return rows


def build_matched_rows(
    system_records: Sequence[dict],
    courier_group: Dict[str, List[dict]],
    system_columns: Sequence[str],
    courier_columns: Sequence[str],
) -> List[dict]:
    rows = []
    for system_record in system_records:
        tracking = system_record["_tracking"]
        courier_records = courier_group.get(tracking, [])
        first_courier = courier_records[0] if courier_records else {}
        row = {
            "物流单号": tracking,
            "系统CSV行号": system_record.get("_row", ""),
            "快递记录数": str(len(courier_records)),
            "快递Sheet": first_courier.get("_sheet", ""),
            "快递Excel行号": first_courier.get("_row", ""),
        }
        row.update(prefixed_row("系统_", system_record, system_columns))
        row.update(prefixed_row("快递_", first_courier, courier_columns))
        rows.append(row)
    return rows


def build_supplement_match_rows(
    supplement_numbers: Sequence[str],
    system_group: Dict[str, List[dict]],
    courier_group: Dict[str, List[dict]],
    supplement_group: Dict[str, List[dict]],
    system_columns: Sequence[str],
    courier_columns: Sequence[str],
    supplement_columns: Sequence[str],
    matched_numbers: set[str],
    only_system_numbers: set[str],
    only_courier_numbers: set[str],
) -> List[dict]:
    rows = []
    for tracking in sorted(supplement_numbers):
        system_records = system_group.get(tracking, [])
        courier_records = courier_group.get(tracking, [])
        supplement_records = supplement_group.get(tracking, [])
        first_system = system_records[0] if system_records else {}
        first_courier = courier_records[0] if courier_records else {}
        first_supplement = supplement_records[0] if supplement_records else {}
        if tracking in matched_numbers:
            original_status = "原已匹配成功"
        elif tracking in only_system_numbers:
            original_status = "原仅系统"
        elif tracking in only_courier_numbers:
            original_status = "原仅快递"
        else:
            original_status = "原不在当前快递对比范围"
        row = {
            "物流单号": tracking,
            "原匹配状态": original_status,
            "系统记录数": str(len(system_records)),
            "快递记录数": str(len(courier_records)),
            "补充记录数": str(len(supplement_records)),
            "系统CSV行号": first_system.get("_row", ""),
            "快递Sheet": first_courier.get("_sheet", ""),
            "快递Excel行号": first_courier.get("_row", ""),
            "补充文件": first_supplement.get("_source_file", ""),
            "补充Sheet": first_supplement.get("_sheet", ""),
            "补充Excel行号": first_supplement.get("_row", ""),
        }
        row.update(prefixed_row("系统_", first_system, system_columns))
        row.update(prefixed_row("快递_", first_courier, courier_columns))
        row.update(prefixed_row("补充_", first_supplement, supplement_columns))
        rows.append(row)
    return rows


def build_supplement_status_rows(
    supplement_group: Dict[str, List[dict]],
    courier_groups: Dict[str, Dict[str, List[dict]]],
    supplement_columns: Sequence[str],
    wanted_status: str,
) -> List[dict]:
    rows = []
    for tracking in sorted(supplement_group):
        courier_name = infer_courier_from_tracking(tracking)
        courier_records = courier_groups.get(courier_name, {}).get(tracking, [])
        status = "匹配成功" if courier_records else "未匹配"
        if status != wanted_status:
            continue

        supplement_records = supplement_group[tracking]
        first_supplement = supplement_records[0]
        first_courier = courier_records[0] if courier_records else {}
        row = {
            "物流单号": tracking,
            "判断快递": courier_name,
            "匹配状态": status,
            "补充记录数": str(len(supplement_records)),
            "快递账单记录数": str(len(courier_records)),
            "补充文件": first_supplement.get("_source_file", ""),
            "补充Sheet": first_supplement.get("_sheet", ""),
            "补充Excel行号": first_supplement.get("_row", ""),
            "快递Sheet": first_courier.get("_sheet", ""),
            "快递Excel行号": first_courier.get("_row", ""),
        }
        row.update(prefixed_row("补充_", first_supplement, supplement_columns))
        rows.append(row)
    return rows


def build_return_status_rows(
    return_records: Sequence[dict],
    courier_groups: Dict[str, Dict[str, List[dict]]],
    return_columns: Sequence[str],
    wanted_status: str,
) -> List[dict]:
    rows = []
    for record in return_records:
        tracking = record["_tracking"]
        courier_name = record["_courier"]
        if courier_name not in courier_groups:
            continue
        courier_records = courier_groups.get(courier_name, {}).get(tracking, [])
        status = "匹配成功" if courier_records else "未匹配"
        if status != wanted_status:
            continue

        first_courier = courier_records[0] if courier_records else {}
        row = {
            "物流单号": tracking,
            "匹配字段": record.get("_tracking_column", ""),
            "判断快递": courier_name,
            "匹配状态": status,
            "退换货CSV行号": record.get("_row", ""),
            "快递账单记录数": str(len(courier_records)),
            "快递Sheet": first_courier.get("_sheet", ""),
            "快递Excel行号": first_courier.get("_row", ""),
        }
        for column in return_columns:
            row[f"退换货_{column}"] = record.get(column, "")
        rows.append(row)
    return rows


def ordered_columns(rows: Sequence[dict]) -> List[str]:
    columns = OrderedDict()
    for row in rows:
        for key in row:
            columns.setdefault(key, None)
    return list(columns.keys())


def rows_to_table(rows: Sequence[dict], preferred_columns: Sequence[str] | None = None) -> List[List[str]]:
    if preferred_columns is None:
        columns = ordered_columns(rows)
    else:
        columns = list(preferred_columns)
        for key in ordered_columns(rows):
            if key not in columns:
                columns.append(key)
    table = [columns]
    for row in rows:
        table.append([str(row.get(column, "")) for column in columns])
    return table


def xml_escape_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def safe_sheet_name(name: str, used: set[str]) -> str:
    clean = re.sub(r"[\[\]:*?/\\]", "_", name)[:31] or "Sheet"
    base = clean
    index = 2
    while clean in used:
        suffix = f"_{index}"
        clean = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1
    used.add(clean)
    return clean


def write_sheet_xml(handle, rows: Sequence[Sequence[str]]) -> None:
    handle.write(
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<sheetData>"
    )
    for row_index, row in enumerate(rows, start=1):
        handle.write(f'<row r="{row_index}">'.encode("utf-8"))
        for col_index, value in enumerate(row):
            cell_ref = f"{index_to_col(col_index)}{row_index}"
            text = xml_escape_text(value)
            preserve = ' xml:space="preserve"' if text != text.strip() else ""
            handle.write(
                f'<c r="{cell_ref}" t="inlineStr"><is><t{preserve}>{text}</t></is></c>'.encode(
                    "utf-8"
                )
            )
        handle.write(b"</row>")
    handle.write(b"</sheetData></worksheet>")


def write_xlsx(path: Path, sheets: Sequence[Tuple[str, Sequence[Sequence[str]]]]) -> None:
    used_names: set[str] = set()
    safe_sheets = [(safe_sheet_name(name, used_names), rows) for name, rows in sheets]

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        overrides = [
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        ]
        for index in range(1, len(safe_sheets) + 1):
            overrides.append(
                f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(overrides)
            + "</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL_NS}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )

        workbook_sheets = []
        workbook_rels = []
        for index, (sheet_name, _) in enumerate(safe_sheets, start=1):
            workbook_sheets.append(
                f'<sheet name="{xml_escape_text(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>'
            )
            workbook_rels.append(
                f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
            )
        zf.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets>'
            + "".join(workbook_sheets)
            + "</sheets></workbook>",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL_NS}">'
            + "".join(workbook_rels)
            + "</Relationships>",
        )
        for index, (_, rows) in enumerate(safe_sheets, start=1):
            with zf.open(f"xl/worksheets/sheet{index}.xml", "w") as handle:
                write_sheet_xml(handle, rows)


def compare_one(
    config: CourierConfig,
    base_dir: Path,
    delivery_file: Path,
    courier_file: Path,
    supplement_group: Dict[str, List[dict]],
    supplement_headers: Sequence[str],
    system_date_column: str,
) -> Tuple[List[List[str]], List[Tuple[str, List[List[str]]]]]:
    system_records, system_headers = read_system_records(
        delivery_file,
        config.company_names,
        config.start_date,
        config.end_date,
        system_date_column,
    )
    courier_records, courier_headers = read_courier_records(courier_file, config.tracking_headers)

    system_group = group_by_tracking(system_records)
    courier_group = group_by_tracking(courier_records)
    system_numbers = set(system_group)
    courier_numbers = set(courier_group)
    matched_numbers = system_numbers & courier_numbers
    only_system_numbers = system_numbers - courier_numbers
    only_courier_numbers = courier_numbers - system_numbers
    supplement_numbers = set(supplement_group)
    supplement_only_system_numbers = only_system_numbers & supplement_numbers
    supplement_only_courier_numbers = only_courier_numbers & supplement_numbers
    supplement_courier_numbers = courier_numbers & supplement_numbers
    supplement_courier_already_matched_numbers = supplement_courier_numbers & matched_numbers
    remaining_only_system_numbers = only_system_numbers - supplement_numbers
    remaining_only_courier_numbers = only_courier_numbers - supplement_numbers

    system_columns = [col for col in SYSTEM_OUTPUT_COLUMNS if col in system_headers]
    courier_columns = courier_headers
    supplement_columns = supplement_headers

    matched_system_records = [
        record for record in system_records if record["_tracking"] in matched_numbers
    ]
    only_system_records = [
        record for record in system_records if record["_tracking"] in remaining_only_system_numbers
    ]
    only_courier_records = [
        record for record in courier_records if record["_tracking"] in remaining_only_courier_numbers
    ]
    supplement_only_system_records = [
        record for record in system_records if record["_tracking"] in supplement_only_system_numbers
    ]
    supplement_only_courier_records = [
        record for record in courier_records if record["_tracking"] in supplement_only_courier_numbers
    ]

    summary = [
        config.name,
        courier_file.name,
        "、".join(config.company_names),
        f"{config.start_date} <= {system_date_column} < {config.end_date}",
        str(len(system_records)),
        str(len(system_numbers)),
        str(len(courier_records)),
        str(len(courier_numbers)),
        str(len(matched_numbers)),
        str(len(matched_system_records)),
        str(len([r for r in courier_records if r["_tracking"] in matched_numbers])),
        str(len(only_system_numbers)),
        str(len([r for r in system_records if r["_tracking"] in only_system_numbers])),
        str(len(only_courier_numbers)),
        str(len([r for r in courier_records if r["_tracking"] in only_courier_numbers])),
        str(len(supplement_only_system_numbers)),
        str(len(supplement_only_system_records)),
        str(len(supplement_only_courier_numbers)),
        str(len(supplement_only_courier_records)),
        str(len(remaining_only_system_numbers)),
        str(len(only_system_records)),
        str(len(remaining_only_courier_numbers)),
        str(len(only_courier_records)),
        "不限制",
        str(len(supplement_courier_numbers)),
        str(len(supplement_courier_already_matched_numbers)),
    ]

    matched_rows = build_matched_rows(
        matched_system_records, courier_group, system_columns, courier_columns
    )
    only_system_rows = build_system_detail_rows(only_system_records, system_columns)
    only_courier_rows = build_courier_detail_rows(only_courier_records, courier_columns)

    matched_columns = [
        "物流单号",
        "系统CSV行号",
        "快递记录数",
        "快递Sheet",
        "快递Excel行号",
    ] + [f"系统_{col}" for col in system_columns] + [f"快递_{col}" for col in courier_columns]
    supplement_columns_out = [
        "物流单号",
        "原匹配状态",
        "系统记录数",
        "快递记录数",
        "补充记录数",
        "系统CSV行号",
        "快递Sheet",
        "快递Excel行号",
        "补充文件",
        "补充Sheet",
        "补充Excel行号",
    ] + [f"系统_{col}" for col in system_columns] + [
        f"快递_{col}" for col in courier_columns
    ] + [f"补充_{col}" for col in supplement_columns]
    only_system_columns = ["物流单号", "系统CSV行号"] + system_columns
    only_courier_columns = ["物流单号", "快递Sheet", "快递Excel行号"] + courier_columns

    sheets = [
        (f"{config.name}_匹配成功", rows_to_table(matched_rows, matched_columns)),
        (f"{config.name}_仅系统", rows_to_table(only_system_rows, only_system_columns)),
        (f"{config.name}_仅快递", rows_to_table(only_courier_rows, only_courier_columns)),
    ]
    return summary, sheets


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="对比系统发货单号和快递公司账单单号")
    parser.add_argument("--base-dir", default=Path(__file__).resolve().parent, type=Path)
    parser.add_argument("--delivery", default="delivery.csv", help="系统导出的 CSV 文件名或路径")
    parser.add_argument("--jitu", default="26极兔.xlsx", help="极兔 Excel 文件名或路径")
    parser.add_argument("--zhongtong", default="26淘品中通.xlsx", help="中通 Excel 文件名或路径")
    parser.add_argument("--shunfeng", default="26顺丰.xlsx", help="顺丰 Excel 文件名或路径")
    parser.add_argument("--supplement-feb", default="2月发货在线汇总表.xlsx", help="2月补充发货表")
    parser.add_argument("--supplement-may", default="5月骊威发货汇总表.xlsx", help="5月补充发货表")
    parser.add_argument("--returns", default="returnandexchangestop.csv", help="退换货 CSV 文件名或路径")
    parser.add_argument(
        "--system-date-column",
        default=DEFAULT_SYSTEM_DATE_COL,
        help="系统表用于筛选时间段的日期字段，默认 created_at；也可用 transaction_time",
    )
    parser.add_argument("--output", default="", help="输出 xlsx 文件路径；默认自动生成")
    return parser


def resolve_path(base_dir: Path, value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return base_dir / path


def main(argv: Sequence[str] | None = None) -> int:
    args = build_arg_parser().parse_args(argv)
    base_dir = args.base_dir.resolve()
    delivery_file = resolve_path(base_dir, args.delivery)
    courier_files = {
        "极兔": resolve_path(base_dir, args.jitu),
        "中通": resolve_path(base_dir, args.zhongtong),
        "顺丰": resolve_path(base_dir, args.shunfeng),
    }
    supplement_files = [
        resolve_path(base_dir, args.supplement_feb),
        resolve_path(base_dir, args.supplement_may),
    ]
    return_file = resolve_path(base_dir, args.returns)

    for path in [delivery_file, *courier_files.values(), *supplement_files, return_file]:
        if not path.exists():
            print(f"文件不存在：{path}", file=sys.stderr)
            return 2

    output = Path(args.output) if args.output else base_dir / f"快递对比结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    if not output.is_absolute():
        output = base_dir / output

    summary_header = [
        "快递",
        "快递文件",
        "系统快递公司名称",
        "系统筛选时间",
        "系统记录数",
        "系统唯一单号数",
        "快递记录数",
        "快递唯一单号数",
        "匹配成功唯一单号数",
        "匹配成功系统记录数",
        "匹配成功快递记录数",
        "原仅系统唯一单号数",
        "原仅系统记录数",
        "原仅快递唯一单号数",
        "原仅快递记录数",
        "补充命中仅系统唯一单号数",
        "补充命中仅系统记录数",
        "补充命中仅快递唯一单号数",
        "补充命中仅快递记录数",
        "补充后仍仅系统唯一单号数",
        "补充后仍仅系统记录数",
        "补充后仍仅快递唯一单号数",
        "补充后仍仅快递记录数",
        "补充表时间限制",
        "补充表命中快递账单唯一单号数",
        "其中原已匹配成功唯一单号数",
    ]
    summary_rows = [summary_header]
    result_sheets: List[Tuple[str, List[List[str]]]] = []
    supplement_records, supplement_headers = read_all_supplements(supplement_files)
    supplement_group = group_by_tracking(supplement_records)
    return_records, return_headers = read_return_records(return_file)
    courier_groups = {}
    for config in COURIERS:
        courier_records, _ = read_courier_records(courier_files[config.name], config.tracking_headers)
        courier_groups[config.name] = group_by_tracking(courier_records)

    for config in COURIERS:
        summary, sheets = compare_one(
            config,
            base_dir,
            delivery_file,
            courier_files[config.name],
            supplement_group,
            supplement_headers,
            args.system_date_column,
        )
        summary_rows.append(summary)
        result_sheets.extend(sheets)

    supplement_status_columns = [
        "物流单号",
        "判断快递",
        "匹配状态",
        "补充记录数",
        "快递账单记录数",
        "补充文件",
        "补充Sheet",
        "补充Excel行号",
        "快递Sheet",
        "快递Excel行号",
    ] + [f"补充_{col}" for col in supplement_headers]
    supplement_unmatched_rows = build_supplement_status_rows(
        supplement_group, courier_groups, supplement_headers, "未匹配"
    )
    supplement_matched_rows = build_supplement_status_rows(
        supplement_group, courier_groups, supplement_headers, "匹配成功"
    )
    result_sheets.extend(
        [
            ("补充_未匹配", rows_to_table(supplement_unmatched_rows, supplement_status_columns)),
            ("补充_匹配成功", rows_to_table(supplement_matched_rows, supplement_status_columns)),
        ]
    )

    return_columns = [col for col in RETURN_OUTPUT_COLUMNS if col in return_headers]
    return_status_columns = [
        "物流单号",
        "匹配字段",
        "判断快递",
        "匹配状态",
        "退换货CSV行号",
        "快递账单记录数",
        "快递Sheet",
        "快递Excel行号",
    ] + [f"退换货_{col}" for col in return_columns]
    return_unmatched_rows = build_return_status_rows(
        return_records, courier_groups, return_columns, "未匹配"
    )
    return_matched_rows = build_return_status_rows(
        return_records, courier_groups, return_columns, "匹配成功"
    )
    result_sheets.extend(
        [
            ("退换货_未匹配", rows_to_table(return_unmatched_rows, return_status_columns)),
            ("退换货_匹配成功", rows_to_table(return_matched_rows, return_status_columns)),
        ]
    )

    write_xlsx(output, [("汇总", summary_rows), *result_sheets])

    print(f"对比完成：{output}")
    print("")
    for row in summary_rows[1:]:
        print(
            f"{row[0]}：匹配成功唯一单号 {row[8]}，补充命中唯一单号 {int(row[15]) + int(row[17])}，"
            f"补充表命中快递账单 {row[24]}，补充后仍仅系统 {row[19]}，补充后仍仅快递 {row[21]}"
        )
    print("")
    print(f"系统日期字段：{args.system_date_column}")
    print("补充表时间限制：不限制")
    print(f"补充表唯一单号数：{len(supplement_group)}")
    print(f"补充表匹配成功唯一单号数：{len(supplement_matched_rows)}")
    print(f"补充表未匹配唯一单号数：{len(supplement_unmatched_rows)}")
    print(f"退换货匹配成功记录数：{len(return_matched_rows)}")
    print(f"退换货未匹配记录数：{len(return_unmatched_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
