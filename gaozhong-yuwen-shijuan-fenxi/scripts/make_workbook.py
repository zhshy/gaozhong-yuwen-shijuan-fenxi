#!/usr/bin/env python3
"""把试卷编码结果与学生得分数据生成 xlsx 数据表。

仅使用标准库，不联网，不改动输入文件。
xlsx 为 zip 包结构，此处直接按 OOXML 规范写出，无需第三方库。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

MIN_PYTHON = (3, 10)

SHEET_ENCODING = "逐题编码表"
SHEET_STUDENT = "学生得分对照表"
SHEET_SPEC = "双向细目表"

INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
COLUMN_NAMES = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def column_ref(index: int) -> str:
    """把 0 基列号转为 Excel 列名，支持超过 26 列。"""
    if index < 0:
        raise ValueError("列号不能为负")
    name = ""
    current = index
    while True:
        name = COLUMN_NAMES[current % 26] + name
        current = current // 26 - 1
        if current < 0:
            break
    return name


def cell_ref(row: int, col: int) -> str:
    return f"{column_ref(col)}{row}"


def is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def xml_text(value) -> str:
    return escape("" if value is None else str(value))


def build_sheet(rows: list[list], freeze_header: bool = True) -> str:
    """按 OOXML 规范构造一张工作表。"""
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ]
    if freeze_header and rows:
        parts.append(
            '<sheetViews><sheetView workbookViewId="0">'
            f'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews>'
        )
    parts.append("<sheetData>")
    for row_index, row in enumerate(rows, start=1):
        parts.append(f'<row r="{row_index}">')
        for col_index, value in enumerate(row):
            ref = cell_ref(row_index, col_index)
            if is_number(value):
                parts.append(f'<c r="{ref}"><v>{value}</v></c>')
            else:
                text = xml_text(value)
                if text == "":
                    continue
                parts.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData>")
    parts.append("</worksheet>")
    return "".join(parts)


def build_workbook(sheets: list[tuple[str, str]]) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheets>",
    ]
    for index, (name, _) in enumerate(sheets, start=1):
        safe = INVALID_SHEET_CHARS.sub("", name)[:31] or f"sheet{index}"
        parts.append(f'<sheet name="{xml_text(safe)}" sheetId="{index}" r:id="rId{index}"/>')
    parts.append("</sheets></workbook>")
    return "".join(parts)


def build_workbook_rels(count: int) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for index in range(1, count + 1):
        parts.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    return "".join(parts) + "</Relationships>"


def build_content_types(count: int) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    ]
    for index in range(1, count + 1):
        parts.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return "".join(parts) + "</Types>"


def build_root_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def write_xlsx(path: Path, sheets: list[tuple[str, list[list]]]) -> None:
    rendered = [(name, build_sheet(rows)) for name, rows in sheets]
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", build_content_types(len(rendered)))
        archive.writestr("_rels/.rels", build_root_rels())
        archive.writestr("xl/workbook.xml", build_workbook(rendered))
        archive.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(len(rendered)))
        for index, (_, xml) in enumerate(rendered, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", xml)


def load_payload(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict):
        raise ValueError("输入 JSON 根节点必须是对象")
    return data


def dictionary(value, default=None):
    return value if isinstance(value, dict) else (default if default is not None else {})


def as_list(value) -> list:
    return value if isinstance(value, list) else []


def rows_from_encoding(items: list) -> list[list]:
    header = ["题目编号", "卷别", "年份", "板块", "文本类型", "体裁", "题型", "考点",
              "指令动词", "认知层级", "情境类型", "实践活动", "分值", "评分标准形态", "教材关联"]
    rows = [header]
    for item in items:
        entry = dictionary(item)
        rows.append([
            entry.get("题目编号", ""),
            entry.get("卷别", ""),
            entry.get("年份", ""),
            entry.get("板块", ""),
            entry.get("文本类型", ""),
            entry.get("体裁", ""),
            entry.get("题型", ""),
            entry.get("考点", ""),
            entry.get("指令动词", ""),
            entry.get("认知层级", ""),
            entry.get("情境类型", ""),
            entry.get("实践活动", ""),
            entry.get("分值", ""),
            entry.get("评分标准形态", ""),
            entry.get("教材关联", ""),
        ])
    return rows


def rows_from_students(students: list) -> list[list]:
    header = ["学生标识", "题目编号", "考点", "满分", "得分", "得分率", "失分机理", "证据等级"]
    rows = [header]
    for student in students:
        entry = dictionary(student)
        name = entry.get("学生标识", "")
        for record in as_list(entry.get("作答记录")):
            item = dictionary(record)
            full = item.get("满分", "")
            score = item.get("得分", "")
            rate = ""
            if is_number(full) and is_number(score) and full:
                rate = round(float(score) / float(full) * 100, 1)
            rows.append([name, item.get("题目编号", ""), item.get("考点", ""), full, score, rate,
                         item.get("失分机理", ""), item.get("证据等级", "")])
    return rows


def rows_from_matrix(matrix: list) -> list[list]:
    levels = ["识别确认", "理解转述", "推断判断", "分析综合", "鉴赏评价", "探究创新"]
    header = ["内容板块"] + levels + ["合计"]
    rows = [header]
    for row in matrix:
        entry = dictionary(row)
        values = [entry.get(level, "") for level in levels]
        total = sum(value for value in values if is_number(value))
        rows.append([entry.get("内容板块", "")] + values + [total if total else ""])
    return rows


def main() -> int:
    parser = argparse.ArgumentParser(description="生成试卷分析数据表 xlsx")
    parser.add_argument("--input", type=Path, required=True, help="编码结果 JSON 路径")
    parser.add_argument("--output", type=Path, required=True, help="输出 xlsx 路径")
    parser.add_argument("--overwrite", action="store_true", help="允许覆盖已存在的输出文件")
    args = parser.parse_args()

    if sys.version_info[:2] < MIN_PYTHON:
        print(f"Python 版本过低，需要 {MIN_PYTHON[0]}.{MIN_PYTHON[1]} 及以上", file=sys.stderr)
        return 2

    if not args.input.is_file():
        print(f"输入文件不存在：{args.input}", file=sys.stderr)
        return 1

    if args.output.exists() and not args.overwrite:
        print(f"输出文件已存在，未覆盖：{args.output}", file=sys.stderr)
        print("如需覆盖请加 --overwrite", file=sys.stderr)
        return 1

    try:
        payload = load_payload(args.input)
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
        print(f"输入 JSON 无法解析：{exc}", file=sys.stderr)
        return 1

    sheets = [
        (SHEET_ENCODING, rows_from_encoding(as_list(payload.get("逐题编码")))),
        (SHEET_STUDENT, rows_from_students(as_list(payload.get("学生数据")))),
        (SHEET_SPEC, rows_from_matrix(as_list(payload.get("双向细目表")))),
    ]

    if all(len(rows) <= 1 for _, rows in sheets):
        print("输入 JSON 没有任何可写入的数据行", file=sys.stderr)
        return 1

    try:
        write_xlsx(args.output, sheets)
    except (OSError, zipfile.BadZipFile, ValueError) as exc:
        print(f"数据表生成失败：{exc}", file=sys.stderr)
        print("请改用 Markdown 表格输出", file=sys.stderr)
        return 1

    total_rows = sum(len(rows) - 1 for _, rows in sheets)
    print(f"已生成：{args.output}")
    print(f"工作表：{SHEET_ENCODING}、{SHEET_STUDENT}、{SHEET_SPEC}")
    print(f"数据行合计：{total_rows}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
