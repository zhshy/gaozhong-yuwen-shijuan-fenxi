#!/usr/bin/env python3
"""只读检查试卷与答卷文件的可用性，确定证据等级。

不修改任何文件，不联网，不读取环境变量。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

MIN_PYTHON = (3, 10)

TEXT_SUFFIXES = {".txt", ".md"}
DOC_SUFFIXES = {".docx"}
PDF_SUFFIXES = {".pdf"}
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
SHEET_SUFFIXES = {".xlsx", ".csv"}

PAPER_SUFFIXES = TEXT_SUFFIXES | DOC_SUFFIXES | PDF_SUFFIXES | IMAGE_SUFFIXES
ANSWER_SUFFIXES = PAPER_SUFFIXES | SHEET_SUFFIXES

SCORING_HINTS = ("评分细则", "评分标准", "给分标准", "评分参考", "采分点", "得分点")
REFERENCE_HINTS = ("参考答案", "答案", "解析")


def is_python_supported() -> tuple[bool, str]:
    current = sys.version_info[:2]
    if current < MIN_PYTHON:
        return False, f"当前 Python {current[0]}.{current[1]}，需要 {MIN_PYTHON[0]}.{MIN_PYTHON[1]} 及以上"
    return True, f"Python {current[0]}.{current[1]} 满足要求"


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return ""


def has_text_layer(path: Path) -> tuple[bool, str]:
    if path.suffix.lower() != ".pdf":
        return True, "非 PDF，跳过文字层检查"
    raw = path.read_bytes()[:4]
    if raw != b"%PDF":
        return False, "文件头不是 PDF，可能已损坏"
    size = path.stat().st_size
    if size < 2048:
        return False, "PDF 体积过小，可能为空白或损坏"
    try:
        data = path.read_bytes()
    except OSError:
        return False, "无法读取 PDF 内容"
    if b"/Font" not in data and b"/Image" not in data:
        return False, "未检出文字层或图像资源，可能为空文档"
    return True, "PDF 可读"


def classify(path: Path, role: str) -> dict:
    item = {
        "path": str(path),
        "name": path.name,
        "role": role,
        "exists": path.is_file(),
        "readable": False,
        "suffix": path.suffix.lower(),
        "size_bytes": 0,
        "status": "unavailable",
        "reason": "",
        "notes": [],
    }
    if not item["exists"]:
        item["reason"] = "文件不存在"
        return item
    if not path.is_file():
        item["reason"] = "路径不是文件"
        return item

    allowed = PAPER_SUFFIXES if role == "paper" else ANSWER_SUFFIXES
    if item["suffix"] not in allowed:
        item["reason"] = f"不支持的格式 {item['suffix'] or '(无扩展名)'}"
        return item

    try:
        item["size_bytes"] = path.stat().st_size
    except OSError:
        item["reason"] = "无法读取文件属性"
        return item

    try:
        with path.open("rb") as handle:
            handle.read(1)
        item["readable"] = True
    except OSError:
        item["reason"] = "文件不可读，可能被占用或权限不足"
        return item

    if item["suffix"] in IMAGE_SUFFIXES:
        item["status"] = "ready"
        item["reason"] = "图像文件，由模型读取内容"
        item["notes"].append("手写内容需逐题识别，字迹不清处将标记待确认")
        return item

    ok, detail = has_text_layer(path)
    if not ok:
        item["status"] = "unavailable"
        item["reason"] = detail
        item["notes"].append("建议提供含文字层的 PDF，或改传图片")
        return item

    if item["suffix"] in TEXT_SUFFIXES:
        text = read_text(path)
        if not text.strip():
            item["status"] = "unavailable"
            item["reason"] = "文本文件为空"
            return item
        item["status"] = "ready"
        item["reason"] = "文本可读"
        item["notes"].append(f"检出字符数 {len(text.strip())}")
        for hint in SCORING_HINTS:
            if hint in text:
                item["notes"].append(f"含评分依据关键词：{hint}")
                break

    item["status"] = "ready"
    item["reason"] = item["reason"] or "文件可读"
    return item


def decide_evidence(paper: dict, answers: list[dict], reference_state: str, scoring_state: str) -> dict:
    if paper["status"] != "ready":
        return {
            "level": "D",
            "label": "未验证",
            "reason": "试卷不可读，无法建立判分基准",
        }
    if reference_state == "unknown":
        return {
            "level": "B",
            "label": "部分验证",
            "reason": "试卷非文本格式，参考答案与评分细则需人工确认后再定证据等级",
        }
    if reference_state == "absent":
        return {
            "level": "D",
            "label": "未验证",
            "reason": "试卷中未检出参考答案，无法建立判分基准",
        }
    if not answers:
        return {
            "level": "B",
            "label": "部分验证",
            "reason": "仅分析试卷考查范围，未提供答卷",
        }
    if all(item["status"] != "ready" for item in answers):
        return {
            "level": "D",
            "label": "未验证",
            "reason": "答卷均不可读",
        }
    if scoring_state == "present":
        return {
            "level": "A",
            "label": "已核验",
            "reason": "试卷、参考答案与评分细则齐全，可做单点失分分析",
        }
    return {
        "level": "B",
        "label": "部分验证",
        "reason": "有参考答案但未提供评分细则，只能判要点方向，不给精确分",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="检查试卷与答卷文件，输出可用性与证据等级")
    parser.add_argument("paper", type=Path, help="试卷文件路径")
    parser.add_argument("answers", nargs="*", type=Path, help="答卷文件路径，可多个")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出")
    args = parser.parse_args()

    ok, version_detail = is_python_supported()
    paper_item = classify(args.paper, "paper")
    answer_items = [classify(path, "answer") for path in args.answers]

    reference_state = "unknown"
    scoring_state = "unknown"
    if paper_item["status"] != "ready":
        reference_state = "unknown"
        scoring_state = "unknown"
        paper_item["notes"].append("试卷不可读，参考答案与评分细则存在性无法判定")
    elif paper_item["suffix"] in TEXT_SUFFIXES:
        combined_text = read_text(args.paper)
        reference_state = "present" if any(hint in combined_text for hint in REFERENCE_HINTS) else "absent"
        scoring_state = "present" if any(hint in combined_text for hint in SCORING_HINTS) else "absent"
    else:
        reference_state = "unknown"
        scoring_state = "unknown"
        paper_item["notes"].append("非文本格式，参考答案与评分细则存在性需人工确认")

    reference_available = reference_state == "present"
    scoring_detail_available = scoring_state == "present"

    evidence = decide_evidence(paper_item, answer_items, reference_state, scoring_state)

    if not ok or paper_item["status"] != "ready":
        status = "missing" if not paper_item["exists"] else "unavailable"
    elif evidence["level"] in {"A", "B"} and all(item["status"] == "ready" for item in answer_items) if answer_items else True:
        status = "ready"
    else:
        status = "partial"

    if not ok:
        status = "unavailable"

    result = {
        "status": status,
        "python": {"supported": ok, "detail": version_detail},
        "paper": paper_item,
        "answers": answer_items,
        "materials": {
            "reference": reference_state,
            "scoring_detail": scoring_state,
            "reference_available": reference_available,
            "scoring_detail_available": scoring_detail_available,
        },
        "evidence_level": evidence,
        "next_action": {
            "ready": "进入正常流程",
            "partial": "进入受限流程，按证据等级标注结论",
            "missing": "请补充试卷文件",
            "unavailable": "按原因恢复或改用可读格式",
        }[status],
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"总体状态：{status}")
        print(f"Python：{version_detail}")
        print(f"试卷：{paper_item['name']} — {paper_item['status']}（{paper_item['reason']}）")
        for note in paper_item["notes"]:
            print(f"  · {note}")
        if answer_items:
            print("答卷：")
            for item in answer_items:
                print(f"  · {item['name']} — {item['status']}（{item['reason']}）")
                for note in item["notes"]:
                    print(f"    - {note}")
        else:
            print("答卷：未提供")
        label_map = {"present": "有", "absent": "未检出", "unknown": "无法判定"}
        print(f"参考答案：{label_map[reference_state]}")
        print(f"评分细则：{label_map[scoring_state]}")
        print(f"证据等级：{evidence['level']} {evidence['label']} — {evidence['reason']}")
        print(f"下一步：{result['next_action']}")

    return 0 if status in {"ready", "partial"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
