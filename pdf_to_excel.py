#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF → Excel 小工具：每页一个工作表，优先抽取表格；若无表格则导出版式文本行。
依赖: pip install pdfplumber pandas openpyxl
可选: pip install pymupdf  （仅作无表格时的文字回退，与主程序一致）

扫描件/纯图片 PDF 需先 OCR，本工具无法从像素恢复表格。
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

import pandas as pd


def _safe_sheet_base(name: str) -> str:
    name = re.sub(r"[\[\]:*?/\\]", "_", (name or "").strip())
    return (name or "Page")[:31]


def _unique_sheet_names(bases: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out = []
    for b in bases:
        key = _safe_sheet_base(b)
        if key not in seen:
            seen[key] = 0
            out.append(key)
        else:
            seen[key] += 1
            suffix = f"_{seen[key]}"
            out.append((key[: 31 - len(suffix)] + suffix)[:31])
    return out


def parse_page_spec(spec: str | None, total_pages: int) -> list[int]:
    """1-based page numbers → 0-based indices."""
    if not spec or not spec.strip():
        return list(range(total_pages))
    indices: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            lo, hi = int(a.strip()), int(b.strip())
            for p in range(lo, hi + 1):
                if 1 <= p <= total_pages:
                    indices.add(p - 1)
        else:
            p = int(part)
            if 1 <= p <= total_pages:
                indices.add(p - 1)
    return sorted(indices)


def _normalize_table_rows(table: list[list[str | None]]) -> list[list[str]]:
    rows: list[list[str]] = []
    for r in table:
        if r is None:
            rows.append([])
            continue
        rows.append(["" if c is None else str(c) for c in r])
    if not rows:
        return [[""]]
    mc = max(len(r) for r in rows)
    return [r + [""] * (mc - len(r)) for r in rows]


def _page_tables_to_rows(tables: list) -> list[list[str]]:
    """Stack multiple tables on one sheet with a blank separator row."""
    blocks: list[list[list[str]]] = []
    for t in tables:
        if not t:
            continue
        blocks.append(_normalize_table_rows(t))
    if not blocks:
        return [[""]]
    out: list[list[str]] = []
    max_cols = max(max(len(r) for r in b) for b in blocks)
    for bi, b in enumerate(blocks):
        if bi > 0:
            out.append([""] * max_cols)
        for r in b:
            rr = r + [""] * (max_cols - len(r))
            out.append(rr[:max_cols])
    return out


def _page_text_fallback_pdfplumber(page) -> list[list[str]]:
    t = page.extract_text(layout=True) or page.extract_text() or ""
    lines = t.splitlines()
    if not lines:
        return [["(此页无文字，可能是扫描图；请先 OCR)"]]
    return [[ln] for ln in lines]


def _page_text_fallback_pymupdf(doc: "fitz.Document", page_index: int) -> list[list[str]]:
    page = doc.load_page(page_index)
    t = page.get_text("text") or ""
    lines = t.splitlines()
    if not lines:
        return [["(此页无文字，可能是扫描图；请先 OCR)"]]
    return [[ln] for ln in lines]


def convert_pdf_to_excel(
    pdf_path: Path,
    xlsx_path: Path,
    page_indices: list[int] | None = None,
) -> tuple[int, str | None]:
    if not pdf_path.is_file():
        return 1, f"找不到文件: {pdf_path}"

    if pdfplumber is not None:
        with pdfplumber.open(str(pdf_path)) as pdf:
            n = len(pdf.pages)
            idxs = page_indices if page_indices is not None else list(range(n))
            idxs = [i for i in idxs if 0 <= i < n]
            if not idxs:
                return 1, "没有选中任何有效页码。"

            bases = [f"P{i + 1}" for i in idxs]
            names = _unique_sheet_names(bases)

            sheets: list[tuple[str, list[list[str]]]] = []
            for i, name in zip(idxs, names):
                page = pdf.pages[i]
                try:
                    tables = page.extract_tables(
                        table_settings={
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines",
                            "intersection_tolerance": 3,
                        }
                    )
                except Exception:
                    tables = page.extract_tables()

                if not tables or all(not t for t in tables):
                    try:
                        tables = page.extract_tables(
                            table_settings={
                                "vertical_strategy": "text",
                                "horizontal_strategy": "text",
                            }
                        )
                    except Exception:
                        tables = []

                rows: list[list[str]]
                if tables and any(t for t in tables):
                    rows = _page_tables_to_rows([t for t in tables if t])
                else:
                    rows = _page_text_fallback_pdfplumber(page)
                sheets.append((name, rows))

        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            for name, rows in sheets:
                pd.DataFrame(rows).to_excel(writer, sheet_name=name, index=False, header=False)
        return 0, None

    # No pdfplumber: try PyMuPDF text-only export
    if fitz is None:
        return (
            1,
            "未安装 pdfplumber。请执行: pip install pdfplumber pandas openpyxl\n"
            "或至少安装: pip install pymupdf pandas openpyxl （仅文本按行导出）",
        )

    doc = fitz.open(str(pdf_path))
    try:
        n = len(doc)
        idxs = page_indices if page_indices is not None else list(range(n))
        idxs = [i for i in idxs if 0 <= i < n]
        if not idxs:
            return 1, "没有选中任何有效页码。"
        bases = [f"P{i + 1}" for i in idxs]
        names = _unique_sheet_names(bases)
        sheets = []
        for i, name in zip(idxs, names):
            rows = _page_text_fallback_pymupdf(doc, i)
            sheets.append((name, rows))
    finally:
        doc.close()

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for name, rows in sheets:
            pd.DataFrame(rows).to_excel(writer, sheet_name=name, index=False, header=False)
    return 0, None


def _print_usage_hint() -> None:
    print(
        "用法（必须把 PDF 路径写在命令后面）：\n"
        "  python pdf_to_excel.py  工资单.pdf\n"
        "  python pdf_to_excel.py  工资单.pdf -o 输出.xlsx\n"
        "  python pdf_to_excel.py  工资单.pdf --pages \"1,3,5-8\"\n"
        "\n"
        "Windows：请把 PDF 拖到 pdf_to_excel.bat 上再松开，不要只双击 BAT（那样没有文件可转）。",
        file=sys.stderr,
    )


def main(argv: list[str]) -> int:
    if not argv:
        _print_usage_hint()
        return 2

    p = argparse.ArgumentParser(
        description="将 PDF 转为 Excel（每页一表；优先表格识别，否则文本行）。"
    )
    p.add_argument("pdf", type=Path, help="输入 PDF 路径")
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="输出 .xlsx 路径（默认与 PDF 同目录、同名）",
    )
    p.add_argument(
        "--pages",
        type=str,
        default=None,
        help='只导出部分页，1 起始，如: "1,3,5-8"',
    )
    args = p.parse_args(argv)

    pdf_path = args.pdf.expanduser().resolve()
    if args.output:
        out = args.output.expanduser().resolve()
    else:
        out = pdf_path.with_suffix(".xlsx")

    page_indices = None
    if args.pages and pdfplumber is not None:
        with pdfplumber.open(str(pdf_path)) as pdf:
            page_indices = parse_page_spec(args.pages, len(pdf.pages))
    elif args.pages and fitz is not None:
        doc = fitz.open(str(pdf_path))
        try:
            page_indices = parse_page_spec(args.pages, len(doc))
        finally:
            doc.close()

    code, err = convert_pdf_to_excel(pdf_path, out, page_indices)
    if err:
        print(err, file=sys.stderr)
        return code
    print(f"已写入: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
