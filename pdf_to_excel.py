#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF → Excel 小工具：每页一个工作表。
默认按页导出网格（--sheet-format grid）。工资单可改用 --sheet-format payslip（Table 1…）；
payslip 模式会在写入前用明细行修正汇总区 Total/Net（避免 PDF 汇总行为 0 的占位数字）。
依赖: pip install pdfplumber pandas openpyxl
可选: pip install pymupdf  （仅作无 pdfplumber 时的文字回退）

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


def _text_rows_pdfplumber(page) -> list[list[str]]:
    """Full-page text as one column per line (layout mode matches payroll_analyzer)."""
    t = None
    try:
        t = page.extract_text(layout=True, x_tolerance=2, y_tolerance=2)
    except TypeError:
        t = page.extract_text(layout=True)
    if not t:
        try:
            t = page.extract_text(x_tolerance=2, y_tolerance=2)
        except TypeError:
            t = page.extract_text()
    if not t:
        t = page.extract_text(layout=True) or page.extract_text() or ""
    lines = [ln.rstrip() for ln in t.splitlines()]
    if not any(x.strip() for x in lines):
        return [["(此页无文字，可能是扫描图；请先 OCR)"]]
    return [[ln] for ln in lines]


def _page_text_fallback_pdfplumber(page) -> list[list[str]]:
    return _text_rows_pdfplumber(page)


def _extract_tables_merged(page) -> list[list[str]]:
    """Return merged grid rows from pdfplumber, or [] if no usable tables."""
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

    if not tables or all(not t for t in tables):
        return []
    merged = _page_tables_to_rows([t for t in tables if t])
    if merged == [[""]]:
        return []
    return merged


def _nonblank_line_count(rows: list[list[str]]) -> int:
    n = 0
    for r in rows:
        if any((c or "").strip() for c in r):
            n += 1
    return n


def _nonblank_char_count(rows: list[list[str]]) -> int:
    return sum(len((c or "").strip()) for r in rows for c in r)


def _text_char_count(text_rows: list[list[str]]) -> int:
    return sum(len((r[0] or "").strip()) for r in text_rows if r)


def _pad_rows(rows: list[list[str]], width: int) -> list[list[str]]:
    out: list[list[str]] = []
    for r in rows:
        rr = list(r)
        if len(rr) < width:
            rr.extend([""] * (width - len(rr)))
        else:
            rr = rr[:width]
        out.append(rr)
    return out


def _page_rows_pdfplumber(page, mode: str) -> list[list[str]]:
    """
    mode:
      auto  — tables if they look complete vs full-page text; else layout text
      text  — always layout text (best for many payslips)
      tables— tables only, fallback to text if none
      both  — tables then separator then full text
    """
    text_rows = _text_rows_pdfplumber(page)
    table_rows = _extract_tables_merged(page)

    if mode == "text":
        return text_rows

    if mode == "tables":
        return table_rows if table_rows else text_rows

    if mode == "both":
        if not table_rows:
            return text_rows
        maxc = max(max(len(r) for r in table_rows), max((len(r) for r in text_rows), 1))
        sep = _pad_rows([["--- full page text below ---"]], maxc)
        blank = _pad_rows([[""]], maxc)
        return (
            _pad_rows(table_rows, maxc)
            + blank
            + sep
            + blank
            + _pad_rows(text_rows, maxc)
        )

    # ---- auto ----
    if not table_rows:
        return text_rows

    tl = _nonblank_line_count(text_rows)
    tb = _nonblank_line_count(table_rows)
    tc_txt = _text_char_count(text_rows)
    tc_tbl = _nonblank_char_count(table_rows)

    if tl >= 10:
        if tb <= max(5, tl // 3) and tc_tbl < tc_txt * 0.38:
            return text_rows
        if tc_txt > 450 and tc_tbl < max(350, int(tc_txt * 0.34)):
            return text_rows

    return table_rows


def _page_text_fallback_pymupdf(doc: "fitz.Document", page_index: int) -> list[list[str]]:
    page = doc.load_page(page_index)
    t = page.get_text("text") or ""
    lines = t.splitlines()
    if not lines:
        return [["(此页无文字，可能是扫描图；请先 OCR)"]]
    return [[ln] for ln in lines]


PAYSLIP_COLS = 10


def _pad_row(cells: list) -> list[str]:
    r = ["" if c is None else str(c) for c in cells]
    if len(r) >= PAYSLIP_COLS:
        return r[:PAYSLIP_COLS]
    return r + [""] * (PAYSLIP_COLS - len(r))


def _fmt_money(v) -> str:
    if v is None:
        return ""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{x:,.2f}"


def _fmt_qty(v) -> str:
    if v is None:
        return ""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return str(v)
    s = f"{x:.6f}".rstrip("0").rstrip(".")
    return s or "0"


def _extract_page_plain_text(pdf_path: Path, page_index: int) -> str:
    chunks: list[str] = []
    if pdfplumber is not None:
        with pdfplumber.open(str(pdf_path)) as pdf:
            pg = pdf.pages[page_index]
            t = None
            try:
                t = pg.extract_text(layout=True, x_tolerance=2, y_tolerance=2)
            except TypeError:
                t = pg.extract_text(layout=True)
            if not t:
                try:
                    t = pg.extract_text(x_tolerance=2, y_tolerance=2)
                except TypeError:
                    t = pg.extract_text()
            if not t:
                t = pg.extract_text(layout=True) or pg.extract_text() or ""
            chunks.append(t)
    if fitz is not None:
        doc = fitz.open(str(pdf_path))
        try:
            chunks.append(doc[page_index].get_text("text") or "")
        finally:
            doc.close()
    if not chunks:
        return ""
    return max(chunks, key=lambda s: len((s or "").strip()))


def _looks_like_person_line(s: str) -> bool:
    s = (s or "").strip()
    if len(s) < 4 or len(s) > 80:
        return False
    if re.search(r"\d{2}-\d{4}-\d+", s):
        return False
    if any(k in s.lower() for k in ("pay period", "payment date", "earnings", "employment", "weekly", "fortnight")):
        return False
    parts = s.split()
    if len(parts) < 2:
        return False
    letters = sum(ch.isalpha() for ch in s)
    return letters >= max(6, len(s) // 3)


def _guess_name_address(lines: list[str]) -> tuple[str | None, list[str]]:
    name = None
    addr: list[str] = []
    i0 = 0
    for i, raw in enumerate(lines[:30]):
        if "EMPLOYMENT" in raw.upper() and "DETAIL" in raw.upper():
            i0 = i + 1
            break
    for j in range(i0, min(len(lines), i0 + 18)):
        s = lines[j].strip()
        if not s:
            continue
        u = s.upper()
        if "PAY PERIOD" in u and "PAYMENT" in u:
            break
        if "EARNINGS" in u and "QUANTITY" in u:
            break
        if name is None and _looks_like_person_line(s):
            name = s
            continue
        if name and not addr and (re.search(r"\d", s) or "auckland" in s.lower() or "/" in s):
            addr.append(s)
            if len(addr) >= 3:
                break
        elif name and len(addr) == 1 and len(s) > 8:
            addr.append(s)
            break
    return name, addr


def _label_money_first_nonzero(text: str, patterns: tuple[str, ...]) -> float | None:
    """Same idea as payroll Excel import: skip leading $0.00 from noisy PDF text."""
    for pat in patterns:
        for m in re.finditer(pat, text, re.I | re.MULTILINE):
            try:
                v = float(m.group(1).replace(",", ""))
                if abs(v) > 1e-9:
                    return v
            except (ValueError, IndexError):
                continue
    return None


def _earning_line_money(e: dict) -> float:
    """Amount used for reconciling totals when THIS PAY is 0 but YTD has the figure."""
    tp = e.get("this_pay")
    yd = e.get("ytd")
    if isinstance(tp, (int, float)) and abs(float(tp)) > 1e-9:
        return float(tp)
    if isinstance(yd, (int, float)) and abs(float(yd)) > 1e-9:
        return float(yd)
    return 0.0


def finalize_payslip_totals(d: dict) -> None:
    """
    PDF 转表时汇总行常为 $0；在写入 Excel 前用明细 + 税务 + 银行行对齐 Total / Net，
    减轻下游 payroll_analyzer 的误判。
    """
    earnings = d.get("earnings") or []
    inferred_total = sum(_earning_line_money(e) for e in earnings)

    te = d.get("total_earnings")
    if te is None or (isinstance(te, (int, float)) and abs(float(te)) < 1e-9):
        if inferred_total > 1e-9:
            d["total_earnings"] = inferred_total
            te = inferred_total

    np = d.get("net_pay")
    if np is None or (isinstance(np, (int, float)) and abs(float(np)) < 1e-9):
        payments = d.get("payments") or []
        bank_max = 0.0
        for p in payments:
            a = p.get("amount")
            if isinstance(a, (int, float)) and float(a) > bank_max:
                bank_max = float(a)
        if bank_max > 1e-9:
            d["net_pay"] = bank_max
            return

        te_f = float(d.get("total_earnings") or 0)
        paye_this = 0.0
        for tr in d.get("tax_rows") or []:
            x = tr.get("this_pay")
            if isinstance(x, (int, float)) and float(x) > paye_this:
                paye_this = float(x)
        if te_f > 1e-9 and paye_this > 1e-9 and te_f > paye_this + 1e-6:
            d["net_pay"] = te_f - paye_this
        elif te_f > 1e-9:
            d["net_pay"] = te_f


def parse_payslip_export_dict(text: str) -> dict:
    """Parse one payslip page (Xero-style text). Keeps logic self-contained (no PyQt import)."""
    lines = text.split("\n")
    t = text

    name = None
    nm = re.search(r"EMPLOYMENT DETAILS\s*\n?\s*([^\n]+)", t)
    if nm:
        cand = nm.group(1).strip()
        if not any(x in cand.lower() for x in ("pay frequency", "weekly", "ird", "tax code")):
            name = cand

    pay_period = None
    m = re.search(r"Pay Period:\s*([^\n]+?)(?:\s{2,}Payment|\n)", t, re.I)
    if not m:
        m = re.search(r"Pay Period:\s*([\d\s\w\-–]+)", t, re.I)
    if m:
        pay_period = m.group(1).strip()

    payment_date = None
    m = re.search(r"Payment Date:\s*([\d\s\w]+)", t, re.I)
    if m:
        payment_date = m.group(1).strip()

    total_earnings = _label_money_first_nonzero(
        t,
        (
            r"Total\s+Earnings:\s*\$?([\d,]+\.\d+)",
            r"Gross\s+(?:Pay|Earnings|Income):\s*\$?([\d,]+\.\d+)",
            r"Total\s+Gross:\s*\$?([\d,]+\.\d+)",
        ),
    )

    net_pay = _label_money_first_nonzero(
        t,
        (
            r"Net\s+Pay:\s*\$?([\d,]+\.\d+)",
            r"(?:Take\s*Home|Take-home|Net\s+Amount|Amount\s+Payable):\s*\$?([\d,]+\.\d+)",
            r"Pay\s+into\s+bank:\s*\$?([\d,]+\.\d+)",
        ),
    )

    employment_lines: list[str] = []
    for raw in lines[:45]:
        s = raw.strip()
        if not s:
            continue
        low = s.lower()
        if "employment details" in low:
            continue
        if s.upper().startswith("EARNINGS") and "QUANTITY" in s.upper():
            break
        if low.startswith("pay frequency") or "ird number" in low or "ird no" in low:
            employment_lines.append(s)
        elif low.startswith("tax code") or low.startswith("tax period"):
            employment_lines.append(s)
        if "pay period:" in low and "payment date" in low:
            break

    g_name, g_addr = _guess_name_address(lines)
    if not name:
        name = g_name
    address_lines = g_addr

    earnings_start = -1
    earnings_end = -1
    for i, line in enumerate(lines):
        if "EARNINGS" in line and earnings_start == -1:
            earnings_start = i
        if earnings_start != -1 and ("TAX" in line and "PAYE" in line):
            earnings_end = i
            break
        if earnings_start != -1 and line.strip().upper().startswith("TAX"):
            earnings_end = i
            break
        if earnings_start != -1 and "LEAVE" in line and "Annual" in line:
            earnings_end = i
            break

    if earnings_start == -1:
        earnings_start = 0
    if earnings_end == -1:
        earnings_end = len(lines)

    earnings: list[dict] = []
    for line in lines[earnings_start:earnings_end]:
        line = line.strip()
        if not line or line == "EARNINGS" or line == "TOTAL":
            continue
        if "QUANTITY" in line and "RATE" in line:
            continue
        if "THIS PAY" in line and "YTD" in line:
            continue

        amounts_found = re.findall(r"\$([\d,]+(?:\.\d+)?)", line)
        if not amounts_found:
            amounts_found = re.findall(r"(?:^|\s)([\d,]+\.\d{2})(?=\s|$)", line)

        if not amounts_found:
            continue

        if len(amounts_found) >= 2:
            this_pay = float(amounts_found[-2].replace(",", ""))
            ytd = float(amounts_found[-1].replace(",", ""))
        else:
            this_pay = float(amounts_found[-1].replace(",", ""))
            ytd = None

        item_text = line
        for amt in amounts_found:
            item_text = item_text.replace(f"${amt}", "")
        item_name = re.sub(r"\s+", " ", item_text.strip())
        item_name = re.sub(r"\d+\.\d+\s*(hours?|Hours?|hour)?", "", item_name, flags=re.I).strip()
        item_name = re.sub(r"\s+", " ", item_name).strip()
        if not item_name:
            continue

        quantity = None
        rate = None
        qty_match = re.search(r"(\d+\.\d+)\s*(hours?|Hours?)", line, re.I)
        if qty_match:
            quantity = float(qty_match.group(1))
        rate_match = re.search(r"\$(\d+\.\d+)\s+\$", line)
        if rate_match:
            rate = float(rate_match.group(1))

        earnings.append(
            {
                "name": item_name,
                "quantity": quantity,
                "rate": rate,
                "this_pay": this_pay,
                "ytd": ytd,
            }
        )

    tax_rows: list[dict] = []
    in_tax = False
    for line in lines:
        u = line.strip().upper()
        if u.startswith("TAX") and "THIS" in u:
            in_tax = True
            continue
        if in_tax and u.startswith("PAYMENTS"):
            break
        if in_tax and re.search(r"\bPAYE\b", line, re.I):
            amts = re.findall(r"\$([\d,]+(?:\.\d+)?)", line)
            if amts:
                tax_rows.append(
                    {
                        "name": "PAYE",
                        "this_pay": float(amts[0].replace(",", "")),
                        "ytd": float(amts[1].replace(",", "")) if len(amts) > 1 else None,
                    }
                )

    payments: list[dict] = []
    in_pay = False
    for line in lines:
        u = line.strip().upper()
        if u.startswith("PAYMENTS"):
            in_pay = True
            continue
        if in_pay and u.startswith("LEAVE"):
            break
        if in_pay:
            mba = re.search(r"(\d{2}-\d{4}-\d+-\d+)", line)
            amts = re.findall(r"\$([\d,]+(?:\.\d+)?)", line)
            if mba and amts:
                payments.append({"particulars": mba.group(1), "amount": float(amts[-1].replace(",", ""))})

    leave_rows: list[dict] = []
    for line in lines:
        if "Annual Leave" in line and "Hours" in line:
            nums = re.findall(r"\d+\.\d+", line)
            if len(nums) >= 3:
                leave_rows.append(
                    {"kind": "Annual Leave (Hours)", "accrued": nums[0], "used": nums[1], "balance": nums[2]}
                )
        if "Sick Leave" in line and "Hours" in line:
            nums = re.findall(r"\d+\.\d+", line)
            if len(nums) >= 3:
                leave_rows.append(
                    {"kind": "Sick Leave (Hours)", "accrued": nums[0], "used": nums[1], "balance": nums[2]}
                )

    return {
        "name": name,
        "address_lines": address_lines,
        "employment_lines": employment_lines,
        "pay_period": pay_period,
        "payment_date": payment_date,
        "total_earnings": total_earnings,
        "net_pay": net_pay,
        "earnings": earnings,
        "tax_rows": tax_rows,
        "payments": payments,
        "leave_rows": leave_rows,
    }


def _payslip_parse_looks_ok(d: dict) -> bool:
    if d.get("name"):
        return True
    if d.get("earnings"):
        return True
    te = d.get("total_earnings")
    np = d.get("net_pay")
    if te is not None and te > 0:
        return True
    if np is not None and np > 0:
        return True
    return False


def payslip_dict_to_sheet_rows(d: dict) -> list[list[str]]:
    out: list[list[str]] = []

    def add(cells: list):
        out.append(_pad_row(cells))

    add(["EMPLOYMENT DETAILS"])
    for el in d.get("employment_lines") or []:
        add([el])
    add([])
    nm = d.get("name") or ""
    add(["", "", "", "", "", "", nm, "", "", ""])
    for al in d.get("address_lines") or []:
        add(["", "", "", "", "", "", al, "", "", ""])
    add([])
    add(
        [
            f"Pay Period: {d.get('pay_period') or ''}",
            "",
            "",
            f"Payment Date: {d.get('payment_date') or ''}",
            "",
            "",
            f"Total Earnings: {_fmt_money(d.get('total_earnings'))}",
            "",
            f"Net Pay: {_fmt_money(d.get('net_pay'))}",
            "",
        ]
    )
    add([])
    add(["EARNINGS", "QUANTITY", "RATE", "THIS PAY", "YTD"])
    for e in d.get("earnings") or []:
        add(
            [
                e.get("name") or "",
                _fmt_qty(e.get("quantity")) if e.get("quantity") is not None else "",
                _fmt_qty(e.get("rate")) if e.get("rate") is not None else "",
                _fmt_money(e.get("this_pay")),
                _fmt_money(e.get("ytd")) if e.get("ytd") is not None else "",
            ]
        )
    te_sum = sum((x.get("this_pay") or 0) for x in (d.get("earnings") or []) if isinstance(x.get("this_pay"), (int, float)))
    ytd_sum = sum((x.get("ytd") or 0) for x in (d.get("earnings") or []) if isinstance(x.get("ytd"), (int, float)))
    if d.get("earnings"):
        te_cell = d.get("total_earnings")
        tp = _fmt_money(te_cell if te_cell is not None else te_sum)
        ytd_cell = _fmt_money(ytd_sum) if ytd_sum else ""
        add(["TOTAL", "", "", tp, ytd_cell])
    if d.get("tax_rows"):
        add([])
        add(["TAX", "", "", "THIS PAY", "YTD"])
        for tr in d["tax_rows"]:
            add([tr.get("name") or "", "", "", _fmt_money(tr.get("this_pay")), _fmt_money(tr.get("ytd"))])
        t_this = sum(x.get("this_pay") or 0 for x in d["tax_rows"] if isinstance(x.get("this_pay"), (int, float)))
        t_ytd = sum(x.get("ytd") or 0 for x in d["tax_rows"] if isinstance(x.get("ytd"), (int, float)))
        add(["TOTAL", "", "", _fmt_money(t_this), _fmt_money(t_ytd) if t_ytd else ""])
    if d.get("payments"):
        add([])
        add(["PAYMENTS", "PARTICULARS", "CODE", "REFERENCE", "AMOUNT"])
        for p in d["payments"]:
            add(["Bank Account Number", p.get("particulars") or "", "", "", _fmt_money(p.get("amount"))])
        pam = sum(x.get("amount") or 0 for x in d["payments"] if isinstance(x.get("amount"), (int, float)))
        add(["TOTAL", "", "", "", _fmt_money(pam)])
    if d.get("leave_rows"):
        add([])
        add(["LEAVE", "ACCRUED", "USED", "BALANCE"])
        for lr in d["leave_rows"]:
            add([lr.get("kind") or "", lr.get("accrued") or "", lr.get("used") or "", lr.get("balance") or ""])
    return out


def _pdf_page_count(pdf_path: Path) -> int:
    if pdfplumber is not None:
        with pdfplumber.open(str(pdf_path)) as pdf:
            return len(pdf.pages)
    if fitz is not None:
        doc = fitz.open(str(pdf_path))
        try:
            return len(doc)
        finally:
            doc.close()
    return 0


def convert_pdf_to_excel(
    pdf_path: Path,
    xlsx_path: Path,
    page_indices: list[int] | None = None,
    mode: str = "auto",
    sheet_format: str = "grid",
) -> tuple[int, str | None]:
    if not pdf_path.is_file():
        return 1, f"找不到文件: {pdf_path}"

    if sheet_format == "payslip":
        if pdfplumber is None and fitz is None:
            return (
                1,
                "payslip 版式需要 pdfplumber 或 pymupdf：pip install pdfplumber 或 pip install pymupdf",
            )
        n = _pdf_page_count(pdf_path)
        if n == 0:
            return 1, "无法读取 PDF 页数。"
        idxs = page_indices if page_indices is not None else list(range(n))
        idxs = [i for i in idxs if 0 <= i < n]
        if not idxs:
            return 1, "没有选中任何有效页码。"
        bases = [f"Table {si + 1}" for si in range(len(idxs))]
        names = _unique_sheet_names(bases)
        sheets: list[tuple[str, list[list[str]]]] = []
        for si, i in enumerate(idxs):
            text = _extract_page_plain_text(pdf_path, i)
            data = parse_payslip_export_dict(text)
            finalize_payslip_totals(data)
            rows = payslip_dict_to_sheet_rows(data)
            if not _payslip_parse_looks_ok(data):
                rows.append(_pad_row([]))
                rows.append(_pad_row(["(Could not fully parse; raw lines below)"]))
                for ln in text.splitlines():
                    rows.append(_pad_row([ln]))
            sheets.append((names[si], rows))

        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            for name, rows in sheets:
                pd.DataFrame(rows).to_excel(writer, sheet_name=name, index=False, header=False)
        return 0, None

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
                rows = _page_rows_pdfplumber(page, mode)
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
        "  python pdf_to_excel.py  工资单.pdf --sheet-format payslip\n"
        "\n"
        "工资单版式（多列、Table 1…工作表）用 --sheet-format payslip；"
        "若只要整页文字用 --mode text。\n"
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
        "--mode",
        choices=("auto", "text", "tables", "both"),
        default="auto",
        help="grid 模式：auto=智能在表格与版式文本间选择；text=整页文本行；tables=仅表格；both=表格后附全文",
    )
    p.add_argument(
        "--sheet-format",
        choices=("grid", "payslip"),
        default="grid",
        help="grid=按页导出单元格网格（默认）；payslip=按工资单版式导出，工作表名为 Table 1…（接近 Xero 多表）",
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

    code, err = convert_pdf_to_excel(
        pdf_path,
        out,
        page_indices,
        mode=args.mode,
        sheet_format=args.sheet_format,
    )
    if err:
        print(err, file=sys.stderr)
        return code
    print(f"已写入: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
