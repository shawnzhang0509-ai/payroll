#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cloud 7 Payroll Analyzer - 工资单分析系统
Version: 1.4.3 (week-ID preview: plain text + truncation + safe commit)
"""

import sys
import os
import json
import re
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
from collections import defaultdict

import traceback

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QFileDialog, QMessageBox, QTabWidget, QHeaderView,
    QDialog, QLineEdit, QDialogButtonBox,
    QPlainTextEdit,
    QStatusBar, QMenuBar, QMenu,
    QListWidget, QListWidgetItem, QInputDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    import matplotlib
    matplotlib.use('Qt5Agg')
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

DEFAULT_CONFIG = {
    "branches": [
        "AKL仓库", "chc-warehouse", "chc-showroom", "ham-warehouse",
        "ham-showroom", "onehunga", "westgate", "支援澳洲", "未分配"
    ],
    "category_mapping": {
        "labor_cost": {
            "name": "真正工时成本",
            "color": "#2E86AB",
            "keywords": ["Ordinary Time", "Salary", "Ordinary"],
            "description": "基于实际工作时间的直接劳动成本"
        },
        "benefits": {
            "name": "福利",
            "color": "#A23B72",
            "keywords": ["Public Holiday", "Holiday Pay", "Extra half working on Holiday", 
                        "Anzac Day", "Annual Leave", "Sick Leave"],
            "description": "非工作时间的补偿性收入"
        },
        "performance": {
            "name": "Performance奖金",
            "color": "#F18F01",
            "keywords": ["Commission", "Sales Commission", "Big order commission",
                        "Review Award", "Shop Benchmark", "Bonus", "Incentive"],
            "description": "基于绩效的变动收入"
        },
        "fixed_salary": {
            "name": "Fixed年薪制工时成本",
            "color": "#C73E1D",
            "keywords": ["Salary"],
            "description": "年薪制员工的固定薪资部分"
        },
        "other": {
            "name": "其他/未分类",
            "color": "#6A4C93",
            "keywords": ["Other Earnings", "support subsidy", "KiwiSaver", "ESCT"],
            "description": "其他收入或未分类项目"
        }
    },
    "history_file": "payroll_history.json"
}

class Employee:
    def __init__(self, name, name_key=None, branch="未分配", english_name=""):
        self.name = name
        self.name_key = name_key or name.lower().strip()
        self.branch = branch
        self.english_name = english_name
        self.history = []

class PayrollDatabase:
    def __init__(self, config):
        self.config = config
        self.employees = {}
        self.branches = set(config.get('branches', DEFAULT_CONFIG['branches']))
        self.mapping_rules = config.get('category_mapping', DEFAULT_CONFIG['category_mapping'])
        self.history = {}
        self.load_data()

    def load_data(self):
        data_file = self.config.get('history_file', 'payroll_history.json')
        if os.path.exists(data_file):
            try:
                with open(data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for emp_data in data.get('employees', []):
                    emp = Employee(emp_data['name'], emp_data['name_key'], 
                                 emp_data.get('branch', '未分配'),
                                 emp_data.get('english_name', ''))
                    self.employees[emp.name_key] = emp
                self.branches.update(data.get('branches', []))
                self.history = data.get('history', {})
            except Exception as e:
                print(f"加载历史数据失败: {e}")

    def save_data(self):
        data = {
            'employees': [
                {'name': e.name, 'name_key': e.name_key, 'branch': e.branch,
                 'english_name': e.english_name}
                for e in self.employees.values()
            ],
            'branches': list(self.branches),
            'category_mapping': self.mapping_rules,
            'history': self.history
        }
        with open(self.config.get('history_file', 'payroll_history.json'), 
                  'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def get_or_create_employee(self, name, name_key=None, branch="未分配"):
        key = name_key or name.lower().strip()
        if key not in self.employees:
            self.employees[key] = Employee(name, key, branch)
            self.save_data()
        return self.employees[key]

    def set_employee_branch(self, name_key, branch):
        if name_key in self.employees:
            self.employees[name_key].branch = branch
            self.save_data()

    def get_branch_summary(self, week_id=None):
        summary = defaultdict(lambda: {
            'labor_cost': 0, 'benefits': 0, 'performance': 0,
            'fixed_salary': 0, 'other': 0, 'total_earnings': 0,
            'hours': 0, 'employee_count': 0, 'net_pay': 0
        })

        target_weeks = [week_id] if week_id else list(self.history.keys())

        for wk in target_weeks:
            if wk not in self.history:
                continue
            for emp_key, payroll_data in self.history[wk].items():
                emp = self.employees.get(emp_key)
                if not emp:
                    continue
                branch = emp.branch
                summary[branch]['labor_cost'] += payroll_data.get('labor_cost', 0)
                summary[branch]['benefits'] += payroll_data.get('benefits', 0)
                summary[branch]['performance'] += payroll_data.get('performance', 0)
                summary[branch]['fixed_salary'] += payroll_data.get('fixed_salary', 0)
                summary[branch]['other'] += payroll_data.get('other', 0)
                summary[branch]['total_earnings'] += payroll_data.get('total_earnings', 0)
                summary[branch]['net_pay'] += payroll_data.get('net_pay', 0)
                summary[branch]['hours'] += payroll_data.get('hours', 0)
                summary[branch]['employee_count'] += 1

        return dict(summary)

# ============================================================
# PDF text extraction: pdfplumber + PyMuPDF fallback (scanned PDFs still need OCR)
# ============================================================

MIN_PDF_TEXT_CHARS = 40


def _pdf_text_pdfplumber(filepath):
    """Extract text with tolerances; layout mode helps some payroll PDFs."""
    chunks = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            t = page.extract_text(x_tolerance=2, y_tolerance=2)
            if not t:
                t = page.extract_text(layout=True)
            if not t:
                t = page.extract_text()
            if t:
                chunks.append(t)
    return "\n".join(chunks)


def _pdf_text_pymupdf(filepath):
    doc = fitz.open(filepath)
    try:
        parts = []
        for page in doc:
            parts.append(page.get_text("text"))
        return "\n".join(parts)
    finally:
        doc.close()


def extract_pdf_text(filepath):
    """
    Return (text, error_message). Chooses the longest non-empty extraction
    between pdfplumber and PyMuPDF—many payroll PDFs work with only one engine.
    """
    path = Path(filepath)
    if not path.is_file():
        return None, f"找不到文件或无法读取: {filepath}"

    candidates = []
    errors = []

    if PDF_AVAILABLE:
        try:
            candidates.append(_pdf_text_pdfplumber(filepath))
        except Exception as e:
            errors.append(f"pdfplumber: {e}")

    if PYMUPDF_AVAILABLE:
        try:
            candidates.append(_pdf_text_pymupdf(filepath))
        except Exception as e:
            errors.append(f"PyMuPDF: {e}")

    if not candidates:
        msg = "无法读取PDF（请安装: pip install pdfplumber pymupdf）。"
        if errors:
            msg += " " + " ".join(errors)
        return None, msg

    text = max(candidates, key=lambda s: len((s or "").strip()))

    if len(text.strip()) < MIN_PDF_TEXT_CHARS:
        msg = (
            "从PDF中提取的文字过少。常见原因：1) 扫描件/图片型工资单（需先用OCR生成可选中文字）；"
            "2) 加密或权限受限的PDF。"
        )
        if errors:
            msg += " 解析器报错: " + " ".join(errors)
        return None, msg

    return text, None


class PDFPayrollParser:
    @staticmethod
    def parse_pdf(filepath):
        if not PDF_AVAILABLE and not PYMUPDF_AVAILABLE:
            return None, "请至少安装其一: pip install pdfplumber 或 pip install pymupdf"

        text, ext_err = extract_pdf_text(filepath)
        if ext_err:
            return None, ext_err

        data, parse_err = PDFPayrollParser.parse_text(text)
        if parse_err:
            return None, parse_err
        return data, None

    @staticmethod
    def parse_text(text):
        lines = text.split('\n')

        # 提取基本信息
        name = None
        pay_period = None
        payment_date = None
        total_earnings = 0
        net_pay = 0

        # 姓名：多种常见工资单抬头
        name_match = re.search(r'EMPLOYMENT DETAILS\s*\n?\s*([^\n]+)', text)
        if name_match:
            name = name_match.group(1).strip()
        if not name:
            for pattern in (
                r'Employee\s*Name\s*[:\s]+\s*([^\n]+)',
                r'(?:Full\s*Name|Staff\s*Name)\s*[:\s]+\s*([^\n]+)',
                r'EMPLOYEE\s*[:\s]+\s*([^\n]+)',
            ):
                m = re.search(pattern, text, re.I)
                if m:
                    name = m.group(1).strip()
                    break

        pp_match = re.search(r'Pay Period:\s*([\d\s\w\-–]+)', text)
        if pp_match:
            pay_period = pp_match.group(1).strip()

        pd_match = re.search(r'Payment Date:\s*([\d\s\w]+)', text)
        if pd_match:
            payment_date = pd_match.group(1).strip()

        te_match = re.search(r'Total Earnings:\s*\$?([\d,]+\.\d+)', text)
        if te_match:
            total_earnings = float(te_match.group(1).replace(',', ''))

        np_match = re.search(r'Net Pay:\s*\$?([\d,]+\.\d+)', text)
        if np_match:
            net_pay = float(np_match.group(1).replace(',', ''))

        # ========== 核心修复：更智能的收入项目提取 ==========
        earnings = []

        # 找到EARNINGS区域
        earnings_start = -1
        earnings_end = -1

        for i, line in enumerate(lines):
            if 'EARNINGS' in line and earnings_start == -1:
                earnings_start = i
            if earnings_start != -1 and 'TAX' in line and 'PAYE' in line:
                earnings_end = i
                break
            if earnings_start != -1 and 'LEAVE' in line:
                earnings_end = i
                break

        if earnings_start == -1:
            earnings_start = 0
        if earnings_end == -1:
            earnings_end = len(lines)

        # 解析EARNINGS区域内的每一行
        earnings_lines = lines[earnings_start:earnings_end]

        for line in earnings_lines:
            line = line.strip()
            if not line or line == 'EARNINGS' or line == 'TOTAL':
                continue
            if 'QUANTITY' in line and 'RATE' in line:
                continue
            if 'THIS PAY' in line and 'YTD' in line:
                continue

            # 尝试提取金额 - 找所有$开头的数字
            # 格式1: "Ordinary Time  24.3500 hours  $23.950000  $583.18  $780.77"
            # 格式2: "Other Earnings          $3,101.34"
            # 格式3: "Holiday Pay             $53.91  $69.72"

            amounts_found = re.findall(r'\$([\d,]+(?:\.\d+)?)', line)
            if not amounts_found:
                amounts_found = re.findall(r'(?:^|\s)([\d,]+\.\d{2})(?=\s|$)', line)

            if len(amounts_found) >= 1:
                # 最后一个金额通常是 "THIS PAY"
                this_pay = float(amounts_found[-1].replace(',', ''))

                # 提取项目名称（去掉所有金额部分）
                item_text = line
                for amt in amounts_found:
                    item_text = item_text.replace(f'${amt}', '')

                item_name = item_text.strip()
                # 清理数字残留（hours, rate等）
                item_name = re.sub(r'\d+\.\d+\s*(hours?|Hours?|hour)?', '', item_name).strip()
                item_name = re.sub(r'\s+', ' ', item_name).strip()

                # 尝试提取quantity和rate
                quantity = 0
                rate = 0
                qty_match = re.search(r'(\d+\.\d+)\s*(hours?|Hours?)', line)
                if qty_match:
                    quantity = float(qty_match.group(1))
                    # 找rate（$xx.xxxxxx 格式，在quantity后面）
                    rate_match = re.search(r'\$(\d+\.\d+)\s+\$', line)
                    if rate_match:
                        rate = float(rate_match.group(1))

                earnings.append({
                    'name': item_name,
                    'quantity': quantity,
                    'rate': rate,
                    'amount': this_pay
                })

        result = {
            'name': name,
            'pay_period': pay_period,
            'payment_date': payment_date,
            'total_earnings': total_earnings,
            'net_pay': net_pay,
            'earnings': earnings
        }

        if (
            not name
            and not earnings
            and total_earnings == 0
            and net_pay == 0
        ):
            return None, (
                "已从PDF提取文字，但未匹配到姓名或收入行。可能版式与内置规则不一致；"
                "若为扫描版请先使用OCR。可将提取样本导出以便调整解析规则。"
            )

        return result, None

class ExcelHoursImporter:
    @staticmethod
    def import_hours(filepath):
        try:
            df = pd.read_excel(filepath, sheet_name='Sheet1')
            if 'full name' in df.columns:
                df = df.rename(columns={'full name': 'full_name', 'English name': 'english_name'})

            df = df[df['full_name'].notna()]
            df = df[~df['full_name'].astype(str).str.contains('求和项|总计', na=False)]
            df = df[df['hours'].notna()]

            result = []
            for _, row in df.iterrows():
                result.append({
                    'full_name': str(row.get('full_name', '')).strip(),
                    'english_name': str(row.get('english_name', '')).strip(),
                    'branch': str(row.get('branch', '未分配')).strip(),
                    'hours': float(row.get('hours', 0))
                })
            return result, None
        except Exception as e:
            return None, str(e)


class ExcelPayslipImporter:
    """
    Xero 等导出的「每人一页」Excel：工作簿中多个工作表，版式与 PDF 工资单类似。
    """

    _SKIP_NAME_SUB = (
        'employment', 'pay frequency', 'ird', 'tax code', 'tax period',
        'details', 'weekly', 'fortnightly', 'monthly', 'number',
        'pay period', 'payment date', 'total earnings', 'net pay',
    )

    @staticmethod
    def _is_na(v):
        return v is None or (isinstance(v, float) and pd.isna(v))

    @staticmethod
    def _float_cell(v):
        if ExcelPayslipImporter._is_na(v):
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace(',', '').replace('$', '')
        if not s:
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0

    @staticmethod
    def _normalize_rows(rows):
        if not rows:
            return []
        max_c = max((len(r) for r in rows), default=0)
        out = []
        for r in rows:
            r = list(r)
            if len(r) < max_c:
                r.extend([None] * (max_c - len(r)))
            out.append(r)
        return out

    @staticmethod
    def _rows_to_blob(rows, end_row):
        lines = []
        for r in rows[:max(0, end_row)]:
            parts = [
                str(c).strip()
                for c in r
                if not ExcelPayslipImporter._is_na(c) and str(c).strip()
            ]
            if parts:
                lines.append(' '.join(parts))
        return '\n'.join(lines)

    @staticmethod
    def _find_pay_period_row(rows):
        for ri, row in enumerate(rows):
            for c in row:
                if c is None:
                    continue
                low = str(c).lower()
                if 'pay period' in low and 'tax period' not in low:
                    return ri
        return None

    @classmethod
    def _find_earnings_columns(cls, rows):
        """
        Locate THIS PAY / YTD columns (often on same row; Xero 也可能拆成两行表头).
        Returns (header_last_row_idx, data_start_row, this_pay_col, ytd_col, qty_col, rate_col)
        """
        this_pay_col = ytd_col = qty_col = rate_col = None
        header_last = None

        for ri, row in enumerate(rows):
            tp = yt = qc = rc = None
            for ci, c in enumerate(row):
                if cls._is_na(c):
                    continue
                t = str(c).strip().lower()
                t = re.sub(r'\s+', ' ', t)
                if t in ('this pay',) or t.startswith('this pay'):
                    tp = ci
                elif t == 'ytd':
                    yt = ci
                elif 'quantity' in t or t in ('qty', 'hours'):
                    qc = ci
                elif t == 'rate' or (t.startswith('rate') and len(t) < 14):
                    rc = ci

            if tp is not None and yt is not None:
                header_last = ri
                this_pay_col, ytd_col, qty_col, rate_col = tp, yt, qc, rc
                break

        if this_pay_col is None or ytd_col is None:
            for ri in range(len(rows) - 1):
                tp = yt = qc = rc = None
                max_c = max(len(rows[ri]), len(rows[ri + 1]))
                for ci in range(max_c):
                    for cand_row in (rows[ri], rows[ri + 1]):
                        if ci >= len(cand_row):
                            continue
                        c = cand_row[ci]
                        if cls._is_na(c):
                            continue
                        t = str(c).strip().lower()
                        t = re.sub(r'\s+', ' ', t)
                        if t in ('this pay',) or t.startswith('this pay'):
                            tp = ci
                        elif t == 'ytd':
                            yt = ci
                        elif 'quantity' in t or t in ('qty', 'hours'):
                            qc = ci
                        elif t == 'rate' or (t.startswith('rate') and len(t) < 14):
                            rc = ci
                if tp is not None and yt is not None:
                    header_last = ri + 1
                    this_pay_col, ytd_col, qty_col, rate_col = tp, yt, qc, rc
                    break

        if this_pay_col is None or ytd_col is None:
            for ri, row in enumerate(rows):
                joined = ' '.join(str(c).lower() for c in row if c is not None)
                if 'earnings' not in joined or 'quantity' not in joined:
                    continue
                tp_c = yt_c = qc_c = rc_c = None
                for ci, c in enumerate(row):
                    if cls._is_na(c):
                        continue
                    t = str(c).strip().lower()
                    t = re.sub(r'\s+', ' ', t)
                    if t in ('this pay',) or t.startswith('this pay'):
                        tp_c = ci
                    elif t == 'ytd':
                        yt_c = ci
                    elif 'quantity' in t or t in ('qty', 'hours'):
                        qc_c = ci
                    elif t == 'rate' or (t.startswith('rate') and len(t) < 14):
                        rc_c = ci
                if tp_c is not None and yt_c is not None:
                    header_last = ri
                    this_pay_col, ytd_col, qty_col, rate_col = tp_c, yt_c, qc_c, rc_c
                    break
                if len(row) >= 5:
                    header_last = ri
                    this_pay_col = 3
                    ytd_col = 4
                    break

        if header_last is None:
            return None, None, None, None, None, None

        data_start = header_last + 1
        while data_start < len(rows):
            r = rows[data_start]
            joined = ' '.join(str(c).strip().lower() for c in r if not cls._is_na(c))
            if joined and all(
                x in joined
                for x in ('this pay', 'ytd')
            ):
                data_start += 1
                continue
            if joined == 'quantity rate this pay ytd' or joined.startswith('quantity '):
                data_start += 1
                continue
            break

        return header_last, data_start, this_pay_col, ytd_col, qty_col, rate_col

    @classmethod
    def _looks_like_person_name(cls, line):
        line = line.strip()
        if len(line) < 3 or len(line) > 70:
            return False
        low = line.lower()
        if any(sk in low for sk in cls._SKIP_NAME_SUB):
            return False
        if ':' in line and len(line) < 55:
            return False
        if re.search(r'\d{4,}', line):
            return False
        if re.match(r'^[\d\s$,.%-]+$', line):
            return False
        # Latin letters + Māori diacritics + 中文
        return bool(
            re.match(
                r"^[\s'A-Za-z\u0080-\u024f\u4e00-\u9fff]"
                r"[\s'A-Za-z\u0080-\u024f\u4e00-\u9fff.\-]{2,}$",
                line,
            )
        )

    @classmethod
    def _guess_name(cls, rows, pay_row_idx):
        """优先 Pay Period 行上方的姓名格；支持合并单元格内换行。"""
        ordered_indices = []
        for d in range(1, 10):
            ri = pay_row_idx - d
            if ri >= 0:
                ordered_indices.append(ri)
        for ri in range(min(pay_row_idx, 24)):
            if ri not in ordered_indices:
                ordered_indices.append(ri)

        seen = set()
        for ri in ordered_indices:
            if ri in seen:
                continue
            seen.add(ri)
            for c in rows[ri]:
                if cls._is_na(c):
                    continue
                raw = str(c).replace('\r', '\n')
                for line in raw.split('\n'):
                    line = line.strip()
                    if cls._looks_like_person_name(line):
                        return line
        return None

    @classmethod
    def _parse_meta_blob(cls, blob):
        """汇总区的 Pay Period / Total / Net 与 PDF 规则对齐。"""
        name = None
        pay_period = None
        payment_date = None

        m = re.search(r'EMPLOYMENT DETAILS\s*\n?\s*([^\n]{1,120})', blob)
        if m:
            raw = m.group(1).strip()
            if 'pay frequency' not in raw.lower() and len(raw) < 90:
                name = raw

        if not name:
            for pattern in (
                r'Employee\s*Name\s*[:\s]+\s*([^\n]+)',
                r'(?:Full\s*Name|Staff\s*Name)\s*[:\s]+\s*([^\n]+)',
                r'EMPLOYEE\s*[:\s]+\s*([^\n]+)',
            ):
                m = re.search(pattern, blob, re.I)
                if m:
                    name = m.group(1).strip()
                    break

        m = re.search(r'Pay Period:\s*([\d\s\w\-–]+)', blob)
        if m:
            pay_period = m.group(1).strip()

        m = re.search(r'Payment Date:\s*([\d\s\w\-–]+)', blob)
        if m:
            payment_date = m.group(1).strip()

        money_num = r'([\d,]+(?:\.\d{1,2})?)'

        def scan_money(patterns):
            for pat in patterns:
                m = re.search(pat, blob, re.I | re.MULTILINE)
                if m:
                    try:
                        return float(m.group(1).replace(',', ''))
                    except ValueError:
                        continue
            return 0.0

        total_earnings = scan_money([
            rf'Total\s+Earnings:\s*\$?{money_num}',
            rf'Gross\s+(?:Pay|Earnings|Income):\s*\$?{money_num}',
            rf'Total\s+Gross:\s*\$?{money_num}',
            rf'Earnings\s+Total:\s*\$?{money_num}',
        ])
        net_pay = scan_money([
            rf'Net\s+Pay:\s*\$?{money_num}',
            rf'(?:Take\s*Home|Take-home|Net\s+Amount|Amount\s+Payable):\s*\$?{money_num}',
            rf'Pay\s+into\s+bank:\s*\$?{money_num}',
        ])

        return {
            'name': name,
            'pay_period': pay_period,
            'payment_date': payment_date,
            'total_earnings': total_earnings,
            'net_pay': net_pay,
        }

    @classmethod
    def _first_label_cell(cls, row, max_col):
        """收入项目名称可能在 A 列或 B 列（左侧有空列）。"""
        for ci in range(min(max_col + 1, len(row))):
            if cls._is_na(row[ci]):
                continue
            s = str(row[ci]).strip()
            if not s:
                continue
            if re.match(r'^[\d$,.%\s-]+$', s) and len(s) < 18:
                continue
            return s
        return ''

    @classmethod
    def _parse_earnings_rows(cls, rows, data_start, this_pay_col, ytd_col, qty_col, rate_col):
        earnings = []
        if this_pay_col is None:
            return earnings

        for ri in range(data_start, len(rows)):
            row = rows[ri]
            if not row:
                continue

            lim = this_pay_col + 1 if this_pay_col is not None else 6
            ls = cls._first_label_cell(row, max(this_pay_col - 1, 5))
            ul = ls.upper()

            if not ls:
                continue
            if ul == 'TAX' or ul == 'LEAVE':
                break
            if ul.startswith('DEDUCTION'):
                break
            if ul == 'EARNINGS':
                continue
            if ul == 'TOTAL' or ls.lower() == 'total':
                break

            qty = (
                cls._float_cell(row[qty_col])
                if qty_col is not None and len(row) > qty_col
                else 0.0
            )
            rate = (
                cls._float_cell(row[rate_col])
                if rate_col is not None and len(row) > rate_col
                else 0.0
            )
            amt = (
                cls._float_cell(row[this_pay_col])
                if len(row) > this_pay_col
                else 0.0
            )

            if abs(amt) < 1e-9 and ytd_col is not None and len(row) > ytd_col:
                amt = cls._float_cell(row[ytd_col])

            if abs(amt) < 1e-9:
                continue

            earnings.append({
                'name': ls,
                'quantity': qty,
                'rate': rate,
                'amount': amt,
            })

        return earnings

    @classmethod
    def parse_sheet(cls, rows, sheet_label=''):
        rows = cls._normalize_rows(rows)
        if not rows:
            return None, "空表"

        pay_idx = cls._find_pay_period_row(rows)
        if pay_idx is None:
            return None, "未找到包含 Pay Period 的汇总行"

        hdr_last, data_start, tp_col, ytd_col, qty_col, rate_col = cls._find_earnings_columns(
            rows
        )

        hdr_idx = hdr_last
        blob_end = hdr_idx if hdr_idx is not None else min(pay_idx + 8, len(rows))
        blob = cls._rows_to_blob(rows, blob_end)

        meta = cls._parse_meta_blob(blob)
        name = cls._guess_name(rows, pay_idx)
        if not name:
            cand = meta.get('name')
            if cand and 'pay frequency' not in cand.lower() and len(cand) < 90:
                name = cand.strip()

        pay_period = meta.get('pay_period') or ''
        payment_date = meta.get('payment_date') or ''
        total_earnings = meta.get('total_earnings') or 0.0
        net_pay = meta.get('net_pay') or 0.0

        earnings = []
        if data_start is not None and tp_col is not None:
            earnings = cls._parse_earnings_rows(
                rows, data_start, tp_col, ytd_col, qty_col, rate_col
            )

        if not earnings:
            full_text = cls._rows_to_blob(rows, len(rows))
            alt, _err = PDFPayrollParser.parse_text(full_text)
            if alt and alt.get('earnings'):
                earnings = alt['earnings']
                if not name and alt.get('name'):
                    name = alt['name']
                if total_earnings == 0 and alt.get('total_earnings'):
                    total_earnings = alt['total_earnings']
                if net_pay == 0 and alt.get('net_pay'):
                    net_pay = alt['net_pay']

        sum_lines = sum(e['amount'] for e in earnings)
        if total_earnings == 0 and sum_lines > 0:
            total_earnings = sum_lines

        result = {
            'name': name,
            'pay_period': pay_period,
            'payment_date': payment_date,
            'total_earnings': total_earnings,
            'net_pay': net_pay,
            'earnings': earnings,
        }

        if (
            not name
            and not earnings
            and total_earnings == 0
            and net_pay == 0
        ):
            return None, "未能解析姓名或收入（版式是否与 Xero 导出一致？）"

        return result, None

    @staticmethod
    def import_workbook(filepath):
        results = []
        errors = []
        try:
            xl = pd.ExcelFile(filepath)
        except Exception as e:
            return [], [str(e)]

        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
                rows = []
                for _, row in df.iterrows():
                    rows.append([None if pd.isna(v) else v for v in row])
                data, err = ExcelPayslipImporter.parse_sheet(rows, sheet_name)
                if err:
                    errors.append(f"{sheet_name}: {err}")
                elif data:
                    results.append(data)
            except Exception as e:
                errors.append(f"{sheet_name}: {e}")

        return results, errors


class PayrollAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cloud 7 Payroll Analyzer - 工资单分析系统 v1.4.3")
        self.setGeometry(100, 100, 1400, 900)

        self.config = self.load_config()
        self.db = PayrollDatabase(self.config)
        self.current_week = None

        self.setup_ui()
        self.setup_menu()
        self.refresh_all()

    def load_config(self):
        config_file = 'payroll_config.json'
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return DEFAULT_CONFIG.copy()

    def save_config(self):
        with open('payroll_config.json', 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def setup_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu('文件(&F)')

        import_pdf_action = QAction('导入PDF工资单(&P)', self)
        import_pdf_action.setShortcut('Ctrl+P')
        import_pdf_action.triggered.connect(self.import_pdf)
        file_menu.addAction(import_pdf_action)

        import_xlsx_payslip_action = QAction('导入Excel工资簿(&X)', self)
        import_xlsx_payslip_action.setShortcut('Ctrl+Shift+X')
        import_xlsx_payslip_action.triggered.connect(self.import_excel_payslips)
        file_menu.addAction(import_xlsx_payslip_action)

        import_excel_action = QAction('导入Excel工时表(&E)', self)
        import_excel_action.setShortcut('Ctrl+E')
        import_excel_action.triggered.connect(self.import_excel)
        file_menu.addAction(import_excel_action)

        file_menu.addSeparator()

        export_action = QAction('导出报告(&X)', self)
        export_action.triggered.connect(self.export_report)
        file_menu.addAction(export_action)

        settings_menu = menubar.addMenu('设置(&S)')

        mapping_action = QAction('分类Mapping规则(&M)', self)
        mapping_action.triggered.connect(self.edit_mapping)
        settings_menu.addAction(mapping_action)

        branches_action = QAction('部门管理(&B)', self)
        branches_action.triggered.connect(self.manage_branches)
        settings_menu.addAction(branches_action)

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)

        toolbar = QHBoxLayout()

        self.week_label = QLabel("当前周: 未选择")
        self.week_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #2E86AB;")
        toolbar.addWidget(self.week_label)

        toolbar.addStretch()

        btn_pdf = QPushButton("📄 导入PDF工资单")
        btn_pdf.setStyleSheet("padding: 8px 16px; font-size: 12px;")
        btn_pdf.clicked.connect(self.import_pdf)
        toolbar.addWidget(btn_pdf)

        btn_xlsx_slips = QPushButton("📑 导入Excel工资簿")
        btn_xlsx_slips.setStyleSheet("padding: 8px 16px; font-size: 12px;")
        btn_xlsx_slips.setToolTip("每人一个工作表，如 Table 1 / Table 2 …")
        btn_xlsx_slips.clicked.connect(self.import_excel_payslips)
        toolbar.addWidget(btn_xlsx_slips)

        btn_excel = QPushButton("📊 导入Excel工时")
        btn_excel.setStyleSheet("padding: 8px 16px; font-size: 12px;")
        btn_excel.clicked.connect(self.import_excel)
        toolbar.addWidget(btn_excel)

        btn_refresh = QPushButton("🔄 刷新")
        btn_refresh.clicked.connect(self.refresh_all)
        toolbar.addWidget(btn_refresh)

        layout.addLayout(toolbar)

        self.tabs = QTabWidget()

        self.tab_summary = self.create_summary_tab()
        self.tabs.addTab(self.tab_summary, "📊 部门汇总")

        self.tab_employees = self.create_employees_tab()
        self.tabs.addTab(self.tab_employees, "👥 员工明细")

        self.tab_charts = self.create_charts_tab()
        self.tabs.addTab(self.tab_charts, "📈 趋势图表")

        self.tab_mapping = self.create_mapping_tab()
        self.tabs.addTab(self.tab_mapping, "⚙️ Mapping规则")

        self.tab_data = self.create_data_tab()
        self.tabs.addTab(self.tab_data, "🗂️ 数据管理")

        layout.addWidget(self.tabs)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("就绪 | PDF / Excel工资簿工资单 / Excel工时表")

    def create_summary_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        week_layout = QHBoxLayout()
        week_layout.addWidget(QLabel("选择周:"))
        self.week_combo = QComboBox()
        self.week_combo.currentTextChanged.connect(self.on_week_changed)
        week_layout.addWidget(self.week_combo)
        week_layout.addStretch()
        layout.addLayout(week_layout)

        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(10)
        self.summary_table.setHorizontalHeaderLabels([
            '部门', '人数', '工时', '1.工时成本', '2.福利', '3.Performance',
            '4.Fixed年薪', '5.其他', '工资单Total', 'Net Pay'
        ])
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.summary_table.setAlternatingRowColors(True)
        self.summary_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #ddd;
                font-size: 11px;
            }
            QHeaderView::section {
                background-color: #2E86AB;
                color: white;
                padding: 5px;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.summary_table)

        self.total_label = QLabel()
        self.total_label.setStyleSheet("font-size: 13px; font-weight: bold; color: #C73E1D; padding: 10px;")
        layout.addWidget(self.total_label)

        return widget

    def create_employees_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("部门筛选:"))
        self.emp_branch_filter = QComboBox()
        self.emp_branch_filter.addItem("全部")
        self.emp_branch_filter.currentTextChanged.connect(self.refresh_employees)
        filter_layout.addWidget(self.emp_branch_filter)

        filter_layout.addWidget(QLabel("搜索:"))
        self.emp_search = QLineEdit()
        self.emp_search.setPlaceholderText("输入姓名...")
        self.emp_search.textChanged.connect(self.refresh_employees)
        filter_layout.addWidget(self.emp_search)
        filter_layout.addStretch()

        btn_assign = QPushButton("✏️ 批量分配部门")
        btn_assign.clicked.connect(self.batch_assign_branch)
        filter_layout.addWidget(btn_assign)

        layout.addLayout(filter_layout)

        self.emp_table = QTableWidget()
        self.emp_table.setColumnCount(12)
        self.emp_table.setHorizontalHeaderLabels([
            '姓名', '英文名', '部门', '工时', 'Total', 'Net Pay',
            '工时成本', '福利', 'Performance', 'Fixed年薪', '其他', '操作'
        ])
        self.emp_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.emp_table.setAlternatingRowColors(True)
        layout.addWidget(self.emp_table)

        return widget

    def create_charts_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        if not MATPLOTLIB_AVAILABLE:
            layout.addWidget(QLabel("matplotlib未安装，无法显示图表。\n运行: pip install matplotlib"))
            return widget

        ctrl_layout = QHBoxLayout()
        ctrl_layout.addWidget(QLabel("图表类型:"))
        self.chart_type = QComboBox()
        self.chart_type.addItems([
            '部门成本堆叠图', '成本占比饼图', '工时效率对比',
            '历史趋势折线图', '员工收入 breakdown'
        ])
        self.chart_type.currentTextChanged.connect(self.refresh_charts)
        ctrl_layout.addWidget(self.chart_type)

        ctrl_layout.addStretch()
        btn_export_chart = QPushButton("💾 导出图表")
        btn_export_chart.clicked.connect(self.export_chart)
        ctrl_layout.addWidget(btn_export_chart)

        layout.addLayout(ctrl_layout)

        self.figure = Figure(figsize=(12, 8), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)

        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)

        return widget

    def create_mapping_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        info = QLabel("分类Mapping说明:<br>系统根据关键词自动将工资单项目分类到以下5个类别。<br>你可以添加/删除关键词来调整分类规则。<br>如果某个项目无法匹配任何关键词，会被归类到\"其他\"。")
        info.setWordWrap(True)
        layout.addWidget(info)

        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(4)
        self.mapping_table.setHorizontalHeaderLabels([
            '类别ID', '类别名称', '关键词 (逗号分隔)', '描述'
        ])
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.mapping_table)

        btn_layout = QHBoxLayout()
        btn_save = QPushButton("💾 保存Mapping规则")
        btn_save.clicked.connect(self.save_mapping)
        btn_layout.addWidget(btn_save)

        btn_reset = QPushButton("🔄 恢复默认")
        btn_reset.clicked.connect(self.reset_mapping)
        btn_layout.addWidget(btn_reset)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        return widget

    def create_data_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        layout.addWidget(QLabel("<b>已导入的周:</b>"))
        self.history_list = QListWidget()
        self.history_list.itemClicked.connect(self.on_history_selected)
        layout.addWidget(self.history_list)

        btn_layout = QHBoxLayout()

        btn_delete = QPushButton("🗑️ 删除选中周")
        btn_delete.clicked.connect(self.delete_week)
        btn_layout.addWidget(btn_delete)

        btn_export_json = QPushButton("📤 导出JSON")
        btn_export_json.clicked.connect(lambda: self.export_data('json'))
        btn_layout.addWidget(btn_export_json)

        btn_export_excel = QPushButton("📊 导出Excel")
        btn_export_excel.clicked.connect(lambda: self.export_data('excel'))
        btn_layout.addWidget(btn_export_excel)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        layout.addWidget(QLabel("<b>员工列表 (点击修改部门):</b>"))
        self.data_emp_list = QTableWidget()
        self.data_emp_list.setColumnCount(4)
        self.data_emp_list.setHorizontalHeaderLabels(['姓名', '英文名', '当前部门', '修改部门'])
        self.data_emp_list.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.data_emp_list)

        return widget

    def import_pdf(self):
        if not PDF_AVAILABLE and not PYMUPDF_AVAILABLE:
            QMessageBox.warning(
                self, "缺少依赖",
                "请至少安装其一:\n  pip install pdfplumber\n  pip install pymupdf",
            )
            return

        files, _ = QFileDialog.getOpenFileNames(
            self, "选择PDF工资单", "", "PDF Files (*.pdf)"
        )
        if not files:
            return

        all_data = []
        failed_files = []
        for filepath in files:
            data, error = PDFPayrollParser.parse_pdf(filepath)
            if error:
                failed_files.append(f"{os.path.basename(filepath)}: {error}")
                continue
            if data:
                all_data.append(data)

        if failed_files:
            QMessageBox.warning(self, "部分解析失败", "\n".join(failed_files))

        if not all_data:
            QMessageBox.warning(self, "导入失败", "无法解析任何PDF文件")
            return

        # 显示解析结果预览（人多时勿拼接数万行再截断，避免卡顿）
        if len(all_data) > 350:
            preview = "解析成功:\n"
            preview += f"（共 {len(all_data)} 条；下方仅列出前 35 条，其余略 — 不影响导入）\n\n"
            for d in all_data[:35]:
                preview += f"  {d['name']}: ${d['total_earnings']:.2f} ({len(d['earnings'])}项收入)\n"
            preview += f"\n… 省略 {len(all_data) - 35} 条。\n"
        else:
            preview = "解析成功:\n"
            for d in all_data:
                preview += f"  {d['name']}: ${d['total_earnings']:.2f} ({len(d['earnings'])}项收入)\n"

        week_id, ok = self._prompt_week_id(
            preview, self.guess_week_id(all_data[0])
        )
        if not ok or not week_id:
            return

        self.current_week = week_id
        try:
            self._commit_payslip_week(all_data, week_id)
        except Exception as e:
            mb = QMessageBox(self)
            mb.setIcon(QMessageBox.Icon.Critical)
            mb.setWindowTitle("导入失败")
            mb.setText(str(e))
            mb.setDetailedText(traceback.format_exc())
            mb.exec()
            traceback.print_exc()

    def import_excel_payslips(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel工资簿（每位员工一个工作表）",
            "",
            "Excel Files (*.xlsx *.xls)",
        )
        if not filepath:
            return

        all_data, errors = ExcelPayslipImporter.import_workbook(filepath)
        if errors:
            msg = "\n".join(errors[:40])
            if len(errors) > 40:
                msg += f"\n… 其余 {len(errors) - 40} 条略"
            QMessageBox.warning(self, "部分工作表未导入", msg)

        if not all_data:
            QMessageBox.warning(self, "导入失败", "未能从任何工作表解析出工资单")
            return

        if len(all_data) > 350:
            preview = "解析成功:\n"
            preview += f"（共 {len(all_data)} 条；下方仅列出前 35 条，其余略 — 不影响导入）\n\n"
            for d in all_data[:35]:
                nm = d.get('name') or '(未识别姓名)'
                preview += (
                    f"  {nm}: ${d.get('total_earnings', 0):.2f} "
                    f"({len(d.get('earnings', []))}项收入)\n"
                )
            preview += f"\n… 省略 {len(all_data) - 35} 条。\n"
        else:
            preview = "解析成功:\n"
            for d in all_data:
                nm = d.get('name') or '(未识别姓名)'
                preview += (
                    f"  {nm}: ${d.get('total_earnings', 0):.2f} "
                    f"({len(d.get('earnings', []))}项收入)\n"
                )

        week_id, ok = self._prompt_week_id(
            preview, self.guess_week_id(all_data[0])
        )
        if not ok or not week_id:
            return

        self.current_week = week_id
        try:
            self._commit_payslip_week(all_data, week_id)
        except Exception as e:
            mb = QMessageBox(self)
            mb.setIcon(QMessageBox.Icon.Critical)
            mb.setWindowTitle("导入失败")
            mb.setText(str(e))
            mb.setDetailedText(traceback.format_exc())
            mb.exec()
            traceback.print_exc()

    def _commit_payslip_week(self, all_data, week_id):
        if week_id not in self.db.history:
            self.db.history[week_id] = {}

        for data in all_data:
            name = data['name']
            if not name:
                continue
            name_key = name.lower().strip()

            emp = self.db.get_or_create_employee(name, name_key)

            payroll = {
                'name': name,
                'pay_period': data.get('pay_period', ''),
                'payment_date': data.get('payment_date', ''),
                'total_earnings': data.get('total_earnings', 0),
                'net_pay': data.get('net_pay', 0),
                'hours': 0,
                'raw_earnings': data.get('earnings', [])
            }

            categorized = {'labor_cost': 0, 'benefits': 0, 'performance': 0,
                          'fixed_salary': 0, 'other': 0}

            for item in data.get('earnings', []):
                assigned = False
                item_name_lower = item['name'].lower()

                for cat_id, rule in self.db.mapping_rules.items():
                    keywords = rule.get('keywords', [])
                    if any(kw.lower() in item_name_lower for kw in keywords):
                        categorized[cat_id] += item['amount']
                        assigned = True
                        break

                if not assigned:
                    categorized['other'] += item['amount']

            payroll.update(categorized)
            self.db.history[week_id][name_key] = payroll

        self.db.save_data()
        self.status.showMessage(f"成功导入 {len(all_data)} 条工资单到 {week_id}")
        self.refresh_all()

    def import_excel(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "选择Excel工时表", "", "Excel Files (*.xlsx *.xls)"
        )
        if not filepath:
            return

        data, error = ExcelHoursImporter.import_hours(filepath)
        if error:
            QMessageBox.warning(self, "导入失败", f"错误: {error}")
            return

        updated = 0
        for item in data:
            name_key = item['full_name'].lower().strip()
            emp = self.db.get_or_create_employee(
                item['full_name'], name_key, item['branch']
            )
            emp.english_name = item['english_name']
            emp.branch = item['branch']

            if self.current_week and self.current_week in self.db.history:
                if name_key in self.db.history[self.current_week]:
                    self.db.history[self.current_week][name_key]['hours'] = item['hours']
                    updated += 1

        self.db.save_data()
        self.status.showMessage(f"导入 {len(data)} 条工时记录，更新 {updated} 条工资单工时")
        self.refresh_all()

    @staticmethod
    def _clip_preview_text(text: str, max_lines: int = 500, max_chars: int = 180_000) -> str:
        """超长预览会拖垮文本控件；截断只影响弹窗显示，不影响导入数据。"""
        text = text.replace('\x00', '')
        raw_lines = text.splitlines()
        n = len(raw_lines)
        if n > max_lines:
            head_n = max_lines // 2
            tail_n = max_lines - head_n - 2
            tail_n = max(1, tail_n)
            if head_n + tail_n > n:
                head_n = n // 2
                tail_n = n - head_n
            omitted = max(0, n - head_n - tail_n)
            lines = (
                [f"【共 {n} 行，预览仅显示前 {head_n} 行与后 {tail_n} 行（导入不受影响）】", ""]
                + raw_lines[:head_n]
                + [f"…… 省略约 {omitted} 行 ……"]
                + raw_lines[-tail_n:]
            )
            text = '\n'.join(lines)
        if len(text) > max_chars:
            keep = max(4000, max_chars // 2 - 120)
            text = (
                "【预览过长已截断字符（导入不受影响）】\n\n"
                + text[:keep]
                + "\n\n…… 中间省略 ……\n\n"
                + text[-keep:]
            )
        return text

    def _prompt_week_id(self, preview_text: str, default_week: str):
        """
        使用 QPlainTextEdit（大段纯文本比 QTextEdit 更省内存）；预览过长时截断。
        """
        safe = self._clip_preview_text(preview_text)
        dlg = QDialog(self)
        dlg.setWindowTitle("设置周ID")
        dlg.setMinimumWidth(560)
        dlg.resize(660, 520)

        layout = QVBoxLayout(dlg)
        tip = QLabel(
            "解析预览（纯文本；人数很多时会自动截断预览，不影响实际导入）。"
        )
        tip.setWordWrap(True)
        layout.addWidget(tip)

        viewer = QPlainTextEdit()
        viewer.setReadOnly(True)
        viewer.setUndoRedoEnabled(False)
        viewer.setPlainText(safe)
        viewer.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        viewer.setMinimumHeight(260)
        viewer.setMaximumHeight(400)
        layout.addWidget(viewer)

        layout.addWidget(QLabel("请输入周标识（例如 2026-W18）："))
        entry = QLineEdit()
        entry.setText(default_week)
        layout.addWidget(entry)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        entry.setFocus()
        entry.selectAll()

        result = dlg.exec()
        accepted = result == QDialog.DialogCode.Accepted
        return entry.text().strip(), accepted

    def guess_week_id(self, data):
        pay_period = data.get('pay_period', '')
        match = re.search(r'(\d{1,2})\s*(Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar)', pay_period)
        if match:
            from datetime import datetime
            now = datetime.now()
            return f"{now.year}-W{now.isocalendar()[1]}"
        return "2026-W18"

    def on_week_changed(self, week_id):
        if week_id and week_id != "全部":
            self.current_week = week_id
            self.week_label.setText(f"当前周: {week_id}")
            self.refresh_summary()
            self.refresh_employees()

    def refresh_all(self):
        self.week_combo.clear()
        weeks = sorted(self.db.history.keys(), reverse=True)
        self.week_combo.addItem("全部")
        self.week_combo.addItems(weeks)

        self.emp_branch_filter.clear()
        self.emp_branch_filter.addItem("全部")
        self.emp_branch_filter.addItems(sorted(self.db.branches))

        self.refresh_summary()
        self.refresh_employees()
        self.refresh_mapping()
        self.refresh_data_tab()
        if MATPLOTLIB_AVAILABLE:
            self.refresh_charts()

    def refresh_summary(self):
        week_id = self.current_week if self.current_week else None
        summary = self.db.get_branch_summary(week_id)

        self.summary_table.setRowCount(len(summary))

        totals = {'labor_cost': 0, 'benefits': 0, 'performance': 0,
                 'fixed_salary': 0, 'other': 0, 'total_earnings': 0,
                 'hours': 0, 'employee_count': 0, 'net_pay': 0}

        for i, (branch, data) in enumerate(sorted(summary.items())):
            self.summary_table.setItem(i, 0, QTableWidgetItem(branch))
            self.summary_table.setItem(i, 1, QTableWidgetItem(str(data['employee_count'])))
            self.summary_table.setItem(i, 2, QTableWidgetItem(f"{data['hours']:.1f}"))
            self.summary_table.setItem(i, 3, QTableWidgetItem(f"${data['labor_cost']:.2f}"))
            self.summary_table.setItem(i, 4, QTableWidgetItem(f"${data['benefits']:.2f}"))
            self.summary_table.setItem(i, 5, QTableWidgetItem(f"${data['performance']:.2f}"))
            self.summary_table.setItem(i, 6, QTableWidgetItem(f"${data['fixed_salary']:.2f}"))
            self.summary_table.setItem(i, 7, QTableWidgetItem(f"${data['other']:.2f}"))
            self.summary_table.setItem(i, 8, QTableWidgetItem(f"${data['total_earnings']:.2f}"))
            self.summary_table.setItem(i, 9, QTableWidgetItem(f"${data['net_pay']:.2f}"))

            for k in totals:
                totals[k] += data[k]

        self.total_label.setText(
            f"总计: {totals['employee_count']}人 | "
            f"工时: {totals['hours']:.1f}h | "
            f"工时成本: ${totals['labor_cost']:.2f} | "
            f"福利: ${totals['benefits']:.2f} | "
            f"Performance: ${totals['performance']:.2f} | "
            f"Fixed年薪: ${totals['fixed_salary']:.2f} | "
            f"其他: ${totals['other']:.2f} | "
            f"Total: ${totals['total_earnings']:.2f}"
        )

    def refresh_employees(self):
        branch_filter = self.emp_branch_filter.currentText()
        search_text = self.emp_search.text().lower()
        week_id = self.current_week

        rows = []
        for emp_key, emp in self.db.employees.items():
            if branch_filter != "全部" and emp.branch != branch_filter:
                continue
            if search_text and search_text not in emp.name.lower() and search_text not in emp.english_name.lower():
                continue

            payroll = None
            if week_id and week_id in self.db.history and emp_key in self.db.history[week_id]:
                payroll = self.db.history[week_id][emp_key]

            rows.append({
                'name': emp.name,
                'english_name': emp.english_name,
                'branch': emp.branch,
                'hours': payroll.get('hours', 0) if payroll else 0,
                'total': payroll.get('total_earnings', 0) if payroll else 0,
                'net_pay': payroll.get('net_pay', 0) if payroll else 0,
                'labor_cost': payroll.get('labor_cost', 0) if payroll else 0,
                'benefits': payroll.get('benefits', 0) if payroll else 0,
                'performance': payroll.get('performance', 0) if payroll else 0,
                'fixed_salary': payroll.get('fixed_salary', 0) if payroll else 0,
                'other': payroll.get('other', 0) if payroll else 0,
                'emp_key': emp_key
            })

        self.emp_table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            self.emp_table.setItem(i, 0, QTableWidgetItem(row['name']))
            self.emp_table.setItem(i, 1, QTableWidgetItem(row['english_name']))

            branch_combo = QComboBox()
            branch_combo.addItems(sorted(self.db.branches))
            branch_combo.setCurrentText(row['branch'])
            branch_combo.currentTextChanged.connect(
                lambda text, key=row['emp_key']: self.on_branch_changed(key, text)
            )
            self.emp_table.setCellWidget(i, 2, branch_combo)

            self.emp_table.setItem(i, 3, QTableWidgetItem(f"{row['hours']:.1f}"))
            self.emp_table.setItem(i, 4, QTableWidgetItem(f"${row['total']:.2f}"))
            self.emp_table.setItem(i, 5, QTableWidgetItem(f"${row['net_pay']:.2f}"))
            self.emp_table.setItem(i, 6, QTableWidgetItem(f"${row['labor_cost']:.2f}"))
            self.emp_table.setItem(i, 7, QTableWidgetItem(f"${row['benefits']:.2f}"))
            self.emp_table.setItem(i, 8, QTableWidgetItem(f"${row['performance']:.2f}"))
            self.emp_table.setItem(i, 9, QTableWidgetItem(f"${row['fixed_salary']:.2f}"))
            self.emp_table.setItem(i, 10, QTableWidgetItem(f"${row['other']:.2f}"))

            btn = QPushButton("🔍")
            btn.setMaximumWidth(40)
            btn.clicked.connect(lambda checked, key=row['emp_key']: self.show_employee_detail(key))
            self.emp_table.setCellWidget(i, 11, btn)

    def on_branch_changed(self, emp_key, new_branch):
        self.db.set_employee_branch(emp_key, new_branch)
        self.refresh_summary()
        self.status.showMessage(f"已更新 {self.db.employees[emp_key].name} 的部门为 {new_branch}")

    def refresh_charts(self):
        if not MATPLOTLIB_AVAILABLE:
            return

        chart_type = self.chart_type.currentText()
        self.figure.clear()
        ax = self.figure.add_subplot(111)

        if chart_type == '部门成本堆叠图':
            self.plot_stacked_bar(ax)
        elif chart_type == '成本占比饼图':
            self.plot_pie(ax)
        elif chart_type == '工时效率对比':
            self.plot_efficiency(ax)
        elif chart_type == '历史趋势折线图':
            self.plot_trend(ax)
        elif chart_type == '员工收入 breakdown':
            self.plot_employee_breakdown(ax)

        self.canvas.draw()

    def plot_stacked_bar(self, ax):
        week_id = self.current_week
        summary = self.db.get_branch_summary(week_id)

        branches = list(summary.keys())
        labor = [summary[b]['labor_cost'] for b in branches]
        benefits = [summary[b]['benefits'] for b in branches]
        performance = [summary[b]['performance'] for b in branches]
        fixed = [summary[b]['fixed_salary'] for b in branches]
        other = [summary[b]['other'] for b in branches]

        x = range(len(branches))
        ax.bar(x, labor, label='工时成本', color='#2E86AB')
        ax.bar(x, benefits, bottom=labor, label='福利', color='#A23B72')
        ax.bar(x, performance, bottom=[l+b for l,b in zip(labor, benefits)], label='Performance', color='#F18F01')
        ax.bar(x, fixed, bottom=[l+b+p for l,b,p in zip(labor, benefits, performance)], label='Fixed年薪', color='#C73E1D')
        ax.bar(x, other, bottom=[l+b+p+f for l,b,p,f in zip(labor, benefits, performance, fixed)], label='其他', color='#6A4C93')

        ax.set_xticks(x)
        ax.set_xticklabels(branches, rotation=45, ha='right')
        ax.set_ylabel('Amount ($NZD)')
        ax.set_title('Cost Breakdown by Branch')
        ax.legend()
        ax.grid(axis='y', alpha=0.3)

    def plot_pie(self, ax):
        week_id = self.current_week
        summary = self.db.get_branch_summary(week_id)

        totals = {
            '工时成本': sum(s['labor_cost'] for s in summary.values()),
            '福利': sum(s['benefits'] for s in summary.values()),
            'Performance': sum(s['performance'] for s in summary.values()),
            'Fixed年薪': sum(s['fixed_salary'] for s in summary.values()),
            '其他': sum(s['other'] for s in summary.values())
        }

        colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A4C93']
        ax.pie(totals.values(), labels=totals.keys(), autopct='%1.1f%%', 
               colors=colors, startangle=90)
        ax.set_title('Overall Cost Distribution')

    def plot_trend(self, ax):
        weeks = sorted(self.db.history.keys())
        if len(weeks) < 2:
            ax.text(0.5, 0.5, '需要至少2周数据才能显示趋势', ha='center', va='center')
            return

        categories = ['labor_cost', 'benefits', 'performance', 'fixed_salary', 'other']
        colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A4C93']
        labels = ['工时成本', '福利', 'Performance', 'Fixed年薪', '其他']

        for cat, color, label in zip(categories, colors, labels):
            values = []
            for wk in weeks:
                total = sum(p.get(cat, 0) for p in self.db.history[wk].values())
                values.append(total)
            ax.plot(weeks, values, marker='o', label=label, color=color, linewidth=2)

        ax.set_xlabel('Week')
        ax.set_ylabel('Amount ($NZD)')
        ax.set_title('Historical Cost Trend')
        ax.legend()
        ax.grid(alpha=0.3)
        import matplotlib.pyplot as plt
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)

    def plot_efficiency(self, ax):
        week_id = self.current_week
        summary = self.db.get_branch_summary(week_id)

        eff_data = [(b, s['total_earnings']/s['hours']) for b, s in summary.items() if s['hours'] > 0]
        if not eff_data:
            ax.text(0.5, 0.5, '无工时数据', ha='center', va='center')
            return

        branches, efficiencies = zip(*sorted(eff_data, key=lambda x: x[1]))
        bars = ax.bar(branches, efficiencies, color='#2E86AB', alpha=0.8)
        ax.set_ylabel('Cost per Hour ($NZD)')
        ax.set_title('Labor Cost Efficiency by Branch')
        ax.grid(axis='y', alpha=0.3)

        for bar, val in zip(bars, efficiencies):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                   f'${val:.2f}', ha='center', va='bottom', fontsize=9)

    def plot_employee_breakdown(self, ax):
        week_id = self.current_week
        if not week_id or week_id not in self.db.history:
            ax.text(0.5, 0.5, '请先选择一周', ha='center', va='center')
            return

        emp_data = []
        for emp_key, payroll in self.db.history[week_id].items():
            emp = self.db.employees.get(emp_key)
            if emp and payroll.get('total_earnings', 0) > 0:
                emp_data.append({
                    'name': emp.name[:15],
                    'labor': payroll.get('labor_cost', 0),
                    'benefits': payroll.get('benefits', 0),
                    'performance': payroll.get('performance', 0),
                    'fixed': payroll.get('fixed_salary', 0),
                    'other': payroll.get('other', 0)
                })

        emp_data.sort(key=lambda x: sum([x['labor'], x['benefits'], x['performance'], x['fixed'], x['other']]))

        y_pos = range(len(emp_data))
        left1 = [e['labor'] for e in emp_data]
        left2 = [l+b for l,b in zip(left1, [e['benefits'] for e in emp_data])]
        left3 = [l+p for l,p in zip(left2, [e['performance'] for e in emp_data])]
        left4 = [l+f for l,f in zip(left3, [e['fixed'] for e in emp_data])]

        ax.barh(y_pos, left1, color='#2E86AB', label='Labor')
        ax.barh(y_pos, [e['benefits'] for e in emp_data], left=left1, color='#A23B72', label='Benefits')
        ax.barh(y_pos, [e['performance'] for e in emp_data], left=left2, color='#F18F01', label='Performance')
        ax.barh(y_pos, [e['fixed'] for e in emp_data], left=left3, color='#C73E1D', label='Fixed')
        ax.barh(y_pos, [e['other'] for e in emp_data], left=left4, color='#6A4C93', label='Other')

        ax.set_yticks(y_pos)
        ax.set_yticklabels([e['name'] for e in emp_data], fontsize=8)
        ax.set_xlabel('Amount ($NZD)')
        ax.set_title('Employee Earnings Breakdown')
        ax.legend(loc='lower right')
        ax.grid(axis='x', alpha=0.3)

    def refresh_mapping(self):
        rules = self.db.mapping_rules
        self.mapping_table.setRowCount(len(rules))

        for i, (cat_id, rule) in enumerate(rules.items()):
            self.mapping_table.setItem(i, 0, QTableWidgetItem(cat_id))
            self.mapping_table.setItem(i, 1, QTableWidgetItem(rule.get('name', '')))

            keywords = ', '.join(rule.get('keywords', []))
            self.mapping_table.setItem(i, 2, QTableWidgetItem(keywords))
            self.mapping_table.setItem(i, 3, QTableWidgetItem(rule.get('description', '')))

    def refresh_data_tab(self):
        self.history_list.clear()
        for wk in sorted(self.db.history.keys(), reverse=True):
            count = len(self.db.history[wk])
            item = QListWidgetItem(f"{wk} ({count}人)")
            item.setData(Qt.ItemDataRole.UserRole, wk)
            self.history_list.addItem(item)

        self.data_emp_list.setRowCount(len(self.db.employees))
        for i, (key, emp) in enumerate(sorted(self.db.employees.items())):
            self.data_emp_list.setItem(i, 0, QTableWidgetItem(emp.name))
            self.data_emp_list.setItem(i, 1, QTableWidgetItem(emp.english_name))
            self.data_emp_list.setItem(i, 2, QTableWidgetItem(emp.branch))

            combo = QComboBox()
            combo.addItems(sorted(self.db.branches))
            combo.setCurrentText(emp.branch)
            combo.currentTextChanged.connect(lambda text, k=key: self.db.set_employee_branch(k, text))
            self.data_emp_list.setCellWidget(i, 3, combo)

    def edit_mapping(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑分类Mapping规则")
        dialog.setGeometry(200, 200, 600, 400)

        layout = QVBoxLayout(dialog)

        info = QLabel("修改关键词后点击保存，系统会自动重新分类所有历史数据。")
        layout.addWidget(info)

        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(['类别', '关键词 (逗号分隔)', '描述'])
        table.setRowCount(len(self.db.mapping_rules))

        for i, (cat_id, rule) in enumerate(self.db.mapping_rules.items()):
            table.setItem(i, 0, QTableWidgetItem(cat_id))
            table.setItem(i, 1, QTableWidgetItem(', '.join(rule.get('keywords', []))))
            table.setItem(i, 2, QTableWidgetItem(rule.get('description', '')))

        layout.addWidget(table)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            for i in range(table.rowCount()):
                cat_id = table.item(i, 0).text()
                keywords = [k.strip() for k in table.item(i, 1).text().split(',')]
                desc = table.item(i, 2).text()
                if cat_id in self.db.mapping_rules:
                    self.db.mapping_rules[cat_id]['keywords'] = keywords
                    self.db.mapping_rules[cat_id]['description'] = desc

            self.save_config()
            self.db.save_data()
            self.refresh_mapping()
            QMessageBox.information(self, "保存成功", "Mapping规则已更新，请重新导入PDF以应用新规则。")

    def manage_branches(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("部门管理")
        dialog.setGeometry(200, 200, 400, 300)

        layout = QVBoxLayout(dialog)

        list_widget = QListWidget()
        for b in sorted(self.db.branches):
            list_widget.addItem(b)
        layout.addWidget(list_widget)

        input_layout = QHBoxLayout()
        new_branch = QLineEdit()
        new_branch.setPlaceholderText("新部门名称")
        input_layout.addWidget(new_branch)

        btn_add = QPushButton("添加")
        btn_add.clicked.connect(lambda: self.add_branch(new_branch.text(), list_widget))
        input_layout.addWidget(btn_add)

        btn_del = QPushButton("删除选中")
        btn_del.clicked.connect(lambda: self.delete_branch(list_widget))
        input_layout.addWidget(btn_del)

        layout.addLayout(input_layout)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        dialog.exec()

    def add_branch(self, name, list_widget):
        if name and name not in self.db.branches:
            self.db.branches.add(name)
            list_widget.addItem(name)
            self.db.save_data()

    def delete_branch(self, list_widget):
        item = list_widget.currentItem()
        if item:
            self.db.branches.discard(item.text())
            list_widget.takeItem(list_widget.row(item))
            self.db.save_data()

    def batch_assign_branch(self):
        selected = []
        for item in self.emp_table.selectedItems():
            row = item.row()
            emp_key = self.emp_table.item(row, 0).text().lower().strip()
            selected.append(emp_key)

        if not selected:
            QMessageBox.information(self, "提示", "请先选择要分配部门的员工行")
            return

        branch, ok = QInputDialog.getItem(self, "批量分配部门", 
            "选择部门:", sorted(self.db.branches), 0, False)
        if ok and branch:
            for emp_key in selected:
                if emp_key in self.db.employees:
                    self.db.employees[emp_key].branch = branch
            self.db.save_data()
            self.refresh_all()
            self.status.showMessage(f"已批量分配 {len(selected)} 人到 {branch}")

    def show_employee_detail(self, emp_key):
        emp = self.db.employees.get(emp_key)
        if not emp:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle(f"员工详情: {emp.name}")
        dialog.setGeometry(200, 200, 500, 400)

        layout = QVBoxLayout(dialog)

        info_text = f"<b>姓名:</b> {emp.name}<br><b>英文名:</b> {emp.english_name}<br><b>部门:</b> {emp.branch}<br><b>历史周数:</b> {len(emp.history)}<br>"
        info = QLabel(info_text)
        info.setWordWrap(True)
        layout.addWidget(info)

        hist_table = QTableWidget()
        hist_table.setColumnCount(6)
        hist_table.setHorizontalHeaderLabels(['周', 'Total', 'Net Pay', '工时成本', '福利', 'Performance'])

        rows = []
        for wk, data in self.db.history.items():
            if emp_key in data:
                p = data[emp_key]
                rows.append((wk, p.get('total_earnings', 0), p.get('net_pay', 0),
                           p.get('labor_cost', 0), p.get('benefits', 0), p.get('performance', 0)))

        rows.sort()
        hist_table.setRowCount(len(rows))
        for i, (wk, total, net, labor, benefits, perf) in enumerate(rows):
            hist_table.setItem(i, 0, QTableWidgetItem(wk))
            hist_table.setItem(i, 1, QTableWidgetItem(f"${total:.2f}"))
            hist_table.setItem(i, 2, QTableWidgetItem(f"${net:.2f}"))
            hist_table.setItem(i, 3, QTableWidgetItem(f"${labor:.2f}"))
            hist_table.setItem(i, 4, QTableWidgetItem(f"${benefits:.2f}"))
            hist_table.setItem(i, 5, QTableWidgetItem(f"${perf:.2f}"))

        layout.addWidget(hist_table)

        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(dialog.accept)
        layout.addWidget(btn_close)

        dialog.exec()

    def export_report(self):
        filepath, _ = QFileDialog.getSaveFileName(
            self, "导出报告", "payroll_report.xlsx", "Excel Files (*.xlsx)"
        )
        if not filepath:
            return

        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                week_id = self.current_week
                summary = self.db.get_branch_summary(week_id)
                df_summary = pd.DataFrame.from_dict(summary, orient='index')
                df_summary.to_excel(writer, sheet_name='部门汇总')

                emp_rows = []
                for emp_key, emp in self.db.employees.items():
                    if week_id and week_id in self.db.history and emp_key in self.db.history[week_id]:
                        p = self.db.history[week_id][emp_key]
                        emp_rows.append({
                            '姓名': emp.name,
                            '英文名': emp.english_name,
                            '部门': emp.branch,
                            '工时': p.get('hours', 0),
                            'Total': p.get('total_earnings', 0),
                            'Net Pay': p.get('net_pay', 0),
                            '工时成本': p.get('labor_cost', 0),
                            '福利': p.get('benefits', 0),
                            'Performance': p.get('performance', 0),
                            'Fixed年薪': p.get('fixed_salary', 0),
                            '其他': p.get('other', 0)
                        })
                df_emp = pd.DataFrame(emp_rows)
                df_emp.to_excel(writer, sheet_name='员工明细', index=False)

            QMessageBox.information(self, "导出成功", f"报告已保存到:\n{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def export_data(self, format_type):
        if format_type == 'json':
            filepath, _ = QFileDialog.getSaveFileName(
                self, "导出JSON", "payroll_data.json", "JSON Files (*.json)"
            )
            if filepath:
                self.db.save_data()
                QMessageBox.information(self, "导出成功", f"数据已保存到:\n{filepath}")
        else:
            self.export_report()

    def export_chart(self):
        if not MATPLOTLIB_AVAILABLE:
            return
        filepath, _ = QFileDialog.getSaveFileName(
            self, "保存图表", "chart.png", "PNG Images (*.png);;PDF Files (*.pdf)"
        )
        if filepath:
            self.figure.savefig(filepath, dpi=150, bbox_inches='tight')
            QMessageBox.information(self, "保存成功", f"图表已保存到:\n{filepath}")

    def delete_week(self):
        item = self.history_list.currentItem()
        if not item:
            return

        week_id = item.data(Qt.ItemDataRole.UserRole)
        reply = QMessageBox.question(self, "确认删除", 
            f"确定要删除 {week_id} 的所有数据吗?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            if week_id in self.db.history:
                del self.db.history[week_id]
                self.db.save_data()
                self.refresh_all()
                self.status.showMessage(f"已删除 {week_id}")

    def on_history_selected(self, item):
        week_id = item.data(Qt.ItemDataRole.UserRole)
        self.current_week = week_id
        self.week_label.setText(f"当前周: {week_id}")
        self.week_combo.setCurrentText(week_id)
        self.refresh_summary()
        self.refresh_employees()

    def save_mapping(self):
        self.save_config()
        self.db.save_data()
        QMessageBox.information(self, "保存成功", "Mapping规则已保存")

    def reset_mapping(self):
        reply = QMessageBox.question(self, "确认", 
            "确定要恢复默认Mapping规则吗? 当前自定义规则将丢失。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.config['category_mapping'] = DEFAULT_CONFIG['category_mapping']
            self.db.mapping_rules = DEFAULT_CONFIG['category_mapping']
            self.save_config()
            self.db.save_data()
            self.refresh_mapping()

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    app.setStyleSheet("""
        QMainWindow { background-color: #f5f5f5; }
        QTabWidget::pane { border: 1px solid #ddd; background: white; }
        QTabBar::tab { padding: 10px 20px; background: #e0e0e0; border: 1px solid #ddd; border-bottom: none; }
        QTabBar::tab:selected { background: white; border-top: 3px solid #2E86AB; }
        QPushButton { background-color: #2E86AB; color: white; border: none; padding: 6px 12px; border-radius: 4px; }
        QPushButton:hover { background-color: #1a5276; }
        QTableWidget { border: 1px solid #ddd; gridline-color: #eee; }
        QComboBox { padding: 4px; border: 1px solid #ddd; }
    """)

    window = PayrollAnalyzer()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
