#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared heuristics for Xero-style payslip employee display names (no PyQt)."""

import re


def payslip_name_is_plausible(line: str) -> bool:
    """
    Reject payroll labels (Pay Frequency…), lines starting with pay/payment,
    obvious NZ address fragments, and street-style tokens.
    """
    line = (line or "").strip()
    if len(line) < 3 or len(line) > 70:
        return False
    low = line.lower()
    toxic = (
        "employment details",
        "pay frequency",
        "pay period:",
        "payment date:",
        "total earnings",
        "net pay:",
        "gross pay",
        "tax code:",
        "tax period:",
        "ird number",
        "ird no",
        "annual leave",
        "sick leave",
        "ordinary time",
        "kiwisaver",
        "deduction",
        "quantity rate",
        "this pay",
        " ytd",
    )
    if any(t in low for t in toxic):
        return False
    if re.match(r"^pay\s", low) or re.match(r"^payment\s", low):
        return False
    if re.search(r"\d{4,}", line):
        return False
    if re.match(r"^\d+\s+[A-Za-z]", line):
        return False
    if re.search(
        r"\b("
        r"auckland|christchurch|wellington|hamilton|tauranga|dunedin|"
        r"napier|hastings|rotorua|whangarei|invercargill|nelson|gisborne|"
        r"blenheim|timaru|porirua|queenstown|pakuranga|onehunga|"
        r"sockburn|newmarket|ponsonby|remuera|manukau|northshore"
        r")\b",
        low,
    ):
        return False
    if re.search(r"\b(?:palmerston\s+north|new\s+plymouth)\b", low):
        return False

    parts = line.split()
    if len(parts) >= 3 and re.search(
        r"\b(place|road|street|crescent|drive|lane|avenue|boulevard|"
        r"highway|heights|ridge|terrace|mews|court|close|park|square)\s*$",
        low,
    ):
        return False
    if line == line.upper() and 2 <= len(parts) <= 3 and re.search(
        r"\b(PLACE|ROAD|STREET|CRESCENT|AVENUE|DRIVE|LANE|COURT|CLOSE|HIGHWAY)\s*$",
        line,
    ):
        return False

    if ":" in line and len(line) < 55:
        return False
    if re.match(r"^[\d\s$,.%-]+$", line):
        return False
    return bool(
        re.match(
            r"^[\s'A-Za-z\u0080-\u024f\u4e00-\u9fff]"
            r"[\s'A-Za-z\u0080-\u024f\u4e00-\u9fff.\-]{2,}$",
            line,
        )
    )
