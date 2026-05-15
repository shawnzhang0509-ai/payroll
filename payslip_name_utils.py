#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared heuristics for Xero-style payslip employee display names (no PyQt)."""

import re

_STREET_SUFFIX = re.compile(
    r"\b("
    r"road|rd\.?|street|st\.?|terrace|drive|dr\.?|crescent|cr\.?|"
    r"avenue|ave\.?|way|place|court|ct\.?|close|lane|rise|grove|mews|"
    r"crest|heights|ridge|boulevard|motorway|highway|parkway|"
    r"esplanade|parade|circuit|loop|square"
    r")\s*$",
    re.I,
)


def payslip_name_is_plausible(line: str) -> bool:
    """
    Reject payroll labels (Pay Frequency…), lines starting with pay/payment,
    obvious NZ address fragments, and street-style tokens (incl. two-word
    addresses like "Andrews Terrace").
    """
    line = (line or "").strip()
    if len(line) < 3 or len(line) > 70:
        return False
    mono = re.sub(r"\s+", " ", line.lower())

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
        "superannuation",
        "kiwi saver",
    )
    if any(t in mono for t in toxic):
        return False
    if re.match(r"^pay\s", mono) or re.match(r"^payment\s", mono):
        return False
    if re.match(r"^tax\s", mono) or re.match(r"^ird\s", mono):
        return False
    if re.match(r"^(weekly|fortnightly|monthly)\s*$", mono):
        return False

    if re.search(r"\d\s*/\s*\d", line):
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
        r"sockburn|newmarket|ponsonby|remuera|manukau|northshore|"
        r"mount\s+eden|mt\s+eden|penrose|epsom|grey\s+lynn|"
        r"mission\s+bay|howick|botany|papakura|pukekohe|"
        r"riccarton|hornby|rolleston|linwood|"
        r"te\s+aro|kilbirnie|johnsonville|petone|lower\s+hutt|upper\s+hutt"
        r")\b",
        mono,
    ):
        return False
    if re.search(r"\b(?:palmerston\s+north|new\s+plymouth)\b", mono):
        return False

    parts = line.split()
    if len(parts) >= 2 and _STREET_SUFFIX.search(line):
        return False

    if len(parts) >= 3 and re.search(
        r"\b(place|road|street|crescent|drive|lane|avenue|boulevard|"
        r"highway|heights|ridge|terrace|mews|court|close|park|square)\s*$",
        mono,
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


def payslip_name_pick_from_cell_text(raw: str) -> str | None:
    """
    From a merged Excel cell (name + flat number + street lines), return the
    first plausible employee name line, preferring text before '1/67' style.
    """
    raw = (raw or "").replace("\r", "\n")
    for line in raw.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^(.{2,60}?)\s+\d+\s*/\s*\d+", line)
        if m:
            inner = m.group(1).strip()
            if payslip_name_is_plausible(inner):
                return inner
        if payslip_name_is_plausible(line):
            return line
    return None
