import csv
import importlib.util
from collections import Counter
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location(
    "reconcile_demo", ROOT / "scripts" / "reconcile_demo.py"
)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


def load(name):
    with open(ROOT / "demo" / name, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


@pytest.fixture(scope="module")
def report():
    return mod.reconcile(load("bank_statement_demo.csv"), load("1c_export_demo.csv"))


def test_expected_discrepancies(report):
    counts = Counter(r["status"] for r in report)
    assert counts[mod.STATUS_MATCHED] == 7
    assert counts[mod.STATUS_MISSING_IN_1C] == 1
    assert counts[mod.STATUS_AMOUNT_MISMATCH] == 1
    assert counts[mod.STATUS_DATE_MISMATCH] == 1
    assert counts[mod.STATUS_MISSING_IN_BANK] == 1
    assert len(report) == 11


def test_specific_mismatches(report):
    by_doc = {r["doc_no"]: r for r in report}
    assert by_doc["BS-1003"]["status"] == mod.STATUS_MISSING_IN_1C
    assert by_doc["BS-1005"]["status"] == mod.STATUS_AMOUNT_MISMATCH
    assert by_doc["BS-1007"]["status"] == mod.STATUS_DATE_MISMATCH
    assert by_doc["ZH-2001"]["status"] == mod.STATUS_MISSING_IN_BANK


@pytest.mark.parametrize("raw,expected", [
    ("15.03.2025", "2025-03-15"),
    ("2025-03-15", "2025-03-15"),
    ("15/03/2025", "2025-03-15"),
])
def test_parse_date_formats(raw, expected):
    assert mod.parse_date(raw).isoformat() == expected


def test_parse_date_invalid():
    assert mod.parse_date("not-a-date") is None
    assert mod.parse_date("") is None
