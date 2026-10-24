#!/usr/bin/env python3
"""Демо сверки банковской выписки и выгрузки 1С (синтетические данные)."""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

DATE_FORMATS = ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y")
AMOUNT_TOLERANCE = 0.01

STATUS_MATCHED = "MATCHED"
STATUS_MISSING_IN_1C = "MISSING_IN_1C"
STATUS_MISSING_IN_BANK = "MISSING_IN_BANK"
STATUS_AMOUNT_MISMATCH = "AMOUNT_MISMATCH"
STATUS_DATE_MISMATCH = "DATE_MISMATCH"

REPORT_FIELDS = [
    "doc_no", "counterparty", "status",
    "bank_date", "c1c_date", "bank_amount", "c1c_amount", "comment",
]


def parse_date(value: str):
    value = (value or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def load_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def index_by_doc(rows):
    return {(r.get("doc_no") or "").strip().upper(): r for r in rows}


def reconcile(bank_rows, c1c_rows):
    bank_by_doc = index_by_doc(bank_rows)
    c1c_by_doc = index_by_doc(c1c_rows)

    report = []
    for doc_no, b in bank_by_doc.items():
        c = c1c_by_doc.get(doc_no)
        row = {
            "doc_no": doc_no,
            "counterparty": b.get("counterparty", ""),
            "bank_date": b.get("date", ""),
            "bank_amount": b.get("amount", ""),
        }
        if c is None:
            row.update(status=STATUS_MISSING_IN_1C, comment="Документ есть в банке, отсутствует в 1С")
        else:
            row["c1c_date"] = c.get("date", "")
            row["c1c_amount"] = c.get("amount", "")
            if abs(float(b["amount"]) - float(c["amount"])) > AMOUNT_TOLERANCE:
                row.update(status=STATUS_AMOUNT_MISMATCH, comment="Расхождение сумм")
            else:
                b_dt, c_dt = parse_date(b["date"]), parse_date(c["date"])
                if b_dt and c_dt and b_dt != c_dt:
                    row.update(status=STATUS_DATE_MISMATCH, comment="Расхождение дат")
                else:
                    row.update(status=STATUS_MATCHED, comment="")
        report.append(row)

    for doc_no, c in c1c_by_doc.items():
        if doc_no not in bank_by_doc:
            report.append({
                "doc_no": doc_no,
                "counterparty": c.get("counterparty", ""),
                "bank_date": "", "bank_amount": "",
                "c1c_date": c.get("date", ""), "c1c_amount": c.get("amount", ""),
                "status": STATUS_MISSING_IN_BANK,
                "comment": "Документ есть в 1С, отсутствует в банке",
            })
    return report


def main():
    demo = Path(__file__).resolve().parent.parent / "demo"
    report = reconcile(load_csv(demo / "bank_statement_demo.csv"),
                       load_csv(demo / "1c_export_demo.csv"))
    out = demo / "report_demo.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=REPORT_FIELDS)
        writer.writeheader()
        for row in report:
            writer.writerow({k: row.get(k, "") for k in REPORT_FIELDS})
    counts = {}
    for row in report:
        counts[row["status"]] = counts.get(row["status"], 0) + 1
    print(f"Всего документов: {len(report)}")
    for status, n in sorted(counts.items()):
        print(f"  {status}: {n}")
    print(f"Отчёт: {out}")


if __name__ == "__main__":
    main()
