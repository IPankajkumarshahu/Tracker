#!/usr/bin/env python3
"""Convert bank statement PDFs to Excel.

Columns: Date | Value Date | Description | Chq./Ref. No. | Debit (Dr) | Credit (Cr) | Balance | Balance Check

Usage:
    python bank_statement_to_excel.py statement.pdf                 # -> statement.xlsx
    python bank_statement_to_excel.py a.pdf b.pdf -o client.xlsx    # both in one workbook
    python bank_statement_to_excel.py statements/                   # every PDF in a folder
    python bank_statement_to_excel.py locked.pdf -p PASSWORD        # password-protected PDF
"""

import argparse
import os
import sys

from bank_statement import StatementError, parse_statement, write_workbook


def _pdfs(inputs):
    for item in inputs:
        if os.path.isdir(item):
            for name in sorted(os.listdir(item)):
                if name.lower().endswith(".pdf"):
                    yield os.path.join(item, name)
        else:
            yield item


def main(argv=None):
    ap = argparse.ArgumentParser(description="Convert bank statement PDFs to Excel.")
    ap.add_argument("inputs", nargs="+", help="PDF files or folders of PDFs")
    ap.add_argument("-o", "--output",
                    help="one Excel file for all statements (default: one .xlsx next to each PDF)")
    ap.add_argument("-p", "--password", help="PDF password, if the statement is locked")
    args = ap.parse_args(argv)

    parsed, failed = [], 0
    for pdf in _pdfs(args.inputs):
        try:
            stmt = parse_statement(pdf, password=args.password)
        except (StatementError, FileNotFoundError) as e:
            print(f"✗ {pdf}: {e}", file=sys.stderr)
            failed += 1
            continue
        print(f"✓ {pdf}: {stmt.bank or 'bank not detected'}, {len(stmt.transactions)} transactions, "
              f"Dr {stmt.total_debit:,.2f} / Cr {stmt.total_credit:,.2f}")
        for w in stmt.warnings:
            print(f"    ! {w}")
        parsed.append((pdf, stmt))
        if not args.output:
            out = os.path.splitext(pdf)[0] + ".xlsx"
            write_workbook([stmt], out)
            print(f"    → {out}")

    if args.output and parsed:
        write_workbook([s for _, s in parsed], args.output)
        print(f"→ {args.output}")
    return 1 if failed and not parsed else 0


if __name__ == "__main__":
    sys.exit(main())
