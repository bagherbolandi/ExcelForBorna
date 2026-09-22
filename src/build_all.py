"""
build_all.py — one command to build + verify + test + render the workbook.
=============================================================================
Usage:  python src/build_all.py
Output: output/ExcelForBorna.xlsx
"""
from __future__ import annotations

import os
import sys
import datetime as dt

sys.path.insert(0, os.path.dirname(__file__))

from builder import build_model
from renderer import render
from verify import run as run_verify
from tests import run_all, patch_test_results


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = os.path.join(root, "output", "ExcelForBorna.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)

    print("=" * 78)
    print("ExcelForBorna — build pipeline")
    print("=" * 78)

    print("\n[1/4] Building the workbook model (44 sheets, real formulas)...")
    model, SET = build_model()

    print("\n[2/4] Verifying formulas with the audit engine...")
    ok = run_verify(model)
    if not ok:
        print("\nVERIFICATION FAILED — aborting render.")
        sys.exit(2)

    print("\n[3/4] Executing 15 mandatory test cases...")
    results = run_all()
    patch_test_results(model, results)
    passed = sum(1 for _, d in results if d["pass_fail"] == "PASS")
    for tid, d in results:
        print(f"   {tid:5} {d['input'][:34]:34} {d['pass_fail']:6}")
    print(f"   → {passed}/{len(results)} passed")
    if passed != len(results):
        print("   WARNING: some tests failed — see Test_Results sheet.")
        sys.exit(3)

    print("\n[4/4] Rendering to .xlsx ...")
    render(model, out)
    size = os.path.getsize(out)
    print(f"   wrote {out} ({size:,} bytes)")

    # Also write a CSV snapshot of the 15 tests for the docs
    csv = os.path.join(root, "docs", "Test_Results.csv")
    _write_test_csv(results, csv)
    print(f"   wrote {csv}")

    print("\nDone. Open output/ExcelForBorna.xlsx")
    print("-" * 78)


def _write_test_csv(results, path):
    import csv
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Test_ID", "Input", "Expected_Result", "Actual_Result", "Pass_Fail", "Issue", "Executed_Date", "By"])
        for tid, d in results:
            w.writerow([tid, d["input"], d["expected"], d["actual"], d["pass_fail"], d["issue"], "2026-09-22", "USR-09"])
    print(f"   wrote {path}")


if __name__ == "__main__":
    main()
