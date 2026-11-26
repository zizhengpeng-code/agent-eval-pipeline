"""核对教学题的输入边界、参考答案和两条判分路径。"""
import sys
import unittest
from pathlib import Path

from openpyxl import load_workbook

CASE = Path(__file__).resolve().parents[1] / "examples/circular_audit"
sys.path.insert(0, str(CASE))
import run as audit_case


class TestCircularAudit(unittest.TestCase):
    def test_frozen_example(self):
        report = audit_case.run(None)
        self.assertEqual(report["review"], "待复核")
        self.assertEqual(report["visible"], ["reference_files/benefits.xlsx", "task/task.md"])
        errors = report["errors"]
        self.assertEqual(len(errors), 1)
        self.assertIn("#CIRC!", errors[0]["error"])
        ref = audit_case.read_json(CASE / "pack/ground_truth/answer.json")
        self.assertEqual([errors[0]["sheet"] + "!" + errors[0]["coord"]], ref["circular_cells"])
        workbook = load_workbook(CASE / "pack/reference_files/benefits.xlsx", data_only=True)
        sheet = workbook["Expenses"]
        self.assertEqual(sum(sheet[f"B{n}"].value for n in (2, 3, 4)), ref["corrected_total"])
        self.assertEqual(sheet["B5"].value, ref["corrected_total"])
        workbook.close()
        self.assertEqual(report["reference"]["score"], 100)
        self.assertEqual(report["cache_only"]["score"], 40)

    def test_duplicate_addresses_and_non_numeric_amount(self):
        ref = audit_case.read_json(CASE / "pack/ground_truth/answer.json")
        rubric = audit_case.read_json(CASE / "pack/rubric/criteria.json")
        answer = {"circular_cells": ref["circular_cells"] * 2, "corrected_total": "600"}
        self.assertEqual(audit_case.grade(answer, ref, rubric)["score"], 0)
