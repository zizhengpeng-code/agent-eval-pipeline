#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pack_gate 公式重算（数组运算、跨类型比较与财务常用函数）单元测试。"""

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import pack_gate  # noqa: E402


class TestPackGateRecalc(unittest.TestCase):
    """验证 pack_gate.recalc_workbook 的数组计算与常用函数语义。"""

    def _run_book(self, sheets_spec):
        with tempfile.TemporaryDirectory() as td:
            xlsx = Path(td) / "sample.xlsx"
            pack_gate.write_ooxml_xlsx(xlsx, sheets_spec)
            return pack_gate.recalc_workbook(xlsx)

    def test_sumproduct_and_array_comparisons(self):
        """验证 SUMPRODUCT(--(ISNUMBER(rng))*(rng<>0)) 与按列条件计数数组运算（含真空白与文本/公式空串差异）。"""
        res = self._run_book([
            {
                "name": "明细",
                "widths": [12, 12, 12],
                "rows": [
                    ["类别", "原值", "差额"],
                    ["A类", 10, {"f": "IF(B2=0,\"无值\",B2-5)", "v": 5}],
                    ["A类", 0, {"f": "IF(B3=0,\"无值\",B3-5)", "v": "无值"}],
                    ["B类", 20, {"f": "IF(B4=0,\"无值\",B4-5)", "v": 15}],
                    ["A类", -3, {"f": "IF(B5=0,\"无值\",B5-5)", "v": -8}],
                    ["B类", None, {"f": "IF(B6=\"\",\"\",B6)", "v": ""}],
                    ["B类", "无值", {"f": "IF(ISTEXT(B7),\"\",0)", "v": ""}],
                ],
            },
            {
                "name": "汇总",
                "widths": [24, 12, 12],
                "rows": [
                    ["指标", "结果", "校验"],
                    ["非零数值差额数", {"f": "=SUMPRODUCT(--(ISNUMBER(明细!C2:C7))*(明细!C2:C7<>0))", "v": 3}, 1],
                    ["A类且原值大于0", {"f": "=SUMPRODUCT((明细!A2:A7=\"A类\")*(明细!B2:B7>0))", "v": 1}, 1],
                    ["B类且原值>0(含文本>数字,空白不计)", {"f": "=SUMPRODUCT((明细!A2:A7=\"B类\")*(明细!B2:B7>0))", "v": 2}, 1],
                    ["B类且公式空串C列>0(Excel文本>数字)", {"f": "=SUMPRODUCT((明细!A2:A7=\"B类\")*(明细!C2:C7>0))", "v": 3}, 1],
                    ["数组加减求和", {"f": "=SUMPRODUCT(明细!B2:B4+1, --(明细!A2:A4=\"A类\"))", "v": 12}, 1],
                ],
            },
        ])
        self.assertEqual(res["error_count"], 0, res["errors"])
        self.assertEqual(res["cached_mismatch"], 0, res["mismatch_samples"])
        vals = res["values"]["汇总"]
        self.assertEqual(vals["B2"], 3)
        self.assertEqual(vals["B3"], 1)
        self.assertEqual(vals["B4"], 2)
        self.assertEqual(vals["B5"], 3)
        self.assertEqual(vals["B6"], 12)

    def test_if_lazy_and_array_broadcast(self):
        """验证标量 IF 惰性求值不误报除零，且数组 IF 可正常逐元素广播。"""
        res = self._run_book([
            {
                "name": "测试",
                "widths": [12, 12, 12],
                "rows": [
                    ["分母", "惰性除零", "数组分支"],
                    [0, {"f": "=IF(A2=0, 99, 100/A2)", "v": 99}, {"f": "=SUMPRODUCT(IF(A2:A4>0, A2:A4*2, 0))", "v": 14}],
                    [2, {"f": "=IF(A3=0, 99, 100/A3)", "v": 50}, 0],
                    [5, {"f": "=IFERROR(1/A2, -1)", "v": -1}, 0],
                ],
            }
        ])
        self.assertEqual(res["error_count"], 0, res["errors"])
        self.assertEqual(res["cached_mismatch"], 0, res["mismatch_samples"])
        vals = res["values"]["测试"]
        self.assertEqual(vals["B2"], 99)
        self.assertEqual(vals["B3"], 50)
        self.assertEqual(vals["B4"], -1)
        self.assertEqual(vals["C2"], 14)

    def test_financial_and_lookup_functions(self):
        """验证 ROUNDUP/ROUNDDOWN/MOD/SIGN/VLOOKUP/INDEX/MATCH/IFNA/CONCAT/LEFT/RIGHT/MID。"""
        res = self._run_book([
            {
                "name": "维表",
                "widths": [14, 14, 14],
                "rows": [
                    ["编码", "名称", "比例"],
                    ["K01", "设备收入", 0.13],
                    ["K02", "服务代理", 0.08],
                    ["K03", "研发软件", 0.15],
                ],
            },
            {
                "name": "运算",
                "widths": [18, 18],
                "rows": [
                    ["函数", "结果"],
                    ["ROUNDUP整数位", {"f": "=ROUNDUP(3.00, 2)", "v": 3}],
                    ["ROUNDUP进位", {"f": "=ROUNDUP(3.121, 2)", "v": 3.13}],
                    ["ROUNDDOWN截断", {"f": "=ROUNDDOWN(-3.129, 2)", "v": -3.12}],
                    ["ROUND负位数", {"f": "=ROUND(1234, -2)", "v": 1200}],
                    ["MOD与SIGN", {"f": "=MOD(17, 5)*SIGN(-8)", "v": -2}],
                    ["VLOOKUP精确", {"f": "=VLOOKUP(\"K02\", 维表!A2:C4, 3, FALSE)", "v": 0.08}],
                    ["INDEX_MATCH", {"f": "=INDEX(维表!B2:B4, MATCH(\"K03\", 维表!A2:A4, 0))", "v": "研发软件"}],
                    ["IFNA兜底", {"f": "=IFNA(VLOOKUP(\"K99\", 维表!A2:C4, 2, 0), \"未找到\")", "v": "未找到"}],
                    ["ISERROR检测", {"f": "=ISERROR(1/0)", "v": True}],
                    ["TEXT裁剪拼接", {"f": "=CONCAT(LEFT(\"星辰科技\",2), MID(\"2026示例\",3,2), RIGHT(\"D01\",2))", "v": "星辰2601"}],
                ],
            },
        ])
        self.assertEqual(res["error_count"], 0, res["errors"])
        self.assertEqual(res["cached_mismatch"], 0, res["mismatch_samples"])
        v = res["values"]["运算"]
        self.assertEqual(v["B2"], 3)
        self.assertAlmostEqual(v["B3"], 3.13)
        self.assertAlmostEqual(v["B4"], -3.12)
        self.assertEqual(v["B5"], 1200)
        self.assertEqual(v["B6"], -2)
        self.assertAlmostEqual(v["B7"], 0.08)
        self.assertEqual(v["B8"], "研发软件")
        self.assertEqual(v["B9"], "未找到")
        self.assertEqual(v["B10"], True)
        self.assertEqual(v["B11"], "星辰2601")


if __name__ == "__main__":
    unittest.main()
