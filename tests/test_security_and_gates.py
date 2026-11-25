#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""安全边界与质量门禁负向测试：验证循环引用、缓存造假、冻结篡改、剧透拦截与路径逃逸防护。"""

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import openai_compat_agent  # noqa: E402
import pack_gate  # noqa: E402
import run_blind_eval  # noqa: E402


class TestSecurityAndGates(unittest.TestCase):
    """验证关键门禁与隔离机制在异常/对抗输入下能够可靠拦截。"""

    def test_detects_circular_reference(self):
        """验证离线 AST 引擎能够检出直接/间接循环引用（#CIRC!）。"""
        with tempfile.TemporaryDirectory() as td:
            xlsx = Path(td) / "circ.xlsx"
            pack_gate.write_ooxml_xlsx(
                xlsx,
                [
                    {
                        "name": "循环表",
                        "widths": [12, 12],
                        "rows": [
                            ["甲", "乙"],
                            [{"f": "=B2+1", "v": 10}, {"f": "=A2+1", "v": 9}],
                        ],
                    }
                ],
            )
            res = pack_gate.recalc_workbook(xlsx)
            self.assertGreaterEqual(res["error_count"], 2)
            self.assertTrue(any("#CIRC!" in str(e.get("error", "")) for e in res["errors"]), res["errors"])

    def test_detects_cached_value_mismatch(self):
        """验证模型写入虚假缓存值（公式与缓存不一致）时能被精准检出。"""
        with tempfile.TemporaryDirectory() as td:
            xlsx = Path(td) / "mismatch.xlsx"
            pack_gate.write_ooxml_xlsx(
                xlsx,
                [
                    {
                        "name": "核验表",
                        "widths": [12, 12, 12],
                        "rows": [
                            ["A", "B", "合计"],
                            [100, 250, {"f": "=A2+B2", "v": 999}],
                        ],
                    }
                ],
            )
            res = pack_gate.recalc_workbook(xlsx)
            self.assertEqual(res["error_count"], 0)
            self.assertEqual(res["cached_mismatch"], 1)
            self.assertEqual(res["values"]["核验表"]["C2"], 350)

    def test_freeze_detects_single_byte_tampering(self):
        """验证任务包冻结后，哪怕修改 1 个字符或新增文件，verify_freeze_pack 必返回非零。"""
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            task_dir = root / "task_pack"
            (task_dir / "task").mkdir(parents=True)
            (task_dir / "reference_files").mkdir(parents=True)
            (task_dir / "task" / "task.md").write_text("原始题面内容", encoding="utf-8")
            (task_dir / "reference_files" / "a.txt").write_text("材料A", encoding="utf-8")

            freeze_json = root / "freeze.json"
            # 直接写入合法冻结清单（模拟冻结后的状态）
            files_map = {
                "task/task.md": hashlib.sha256((task_dir / "task" / "task.md").read_bytes()).hexdigest(),
                "reference_files/a.txt": hashlib.sha256((task_dir / "reference_files" / "a.txt").read_bytes()).hexdigest(),
            }
            freeze_json.write_text(json.dumps({"files": files_map}, ensure_ascii=False), encoding="utf-8")

            # 未篡改时通过
            self.assertEqual(pack_gate.verify_freeze_pack(task_dir, freeze_json), 0)

            # 修改 1 个字符后必须失败
            (task_dir / "task" / "task.md").write_text("原始题面内容!", encoding="utf-8")
            self.assertNotEqual(pack_gate.verify_freeze_pack(task_dir, freeze_json), 0)

    def test_gate1_blocks_spoiler_filename_and_header(self):
        """验证 Gate 1 能拦截剧透文件名和预标答案辅助列。"""
        with tempfile.TemporaryDirectory() as td:
            task_dir = Path(td) / "bad_task"
            for sub in ("task", "reference_files", "ground_truth", "rubric"):
                (task_dir / sub).mkdir(parents=True)
            (task_dir / "task" / "task.md").write_text("题面", encoding="utf-8")
            ref = task_dir / "reference_files"
            (ref / "01_复核口径说明.txt").write_text("剧透文件", encoding="utf-8")
            (ref / "02_明细.csv").write_text("订单号,金额,复核结论\n1,100,错账\n", encoding="utf-8")
            (ref / "03_函件.md").write_text("文本", encoding="utf-8")
            (ref / "04_补充.txt").write_text("文本", encoding="utf-8")
            (ref / "05_记录.txt").write_text("文本", encoding="utf-8")

            errs = pack_gate.validate_task_pack(task_dir)
            joined = "\n".join(errs)
            self.assertIn("违规剧透文件", joined)
            self.assertIn("违规辅助列", joined)

    def test_sandbox_blocks_path_traversal(self):
        """验证隔离 Agent 工具严格拦截 ../ 路径逃逸与外部绝对路径。"""
        with tempfile.TemporaryDirectory() as td:
            workdir = Path(td) / "sandbox"
            workdir.mkdir()
            (workdir / "ok.txt").write_text("safe", encoding="utf-8")

            # 合法相对路径放行
            resolved = openai_compat_agent._resolve_safe(workdir, "ok.txt")
            self.assertEqual(resolved, (workdir / "ok.txt").resolve())

            # 越界相对路径与绝对路径必须抛出 ValueError
            with self.assertRaises(ValueError):
                openai_compat_agent._resolve_safe(workdir, "../secret.txt")
            with self.assertRaises(ValueError):
                openai_compat_agent._resolve_safe(workdir, "/etc/passwd")

    def test_event_scanner_catches_leak_and_network_keywords(self):
        """验证盲测事后审计能检出越界路径与联网命令关键词。"""
        with tempfile.TemporaryDirectory() as td:
            ev_file = Path(td) / "events.jsonl"
            lines = [
                json.dumps({"type": "item.completed", "item": {"command": "ls reference_files"}}),
                json.dumps({"type": "item.completed", "item": {"command": "cat ../ground_truth/ans.xlsx"}}),
                json.dumps({"type": "tool_result", "arguments": {"command": "curl https://example.com"}}),
            ]
            ev_file.write_text("\n".join(lines) + "\n", encoding="utf-8")

            hits = run_blind_eval._scan_events_for_violations(ev_file)
            keywords = {h["keyword"] for h in hits}
            self.assertIn("ground_truth", keywords)
            self.assertIn("curl", keywords)


if __name__ == "__main__":
    unittest.main()
