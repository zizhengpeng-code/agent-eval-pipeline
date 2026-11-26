"""复现合成审题记录、盲测投放和逐项判分。"""
import argparse
import csv
import json
import sys
import tempfile
from pathlib import Path

CASE = Path(__file__).resolve().parent
sys.path.insert(0, str(CASE.parents[1] / "scripts"))
import pack_gate
from run_blind_eval import _prepare_stage_dir


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def review_record(path):
    with path.open(newline="", encoding="utf-8-sig") as stream:
        rows = list(csv.DictReader(stream))
    for revision, row in enumerate(rows, 1):
        if row["版本"] != str(revision) or not row["原因"].strip():
            raise ValueError("审核记录缺少连续版本或原因")
        if row["结论"] == "要改" and not row["改成什么"].strip():
            raise ValueError("退回记录缺少修改要求")
        if row["结论"] not in ("好", "要改", "不要", "待复核"):
            raise ValueError("未知审核结论")
    if not rows or rows[-1]["记录类型"] != "专家复核" or rows[-1]["结论"] in ("要改", "不要"):
        raise ValueError("当前版本缺少可演示的复核记录")
    return rows[-1]["结论"]


def grade(answer, reference, rubric):
    if not isinstance(answer, dict):
        raise ValueError("提交必须是 JSON 对象")
    cells = answer.get("circular_cells")
    expected = reference["circular_cells"]
    cells_match = (isinstance(cells, list) and all(isinstance(c, str) for c in cells)
                   and sorted(cells) == sorted(expected))
    total = answer.get("corrected_total")
    amount_matches = type(total) in (int, float) and total == reference["corrected_total"]
    checks = {}
    for key, match in (("circular_cells", cells_match), ("corrected_total", amount_matches)):
        checks[key] = {"submitted": answer.get(key), "expected": reference[key],
                       "earned": rubric[key] if match else 0, "possible": rubric[key]}
    return {"score": sum(check["earned"] for check in checks.values()), "criteria": checks}


def run(submission):
    pack = CASE / "pack"
    freeze = CASE / "freeze.json"
    if pack_gate.verify_freeze_pack(pack, freeze):
        raise ValueError("任务包与示例指纹不一致")
    status = review_record(pack / "review.csv")
    with tempfile.TemporaryDirectory() as scratch:
        stage = Path(scratch) / "workspace"
        manifest = _prepare_stage_dir(pack, read_json(freeze), stage)
        visible = sorted(p.relative_to(stage).as_posix() for p in stage.rglob("*") if p.is_file())
        if visible != ["reference_files/benefits.xlsx", "task/task.md"] or not all(
                item["matches_freeze"] for item in manifest["files"].values()):
            raise ValueError("盲测输入清单或指纹异常")
        audit = pack_gate.recalc_workbook(stage / visible[0])
    reference = read_json(pack / "ground_truth/answer.json")
    rubric = read_json(pack / "rubric/criteria.json")
    results = {"review": status, "visible": visible, "errors": audit["errors"]}
    if submission:
        results["submission"] = grade(read_json(submission), reference, rubric)
    else:
        results["reference"] = grade(reference, reference, rubric)
        results["cache_only"] = grade(read_json(CASE / "submissions/cache_only.json"), reference, rubric)
    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--submission", type=Path, help="模型或人工提交的 answer.json")
    args = parser.parse_args()
    try:
        print(json.dumps(run(args.submission), ensure_ascii=False, indent=2))
    except (ValueError, OSError) as exc:
        parser.exit(2, f"{exc}\n")
