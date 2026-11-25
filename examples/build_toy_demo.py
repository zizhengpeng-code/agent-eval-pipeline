#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
零依赖端到端合成示例（Toy Demo）：
在 demo_output/ 下生成一套 100% 虚构的微型任务包（星辰商贸 2026Q1 订单对账示例），
并依次演示：
1. 生成多格式 reference_files/（.txt/.csv/.xlsx/.docx）
2. 生成含动态公式的 ground_truth/（.xlsx + .docx）与双表 rubric/（.xlsx + .csv）
3. 执行 Gate 1 & Gate 2 自动化校验（pack_gate.validate_task_pack）
4. 执行离线 AST Excel 公式重算审计（pack_gate.recalc_workbook）
5. 生成 SHA-256 冻结指纹并完成完整性核验（freeze_task_pack + verify_freeze_pack）
"""

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import pack_gate  # noqa: E402


def build_demo(out_root: Path) -> int:
    task_dir = out_root / "任务包_DEMO_星辰商贸订单对账"
    freeze_json = out_root / "冻结指纹.json"
    spec_json = out_root / "rubric_spec.json"

    if out_root.exists():
        shutil.rmtree(out_root)
    for sub in ("task", "reference_files", "ground_truth", "rubric"):
        (task_dir / sub).mkdir(parents=True, exist_ok=True)

    # 1. 题面 task/task.md
    (task_dir / "task" / "task.md").write_text(
        "# 星辰商贸 2026Q1 订单与回款对账任务（合成示例）\n\n"
        "你是星辰商贸有限公司财务运营专员，请根据 `reference_files/` 中的 5 份原始资料，"
        "核对 2026 年第一季度 10 笔重点客户订单的应收金额、实收金额与差异原因，"
        "并将最终对账底稿（`.xlsx`，要求全表动态公式联动）与对账说明（`.docx`）保存至 `output/`。\n",
        encoding="utf-8",
    )

    # 2. 多格式 reference_files/（5 份文件，4 种格式：txt/csv/xlsx/docx）
    ref_dir = task_dir / "reference_files"
    (ref_dir / "01_业务交接邮件.txt").write_text(
        "发件人：运营部-林敏\n收件人：财务部\n主题：2026Q1 重点客户对账说明\n\n"
        "附 Q1 订单导出明细、银行流水与客户往来函件，请核对差异并出具复核底稿。\n",
        encoding="utf-8",
    )
    (ref_dir / "02_系统导出订单.csv").write_text(
        "订单号,客户,不含税金额,税率,含税金额\n"
        "SO-001,甲客户,10000,0.13,11300\n"
        "SO-002,乙客户,20000,0.13,22600\n"
        "SO-003,丙客户,15000,0.13,16950\n"
        "SO-004,丁客户,8000,0.13,9040\n"
        "SO-005,戊客户,12000,0.13,13560\n",
        encoding="utf-8-sig",
    )
    (ref_dir / "03_客户对账函摘要.txt").write_text(
        "客户乙反馈：SO-002 存在 600 元质量折让，已于 3 月底签署补充协议。\n",
        encoding="utf-8",
    )
    pack_gate.write_ooxml_xlsx(
        ref_dir / "04_银行回款流水.xlsx",
        [
            {
                "name": "回款流水",
                "widths": [14, 14, 14, 18],
                "rows": [
                    ["订单号", "客户", "实收金额", "到账日期"],
                    ["SO-001", "甲客户", 11300, "2026-01-20"],
                    ["SO-002", "乙客户", 22000, "2026-02-18"],
                    ["SO-003", "丙客户", 16950, "2026-03-05"],
                    ["SO-004", "丁客户", 9040, "2026-03-19"],
                    ["SO-005", "戊客户", 13000, "2026-03-28"],
                ],
            }
        ],
    )
    pack_gate.write_ooxml_docx(
        ref_dir / "05_补充协议记录.docx",
        title="星辰商贸 2026Q1 客户折让补充协议摘录",
        subtitle="编号：XC-2026-Q1-01",
        sections=[
            {
                "heading": "一、质量折让约定",
                "paragraphs": [
                    "经双方确认，订单 SO-002 给予含税折让 600 元，调整后应收金额为 22,000 元。",
                    "订单 SO-005 尚有 560 元尾款未到账，需业务部门跟进催收。",
                ],
            }
        ],
    )

    # 3. 标准答案 ground_truth/（5 表动态公式 Excel + >=2200 字 2 表 Word）
    orders = [
        ("SO-001", "甲客户", 10000, 0.13, 11300, 0, "02_系统导出订单.csv#L2", "常规全额回款"),
        ("SO-002", "乙客户", 20000, 0.13, 22000, 600, "05_补充协议记录.docx#1", "质量折让调减"),
        ("SO-003", "丙客户", 15000, 0.13, 16950, 0, "02_系统导出订单.csv#L4", "常规全额回款"),
        ("SO-004", "丁客户", 8000, 0.13, 9040, 0, "02_系统导出订单.csv#L5", "常规全额回款"),
        ("SO-005", "戊客户", 12000, 0.13, 13000, 0, "04_银行回款流水.xlsx#R6", "存在未回款尾款"),
        ("SO-006", "己客户", 10000, 0.13, 11300, 0, "02_系统导出订单.csv#示例", "合成扩展行"),
        ("SO-007", "庚客户", 5000, 0.13, 5650, 0, "02_系统导出订单.csv#示例", "合成扩展行"),
        ("SO-008", "辛客户", 9000, 0.13, 10170, 0, "02_系统导出订单.csv#示例", "合成扩展行"),
        ("SO-009", "壬客户", 7000, 0.13, 7910, 0, "02_系统导出订单.csv#示例", "合成扩展行"),
        ("SO-010", "癸客户", 6000, 0.13, 6780, 0, "02_系统导出订单.csv#示例", "合成扩展行"),
    ]

    detail_rows = [["订单号", "客户", "不含税金额", "税率", "合同含税额", "折让调减", "调整后应收", "银行实收", "应收差异", "核对依据", "证据出处"]]
    for idx, (oid, cust, net, rate, paid, disc, ev, basis) in enumerate(orders, start=2):
        gross = round(net * (1 + rate), 2)
        adj = round(gross - disc, 2)
        diff = round(adj - paid, 2)
        detail_rows.append([
            oid,
            cust,
            net,
            rate,
            {"f": f"=ROUND(C{idx}*(1+D{idx}),2)", "v": gross},
            disc,
            {"f": f"=ROUND(E{idx}-F{idx},2)", "v": adj},
            paid,
            {"f": f"=ROUND(G{idx}-H{idx},2)", "v": diff},
            basis,
            ev,
        ])
    last_r = len(orders) + 1
    tot_net = sum(o[2] for o in orders)
    tot_gross = round(sum(o[2] * (1 + o[3]) for o in orders), 2)
    tot_disc = sum(o[5] for o in orders)
    tot_adj = round(tot_gross - tot_disc, 2)
    tot_paid = sum(o[4] for o in orders)
    tot_diff = round(tot_adj - tot_paid, 2)
    detail_rows.append([
        "合计",
        "",
        {"f": f"=SUM(C2:C{last_r})", "v": tot_net},
        "",
        {"f": f"=SUM(E2:E{last_r})", "v": tot_gross},
        {"f": f"=SUM(F2:F{last_r})", "v": tot_disc},
        {"f": f"=SUM(G2:G{last_r})", "v": tot_adj},
        {"f": f"=SUM(H2:H{last_r})", "v": tot_paid},
        {"f": f"=SUM(I2:I{last_r})", "v": tot_diff},
        "汇总校验",
        "全表联动",
    ])

    gt_xlsx = task_dir / "ground_truth" / "星辰商贸_2026Q1_对账复核底稿.xlsx"
    pack_gate.write_ooxml_xlsx(
        gt_xlsx,
        [
            {
                "name": "底稿说明与口径",
                "widths": [18, 48, 22],
                "rows": [
                    ["项目", "口径说明", "依据与证据索引"],
                    ["主体名称", "星辰商贸有限公司（100% 合成演示数据）", "01_业务交接邮件.txt"],
                    ["税率口径", "统一适用 13% 增值税率，保留两位小数", "02_系统导出订单.csv"],
                    ["折让处理", "已签署补充协议的质量折让冲减调整后应收", "05_补充协议记录.docx"],
                ],
            },
            {"name": "订单回款取数明细", "widths": [12, 10, 12, 8, 12, 10, 12, 12, 12, 16, 22], "rows": detail_rows},
            {
                "name": "差异分类复算",
                "widths": [20, 14, 24],
                "rows": [
                    ["差异类型", "金额", "依据与证据出处"],
                    ["质量折让调减", {"f": f"=订单回款取数明细!F{last_r + 1}", "v": tot_disc}, "05_补充协议记录.docx"],
                    ["未到账尾款余额", {"f": f"=订单回款取数明细!I{last_r + 1}", "v": tot_diff}, "04_银行回款流水.xlsx"],
                    ["合计差额", {"f": "=SUM(B2:B3)", "v": round(tot_disc + tot_diff, 2)}, "跨表汇总"],
                ],
            },
            {
                "name": "汇总勾稽校验",
                "widths": [22, 14, 14],
                "rows": [
                    ["勾稽指标", "数值", "校验状态"],
                    ["调整后应收合计", {"f": f"=订单回款取数明细!G{last_r + 1}", "v": tot_adj}, "通过"],
                    ["实收加尾款合计", {"f": f"=订单回款取数明细!H{last_r + 1}+差异分类复算!B3", "v": tot_adj}, "通过"],
                    ["勾稽差额", {"f": "=ROUND(B2-B3,2)", "v": 0}, "通过"],
                ],
            },
            {
                "name": "待补证与跟进清单",
                "widths": [14, 34, 22],
                "rows": [
                    ["订单号", "待跟进事项", "证据索引"],
                    ["SO-005", "跟踪戊客户 560 元尾款回款进度并补充回单", "04_银行回款流水.xlsx#R6"],
                ],
            },
        ],
    )

    para_block = (
        "本合成示例用于演示自动化出题门禁、离线 Excel 公式重算引擎与 SHA-256 防篡改冻结流程。"
        "在实际业务复核场景中，审核人员需要逐笔比对系统导出订单、银行回款流水与客户往来补充协议，"
        "严格区分已取得书面确认的商业折让与尚未收回的逾期尾款，确保财务底稿中的每一项调整都具备清晰的证据索引与动态公式勾稽关系。"
        "通过对不含税金额、增值税税率、合同含税总额、协议折让金额及银行实收金额建立逐行联动公式，"
        "可以在不依赖外部电子表格软件的前提下，利用离线抽象语法树求值器完成全工作簿重算、循环引用检测与缓存一致性核验。"
    )
    gt_docx = task_dir / "ground_truth" / "星辰商贸_2026Q1_对账复核意见书.docx"
    pack_gate.write_ooxml_docx(
        gt_docx,
        title="星辰商贸有限公司 2026 年第一季度订单与回款对账复核意见书",
        subtitle="合成演示标准答案（Toy Demo）",
        sections=[
            {
                "heading": "一、复核背景与执行摘要",
                "paragraphs": [para_block] * 4,
                "table": {
                    "headers": ["核心指标", "合同含税额", "折让调减", "调整后应收", "银行实收", "期末尾款"],
                    "rows": [["2026Q1 汇总", f"{tot_gross:.2f}", f"{tot_disc:.2f}", f"{tot_adj:.2f}", f"{tot_paid:.2f}", f"{tot_diff:.2f}"]],
                },
            },
            {
                "heading": "二、重点差异事项与证据溯源",
                "paragraphs": [para_block] * 4,
                "table": {
                    "headers": ["订单号", "客户", "差异类型", "金额", "证据索引"],
                    "rows": [
                        ["SO-002", "乙客户", "质量折让调减", "600.00", "05_补充协议记录.docx"],
                        ["SO-005", "戊客户", "未回款尾款", "560.00", "04_银行回款流水.xlsx"],
                    ],
                },
            },
            {
                "heading": "三、后续整改与内控建议",
                "paragraphs": [para_block] * 4,
            },
        ],
    )

    # 4. 生成双表 rubric.xlsx 与 rubric.csv（15 条得分项合计 100 分 + 3 条扣分项）
    dims = sorted(pack_gate.OFFICIAL_DIMENSIONS)
    rubric_items = [
        {"id": f"R{i:02d}", "dimension": dims[(i - 1) % 8], "check": f"示例核心检查项 {i:02d}", "standard": "客观二元判定达标", "points": pts, "type": "得分", "location": "对账复核底稿.xlsx"}
        for i, pts in enumerate([7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 6, 6, 6, 6, 6], start=1)
    ] + [
        {"id": "D1", "dimension": dims[0], "check": "捏造材料不存在的依据或金额", "standard": "出现即扣分", "points": -5, "type": "扣分", "location": "全交付物"},
        {"id": "D2", "dimension": dims[8], "check": "Excel 存在循环引用或公式报错", "standard": "AST 重算发现错误即扣分", "points": -5, "type": "扣分", "location": "对账复核底稿.xlsx"},
        {"id": "D3", "dimension": dims[2], "check": "缺失关键交付文件", "standard": "未交付核心文件即扣分", "points": -5, "type": "扣分", "location": "全交付物"},
    ]
    spec_data = {
        "task_name": "星辰商贸 2026Q1 订单对账示例",
        "items": rubric_items,
        "instructions": {
            "judgment_principle": "二元客观判定，达标得全部分值，未达标得 0 分。",
            "accepted_alternatives": "金额四舍五入允许 ±0.01 元尾差；等价公式写法均予认可。",
            "non_cascading_rule": "上游失分时，下游汇总项按答题人自身上游数据复核，不重复扣分。",
            "deduction_rule": "D1–D3 独立判定，触发即扣对应负分，总分最低扣至 0 分。",
        },
    }
    spec_json.write_text(json.dumps(spec_data, ensure_ascii=False, indent=2), encoding="utf-8")
    pack_gate.build_rubric_from_spec(spec_json, task_dir / "rubric")

    # 5. 执行 Gate 1 & 2 验收
    errs = pack_gate.validate_task_pack(task_dir)
    if errs:
        print("❌ Toy Demo Gate 1/2 未通过：")
        for e in errs:
            print("  -", e)
        return 1
    print(f"✅ Gate 1 & 2 验收通过：{task_dir}")

    # 6. 执行离线 AST Excel 公式重算
    rec = pack_gate.recalc_workbook(gt_xlsx)
    print(
        f"✅ 离线 AST 重算完成：formulas={rec['formula_total']}, "
        f"errors={rec['error_count']}, cached_mismatch={rec['cached_mismatch']}"
    )
    if rec["error_count"] != 0 or rec["cached_mismatch"] != 0:
        return 1

    # 7. 执行冻结与防篡改核验
    if pack_gate.freeze_task_pack(task_dir, freeze_json) != 0:
        return 1
    if pack_gate.verify_freeze_pack(task_dir, freeze_json) != 0:
        return 1

    print("\n🎉 全链路合成示例运行成功！产物目录：", out_root)
    return 0


if __name__ == "__main__":
    sys.exit(build_demo(ROOT / "demo_output"))
