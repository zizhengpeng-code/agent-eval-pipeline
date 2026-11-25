#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务包与体验评测包通用生成与四道闸门（Gate 1~4）校验工具。
兼容 Python 3.9+ 标准库（使用 zipfile 与 xml.etree.ElementTree 解析与生成 OOXML）。

用法：
1. 校验任务包（Gate 1 材料防剧透与多格式 + Gate 2 标准答案与 Rubric 规格）：
   python3 scripts/pack_gate.py validate-task <任务包目录>
2. 校验体验评测包（Gate 3 & Gate 4 评测记录 5 表 + 评测报告 6 节）：
   python3 scripts/pack_gate.py validate-eval <体验评测目录> --rubric <rubric.xlsx路径>
3. 从单题配置 JSON 自动生成双工作表 rubric.xlsx 与 rubric.csv：
   python3 scripts/pack_gate.py build-rubric <rubric_spec.json> <rubric输出目录>
4. 从评测配置 JSON 自动生成体验评测记录.xlsx、体验评测报告.docx 与 操作轨迹摘要.md：
   python3 scripts/pack_gate.py build-eval <eval_spec.json> <体验评测输出目录>
5. 重算并校验 Excel 公式与缓存一致性（不依赖 LibreOffice）：
   python3 scripts/pack_gate.py recalc-xlsx <xlsx路径> [--json-out <结果json>]
6. 冻结任务包指纹与核验任务包未篡改（先冻结后盲测）：
   python3 scripts/pack_gate.py freeze-task <任务包目录> --out <冻结指纹.json>
   python3 scripts/pack_gate.py verify-freeze <任务包目录> --freeze <冻结指纹.json>
"""

import argparse
import csv
from datetime import datetime, timedelta, timezone
import hashlib
import json
import math
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

# 标准 9 大评价维度
OFFICIAL_DIMENSIONS = {
    "1 关键事实准确性",
    "2 关键指标与计算准确性",
    "3 方法与业务逻辑正确性",
    "4 需求覆盖与完成度",
    "5 证据依据与可追溯性",
    "6 数据与内容的一致性",
    "7 专业判断与建议质量",
    "8 交付物可用性与适用条件",
    "9 交付规范与表达质量",
}

RUBRIC_HEADERS = ["编号", "评价维度", "检查项", "判定标准", "分值", "类型", "核对出处"]
EVAL_SCORING_HEADERS = [
    "编号",
    "评价维度",
    "检查项",
    "判定标准",
    "分值",
    "类型",
    "核对出处",
    "判定",
    "得分",
    "判定理由（被测交付物原文位置／参考材料位置）",
    "备注",
]
EVAL_SHEETS_EXPECTED = ["基本信息", "关键操作轨迹", "逐项评分", "问题记录", "定性反馈"]
EXCEL_ERRORS = {"#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}

# 禁止在 reference_files 文件名中出现的剧透词
FORBIDDEN_REF_FILENAME_WORDS = ("复核口径", "解题口径", "答题口径", "评分口径", "解题指南", "参考答案")
# 禁止在 reference_files 表格表头中人为添加的剧透辅助列
FORBIDDEN_REF_HEADER_WORDS = ("选修组", "计划计入", "是否计入", "计入系数", "复核结论", "建议更正")


def count_chinese_chars(text: str) -> int:
    """统计字符串中的中文字符数量。"""
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def col_letter(n: int) -> str:
    """将 1-based 列号转为 Excel 列字母。"""
    res = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        res = chr(65 + r) + res
    return res


def read_xlsx_book(path: Path) -> List[Tuple[str, List[List[Dict[str, Any]]]]]:
    """
    使用标准库解析 .xlsx，返回 [(sheet_name, rows), ...]。
    每个单元格字典包含: {'coord': 'A1', 'value': str, 'formula': Optional[str], 'type': Optional[str]}
    """
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path, "r") as z:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_xml}
        ss: List[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_xml.findall(".//main:si", ns):
                ss.append("".join(t.text or "" for t in si.findall(".//main:t", ns)))

        result = []
        rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        for s in wb_xml.findall(".//main:sheet", ns):
            sname = s.attrib["name"]
            rid = s.attrib[rel_ns]
            target = rels[rid].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            ws_xml = ET.fromstring(z.read(target))
            sheet_rows = []
            for r in ws_xml.findall(".//main:row", ns):
                row_cells = []
                for c in r.findall("main:c", ns):
                    coord = c.attrib.get("r", "")
                    t = c.attrib.get("t")
                    f_el = c.find("main:f", ns)
                    v_el = c.find("main:v", ns)
                    is_el = c.find("main:is", ns)
                    val = ""
                    if v_el is not None and v_el.text is not None:
                        if t == "s":
                            idx = int(v_el.text)
                            val = ss[idx] if 0 <= idx < len(ss) else v_el.text
                        else:
                            val = v_el.text
                    elif is_el is not None:
                        val = "".join(x.text or "" for x in is_el.findall(".//main:t", ns))
                    formula = f_el.text if (f_el is not None and f_el.text is not None) else None
                    row_cells.append({
                        "coord": coord,
                        "value": val.strip() if isinstance(val, str) else str(val),
                        "formula": formula,
                        "type": t,
                    })
                sheet_rows.append(row_cells)
            result.append((sname, sheet_rows))
        return result


def read_docx_info(path: Path) -> Dict[str, Any]:
    """使用标准库读取 .docx 的段落文本、正文汉字数（不含表格）和表格列表。"""
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(path, "r") as z:
        xml = ET.fromstring(z.read("word/document.xml"))
        body = xml.find("w:body", ns)
        paras = []
        if body is not None:
            for p in body.findall("w:p", ns):
                txt = "".join(t.text or "" for t in p.findall(".//w:t", ns)).strip()
                if txt:
                    paras.append(txt)
        tables = []
        if body is not None:
            for tbl in body.findall("w:tbl", ns):
                t_rows = []
                for tr in tbl.findall("w:tr", ns):
                    t_cells = [
                        "".join(t.text or "" for t in tc.findall(".//w:t", ns)).strip()
                        for tc in tr.findall("w:tc", ns)
                    ]
                    t_rows.append(t_cells)
                tables.append(t_rows)
        body_text = "\n".join(paras)
        return {
            "paragraphs": paras,
            "tables": tables,
            "body_chinese_chars": count_chinese_chars(body_text),
            "full_text": body_text + "\n" + "\n".join(" ".join(r) for t in tables for r in t),
        }


# ==============================================================================
# 纯标准库生成 OOXML (.xlsx / .docx) 辅助函数（兼容无第三方库的 Python 3.9）
# ==============================================================================

def write_ooxml_xlsx(path: Path, sheets_spec: List[Dict[str, Any]]) -> None:
    """
    生成标准可打开、带表头样式与自动换行的 .xlsx 文件（支持公式与缓存值同时写入）。
    sheets_spec 每项结构：
    {
      "name": "工作表名",
      "widths": [12, 30, ...],
      "rows": [ [val_or_dict, ...], ... ]
    }
    其中单元格如果是 dict，可为 {"f": "SUM(B2:B9)", "v": 100}；否则按普通值写入。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml
        sheet_overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, len(sheets_spec) + 1)
        )
        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            f"{sheet_overrides}</Types>"
        )
        z.writestr("[Content_Types].xml", content_types)

        # _rels/.rels
        rels_root = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/></Relationships>'
        )
        z.writestr("_rels/.rels", rels_root)

        # xl/workbook.xml & xl/_rels/workbook.xml.rels
        sheets_xml = "".join(
            f'<sheet name="{xml_escape(s["name"])}" sheetId="{i}" r:id="rId{i}"/>'
            for i, s in enumerate(sheets_spec, 1)
        )
        workbook_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheets_xml}</sheets></workbook>"
        )
        z.writestr("xl/workbook.xml", workbook_xml)

        wb_rels = "".join(
            f'<Relationship Id="rId{i}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, len(sheets_spec) + 1)
        )
        wb_rels += (
            f'<Relationship Id="rId{len(sheets_spec)+1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{wb_rels}</Relationships>",
        )

        # xl/styles.xml (style 0: 默认, style 1: 表头深蓝底白字加粗, style 2: 正文自动换行带细边框)
        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="2">'
            '<font><sz val="10"/><name val="Arial Unicode MS"/></font>'
            '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial Unicode MS"/></font>'
            '</fonts>'
            '<fills count="3">'
            '<fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/></patternFill></fill>'
            '</fills>'
            '<borders count="2">'
            '<border><left/><right/><top/><bottom/><diagonal/></border>'
            '<border>'
            '<left style="thin"><color rgb="FFBFBFBF"/></left>'
            '<right style="thin"><color rgb="FFBFBFBF"/></right>'
            '<top style="thin"><color rgb="FFBFBFBF"/></top>'
            '<bottom style="thin"><color rgb="FFBFBFBF"/></bottom>'
            '<diagonal/>'
            '</border>'
            '</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="3">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
            '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
            '<alignment vertical="top" wrapText="1"/></xf>'
            '</cellXfs>'
            '</styleSheet>'
        )
        z.writestr("xl/styles.xml", styles_xml)

        # 逐张写入工作表
        for idx, s in enumerate(sheets_spec, 1):
            widths = s.get("widths", [])
            cols_xml = ""
            if widths:
                col_items = "".join(
                    f'<col min="{c_i}" max="{c_i}" width="{w}" customWidth="1"/>'
                    for c_i, w in enumerate(widths, 1)
                )
                cols_xml = f"<cols>{col_items}</cols>"

            rows_xml_list = []
            for r_i, row in enumerate(s.get("rows", []), 1):
                style_id = "1" if r_i == 1 else "2"
                c_xml_list = []
                for c_i, cell in enumerate(row, 1):
                    coord = f"{col_letter(c_i)}{r_i}"
                    if isinstance(cell, dict) and "f" in cell:
                        formula_txt = xml_escape(str(cell["f"]).lstrip("="))
                        cached_v = cell.get("v")
                        if cached_v is None:
                            c_xml_list.append(f'<c r="{coord}" s="{style_id}"><f>{formula_txt}</f></c>')
                        elif isinstance(cached_v, (int, float)) and not isinstance(cached_v, bool):
                            c_xml_list.append(
                                f'<c r="{coord}" s="{style_id}"><f>{formula_txt}</f><v>{cached_v}</v></c>'
                            )
                        else:
                            c_xml_list.append(
                                f'<c r="{coord}" s="{style_id}" t="str"><f>{formula_txt}</f>'
                                f"<v>{xml_escape(str(cached_v))}</v></c>"
                            )
                    elif cell is None or cell == "":
                        c_xml_list.append(f'<c r="{coord}" s="{style_id}"/>')
                    elif isinstance(cell, (int, float)) and not isinstance(cell, bool):
                        c_xml_list.append(f'<c r="{coord}" s="{style_id}"><v>{cell}</v></c>')
                    else:
                        text_escaped = xml_escape(str(cell))
                        c_xml_list.append(
                            f'<c r="{coord}" s="{style_id}" t="inlineStr"><is><t>{text_escaped}</t></is></c>'
                        )
                rows_xml_list.append(f'<row r="{r_i}">{"".join(c_xml_list)}</row>')

            ws_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                f'{cols_xml}<sheetData>{"".join(rows_xml_list)}</sheetData></worksheet>'
            )
            z.writestr(f"xl/worksheets/sheet{idx}.xml", ws_xml)


def write_ooxml_docx(path: Path, title: str, subtitle: str, sections: List[Dict[str, Any]]) -> None:
    """
    生成标准可打开的 .docx 文件（支持标题、多级小节、段落与表格）。
    sections 每项结构：
    {
      "heading": "一、...",
      "paragraphs": ["段落1", ...],
      "table": [ ["列1", "列2"], ["值1", "值2"], ... ] (可选),
      "after_paragraphs": ["表后段落1", ...] (可选)
    }
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    body_parts = []

    def make_p(text: str, bold: bool = False, size_half_pt: int = 22, space_after: int = 120) -> str:
        b_xml = "<w:b/>" if bold else ""
        rpr = (
            f"<w:rPr><w:rFonts w:ascii=\"Arial Unicode MS\" w:eastAsia=\"Arial Unicode MS\" "
            f"w:hAnsi=\"Arial Unicode MS\"/>{b_xml}<w:sz w:val=\"{size_half_pt}\"/>"
            f"<w:szCs w:val=\"{size_half_pt}\"/></w:rPr>"
        )
        ppr = f"<w:pPr><w:spacing w:after=\"{space_after}\" w:line=\"300\" w:lineRule=\"auto\"/></w:pPr>"
        return f"<w:p>{ppr}<w:r>{rpr}<w:t xml:space=\"preserve\">{xml_escape(text)}</w:t></w:r></w:p>"

    def make_tbl(rows: List[List[Any]]) -> str:
        tbl_pr = (
            "<w:tblPr><w:tblW w:w=\"5000\" w:type=\"pct\"/>"
            "<w:tblBorders>"
            "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"BFBFBF\"/>"
            "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"BFBFBF\"/>"
            "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"BFBFBF\"/>"
            "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"BFBFBF\"/>"
            "<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"BFBFBF\"/>"
            "<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"BFBFBF\"/>"
            "</w:tblBorders></w:tblPr>"
        )
        tr_list = []
        for r_idx, r in enumerate(rows):
            tc_list = []
            for c in r:
                p_xml = make_p(str(c), bold=(r_idx == 0), size_half_pt=20, space_after=40)
                tc_list.append(f"<w:tc><w:tcPr/>{p_xml}</w:tc>")
            tr_list.append(f"<w:tr>{''.join(tc_list)}</w:tr>")
        return f"<w:tbl>{tbl_pr}{''.join(tr_list)}</w:tbl>"

    if title:
        body_parts.append(make_p(title, bold=True, size_half_pt=36, space_after=160))
    if subtitle:
        body_parts.append(make_p(subtitle, bold=False, size_half_pt=21, space_after=200))

    for sec in sections:
        heading = sec.get("heading")
        if heading:
            body_parts.append(make_p(heading, bold=True, size_half_pt=26, space_after=120))
        for p_txt in sec.get("paragraphs", []):
            body_parts.append(make_p(p_txt, bold=False, size_half_pt=22, space_after=120))
        if sec.get("table"):
            body_parts.append(make_tbl(sec["table"]))
        for p_txt in sec.get("after_paragraphs", []):
            body_parts.append(make_p(p_txt, bold=False, size_half_pt=22, space_after=120))

    sect_pr = (
        "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
        "<w:pgMar w:top=\"1134\" w:right=\"1247\" w:bottom=\"1134\" w:left=\"1247\"/></w:sectPr>"
    )
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{''.join(body_parts)}{sect_pr}</w:body></w:document>"
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>",
        )
        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>',
        )
        z.writestr("word/document.xml", doc_xml)


# ==============================================================================
# Gate 1 & Gate 2：任务包校验器
# ==============================================================================

def validate_task_pack(task_dir: Path) -> List[str]:
    """校验任务包目录是否符合 Gate 1（输入材料防剧透与多格式）和 Gate 2（GT 与 Rubric 规格）。"""
    errors: List[str] = []

    # 1. 基础目录存在性
    for sub in ("task", "reference_files", "ground_truth", "rubric"):
        if not (task_dir / sub).is_dir():
            errors.append(f"[目录缺失] 缺少子目录: {sub}/")
    if errors:
        return errors

    task_md = task_dir / "task" / "task.md"
    if not task_md.is_file():
        errors.append("[Gate 1] 缺少 task/task.md")

    # 2. Gate 1: reference_files 校验
    ref_files = sorted(
        p for p in (task_dir / "reference_files").iterdir()
        if p.is_file() and not p.name.startswith(".")
    )
    if len(ref_files) < 5:
        errors.append(
            f"[Gate 1 材料数量不足] reference_files/ 仅有 {len(ref_files)} 份文件（要求至少 5~12 份多源材料，禁止单表拆片）"
        )
    exts = {p.suffix.lower() for p in ref_files}
    if len(exts) < 3:
        errors.append(
            f"[Gate 1 格式单一(D2不足)] reference_files/ 仅包含 {sorted(exts)} 共 {len(exts)} 种格式（要求至少 3 种格式，如 txt/pdf/xlsx/docx/csv 组合）"
        )

    for p in ref_files:
        for bad_word in FORBIDDEN_REF_FILENAME_WORDS:
            if bad_word in p.name:
                errors.append(
                    f"[Gate 1 违规剧透文件] reference_files/{p.name} 文件名含「{bad_word}」，严禁将出题人总结的复核口径说明直接发给考生！"
                )
        if p.suffix.lower() == ".csv":
            first_line = p.read_text(encoding="utf-8-sig", errors="ignore").splitlines()[:1]
            if first_line:
                for bad_col in FORBIDDEN_REF_HEADER_WORDS:
                    if bad_col in first_line[0]:
                        errors.append(
                            f"[Gate 1 违规辅助列] reference_files/{p.name} 表头含人为标注列「{bad_col}」，严禁替答题模型预标答案线索！"
                        )
        elif p.suffix.lower() == ".xlsx":
            try:
                sheets = read_xlsx_book(p)
                for sname, srows in sheets:
                    if srows:
                        headers_str = " | ".join(c["value"] for c in srows[0])
                        for bad_col in FORBIDDEN_REF_HEADER_WORDS:
                            if bad_col in headers_str:
                                errors.append(
                                    f"[Gate 1 违规辅助列] reference_files/{p.name} [{sname}] 表头含人为标注列「{bad_col}」！"
                                )
            except Exception as exc:
                errors.append(f"[Gate 1 文件损坏] reference_files/{p.name} 无法作为 xlsx 读取: {exc}")

    # 3. Gate 2: rubric/rubric.xlsx 与 rubric.csv 校验
    rubric_xlsx = task_dir / "rubric" / "rubric.xlsx"
    rubric_csv = task_dir / "rubric" / "rubric.csv"
    if not rubric_xlsx.is_file():
        errors.append("[Gate 2] 缺少 rubric/rubric.xlsx")
    else:
        try:
            r_sheets = read_xlsx_book(rubric_xlsx)
            r_sheet_names = [s[0] for s in r_sheets]
            if r_sheet_names != ["评分表", "评分说明"]:
                errors.append(
                    f"[Gate 2 Rubric 工作表不合规] rubric.xlsx 工作表为 {r_sheet_names}，必须严格为 ['评分表', '评分说明'] 双表结构"
                )
            if r_sheets:
                _, s1_rows = r_sheets[0]
                if not s1_rows:
                    errors.append("[Gate 2] rubric.xlsx『评分表』为空")
                else:
                    hdr = [c["value"] for c in s1_rows[0]]
                    if hdr != RUBRIC_HEADERS:
                        errors.append(
                            f"[Gate 2 Rubric 列结构错误]『评分表』表头为 {hdr}，必须严格为 7 列: {RUBRIC_HEADERS}"
                        )
                    else:
                        pos_count = 0
                        neg_count = 0
                        pos_sum = 0
                        max_single = 0
                        dim9_points = 0
                        for row in s1_rows[1:]:
                            vals = [c["value"] for c in row]
                            if not any(vals) or any(k in vals[0] for k in ("合计", "总计", "总分")):
                                continue
                            if len(vals) < 7:
                                errors.append(f"[Gate 2] 评分项 {vals} 列数不足 7 列")
                                continue
                            code, dim, check, std, pts_str, rtype, loc = vals[:7]
                            if dim not in OFFICIAL_DIMENSIONS:
                                errors.append(
                                    f"[Gate 2 维度不合规] 条目 {code} 评价维度「{dim}」不在标准 9 大评价维度内"
                                )
                            try:
                                pts = int(float(pts_str))
                            except ValueError:
                                errors.append(f"[Gate 2] 条目 {code} 分值「{pts_str}」不是整数")
                                continue
                            if rtype == "得分":
                                pos_count += 1
                                pos_sum += pts
                                max_single = max(max_single, pts)
                                if dim.startswith("9 "):
                                    dim9_points += pts
                            elif rtype == "扣分":
                                neg_count += 1
                                if pts >= 0:
                                    errors.append(f"[Gate 2] 扣分项 {code} 分值必须为负数，当前为 {pts}")
                            else:
                                errors.append(f"[Gate 2] 条目 {code} 类型「{rtype}」必须为『得分』或『扣分』")

                        if not (12 <= pos_count <= 30):
                            errors.append(
                                f"[Gate 2 颗粒度不合规] 得分项数量为 {pos_count} 条（要求 12~30 条，避免过粗导致一错扣十几分）"
                            )
                        if pos_sum != 100:
                            errors.append(f"[Gate 2] 得分项分值合计为 {pos_sum}，必须等于 100")
                        if max_single > 8:
                            errors.append(
                                f"[Gate 2 单条权重过高] 单条最高分为 {max_single} 分（标准规范要求单条不超过 6~8 分）"
                            )
                        if neg_count < 3:
                            errors.append(
                                f"[Gate 2 缺少负分扣分项] 当前扣分项仅 {neg_count} 条（要求至少 3 条 D1~D3 扣分项）"
                            )
                        if dim9_points > 4:
                            errors.append(
                                f"[Gate 2 规范分过高] 维度 9（交付规范与表达质量）分值为 {dim9_points} 分（要求 ≤ 4 分，核心内容不可被格式项稀释）"
                            )
            if len(r_sheets) >= 2:
                _, s2_rows = r_sheets[1]
                if len(s2_rows) < 7:
                    errors.append(
                        f"[Gate 2] rubric.xlsx『评分说明』仅有 {len(s2_rows)} 行（要求包含表头 + 7 条任务级规则）"
                    )
        except Exception as exc:
            errors.append(f"[Gate 2] 读取 rubric.xlsx 失败: {exc}")

    if not rubric_csv.is_file():
        errors.append("[Gate 2] 缺少 rubric/rubric.csv")

    # 4. Gate 2: ground_truth 校验
    gt_xlsx_files = sorted((task_dir / "ground_truth").glob("*.xlsx"))
    gt_docx_files = sorted((task_dir / "ground_truth").glob("*.docx"))
    if not gt_xlsx_files:
        errors.append("[Gate 2] ground_truth/ 缺少标准答案 .xlsx 底稿")
    else:
        gt_sheets = read_xlsx_book(gt_xlsx_files[0])
        if len(gt_sheets) < 5:
            errors.append(
                f"[Gate 2 标准底稿工作表过少] {gt_xlsx_files[0].name} 仅有 {[s[0] for s in gt_sheets]} 共 {len(gt_sheets)} 张表（要求至少 5 张表，含底稿说明/口径、取数明细、分类复算、汇总勾稽、待补证与勘误）"
            )
        total_formulas = 0
        has_basis_col = False
        has_evidence_col = False
        for sname, srows in gt_sheets:
            for r_i, row in enumerate(srows):
                for c in row:
                    if c["formula"]:
                        total_formulas += 1
                    if c["value"] in EXCEL_ERRORS or c["type"] == "e":
                        errors.append(
                            f"[Gate 2 标准底稿存在公式报错] {gt_xlsx_files[0].name} [{sname}!{c['coord']}] 值为 {c['value']}"
                        )
                    if r_i <= 3:
                        if "依据" in c["value"]:
                            has_basis_col = True
                        if any(k in c["value"] for k in ("证据", "索引", "出处", "定位")):
                            has_evidence_col = True
        if total_formulas < 30:
            errors.append(f"[Gate 2 标准底稿公式过少] 公式单元格仅 {total_formulas} 个（要求 ≥ 30 个）")
        if not (has_basis_col and has_evidence_col):
            errors.append(
                f"[Gate 2 标准底稿缺少追溯列] {gt_xlsx_files[0].name} 表头必须包含『依据』与『证据索引/出处』列"
            )

    if not gt_docx_files:
        errors.append("[Gate 2] ground_truth/ 缺少标准答案 .docx 意见书/报告")
    else:
        dinfo = read_docx_info(gt_docx_files[0])
        if dinfo["body_chinese_chars"] < 2200:
            errors.append(
                f"[Gate 2 标准答案 Word 篇幅严重缩水] {gt_docx_files[0].name} 正文汉字数（不含表格）仅 {dinfo['body_chinese_chars']} 字（要求 ≥ 2,200 字，推荐 2,500~5,000 字）"
            )
        if len(dinfo["tables"]) < 2:
            errors.append(
                f"[Gate 2 标准答案 Word 缺少汇总表] {gt_docx_files[0].name} 仅有 {len(dinfo['tables'])} 张表格（要求至少包含 2 张关键汇总/对照表）"
            )

    return errors


# ==============================================================================
# Gate 3 & Gate 4：体验评测包校验器
# ==============================================================================

def validate_eval_pack(eval_dir: Path, rubric_xlsx: Path) -> List[str]:
    """校验体验评测包目录是否符合 5 表体验评测记录.xlsx 与 6 节体验评测报告.docx 规范。"""
    errors: List[str] = []
    if not eval_dir.is_dir():
        return [f"[Gate 3] 评测目录不存在: {eval_dir}"]

    deliverables_dir = eval_dir / "被测交付物"
    if not deliverables_dir.is_dir():
        errors.append("[Gate 3] 缺少 被测交付物/ 子目录")
    else:
        d_xlsx = list(deliverables_dir.glob("*.xlsx"))
        d_docx = list(deliverables_dir.glob("*.docx"))
        missing_note = deliverables_dir / "缺失交付物说明.txt"
        if not d_xlsx and not d_docx:
            errors.append("[Gate 3] 被测交付物/ 下未找到任何 .xlsx 或 .docx 被测模型原件")
        elif (not d_xlsx or not d_docx) and not missing_note.is_file():
            errors.append(
                "[Gate 3] 被测交付物/ 缺少 .xlsx 或 .docx 原件（若系被测模型盲测超时/漏产，须如实留空并附『缺失交付物说明.txt』说明运行事实）"
            )

    trace_md = eval_dir / "操作轨迹摘要.md"
    if not trace_md.is_file():
        errors.append("[Gate 3] 缺少 操作轨迹摘要.md")

    eval_xlsx = eval_dir / "体验评测记录.xlsx"
    if not eval_xlsx.is_file():
        errors.append("[Gate 3] 缺少 体验评测记录.xlsx")
    else:
        sheets = read_xlsx_book(eval_xlsx)
        snames = [s[0] for s in sheets]
        if snames != EVAL_SHEETS_EXPECTED:
            errors.append(
                f"[Gate 3 评测记录表结构错误] {eval_dir.name}/体验评测记录.xlsx 工作表为 {snames}，必须严格为 5 张表: {EVAL_SHEETS_EXPECTED}"
            )
        else:
            # 检查基本信息
            info_rows = sheets[0][1]
            if len(info_rows) < 15:
                errors.append(f"[Gate 3] 『基本信息』行数仅 {len(info_rows)} 行（要求 ≥ 15 行）")
            # 检查逐项评分 11 列并与 rubric 对齐
            scoring_rows = sheets[2][1]
            if not scoring_rows:
                errors.append("[Gate 3] 『逐项评分』为空")
            else:
                shdr = [c["value"] for c in scoring_rows[0]]
                if len(shdr) != 11 or shdr[:9] != EVAL_SCORING_HEADERS[:9]:
                    errors.append(
                        f"[Gate 3 逐项评分列数错误] 『逐项评分』表头为 {shdr}，必须为 11 列标准表头: {EVAL_SCORING_HEADERS}"
                    )
                if rubric_xlsx.is_file():
                    r_sheets = read_xlsx_book(rubric_xlsx)
                    r_rows = [
                        [c["value"] for c in r][:7]
                        for r in r_sheets[0][1][1:]
                        if any(c["value"] for c in r) and not any(k in r[0]["value"] for k in ("合计", "总计", "总分"))
                    ]
                    s_data_rows = [
                        [c["value"] for c in r][:7]
                        for r in scoring_rows[1:]
                        if any(c["value"] for c in r) and not any(k in r[0]["value"] for k in ("合计", "总计", "总分"))
                    ]
                    if len(r_rows) != len(s_data_rows):
                        errors.append(
                            f"[Gate 3] 『逐项评分』条目数（{len(s_data_rows)}）与 rubric.xlsx 条目数（{len(r_rows)}）不一致"
                        )
            # 检查定性反馈 11 行
            fb_rows = sheets[4][1]
            if len(fb_rows) < 11:
                errors.append(
                    f"[Gate 3 定性反馈缩水] 『定性反馈』仅有 {len(fb_rows)} 行（要求包含表头 + 长处/短处/典型失误模式/7大能力维度/区分度评价 共 12 行）"
                )

    eval_docx = eval_dir / "体验评测报告.docx"
    if not eval_docx.is_file():
        errors.append("[Gate 3] 缺少 体验评测报告.docx")
    else:
        dinfo = read_docx_info(eval_docx)
        if dinfo["body_chinese_chars"] < 1200:
            errors.append(
                f"[Gate 3 评测报告严重缩水] {eval_dir.name}/体验评测报告.docx 正文汉字数仅 {dinfo['body_chinese_chars']} 字（要求 ≥ 1,200 字，完整包含 6 大章节）"
            )
        if len(dinfo["tables"]) < 1:
            errors.append(f"[Gate 3] {eval_dir.name}/体验评测报告.docx 缺少汇总/差异对比表格")

    return errors


# ==============================================================================
# 通用生成器 1：build-rubric（从 JSON 自动生成 7 列评分表 + 7 行评分说明双表）
# ==============================================================================

def build_rubric_from_spec(spec_path: Path, out_dir: Path) -> None:
    """读取单题 rubric_spec.json，自动统计分类与维度分布并生成 rubric.xlsx 与 rubric.csv。"""
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    items: List[Dict[str, Any]] = spec["items"]

    # 校验并统计
    pos_items = [it for it in items if it["type"] == "得分"]
    neg_items = [it for it in items if it["type"] == "扣分"]
    pos_sum = sum(int(it["points"]) for it in pos_items)
    if pos_sum != 100:
        raise ValueError(f"得分项分值合计必须为 100，当前为 {pos_sum}")

    # 统计条目性质（判断深度、细节完整性、跨材料综合、覆盖与交付规范）
    cat_buckets: Dict[str, List[Tuple[str, int]]] = {
        "判断深度": [],
        "细节完整性": [],
        "跨材料综合": [],
        "覆盖与交付规范": [],
    }
    dim_buckets: Dict[str, List[Tuple[str, int]]] = {}
    for it in pos_items:
        code = it["id"]
        pts = int(it["points"])
        cat = it.get("category", "细节完整性")
        cat_buckets.setdefault(cat, []).append((code, pts))
        dim = it["dimension"]
        dim_buckets.setdefault(dim, []).append((code, pts))

    def fmt_bucket(pairs: List[Tuple[str, int]]) -> Tuple[str, int]:
        total = sum(p for _, p in pairs)
        codes = "、".join(f"{c}（{p}）" for c, p in pairs)
        return codes, total

    jd_str, jd_pts = fmt_bucket(cat_buckets["判断深度"])
    dt_str, dt_pts = fmt_bucket(cat_buckets["细节完整性"])
    cm_str, cm_pts = fmt_bucket(cat_buckets["跨材料综合"])
    cv_str, cv_pts = fmt_bucket(cat_buckets["覆盖与交付规范"])
    core_three_pts = jd_pts + dt_pts + cm_pts

    dim_parts = []
    for d in sorted(dim_buckets.keys()):
        c_list = "、".join(c for c, _ in dim_buckets[d])
        d_sum = sum(p for _, p in dim_buckets[d])
        dim_parts.append(f"{d}：{c_list}（{d_sum}分）")
    dim_summary_text = "；".join(dim_parts) + "。"

    # Sheet 1 行数据
    s1_rows: List[List[Any]] = [RUBRIC_HEADERS]
    for it in items:
        s1_rows.append([
            it["id"],
            it["dimension"],
            it["check"],
            it["standard"],
            int(it["points"]),
            it["type"],
            it["location"],
        ])

    # Sheet 2 评分说明（7 行任务级规则，自动合成统计项）
    neg_sum = abs(sum(int(it["points"]) for it in neg_items))
    max_single = max(int(it["points"]) for it in pos_items)
    s2_rows: List[List[Any]] = [
        ["项目", "说明"],
        [
            "1. 总分计算",
            f"得分项共{len(pos_items)}条（{pos_items[0]['id']}—{pos_items[-1]['id']}），正向满分100分；"
            f"每条二元判定，达到得全部分值，未达到得0分。扣分项{len(neg_items)}条（{neg_items[0]['id']}—{neg_items[-1]['id']}），"
            f"触发即按第5条差额或独立规则扣分，每条最多扣一次。最终得分＝得分项合计＋扣分项合计（负数），最低0分。",
        ],
        ["2. 整项判定原则", spec["instructions"]["judgment_principle"]],
        ["3. 列明的等价处理（封闭清单）", spec["instructions"]["accepted_alternatives"]],
        ["4. 扣分不连锁原则", spec["instructions"]["non_cascading_rule"]],
        ["5. 扣分项与得分项的关系", spec["instructions"]["deduction_rule"]],
        [
            "6. 检查项总数与条目性质",
            f"检查项共{len(items)}条：得分项{len(pos_items)}条，合计100分，单条最高{max_single}分；"
            f"扣分项{len(neg_items)}条，合计最多扣{neg_sum}分。"
            f"条目性质（核心三类合计{core_three_pts}分）："
            f"判断深度类——{jd_str}，计{jd_pts}分；"
            f"细节完整性类——{dt_str}，计{dt_pts}分；"
            f"跨材料综合类——{cm_str}，计{cm_pts}分；"
            f"其余为覆盖与交付规范类——{cv_str}，计{cv_pts}分。",
        ],
        ["7. 维度分布", dim_summary_text],
    ]

    out_dir.mkdir(parents=True, exist_ok=True)
    write_ooxml_xlsx(
        out_dir / "rubric.xlsx",
        [
            {"name": "评分表", "widths": [10, 24, 28, 65, 8, 8, 38], "rows": s1_rows},
            {"name": "评分说明", "widths": [26, 100], "rows": s2_rows},
        ],
    )
    with (out_dir / "rubric.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(s1_rows)


# ==============================================================================
# 通用生成器 2：build-eval（从 JSON 自动生成 5 表评测记录 + 6 节评测报告 + 轨迹摘要）
# ==============================================================================

def build_eval_from_spec(spec_path: Path, out_dir: Path) -> None:
    """读取评测配置 JSON 与 rubric.xlsx，一键生成完全合规的体验评测包文件。"""
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    rubric_xlsx = Path(spec["rubric_xlsx"])
    r_sheets = read_xlsx_book(rubric_xlsx)
    r_rows = [
        [c["value"] for c in r][:7]
        for r in r_sheets[0][1][1:]
        if any(c["value"] for c in r) and r[0]["value"] not in ("合计", "总计")
    ]

    judgments: Dict[str, Dict[str, Any]] = {j["id"]: j for j in spec["judgments"]}
    scoring_rows: List[List[Any]] = [EVAL_SCORING_HEADERS]

    pos_passed = 0
    pos_total = 0
    neg_triggered = 0
    total_score = 0
    pos_end_row = 1
    for idx, r in enumerate(r_rows, start=2):
        code, dim, check, std, pts_str, rtype, loc = r
        pts = int(float(pts_str))
        j = judgments[code]
        status = j["status"]  # 得分项: 达到/未达到; 扣分项: 未触发/触发
        actual_pts = int(j["score"])
        total_score += actual_pts
        if rtype == "得分":
            pos_total += 1
            pos_end_row = idx
            if status == "达到":
                pos_passed += 1
            formula_str = f'IF(H{idx}="达到",E{idx},0)'
        else:
            if status == "触发":
                neg_triggered += 1
            formula_str = f'IF(H{idx}="触发",{actual_pts},0)'

        scoring_rows.append([
            code,
            dim,
            check,
            std,
            pts,
            rtype,
            loc,
            status,
            {"f": formula_str, "v": actual_pts},
            j["reason"],
            j.get("note", ""),
        ])

    last_row_idx = len(r_rows) + 1
    total_row_idx = last_row_idx + 1
    scoring_rows.append([
        "合计",
        "",
        "",
        "",
        {"f": f"SUM(E2:E{pos_end_row})", "v": 100},
        "",
        "",
        "",
        {"f": f"MAX(0,SUM(I2:I{last_row_idx}))", "v": max(0, total_score)},
        f"得分项 {pos_passed}/{pos_total} 达到；扣分项触发 {neg_triggered} 项。",
        "",
    ])

    # Sheet 1: 基本信息（16 行标准结构）
    b = spec["basic_info"]
    info_rows = [
        ["项目", "内容"],
        ["评测对象", b["target"]],
        ["被测模型", b["model"]],
        ["运行环境", b["environment"]],
        ["运行说明", b["run_note"]],
        ["开始时间", b["start_time"]],
        ["结束时间", b["end_time"]],
        ["耗时", b["duration"]],
        ["工具调用次数", b["tool_calls"]],
        ["令牌用量", b["token_usage"]],
        ["交付文件清单", b["deliverables"]],
        ["被测交付物副本", b["copy_note"]],
        ["评测方复算", b["recalc_note"]],
        ["评分依据", b["rubric_note"]],
        ["总分", {"f": f"'逐项评分'!I{total_row_idx}", "v": max(0, total_score)}],
        ["评测人与日期", f"{b['evaluator']}（{b['date']}）"],
    ]

    # Sheet 2: 关键操作轨迹
    traj_rows = [["序号", "事件位置", "操作", "方式/命令要点", "结果", "对结果的影响"]]
    for t in spec["trajectory"]:
        traj_rows.append([t["seq"], t["event"], t["action"], t["command"], t["result"], t["impact"]])

    # Sheet 4: 问题记录
    issue_rows = [["编号", "分类", "出现位置", "问题描述", "影响", "是否影响主要结论"]]
    for iss in spec["issues"]:
        issue_rows.append([
            iss["id"],
            iss["category"],
            iss["location"],
            iss["description"],
            iss["impact"],
            iss["affects_main"],
        ])

    # Sheet 5: 定性反馈（12 行标准结构）
    fb = spec["qualitative_feedback"]
    fb_rows = [
        ["方面", "反馈"],
        ["长处", fb["strengths"]],
        ["短处", fb["weaknesses"]],
        ["典型失误模式", fb["error_patterns"]],
        ["专业知识理解", fb["domain_knowledge"]],
        ["多步推理", fb["multi_step_reasoning"]],
        ["多文档整合", fb["multi_doc_integration"]],
        ["工具使用", fb["tool_use"]],
        ["风险识别", fb["risk_identification"]],
        ["方案判断", fb["solution_judgment"]],
        ["结果交付", fb["delivery_quality"]],
        ["区分度评价", fb["discrimination"]],
    ]

    out_dir.mkdir(parents=True, exist_ok=True)
    write_ooxml_xlsx(
        out_dir / "体验评测记录.xlsx",
        [
            {"name": "基本信息", "widths": [22, 95], "rows": info_rows},
            {"name": "关键操作轨迹", "widths": [8, 16, 18, 48, 18, 32], "rows": traj_rows},
            {
                "name": "逐项评分",
                "widths": [8, 22, 24, 52, 8, 8, 30, 10, 8, 52, 22],
                "rows": scoring_rows,
            },
            {"name": "问题记录", "widths": [8, 22, 28, 50, 36, 16], "rows": issue_rows},
            {"name": "定性反馈", "widths": [20, 95], "rows": fb_rows},
        ],
    )

    # 写入 操作轨迹摘要.md
    md_lines = [
        f"# 操作轨迹摘要（{spec['report_title']}）",
        "",
        f"- 运行：{b['start_time']} 至 {b['end_time']}，耗时 {b['duration']}",
        f"- 工具与命令：{b['tool_calls']}",
        f"- 令牌用量：{b['token_usage']}",
        f"- 交付文件：{b['deliverables']}",
        f"- 主口径结果与总分：{spec['headline_result']}（总分 {max(0, total_score)} / 100）",
        "",
        "| 序号 | 事件位置 | 操作 | 方式/命令要点 | 结果 | 对结果的影响 |",
        "|---|---|---|---|---|---|",
    ]
    for t in spec["trajectory"]:
        md_lines.append(
            f"| {t['seq']} | {t['event']} | {t['action']} | {t['command']} | {t['result']} | {t['impact']} |"
        )
    (out_dir / "操作轨迹摘要.md").write_text("\n".join(md_lines) + "\n", encoding="utf-8")

    # 写入 体验评测报告.docx
    write_ooxml_docx(
        out_dir / "体验评测报告.docx",
        title=spec["report_title"],
        subtitle=spec["report_subtitle"],
        sections=spec["report_sections"],
    )


# ==============================================================================
# 通用校验与冻结工具：recalc-xlsx / freeze-task / verify-freeze
# ==============================================================================

class _FormulaErr(Exception):
    """表示 Excel 公式求值过程中遇到的标准错误值。"""


class _BlankCell(int):
    """空单元格占位：参与算术时为 0，COUNTA/COUNT 不计入，等价比较兼容 0 与空串。"""

    def __repr__(self) -> str:
        return "Blank"


_BLANK = _BlankCell(0)


class _RangeArray(list):
    """保留二维行列形状信息的单元格区域数组，向下兼容一维 list。"""

    def __init__(self, items: List[Any], nrows: int, ncols: int) -> None:
        super().__init__(items)
        self.nrows = nrows
        self.ncols = ncols


def _excel_type_rank(val: Any) -> int:
    """Excel 有序比较类型优先级：数值(0) < 文本(1) < 布尔(2)。"""
    if isinstance(val, bool):
        return 2
    if isinstance(val, (int, float)):
        return 0
    return 1


def _excel_num(val: Any) -> float:
    if val is None or isinstance(val, _BlankCell):
        return 0
    if isinstance(val, bool):
        return int(val)
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        s = val.strip()
        if s == "":
            return 0
        if s in EXCEL_ERRORS:
            raise _FormulaErr(s)
        try:
            f = float(s)
            return int(f) if f.is_integer() and "." not in s else f
        except ValueError:
            pass
    raise _FormulaErr(f"#VALUE!（非数值参与算术: {val!r}）")


def _excel_str(val: Any) -> str:
    if val is None or isinstance(val, _BlankCell):
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(val)
    return str(val)


def _excel_bool(val: Any) -> bool:
    if val is None or isinstance(val, _BlankCell):
        return False
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val != 0
    if isinstance(val, str):
        u = val.strip().upper()
        if u in ("TRUE", ""):
            return u == "TRUE"
        if u == "FALSE":
            return False
    raise _FormulaErr(f"#VALUE!（无法转为布尔值: {val!r}）")


def _excel_cmp(left: Any, right: Any, op: str) -> bool:
    """按 Excel 语义执行 =, <>, <, <=, >, >= 比较。"""
    if isinstance(left, _BlankCell) and isinstance(right, _BlankCell):
        return op in ("=", "<=", ">=")
    if isinstance(left, _BlankCell):
        left = 0 if isinstance(right, (int, float)) and not isinstance(right, bool) else ""
    if isinstance(right, _BlankCell):
        right = 0 if isinstance(left, (int, float)) and not isinstance(left, bool) else ""

    if op == "=":
        if isinstance(left, str) and isinstance(right, str):
            return left.lower() == right.lower()
        if type(left) is bool or type(right) is bool:
            return type(left) is type(right) and left == right
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            return abs(left - right) < 1e-12
        return False
    if op == "<>":
        return not _excel_cmp(left, right, "=")

    # 有序比较：不同类型遵循 Excel 规则（数值 < 文本 < 布尔），同类型按对应语义比较
    rl, rr = _excel_type_rank(left), _excel_type_rank(right)
    if rl != rr:
        lv_cmp: Any = rl
        rv_cmp: Any = rr
    elif rl == 0:
        lv_cmp, rv_cmp = left, right
    elif rl == 2:
        lv_cmp, rv_cmp = int(left), int(right)
    else:
        lv_cmp, rv_cmp = _excel_str(left).lower(), _excel_str(right).lower()
    return {"<": lv_cmp < rv_cmp, "<=": lv_cmp <= rv_cmp, ">": lv_cmp > rv_cmp, ">=": lv_cmp >= rv_cmp}[op]


def _make_crit(cond: Any):
    """构造 SUMIF/COUNTIF 条件匹配器。"""
    if isinstance(cond, _BlankCell):
        return lambda x: isinstance(x, _BlankCell) or x == ""
    s = _excel_str(cond)
    m = re.match(r"^(<=|>=|<>|<|>|=)?(.*)$", s)
    op = m.group(1) or "="
    raw = m.group(2)
    try:
        num_val: Optional[float] = float(raw)
    except ValueError:
        num_val = None

    def _match(x: Any) -> bool:
        if isinstance(x, _BlankCell):
            return op == "=" and raw == ""
        if num_val is not None and op in ("=", "<>", "<", "<=", ">", ">="):
            if isinstance(x, (int, float)) and not isinstance(x, bool):
                if op == "=":
                    return abs(x - num_val) < 1e-12
                if op == "<>":
                    return abs(x - num_val) >= 1e-12
                return {"<": x < num_val, "<=": x <= num_val, ">": x > num_val, ">=": x >= num_val}[op]
            return op == "<>"
        if op in ("=", "<>") and ("*" in raw or "?" in raw):
            pat = "^" + re.escape(raw).replace(r"\*", ".*").replace(r"\?", ".") + "$"
            matched = bool(re.match(pat, _excel_str(x), flags=re.IGNORECASE))
            return matched if op == "=" else not matched
        xs = _excel_str(x).lower()
        rs = raw.lower()
        return {"=": xs == rs, "<>": xs != rs, "<": xs < rs, "<=": xs <= rs, ">": xs > rs, ">=": xs >= rs}[op]

    return _match


def recalc_workbook(xlsx_path: Path) -> Dict[str, Any]:
    """
    在无 LibreOffice 环境下使用 AST 解析并重算 .xlsx 全部公式单元格，
    同步核对文件自带缓存值并汇总关键工作表/总计值。
    """
    import openpyxl
    from openpyxl.formula.tokenizer import Token, Tokenizer
    from openpyxl.utils import column_index_from_string, get_column_letter

    wb_f = openpyxl.load_workbook(xlsx_path, data_only=False)
    wb_v = openpyxl.load_workbook(xlsx_path, data_only=True)
    memo: Dict[Tuple[str, str], Any] = {}
    visiting = set()

    def split_addr(addr: str) -> Tuple[int, int]:
        m = re.match(r"^\$?([A-Za-z]+)\$?(\d+)$", addr.strip())
        if not m:
            raise _FormulaErr(f"#REF!（无效单元格地址: {addr!r}）")
        return column_index_from_string(m.group(1).upper()), int(m.group(2))

    def get_cell(sheet_name: str, addr: str) -> Any:
        if sheet_name not in wb_f.sheetnames:
            raise _FormulaErr(f"#REF!（工作表不存在: {sheet_name}）")
        norm = addr.replace("$", "").upper()
        key = (sheet_name, norm)
        if key in memo:
            return memo[key]
        if key in visiting:
            raise _FormulaErr("#CIRC!（检测到公式循环引用）")
        col_idx, row_idx = split_addr(norm)
        raw = wb_f[sheet_name].cell(row_idx, col_idx).value
        if raw is None:
            memo[key] = _BLANK
            return _BLANK
        if isinstance(raw, str) and raw.startswith("="):
            visiting.add(key)
            try:
                val = eval_formula_str(sheet_name, raw[1:])
            finally:
                visiting.discard(key)
            memo[key] = val
            return val
        memo[key] = raw
        return raw

    def get_range(sheet_name: str, start_addr: str, end_addr: str) -> _RangeArray:
        c1, r1 = split_addr(start_addr)
        c2, r2 = split_addr(end_addr)
        c_lo, c_hi = min(c1, c2), max(c1, c2)
        r_lo, r_hi = min(r1, r2), max(r1, r2)
        items = [
            get_cell(sheet_name, f"{get_column_letter(c)}{r}")
            for r in range(r_lo, r_hi + 1)
            for c in range(c_lo, c_hi + 1)
        ]
        return _RangeArray(items, r_hi - r_lo + 1, c_hi - c_lo + 1)

    def flatten(args: Tuple[Any, ...]) -> List[Any]:
        out: List[Any] = []
        for a in args:
            if isinstance(a, list):
                out.extend(a)
            else:
                out.append(a)
        return out

    def parse_ast(default_sheet: str, formula_body: str):
        tok = Tokenizer("=" + formula_body)
        items = [t for t in tok.items if t.type != Token.WSPACE]
        pos = [0]

        def peek() -> Optional[Token]:
            return items[pos[0]] if pos[0] < len(items) else None

        def consume() -> Token:
            t = items[pos[0]]
            pos[0] += 1
            return t

        def parse_expr(min_prec: int = 1):
            t = peek()
            if t is None:
                raise _FormulaErr("#VALUE!（空表达式）")
            if t.type == Token.OP_PRE:
                op = consume().value
                node = parse_expr(6)
                node = ("unary", op, node)
            elif t.type == Token.PAREN and t.subtype == Token.OPEN:
                consume()
                node = parse_expr(1)
                if peek() and peek().type == Token.PAREN and peek().subtype == Token.CLOSE:
                    consume()
            elif t.type == Token.FUNC and t.subtype == Token.OPEN:
                fn_name = consume().value[:-1].upper()
                args = []
                if peek() and not (peek().type == Token.FUNC and peek().subtype == Token.CLOSE):
                    while True:
                        args.append(parse_expr(1))
                        if peek() and peek().type == Token.SEP:
                            consume()
                        else:
                            break
                if peek() and peek().type == Token.FUNC and peek().subtype == Token.CLOSE:
                    consume()
                node = ("call", fn_name, args)
            elif t.type == Token.OPERAND:
                consume()
                if t.subtype == Token.RANGE:
                    m = re.match(r"^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$", t.value)
                    if m:
                        s_name = (m.group(1).replace("''", "'") if m.group(1) else m.group(2))
                        ref = m.group(3)
                    else:
                        s_name, ref = default_sheet, t.value
                    if ":" in ref:
                        a, b = ref.split(":", 1)
                        node = ("range", s_name, a, b)
                    else:
                        node = ("cell", s_name, ref)
                elif t.subtype == Token.TEXT:
                    node = ("lit", t.value[1:-1].replace('""', '"'))
                elif t.subtype == Token.LOGICAL:
                    node = ("lit", t.value.upper() == "TRUE")
                elif t.subtype == Token.ERROR:
                    node = ("err", t.value)
                else:
                    raw_num = t.value
                    f = float(raw_num)
                    node = ("lit", int(f) if f.is_integer() and "." not in raw_num and "e" not in raw_num.lower() else f)
            else:
                raise _FormulaErr(f"#VALUE!（未识别语法单元: {t.type}/{t.value!r}）")

            while peek() and peek().type == Token.OP_POST:
                consume()
                node = ("post%", node)

            prec_map = {"=": 1, "<>": 1, "<": 1, "<=": 1, ">": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}
            while peek() and peek().type == Token.OP_IN:
                op = peek().value
                prec = prec_map.get(op)
                if prec is None or prec < min_prec:
                    break
                consume()
                next_min = prec if op == "^" else prec + 1
                right = parse_ast_or_expr(next_min)
                node = ("bin", op, node, right)
                while peek() and peek().type == Token.OP_POST:
                    consume()
                    node = ("post%", node)
            return node

        def parse_ast_or_expr(min_prec: int):
            return parse_expr(min_prec)

        return parse_expr(1)

    def _is_err_val(v: Any) -> bool:
        return isinstance(v, str) and (
            v in EXCEL_ERRORS or any(v.startswith(e) for e in EXCEL_ERRORS) or v.startswith("#CIRC!")
        )

    def _wrap_shape(items: List[Any], ref_a: Any, ref_b: Any = None) -> Any:
        for ref in (ref_a, ref_b):
            if isinstance(ref, _RangeArray) and len(items) == len(ref):
                return _RangeArray(items, ref.nrows, ref.ncols)
        return items

    def _vec1(val: Any, fn) -> Any:
        if isinstance(val, list):
            return _wrap_shape([fn(x) for x in val], val)
        return fn(val)

    def _apply_bin_scalar(op: str, lv: Any, rv: Any) -> Any:
        if op in ("=", "<>", "<", "<=", ">", ">="):
            return _excel_cmp(lv, rv, op)
        if op == "&":
            return _excel_str(lv) + _excel_str(rv)
        a, b = _excel_num(lv), _excel_num(rv)
        if op == "+":
            return a + b
        if op == "-":
            return a - b
        if op == "*":
            return a * b
        if op == "/":
            if b == 0:
                raise _FormulaErr("#DIV/0!")
            return a / b
        if op == "^":
            return a ** b
        raise _FormulaErr(f"#VALUE!（未知运算符: {op}）")

    def _apply_bin(op: str, lv: Any, rv: Any) -> Any:
        l_is_list = isinstance(lv, list)
        r_is_list = isinstance(rv, list)
        if not l_is_list and not r_is_list:
            return _apply_bin_scalar(op, lv, rv)
        l_arr = lv if l_is_list else [lv]
        r_arr = rv if r_is_list else [rv]
        if len(l_arr) != len(r_arr):
            if len(l_arr) == 1:
                l_arr = [l_arr[0]] * len(r_arr)
            elif len(r_arr) == 1:
                r_arr = [r_arr[0]] * len(l_arr)
            else:
                raise _FormulaErr("#VALUE!（数组长度不一致）")
        out = [_apply_bin_scalar(op, x, y) for x, y in zip(l_arr, r_arr)]
        return _wrap_shape(out, lv, rv)

    def _excel_round_num(val: float, ndigits: int, mode: str) -> float:
        factor = 10 ** ndigits
        scaled = val * factor
        if mode in ("up", "down") and abs(scaled - round(scaled)) < 1e-11:
            r_int = int(round(scaled))
        elif mode == "up":
            r_int = math.ceil(scaled) if scaled >= 0 else math.floor(scaled)
        elif mode == "down":
            r_int = math.floor(scaled) if scaled >= 0 else math.ceil(scaled)
        else:
            r_int = math.floor(scaled + 0.5) if scaled >= 0 else math.ceil(scaled - 0.5)
        if ndigits <= 0:
            return int(round(r_int / factor)) if ndigits < 0 else int(r_int)
        return r_int / factor

    def _lookup_match_exact(item: Any, target: Any) -> bool:
        if isinstance(item, _BlankCell) != isinstance(target, _BlankCell):
            return False
        if isinstance(target, str) and ("*" in target or "?" in target):
            pat = "^" + re.escape(target).replace(r"\*", ".*").replace(r"\?", ".") + "$"
            return bool(re.match(pat, _excel_str(item), flags=re.IGNORECASE))
        return _excel_cmp(item, target, "=")

    def eval_node(sh: str, node: Tuple[Any, ...]) -> Any:
        kind = node[0]
        if kind == "lit":
            return node[1]
        if kind == "err":
            raise _FormulaErr(node[1])
        if kind == "cell":
            return get_cell(node[1], node[2])
        if kind == "range":
            return get_range(node[1], node[2], node[3])
        if kind == "post%":
            return _vec1(eval_node(sh, node[1]), lambda x: _excel_num(x) / 100.0)
        if kind == "unary":
            op_u = node[1]
            return _vec1(eval_node(sh, node[2]), lambda x: (-_excel_num(x) if op_u == "-" else _excel_num(x)))
        if kind == "bin":
            op = node[1]
            lv = eval_node(sh, node[2])
            rv = eval_node(sh, node[3])
            return _apply_bin(op, lv, rv)
        if kind == "call":
            fn = node[1]
            args = node[2]
            if fn == "IF":
                cond = eval_node(sh, args[0]) if args else False
                if not isinstance(cond, list):
                    if _excel_bool(cond):
                        return eval_node(sh, args[1]) if len(args) > 1 else True
                    return eval_node(sh, args[2]) if len(args) > 2 else False
                t_val = eval_node(sh, args[1]) if len(args) > 1 else True
                f_val = eval_node(sh, args[2]) if len(args) > 2 else False
                n = len(cond)
                t_arr = t_val if isinstance(t_val, list) else [t_val] * n
                f_arr = f_val if isinstance(f_val, list) else [f_val] * n
                if len(t_arr) == 1 and n > 1:
                    t_arr = [t_arr[0]] * n
                if len(f_arr) == 1 and n > 1:
                    f_arr = [f_arr[0]] * n
                if len(t_arr) != n or len(f_arr) != n:
                    raise _FormulaErr("#VALUE!（IF 数组分支长度不一致）")
                out = [t_arr[i] if _excel_bool(cond[i]) else f_arr[i] for i in range(n)]
                return _wrap_shape(out, cond)
            if fn == "IFERROR":
                try:
                    v = eval_node(sh, args[0])
                    if _is_err_val(v):
                        raise _FormulaErr(v)
                    return v
                except _FormulaErr:
                    return eval_node(sh, args[1]) if len(args) > 1 else ""
            if fn == "IFNA":
                try:
                    v = eval_node(sh, args[0])
                    if isinstance(v, str) and (v == "#N/A" or v.startswith("#N/A")):
                        raise _FormulaErr(v)
                    return v
                except _FormulaErr as exc:
                    if "#N/A" in str(exc):
                        return eval_node(sh, args[1]) if len(args) > 1 else ""
                    raise
            if fn == "ISERROR":
                try:
                    v = eval_node(sh, args[0])
                except _FormulaErr:
                    return True
                return _vec1(v, _is_err_val)

            ev = [eval_node(sh, a) for a in args]
            if fn == "SUM":
                vals = []
                for x in flatten(tuple(ev)):
                    if isinstance(x, (int, float)) and not isinstance(x, (_BlankCell, bool)):
                        vals.append(x)
                return sum(vals)
            if fn == "SUMIF":
                rng_vals = ev[0] if isinstance(ev[0], list) else [ev[0]]
                matcher = _make_crit(ev[1])
                sum_vals = (ev[2] if isinstance(ev[2], list) else [ev[2]]) if len(ev) > 2 else rng_vals
                return sum(
                    v for k, v in zip(rng_vals, sum_vals)
                    if matcher(k) and isinstance(v, (int, float)) and not isinstance(v, (_BlankCell, bool))
                )
            if fn == "SUMIFS":
                s_vals = ev[0] if isinstance(ev[0], list) else [ev[0]]
                pairs = [(ev[i] if isinstance(ev[i], list) else [ev[i]], _make_crit(ev[i + 1])) for i in range(1, len(ev), 2)]
                return sum(
                    s_vals[idx] for idx in range(len(s_vals))
                    if all(idx < len(r_v) and m_fn(r_v[idx]) for r_v, m_fn in pairs)
                    and isinstance(s_vals[idx], (int, float)) and not isinstance(s_vals[idx], (_BlankCell, bool))
                )
            if fn == "COUNTIF":
                rng_vals = ev[0] if isinstance(ev[0], list) else [ev[0]]
                matcher = _make_crit(ev[1])
                return sum(1 for k in rng_vals if matcher(k))
            if fn == "COUNTIFS":
                pairs = [(ev[i] if isinstance(ev[i], list) else [ev[i]], _make_crit(ev[i + 1])) for i in range(0, len(ev), 2)]
                if not pairs:
                    return 0
                n = len(pairs[0][0])
                return sum(1 for idx in range(n) if all(idx < len(r_v) and m_fn(r_v[idx]) for r_v, m_fn in pairs))
            if fn == "COUNTA":
                return sum(1 for x in flatten(tuple(ev)) if not isinstance(x, _BlankCell) and x != "")
            if fn == "COUNT":
                return sum(1 for x in flatten(tuple(ev)) if isinstance(x, (int, float)) and not isinstance(x, (_BlankCell, bool)))
            if fn in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
                ndigits = int(_excel_num(ev[1])) if len(ev) > 1 else 0
                rmode = "round" if fn == "ROUND" else ("up" if fn == "ROUNDUP" else "down")
                return _vec1(ev[0], lambda x: _excel_round_num(_excel_num(x), ndigits, rmode))
            if fn == "MOD":
                dv = _excel_num(ev[1])
                if dv == 0:
                    raise _FormulaErr("#DIV/0!")
                return _excel_num(ev[0]) % dv
            if fn == "AND":
                items_f = [x for x in flatten(tuple(ev)) if not isinstance(x, _BlankCell)]
                return all(_excel_bool(x) for x in items_f) if items_f else False
            if fn == "OR":
                items_f = [x for x in flatten(tuple(ev)) if not isinstance(x, _BlankCell)]
                return any(_excel_bool(x) for x in items_f)
            if fn == "NOT":
                return _vec1(ev[0], lambda x: not _excel_bool(x))
            if fn == "ABS":
                return _vec1(ev[0], lambda x: abs(_excel_num(x)))
            if fn == "SIGN":
                return _vec1(ev[0], lambda x: (1 if _excel_num(x) > 0 else (-1 if _excel_num(x) < 0 else 0)))
            if fn == "N":
                return _vec1(
                    ev[0],
                    lambda x: (int(x) if isinstance(x, bool) else (x if isinstance(x, (int, float)) and not isinstance(x, _BlankCell) else 0)),
                )
            if fn == "MAX":
                nums = [x for x in flatten(tuple(ev)) if isinstance(x, (int, float)) and not isinstance(x, (_BlankCell, bool))]
                return max(nums) if nums else 0
            if fn == "MIN":
                nums = [x for x in flatten(tuple(ev)) if isinstance(x, (int, float)) and not isinstance(x, (_BlankCell, bool))]
                return min(nums) if nums else 0
            if fn == "AVERAGE":
                nums = [x for x in flatten(tuple(ev)) if isinstance(x, (int, float)) and not isinstance(x, (_BlankCell, bool))]
                if not nums:
                    raise _FormulaErr("#DIV/0!")
                return sum(nums) / len(nums)
            if fn == "INT":
                return _vec1(ev[0], lambda x: math.floor(_excel_num(x)))
            if fn == "LEN":
                return _vec1(ev[0], lambda x: len(_excel_str(x)))
            if fn == "TRIM":
                return _vec1(ev[0], lambda x: " ".join(_excel_str(x).split()))
            if fn == "LEFT":
                cnt = int(_excel_num(ev[1])) if len(ev) > 1 else 1
                if cnt < 0:
                    raise _FormulaErr("#VALUE!")
                return _vec1(ev[0], lambda x: _excel_str(x)[:cnt])
            if fn == "RIGHT":
                cnt = int(_excel_num(ev[1])) if len(ev) > 1 else 1
                if cnt < 0:
                    raise _FormulaErr("#VALUE!")
                return _vec1(ev[0], lambda x: "" if cnt == 0 else _excel_str(x)[-cnt:])
            if fn == "MID":
                start_p = int(_excel_num(ev[1]))
                cnt = int(_excel_num(ev[2]))
                if start_p < 1 or cnt < 0:
                    raise _FormulaErr("#VALUE!")
                return _vec1(ev[0], lambda x: _excel_str(x)[start_p - 1 : start_p - 1 + cnt])
            if fn in ("CONCAT", "CONCATENATE"):
                return "".join(_excel_str(x) for x in flatten(tuple(ev)))
            if fn == "ISNUMBER":
                return _vec1(ev[0], lambda x: isinstance(x, (int, float)) and not isinstance(x, (_BlankCell, bool)))
            if fn == "ISBLANK":
                return _vec1(ev[0], lambda x: isinstance(x, _BlankCell))
            if fn == "ISTEXT":
                return _vec1(ev[0], lambda x: isinstance(x, str) and not _is_err_val(x))
            if fn == "MATCH":
                lookup_val = ev[0]
                arr = ev[1] if isinstance(ev[1], list) else [ev[1]]
                mtype = int(_excel_num(ev[2])) if len(ev) > 2 else 1
                if mtype == 0:
                    for idx, item in enumerate(arr, start=1):
                        if _lookup_match_exact(item, lookup_val):
                            return idx
                    raise _FormulaErr("#N/A")
                t_rank = _excel_type_rank(lookup_val)
                best_idx = None
                for idx, item in enumerate(arr, start=1):
                    if isinstance(item, _BlankCell) or _excel_type_rank(item) != t_rank:
                        continue
                    if mtype > 0:
                        if _excel_cmp(item, lookup_val, "<="):
                            best_idx = idx
                        else:
                            break
                    else:
                        if _excel_cmp(item, lookup_val, ">="):
                            best_idx = idx
                        else:
                            break
                if best_idx is None:
                    raise _FormulaErr("#N/A")
                return best_idx
            if fn == "INDEX":
                arr = ev[0]
                if not isinstance(arr, list):
                    r_num = int(_excel_num(ev[1])) if len(ev) > 1 else 1
                    c_num = int(_excel_num(ev[2])) if len(ev) > 2 else 1
                    if r_num in (0, 1) and c_num in (0, 1):
                        return arr
                    raise _FormulaErr("#REF!")
                nrows = getattr(arr, "nrows", len(arr))
                ncols = getattr(arr, "ncols", 1)
                r_num = int(_excel_num(ev[1])) if len(ev) > 1 else 0
                if len(ev) <= 2:
                    if nrows == 1 and ncols > 1:
                        if r_num == 0:
                            return list(arr)
                        if 1 <= r_num <= ncols:
                            return arr[r_num - 1]
                        raise _FormulaErr("#REF!")
                    if r_num == 0:
                        return list(arr)
                    if 1 <= r_num <= nrows and ncols == 1:
                        return arr[r_num - 1]
                    if 1 <= r_num <= nrows and ncols > 1:
                        return arr[(r_num - 1) * ncols : r_num * ncols]
                    raise _FormulaErr("#REF!")
                c_num = int(_excel_num(ev[2]))
                if r_num < 0 or c_num < 0 or r_num > nrows or c_num > ncols:
                    raise _FormulaErr("#REF!")
                if r_num == 0 and c_num == 0:
                    return arr
                if r_num == 0:
                    return _RangeArray([arr[row_i * ncols + (c_num - 1)] for row_i in range(nrows)], nrows, 1)
                if c_num == 0:
                    return _RangeArray(arr[(r_num - 1) * ncols : r_num * ncols], 1, ncols)
                return arr[(r_num - 1) * ncols + (c_num - 1)]
            if fn == "VLOOKUP":
                lookup_val = ev[0]
                table = ev[1]
                if not isinstance(table, list):
                    raise _FormulaErr("#VALUE!（VLOOKUP 第二参数须为区域）")
                nrows = getattr(table, "nrows", len(table))
                ncols = getattr(table, "ncols", 1)
                c_idx = int(_excel_num(ev[2]))
                if c_idx < 1:
                    raise _FormulaErr("#VALUE!")
                if c_idx > ncols:
                    raise _FormulaErr("#REF!")
                approx = _excel_bool(ev[3]) if len(ev) > 3 else True
                if not approx:
                    for row_i in range(nrows):
                        key_v = table[row_i * ncols]
                        if _lookup_match_exact(key_v, lookup_val):
                            return table[row_i * ncols + c_idx - 1]
                    raise _FormulaErr("#N/A")
                t_rank = _excel_type_rank(lookup_val)
                best_row = None
                for row_i in range(nrows):
                    key_v = table[row_i * ncols]
                    if isinstance(key_v, _BlankCell) or _excel_type_rank(key_v) != t_rank:
                        continue
                    if _excel_cmp(key_v, lookup_val, "<="):
                        best_row = row_i
                    else:
                        break
                if best_row is None:
                    raise _FormulaErr("#N/A")
                return table[best_row * ncols + c_idx - 1]
            if fn == "SUMPRODUCT":
                arrs = [x if isinstance(x, list) else [x] for x in ev]
                if not arrs:
                    return 0
                non_scalar_lens = {len(a) for a in arrs if len(a) != 1}
                if len(non_scalar_lens) > 1:
                    raise _FormulaErr("#VALUE!（SUMPRODUCT 数组长度不一致）")
                n = max((len(a) for a in arrs), default=0)
                total = 0.0
                for idx in range(n):
                    prod = 1.0
                    for a in arrs:
                        v = a[idx] if len(a) > 1 else a[0]
                        if _is_err_val(v):
                            raise _FormulaErr(v)
                        prod *= _excel_num(v) if isinstance(v, (int, float, _BlankCell, bool)) else 0.0
                    total += prod
                return int(total) if total.is_integer() else total
            raise _FormulaErr(f"#NAME?（不支持的函数: {fn}）")
        raise _FormulaErr(f"#VALUE!（未知 AST 节点: {kind}）")

    def eval_formula_str(sh: str, f_body: str) -> Any:
        ast_root = parse_ast(sh, f_body)
        return eval_node(sh, ast_root)

    formula_total = 0
    errors: List[Dict[str, Any]] = []
    cached_checked = 0
    cached_mismatch = 0
    mismatch_samples: List[Dict[str, Any]] = []
    values_by_sheet: Dict[str, Dict[str, Any]] = {}

    for sname in wb_f.sheetnames:
        ws_f = wb_f[sname]
        ws_v = wb_v[sname]
        s_vals: Dict[str, Any] = {}
        for row in ws_f.iter_rows():
            for cell_obj in row:
                raw_v = cell_obj.value
                coord = cell_obj.coordinate
                if isinstance(raw_v, str) and raw_v.startswith("="):
                    formula_total += 1
                    try:
                        calc_v = get_cell(sname, coord)
                        if isinstance(calc_v, str) and calc_v in EXCEL_ERRORS:
                            raise _FormulaErr(calc_v)
                        s_vals[coord] = "" if isinstance(calc_v, _BlankCell) else calc_v
                        cached_v = ws_v[coord].value
                        if cached_v is not None:
                            cached_checked += 1
                            same = False
                            if isinstance(calc_v, bool) or isinstance(cached_v, bool):
                                same = bool(calc_v) == bool(cached_v)
                            elif isinstance(calc_v, (int, float)) and isinstance(cached_v, (int, float)):
                                same = abs(float(calc_v) - float(cached_v)) <= 1e-6
                            elif isinstance(calc_v, _BlankCell):
                                same = cached_v in ("", 0, None)
                            else:
                                same = str(calc_v) == str(cached_v)
                            if not same:
                                cached_mismatch += 1
                                if len(mismatch_samples) < 20:
                                    mismatch_samples.append({
                                        "sheet": sname, "coord": coord, "formula": raw_v,
                                        "calc": calc_v, "cached": cached_v,
                                    })
                    except _FormulaErr as exc:
                        errors.append({"sheet": sname, "coord": coord, "formula": raw_v, "error": str(exc)})
                    except Exception as exc:
                        errors.append({"sheet": sname, "coord": coord, "formula": raw_v, "error": f"{type(exc).__name__}: {exc}"})
        values_by_sheet[sname] = s_vals

    return {
        "file": str(xlsx_path),
        "sheets": wb_f.sheetnames,
        "formula_total": formula_total,
        "error_count": len(errors),
        "errors": errors[:50],
        "cached_checked": cached_checked,
        "cached_mismatch": cached_mismatch,
        "mismatch_samples": mismatch_samples,
        "values": values_by_sheet,
    }


def _collect_task_hashes(task_dir: Path) -> Dict[str, str]:
    """收集任务包内全部非隐藏、非 .ndjson 文件的 SHA-256。"""
    hashes: Dict[str, str] = {}
    for p in sorted(task_dir.rglob("*")):
        if p.is_file() and not p.name.startswith(".") and not p.name.endswith(".ndjson"):
            hashes[p.relative_to(task_dir).as_posix()] = hashlib.sha256(p.read_bytes()).hexdigest()
    return hashes


def freeze_task_pack(task_dir: Path, out_json: Path) -> int:
    """先执行 Gate 1~2 校验，通过后写入包含时间戳与 git 提交的冻结指纹文件。"""
    errs = validate_task_pack(task_dir)
    if errs:
        print(f"❌ 任务包未通过 Gate 1~2 校验，拒绝冻结（{len(errs)} 项问题）：")
        for e in errs:
            print("  -", e)
        return 1

    git_head = "unknown"
    try:
        proc = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=str(task_dir),
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=5,
        )
        if proc.returncode == 0 and proc.stdout.strip():
            git_head = proc.stdout.strip()
    except Exception:
        pass

    files = _collect_task_hashes(task_dir)
    tz_cn = timezone(timedelta(hours=8))
    payload = {
        "frozen_at": datetime.now(tz_cn).isoformat(timespec="seconds"),
        "git_head": git_head,
        "task_dir": task_dir.as_posix(),
        "file_count": len(files),
        "files": files,
    }
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"✅ 已冻结任务包指纹（{len(files)} 个文件，git={git_head}）：{out_json}")
    return 0


def verify_freeze_pack(task_dir: Path, freeze_json: Path) -> int:
    """严格比对任务包现行文件与冻结指纹记录，发现增删改立即报错返回 1。"""
    if not freeze_json.is_file():
        print(f"❌ 找不到冻结指纹文件：{freeze_json}")
        return 1
    data = json.loads(freeze_json.read_text(encoding="utf-8"))
    expected: Dict[str, str] = data["files"] if isinstance(data, dict) and "files" in data else data
    actual = _collect_task_hashes(task_dir)

    missing = sorted(k for k in expected if k not in actual)
    extra = sorted(k for k in actual if k not in expected)
    modified = sorted(k for k in expected if k in actual and expected[k] != actual[k])

    if missing or extra or modified:
        print(f"❌ 冻结指纹校验失败（{task_dir} vs {freeze_json}）：")
        for k in missing:
            print(f"  - 缺失文件: {k}")
        for k in extra:
            print(f"  - 新增未冻结文件: {k}")
        for k in modified:
            print(f"  - 文件内容已变更: {k}")
        return 1

    print(f"✅ 冻结指纹校验通过（{len(actual)} 个文件逐字节一致）：{task_dir}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="任务包与体验评测包四道闸门校验与生成工具")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_vt = sub.add_parser("validate-task", help="校验任务包（Gate 1 + Gate 2）")
    p_vt.add_argument("task_dir", type=Path)

    p_ve = sub.add_parser("validate-eval", help="校验体验评测包（Gate 3 + Gate 4）")
    p_ve.add_argument("eval_dir", type=Path)
    p_ve.add_argument("--rubric", type=Path, required=True)

    p_br = sub.add_parser("build-rubric", help="从 JSON 生成双工作表 rubric.xlsx 与 rubric.csv")
    p_br.add_argument("spec_json", type=Path)
    p_br.add_argument("out_dir", type=Path)

    p_be = sub.add_parser("build-eval", help="从 JSON 生成 5 表体验评测记录.xlsx 与 6 节体验评测报告.docx")
    p_be.add_argument("spec_json", type=Path)
    p_be.add_argument("out_dir", type=Path)

    p_rc = sub.add_parser("recalc-xlsx", help="独立 AST 重算 Excel 公式并校验缓存一致性")
    p_rc.add_argument("xlsx_path", type=Path)
    p_rc.add_argument("--json-out", type=Path, default=None)

    p_fr = sub.add_parser("freeze-task", help="校验任务包后写入冻结指纹 JSON")
    p_fr.add_argument("task_dir", type=Path)
    p_fr.add_argument("--out", type=Path, required=True)

    p_vf = sub.add_parser("verify-freeze", help="核验任务包当前文件与冻结指纹一致")
    p_vf.add_argument("task_dir", type=Path)
    p_vf.add_argument("--freeze", type=Path, required=True)

    args = parser.parse_args()
    if args.cmd == "validate-task":
        errs = validate_task_pack(args.task_dir)
        if errs:
            print(f"❌ 任务包校验未通过（共 {len(errs)} 项问题）：")
            for e in errs:
                print("  -", e)
            return 1
        print(f"✅ 任务包校验通过：{args.task_dir}")
        return 0

    if args.cmd == "validate-eval":
        errs = validate_eval_pack(args.eval_dir, args.rubric)
        if errs:
            print(f"❌ 体验评测包校验未通过（共 {len(errs)} 项问题）：")
            for e in errs:
                print("  -", e)
            return 1
        print(f"✅ 体验评测包校验通过：{args.eval_dir}")
        return 0

    if args.cmd == "build-rubric":
        build_rubric_from_spec(args.spec_json, args.out_dir)
        print(f"✅ 已生成双工作表 rubric.xlsx 与 rubric.csv：{args.out_dir}")
        return 0

    if args.cmd == "build-eval":
        build_eval_from_spec(args.spec_json, args.out_dir)
        print(f"✅ 已生成体验评测包文件：{args.out_dir}")
        return 0

    if args.cmd == "recalc-xlsx":
        res = recalc_workbook(args.xlsx_path)
        summary = {k: v for k, v in res.items() if k != "values"}
        print(
            f"recalc-xlsx: file={args.xlsx_path.name} formulas={res['formula_total']} "
            f"errors={res['error_count']} cached_checked={res['cached_checked']} "
            f"cached_mismatch={res['cached_mismatch']}"
        )
        for err in res["errors"][:10]:
            print("  ERR:", err)
        for mm in res["mismatch_samples"][:10]:
            print("  MISMATCH:", mm)
        if args.json_out is not None:
            args.json_out.parent.mkdir(parents=True, exist_ok=True)
            args.json_out.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return 1 if res["error_count"] > 0 else 0

    if args.cmd == "freeze-task":
        return freeze_task_pack(args.task_dir, args.out)

    if args.cmd == "verify-freeze":
        return verify_freeze_pack(args.task_dir, args.freeze)

    return 0


if __name__ == "__main__":
    sys.exit(main())
