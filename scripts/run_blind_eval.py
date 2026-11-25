#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一公平盲测运行器：
1. 盲测前强校验任务包冻结指纹（pack_gate.verify_freeze_pack）；
2. 在 git 仓库之外创建按「任务包/批次/模型」隔离的运行目录，仅投放 task/task.md 与 reference_files/*；
3. 使用完全一致的中性提示词分别调度 OpenAI 兼容 Function Calling Agent 或 Codex CLI；
4. 无论成功或超时，通过 finally 将全套产物、事件轨迹、输入清单与 Excel AST 公式/越界审计落盘。
"""

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
from typing import Any, Dict, List, Optional

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import openai_compat_agent  # noqa: E402
import pack_gate  # noqa: E402

DEFAULT_STAGE_ROOT = Path(
    os.environ.get("EVAL_STAGE_ROOT") or (Path(tempfile.gettempdir()) / "agent_eval_stage")
)
AUDIT_KEYWORDS = ("rubric", "ground_truth", "curl", "wget", "pip install")


def _now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def _prepare_stage_dir(task_pack: Path, freeze_data: Dict[str, Any], stage_dir: Path) -> Dict[str, Any]:
    """在仓库外构建隔离运行目录，并逐文件核对输入 SHA-256 与冻结指纹一致。"""
    if stage_dir.exists():
        shutil.rmtree(stage_dir)
    (stage_dir / "task").mkdir(parents=True, exist_ok=True)
    (stage_dir / "reference_files").mkdir(parents=True, exist_ok=True)
    (stage_dir / "output").mkdir(parents=True, exist_ok=True)

    expected_files: Dict[str, str] = (
        freeze_data["files"] if isinstance(freeze_data, dict) and "files" in freeze_data else freeze_data
    )

    manifest: Dict[str, Any] = {"staged_at": _now_iso(), "stage_dir": str(stage_dir), "files": {}}

    src_task = task_pack / "task" / "task.md"
    dst_task = stage_dir / "task" / "task.md"
    shutil.copy2(src_task, dst_task)
    task_hash = _sha256(dst_task)
    exp_task = expected_files.get("task/task.md")
    if exp_task and task_hash != exp_task:
        raise RuntimeError(f"隔离目录 task/task.md 指纹与冻结版不一致：{task_hash} != {exp_task}")
    manifest["files"]["task/task.md"] = {
        "sha256": task_hash,
        "bytes": dst_task.stat().st_size,
        "matches_freeze": task_hash == exp_task,
    }

    for src in sorted((task_pack / "reference_files").iterdir()):
        if not src.is_file() or src.name.startswith("."):
            continue
        dst = stage_dir / "reference_files" / src.name
        shutil.copy2(src, dst)
        rel = f"reference_files/{src.name}"
        h = _sha256(dst)
        exp_h = expected_files.get(rel)
        if exp_h and h != exp_h:
            raise RuntimeError(f"隔离目录 {rel} 指纹与冻结版不一致：{h} != {exp_h}")
        manifest["files"][rel] = {"sha256": h, "bytes": dst.stat().st_size, "matches_freeze": h == exp_h}

    return manifest


def _parse_codex_usage(events_path: Path) -> Dict[str, int]:
    """从 Codex CLI --json 事件流提取累计 token 用量。"""
    inp, cached, out = 0, 0, 0
    if not events_path.is_file():
        return {"input_tokens": 0, "cached_input_tokens": 0, "output_tokens": 0}
    for line in events_path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            ev = json.loads(line)
        except json.JSONDecodeError:
            continue
        usage = ev.get("usage")
        if isinstance(usage, dict):
            inp += int(usage.get("input_tokens") or 0)
            cached += int(usage.get("cached_input_tokens") or 0)
            out += int(usage.get("output_tokens") or 0)
    return {"input_tokens": inp, "cached_input_tokens": cached, "output_tokens": out}


def _scan_events_for_violations(events_path: Path) -> List[Dict[str, Any]]:
    """扫描事件流里的命令/参数字段，检查是否出现越界路径或联网关键词。"""
    hits: List[Dict[str, Any]] = []
    if not events_path.is_file():
        return hits
    for idx, line in enumerate(events_path.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
        s = line.strip()
        if not s:
            continue
        targets: List[str] = []
        try:
            ev = json.loads(s)
            if ev.get("type") in ("item.started", "item.completed") and isinstance(ev.get("item"), dict):
                item = ev["item"]
                if item.get("command"):
                    targets.append(str(item["command"]))
            elif ev.get("type") == "tool_result":
                targets.append(json.dumps(ev.get("arguments") or {}, ensure_ascii=False))
            elif ev.get("type") == "assistant" and ev.get("tool_calls"):
                for tc in ev["tool_calls"]:
                    fn = (tc or {}).get("function") or {}
                    targets.append(str(fn.get("arguments") or ""))
        except json.JSONDecodeError:
            targets.append(s)

        for t in targets:
            for kw in AUDIT_KEYWORDS:
                if kw.lower() in t.lower():
                    hits.append({"line": idx, "keyword": kw, "snippet": t[:240]})
    return hits


def _audit_workspace(run_out_dir: Path) -> Dict[str, Any]:
    """扫描运行事件越界/联网情况，并对 output/*.xlsx 执行 AST 公式重算。"""
    events_path = run_out_dir / "运行事件.jsonl"
    violations = _scan_events_for_violations(events_path)

    ws_out = run_out_dir / "workspace" / "output"
    output_files = []
    xlsx_audits = []
    if ws_out.is_dir():
        for f in sorted(ws_out.iterdir()):
            if f.is_file() and not f.name.startswith("."):
                output_files.append({"name": f.name, "bytes": f.stat().st_size, "sha256": _sha256(f)})
                if f.suffix.lower() == ".xlsx":
                    try:
                        rec = pack_gate.recalc_workbook(f)
                        xlsx_audits.append({k: v for k, v in rec.items() if k != "values"})
                    except Exception as exc:
                        xlsx_audits.append(
                            {"file": f.name, "error_count": 1, "errors": [f"{type(exc).__name__}: {exc}"]}
                        )

    return {
        "audited_at": _now_iso(),
        "violation_count": len(violations),
        "violations": violations,
        "output_files": output_files,
        "xlsx_recalc": xlsx_audits,
    }


def _copy_workspace(stage_dir: Path, dst_ws: Path) -> None:
    if dst_ws.exists():
        shutil.rmtree(dst_ws)
    shutil.copytree(
        stage_dir,
        dst_ws,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store", ".git"),
    )


def run_one_model(
    cfg: Dict[str, Any],
    task_pack: Path,
    freeze_data: Dict[str, Any],
    run_id: str,
    runs_dir: Path,
    stage_root: Path,
    timeout_sec: int,
    max_turns: int,
) -> Dict[str, Any]:
    label = str(cfg["label"])
    engine = str(cfg["engine"])
    model = str(cfg["model"])

    stage_dir = stage_root / task_pack.name / run_id / label
    run_out_dir = runs_dir / run_id / label
    if run_out_dir.exists():
        shutil.rmtree(run_out_dir)
    run_out_dir.mkdir(parents=True, exist_ok=True)

    manifest = _prepare_stage_dir(task_pack, freeze_data, stage_dir)
    (run_out_dir / "input_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )

    prompt_record = f"=== USER_PROMPT ===\n{openai_compat_agent.USER_PROMPT}\n"
    if engine == "openai_compat":
        prompt_record += f"\n=== SYSTEM_PROMPT ===\n{openai_compat_agent.AGENT_SYSTEM_PROMPT}\n"
    elif engine == "codex":
        prompt_record += (
            "\n=== CODEX_CLI_NOTE ===\n"
            "使用 Codex CLI 隔离会话：--ephemeral --ignore-rules -s workspace-write\n"
        )
    (run_out_dir / "prompt_used.txt").write_text(prompt_record, encoding="utf-8")

    events_path = run_out_dir / "运行事件.jsonl"
    reply_path = run_out_dir / "最终回复.md"
    stderr_path = run_out_dir / "stderr.log"

    start_iso = _now_iso()
    t0 = time.time()
    rc = -1
    token_usage: Dict[str, int] = {"input_tokens": 0, "cached_input_tokens": 0, "output_tokens": 0}
    resp_model = model
    err_detail: Optional[str] = None

    try:
        if engine == "codex":
            codex_bin = str(cfg.get("codex_bin") or os.environ.get("CODEX_BIN") or shutil.which("codex") or "codex")
            env = os.environ.copy()
            extra_path = os.environ.get("EVAL_EXTRA_PATH", "").strip()
            if extra_path:
                env["PATH"] = f"{extra_path}{os.pathsep}" + env.get("PATH", "")
            reasoning = str(cfg.get("reasoning_effort") or "medium")
            cmd = [
                codex_bin,
                "exec",
                "-m",
                model,
                "-c",
                f'model_reasoning_effort="{reasoning}"',
                "-s",
                "workspace-write",
                "--skip-git-repo-check",
                "--ephemeral",
                "--ignore-rules",
                "-C",
                str(stage_dir),
                "--json",
                "-o",
                str(reply_path),
                "-",
            ]
            print(f"[{label}] 启动 Codex CLI ({model}) ...", flush=True)
            with events_path.open("wb") as fout, stderr_path.open("wb") as ferr:
                try:
                    proc = subprocess.run(
                        cmd,
                        cwd=str(stage_dir),
                        env=env,
                        input=openai_compat_agent.USER_PROMPT.encode("utf-8"),
                        stdout=fout,
                        stderr=ferr,
                        timeout=timeout_sec,
                    )
                    rc = proc.returncode
                except subprocess.TimeoutExpired:
                    rc = 124
                    err_detail = f"Codex CLI 超时（>{timeout_sec}s）"
            token_usage = _parse_codex_usage(events_path)
            if not reply_path.is_file():
                reply_path.write_text(f"({err_detail or f'退出码 {rc}'})", encoding="utf-8")

        elif engine == "openai_compat":
            print(f"[{label}] 启动 OpenAI 兼容 Agent ({model}) ...", flush=True)
            stderr_path.write_text("", encoding="utf-8")
            res = openai_compat_agent.run_agent_session(
                workdir=stage_dir,
                events_jsonl_path=events_path,
                final_reply_path=reply_path,
                model=model,
                base_url=cfg.get("base_url"),
                api_key_env=str(cfg.get("api_key_env") or "OPENAI_API_KEY"),
                system_prompt=openai_compat_agent.AGENT_SYSTEM_PROMPT,
                user_prompt=openai_compat_agent.USER_PROMPT,
                max_turns=max_turns,
                timeout_sec=timeout_sec,
                extra_body=cfg.get("extra_body"),
                use_proxy=bool(cfg.get("use_proxy", False)),
                log_prefix=f"[{label}] ",
            )
            rc = int(res["returncode"])
            resp_model = str(res["model"])
            token_usage = res["token_usage"]
            err_detail = res.get("error")
        else:
            raise ValueError(f"不支持的运行通道: {engine}（仅支持 openai_compat / codex）")

    except Exception as exc:
        err_detail = f"{type(exc).__name__}: {exc}"
        with stderr_path.open("a", encoding="utf-8") as ferr:
            ferr.write(err_detail + "\n")
    finally:
        end_iso = _now_iso()
        elapsed = round(time.time() - t0, 2)
        _copy_workspace(stage_dir, run_out_dir / "workspace")
        audit = _audit_workspace(run_out_dir)
        (run_out_dir / "audit_summary.json").write_text(
            json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        meta = {
            "label": label,
            "run_id": run_id,
            "engine": engine,
            "requested_model": model,
            "response_model": resp_model,
            "start_time": start_iso,
            "end_time": end_iso,
            "elapsed_seconds": elapsed,
            "returncode": rc,
            "error": err_detail,
            "token_usage": token_usage,
            "stage_dir": str(stage_dir),
            "output_dir": str(run_out_dir),
            "output_files": audit["output_files"],
        }
        (run_out_dir / "meta.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(
            f"[{label}] 落盘完成 -> {run_out_dir} (rc={rc}, {elapsed}s, "
            f"files={len(audit['output_files'])}, violations={audit['violation_count']})",
            flush=True,
        )

    return meta


def main() -> int:
    parser = argparse.ArgumentParser(description="统一公平盲测运行器（先校验冻结指纹，再隔离盲答并持久化落盘）")
    parser.add_argument("--task-pack", type=Path, required=True, help="冻结后的任务包目录")
    parser.add_argument("--freeze", type=Path, required=True, help="冻结指纹 JSON 路径")
    parser.add_argument("--runs-dir", type=Path, required=True, help="盲测运行持久化根目录")
    parser.add_argument("--models-config", type=Path, required=True, help="模型矩阵 JSON 配置文件路径")
    parser.add_argument("--run-id", default="run_01", help="本轮盲测批次标识（如 run_01）")
    parser.add_argument("--models", default=None, help="可选：逗号分隔的要运行模型标签（默认全部）")
    parser.add_argument("--stage-root", type=Path, default=DEFAULT_STAGE_ROOT, help="隔离暂存根目录")
    parser.add_argument("--timeout", type=int, default=1800, help="单模型总超时秒数（默认 1800s）")
    parser.add_argument("--max-turns", type=int, default=80, help="OpenAI 兼容通道最大工具交互轮数")
    parser.add_argument("--sequential", action="store_true", help="串行执行（默认并行执行所选模型）")
    args = parser.parse_args()

    rc = pack_gate.verify_freeze_pack(args.task_pack, args.freeze)
    if rc != 0:
        print("❌ 冻结指纹校验未通过，严禁启动盲测！", file=sys.stderr)
        return rc

    freeze_data = json.loads(args.freeze.read_text(encoding="utf-8"))
    all_models: List[Dict[str, Any]] = json.loads(args.models_config.read_text(encoding="utf-8"))
    if args.models:
        wanted = {m.strip() for m in args.models.split(",") if m.strip()}
        selected = [cfg for cfg in all_models if cfg.get("label") in wanted]
    else:
        selected = list(all_models)

    if not selected:
        raise SystemExit("未匹配到任何有效模型配置")

    results: List[Dict[str, Any]] = []
    if args.sequential or len(selected) == 1:
        for cfg in selected:
            results.append(
                run_one_model(
                    cfg,
                    args.task_pack,
                    freeze_data,
                    args.run_id,
                    args.runs_dir,
                    args.stage_root,
                    args.timeout,
                    args.max_turns,
                )
            )
    else:
        with ThreadPoolExecutor(max_workers=len(selected)) as pool:
            futs = [
                pool.submit(
                    run_one_model,
                    cfg,
                    args.task_pack,
                    freeze_data,
                    args.run_id,
                    args.runs_dir,
                    args.stage_root,
                    args.timeout,
                    args.max_turns,
                )
                for cfg in selected
            ]
            for fut in as_completed(futs):
                results.append(fut.result())

    failed = [r for r in results if r["returncode"] != 0]
    if failed:
        print(f"⚠️ {len(failed)}/{len(results)} 个模型非零退出：", [f"{r['label']}(rc={r['returncode']})" for r in failed])
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
