#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用 OpenAI 兼容 Chat Completions 多轮 Agent 运行器（Function Calling）。

在严格限制的隔离工作目录内提供 list_dir / read_file / write_file / bash 四个工具，
自动阻止越界路径访问，并完整持久化每轮模型响应、工具调用与 token 统计。

环境变量：
- OPENAI_BASE_URL：OpenAI 兼容接口地址（如 https://api.openai.com/v1/chat/completions）
- OPENAI_API_KEY：接口密钥（优先环境变量，其次读取同仓库 .env）
- OPENAI_MODEL：默认模型名称（可由参数覆盖）
- EVAL_EXTRA_PATH：可选，追加到子进程 PATH 前缀的工具路径（如自定义 Python 解释器目录）
"""

import argparse
import getpass
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.error import HTTPError, URLError
from urllib.request import HTTPRedirectHandler, ProxyHandler, Request, build_opener

DEFAULT_BASE_URL = "https://api.openai.com/v1/chat/completions"
DEFAULT_MODEL = "gpt-4.1"

USER_PROMPT = (
    "请读取当前工作目录下的 `task/task.md` 以及 `reference_files/` 下的全部资料，"
    "独立完成任务要求，并将最终交付文件保存到 `output/` 目录。\n"
    "只允许使用当前目录内的文件，不要访问外部路径或联网。"
)
AGENT_SYSTEM_PROMPT = (
    "你在隔离工作目录中独立完成任务，可使用 list_dir、read_file、write_file、bash 工具"
    "（工作环境已提供 python3 及常用文档处理库）。所有文件读写和命令执行必须限制在当前工作目录内。"
)


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, hdrs, newurl):
        return None


TOOLS: List[Dict[str, Any]] = [
    {
        "type": "function",
        "function": {
            "name": "list_dir",
            "description": "列出工作目录或指定相对子目录下的文件与文件夹",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "相对工作目录的路径，默认为 '.'"}
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "read_file",
            "description": "按行读取工作目录内的文本文件（带行号）",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "相对工作目录的文件路径"},
                    "start_line": {"type": "integer", "description": "起始行号（1-based，默认 1）"},
                    "end_line": {"type": "integer", "description": "结束行号（默认 250）"},
                },
                "required": ["path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "write_file",
            "description": "在工作目录内创建或覆盖写入文本文件（可用于编写 Python 脚本）",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "相对工作目录的文件路径"},
                    "content": {"type": "string", "description": "完整文件内容"},
                },
                "required": ["path", "content"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "bash",
            "description": "在当前工作目录执行 shell 命令",
            "parameters": {
                "type": "object",
                "properties": {
                    "command": {"type": "string", "description": "要执行的 shell 命令"}
                },
                "required": ["command"],
            },
        },
    },
]


def _resolve_safe(workdir: Path, rel: str) -> Path:
    """将相对路径解析为绝对路径，并禁止逃逸出 workdir。"""
    base = workdir.resolve()
    target = (base / (rel or ".")).resolve()
    if target != base and base not in target.parents:
        raise ValueError(f"越界路径禁止访问: {rel}")
    return target


def _load_env_var(name: str, env_file: Optional[Path] = None) -> Optional[str]:
    val = os.environ.get(name)
    if val and val.strip():
        return val.strip()
    candidate = env_file or (Path(__file__).resolve().parent.parent / ".env")
    if candidate.is_file():
        for line in candidate.read_text(encoding="utf-8").splitlines():
            s = line.strip()
            if not s or s.startswith("#") or "=" not in s:
                continue
            k, v = s.split("=", 1)
            if k.strip() == name:
                cleaned = v.strip().strip("\"'")
                if cleaned:
                    return cleaned
    return None


def exec_tool(workdir: Path, name: str, args: Dict[str, Any], env: Dict[str, str]) -> str:
    """在隔离目录内执行单个工具调用。"""
    try:
        if name == "list_dir":
            p = _resolve_safe(workdir, args.get("path", "."))
            if not p.exists():
                return f"目录不存在: {args.get('path', '.')}"
            items = sorted(p.iterdir(), key=lambda x: (not x.is_dir(), x.name))
            return "\n".join(
                f"{'[DIR] ' if x.is_dir() else '[FILE]'} {x.name}"
                + (f" ({x.stat().st_size} bytes)" if x.is_file() else "")
                for x in items
            ) or "(空目录)"

        if name == "read_file":
            p = _resolve_safe(workdir, args["path"])
            if not p.is_file():
                return f"文件不存在: {args['path']}"
            lines = p.read_text(encoding="utf-8", errors="replace").splitlines()
            s = max(1, int(args.get("start_line") or 1))
            e = min(len(lines), int(args.get("end_line") or (s + 249)))
            out = [f"{i}: {lines[i - 1]}" for i in range(s, e + 1)]
            if e < len(lines):
                out.append(f"...(共 {len(lines)} 行，当前仅显示到第 {e} 行)")
            return "\n".join(out)

        if name == "write_file":
            p = _resolve_safe(workdir, args["path"])
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(args["content"], encoding="utf-8")
            return f"已写入 {args['path']}（{len(args['content'])} 字符）"

        if name == "bash":
            cmd = args["command"]
            shell = "/bin/bash" if Path("/bin/bash").exists() else "/bin/sh"
            proc = subprocess.run(
                [shell, "-lc", cmd],
                cwd=str(workdir),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=120,
            )
            out = proc.stdout or ""
            if len(out) > 16000:
                out = out[:8000] + "\n...[中间输出已截断]...\n" + out[-7000:]
            return f"[exit_code={proc.returncode}]\n{out}"

        return f"未知工具: {name}"
    except Exception as exc:
        return f"[tool_error] {type(exc).__name__}: {exc}"


def call_chat_completions(
    base_url: str,
    api_key: str,
    messages: List[Dict[str, Any]],
    opener: Any,
    model: str,
    extra_body: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """调用 OpenAI 兼容 Chat Completions 接口，内置 3 次指数退避重试。"""
    body: Dict[str, Any] = {
        "model": model,
        "messages": messages,
        "tools": TOOLS,
        "tool_choice": "auto",
        "max_tokens": 16384,
        "stream": False,
    }
    if extra_body:
        body.update(extra_body)

    for attempt in range(3):
        req = Request(
            base_url,
            data=json.dumps(body).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
        )
        try:
            with opener.open(req, timeout=180) as resp:
                return json.load(resp)
        except HTTPError as exc:
            if (exc.code in (408, 429) or 500 <= exc.code < 600) and attempt < 2:
                time.sleep(2 ** attempt)
                continue
            raise
        except (URLError, TimeoutError, OSError):
            if attempt < 2:
                time.sleep(2 ** attempt)
                continue
            raise
    raise RuntimeError("API 请求重试耗尽")


def run_agent_session(
    workdir: Path,
    events_jsonl_path: Path,
    final_reply_path: Path,
    model: Optional[str] = None,
    base_url: Optional[str] = None,
    api_key_env: str = "OPENAI_API_KEY",
    system_prompt: str = AGENT_SYSTEM_PROMPT,
    user_prompt: str = USER_PROMPT,
    max_turns: int = 80,
    timeout_sec: int = 1800,
    extra_body: Optional[Dict[str, Any]] = None,
    use_proxy: bool = False,
    log_prefix: str = "[Agent] ",
) -> Dict[str, Any]:
    """运行中性多轮 Function Calling 会话，无论成功或异常均落盘完整轨迹。"""
    resolved_url = base_url or _load_env_var("OPENAI_BASE_URL") or DEFAULT_BASE_URL
    resolved_model = model or _load_env_var("OPENAI_MODEL") or DEFAULT_MODEL
    api_key = _load_env_var(api_key_env)
    if not api_key and sys.stdin.isatty():
        api_key = getpass.getpass(f"{api_key_env} (hidden): ").strip()
    if not api_key:
        raise SystemExit(f"未提供密钥环境变量：{api_key_env}")

    env = os.environ.copy()
    extra_path = os.environ.get("EVAL_EXTRA_PATH", "").strip()
    if extra_path:
        env["PATH"] = f"{extra_path}{os.pathsep}" + env.get("PATH", "")

    opener = build_opener(_NoRedirect) if use_proxy else build_opener(ProxyHandler({}), _NoRedirect)

    messages: List[Dict[str, Any]] = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    events: List[Dict[str, Any]] = []
    total_prompt = 0
    total_completion = 0
    total_cached = 0
    resp_model = resolved_model
    t0 = time.time()
    final_text = ""
    returncode = 0
    err_msg: Optional[str] = None

    try:
        for turn in range(1, max_turns + 1):
            if time.time() - t0 > timeout_sec:
                returncode = 124
                err_msg = f"达到总超时上限 {timeout_sec}s"
                final_text = f"({err_msg})"
                break
            t_turn = time.time()
            try:
                resp = call_chat_completions(
                    base_url=resolved_url,
                    api_key=api_key,
                    messages=messages,
                    opener=opener,
                    model=resolved_model,
                    extra_body=extra_body,
                )
            except HTTPError as exc:
                err_body = exc.read().decode("utf-8", errors="replace")[:500]
                events.append({"turn": turn, "type": "http_error", "code": exc.code, "detail": err_body})
                returncode = 1
                err_msg = f"Turn {turn} HTTP {exc.code}: {err_body}"
                break
            except (URLError, TimeoutError, OSError, RuntimeError) as exc:
                events.append({"turn": turn, "type": "url_error", "detail": str(exc)})
                returncode = 1
                err_msg = f"Turn {turn} NetworkError: {exc}"
                break

            usage = resp.get("usage") or {}
            total_prompt += int(usage.get("prompt_tokens") or 0)
            total_completion += int(usage.get("completion_tokens") or 0)
            pt_details = usage.get("prompt_tokens_details") or {}
            total_cached += int(pt_details.get("cached_tokens") or 0)
            resp_model = resp.get("model") or resp_model

            choice = resp["choices"][0]
            msg = choice["message"]
            finish_reason = choice.get("finish_reason")
            tool_calls = msg.get("tool_calls") or []

            assistant_record: Dict[str, Any] = {"role": "assistant", "content": msg.get("content")}
            if tool_calls:
                assistant_record["tool_calls"] = tool_calls
            if "reasoning_content" in msg and msg["reasoning_content"] is not None:
                assistant_record["reasoning_content"] = msg["reasoning_content"]
            messages.append(assistant_record)

            events.append(
                {
                    "turn": turn,
                    "type": "assistant",
                    "model": resp_model,
                    "finish_reason": finish_reason,
                    "seconds": round(time.time() - t_turn, 2),
                    "usage": usage,
                    "content": msg.get("content"),
                    "tool_calls": tool_calls,
                }
            )
            print(
                f"{log_prefix}[Turn {turn}] finish={finish_reason} tools={len(tool_calls)} "
                f"tokens={usage.get('prompt_tokens')}/{usage.get('completion_tokens')} "
                f"({round(time.time() - t_turn, 1)}s)",
                flush=True,
            )

            if not tool_calls:
                final_text = msg.get("content") or ""
                break

            for tc in tool_calls:
                fn = tc["function"]["name"]
                raw_args = tc["function"].get("arguments") or "{}"
                try:
                    args = json.loads(raw_args)
                except json.JSONDecodeError:
                    args = {}
                out = exec_tool(workdir, fn, args, env)
                messages.append({"role": "tool", "tool_call_id": tc["id"], "content": out})
                events.append(
                    {
                        "turn": turn,
                        "type": "tool_result",
                        "tool": fn,
                        "tool_call_id": tc["id"],
                        "arguments": args,
                        "output_preview": out[:1500],
                    }
                )
        else:
            final_text = f"(达到最大轮数 {max_turns} 轮提前结束)"
            returncode = 1
            err_msg = final_text
    finally:
        events_jsonl_path.parent.mkdir(parents=True, exist_ok=True)
        with events_jsonl_path.open("w", encoding="utf-8") as fp:
            for ev in events:
                fp.write(json.dumps(ev, ensure_ascii=False) + "\n")
        final_reply_path.parent.mkdir(parents=True, exist_ok=True)
        final_reply_path.write_text(final_text, encoding="utf-8")

    elapsed = round(time.time() - t0, 2)
    turns = len([e for e in events if e.get("type") == "assistant"])
    print(
        f"{log_prefix}DONE: rc={returncode} elapsed={elapsed}s turns={turns} "
        f"prompt_tokens={total_prompt} completion_tokens={total_completion}",
        flush=True,
    )
    return {
        "returncode": returncode,
        "error": err_msg,
        "elapsed_seconds": elapsed,
        "turns": turns,
        "model": resp_model,
        "token_usage": {
            "input_tokens": total_prompt,
            "cached_input_tokens": total_cached,
            "output_tokens": total_completion,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="在隔离工作目录运行 OpenAI 兼容多轮 Agent")
    parser.add_argument("workdir", type=Path, help="隔离工作目录（需包含 task/task.md 与 reference_files/）")
    parser.add_argument("--model", default=None, help="模型名称（默认读取 OPENAI_MODEL）")
    parser.add_argument("--base-url", default=None, help="接口地址（默认读取 OPENAI_BASE_URL）")
    parser.add_argument("--api-key-env", default="OPENAI_API_KEY", help="读取密钥的环境变量名")
    parser.add_argument("--max-turns", type=int, default=80, help="最大交互轮数")
    parser.add_argument("--timeout", type=int, default=1800, help="总超时秒数")
    args = parser.parse_args()

    res = run_agent_session(
        workdir=args.workdir,
        events_jsonl_path=args.workdir.parent / "agent_events.jsonl",
        final_reply_path=args.workdir.parent / "agent_reply.md",
        model=args.model,
        base_url=args.base_url,
        api_key_env=args.api_key_env,
        max_turns=args.max_turns,
        timeout_sec=args.timeout,
    )
    return int(res["returncode"])


if __name__ == "__main__":
    sys.exit(main())
