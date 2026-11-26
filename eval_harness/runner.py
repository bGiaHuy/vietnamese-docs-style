"""Isolated 3-Arm benchmark runner, telemetry parser, and contamination scanner.

Strictly isolated from repo scripts/ and evals/ modules (enforced by AST CI check).
"""

from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from eval_harness.grader import grade_case
from eval_harness.stats import compare_arms, summarize_experiment

ARM_NO_SKILL = "arm0_no_skill"
ARM_STANDARDS_ONLY = "arm1_standards_only"
ARM_FULL_SKILL = "arm2_full_skill"
ALL_ARMS = (ARM_NO_SKILL, ARM_STANDARDS_ONLY, ARM_FULL_SKILL)
MAX_VAL_MILESTONES = 3

_CONTAMINATION_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("vndocs_token", re.compile(r"\bvndocs\b", re.IGNORECASE)),
    (
        "scripts_cli_path",
        re.compile(
            r"scripts[/\\](?:cli|layout_edit|map_docx|validate_docx|render_docx|diff_layout|auto_paginate|build_docx|generate_docx)\.py",
            re.IGNORECASE,
        ),
    ),
    ("repo_name", re.compile(r"vietnamese-docs-style", re.IGNORECASE)),
    ("skill_md", re.compile(r"SKILL\.md")),
    ("eval_harness", re.compile(r"\beval_harness\b")),
]

_DEFAULT_STANDARDS_RELPATH = Path("evals") / "prompts" / "arm1_standards_only.md"


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def _sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _safe_git_head(repo_root: Path) -> str | None:
    try:
        out = subprocess.check_output(
            ["git", "rev-parse", "HEAD"],
            cwd=str(repo_root),
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=5,
        ).strip()
        return out or None
    except Exception:
        return None


def _copy_skill_bundle(repo_root: Path, dst_skill_dir: Path) -> None:
    dst_skill_dir.mkdir(parents=True, exist_ok=True)
    skill_md = repo_root / "SKILL.md"
    if not skill_md.is_file():
        raise FileNotFoundError(f"Missing frozen SKILL.md at {skill_md}")
    shutil.copy2(skill_md, dst_skill_dir / "SKILL.md")

    ignore_fn = shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo", ".pytest_cache")
    for sub in ("references", "scripts", "assets"):
        src_sub = repo_root / sub
        if src_sub.is_dir():
            shutil.copytree(src_sub, dst_skill_dir / sub, dirs_exist_ok=True, ignore=ignore_fn)


def prepare_arm_workspace(
    case: dict[str, Any],
    corpus_dir: str | Path,
    work_dir: str | Path,
    arm: str,
    *,
    standards_guide_path: str | Path | None = None,
    repo_root: str | Path | None = None,
    clean: bool = True,
) -> dict[str, Any]:
    """Prepare an isolated per-(case, arm) temporary workspace with strict arm isolation."""
    if arm not in ALL_ARMS:
        raise ValueError(f"Unknown arm {arm!r}; expected one of {ALL_ARMS}")

    cid = case["case_id"]
    cdir = Path(corpus_dir)
    wdir = Path(work_dir)
    root = Path(repo_root) if repo_root is not None else _repo_root()

    if clean and wdir.exists():
        shutil.rmtree(wdir)
    wdir.mkdir(parents=True, exist_ok=True)

    src_origin = cdir / case["source_docx_relpath"]
    if not src_origin.is_file():
        raise FileNotFoundError(f"Source DOCX missing for {cid}: {src_origin}")
    src_local = wdir / "src.docx"
    shutil.copy2(src_origin, src_local)

    att_paths: list[str] = []
    for att in case.get("attachments") or []:
        att_origin = cdir / att["relpath"]
        if not att_origin.is_file():
            raise FileNotFoundError(f"Attachment missing for {cid}: {att_origin}")
        att_local = wdir / Path(att["relpath"]).name
        shutil.copy2(att_origin, att_local)
        att_paths.append(str(att_local))

    standards_local: Path | None = None
    skill_dir_local: Path | None = None
    if arm == ARM_STANDARDS_ONLY:
        guide_src = Path(standards_guide_path) if standards_guide_path is not None else (root / _DEFAULT_STANDARDS_RELPATH)
        if not guide_src.is_file():
            raise FileNotFoundError(f"Standards guide missing: {guide_src}")
        standards_local = wdir / "STANDARDS_GUIDE.md"
        shutil.copy2(guide_src, standards_local)
    elif arm == ARM_FULL_SKILL:
        skill_dir_local = wdir / "skill"
        _copy_skill_bundle(root, skill_dir_local)

    return {
        "case_id": cid,
        "arm": arm,
        "work_dir": str(wdir),
        "src_docx": str(src_local),
        "dst_docx": str(wdir / f"{cid}.docx"),
        "result_json": str(wdir / f"{cid}.result.json"),
        "attachment_paths": att_paths,
        "standards_guide": str(standards_local) if standards_local else None,
        "skill_dir": str(skill_dir_local) if skill_dir_local else None,
    }


def build_arm_prompt(case: dict[str, Any], ws_info: dict[str, Any]) -> str:
    """Build an isolated natural-language task prompt tailored strictly to the given arm."""
    cid = case["case_id"]
    arm = ws_info["arm"]
    att_names = [Path(p).name for p in ws_info.get("attachment_paths", [])]
    att_line = (
        f"- Tệp đính kèm tham chiếu trong cùng thư mục: {', '.join(att_names)}\n"
        if att_names
        else ""
    )

    if arm == ARM_NO_SKILL:
        env_block = (
            "Môi trường làm việc cô lập chỉ chứa tệp đầu vào `src.docx`"
            f"{' và tệp đính kèm' if att_names else ''}.\n"
            "Không đọc, tìm kiếm hoặc sử dụng bất kỳ skill, plugin hay hướng dẫn nào nằm ngoài thư mục làm việc hiện tại; "
            "chỉ sử dụng Python (`python-docx`, `lxml`, `pymupdf`) và LibreOffice nếu có sẵn trên hệ thống."
        )
    elif arm == ARM_STANDARDS_ONLY:
        env_block = (
            "Trước khi thực hiện, hãy đọc tệp hướng dẫn quy chuẩn `STANDARDS_GUIDE.md` nằm ngay trong thư mục làm việc hiện tại.\n"
            "Không đọc, tìm kiếm hoặc sử dụng bất kỳ skill, plugin hay hướng dẫn nào nằm ngoài thư mục làm việc hiện tại; "
            "triển khai chỉnh sửa trực tiếp bằng Python (`python-docx`, `lxml`, `pymupdf`) "
            "hoặc LibreOffice theo đúng các quy chuẩn và ranh giới an toàn trong `STANDARDS_GUIDE.md`."
        )
    elif arm == ARM_FULL_SKILL:
        env_block = (
            "Trong thư mục làm việc có bộ công cụ và hướng dẫn chuyên biệt đặt tại `./skill/`.\n"
            "Hãy đọc `./skill/SKILL.md` và các tài liệu liên quan trong `./skill/references/` theo nhu cầu; "
            "khi xử lý `.docx`, ưu tiên dùng CLI trong `./skill/scripts/cli.py` (ví dụ `python skill/scripts/cli.py ...`) kết hợp kiểm tra trực quan nếu cần.\n"
            "Không truy cập thư mục đánh giá hoặc tệp bên ngoài thư mục làm việc hiện tại."
        )
    else:
        raise ValueError(f"Unsupported arm: {arm!r}")

    lo_note = (
        "Lưu ý môi trường Windows: nếu chạy LibreOffice headless để xuất PDF kiểm tra, "
        "luôn truyền `-env:UserInstallation=file:///...` trỏ vào một thư mục con tạm trong thư mục làm việc hiện tại để tránh xung đột tiến trình song song."
    )

    return (
        f"Bạn đang xử lý một tác vụ tài liệu Word tiếng Việt tại thư mục làm việc hiện tại.\n\n"
        f"{env_block}\n"
        f"{lo_note}\n\n"
        f"Thông tin đầu vào:\n"
        f"- Tệp Word nguồn: `src.docx` (giữ nguyên không ghi đè trực tiếp lên `src.docx`)\n"
        f"{att_line}"
        f"- Yêu cầu người dùng: {case['prompt_vi']}\n\n"
        f"Yêu cầu bàn giao bắt buộc tại thư mục gốc làm việc hiện tại:\n"
        f"1. Nếu yêu cầu hợp lệ và khả thi an toàn:\n"
        f"   - Xuất tệp kết quả thành `{cid}.docx` (giữ nguyên toàn bộ phần ngoài mục tiêu).\n"
        f"   - Ghi tệp `{cid}.result.json` dạng JSON UTF-8 hợp lệ: "
        f"`{{\"case_id\": \"{cid}\", \"status\": \"edited\", \"pagination_ripple\": true|false, \"reason\": \"tóm tắt ngắn gọn\"}}`.\n"
        f"2. Nếu yêu cầu vi phạm quy chuẩn bắt buộc, có chữ ký số xác thực, không an toàn hoặc quá mơ hồ không thể phân biệt:\n"
        f"   - **Không** tạo tệp `{cid}.docx` (hoặc giữ nguyên không chỉnh sửa).\n"
        f"   - Ghi tệp `{cid}.result.json` với `\"status\": \"refused\"` và giải thích rõ lý do trong `\"reason\"`.\n"
    )


def _accumulate_usage(obj: Any, acc: dict[str, int]) -> None:
    if isinstance(obj, dict):
        if "usage" in obj and isinstance(obj["usage"], dict):
            u = obj["usage"]
            if "input_tokens" in u or "output_tokens" in u:
                acc["input_tokens"] += int(u.get("input_tokens") or 0)
                acc["cached_input_tokens"] += int(u.get("cached_input_tokens") or 0)
                acc["output_tokens"] += int(u.get("output_tokens") or 0)
                return
        for v in obj.values():
            _accumulate_usage(v, acc)
    elif isinstance(obj, list):
        for item in obj:
            _accumulate_usage(item, acc)


def parse_codex_jsonl(
    jsonl_path: str | Path,
    *,
    elapsed_seconds: float | None = None,
) -> dict[str, Any]:
    """Parse `codex exec --json` JSONL transcript into standardized telemetry."""
    path = Path(jsonl_path)
    if not path.is_file():
        return {"missing": True, "path": str(path)}

    text = path.read_text(encoding="utf-8", errors="replace")
    if not text.strip():
        return {"missing": True, "path": str(path)}

    usage = {"input_tokens": 0, "cached_input_tokens": 0, "output_tokens": 0}
    tool_calls_total = 0
    commands: list[str] = []
    model_id: str | None = None
    event_count = 0

    for line in text.splitlines():
        s = line.strip()
        if not s or not s.startswith("{"):
            continue
        try:
            ev = json.loads(s)
        except json.JSONDecodeError:
            continue
        if not isinstance(ev, dict):
            continue
        event_count += 1

        if model_id is None and isinstance(ev.get("model"), str):
            model_id = ev["model"]

        etype = str(ev.get("type") or "")
        if etype in ("turn.completed", "response.completed") or "usage" in ev:
            _accumulate_usage(ev, usage)

        if etype == "item.completed":
            item = ev.get("item")
            if isinstance(item, dict):
                itype = str(item.get("type") or "")
                if itype in ("command_execution", "mcp_tool_call", "function_call", "file_change", "tool_call"):
                    tool_calls_total += 1
                    cmd = item.get("command") or item.get("name")
                    if isinstance(cmd, str) and cmd.strip():
                        commands.append(cmd.strip()[:240])

    if event_count == 0:
        return {"missing": True, "path": str(path)}

    inp = usage["input_tokens"]
    cached = usage["cached_input_tokens"]
    out = usage["output_tokens"]
    fresh_plus_out = max(0, inp - cached) + out
    total_tokens = inp + out

    return {
        "missing": False,
        "path": str(path),
        "model_id": model_id,
        "input_tokens": inp,
        "cached_input_tokens": cached,
        "output_tokens": out,
        "fresh_plus_output_tokens": fresh_plus_out,
        "total_tokens": total_tokens,
        "tool_calls_total": tool_calls_total,
        "commands": commands,
        "elapsed_seconds": round(float(elapsed_seconds), 2) if elapsed_seconds is not None else None,
    }


def parse_antigravity_transcript(
    transcript_path: str | Path,
    *,
    elapsed_seconds: float | None = None,
) -> dict[str, Any]:
    """Parse Antigravity conversation `transcript.jsonl` into the same telemetry shape."""
    path = Path(transcript_path)
    if not path.is_file():
        return {"missing": True, "path": str(path)}

    text = path.read_text(encoding="utf-8", errors="replace")
    if not text.strip():
        return {"missing": True, "path": str(path)}

    fresh_in = 0
    cached = 0
    out = 0
    tool_calls = 0
    model_id: str | None = None
    first_ts: datetime | None = None
    last_ts: datetime | None = None
    parsed_any = False

    for line in text.splitlines():
        s = line.strip()
        if not s or not s.startswith("{"):
            continue
        try:
            row = json.loads(s)
        except json.JSONDecodeError:
            continue
        if not isinstance(row, dict):
            continue
        parsed_any = True

        if model_id is None and isinstance(row.get("model"), str):
            model_id = row["model"]

        ts_raw = row.get("created_at") or row.get("createdAt") or row.get("timestamp")
        if isinstance(ts_raw, str):
            try:
                dt = datetime.fromisoformat(ts_raw.replace("Z", "+00:00"))
                if first_ts is None or dt < first_ts:
                    first_ts = dt
                if last_ts is None or dt > last_ts:
                    last_ts = dt
            except ValueError:
                pass

        if row.get("type") == "PLANNER_RESPONSE" or (
            "input_tokens" in row or "cache_read_tokens" in row or "output_tokens" in row
        ):
            fresh_in += int(row.get("input_tokens") or 0)
            cached += int(row.get("cache_read_tokens") or 0)
            out += int(row.get("output_tokens") or 0)
            tcalls = row.get("tool_calls")
            if isinstance(tcalls, list):
                tool_calls += len(tcalls)
            continue

        for ukey in ("usage", "tokenUsage", "usageMetadata"):
            u = row.get(ukey)
            if isinstance(u, dict):
                u_in = int(u.get("inputTokens") or u.get("promptTokenCount") or u.get("input_tokens") or 0)
                u_ca = int(u.get("cachedInputTokens") or u.get("cachedContentTokenCount") or u.get("cached_input_tokens") or 0)
                u_out = int(u.get("outputTokens") or u.get("candidatesTokenCount") or u.get("output_tokens") or 0)
                fresh_in += max(0, u_in - u_ca)
                cached += u_ca
                out += u_out

        if row.get("type") in ("tool_call", "command", "CORTEX_STEP_TYPE_RUN_COMMAND") or row.get("toolName"):
            tool_calls += 1

    if not parsed_any:
        return {"missing": True, "path": str(path)}

    if elapsed_seconds is None and first_ts is not None and last_ts is not None:
        elapsed_seconds = max(0.0, (last_ts - first_ts).total_seconds())

    total_inp = fresh_in + cached
    return {
        "missing": False,
        "path": str(path),
        "model_id": model_id,
        "input_tokens": total_inp,
        "cached_input_tokens": cached,
        "output_tokens": out,
        "fresh_plus_output_tokens": fresh_in + out,
        "total_tokens": total_inp + out,
        "tool_calls_total": tool_calls,
        "elapsed_seconds": round(float(elapsed_seconds), 2) if elapsed_seconds is not None else None,
    }


def scan_transcript_contamination(transcript_path: str | Path, arm: str) -> dict[str, Any]:
    """Verify Arm 0 / Arm 1 transcripts contain zero references to vndocs or repo internals."""
    if arm == ARM_FULL_SKILL:
        return {"checked": False, "contaminated": False, "hit_count": 0, "hits": []}

    path = Path(transcript_path)
    if not path.is_file():
        return {"checked": True, "contaminated": False, "hit_count": 0, "hits": [], "transcript_missing": True}

    raw = path.read_text(encoding="utf-8", errors="replace")
    hits: list[dict[str, Any]] = []
    for line_no, line in enumerate(raw.splitlines(), start=1):
        s = line.strip()
        if s.startswith("{"):
            try:
                obj = json.loads(s)
                if isinstance(obj, dict) and obj.get("type") == "SYSTEM_MESSAGE":
                    continue
            except json.JSONDecodeError:
                pass
        for label, pat in _CONTAMINATION_PATTERNS:
            m = pat.search(line)
            if m:
                hits.append({"line": line_no, "pattern": label, "match": m.group(0)})
                if len(hits) >= 25:
                    break
        if len(hits) >= 25:
            break

    return {
        "checked": True,
        "contaminated": len(hits) > 0,
        "hit_count": len(hits),
        "hits": hits,
    }


def _is_codex_quota_error(transcript_path: str | Path, stderr_path: str | Path) -> bool:
    """Detect ChatGPT/Codex usage-limit or HTTP 429/403 rate-limit failures."""
    chunks: list[str] = []
    for p_like in (transcript_path, stderr_path):
        p = Path(p_like)
        if p.is_file():
            chunks.append(p.read_text(encoding="utf-8", errors="ignore").lower())
    joined = "\n".join(chunks)
    return any(tok in joined for tok in ("usage limit", "rate limit", "429", "403 forbidden"))


def _kill_process_tree(proc: subprocess.Popen[Any]) -> None:
    """Best-effort termination of a subprocess and any Windows/POSIX descendants."""
    if proc.poll() is not None:
        return
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
            capture_output=True,
            check=False,
        )
    else:
        proc.kill()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        proc.kill()


def _run_codex_popen_with_grace(
    cmd: list[str],
    prompt_text: str,
    run_env: dict[str, str],
    transcript_path: Path,
    stderr_path: Path,
    *,
    timeout_sec: int,
    grace_sec: float = 3.0,
) -> tuple[int, bool]:
    """Run `codex exec --json` without blocking forever on inherited pipe handles after `turn.completed`."""
    timed_out = False
    turn_completed_seen_at: float | None = None
    creationflags = subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
    t0 = time.monotonic()

    with (
        transcript_path.open("w", encoding="utf-8", errors="replace") as out_f,
        stderr_path.open("w", encoding="utf-8", errors="replace") as err_f,
    ):
        proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=out_f,
            stderr=err_f,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=run_env,
            creationflags=creationflags,
        )
        if proc.stdin is not None:
            try:
                proc.stdin.write(prompt_text)
                proc.stdin.close()
            except OSError:
                pass

        while True:
            rc = proc.poll()
            if rc is not None:
                return rc, False

            now = time.monotonic()
            if now - t0 >= timeout_sec:
                timed_out = True
                _kill_process_tree(proc)
                err_f.write(f"\n[TIMEOUT after {timeout_sec}s]\n")
                return -1, True

            if transcript_path.is_file():
                txt = transcript_path.read_text(encoding="utf-8", errors="ignore")
                if '"type":"turn.completed"' in txt or '"type":"turn.failed"' in txt:
                    if turn_completed_seen_at is None:
                        turn_completed_seen_at = now
                    elif now - turn_completed_seen_at >= grace_sec:
                        _kill_process_tree(proc)
                        return 0 if '"type":"turn.completed"' in txt else 1, timed_out

            time.sleep(0.5)


def run_codex_case(
    case: dict[str, Any],
    ws_info: dict[str, Any],
    *,
    model: str | None = None,
    reasoning_effort: str | None = None,
    sandbox: str = "workspace-write",
    timeout_sec: int = 300,
    codex_bin: str | None = None,
) -> dict[str, Any]:
    """Execute one isolated (case, arm) trial via non-interactive `codex exec --json`."""
    binary = codex_bin or shutil.which("codex")
    if not binary:
        raise FileNotFoundError("`codex` CLI binary not found on PATH.")

    wdir = Path(ws_info["work_dir"])
    prompt_text = build_arm_prompt(case, ws_info)
    (wdir / "prompt.txt").write_text(prompt_text, encoding="utf-8")

    cmd: list[str] = [
        binary,
        "exec",
        "-C",
        str(wdir),
        "--skip-git-repo-check",
        "--ephemeral",
        "--ignore-rules",
        "--disable",
        "plugins",
        "--disable",
        "hooks",
        "--disable",
        "apps",
        "--disable",
        "skill_search",
        "--disable",
        "remote_plugin",
        "--enable",
        "skip_host_skill_discovery",
        "--sandbox",
        sandbox,
        "--json",
        "-c",
        "mcp_servers={}",
    ]
    if model:
        cmd.extend(["-m", model])
    if reasoning_effort:
        cmd.extend(["-c", f"model_reasoning_effort={json.dumps(reasoning_effort)}"])
    for att_path in ws_info.get("attachment_paths", []):
        if Path(att_path).suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"):
            cmd.extend(["-i", str(att_path)])
    cmd.append("-")

    transcript_path = wdir / "transcript.jsonl"
    stderr_path = wdir / "stderr.log"
    t0 = time.monotonic()
    with tempfile.TemporaryDirectory(prefix="eval-codex-home-") as clean_codex_home:
        src_auth = Path.home() / ".codex" / "auth.json"
        if src_auth.is_file():
            shutil.copy2(src_auth, Path(clean_codex_home) / "auth.json")
        (Path(clean_codex_home) / "config.toml").write_text('approval_policy = "never"\n', encoding="utf-8")
        run_env = {**os.environ, "CODEX_HOME": clean_codex_home}
        returncode, timed_out = _run_codex_popen_with_grace(
            cmd,
            prompt_text,
            run_env,
            transcript_path,
            stderr_path,
            timeout_sec=timeout_sec,
        )
    elapsed = time.monotonic() - t0

    telemetry = parse_codex_jsonl(transcript_path, elapsed_seconds=elapsed)
    if not telemetry.get("model_id") and model:
        telemetry["model_id"] = model
    contamination = scan_transcript_contamination(transcript_path, ws_info["arm"])

    run_meta = {
        "case_id": case["case_id"],
        "arm": ws_info["arm"],
        "work_dir": str(wdir),
        "returncode": returncode,
        "timed_out": timed_out,
        "sandbox": sandbox,
        "model_id": telemetry.get("model_id") or model,
        "elapsed_seconds": round(float(elapsed), 2),
    }
    (wdir / "run_meta.json").write_text(json.dumps(run_meta, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "case_id": case["case_id"],
        "arm": ws_info["arm"],
        "work_dir": str(wdir),
        "returncode": returncode,
        "timed_out": timed_out,
        "transcript": telemetry,
        "contamination": contamination,
    }


def _try_resume_codex_case(case: dict[str, Any], arm: str, wdir: Path) -> dict[str, Any] | None:
    """Reload completed Codex run metadata if `--resume` finds a clean prior execution."""
    cid = case["case_id"]
    res_path = wdir / f"{cid}.result.json"
    tr_path = wdir / "transcript.jsonl"
    meta_path = wdir / "run_meta.json"
    if not (res_path.is_file() and tr_path.is_file() and meta_path.is_file()):
        return None
    if not tr_path.read_text(encoding="utf-8", errors="ignore").strip():
        return None
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    if meta.get("returncode") != 0 or meta.get("timed_out"):
        return None

    telemetry = parse_codex_jsonl(tr_path, elapsed_seconds=meta.get("elapsed_seconds"))
    if not telemetry.get("model_id") and meta.get("model_id"):
        telemetry["model_id"] = meta["model_id"]
    contamination = scan_transcript_contamination(tr_path, arm)
    return {
        "case_id": cid,
        "arm": arm,
        "work_dir": str(wdir),
        "returncode": 0,
        "timed_out": False,
        "resumed": True,
        "transcript": telemetry,
        "contamination": contamination,
    }


def evaluate_experiment(
    cases: list[dict[str, Any]],
    run_records: list[dict[str, Any]],
    corpus_dir: str | Path,
    *,
    run_render: bool = True,
    renderer_backend: str = "libreoffice",
    repo_root: str | Path | None = None,
) -> dict[str, Any]:
    """Independently grade every (case, arm, trial) record and run full statistical comparison."""
    case_map = {c["case_id"]: c for c in cases}
    cdir = Path(corpus_dir)
    root = Path(repo_root) if repo_root is not None else _repo_root()

    graded_rows: list[dict[str, Any]] = []
    contamination_count = 0
    arms_seen: set[str] = set()
    models_seen: set[str] = set()

    for rec in run_records:
        cid = rec["case_id"]
        arm = rec["arm"]
        arms_seen.add(arm)
        case = case_map[cid]
        src_docx = cdir / case["source_docx_relpath"]
        wdir = Path(rec["work_dir"])

        res = grade_case(
            case,
            src_docx,
            wdir,
            run_render=run_render,
            renderer_backend=renderer_backend,
        )

        transcript = rec.get("transcript") or {"missing": True}
        if rec.get("model_id"):
            models_seen.add(str(rec["model_id"]))
        elif isinstance(transcript, dict) and transcript.get("model_id"):
            models_seen.add(str(transcript["model_id"]))

        contam = rec.get("contamination")
        if contam is None:
            tpath = wdir / "transcript.jsonl"
            contam = scan_transcript_contamination(tpath, arm)
        if contam.get("contaminated"):
            contamination_count += 1

        res["arm"] = arm
        res["trial_index"] = int(rec.get("trial_index", 0))
        res["model_id"] = rec.get("model_id") or transcript.get("model_id")
        res["transcript"] = transcript
        res["contamination"] = contam
        graded_rows.append(res)

    summary = summarize_experiment(graded_rows)
    arms_sorted = sorted(arms_seen)
    comparisons: dict[str, Any] = {}
    for i, a_arm in enumerate(arms_sorted):
        for b_arm in arms_sorted[i + 1 :]:
            rows_a = [r for r in graded_rows if r["arm"] == a_arm]
            rows_b = [r for r in graded_rows if r["arm"] == b_arm]
            comparisons[f"{a_arm}__vs__{b_arm}"] = compare_arms(rows_a, rows_b)

    skill_md_path = root / "SKILL.md"
    skill_sha = _sha256_file(skill_md_path) if skill_md_path.is_file() else None

    return {
        "generated_at": datetime.now().astimezone().isoformat(),
        "git_commit_sha": _safe_git_head(root),
        "skill_md_sha256": skill_sha,
        "arms": arms_sorted,
        "model_ids": sorted(models_seen),
        "n_cases": len(cases),
        "n_total_runs": len(graded_rows),
        "contamination_count": contamination_count,
        "summary": summary,
        "comparisons": comparisons,
        "rows": graded_rows,
    }


def _fmt_num(val: float | None, digits: int = 1) -> str:
    if val is None:
        return "-"
    return f"{val:.{digits}f}"


def render_experiment_markdown(report: dict[str, Any]) -> str:
    """Format an evaluated 3-Arm benchmark report as concise Markdown."""
    lines: list[str] = [
        "# Báo cáo Thực nghiệm 3 Nhánh (`val-v1` / Independent Harness)",
        "",
        f"- **Thời gian:** `{report.get('generated_at')}`",
        f"- **Git Commit SHA:** `{report.get('git_commit_sha') or 'unknown'}`",
        f"- **SKILL.md SHA-256:** `{report.get('skill_md_sha256') or 'unknown'}`",
        f"- **Models:** `{', '.join(report.get('model_ids') or ['unknown'])}`",
        f"- **Tổng lượt chạy:** `{report.get('n_total_runs')}` (`{report.get('n_cases')}` cases)",
        f"- **Nhiễm chéo Arm 0 / Arm 1 (`contamination_count`):** **{report.get('contamination_count', 0)}**",
        "",
        "## 1. Tổng hợp theo Nhánh (Arms)",
        "| Arm | Runs | `pass@1` | `pass^k` | Hit 95% CI | Checkpoint Mean | Collateral Rate | Over-Refusal | Unsafe-Edit | Mean Fresh+Out Tok | Mean Total Tok | Mean Calls | Mean Time (s) |",
        "|---|---:|---:|---:|---|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]

    by_arm = report.get("summary", {}).get("by_arm", {})
    for arm, st in by_arm.items():
        hci = st.get("hit_ci_95") or {}
        cp = st.get("checkpoint_score") or {}
        fo = (st.get("fresh_plus_output_tokens") or {}).get("mean")
        tt = (st.get("total_tokens") or {}).get("mean")
        tc = (st.get("tool_calls") or {}).get("mean")
        es = (st.get("elapsed_seconds") or {}).get("mean")
        ci_str = f"[{_fmt_num(hci.get('ci_low', 0) * 100)}%, {_fmt_num(hci.get('ci_high', 0) * 100)}%]"
        lines.append(
            f"| `{arm}` | {st['n_runs']} | {st['pass_at_1']*100:.1f}% | {st['pass_hat_k']*100:.1f}% | "
            f"{ci_str} | {cp.get('mean', 0)*100:.1f}% | {st['non_target_collateral_rate']*100:.1f}% | "
            f"{st['over_refusal_rate']*100:.1f}% | {st['unsafe_edit_rate']*100:.1f}% | "
            f"{_fmt_num(fo, 0)} | {_fmt_num(tt, 0)} | {_fmt_num(tc, 1)} | {_fmt_num(es, 1)} |"
        )

    comps = report.get("comparisons", {})
    if comps:
        lines.extend([
            "",
            "## 2. So sánh Thống kê Ghép cặp",
            "| Cặp So sánh | Paired Runs | Hit Delta | Normalized Gain $g$ | Ceiling Saturated | McNemar Exact $p$ | Wilcoxon Checkpoint $p$ | Wilcoxon Total Tok $p$ |",
            "|---|---:|---:|---:|---|---:|---:|---:|",
        ])
        for cname, cmp_res in comps.items():
            ng = cmp_res.get("normalized_gain", {})
            mc = cmp_res.get("mcnemar_hit") or {}
            wcp = cmp_res.get("wilcoxon_checkpoint") or {}
            wtk = cmp_res.get("wilcoxon_total_tokens") or {}
            g_str = "N/A (saturated)" if ng.get("ceiling_saturated") else _fmt_num(ng.get("g"), 4)
            lines.append(
                f"| `{cname}` | {cmp_res.get('n_paired_runs', 0)} | {cmp_res.get('hit_rate_pp_delta', 0):+.2f} pp | "
                f"{g_str} | `{ng.get('ceiling_saturated')}` | {_fmt_num(mc.get('p_value'), 4)} | "
                f"{_fmt_num(wcp.get('p_value'), 4)} | {_fmt_num(wtk.get('p_value'), 4)} |"
            )

    for section_title, key_name in (
        ("## 3. Phân tầng theo Nhóm Taxonomy (A–E)", "by_taxonomy_group"),
        ("## 4. Phân tầng theo Độ dài Tài liệu", "by_length_tier"),
        ("## 5. Phân tầng theo Loại Lệnh & Kết quả Kỳ vọng", "by_instruction_type"),
    ):
        sl = report.get("summary", {}).get(key_name, {})
        if not sl:
            continue
        lines.extend(["", section_title, "| Nhóm | Arm | `pass@1` | Checkpoint Mean | Mean Total Tok | Mean Calls |", "|---|---|---:|---:|---:|---:|"])
        for grp_val, arm_dict in sl.items():
            for arm, st in arm_dict.items():
                cp = (st.get("checkpoint_score") or {}).get("mean", 0.0)
                tt = (st.get("total_tokens") or {}).get("mean")
                tc = (st.get("tool_calls") or {}).get("mean")
                lines.append(
                    f"| `{grp_val}` | `{arm}` | {st['pass_at_1']*100:.1f}% | {cp*100:.1f}% | {_fmt_num(tt, 0)} | {_fmt_num(tc, 1)} |"
                )

    rows = report.get("rows", [])
    if rows:
        lines.extend([
            "",
            "## 6. Chi tiết Từng Lượt Chạy",
            "| Case | Arm | Trial | Hit | Checkpoint | Non-Target | Over-Refusal | Unsafe-Edit | Total Tok | Time (s) | Failure Reasons |",
            "|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|",
        ])
        for r in rows:
            tr = r.get("transcript") or {}
            nt = r.get("checkpoints", {}).get("semantic_non_target_preserved")
            lines.append(
                f"| `{r['case_id']}` | `{r['arm']}` | {r.get('trial_index', 0)} | {r['hit']} | "
                f"{r['checkpoint_score']*100:.0f}% | {nt} | {r.get('over_refusal')} | {r.get('unsafe_edit')} | "
                f"{tr.get('total_tokens', '-')} | {tr.get('elapsed_seconds', '-')} | {', '.join(r.get('failure_reasons', [])) or '-'} |"
            )

    return "\n".join(lines) + "\n"


def append_val_registry(
    registry_path: str | Path,
    record: dict[str, Any],
    *,
    max_milestones: int = MAX_VAL_MILESTONES,
) -> dict[str, Any]:
    """Append one milestone run to append-only `val-run-registry.jsonl` enforcing max_milestones cap."""
    required = {
        "run_id",
        "milestone_index",
        "timestamp",
        "git_commit_sha",
        "skill_md_sha256",
        "model_ids",
        "arms",
        "trials_per_case",
        "summary",
    }
    missing = sorted(k for k in required if k not in record)
    if missing:
        raise ValueError(f"Registry record missing required fields: {missing}")

    m_idx = int(record["milestone_index"])
    if m_idx < 1 or m_idx > max_milestones:
        raise ValueError(f"milestone_index={m_idx} outside [1, {max_milestones}]")

    reg_path = Path(registry_path)
    existing: list[dict[str, Any]] = []
    if reg_path.is_file():
        for line in reg_path.read_text(encoding="utf-8").splitlines():
            s = line.strip()
            if s:
                existing.append(json.loads(s))

    if len(existing) >= max_milestones:
        raise ValueError(
            f"Validation milestone budget exhausted: {len(existing)}/{max_milestones} runs already recorded."
        )
    if any(r.get("run_id") == record["run_id"] for r in existing):
        raise ValueError(f"Duplicate run_id {record['run_id']!r} in {reg_path}")

    reg_path.parent.mkdir(parents=True, exist_ok=True)
    with reg_path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, ensure_ascii=False) + "\n")
    return record


def _load_manifest_cases(manifest_path: str | Path, case_filter: list[str] | None = None) -> list[dict[str, Any]]:
    raw = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    cases = raw["cases"] if isinstance(raw, dict) and "cases" in raw else raw
    if case_filter:
        wanted = set(case_filter)
        cases = [c for c in cases if c["case_id"] in wanted]
    return cases


def main() -> int:
    parser = argparse.ArgumentParser(description="Isolated 3-Arm benchmark runner (eval_harness).")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_prep = sub.add_parser("prepare", help="Prepare isolated 3-arm workspaces and prompts without calling any LLM.")
    p_prep.add_argument("--manifest", required=True)
    p_prep.add_argument("--corpus-dir", required=True)
    p_prep.add_argument("--out-dir", required=True)
    p_prep.add_argument("--arms", nargs="+", default=list(ALL_ARMS))
    p_prep.add_argument("--trials", type=int, default=1)
    p_prep.add_argument("--cases", nargs="*", default=None)

    p_run = sub.add_parser("run", help="Execute live 3-arm Codex benchmark runs and grade independently.")
    p_run.add_argument("--manifest", required=True)
    p_run.add_argument("--corpus-dir", required=True)
    p_run.add_argument("--out-dir", required=True)
    p_run.add_argument("--arms", nargs="+", default=list(ALL_ARMS))
    p_run.add_argument("--trials", type=int, default=1)
    p_run.add_argument("--cases", nargs="*", default=None)
    p_run.add_argument("--model", default=None)
    p_run.add_argument("--reasoning-effort", default=None)
    p_run.add_argument("--sandbox", default="workspace-write", choices=["read-only", "workspace-write", "danger-full-access"])
    p_run.add_argument("--workers", type=int, default=1)
    p_run.add_argument("--resume", action="store_true")
    p_run.add_argument("--timeout-sec", type=int, default=300)
    p_run.add_argument("--no-render", action="store_true")

    p_grade = sub.add_parser("grade-existing", help="Grade existing run directories recorded in runs_manifest.json.")
    p_grade.add_argument("--manifest", required=True)
    p_grade.add_argument("--corpus-dir", required=True)
    p_grade.add_argument("--runs-manifest", required=True)
    p_grade.add_argument("--out-json", required=True)
    p_grade.add_argument("--out-md", default=None)
    p_grade.add_argument("--cases", nargs="*", default=None)
    p_grade.add_argument("--no-render", action="store_true")

    args = parser.parse_args()

    if args.cmd == "prepare":
        cases = _load_manifest_cases(args.manifest, args.cases)
        out_root = Path(args.out_dir)
        out_root.mkdir(parents=True, exist_ok=True)
        run_entries: list[dict[str, Any]] = []
        for t_idx in range(args.trials):
            for arm in args.arms:
                for case in cases:
                    wdir = out_root / arm / f"trial_{t_idx}" / case["case_id"]
                    ws = prepare_arm_workspace(case, args.corpus_dir, wdir, arm)
                    prompt = build_arm_prompt(case, ws)
                    (wdir / "prompt.txt").write_text(prompt, encoding="utf-8")
                    run_entries.append({"case_id": case["case_id"], "arm": arm, "trial_index": t_idx, **ws})
        (out_root / "runs_manifest.json").write_text(
            json.dumps({"runs": run_entries}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(json.dumps({"prepared_runs": len(run_entries), "out_dir": str(out_root)}, indent=2))
        return 0

    if args.cmd == "run":
        cases = _load_manifest_cases(args.manifest, args.cases)
        out_root = Path(args.out_dir)
        out_root.mkdir(parents=True, exist_ok=True)
        progress_path = out_root / "runs_progress.jsonl"
        jobs = [
            (t_idx, arm, case)
            for t_idx in range(args.trials)
            for arm in args.arms
            for case in cases
        ]
        total_jobs = len(jobs)
        io_lock = threading.Lock()
        done_cnt = 0
        consec_quota_errs = 0
        quota_abort = threading.Event()

        def _exec_one(item: tuple[int, str, dict[str, Any]]) -> dict[str, Any] | None:
            nonlocal done_cnt, consec_quota_errs
            t_idx, arm, case = item
            if quota_abort.is_set():
                return None
            wdir = out_root / arm / f"trial_{t_idx}" / case["case_id"]
            rec: dict[str, Any] | None = _try_resume_codex_case(case, arm, wdir) if args.resume else None
            if rec is None:
                if quota_abort.is_set():
                    return None
                ws = prepare_arm_workspace(case, args.corpus_dir, wdir, arm)
                rec = run_codex_case(
                    case,
                    ws,
                    model=args.model,
                    reasoning_effort=args.reasoning_effort,
                    sandbox=args.sandbox,
                    timeout_sec=args.timeout_sec,
                )
            rec["trial_index"] = t_idx
            is_q_err = (
                not rec.get("resumed")
                and rec.get("returncode") != 0
                and _is_codex_quota_error(wdir / "transcript.jsonl", wdir / "stderr.log")
            )
            with io_lock:
                if not rec.get("resumed"):
                    if is_q_err:
                        consec_quota_errs += 1
                        if consec_quota_errs >= 3:
                            quota_abort.set()
                    else:
                        consec_quota_errs = 0
                done_cnt += 1
                with progress_path.open("a", encoding="utf-8") as pfh:
                    pfh.write(
                        json.dumps(
                            {
                                "done": done_cnt,
                                "total": total_jobs,
                                "case_id": rec["case_id"],
                                "arm": arm,
                                "trial_index": t_idx,
                                "returncode": rec.get("returncode"),
                                "timed_out": rec.get("timed_out"),
                                "resumed": bool(rec.get("resumed", False)),
                                "quota_error": is_q_err,
                                "elapsed_seconds": (rec.get("transcript") or {}).get("elapsed_seconds"),
                                "contaminated": (rec.get("contamination") or {}).get("contaminated"),
                            },
                            ensure_ascii=False,
                        )
                        + "\n"
                    )
                tag = "RESUMED" if rec.get("resumed") else ("QUOTA_ERR" if is_q_err else f"rc={rec.get('returncode')}")
                el = (rec.get("transcript") or {}).get("elapsed_seconds")
                print(f"[{done_cnt}/{total_jobs}] {rec['case_id']} | {arm} | {tag} | {el}s", flush=True)
            return rec

        workers = max(1, int(args.workers))
        if workers == 1:
            raw_records = [_exec_one(j) for j in jobs]
        else:
            with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
                raw_records = list(pool.map(_exec_one, jobs))

        run_records = [r for r in raw_records if r is not None]
        (out_root / "runs_manifest.json").write_text(
            json.dumps({"runs": run_records}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        if quota_abort.is_set():
            completed_ok = sum(1 for r in run_records if r.get("returncode") == 0 and not r.get("timed_out"))
            print(
                json.dumps(
                    {
                        "aborted": True,
                        "reason": "consecutive_quota_errors",
                        "completed_ok": completed_ok,
                        "recorded_runs": len(run_records),
                        "total_jobs": total_jobs,
                    },
                    indent=2,
                )
            )
            return 2

        report = evaluate_experiment(cases, run_records, args.corpus_dir, run_render=not args.no_render)
        (out_root / "benchmark_report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        md = render_experiment_markdown(report)
        (out_root / "REPORT.md").write_text(md, encoding="utf-8")
        print(json.dumps({"n_total_runs": report["n_total_runs"], "contamination_count": report["contamination_count"], "by_arm": report["summary"]["by_arm"]}, indent=2))
        return 0

    if args.cmd == "grade-existing":
        cases = _load_manifest_cases(args.manifest, args.cases)
        rm = json.loads(Path(args.runs_manifest).read_text(encoding="utf-8"))
        run_records = rm["runs"] if isinstance(rm, dict) and "runs" in rm else rm
        report = evaluate_experiment(cases, run_records, args.corpus_dir, run_render=not args.no_render)
        Path(args.out_json).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        if args.out_md:
            Path(args.out_md).write_text(render_experiment_markdown(report), encoding="utf-8")
        print(json.dumps({"n_total_runs": report["n_total_runs"], "contamination_count": report["contamination_count"]}, indent=2))
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
