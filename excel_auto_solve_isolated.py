from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import win32com.client

import excel_batch_linked_solver as batch


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


ROOT = Path(r"C:\Users\cyc20\Desktop\供应链管理")
TARGET_DIR = Path(r"C:\Users\cyc20\Desktop\excel")
SUMMARY_PATH = TARGET_DIR / "_excel自动求解汇总.json"


def log(message: str) -> None:
    print(f"[excel-auto-solve] {message}", flush=True)


def automation_excel_pids() -> set[int]:
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
        import wmi  # type: ignore
    except Exception:
        pass
    try:
        output = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE'\" | "
                "Where-Object { $_.CommandLine -like '*/automation*' -or $_.CommandLine -like '*-Embedding*' } | "
                "ForEach-Object { $_.ProcessId }",
            ],
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
    except Exception:
        return set()
    return {int(line.strip()) for line in output.splitlines() if line.strip().isdigit()}


def kill_pids(pids: set[int]) -> None:
    if not pids:
        return
    cmd = (
        "$ids=@("
        + ",".join(str(pid) for pid in sorted(pids))
        + "); foreach($id in $ids){ Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", cmd], check=False)


def clean_target_dir(target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    for path in target_dir.iterdir():
        if path.is_file():
            path.unlink()


def run_child(source: Path, target_dir: Path, max_time: int, wall_timeout: int) -> dict[str, Any]:
    before = automation_excel_pids()
    cmd = [
        sys.executable,
        "-X",
        "utf8",
        str(Path(__file__).resolve()),
        "--source-file",
        str(source),
        "--target-dir",
        str(target_dir),
        "--solve",
        "--max-time",
        str(max_time),
    ]
    started = time.time()
    try:
        completed = subprocess.run(
            cmd,
            cwd=str(ROOT),
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            timeout=wall_timeout,
        )
        elapsed = round(time.time() - started, 2)
        result = {
            "file": source.name,
            "returncode": completed.returncode,
            "elapsed_seconds": elapsed,
            "stdout": completed.stdout[-4000:],
            "stderr": completed.stderr[-4000:],
            "timeout": False,
        }
        return result
    except subprocess.TimeoutExpired as exc:
        elapsed = round(time.time() - started, 2)
        after = automation_excel_pids()
        kill_pids(after - before)
        return {
            "file": source.name,
            "returncode": None,
            "elapsed_seconds": elapsed,
            "stdout": (exc.stdout or "")[-4000:] if isinstance(exc.stdout, str) else "",
            "stderr": (exc.stderr or "")[-4000:] if isinstance(exc.stderr, str) else "",
            "timeout": True,
        }


def status_override(source: Path, target_dir: Path, message: str) -> dict[str, Any]:
    cmd = [
        sys.executable,
        "-X",
        "utf8",
        str(Path(__file__).resolve()),
        "--source-file",
        str(source),
        "--target-dir",
        str(target_dir),
        "--status-override",
        message,
    ]
    completed = subprocess.run(
        cmd,
        cwd=str(ROOT),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        timeout=180,
    )
    return {
        "file": source.name,
        "returncode": completed.returncode,
        "stdout": completed.stdout[-2000:],
        "stderr": completed.stderr[-2000:],
        "override": message,
    }


def read_workbook_status(path: Path) -> dict[str, Any]:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(path.resolve()), ReadOnly=True)
        model = workbook.Worksheets("Excel规划模型")
        result = workbook.Worksheets("求解结果")
        return {
            "file": path.name,
            "solver_status": result.Range("B2").Value,
            "objective": result.Range("B7").Value,
            "transport_cost": result.Range("B8").Value,
            "purchase_cost": result.Range("B9").Value,
            "shortage_penalty": result.Range("B10").Value,
            "score": result.Range("B11").Value,
            "model_score": model.Range("B6").Value,
        }
    finally:
        if workbook is not None:
            workbook.Close(False)
        excel.Quit()


def run_batch(root: Path, target_dir: Path, max_time: int, wall_timeout: int) -> None:
    batch.ensure_solver_xlam_for_com()
    clean_target_dir(target_dir)
    sources = batch.discover_source_files(root)
    rows: list[dict[str, Any]] = []
    for source in sources:
        log(f"自动点击 Solver: {source.name}")
        result = run_child(source, target_dir, max_time=max_time, wall_timeout=wall_timeout)
        if result["timeout"]:
            log(f"超时: {source.name}，写入可手动求解状态")
            override = status_override(source, target_dir, f"自动求解超时（外层限制 {wall_timeout} 秒）")
            result["override"] = override
        elif result["returncode"] != 0:
            log(f"失败: {source.name}，写入失败状态")
            override = status_override(source, target_dir, f"自动求解失败（返回码 {result['returncode']}）")
            result["override"] = override
        target = target_dir / source.name
        if target.exists():
            try:
                result["workbook_status"] = read_workbook_status(target)
            except Exception as exc:
                result["workbook_status_error"] = str(exc)
        rows.append(result)
        SUMMARY_PATH.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"汇总: {SUMMARY_PATH}")
    for row in rows:
        status = row.get("workbook_status", {}).get("solver_status", row.get("override", {}).get("override", "未知"))
        score = row.get("workbook_status", {}).get("score", "")
        log(f"{row['file']} | 状态={status} | 分数={score} | 秒={row['elapsed_seconds']}")


def main() -> None:
    parser = argparse.ArgumentParser(description="隔离运行 Excel 内置 Solver，避免单个工作簿卡死整批")
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--target-dir", type=Path, default=TARGET_DIR)
    parser.add_argument("--source-file", type=Path, default=None)
    parser.add_argument("--solve", action="store_true")
    parser.add_argument("--max-time", type=int, default=120)
    parser.add_argument("--wall-timeout", type=int, default=180)
    parser.add_argument("--status-override", default=None)
    args = parser.parse_args()

    if args.source_file:
        args.target_dir.mkdir(parents=True, exist_ok=True)
        source = args.source_file.resolve()
        target = args.target_dir / source.name
        shutil.copy2(source, target)
        result = batch.build_workbook_model(
            target,
            solve_model=args.solve,
            max_time=args.max_time,
            status_override=args.status_override,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    run_batch(args.root, args.target_dir, max_time=args.max_time, wall_timeout=args.wall_timeout)


if __name__ == "__main__":
    main()
