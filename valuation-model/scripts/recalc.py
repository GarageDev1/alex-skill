#!/usr/bin/env python3
"""跨平台重算 xlsx 并安全回写公式缓存。

后端优先级(auto):Windows Excel COM → macOS Excel AppleScript →
LibreOffice headless。所有后端都只操作临时副本;副本通过 xlsx ZIP 完整性检查后
才原子替换源文件,失败时源文件保持不变。

用法:
  python recalc.py model.xlsx [more.xlsx ...]
  python recalc.py --backend auto --dry-run model.xlsx
  python recalc.py --backend excel-com model.xlsx
  python recalc.py --backend excel-applescript model.xlsx
  python recalc.py --backend libreoffice model.xlsx
"""
from __future__ import annotations

import argparse
import importlib.util
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import tempfile
import zipfile


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


BACKENDS = ("auto", "excel-com", "excel-applescript", "libreoffice")


def _libreoffice_command() -> str | None:
    for name in ("soffice", "libreoffice"):
        command = shutil.which(name)
        if command:
            return command
    return None


def _module_available(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ModuleNotFoundError):
        return False


def _mac_excel_installed() -> bool:
    return any(
        path.expanduser().exists()
        for path in (
            Path("/Applications/Microsoft Excel.app"),
            Path("~/Applications/Microsoft Excel.app"),
        )
    )


def available_backends() -> dict[str, bool]:
    system = platform.system()
    return {
        "excel-com": system == "Windows" and _module_available("win32com.client"),
        "excel-applescript": (
            system == "Darwin"
            and shutil.which("osascript") is not None
            and _mac_excel_installed()
        ),
        "libreoffice": _libreoffice_command() is not None,
    }


def choose_backends(requested: str) -> list[str]:
    available = available_backends()
    if requested != "auto":
        if not available[requested]:
            raise RuntimeError(f"后端不可用: {requested}")
        return [requested]
    candidates = [
        candidate
        for candidate in ("excel-com", "excel-applescript", "libreoffice")
        if available[candidate]
    ]
    if not candidates:
        raise RuntimeError("没有可用重算后端:请安装 Microsoft Excel(Windows/macOS)或 LibreOffice")
    return candidates


def _recalc_excel_com(path: Path) -> None:
    import win32com.client as win32

    try:
        excel = win32.DispatchEx("Excel.Application")
    except Exception:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(path.resolve()))
        excel.CalculateFullRebuild()
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


def _recalc_excel_applescript(path: Path) -> None:
    escaped = str(path.resolve()).replace("\\", "\\\\").replace('"', '\\"')
    script = f'''
tell application "Microsoft Excel"
    set display alerts to false
    open POSIX file "{escaped}"
    calculate full
    save active workbook
    close active workbook saving no
end tell
'''
    subprocess.run(["osascript", "-e", script], check=True, capture_output=True, text=True)


def _recalc_libreoffice(path: Path, workdir: Path) -> Path:
    command = _libreoffice_command()
    if not command:
        raise RuntimeError("找不到 soffice/libreoffice")
    converted = workdir / "converted"
    converted.mkdir()
    proc = subprocess.run(
        [command, "--headless", "--convert-to", "xlsx", "--outdir", str(converted), str(path)],
        check=True,
        capture_output=True,
        text=True,
    )
    output = converted / path.name
    if not output.exists():
        raise RuntimeError(f"LibreOffice 未生成输出:{proc.stdout.strip()} {proc.stderr.strip()}")
    return output


def _validate_xlsx(path: Path) -> None:
    if not path.exists() or path.stat().st_size == 0:
        raise RuntimeError("重算结果不存在或为空")
    if not zipfile.is_zipfile(path):
        raise RuntimeError("重算结果不是有效 xlsx ZIP")
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad:
            raise RuntimeError(f"xlsx ZIP 损坏:{bad}")
        if "xl/workbook.xml" not in archive.namelist():
            raise RuntimeError("重算结果缺少 xl/workbook.xml")


def recalc_one(source: Path, backend: str) -> None:
    source = source.expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if source.suffix.lower() != ".xlsx":
        raise ValueError(f"仅支持 .xlsx:{source}")

    with tempfile.TemporaryDirectory(prefix=f".{source.stem}.recalc-", dir=source.parent) as tmp:
        workdir = Path(tmp)
        staged = workdir / source.name
        shutil.copy2(source, staged)
        if backend == "excel-com":
            _recalc_excel_com(staged)
            result = staged
        elif backend == "excel-applescript":
            _recalc_excel_applescript(staged)
            result = staged
        else:
            result = _recalc_libreoffice(staged, workdir)
        _validate_xlsx(result)
        replacement = source.with_name(f".{source.name}.recalc-ready")
        shutil.copy2(result, replacement)
        try:
            os.replace(replacement, source)
        finally:
            if replacement.exists():
                replacement.unlink()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--backend", choices=BACKENDS, default="auto")
    parser.add_argument("--dry-run", action="store_true", help="只检查后端与输入,不创建或修改文件")
    args = parser.parse_args(argv)

    try:
        backends = choose_backends(args.backend)
    except Exception as exc:
        print(f"recalc FAIL:{exc}", file=sys.stderr)
        return 1

    failed = False
    for path in args.paths:
        resolved = path.expanduser().resolve()
        if not resolved.is_file() or resolved.suffix.lower() != ".xlsx":
            print(f"recalc FAIL (输入无效):{path}", file=sys.stderr)
            failed = True
            continue
        if args.dry_run:
            print(f"recalc DRY-RUN backends={','.join(backends)}:{resolved}")
            continue
        errors: list[str] = []
        for backend in backends:
            try:
                recalc_one(resolved, backend)
                print(f"recalc OK backend={backend}:{resolved}")
                break
            except Exception as exc:
                errors.append(f"{backend}: {exc}")
                if args.backend == "auto" and backend != backends[-1]:
                    print(f"recalc WARN backend={backend}，尝试下一后端: {exc}", file=sys.stderr)
        else:
            print(f"recalc FAIL:{resolved} — {' | '.join(errors)}", file=sys.stderr)
            failed = True
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
