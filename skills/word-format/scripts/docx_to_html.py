#!/usr/bin/env python3
"""
把参考 .docx 转为 Filtered HTML，用于后续读取其完整排版样式。

工作原理：驱动 Microsoft Word 自身用"另存为 → 筛选过的网页"导出 HTML。
Word 的 Filtered HTML 会把字体、字号、缩进、行距等全部以 inline CSS
写进 HTML，这是唯一能 100% 保留 Word 排版信息的文本格式。

用法：
    python docx_to_html.py <input.docx> [output.html]

Windows 使用同目录的 docx_to_html.ps1 和 Word COM；macOS 保留原 AppleScript 流程。
若未提供 output 路径，Windows 输出到 %LOCALAPPDATA%\\word-format-skill，macOS 输出到 ~/Library/Caches/word-format-skill。
注意：Equation.DSMT4 MathType/OLE 在 Filtered HTML 中会成为图片；要保留可编辑对象，请在后续渲染时使用续写模式。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


if os.name == "nt":
    CACHE_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData/Local")) / "word-format-skill"
else:
    CACHE_DIR = Path.home() / "Library/Caches/word-format-skill"


def _abs(p: str | Path) -> Path:
    return Path(p).expanduser().resolve()


def _word_running() -> bool:
    r = subprocess.run(
        ["osascript", "-e", 'tell application "System Events" to '
         '(name of processes) contains "Microsoft Word"'],
        capture_output=True, text=True,
    )
    return r.stdout.strip() == "true"


def _windows_export(docx_path: Path, html_path: Path) -> Path:
    """Use the bundled PowerShell COM implementation on Windows."""
    script = Path(__file__).with_name("docx_to_html.ps1")
    powershell = shutil.which("powershell.exe") or shutil.which("pwsh.exe") or "powershell.exe"
    result = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
         str(script), str(docx_path), str(html_path)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(f"PowerShell export failed: {message}")
    if not html_path.is_file():
        raise RuntimeError(f"Word did not produce: {html_path}")
    return html_path


def _start_word() -> None:
    subprocess.run(["open", "-a", "Microsoft Word"], check=True)
    subprocess.run(["osascript", "-e",
                    'tell application "Microsoft Word" to activate'])


def _quit_word() -> None:
    subprocess.run(["osascript", "-e",
                    'tell application "Microsoft Word" to quit saving no'])


def _try_export(docx_path: Path, html_path: Path) -> subprocess.CompletedProcess:
    script = f'''
    tell application "Microsoft Word"
        activate
        delay 2
        open POSIX file "{docx_path.as_posix()}"
        delay 2
        set theDoc to active document
        save as theDoc file name "{html_path.as_posix()}" file format format filtered HTML
        close theDoc saving no
    end tell
    '''
    return subprocess.run(["osascript", "-e", script],
                          capture_output=True, text=True)


def docx_to_html(docx_path: Path, html_path: Path) -> Path:
    docx_path = _abs(docx_path)
    html_path = _abs(html_path)
    if not docx_path.is_file():
        raise FileNotFoundError(docx_path)
    html_path.parent.mkdir(parents=True, exist_ok=True)

    if os.name == "nt":
        return _windows_export(docx_path, html_path)

    if not _word_running():
        _start_word()

    r = _try_export(docx_path, html_path)
    if r.returncode != 0 or not html_path.is_file():
        # Word 处于残留状态时 save as 会失败；quit 再 retry 一次
        _quit_word()
        import time
        time.sleep(2)
        _start_word()
        time.sleep(2)
        r = _try_export(docx_path, html_path)

    if r.returncode != 0:
        raise RuntimeError(f"AppleScript failed: {r.stderr.strip()}")
    if not html_path.is_file():
        raise RuntimeError(f"Word did not produce: {html_path}")
    return html_path


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__, file=sys.stderr)
        return 2
    src = _abs(sys.argv[1])
    if len(sys.argv) >= 3:
        dst = _abs(sys.argv[2])
    else:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        dst = CACHE_DIR / (src.stem + ".html")
    out = docx_to_html(src, dst)
    print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
