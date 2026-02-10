#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HWPX -> HTML batch converter (same folder, same filename).

Default behavior:
- If no paths are provided: convert all *.hwpx in the current working directory.
- If paths are provided: convert those files/folders only.

Notes:
- This script does NOT require Cursor or MCP. It can be run directly in a terminal.
- Requirements: Windows + HWP installed + Python deps installed (see requirements.txt).

Examples (PowerShell):
  cd C:\\github\\hwp-mcp
  python .\\convert_hwpx_to_html.py

  python .\\convert_hwpx_to_html.py .\\L2-proposal-2.hwpx .\\L2-proposal-3.hwpx

  # Convert all *.hwpx in a folder
  python .\\convert_hwpx_to_html.py C:\\path\\to\\folder
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from src.tools.hwp_controller import HwpController


def _iter_hwpx_targets(paths: list[str], recursive: bool) -> list[Path]:
    if paths:
        out: set[Path] = set()
        for p in paths:
            pp = Path(p).expanduser().resolve()

            if pp.is_dir():
                it = pp.rglob("*.hwpx") if recursive else pp.glob("*.hwpx")
                for f in it:
                    out.add(f.resolve())
                continue

            out.add(pp)

        return sorted(out)

    cwd = Path.cwd()
    return sorted(cwd.glob("*.hwpx"))


def main(argv: list[str]) -> int:
    try:
        # Windows 콘솔에서 한글 출력 안정화
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="Convert HWPX to HTML (same folder).")
    parser.add_argument(
        "paths",
        nargs="*",
        help="Optional .hwpx file paths or folder paths. If omitted, converts all *.hwpx in current directory.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="If a folder path is provided, search *.hwpx recursively.",
    )
    args = parser.parse_args(argv)

    targets = _iter_hwpx_targets(args.paths, args.recursive)
    if not targets:
        print("변환 대상 .hwpx 파일이 없습니다.")
        return 0

    controller = HwpController()
    if not controller.connect(visible=True, register_security_module=True):
        print("Error: HWP 프로그램에 연결할 수 없습니다.")
        return 2

    ok_count = 0
    fail_count = 0

    try:
        for inp in targets:
            if inp.suffix.lower() != ".hwpx":
                print(f"[SKIP] {inp} (확장자 아님: {inp.suffix})")
                continue

            if not inp.exists():
                print(f"[FAIL] {inp} (파일 없음)")
                fail_count += 1
                continue

            out = inp.with_suffix(".html")
            print(f"\n== {inp.name} ==")
            print(f"- open: {inp}")
            opened = controller.open_document(str(inp))
            if not opened:
                print(f"[FAIL] open 실패: {inp}")
                fail_count += 1
                continue

            print(f"- save_as_html: {out}")
            saved = controller.save_as_html(str(out))
            if not saved or not out.exists():
                print(f"[FAIL] html 저장 실패: {out}")
                fail_count += 1
            else:
                print(f"[OK]  saved: {out}")
                ok_count += 1

            # 다음 파일 처리를 위해 닫기
            controller.close_document(save=False, suppress_dialog=True)
    finally:
        controller.disconnect()

    print(f"\n완료: 성공 {ok_count}개, 실패 {fail_count}개")
    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

