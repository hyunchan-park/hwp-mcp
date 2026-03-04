#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HWPX -> PDF 일괄 변환 스크립트.

한글 프로그램의 '파일 > PDF로 저장하기' 기능을 COM 자동화로 호출하여
HWPX 파일을 PDF로 변환합니다.

기본 동작:
- 경로를 지정하지 않으면: 현재 폴더의 모든 *.hwpx 파일을 변환.
- 경로를 지정하면: 해당 파일/폴더만 변환.

참고:
- Cursor/MCP 없이 터미널에서 직접 실행 가능.
- 필수: Windows + 한글(HWP) 설치 + Python 의존성 설치 (requirements.txt 참조).

사용 예시 (PowerShell):
  cd C:\\github\\hwp-mcp

  # 특정 파일 변환
  python -X utf8 .\\convert_hwpx_to_pdf.py "C:\\path\\to\\file.hwpx"

  # 폴더 내 모든 hwpx 변환
  python -X utf8 .\\convert_hwpx_to_pdf.py C:\\path\\to\\folder

  # 폴더 재귀 검색
  python -X utf8 .\\convert_hwpx_to_pdf.py C:\\path\\to\\folder --recursive

  # 출력 폴더 지정
  python -X utf8 .\\convert_hwpx_to_pdf.py "C:\\path\\to\\file.hwpx" --output-dir C:\\output
"""

from __future__ import annotations

import argparse
import sys
import time
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

    parser = argparse.ArgumentParser(description="HWPX를 PDF로 변환합니다 (한글 COM 자동화).")
    parser.add_argument(
        "paths",
        nargs="*",
        help="변환할 .hwpx 파일 또는 폴더 경로. 생략 시 현재 폴더의 모든 *.hwpx를 변환.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="폴더 경로 지정 시 하위 폴더까지 재귀 검색.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="PDF 출력 폴더. 생략 시 원본 파일과 같은 폴더에 저장.",
    )
    args = parser.parse_args(argv)

    targets = _iter_hwpx_targets(args.paths, args.recursive)
    if not targets:
        print("변환 대상 .hwpx 파일이 없습니다.")
        return 0

    print(f"변환 대상: {len(targets)}개 파일")
    for t in targets:
        print(f"  - {t.name}")
    print()

    controller = HwpController()
    if not controller.connect(visible=True, register_security_module=True):
        print("Error: HWP 프로그램에 연결할 수 없습니다.")
        return 2

    ok_count = 0
    fail_count = 0

    try:
        for idx, inp in enumerate(targets, 1):
            if inp.suffix.lower() != ".hwpx":
                print(f"[SKIP] {inp} (확장자 아님: {inp.suffix})")
                continue

            if not inp.exists():
                print(f"[FAIL] {inp} (파일 없음)")
                fail_count += 1
                continue

            # 출력 경로 결정
            if args.output_dir:
                out_dir = Path(args.output_dir).expanduser().resolve()
                out_dir.mkdir(parents=True, exist_ok=True)
                out = out_dir / inp.with_suffix(".pdf").name
            else:
                out = inp.with_suffix(".pdf")

            print(f"[{idx}/{len(targets)}] {inp.name}")
            print(f"  열기: {inp}")

            opened = controller.open_document(str(inp))
            if not opened:
                print(f"  [FAIL] 문서 열기 실패")
                fail_count += 1
                continue

            # 문서 로딩 대기
            time.sleep(1)

            print(f"  PDF 저장: {out}")
            saved = controller.save_as_pdf(str(out))

            if not saved or not out.exists():
                print(f"  [FAIL] PDF 저장 실패")
                fail_count += 1
            else:
                size_kb = out.stat().st_size / 1024
                print(f"  [OK] 저장 완료 ({size_kb:.1f} KB)")
                ok_count += 1

            # 다음 파일 처리를 위해 닫기
            controller.close_document(save=False, suppress_dialog=True)
            time.sleep(0.5)

    finally:
        controller.disconnect()

    print(f"\n{'='*50}")
    print(f"완료: 성공 {ok_count}개, 실패 {fail_count}개 (전체 {len(targets)}개)")
    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
