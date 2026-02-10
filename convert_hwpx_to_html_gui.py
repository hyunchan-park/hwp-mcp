#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HWPX -> HTML GUI converter for Windows.

- This is a standalone GUI program. It does NOT require Cursor or MCP.
- Requirements: Windows + HWP installed + Python deps installed (requirements.txt).

Run (PowerShell):
  cd C:\\github\\hwp-mcp
  python -X utf8 .\\convert_hwpx_to_html_gui.py
"""

from __future__ import annotations

import queue
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from tkinter import BooleanVar, END, StringVar, Tk, filedialog, ttk

from src.tools.hwp_controller import HwpController


@dataclass(frozen=True)
class ConvertJob:
    target_folder: Path
    recursive: bool


def _iter_hwpx_in_folder(folder: Path, recursive: bool) -> list[Path]:
    if recursive:
        return sorted(p.resolve() for p in folder.rglob("*.hwpx"))
    return sorted(p.resolve() for p in folder.glob("*.hwpx"))


class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("HWPX → HTML 변환기 (hwp-mcp)")
        self.root.geometry("780x520")

        self.folder_var = StringVar(value=str(Path.cwd()))
        self.recursive_var = BooleanVar(value=False)
        self.running_var = BooleanVar(value=False)

        self._log_queue: queue.Queue[str] = queue.Queue()
        self._worker_thread: threading.Thread | None = None

        self._build_ui()
        self._tick_logs()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 8}

        frm = ttk.Frame(self.root)
        frm.pack(fill="both", expand=True, **pad)

        # Folder row
        row1 = ttk.Frame(frm)
        row1.pack(fill="x")
        ttk.Label(row1, text="대상 폴더").pack(side="left")
        ttk.Entry(row1, textvariable=self.folder_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row1, text="찾기...", command=self._on_browse).pack(side="left")

        # Options row
        row2 = ttk.Frame(frm)
        row2.pack(fill="x", pady=(6, 0))
        ttk.Checkbutton(row2, text="하위 폴더까지(재귀)", variable=self.recursive_var).pack(side="left")

        # Buttons row
        row3 = ttk.Frame(frm)
        row3.pack(fill="x", pady=(10, 0))

        self.btn_run = ttk.Button(row3, text="변환 시작", command=self._on_run)
        self.btn_run.pack(side="left")

        self.btn_stop = ttk.Button(row3, text="중지(요청)", command=self._on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(8, 0))

        self.status = ttk.Label(row3, text="대기 중")
        self.status.pack(side="right")

        # Log box
        ttk.Separator(frm).pack(fill="x", pady=10)
        ttk.Label(frm, text="로그").pack(anchor="w")

        self.log = ttk.Treeview(frm, columns=("msg",), show="headings", height=18)
        self.log.heading("msg", text="메시지")
        self.log.column("msg", width=740, anchor="w")
        self.log.pack(fill="both", expand=True, pady=(6, 0))

        # Scrollbar
        sb = ttk.Scrollbar(frm, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.place(in_=self.log, relx=1.0, rely=0, relheight=1.0, anchor="ne")

    def _append_log(self, msg: str) -> None:
        ts = time.strftime("%H:%M:%S")
        self._log_queue.put(f"[{ts}] {msg}")

    def _tick_logs(self) -> None:
        try:
            while True:
                msg = self._log_queue.get_nowait()
                self.log.insert("", END, values=(msg,))
                # auto-scroll
                children = self.log.get_children()
                if children:
                    self.log.see(children[-1])
        except queue.Empty:
            pass
        self.root.after(200, self._tick_logs)

    def _set_running(self, running: bool) -> None:
        self.running_var.set(running)
        self.btn_run.configure(state=("disabled" if running else "normal"))
        self.btn_stop.configure(state=("normal" if running else "disabled"))
        self.status.configure(text=("변환 중..." if running else "대기 중"))

    def _on_browse(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def _on_run(self) -> None:
        if self.running_var.get():
            return

        folder = Path(self.folder_var.get()).expanduser()
        if not folder.exists() or not folder.is_dir():
            self._append_log(f"[FAIL] 폴더가 유효하지 않습니다: {folder}")
            return

        job = ConvertJob(target_folder=folder.resolve(), recursive=bool(self.recursive_var.get()))
        self._set_running(True)
        self._append_log(f"시작: {job.target_folder} (recursive={job.recursive})")

        self._worker_thread = threading.Thread(target=self._run_job, args=(job,), daemon=True)
        self._worker_thread.start()

    def _on_stop(self) -> None:
        # 하드 중지는 위험하니, 사용자에게 안내용으로만 상태 표기
        self._append_log("중지 요청: 현재 파일 처리 후 멈춥니다.")
        self._stop_requested = True

    def _run_job(self, job: ConvertJob) -> None:
        self._stop_requested = False

        targets = _iter_hwpx_in_folder(job.target_folder, job.recursive)
        if not targets:
            self._append_log("변환 대상 .hwpx 파일이 없습니다.")
            self.root.after(0, lambda: self._set_running(False))
            return

        controller = HwpController()
        if not controller.connect(visible=True, register_security_module=True):
            self._append_log("[FAIL] HWP 프로그램에 연결할 수 없습니다.")
            self.root.after(0, lambda: self._set_running(False))
            return

        ok = 0
        fail = 0

        try:
            for inp in targets:
                if self._stop_requested:
                    self._append_log("중지: 요청에 의해 작업을 종료합니다.")
                    break

                out = inp.with_suffix(".html")
                self._append_log(f"open: {inp}")
                opened = controller.open_document(str(inp))
                if not opened:
                    self._append_log(f"[FAIL] open 실패: {inp}")
                    fail += 1
                    continue

                self._append_log(f"save_as_html: {out}")
                saved = controller.save_as_html(str(out))
                if saved and out.exists():
                    self._append_log(f"[OK] saved: {out}")
                    ok += 1
                else:
                    self._append_log(f"[FAIL] 저장 실패: {out}")
                    fail += 1

                controller.close_document(save=False, suppress_dialog=True)
        finally:
            controller.disconnect()

        self._append_log(f"완료: 성공 {ok}개, 실패 {fail}개")
        self.root.after(0, lambda: self._set_running(False))

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    App().run()

