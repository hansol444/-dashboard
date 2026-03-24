# -*- coding: utf-8 -*-
"""
gui.py — PPT 장표 번역기 데스크탑 GUI
실행: python gui.py
"""
import os, sys, queue, threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import config
from translate import translate_pptx

NAVY    = "#1B3A6B"
BG      = "#F1F5F9"
WHITE   = "#FFFFFF"
TEXT    = "#1E293B"
MUTED   = "#64748B"
SUCCESS = "#059669"
WARNING = "#D97706"
DANGER  = "#DC2626"
FONT    = "맑은 고딕"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT 장표 번역기")
        self.geometry("560x380")
        self.resizable(False, False)
        self.configure(bg=BG)

        self.file_path  = tk.StringVar()
        self.direction  = tk.StringVar(value="ko_to_en")
        self._overflows = []
        self._q         = queue.Queue()

        self._build()
        self._poll()

    # ── UI 구성 ──────────────────────────────────────────────────────────────

    def _build(self):
        # 헤더
        hdr = tk.Frame(self, bg=NAVY, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="PPT 장표 번역기", font=(FONT, 15, "bold"),
                 bg=NAVY, fg=WHITE).pack(side="left", padx=20, pady=14)
        tk.Label(hdr, text="Claude claude-sonnet-4-6 · Australian English",
                 font=(FONT, 9), bg=NAVY, fg="#94A3B8").pack(side="left", pady=18)

        # 본문
        body = tk.Frame(self, bg=BG, padx=28, pady=22)
        body.pack(fill="both", expand=True)

        # 파일 선택
        tk.Label(body, text="파일 선택", font=(FONT, 10, "bold"),
                 bg=BG, fg=TEXT).grid(row=0, column=0, sticky="w")

        file_row = tk.Frame(body, bg=BG)
        file_row.grid(row=1, column=0, sticky="ew", pady=(4, 18))
        body.columnconfigure(0, weight=1)

        self.file_entry = tk.Entry(file_row, textvariable=self.file_path,
                                   font=(FONT, 9), state="readonly",
                                   bg=WHITE, fg=TEXT, relief="solid", bd=1,
                                   readonlybackground=WHITE)
        self.file_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        tk.Button(file_row, text="찾아보기", font=(FONT, 9),
                  bg=NAVY, fg=WHITE, relief="flat", padx=14, pady=6,
                  cursor="hand2", command=self._pick).pack(side="left")

        # 번역 방향
        tk.Label(body, text="번역 방향", font=(FONT, 10, "bold"),
                 bg=BG, fg=TEXT).grid(row=2, column=0, sticky="w")

        dir_row = tk.Frame(body, bg=BG)
        dir_row.grid(row=3, column=0, sticky="w", pady=(4, 20))
        for label, val in [("한 → 영", "ko_to_en"), ("영 → 한", "en_to_ko")]:
            tk.Radiobutton(dir_row, text=label, variable=self.direction, value=val,
                           bg=BG, fg=TEXT, font=(FONT, 10),
                           selectcolor=WHITE, activebackground=BG).pack(side="left", padx=(0, 20))

        # 번역 버튼
        self.btn = tk.Button(body, text="번역 시작", font=(FONT, 11, "bold"),
                             bg=NAVY, fg=WHITE, relief="flat", pady=11,
                             cursor="hand2", command=self._start)
        self.btn.grid(row=4, column=0, sticky="ew", pady=(0, 14))

        # 진행률
        self.progress = ttk.Progressbar(body, maximum=100, length=504)
        self.progress.grid(row=5, column=0, sticky="ew")

        self.status = tk.Label(body, text="", font=(FONT, 9), bg=BG, fg=MUTED)
        self.status.grid(row=6, column=0, sticky="w", pady=(6, 0))

        self.overflow_lbl = tk.Label(body, text="", font=(FONT, 9),
                                     bg=BG, fg=WARNING, wraplength=500, justify="left")
        self.overflow_lbl.grid(row=7, column=0, sticky="w", pady=(4, 0))

    # ── 이벤트 ───────────────────────────────────────────────────────────────

    def _pick(self):
        path = filedialog.askopenfilename(
            title="PPT 파일 선택",
            filetypes=[("PowerPoint 파일", "*.pptx"), ("모든 파일", "*.*")]
        )
        if path:
            self.file_path.set(path)
            self._overflows.clear()
            self.overflow_lbl.config(text="")
            self.status.config(text="", fg=MUTED)

    def _start(self):
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("파일 없음", ".pptx 파일을 먼저 선택해주세요.")
            return

        direction = self.direction.get()
        suffix    = "_EN" if direction == "ko_to_en" else "_KO"
        out_dir   = Path(config.OUTPUT_DIR)
        out_dir.mkdir(exist_ok=True)
        output    = str(out_dir / (Path(path).stem + suffix + ".pptx"))

        self.btn.config(state="disabled", text="번역 중...")
        self.progress["value"] = 0
        self.status.config(text="번역 준비 중...", fg=MUTED)
        self._overflows.clear()
        self.overflow_lbl.config(text="")

        def run():
            try:
                translate_pptx(
                    input_path=path,
                    output_path=output,
                    direction=direction,
                    quality="fast",
                    postprocess=True,
                    make_report=True,
                    progress_callback=lambda t, d: self._q.put((t, d)),
                )
                self._q.put(("done", {"output": output}))
            except Exception as e:
                self._q.put(("error", {"msg": str(e)}))

        threading.Thread(target=run, daemon=True).start()

    # ── 큐 폴링 ──────────────────────────────────────────────────────────────

    def _poll(self):
        try:
            while True:
                ev, data = self._q.get_nowait()

                if ev == "progress":
                    self.progress["value"] = data["percent"]
                    self.status.config(
                        text=f"슬라이드 {data['slide']} / {data['total']} "
                             f"({data['elapsed']}초 경과, 약 {data['eta']}초 남음)",
                        fg=MUTED,
                    )

                elif ev == "status":
                    self.status.config(text=data["message"], fg=MUTED)

                elif ev == "overflow":
                    self._overflows.append(f"슬라이드 {data['slide']}")
                    slides = ", ".join(dict.fromkeys(self._overflows))
                    self.overflow_lbl.config(text=f"[!] 글자 초과 → 빨간 표시: {slides}")

                elif ev == "done":
                    self.progress["value"] = 100
                    out = data["output"]
                    self.status.config(text=f"완료: {out}", fg=SUCCESS)
                    self.btn.config(state="normal", text="번역 시작")
                    messagebox.showinfo("번역 완료", f"저장 위치:\n{out}")

                elif ev == "error":
                    self.status.config(text=f"오류: {data['msg']}", fg=DANGER)
                    self.btn.config(state="normal", text="번역 시작")
                    messagebox.showerror("오류 발생", data["msg"])

        except queue.Empty:
            pass

        self.after(100, self._poll)


if __name__ == "__main__":
    app = App()
    app.mainloop()
