"""
KCS 문서 정리 도구 (GUI)

hwp_to_hwpx.py / hwpx_cleanup.py / hwpx_to_hwp.py 를 명령어 없이 클릭만으로
실행할 수 있게 만든 화면입니다. 처음 보는 사람도 이 창 안의 설명만으로
바로 사용할 수 있도록 만들었습니다.

사전 준비 (한 번만):
    pip install pywin32 lxml tkinterdnd2

사용법:
    이 파일(kcs_gui.pyw)을 hwp_to_hwpx.py, hwpx_cleanup.py, hwpx_to_hwp.py 와
    같은 폴더에 놓고 더블클릭하면 됩니다. (검은 콘솔창 없이 바로 창이 뜹니다)
    탐색기에서 폴더나 파일을 목록으로 드래그해서 놓아도 됩니다.
"""

import os
import sys
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 같은 폴더의 hwp_to_hwpx.py / hwpx_cleanup.py / hwpx_to_hwp.py 를 불러오기 위함
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

INTRO_TEXT = (
    "이 프로그램은 KCS 표준시방서류 한글(HWP) 문서를 정리하는 도구입니다.\n"
    "① .hwp를 .hwpx로 변환  ②표지의 불필요한 요소/뒤쪽 불필요한 페이지 삭제  ③ 필요하면 다시 .hwp로 되돌리기.\n"
    "원본 파일은 절대 수정하지 않고, 항상 새 파일(\"..._정리됨\")로 따로 저장합니다."
)

HELP_TEXT = """\
[전체 흐름]
 1) HWP -> HWPX 변환      : .hwp 파일을 같은 이름의 .hwpx로 새로 만듭니다. (원본 유지)
 2) 표지 / 페이지 정리     : 아래 체크된 항목만 골라서 정리합니다.
 3) HWPX -> HWP 되돌리기   : 정리가 끝난 .hwpx를 다시 .hwp로 저장합니다. (필요할 때만 켜세요)

[저장 위치]
 2번(표지/페이지 정리) 작업의 결과 파일이 어디에 저장될지를 고릅니다.
 - 원본과 같은 폴더에 저장 : 기존 방식. 원본 옆에 "파일명_정리됨.hwpx"로 저장됩니다.
 - 지정한 폴더에 저장       : 폴더 하나를 정해두면, 실행할 때마다 결과 파일이 항상 그
                             폴더로 바로 저장(재실행 시 자동 덮어쓰기)됩니다. 오류를 발견해서
                             원본을 고치고 다시 돌릴 때, 결과 파일을 수동으로 옮기거나
                             덮어쓸 필요가 없어집니다. 여러 하위 폴더를 함께 처리하는
                             경우(하위 폴더까지 포함 체크) 이름이 겹치지 않도록 지정한
                             폴더 밑에 원래의 하위 폴더 구조를 그대로 만들어 저장합니다.

[표지/바탕쪽에서 정리하는 항목]
 - 워터마크/로고 이미지 제거   : 표지 및 바탕쪽(마스터페이지)에 박힌 워터마크, 기관 로고 등
                               이미지를 지웁니다.
 - 색깔·이미지 배경 제거       : 파란 세로 바, 회색 줄 같은 색상 배경이나 이미지로 채워진
                               배경을 흰 배경으로 바꿉니다. 표지뿐 아니라 바탕쪽(마스터페이지)에
                               표 형태로 박혀있는 배경도 함께 처리합니다.
 - 정보란 표 제거              : "제정 : 2016년 6월 30일 / 심의 : .../ 소관부서 : .../
                               관련단체(작성기관) : ..." 처럼 라벨이 모여있는 제정·개정
                               이력 정보란 표를 통째로 지웁니다. 그림이나 색이 아니라 순수
                               텍스트 표로 되어 있어도, 표지든 바탕쪽(마스터페이지)이든 지웁니다.
 - "표준시방서 ..." 부제 제거  : 표지의 영문 부제 문구를 지웁니다.
 - 개정/제정 날짜 제거         : "2021년 5월 12일 개정" 같은 날짜 문구를 지웁니다.
 - URL 제거                    : "http://..." 로 시작하는 문구를 지웁니다.
 - 다른 기준 잔재 텍스트 제거  : KDS 등 이 문서와 무관한 다른 기준 계열 옛 텍스트를 지웁니다.
 - 중복된 옛 표지 블록 제거    : 같은 KCS 코드/제목 묶음이 표지에 통째로 두 번 들어있는 경우
                               (예: 예전 버전 표지를 안 지우고 새 표지를 덧붙인 경우),
                               문서 고유 코드가 있는 마지막(진짜) 묶음만 남기고 앞의 옛 묶음을 지웁니다.
   (표지 자체, 즉 "KCS OO OO OO / 문서 제목" 표기는 항상 남습니다)

[페이지 삭제 항목]
 - 경과조치/연혁 페이지 삭제   : 표지와 목차 사이에 있는 개정 이력 페이지를 지웁니다.
                               (본문과 목차는 절대 건드리지 않습니다)
 - 집필위원 이후 페이지 삭제   : 본문이 끝난 뒤 나오는 집필위원/자문위원 명단, 마지막
                               작성기관 페이지를 전부 지웁니다. (그 앞의 빈 페이지도 같이 제거)

[쪽번호/꼬리말 항목]
 - 꼬리말 삭제                 : 문서에 설정된 꼬리말(페이지 하단 반복 영역) 자체를
                               통째로 지웁니다. 꼬리말 안에 들어있던 쪽번호도 함께 사라집니다.
 - 쪽번호 삭제                 : 꼬리말과 별도로(예: 머리말 쪽에) 남아있는 쪽번호 자동
                               채번 필드를 지웁니다. (꼬리말을 지우면 그 안의 쪽번호는
                               이미 같이 지워지므로, 이 항목은 보조적인 처리입니다)

[안전장치]
 - 원본은 절대 바꾸지 않습니다. 결과는 항상 "원본이름_정리됨.hwpx" 새 파일로 저장됩니다.
 - 어떤 문서에 특정 항목(예: 경과조치 페이지)이 아예 없으면, 억지로 지우지 않고
   로그에 "확인 필요"로만 표시합니다.
 - 이미 "_정리됨.hwpx"로 끝나는 파일은 다시 처리 대상에서 자동 제외됩니다.

[사용 순서]
 1. 처리할 폴더나 파일을 목록에 넣습니다 - "폴더 추가"/"파일 추가" 버튼을 쓰거나,
    탐색기에서 목록 칸으로 그냥 끌어다 놓아도(드래그 앤 드롭) 됩니다.
 2. 원하는 작업만 체크합니다 (기본값: 1, 2번의 모든 세부 항목이 켜져 있음).
 3. "실행" 버튼을 누르고 아래 로그창에서 진행 상황을 확인합니다.
"""


class LogWriter:
    """print() 출력을 큐로 보내서 GUI 로그창에 실시간으로 띄우기 위한 가짜 파일 객체."""

    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(text)

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title("KCS 문서 정리 도구")
        root.geometry("880x760")
        root.minsize(760, 600)

        self.folders = []
        self.log_queue = queue.Queue()
        self.worker_running = False

        # --- 상단 소개 ---
        intro_frame = ttk.Frame(root)
        intro_frame.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Label(intro_frame, text=INTRO_TEXT, justify="left", wraplength=840).pack(
            side="left", anchor="w")
        ttk.Button(intro_frame, text="❓ 자세한 설명", command=self.show_help).pack(
            side="right", anchor="ne")

        # --- 폴더 목록 ---
        list_title = "① 처리할 폴더/파일 목록"
        if DND_AVAILABLE:
            list_title += "  (탐색기에서 여기로 끌어다 놓아도 됩니다)"
        frame_top = ttk.LabelFrame(root, text=list_title)
        frame_top.pack(fill="both", expand=False, padx=10, pady=(10, 5))

        self.listbox = tk.Listbox(frame_top, height=5, selectmode="extended")
        self.listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)

        scrollbar = ttk.Scrollbar(frame_top, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="left", fill="y", pady=8)
        self.listbox.config(yscrollcommand=scrollbar.set)

        if DND_AVAILABLE:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind('<<Drop>>', self.on_drop)

        btn_frame = ttk.Frame(frame_top)
        btn_frame.pack(side="left", fill="y", padx=8, pady=8)
        ttk.Button(btn_frame, text="폴더 추가", command=self.add_folder).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="파일 추가", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="선택 삭제", command=self.remove_selected).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="전체 삭제", command=self.clear_folders).pack(fill="x", pady=2)

        self.var_recursive = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_top, text="하위 폴더까지 포함",
                        variable=self.var_recursive).pack(anchor="w", padx=8, pady=(0, 6))

        # --- 저장 위치 ---
        frame_out = ttk.LabelFrame(root, text="② 저장 위치 (2번 표지/페이지 정리 결과가 저장될 곳)")
        frame_out.pack(fill="x", padx=10, pady=5)

        self.var_output_same = tk.BooleanVar(value=True)
        ttk.Radiobutton(frame_out, text="원본과 같은 폴더에 저장 (파일명_정리됨.hwpx)",
                        variable=self.var_output_same, value=True,
                        command=self.on_output_mode_change).pack(anchor="w", padx=10, pady=(6, 2))

        out_row = ttk.Frame(frame_out)
        out_row.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Radiobutton(out_row, text="지정한 폴더에 저장:", variable=self.var_output_same,
                        value=False, command=self.on_output_mode_change).pack(side="left")
        self.var_output_dir = tk.StringVar(value="")
        self.output_dir_entry = ttk.Entry(out_row, textvariable=self.var_output_dir,
                                           state="disabled")
        self.output_dir_entry.pack(side="left", fill="x", expand=True, padx=6)
        self.output_dir_button = ttk.Button(out_row, text="폴더 선택", command=self.choose_output_dir,
                                             state="disabled")
        self.output_dir_button.pack(side="left")
        ttk.Label(frame_out, text="(재실행할 때마다 이 폴더로 바로 덮어써서 저장되므로,"
                                   " 오류를 고친 뒤 결과 파일을 따로 옮길 필요가 없습니다)",
                  foreground="#666666").pack(anchor="w", padx=10, pady=(0, 6))

        # --- 작업 선택 ---
        frame_opt = ttk.LabelFrame(root, text="③ 실행할 작업 (위에서부터 순서대로 실행됩니다)")
        frame_opt.pack(fill="x", padx=10, pady=5)

        self.var_step1 = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_opt, text="1) HWP → HWPX 변환",
                        variable=self.var_step1).pack(anchor="w", padx=10, pady=(6, 2))

        self.var_step2 = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_opt, text="2) 표지 정리 + 불필요한 페이지 삭제 (아래에서 세부 항목 선택)",
                        variable=self.var_step2).pack(anchor="w", padx=10, pady=2)

        sub_frame = ttk.Frame(frame_opt)
        sub_frame.pack(fill="x", padx=32, pady=(0, 4))

        import hwpx_cleanup  # 옵션 키/설명 재사용
        self.option_vars = {}
        cover_keys = ['cover_image', 'cover_color', 'cover_infobox', 'cover_subtitle',
                      'cover_date', 'cover_url', 'cover_foreign', 'cover_duplicate']
        page_keys = ['frontmatter_gap', 'trailing_matter']
        footer_keys = ['remove_footer', 'remove_page_number']

        ttk.Label(sub_frame, text="[표지 정리]", font=("", 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=(4, 0))
        for i, key in enumerate(cover_keys, start=1):
            var = tk.BooleanVar(value=hwpx_cleanup.DEFAULT_OPTIONS[key])
            self.option_vars[key] = var
            ttk.Checkbutton(sub_frame, text=hwpx_cleanup.OPTION_LABELS[key],
                            variable=var).grid(row=i, column=0, sticky="w")

        base_row = len(cover_keys) + 1
        ttk.Label(sub_frame, text="[페이지 삭제]", font=("", 9, "bold")).grid(
            row=base_row, column=0, sticky="w", pady=(8, 0))
        for i, key in enumerate(page_keys, start=1):
            var = tk.BooleanVar(value=hwpx_cleanup.DEFAULT_OPTIONS[key])
            self.option_vars[key] = var
            ttk.Checkbutton(sub_frame, text=hwpx_cleanup.OPTION_LABELS[key],
                            variable=var).grid(row=base_row + i, column=0, sticky="w")

        base_row2 = base_row + len(page_keys) + 1
        ttk.Label(sub_frame, text="[쪽번호/꼬리말]", font=("", 9, "bold")).grid(
            row=base_row2, column=0, sticky="w", pady=(8, 0))
        for i, key in enumerate(footer_keys, start=1):
            var = tk.BooleanVar(value=hwpx_cleanup.DEFAULT_OPTIONS[key])
            self.option_vars[key] = var
            ttk.Checkbutton(sub_frame, text=hwpx_cleanup.OPTION_LABELS[key],
                            variable=var).grid(row=base_row2 + i, column=0, sticky="w")

        quick_frame = ttk.Frame(sub_frame)
        quick_frame.grid(row=base_row2 + len(footer_keys) + 1, column=0, sticky="w", pady=(6, 0))
        ttk.Button(quick_frame, text="세부 항목 전체 선택",
                   command=lambda: self.set_all_options(True)).pack(side="left", padx=(0, 6))
        ttk.Button(quick_frame, text="세부 항목 전체 해제",
                   command=lambda: self.set_all_options(False)).pack(side="left")

        self.var_step3 = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_opt, text="3) HWPX → HWP로 되돌리기",
                        variable=self.var_step3).pack(anchor="w", padx=10, pady=(6, 8))

        # --- 실행 버튼 ---
        frame_run = ttk.Frame(root)
        frame_run.pack(fill="x", padx=10, pady=(0, 5))
        self.run_button = ttk.Button(frame_run, text="④ 실행", command=self.start_run)
        self.run_button.pack(side="left")
        self.status_label = ttk.Label(frame_run, text="대기 중")
        self.status_label.pack(side="left", padx=10)

        # --- 로그 ---
        frame_log = ttk.LabelFrame(root, text="진행 로그")
        frame_log.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.log_text = tk.Text(frame_log, wrap="word", state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        log_scroll = ttk.Scrollbar(frame_log, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side="left", fill="y", pady=8)
        self.log_text.config(yscrollcommand=log_scroll.set)

        self.root.after(100, self.poll_log_queue)

    # ---------- 도움말 ----------
    def show_help(self):
        win = tk.Toplevel(self.root)
        win.title("자세한 설명")
        win.geometry("640x600")
        text = tk.Text(win, wrap="word")
        text.insert("1.0", HELP_TEXT)
        text.config(state="disabled")
        text.pack(fill="both", expand=True, padx=10, pady=10)

    # ---------- 세부 옵션 ----------
    def set_all_options(self, value):
        for var in self.option_vars.values():
            var.set(value)

    # ---------- 폴더 목록 ----------
    def add_folder(self):
        path = filedialog.askdirectory(title="처리할 폴더 선택")
        if path:
            self.add_path(path)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="처리할 파일 선택",
            filetypes=[("한글 문서", "*.hwp *.hwpx"), ("모든 파일", "*.*")])
        for path in paths:
            self.add_path(path)

    def add_path(self, path):
        path = os.path.normpath(path)
        if path not in self.folders:
            self.folders.append(path)
            self.listbox.insert("end", path)

    def on_drop(self, event):
        for path in self.root.tk.splitlist(event.data):
            if os.path.exists(path):
                self.add_path(path)

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.folders[idx]

    def clear_folders(self):
        self.listbox.delete(0, "end")
        self.folders = []

    # ---------- 저장 위치 ----------
    def on_output_mode_change(self):
        if self.var_output_same.get():
            self.output_dir_entry.config(state="disabled")
            self.output_dir_button.config(state="disabled")
        else:
            self.output_dir_entry.config(state="normal")
            self.output_dir_button.config(state="normal")

    def choose_output_dir(self):
        path = filedialog.askdirectory(title="정리 결과를 저장할 폴더 선택")
        if path:
            self.var_output_dir.set(os.path.normpath(path))

    # ---------- 로그 ----------
    def poll_log_queue(self):
        try:
            while True:
                text = self.log_queue.get_nowait()
                self.log_text.config(state="normal")
                self.log_text.insert("end", text)
                self.log_text.see("end")
                self.log_text.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self.poll_log_queue)

    def append_log(self, text):
        self.log_queue.put(text)

    # ---------- 실행 ----------
    def start_run(self):
        if self.worker_running:
            messagebox.showinfo("알림", "이미 작업이 진행 중입니다.")
            return
        if not self.folders:
            messagebox.showwarning("알림", "먼저 폴더를 추가해주세요.")
            return
        if not (self.var_step1.get() or self.var_step2.get() or self.var_step3.get()):
            messagebox.showwarning("알림", "실행할 작업을 하나 이상 선택해주세요.")
            return
        if self.var_step2.get() and not any(v.get() for v in self.option_vars.values()):
            messagebox.showwarning("알림", "2번 작업의 세부 항목을 하나 이상 선택해주세요.")
            return
        if not self.var_output_same.get() and not self.var_output_dir.get().strip():
            messagebox.showwarning("알림", "저장할 폴더를 지정해주세요.")
            return

        self.run_button.config(state="disabled")
        self.status_label.config(text="작업 중...")
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        self.worker_running = True
        t = threading.Thread(target=self.run_worker, daemon=True)
        t.start()

    def run_worker(self):
        old_stdout = sys.stdout
        sys.stdout = LogWriter(self.log_queue)
        try:
            folders = list(self.folders)
            recursive = self.var_recursive.get()
            output_dir = None if self.var_output_same.get() else self.var_output_dir.get().strip()

            if self.var_step1.get():
                self.append_log("===== 1) HWP -> HWPX 변환 시작 =====\n")
                try:
                    import hwp_to_hwpx
                    hwp_to_hwpx.run(folders, recursive)
                except Exception:
                    self.append_log("[오류] HWP -> HWPX 변환 중 문제 발생:\n" + traceback.format_exc() + "\n")

            if self.var_step2.get():
                self.append_log("\n===== 2) 표지 정리 / 불필요 페이지 삭제 시작 =====\n")
                try:
                    import hwpx_cleanup
                    opts = {key: var.get() for key, var in self.option_vars.items()}
                    self.append_log("선택된 항목: " +
                                     ", ".join(hwpx_cleanup.OPTION_LABELS[k] for k, v in opts.items() if v) +
                                     "\n\n")
                    hwpx_cleanup.run(folders, recursive, opts, output_dir=output_dir)
                except Exception:
                    self.append_log("[오류] 정리 작업 중 문제 발생:\n" + traceback.format_exc() + "\n")

            if self.var_step3.get():
                self.append_log("\n===== 3) HWPX -> HWP 되돌리기 시작 =====\n")
                try:
                    import hwpx_to_hwp
                    hwpx_to_hwp.run(folders, recursive)
                except Exception:
                    self.append_log("[오류] HWPX -> HWP 변환 중 문제 발생:\n" + traceback.format_exc() + "\n")

            self.append_log("\n모든 작업이 끝났습니다.\n")
        finally:
            sys.stdout = old_stdout
            self.worker_running = False
            self.root.after(0, self.on_worker_done)

    def on_worker_done(self):
        self.run_button.config(state="normal")
        self.status_label.config(text="완료")
        messagebox.showinfo("완료", "선택한 작업이 모두 끝났습니다. 로그를 확인해주세요.")


def main():
    missing = []
    try:
        import lxml  # noqa: F401
    except ImportError:
        missing.append("lxml")
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        missing.append("pywin32")

    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    if missing:
        messagebox.showwarning(
            "패키지 필요",
            "다음 패키지가 설치되어 있지 않습니다: " + ", ".join(missing) +
            "\n\n명령 프롬프트에서 아래 명령어로 설치 후 다시 실행해주세요.\n\n"
            "pip install pywin32 lxml"
        )
    if not DND_AVAILABLE:
        messagebox.showinfo(
            "참고",
            "드래그 앤 드롭 기능을 쓰려면 아래 명령어로 tkinterdnd2를 설치한 뒤 "
            "다시 실행해주세요. (설치 안 해도 '폴더 추가'/'파일 추가' 버튼은 그대로 사용 가능합니다)\n\n"
            "pip install tkinterdnd2"
        )
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
