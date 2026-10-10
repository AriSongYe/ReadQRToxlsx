"""Single-window Korean UI. File I/O runs on a worker; Tk stays on its thread."""
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
from queue import Queue, Empty
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from DateFucntion import get_current_sunday
from QRParsing import decode_image
from readxlsx import ExcelManager, parse_amount


class MyApp:
    def __init__(self, root):
        self.root = root
        self.manager = None
        self.current = None
        self.busy = False
        self.closed = False
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.results = Queue()
        self.controls = []
        root.title("서울반석교회 · 선교헌금")
        root.geometry("960x760")
        root.minsize(840, 700)
        root.configure(bg="#f3f5f8")
        root.protocol("WM_DELETE_WINDOW", self.close_app)
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f3f5f8")
        style.configure("Card.TFrame", background="white")
        style.configure("TLabel", background="#f3f5f8", foreground="#243247", font=("Arial", 12))
        style.configure("Card.TLabel", background="white", foreground="#243247", font=("Arial", 12))
        style.configure("Title.TLabel", background="white", foreground="#172d4b", font=("Arial", 20, "bold"))
        style.configure("TButton", font=("Arial", 12), padding=(12, 8))
        style.configure("Accent.TButton", foreground="white", background="#2166c0")
        style.map("Accent.TButton", background=[("active", "#174d94"), ("disabled", "#aab6c6")])
        style.configure("TEntry", padding=8, font=("Arial", 14))
        style.configure("TCombobox", padding=7)
        style.configure("Treeview", rowheight=30, font=("Arial", 11))
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"), padding=6)

        banner = tk.Frame(root, bg="#172d4b", padx=26, pady=18)
        banner.pack(fill="x")
        tk.Label(banner, text="선교헌금 입력", bg="#172d4b", fg="white", font=("Arial", 23, "bold")).pack(anchor="w")
        tk.Label(banner, text="엑셀 연결 → 연도와 입금일 선택 → QR 스캔", bg="#172d4b", fg="#c7d6ea", font=("Arial", 12)).pack(anchor="w", pady=(6, 0))
        body = ttk.Frame(root, padding=(24, 16))
        body.pack(fill="both", expand=True)
        toolbar = ttk.Frame(body)
        toolbar.pack(fill="x")
        self.open_button = self._button(toolbar, "엑셀 열기", self.modify_excel_path)
        self.open_button.pack(side="left")
        self.file_label = ttk.Label(toolbar, text="연결된 파일 없음", width=53)
        self.file_label.pack(side="left", padx=12)
        self.save_button = self._button(toolbar, "파일 저장", self.save_data, accent=True)
        self.save_button.pack(side="right")

        selection = ttk.Frame(body)
        selection.pack(fill="x", pady=(16, 14))
        self.sheet_var, self.year_var, self.date_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.sheet_combo = self._combo(selection, "데이터 시트", self.sheet_var, 28, self.change_sheet)
        self.year_combo = self._combo(selection, "작정 연도", self.year_var, 10, self.change_year)
        self.date_combo = self._combo(selection, "입금일", self.date_var, 16, self.change_date)
        self.pending_var = tk.StringVar(value="저장 대기 0건")
        ttk.Label(selection, textvariable=self.pending_var).pack(side="right", anchor="s", pady=8)

        card = ttk.Frame(body, style="Card.TFrame", padding=20)
        card.pack(fill="x")
        ttk.Label(card, text="QR 스캔", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.auto_var = tk.BooleanVar(value=True)
        check = ttk.Checkbutton(card, text="작정금액이 있으면 스캔 후 자동 기록", variable=self.auto_var)
        check.grid(row=0, column=1, columnspan=2, sticky="e")
        self.controls.append((check, "normal"))
        self.qr_var, self.amount_var = tk.StringVar(), tk.StringVar()
        self.entry_qr = ttk.Entry(card, textvariable=self.qr_var, font=("Arial", 17))
        self.entry_qr.grid(row=1, column=0, sticky="ew", pady=(14, 12), padx=(0, 10))
        self.entry_qr.bind("<Return>", self.read_qr)
        self.controls.append((self.entry_qr, "normal"))
        self.read_button = self._button(card, "QR 조회", self.read_qr, accent=True)
        self.read_button.grid(row=1, column=1, padx=(0, 8))
        self.photo_button = self._button(card, "사진에서 QR 읽기", self.read_photo)
        self.photo_button.grid(row=1, column=2)
        self.result_var = tk.StringVar(value="QR 스캐너로 읽거나 코드(예: 1-2386)를 입력해 주세요.")
        ttk.Label(card, textvariable=self.result_var, style="Card.TLabel", wraplength=820, justify="left").grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 14))
        money = ttk.Frame(card, style="Card.TFrame")
        money.grid(row=3, column=0, columnspan=3, sticky="ew")
        ttk.Label(money, text="입금액 (원)", style="Card.TLabel").pack(side="left", padx=(0, 10))
        self.entry_fix = ttk.Entry(money, textvariable=self.amount_var, width=20, font=("Arial", 15))
        self.entry_fix.pack(side="left")
        self.entry_fix.bind("<Return>", self.record_current)
        self.controls.append((self.entry_fix, "normal"))
        self.record_button = self._button(money, "입금 기록 / 금액 변경", self.record_current, accent=True)
        self.record_button.pack(side="left", padx=10)
        card.columnconfigure(0, weight=1)

        heading = ttk.Frame(body)
        heading.pack(fill="x", pady=(16, 8))
        ttk.Label(heading, text="이번 실행의 입력 내역", font=("Arial", 13, "bold")).pack(side="left")
        ttk.Label(heading, text="입금 기록 후 ‘파일 저장’ 또는 종료 시 저장").pack(side="right")
        history_frame = ttk.Frame(body)
        history_frame.pack(fill="both", expand=True)
        columns = ("day", "name", "category", "amount", "cell")
        self.history = ttk.Treeview(history_frame, columns=columns, show="headings", height=5)
        for col, label, width in zip(columns, ("입금일", "이름", "항목", "입금액", "시트 / 셀"), (110, 140, 90, 110, 250)):
            self.history.heading(col, text=label)
            self.history.column(col, width=width, anchor="e" if col == "amount" else "w")
        scroll = ttk.Scrollbar(history_frame, orient="vertical", command=self.history.yview)
        self.history.configure(yscrollcommand=scroll.set)
        self.history.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.status_var = tk.StringVar(value="먼저 엑셀 파일을 열어 주세요.")
        ttk.Label(body, textvariable=self.status_var, wraplength=880).pack(fill="x", pady=(12, 0))
        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.pack(fill="x", pady=(8, 0))
        root.bind("<Control-s>", self.save_data)
        root.bind("<Command-s>", self.save_data)
        self._refresh_controls()
        self.poll_id = root.after(50, self._poll)

    def _button(self, parent, text, command, accent=False):
        button = ttk.Button(parent, text=text, command=command, style="Accent.TButton" if accent else "TButton")
        self.controls.append((button, "normal"))
        return button

    def _combo(self, parent, label, variable, width, command):
        frame = ttk.Frame(parent)
        frame.pack(side="left", padx=(0, 16))
        ttk.Label(frame, text=label).pack(anchor="w", pady=(0, 5))
        combo = ttk.Combobox(frame, textvariable=variable, width=width, state="readonly")
        combo.pack()
        combo.bind("<<ComboboxSelected>>", command)
        self.controls.append((combo, "readonly"))
        return combo

    def _refresh_controls(self):
        for widget, state in self.controls:
            widget.configure(state="disabled" if self.busy or (not self.manager and widget != self.open_button) else state)
        if not self.busy:
            ready = self.manager is not None and bool(self.date_var.get())
            for widget in (self.entry_qr, self.read_button, self.photo_button):
                widget.configure(state="normal" if ready else "disabled")
            for widget in (self.entry_fix, self.record_button):
                widget.configure(state="normal" if ready and self.current else "disabled")
            self.save_button.configure(state="normal" if self.manager and self.manager.dirty else "disabled")
        self.pending_var.set(f"저장 대기 {len(self.manager.changes) if self.manager else 0}건")
        self.root.title("서울반석교회 · 선교헌금" + (" · 저장 대기" if self.manager and self.manager.dirty else ""))

    def _run(self, work, done, status):
        if self.busy:
            return
        self.busy = True
        self.status_var.set(status)
        self.progress.start(12)
        self._refresh_controls()

        def worker():
            try:
                result = work()
            except Exception as error:
                self.results.put((False, error, done))
            else:
                self.results.put((True, result, done))
        self.executor.submit(worker)

    def _poll(self):
        try:
            success, result, done = self.results.get_nowait()
        except Empty:
            pass
        else:
            self.busy = False
            self.progress.stop()
            self._refresh_controls()
            if success:
                done(result)
            else:
                self.status_var.set(f"처리 실패: {result}")
                messagebox.showerror("처리 실패", str(result), parent=self.root)
        if not self.closed:
            self.poll_id = self.root.after(50, self._poll)

    def modify_excel_path(self):
        if self.busy:
            return
        if self.manager and self.manager.dirty:
            self.save_data(after=self._choose_file)
        else:
            self._choose_file()

    def _choose_file(self):
        path = filedialog.askopenfilename(parent=self.root, title="선교헌금 엑셀 선택", filetypes=[("Excel 통합문서", "*.xlsx")])
        if path:
            self.open_file(path)

    def open_file(self, path):
        self._run(lambda: ExcelManager(path), self._opened, "엑셀을 읽고 전체 QR을 색인하고 있습니다…")

    def _opened(self, manager):
        if self.manager:
            self.manager.close()
        self.manager = manager
        self.history.delete(*self.history.get_children())
        name = Path(manager.filepath).name
        self.file_label.configure(text=name if len(name) <= 43 else name[:40] + "…")
        self.sheet_combo.configure(values=list(manager.schemas))
        self.sheet_var.set(manager.sheet_name)
        self._set_years()
        self.status_var.set(f"연결 완료 · {manager.sheet_name} · QR {len(manager._indexes[manager.sheet_name]):,}개")
        self._refresh_controls()
        self.entry_qr.focus_set()

    def _reset_result(self):
        self.current = None
        self.amount_var.set("")
        self.result_var.set("QR 스캐너로 읽거나 코드(예: 1-2386)를 입력해 주세요.")
        self._refresh_controls()

    def change_sheet(self, event=None):
        if not self.manager or self.busy:
            return
        self._reset_result()
        previous = self.manager.sheet_name
        def done(_):
            self.sheet_var.set(self.manager.sheet_name)
            self._set_years()
            self.status_var.set(f"시트 선택: {self.manager.sheet_name}")
        # Restore visual selection immediately; apply only on successful load.
        selected = self.sheet_var.get()
        self.sheet_var.set(previous)
        self._run(lambda: self.manager.select_sheet(selected), done, "선택한 시트를 읽고 있습니다…")

    def _set_years(self):
        years = self.manager.years
        self.year_combo.configure(values=[str(year) for year in years])
        with_dates = [year for year in years if self.manager.dates(year)]
        self.year_var.set(str(date.today().year if date.today().year in with_dates else max(with_dates or years)))
        self.change_year()

    def change_year(self, event=None):
        if not self.manager or self.busy:
            return
        self._reset_result()
        dates = self.manager.dates(int(self.year_var.get()))
        self.date_combo.configure(values=[day.isoformat() for day in dates])
        sunday = get_current_sunday().date()
        past = [day for day in dates if day <= sunday]
        self.date_var.set((sunday if sunday in dates else max(past) if past else dates[0]).isoformat() if dates else "")
        self.status_var.set(f"{self.year_var.get()}년 입금일을 선택했습니다." if dates else f"{self.year_var.get()}년 작정금액은 있지만 이 시트에는 해당 연도의 입금일 열이 없습니다.")
        self._refresh_controls()

    def change_date(self, event=None):
        self._reset_result()
        self.status_var.set(f"입금일: {self.date_var.get()}")

    def read_qr(self, event=None):
        if self.busy or not self.manager or not self.date_var.get():
            return
        self.current = None
        self.amount_var.set("")
        try:
            donation = self.manager.lookup(self.qr_var.get(), int(self.year_var.get()), date.fromisoformat(self.date_var.get()))
        except ValueError as error:
            self.result_var.set(str(error))
            self.status_var.set("QR을 확인하고 다시 입력해 주세요.")
            self._refresh_controls()
            self.entry_qr.selection_range(0, "end")
            self.entry_qr.focus_set()
            return
        self.current = donation
        pledged = donation.pledged
        amount = None
        try:
            if pledged is not None:
                amount = parse_amount(pledged)
        except ValueError:
            pass
        pledge_text = f"{amount:,}원" if amount is not None else "없음 · 입금액을 직접 입력해 주세요"
        existing = f" / 기존 입금: {donation.existing:,}원" if isinstance(donation.existing, (int, float)) else " / 기존 입금 있음" if donation.existing is not None else ""
        self.result_var.set(f"{donation.name} · {donation.category} · {donation.qr}\n{donation.year}년 작정금액: {pledge_text}\n입금일: {donation.day} · {donation.sheet}!{donation.coordinate}{existing}")
        self.amount_var.set(str(donation.existing) if donation.existing is not None else str(amount) if amount is not None else "")
        self._refresh_controls()
        if self.auto_var.get() and amount is not None and donation.existing is None:
            self.record_current()
        else:
            self.status_var.set("이미 입금된 셀입니다. 변경하려면 금액을 입력하고 기록 버튼을 누르세요." if donation.existing is not None else "입금액을 확인하고 기록 버튼을 누르세요.")
            self.entry_fix.focus_set()
            self.entry_fix.selection_range(0, "end")

    def record_current(self, event=None):
        if self.busy or not self.current:
            return
        try:
            amount = parse_amount(self.amount_var.get())
            current = self.manager.lookup(self.current.qr, self.current.year, self.current.day)
            overwrite = current.existing is not None
            if overwrite and not messagebox.askyesno("기존 입금액 변경", f"{current.name} · {current.day}\n기존 {current.existing}원 → {amount:,}원으로 변경할까요?", parent=self.root):
                return
            self.manager.record(current, amount, overwrite=overwrite)
        except ValueError as error:
            self.status_var.set(str(error))
            self.entry_fix.focus_set()
            return
        self.history.insert("", 0, values=(current.day.isoformat(), current.name, current.category, f"{amount:,}원", f"{current.sheet}!{current.coordinate}"))
        self.current = self.manager.lookup(current.qr, current.year, current.day)
        self.result_var.set(f"{current.name} · {current.category} · {amount:,}원 기록 완료\n{current.day} · {current.sheet}!{current.coordinate}")
        self.status_var.set("입금 기록 완료 · 파일에 반영하려면 저장해 주세요. 다음 QR을 스캔할 수 있습니다.")
        self.qr_var.set("")
        self._refresh_controls()
        self.entry_qr.focus_set()

    def read_photo(self):
        if self.busy:
            return
        path = filedialog.askopenfilename(parent=self.root, title="QR 사진 선택", filetypes=[("QR 사진", "*.jpeg *.jpg *.png *.webp *.bmp")])
        if path:
            self._run(lambda: decode_image(path), self._photo_read, "사진의 QR 코드를 읽고 있습니다…")

    def _photo_read(self, codes):
        if len(codes) == 1:
            self.qr_var.set(codes[0])
            self.read_qr()
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("사진에서 찾은 QR 선택")
        dialog.geometry("500x300")
        ttk.Label(dialog, text="입력할 QR 한 개를 선택해 주세요.", padding=16).pack(anchor="w")
        choices = tk.Listbox(dialog, font=("Arial", 14), height=6)
        choices.pack(fill="both", expand=True, padx=16)
        for code in codes:
            try:
                donation = self.manager.lookup(code, int(self.year_var.get()), date.fromisoformat(self.date_var.get()))
                text = f"{code} · {donation.name} · {donation.category}"
            except ValueError:
                text = f"{code} · 선택한 시트에 없음 / 중복"
            choices.insert("end", text)
        choices.selection_set(0)
        def select(event=None):
            selected = choices.curselection()
            if selected:
                self.qr_var.set(codes[selected[0]])
                dialog.destroy()
                self.read_qr()
        ttk.Button(dialog, text="선택한 QR 조회", command=select).pack(pady=12)
        choices.bind("<Double-1>", select)
        dialog.transient(self.root)
        dialog.grab_set()

    def save_data(self, event=None, after=None):
        if self.busy:
            return
        if not self.manager or not self.manager.dirty:
            if after:
                after()
            return
        def done(backup):
            self.status_var.set(f"저장 완료 · 백업: {backup.name}" if backup else "저장 완료")
            self._refresh_controls()
            if after:
                after()
        self._run(self.manager.save, done, "입력 금액을 저장하고 원본을 백업하고 있습니다…")

    def close_app(self):
        if self.busy:
            self.status_var.set("진행 중인 작업이 끝난 후 종료해 주세요.")
            return
        if self.manager and self.manager.dirty:
            self.save_data(after=self._destroy)
        else:
            self._destroy()

    def _destroy(self):
        self.closed = True
        self.root.after_cancel(self.poll_id)
        if self.manager:
            self.manager.close()
        self.executor.shutdown(wait=False)
        self.root.destroy()
