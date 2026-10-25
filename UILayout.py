"""Original two-window Tk layout with indexed, safe workbook operations."""
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
from queue import Queue, Empty
import tkinter as tk
from tkinter import filedialog, messagebox

from DateFucntion import get_current_sunday
from QRParsing import decode_image
from readxlsx import ExcelManager, parse_amount


def set_icon(window):
    icon = Path(__file__).with_name("icon.ico")
    if icon.exists():
        try:
            window.iconbitmap(str(icon))
        except tk.TclError:
            pass


class MyApp:
    def __init__(self, root):
        self.root = root
        self.manager = None
        self.data_window = None
        self.busy = False
        self.closed = False
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.results = Queue()
        self.menus = {}
        self.sheet_var = tk.StringVar(root)
        self.year_var = tk.StringVar(root)
        self.date_var = tk.StringVar(root)
        self.status_var = tk.StringVar(root, value="엑셀 파일을 연결해 주세요.")
        root.title("서울반석교회 선교헌금 Writer")
        root.geometry("500x400")
        root.minsize(500, 400)
        root.protocol("WM_DELETE_WINDOW", self.close_app)

        # Keep the original buttons, order, spacing and native Tk appearance.
        self.send_button = tk.Button(root, text="데이터 전송", command=self.send_data)
        self.send_button.pack(pady=10)
        self.open_button = tk.Button(root, text="엑셀 파일 경로 수정", command=self.modify_excel_path)
        self.open_button.pack(pady=10)
        self.exit_button = tk.Button(root, text="나가기", command=self.close_app)
        self.exit_button.pack(pady=10)
        manual = (
            "메뉴얼\n\n"
            "1. Excel 연결 시 Open 시간이 조금 걸립니다!\n"
            "2. Save는 창을 닫을 시 자동으로 됩니다!\n"
            "3. Save는 시간이 조금 걸립니다!\n"
            "4. 입력 설정에서 연도와 입금일을 선택할 수 있습니다.\n"
            "5. 금액 수정은 입력 후 Enter를 눌러 주세요.\n"
            "6. 저장할 때는 Excel에서 해당 파일을 닫아 주세요.\n"
            "7. 저장 시 원본은 .backups 폴더에 자동 백업됩니다."
        )
        self.Menual_label = tk.Label(root, text=manual, wraplength=470)
        self.Menual_label.pack(pady=10)
        tk.Label(root, textvariable=self.status_var, wraplength=470).pack(pady=(0, 5))
        self._add_menu(root)
        self._bind_save(root)
        self._refresh_controls()
        self.poll_id = root.after(50, self._poll)

    def _bind_save(self, window):
        window.bind("<Control-s>", self.save_data)
        window.bind("<Command-s>", self.save_data)

    def _add_menu(self, window, data_window=None):
        bar = tk.Menu(window, tearoff=False)
        settings = tk.Menu(bar, tearoff=False)
        sheets = tk.Menu(settings, tearoff=False)
        years = tk.Menu(settings, tearoff=False)
        dates = tk.Menu(settings, tearoff=False)
        settings.add_cascade(label="데이터 시트", menu=sheets)
        settings.add_cascade(label="작정 연도", menu=years)
        settings.add_cascade(label="입금일", menu=dates)
        bar.add_cascade(label="입력 설정", menu=settings)
        if data_window:
            options = tk.Menu(bar, tearoff=False)
            options.add_checkbutton(label="스캔 후 작정금액 자동 기록", variable=data_window.auto_var)
            options.add_command(label="사진에서 QR 읽기", command=data_window.read_photo)
            bar.add_cascade(label="추가 기능", menu=options)
        window.configure(menu=bar)
        self.menus[window] = (bar, sheets, years, dates)
        self._refresh_menu_options()

    def _refresh_menu_options(self):
        for window, (bar, sheets, years, dates) in self.menus.items():
            for menu in (sheets, years, dates):
                for child in menu.winfo_children():
                    child.destroy()
                menu.delete(0, "end")
            if not self.manager:
                continue
            for name in self.manager.schemas:
                sheets.add_radiobutton(label=name.strip(), value=name, variable=self.sheet_var, command=self.change_sheet)
            for year in self.manager.years:
                years.add_radiobutton(label=f"{year}년", value=str(year), variable=self.year_var, command=self.change_year)
            if self.year_var.get():
                months = {}
                for day in self.manager.dates(int(self.year_var.get())):
                    if day.month not in months:
                        months[day.month] = tk.Menu(dates, tearoff=False)
                        dates.add_cascade(label=f"{day.month}월", menu=months[day.month])
                    months[day.month].add_radiobutton(label=day.isoformat(), value=day.isoformat(), variable=self.date_var, command=self.change_date)
            bar.entryconfigure(0, label=f"입력 설정 ({self.year_var.get()}년)")

    def _refresh_controls(self):
        for button in (self.send_button, self.open_button, self.exit_button):
            button.configure(state="disabled" if self.busy else "normal")
        for bar, _, _, _ in self.menus.values():
            bar.entryconfigure(0, state="normal" if self.manager and not self.busy else "disabled")
            if bar.index("end") == 1:
                bar.entryconfigure(1, state="normal" if self.manager and self.date_var.get() and not self.busy else "disabled")
        if self.data_window:
            self.data_window.refresh_controls()
        title = "서울반석교회 선교헌금 Writer"
        self.root.title(title + (" *" if self.manager and self.manager.dirty else ""))

    def _run(self, work, done, status):
        if self.busy:
            return
        self.busy = True
        self.status_var.set(status)
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
            self._refresh_controls()
            if success:
                done(result)
            else:
                self.status_var.set(f"처리 실패: {result}")
                messagebox.showerror("처리 실패", str(result), parent=self.root)
        if not self.closed:
            self.poll_id = self.root.after(50, self._poll)

    def send_data(self):
        if self.busy:
            return
        if not self.manager:
            messagebox.showinfo("File Connect Error", "xlsx파일 path를 연결해주세요!", parent=self.root)
            return
        if self.data_window:
            self.data_window.top.deiconify()
            self.data_window.top.lift()
            self.data_window.entry_qr.focus_set()
        else:
            self.data_window = DataWindow(self.root, self)
        return self.data_window

    def modify_excel_path(self):
        if self.busy:
            return
        if self.manager and self.manager.dirty:
            self.save_data(after=self._choose_file)
        else:
            self._choose_file()

    def _choose_file(self):
        path = filedialog.askopenfilename(parent=self.root, title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.open_file(path)

    def open_file(self, path):
        self._run(lambda: ExcelManager(path), self._opened, "엑셀을 읽고 전체 QR을 색인하고 있습니다…")

    def _opened(self, manager):
        if self.data_window:
            self.data_window.close_window()
        if self.manager:
            self.manager.close()
        self.manager = manager
        self.sheet_var.set(manager.sheet_name)
        self._set_years()
        self.status_var.set(f"엑셀 연결 완료: {Path(manager.filepath).name}")
        self._refresh_controls()

    def _reset_result(self):
        if self.data_window:
            self.data_window.reset_result()

    def change_sheet(self):
        if not self.manager or self.busy:
            return
        self._reset_result()
        selected = self.sheet_var.get()
        self.sheet_var.set(self.manager.sheet_name)
        def done(_):
            self.sheet_var.set(self.manager.sheet_name)
            self._set_years()
            self.status_var.set(f"시트 선택: {self.manager.sheet_name}")
        self._run(lambda: self.manager.select_sheet(selected), done, "선택한 시트를 읽고 있습니다…")

    def _set_years(self):
        years = self.manager.years
        with_dates = [year for year in years if self.manager.dates(year)]
        self.year_var.set(str(date.today().year if date.today().year in with_dates else max(with_dates or years)))
        self.change_year()

    def change_year(self):
        if not self.manager or self.busy:
            return
        self._reset_result()
        dates = self.manager.dates(int(self.year_var.get()))
        sunday = get_current_sunday().date()
        past = [day for day in dates if day <= sunday]
        self.date_var.set((sunday if sunday in dates else max(past) if past else dates[0]).isoformat() if dates else "")
        self.status_var.set(f"{self.year_var.get()}년 / {self.date_var.get()}" if dates else f"{self.year_var.get()}년 입금일 열이 없습니다. 입력 설정에서 다른 연도를 선택해 주세요.")
        self._refresh_menu_options()
        self._refresh_controls()

    def change_date(self):
        self._reset_result()
        self.status_var.set(f"입금일: {self.date_var.get()}")
        self._refresh_controls()

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


class DataWindow:
    def __init__(self, master, app):
        self.app = app
        self.current = None
        self.top = tk.Toplevel(master)
        self.top.title("Data Window")
        self.top.geometry("300x400")
        self.top.minsize(300, 400)
        self.top.protocol("WM_DELETE_WINDOW", self.close_window)
        set_icon(self.top)
        self.qr_var = tk.StringVar(self.top)
        self.amount_var = tk.StringVar(self.top)
        self.result_var = tk.StringVar(self.top)
        self.auto_var = tk.BooleanVar(self.top, value=True)
        self.entry_qr = tk.Entry(self.top, textvariable=self.qr_var)
        self.entry_qr.pack(pady=10)
        self.entry_qr.bind("<Return>", self.read_qr)
        self.read_button = tk.Button(self.top, text="Read QR", command=self.read_qr)
        self.read_button.pack(pady=10)
        self.entry_fix = tk.Entry(self.top, textvariable=self.amount_var)
        self.entry_fix.pack(pady=10)
        self.entry_fix.bind("<Return>", self.fixMoney)
        self.result_label = tk.Label(self.top, textvariable=self.result_var, wraplength=280)
        self.result_label.pack(pady=10)
        self.save_button = tk.Button(self.top, text="Save Data", command=self.saveData)
        self.save_button.pack(pady=10)
        self.app._add_menu(self.top, self)
        self.app._bind_save(self.top)
        self.refresh_controls()
        self.entry_qr.focus_set()

    def refresh_controls(self):
        ready = not self.app.busy and bool(self.app.date_var.get())
        for widget in (self.entry_qr, self.read_button):
            widget.configure(state="normal" if ready else "disabled")
        self.entry_fix.configure(state="normal" if ready and self.current else "disabled")
        self.save_button.configure(state="disabled" if self.app.busy else "normal")

    def reset_result(self):
        self.current = None
        self.amount_var.set("")
        self.result_var.set("")
        self.refresh_controls()

    def close_window(self):
        if self.app.busy:
            self.app.status_var.set("진행 중인 작업이 끝난 후 입력창을 닫아 주세요.")
            return
        self.app.menus.pop(self.top, None)
        self.app.data_window = None
        self.top.destroy()

    def read_qr(self, event=None):
        if self.app.busy or not self.app.date_var.get():
            return
        self.reset_result()
        try:
            donation = self.app.manager.lookup(self.qr_var.get(), int(self.app.year_var.get()), date.fromisoformat(self.app.date_var.get()))
        except ValueError as error:
            self.result_var.set(str(error))
            self.entry_qr.selection_range(0, "end")
            self.entry_qr.focus_set()
            return
        self.current = donation
        amount = None
        try:
            if donation.pledged is not None:
                amount = parse_amount(donation.pledged)
        except ValueError:
            pass
        pledge = f"{amount:,}원" if amount is not None else "없음 (금액 입력 후 Enter)"
        existing = f"\n기존 입금액: {donation.existing:,}원" if isinstance(donation.existing, (int, float)) else "\n기존 입금 있음" if donation.existing is not None else ""
        self.result_var.set(f"이름: {donation.name}\n작정항목: {donation.category}QR\n작정금액: {pledge}\n날짜: {donation.day}{existing}")
        self.refresh_controls()
        if self.auto_var.get() and amount is not None and donation.existing is None:
            self._record(amount)
        else:
            self.app.status_var.set("금액을 수정하려면 두 번째 입력칸에 금액을 입력하고 Enter를 누르세요.")
            self.entry_fix.focus_set()

    def fixMoney(self, event=None):
        if self.app.busy or not self.current:
            return
        try:
            amount = parse_amount(self.amount_var.get())
        except ValueError as error:
            self.result_var.set(str(error))
            self.entry_fix.focus_set()
            return
        self._record(amount)

    def _record(self, amount):
        try:
            current = self.app.manager.lookup(self.current.qr, self.current.year, self.current.day)
            overwrite = current.existing is not None
            if overwrite and not messagebox.askyesno("기존 입금액 변경", f"{current.name} · {current.day}\n기존 {current.existing}원 → {amount:,}원으로 변경할까요?", parent=self.top):
                return
            self.app.manager.record(current, amount, overwrite=overwrite)
        except ValueError as error:
            self.result_var.set(str(error))
            self.entry_fix.focus_set()
            return
        self.current = self.app.manager.lookup(current.qr, current.year, current.day)
        self.result_var.set(f"이름: {current.name}\n작정항목: {current.category}QR\n입금액: {amount:,}원\n날짜: {current.day}")
        self.app.status_var.set("입금 기록 완료 · Save Data 또는 종료 시 저장됩니다.")
        self.qr_var.set("")
        self.amount_var.set("")
        self.app._refresh_controls()
        self.entry_qr.focus_set()

    def saveData(self):
        self.app.save_data()

    def read_photo(self):
        if self.app.busy:
            return
        path = filedialog.askopenfilename(parent=self.top, title="QR 사진 선택", filetypes=[("QR 사진", "*.jpeg *.jpg *.png *.webp *.bmp")])
        if path:
            self.app._run(lambda: decode_image(path), self._photo_read, "사진의 QR 코드를 읽고 있습니다…")

    def _photo_read(self, codes):
        if len(codes) == 1:
            self.qr_var.set(codes[0])
            self.read_qr()
            return
        dialog = tk.Toplevel(self.top)
        dialog.title("QR 선택")
        tk.Label(dialog, text="입력할 QR 한 개를 선택해 주세요.").pack(pady=10)
        choices = tk.Listbox(dialog, width=35, height=6)
        choices.pack(padx=10)
        for code in codes:
            try:
                donation = self.app.manager.lookup(code, int(self.app.year_var.get()), date.fromisoformat(self.app.date_var.get()))
                label = f"{code} · {donation.name}"
            except ValueError:
                label = f"{code} · 선택한 시트에 없음 / 중복"
            choices.insert("end", label)
        choices.selection_set(0)
        def select(event=None):
            selected = choices.curselection()
            if selected:
                self.qr_var.set(codes[selected[0]])
                dialog.destroy()
                self.read_qr()
        tk.Button(dialog, text="선택한 QR 조회", command=select).pack(pady=10)
        choices.bind("<Double-1>", select)
        dialog.transient(self.top)
        dialog.grab_set()
