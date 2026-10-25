"""Exercise actual Tk callbacks and background saves against disposable workbooks."""
from datetime import date
import gc
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

from test_workflow import fixture, DAY
from readxlsx import ExcelManager

try:
    import tkinter as tk
    from UILayout import MyApp
except ImportError:
    tk = None


@unittest.skipIf(tk is None, "Tk is not installed")
class UITests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as error:
            self.skipTest(f"Display unavailable: {error}")
        self.root.withdraw()
        self.app = MyApp(self.root)
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / "ui.xlsx"
        fixture(self.path)
        self.app.open_file(self.path)
        self.wait_idle()
        self.window = self.app.send_data()
        self.root.update_idletasks()
        self.window.top.withdraw()
        self.root.update()

    def wait_idle(self):
        deadline = time.monotonic() + 10
        while self.app.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app.busy)
        self.root.update()

    def tearDown(self):
        if hasattr(self, "app") and not self.app.closed:
            self.app._destroy()
        if hasattr(self, "app"):
            self.app.executor.shutdown(wait=True)
        if hasattr(self, "temp"):
            self.temp.cleanup()
        self.window = self.app = self.root = None
        gc.collect()

    def test_original_two_window_layout_and_reopen(self):
        self.root.deiconify()
        self.window.top.deiconify()
        self.root.update_idletasks()
        main_buttons = [widget.cget("text") for widget in self.root.pack_slaves() if isinstance(widget, tk.Button)]
        self.assertEqual(main_buttons, ["데이터 전송", "엑셀 파일 경로 수정", "나가기"])
        for bar, _, _, _ in self.app.menus.values():
            self.assertFalse(int(bar.cget("tearoff")))
            self.assertEqual(bar.type(0), "cascade")
        self.assertEqual(self.root.wm_geometry().split("+")[0], "500x400")
        self.assertEqual(self.window.top.wm_geometry().split("+")[0], "300x400")
        self.assertEqual(self.window.top.pack_slaves(), [self.window.entry_qr, self.window.read_button, self.window.entry_fix, self.window.result_label, self.window.save_button])
        self.assertEqual(self.window.read_button.cget("text"), "Read QR")
        self.assertEqual(self.window.save_button.cget("text"), "Save Data")
        self.assertLessEqual(self.app.Menual_label.winfo_y() + self.app.Menual_label.winfo_height(), 400)
        self.assertLessEqual(self.window.save_button.winfo_y() + self.window.save_button.winfo_height(), 400)
        self.assertIs(self.app.send_data(), self.window)
        self.window.close_window()
        self.assertIsNone(self.app.data_window)
        self.window = self.app.send_data()
        self.window.top.withdraw()
        self.assertIsNotNone(self.window)

    def test_selection_change_clears_stale_result_and_missing_dates_disable_entry(self):
        self.app.date_var.set(DAY.isoformat())
        self.window.qr_var.set("1-7007")
        self.window.read_qr()
        self.app.year_var.set("2025")
        self.app.change_year()
        self.assertIsNone(self.window.current)
        self.window.amount_var.set("5000")
        self.window.fixMoney()
        self.assertIsNone(self.app.manager.lookup("1-7007", 2025, date(2025, 10, 5)).existing)
        self.app.manager.schema.date_columns.pop(date(2025, 10, 5))
        self.app.change_year()
        self.assertEqual(self.app.date_var.get(), "")
        self.assertEqual(str(self.window.entry_qr["state"]), "disabled")

    def test_closing_input_window_preserves_pending_deposit(self):
        self.app.date_var.set(DAY.isoformat())
        self.window.qr_var.set("1-7007")
        self.window.read_qr()
        self.window.close_window()
        self.assertTrue(self.app.manager.dirty)
        self.window = self.app.send_data()
        self.window.top.withdraw()
        self.window.qr_var.set("1-7007")
        self.window.read_qr()
        self.assertEqual(len(self.app.manager.changes), 1)
        self.assertIn("기존 입금액", self.window.result_var.get())

    def test_scan_invalid_duplicate_and_manual_amount(self):
        self.app.date_var.set(DAY.isoformat())
        self.window.qr_var.set("1-7007")
        self.window.read_qr()
        self.assertEqual(self.app.manager.lookup("1-7007", 2026, DAY).existing, 1000)
        self.assertEqual(len(self.app.manager.changes), 1)
        self.window.qr_var.set("1-7007")
        self.window.read_qr()
        self.assertEqual(len(self.app.manager.changes), 1)
        self.window.amount_var.set("3,000")
        with patch("UILayout.messagebox.askyesno", return_value=True):
            self.window.fixMoney()
        self.assertEqual(self.app.manager.lookup("1-7007", 2026, DAY).existing, 3000)
        self.window.qr_var.set("invalid")
        self.window.read_qr()
        self.assertIsNone(self.window.current)
        self.assertEqual(str(self.window.entry_fix["state"]), "disabled")
        self.window.auto_var.set(False)
        self.window.qr_var.set("2-7007")
        self.window.read_qr()
        self.assertIsNone(self.app.manager.lookup("2-7007", 2026, DAY).existing)
        self.window.amount_var.set("0")
        self.window.fixMoney()
        self.assertEqual(self.app.manager.lookup("2-7007", 2026, DAY).existing, 0)

    def test_year_switch_and_background_save_close(self):
        self.app.year_var.set("2025")
        self.app.change_year()
        self.assertEqual(self.app.date_var.get(), "2025-10-05")
        self.window.qr_var.set("1-1")
        self.window.read_qr()
        self.assertEqual(self.app.manager.lookup("1-1", 2025, date(2025, 10, 5)).existing, 500)
        self.window.saveData()
        self.assertTrue(self.app.busy)
        self.wait_idle()
        self.assertFalse(self.app.manager.dirty)
        self.app.close_app()
        self.assertTrue(self.app.closed)

    def test_close_with_dirty_data_and_error_keeps_window(self):
        self.app.date_var.set(DAY.isoformat())
        self.window.qr_var.set("1-7007")
        self.window.read_qr()
        with patch.object(self.app.manager, "save", side_effect=PermissionError("locked")), patch("UILayout.messagebox.showerror"):
            self.app.close_app()
            self.wait_idle()
        self.assertFalse(self.app.closed)
        self.assertTrue(self.app.manager.dirty)
        self.app.close_app()
        deadline = time.monotonic() + 10
        while not self.app.closed and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertTrue(self.app.closed)
        saved = ExcelManager(self.path)
        self.assertEqual(saved.lookup("1-7007", 2026, DAY).existing, 1000)
        saved.close()


if __name__ == "__main__":
    unittest.main()
