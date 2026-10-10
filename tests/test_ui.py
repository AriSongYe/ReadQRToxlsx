"""Exercise actual Tk callbacks and background saves against disposable workbooks."""
from datetime import date
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
        if hasattr(self, "temp"):
            self.temp.cleanup()

    def test_scan_invalid_duplicate_and_manual_amount(self):
        self.app.date_var.set(DAY.isoformat())
        self.app.qr_var.set("1-7007")
        self.app.read_qr()
        self.assertEqual(self.app.manager.lookup("1-7007", 2026, DAY).existing, 1000)
        self.assertEqual(len(self.app.history.get_children()), 1)
        self.app.qr_var.set("1-7007")
        self.app.read_qr()
        self.assertEqual(len(self.app.history.get_children()), 1)
        self.app.amount_var.set("3,000")
        with patch("UILayout.messagebox.askyesno", return_value=True):
            self.app.record_current()
        self.assertEqual(self.app.manager.lookup("1-7007", 2026, DAY).existing, 3000)
        self.app.qr_var.set("invalid")
        self.app.read_qr()
        self.assertIsNone(self.app.current)
        self.assertEqual(str(self.app.record_button["state"]), "disabled")
        self.app.auto_var.set(False)
        self.app.qr_var.set("2-7007")
        self.app.read_qr()
        self.assertIsNone(self.app.manager.lookup("2-7007", 2026, DAY).existing)
        self.app.amount_var.set("0")
        self.app.record_current()
        self.assertEqual(self.app.manager.lookup("2-7007", 2026, DAY).existing, 0)

    def test_year_switch_and_background_save_close(self):
        self.app.year_var.set("2025")
        self.app.change_year()
        self.assertEqual(self.app.date_var.get(), "2025-10-05")
        self.app.qr_var.set("1-1")
        self.app.read_qr()
        self.assertEqual(self.app.manager.lookup("1-1", 2025, date(2025, 10, 5)).existing, 500)
        self.app.save_data()
        self.assertTrue(self.app.busy)
        self.wait_idle()
        self.assertFalse(self.app.manager.dirty)
        self.app.close_app()
        self.assertTrue(self.app.closed)

    def test_close_with_dirty_data_and_error_keeps_window(self):
        self.app.date_var.set(DAY.isoformat())
        self.app.qr_var.set("1-7007")
        self.app.read_qr()
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
