"""Desktop entry point; safe to import in tests and packaging tools."""
import argparse
from pathlib import Path
import tkinter as tk
from UILayout import MyApp


def main():
    parser = argparse.ArgumentParser(description="서울반석교회 선교헌금 QR 입력")
    parser.add_argument("workbook", nargs="?", help="실행 시 열 .xlsx 파일")
    parser.add_argument("--smoke-test", metavar="REPORT", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.smoke_test:
        from packaging_smoke import run
        return run(args.smoke_test)
    root = tk.Tk()
    icon = Path(__file__).with_name("icon.ico")
    if icon.exists():
        try:
            root.iconbitmap(str(icon))
        except tk.TclError:
            pass  # macOS/Linux do not necessarily support Windows ICO icons.
    app = MyApp(root)
    if args.workbook:
        root.after(100, lambda: app.open_file(args.workbook))
    root.mainloop()


if __name__ == "__main__":
    raise SystemExit(main())
