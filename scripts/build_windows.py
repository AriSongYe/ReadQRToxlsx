"""Build a standalone Windows x64 EXE using a Windows Python installation."""
from pathlib import Path
import platform
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]


def main():
    if sys.platform != "win32" or platform.machine().lower() not in ("amd64", "x86_64"):
        raise SystemExit("Windows x64에서 실행해 주세요. Mac에서는 Windows EXE를 만들 수 없습니다.")
    subprocess.run([
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
        "--onefile", "--windowed", "--noupx", "--name", "ReadQRToxlsx",
        "--icon", str(ROOT / "icon.ico"),
        "--add-data", f"{ROOT / 'icon.ico'};.",
        "--hidden-import", "zxingcpp",
        "--distpath", str(ROOT / "dist"),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"),
        str(ROOT / "main.py"),
    ], cwd=ROOT, check=True)
    print(ROOT / "dist" / "ReadQRToxlsx.exe")


if __name__ == "__main__":
    main()
