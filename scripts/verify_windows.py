"""Verify PE architecture and run the actual frozen executable's smoke test."""
import json
from pathlib import Path
import struct
import subprocess

ROOT = Path(__file__).resolve().parents[1]


def main():
    executable = ROOT / "dist" / "ReadQRToxlsx.exe"
    data = executable.read_bytes()
    assert data[:2] == b"MZ", "Not a Windows executable"
    pe_offset = struct.unpack_from("<I", data, 0x3C)[0]
    assert data[pe_offset:pe_offset + 4] == b"PE\0\0", "Invalid PE header"
    assert struct.unpack_from("<H", data, pe_offset + 4)[0] == 0x8664, "Expected x64 EXE"
    report_path = ROOT / "dist" / "smoke-report.json"
    report_path.unlink(missing_ok=True)
    subprocess.run([str(executable), "--smoke-test", str(report_path)], timeout=120, check=True)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["ok"], report
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
