"""Inspect embedded assets and run only the EXE in an isolated directory."""
import argparse
from hashlib import sha256
import json
import os
from pathlib import Path
import shutil
import struct
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[1]


def inspect_executable(executable):
    from PyInstaller.archive.readers import CArchiveReader

    data = executable.read_bytes()
    assert data[:2] == b"MZ", "Not a Windows executable"
    pe = struct.unpack_from("<I", data, 0x3C)[0]
    assert data[pe:pe + 4] == b"PE\0\0", "Invalid PE header"
    assert struct.unpack_from("<H", data, pe + 4)[0] == 0x8664, "Expected x64 EXE"
    optional = pe + 24
    assert struct.unpack_from("<H", data, optional)[0] == 0x20B
    section_start = optional + struct.unpack_from("<H", data, pe + 20)[0]
    section_count = struct.unpack_from("<H", data, pe + 6)[0]

    def offset_for(rva):
        for number in range(section_count):
            section = section_start + 40 * number
            virtual_size, address, raw_size, raw_offset = struct.unpack_from("<IIII", data, section + 8)
            if address <= rva < address + max(virtual_size, raw_size):
                return raw_offset + rva - address
        raise AssertionError(f"Unmapped PE address: {rva}")

    resource_rva = struct.unpack_from("<I", data, optional + 112 + 2 * 8)[0]
    resource_base = offset_for(resource_rva)

    def entries(relative):
        directory = resource_base + relative
        named, numbered = struct.unpack_from("<HH", data, directory + 12)
        return [struct.unpack_from("<II", data, directory + 16 + 8 * number)
                for number in range(named + numbered)]

    def payload(language_directory):
        assert language_directory & 0x80000000
        _, entry = entries(language_directory & 0x7FFFFFFF)[0]
        assert not entry & 0x80000000
        rva, size = struct.unpack_from("<II", data, resource_base + entry)
        start = offset_for(rva)
        return data[start:start + size]

    types = dict(entries(0))
    assert 3 in types and 14 in types, "EXE file icon resources are missing"
    images = {identifier: payload(child) for identifier, child in entries(types[3] & 0x7FFFFFFF)}
    groups = [payload(child) for _, child in entries(types[14] & 0x7FFFFFFF)]
    source_icon = (ROOT / "icon.ico").read_bytes()
    source_count = struct.unpack_from("<H", source_icon, 4)[0]
    source_images = []
    for number in range(source_count):
        size, start = struct.unpack_from("<II", source_icon, 6 + 16 * number + 8)
        source_images.append(sha256(source_icon[start:start + size]).hexdigest())
    matched = False
    for group in groups:
        count = struct.unpack_from("<H", group, 4)[0]
        group_ids = [struct.unpack_from("<H", group, 6 + 14 * number + 12)[0] for number in range(count)]
        matched |= sorted(sha256(images[identifier]).hexdigest() for identifier in group_ids) == sorted(source_images)
    assert matched, "EXE file icon differs from icon.ico"

    archive = CArchiveReader(str(executable))
    names = set(archive.toc)
    assert archive.extract("icon.ico") == source_icon, "Bundled window icon is missing or different"
    required = {
        "Python runtime": lambda name: name.lower().startswith("python3") and name.lower().endswith(".dll"),
        "Tk native module": lambda name: name.lower().endswith("_tkinter.pyd"),
        "QR native module": lambda name: "zxingcpp" in name.lower() and name.lower().endswith(".pyd"),
        "Pillow native module": lambda name: name.lower().endswith("_imaging.cp312-win_amd64.pyd") or name.lower().endswith("_imaging.pyd"),
        "Tcl data": lambda name: name.replace("\\", "/").endswith("/init.tcl"),
        "Tk data": lambda name: name.replace("\\", "/").endswith("/tk.tcl"),
    }
    modules = {}
    for label, matches in required.items():
        found = sorted(name for name in names if matches(name))
        assert found, f"Missing bundled component: {label}"
        modules[label] = found
    return {"sha256": sha256(data).hexdigest(), "size_bytes": len(data),
            "file_icon_matches_source": True, "bundled_window_icon_matches_source": True,
            "bundled_components": modules}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--executable", type=Path, default=ROOT / "dist" / "ReadQRToxlsx.exe")
    parser.add_argument("--report", type=Path, default=ROOT / "dist" / "smoke-report.json")
    parser.add_argument("--expected-sha256")
    args = parser.parse_args()
    executable = args.executable.resolve()
    inspection = inspect_executable(executable)
    if args.expected_sha256:
        assert inspection["sha256"] == args.expected_sha256, "Different executable from the supplied file"
    with tempfile.TemporaryDirectory(prefix="readqrt-standalone-") as temporary:
        temporary = Path(temporary)
        isolated = temporary / "EXE만 있는 폴더"
        isolated.mkdir()
        copied = isolated / "ReadQRToxlsx.exe"
        shutil.copy2(executable, copied)
        assert [entry.name for entry in isolated.iterdir()] == [copied.name]
        env = os.environ.copy()
        for key in list(env):
            if key.upper().startswith(("PYTHON", "TCL_", "TK_", "PYINSTALLER_", "_PYI_")):
                env.pop(key)
        windows = Path(env.get("SystemRoot", r"C:\Windows"))
        env["PATH"] = os.pathsep.join(str(path) for path in (windows / "System32", windows))
        report_path = temporary / "smoke-report.json"
        subprocess.run([str(copied), "--smoke-test", str(report_path)],
                       cwd=isolated, env=env, timeout=120, check=True)
        report = json.loads(report_path.read_text(encoding="utf-8"))
        assert report["ok"], report
        assert [entry.name for entry in isolated.iterdir()] == [copied.name]
    report["standalone_verification"] = {
        "exe_only_directory": True, "python_removed_from_path": True,
        "python_tcl_environment_removed": True, "working_directory_outside_project": True,
        **inspection,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
