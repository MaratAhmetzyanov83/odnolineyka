"""Check library discovery with AutoCAD Core Console, without touching user apps."""
from pathlib import Path
import argparse
import shutil
import subprocess
import sys
import tempfile


def lisp_string(path):
    return '"' + str(path).replace("\\", "/").replace('"', '\\"') + '"'


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path)
    parser.add_argument("--tests", type=Path)
    parser.add_argument("--expected", type=int, default=9)
    parser.add_argument("--library", type=Path)
    parser.add_argument("--console", type=Path, default=Path("C:/Program Files/Autodesk/AutoCAD 2025/accoreconsole.exe"))
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    source = args.source or root / "2_AutoCAD/Компоновка/SX_Panel.lsp"
    with tempfile.TemporaryDirectory(prefix="cpdraw_qa_") as folder:
        folder = Path(folder)
        # SCR uses the system code page. Temporary filenames contain ASCII only.
        local_source = folder / "source.lsp"
        local_tests = folder / "test.lsp"
        shutil.copyfile(source, local_source)
        shutil.copyfile(args.tests or Path(__file__).with_name("test_panel_library.lsp"), local_tests)
        report = folder / "report.txt"
        script = folder / "run.scr"
        code = '(load ' + lisp_string(local_source) + ')\n'
        code += '(setq qa:panel-report (open ' + lisp_string(report) + ' "w"))\n'
        code += '(load ' + lisp_string(local_tests) + ')\n'
        code += '(write-line (strcat "native_findfile_restored=" (if (findfile ' + lisp_string(local_source) + ') "PASS" "FAIL")) qa:panel-report)\n'
        code += '(write-line "DONE" qa:panel-report) (close qa:panel-report) (princ)\n_QUIT\n' + ("_N" if args.library else "_Y") + "\n"
        script.write_text(code, encoding="ascii")
        command = [str(args.console), "/s", str(script), "/l", "en-US"]
        if args.library:
            local_library = folder / "library.dwg"
            shutil.copyfile(args.library, local_library)
            command.extend(["/i", str(local_library)])
        try:
            run = subprocess.run(command, cwd=folder, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=600)
        except subprocess.TimeoutExpired as error:
            print((error.stdout or b"").decode("utf-16le", errors="replace")[-5000:])
            raise
        if not report.exists():
            print(run.stdout.decode("utf-16le", errors="replace")[-3000:])
            raise RuntimeError("AutoCAD did not produce a test report")
        data = report.read_bytes()
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            text = data.decode("cp1251")
        print(text)
        if "FAIL" in text or "DONE" not in text or text.count("=PASS") != args.expected:
            print(run.stdout.decode("utf-16le", errors="replace")[-5000:])
            raise RuntimeError("Panel regression failed")


if __name__ == "__main__":
    main()
