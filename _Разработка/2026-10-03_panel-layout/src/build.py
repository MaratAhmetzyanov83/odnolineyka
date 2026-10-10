"""Сборка SX_Layout: (R "русский текст") -> (sx:ru '(коды)), проверка скобок и литералов.

Запуск: py build.py
Выход:  2_AutoCAD/Компоновка/SX_Layout.lsp (UTF-8 без BOM, CRLF), SX_Layout.dcl (Windows-1251, CRLF)
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
OUT = os.path.join(ROOT, "2_AutoCAD", "Компоновка")


def ru(m):
    codes = " ".join(str(ord(c)) for c in m.group(1))
    return "(sx:ru '(%s))" % codes


def check_lisp(src):
    """Глубина скобок вне строк и комментариев; не-ASCII внутри строковых литералов."""
    depth, i, line, bad, n = 0, 0, 1, [], len(src)
    while i < n:
        c = src[i]
        if c == "\n":
            line += 1
        if c == ";":
            while i < n and src[i] != "\n":
                i += 1
            continue
        if c == '"':
            j = i + 1
            while j < n and src[j] != '"':
                if src[j] == "\\":
                    j += 1
                if src[j] == "\n":
                    line += 1
                if ord(src[j]) > 127:
                    bad.append(line)
                j += 1
            i = j + 1
            continue
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth < 0:
                raise SystemExit("extra ')' at line %d" % line)
        elif c == "\n" and depth == 0:
            pass
        i += 1
    return depth, sorted(set(bad))


def main():
    os.makedirs(OUT, exist_ok=True)
    src = open(os.path.join(HERE, "SX_Layout.src.lsp"), encoding="utf-8-sig").read()
    out = re.sub(r'\(R "([^"]*)"\)', ru, src)
    # (R0 "...") — русский текст до загрузки SX_Schema (своя функция sxl:ru0)
    out = re.sub(r'\(R0 "([^"]*)"\)', lambda m: ru(m).replace("(sx:ru ", "(sxl:ru0 ", 1), out)
    depth, bad = check_lisp(out)
    print("depth", depth, "nonascii literals at lines", bad)
    if depth != 0 or bad:
        sys.exit(1)
    # top-level формы: каждая (defun ...) должна начинаться с нулевой глубины
    with open(os.path.join(OUT, "SX_Layout.lsp"), "w", encoding="utf-8", newline="\r\n") as f:
        f.write(out)
    dcl = open(os.path.join(HERE, "SX_Layout.src.dcl"), encoding="utf-8-sig").read()
    with open(os.path.join(OUT, "SX_Layout.dcl"), "w", encoding="cp1251", newline="\r\n") as f:
        f.write(dcl)
    print("ok ->", OUT)


if __name__ == "__main__":
    main()
