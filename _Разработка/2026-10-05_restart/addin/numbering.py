# -*- coding: utf-8 -*-
"""Номера строк в процедурах VBA — чтобы Erl в сообщении об ошибке показывал строку."""
import re

SKIP = ("Dim ", "Case ", "Select ", "End Select", "End Sub", "End Function", "Else", "ElseIf", "End If",
        "Next", "Loop", "End With", "Private Const", "Public Const", "Option ", "Attribute ")
PROC = re.compile(r"^(Public |Private )?(Sub|Function) ")
LABEL = re.compile(r"^\w+:$")


def number_lines(txt):
    out, inproc, n, cont = [], False, 0, False
    for line in txt.split("\r\n"):
        st = line.strip()
        if PROC.match(st):
            inproc = True
            out.append(line)
            cont = st.endswith(" _")
            continue
        if st.startswith(("End Sub", "End Function")):
            inproc = False
        ok = (inproc and st and not cont and not st.startswith("'") and not LABEL.match(st)
              and not st.startswith(SKIP))
        cont = st.endswith(" _")
        if ok:
            n += 10
            out.append(f"{n} {line.lstrip()}")
        else:
            out.append(line)
    return "\r\n".join(out)
