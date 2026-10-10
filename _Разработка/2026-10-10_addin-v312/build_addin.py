# -*- coding: utf-8 -*-
"""Собрать надстройку «Однолинейка.xlam» для таблицы v3.12: вкладка на ленте + модули
3_Макросы_Excel/SX_Spec.bas, SX_Groups.bas и SX_Ribbon.bas (наклейки, бирки, справка).

py build_addin.py [выход.xlam]   (по умолчанию 3_Макросы_Excel/Однолинейка.xlam)
Нужно: Excel → Центр управления безопасностью → «Доверять доступ к объектной модели проектов VBA».
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
MACROS = os.path.join(ROOT, "3_Макросы_Excel")
DEFAULT_OUT = os.path.join(MACROS, "Однолинейка.xlam")

RIBBON = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="sxTabOdn" label="Однолинейка">
        <group id="sxGrpSpec" label="Спецификация">
          <button id="sxSpec" label="Обновить спецификацию" size="large" imageMso="TableInsert" onAction="SX_Ribbon"
                  screentip="«Спецификация_авто» → «Спецификация», изменения подсвечиваются против принятой ревизии"/>
          <button id="sxAccept" label="Принять изменения" size="large" imageMso="AcceptInvitation" onAction="SX_Ribbon"
                  screentip="Снабжение получило: снять подсветку, ревизия +1"/>
        </group>
        <group id="sxGrpGroups" label="Типовые группы">
          <button id="sxGroup" label="Вставить группу" size="large" imageMso="GroupInsert" onAction="SX_Ribbon"
                  screentip="Выделите строки на листе «Однолинейка» и вставьте типовую группу"/>
          <button id="sxSaveGroup" label="Сохранить как группу" size="large" imageMso="FileSaveAs" onAction="SX_Ribbon"
                  screentip="Запомнить выделенные строки как свою типовую группу"/>
        </group>
        <group id="sxGrpPrint" label="Печать">
          <button id="sxTags" label="Бирки" size="large" imageMso="FileNew" onAction="SX_Ribbon"
                  screentip="Лист «Бирки» (считается сам)"/>
          <button id="sxLabels" label="Наклейки" size="large" imageMso="FilePrint" onAction="SX_Ribbon"
                  screentip="Лист «Наклейки»: лента под аппаратами, клетка = 1 модуль DIN"/>
        </group>
        <group id="sxGrpHelp" label="Справка">
          <button id="sxHelp" label="Как пользоваться" size="large" imageMso="Help" onAction="SX_Ribbon"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

SKIP = ("Dim ", "Case ", "Select ", "End Select", "End Sub", "End Function", "Else", "ElseIf", "End If",
        "Next", "Loop", "End With", "Private Const", "Public Const", "Option ", "Attribute ")
PROC = re.compile(r"^(Public |Private )?(Sub|Function) ")
LABEL = re.compile(r"^\w+:$")


def number_lines(txt):
    """Номера строк в процедурах — чтобы Erl в сообщении об ошибке показывал строку."""
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


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def quit_excel(app, pid):
    try:
        app.Quit()
    except Exception:
        pass
    time.sleep(1)
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


def build(out=DEFAULT_OUT):
    out = os.path.abspath(out)
    tmpdir = tempfile.mkdtemp()
    ribbon_bas = os.path.join(tmpdir, "SX_Ribbon.bas")
    with open(os.path.join(HERE, "SX_Ribbon.bas"), encoding="utf-8") as f:
        txt = f.read().replace("\r\n", "\n").replace("\n", "\r\n")
    with open(ribbon_bas, "w", encoding="cp1251", newline="") as f:
        f.write(number_lines(txt))
    raw = os.path.join(tmpdir, "raw.xlam")
    app, pid = excel()
    try:
        wb = app.Workbooks.Add()
        for bas in (os.path.join(MACROS, "SX_Spec.bas"), os.path.join(MACROS, "SX_Groups.bas"), ribbon_bas):
            wb.VBProject.VBComponents.Import(bas)
        wb.VBProject.Name = "Odnolineyka"
        # компиляция: Отладка → Compile; при ошибке Excel покажет окно — в скрытом экземпляре не ждём
        wb.SaveAs(raw, FileFormat=55)   # xlOpenXMLAddIn
        wb.Close(False)
    finally:
        quit_excel(app, pid)
    tmp_out = out + ".tmp"
    with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "_rels/.rels":
                t = data.decode("utf-8")
                if "customUI14.xml" not in t:
                    t = t.replace("</Relationships>",
                                  f'<Relationship Id="sxRibbon" Type="{REL_TYPE}" Target="customUI/customUI14.xml"/></Relationships>')
                data = t.encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("customUI/customUI14.xml", RIBBON.encode("utf-8"))
    os.replace(tmp_out, out)
    shutil.rmtree(tmpdir, ignore_errors=True)
    print("надстройка:", out)
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    build(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT)
