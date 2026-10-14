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
          <button id="sxSpec" label="Обновить спецификацию" size="large" imageMso="TableInsert" onAction="SX_OnAction"
                  screentip="«Спецификация_авто» → «Спецификация», изменения подсвечиваются против принятой ревизии"/>
          <button id="sxAccept" label="Принять изменения" size="large" imageMso="AcceptInvitation" onAction="SX_OnAction"
                  screentip="Снабжение получило: снять подсветку, ревизия +1"/>
          <button id="sxCatUpd" label="Каталог: обновить" size="large" imageMso="Refresh" onAction="SX_OnAction"
                  screentip="Перенести новые и изменённые позиции из общего каталога в «Каталог» книги"
                  supertip="Файл 1_Таблицы_Excel\Каталог_общий.xlsx. Ключ: категория + производитель + модель + наименование. Запас (O) и формулы не меняются. Перед записью покажет, что изменится."/>
        </group>
        <group id="sxGrpPrint" label="Печать">
          <button id="sxTags" label="Бирки" size="large" imageMso="FileNew" onAction="SX_OnAction"
                  screentip="Создать лист «Бирки» по «Однолинейке»"/>
          <button id="sxLabels" label="Наклейки" size="large" imageMso="FilePrint" onAction="SX_OnAction"
                  screentip="Лист «Наклейки»: лента под аппаратами, клетка = 1 модуль DIN"/>
        </group>
        <group id="sxGrpLinks" label="Связи щитов">
          <button id="sxLink" label="Связать с вышестоящим" size="large" imageMso="HyperlinkInsert" onAction="SX_OnAction"
                  screentip="Выбрать файл вышестоящего щита и линию-фидер: связи появятся в обеих книгах"/>
          <button id="sxLinksUpd" label="Обновить связи" size="large" imageMso="RefreshAll" onAction="SX_OnAction"
                  screentip="Подтянуть данные из связанных книг и показать, что изменилось"/>
          <button id="sxTree" label="Дерево щитов" size="large" imageMso="SmartArtInsert" onAction="SX_OnAction"
                  screentip="Все щиты папки объекта: кто от кого питается, проверки — лист «Дерево_щитов»"/>
        </group>
        <group id="sxGrpKJ" label="Кабельный журнал">
          <button id="sxKJ" label="КЖ в DWG" size="large" imageMso="TableInsert" onAction="SX_OnAction"
                  screentip="Создать или обновить кабельный журнал в DWG. DWG открывать не нужно"
                  supertip="Лист «Кабельный_журнал» → таблица AutoCAD по образцу КЖ_шаблон.dwg. Листы добавляются сами. Ручные графы (проложено, испытания, дата) в DWG сохраняются. Старый файл — .bak рядом."/>
          <button id="sxKJFile" label="КЖ: другой файл" size="large" imageMso="FileSaveAs" onAction="SX_OnAction"
                  screentip="Выбрать другой DWG для кабельного журнала этой книги (новый или уже готовый)"/>
          <button id="sxKJStamp" label="Штамп КЖ" size="large" imageMso="FileOpen" onAction="SX_OnAction"
                  screentip="Указать файл штампа (DWG). Рамка КЖ ссылается на него внешней ссылкой «Штамп»"/>
        </group>
        <group id="sxGrpPZ" label="Стадия П">
          <button id="sxPZ" label="ПЗ по объекту" size="large" imageMso="FileSaveAs" onAction="SX_OnAction"
                  screentip="Все щиты папки объекта: ПЗ_объект.xlsx (таблицы) и ПЗ_объект.docx (расчётная часть ПЗ)"/>
        </group>
        <group id="sxGrpView" label="Столбцы">
          <button id="sxLvl1" label="Только ввод" size="large" imageMso="ZoomOut" onAction="SX_OnAction"
                  screentip="Видны только столбцы, которые вы заполняете" supertip="Лист «Однолинейка» или «Исходные данные». Данные не меняются — столбцы только скрываются."/>
          <button id="sxLvl2" label="Ввод + расчёт" size="large" imageMso="Zoom100" onAction="SX_OnAction"
                  screentip="Ввод и итоги: ток, автомат, кабель, ΔU, КЗ, проверки"/>
          <button id="sxLvl3" label="Все столбцы" size="large" imageMso="ZoomIn" onAction="SX_OnAction"
                  screentip="Все столбцы, со служебными расчётами"/>
        </group>
        <group id="sxGrpHelp" label="Справка">
          <button id="sxHelp" label="Как пользоваться" size="large" imageMso="Help" onAction="SX_OnAction"/>
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


CRLF = chr(13) + chr(10)


def to_cp1251(txt):
    """Символы вне cp1251 (Δ, →, └…): в строке VBA -> " & ChrW(n) & ", в комментарии -> ?"""
    out = []
    for line in txt.split(CRLF):
        q, comment, buf = False, False, []
        for ch in line:
            if not comment and ch == '"':
                q = not q
            elif not comment and not q and ch == "'":
                comment = True
            try:
                ch.encode("cp1251")
                buf.append(ch)
            except UnicodeEncodeError:
                buf.append(f'" & ChrW({ord(ch)}) & "' if q and not comment else "?")
        out.append("".join(buf))
    return CRLF.join(out)


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
    own = []                                     # свои модули (UTF-8) -> cp1251 + номера строк
    for name in ("SX_Ribbon.bas", "SX_KJ.bas", "SX_Catalog.bas"):
        with open(os.path.join(HERE, name), encoding="utf-8") as f:
            txt = f.read().replace("\r\n", "\n").replace("\n", "\r\n")
        own.append(os.path.join(tmpdir, name))
        with open(own[-1], "w", encoding="cp1251", newline="") as f:
            f.write(to_cp1251(number_lines(txt)))
    raw = os.path.join(tmpdir, "raw.xlam")
    app, pid = excel()
    try:
        wb = app.Workbooks.Add()
        for bas in (os.path.join(MACROS, "SX_Spec.bas"), os.path.join(MACROS, "SX_Groups.bas"), *own):
            wb.VBProject.VBComponents.Import(bas)
        wb.VBProject.Name = "Odnolineyka"
        # компиляция: Отладка → Compile; при ошибке Excel покажет окно — в скрытом экземпляре не ждём
        wb.SaveAs(raw, 55)   # xlOpenXMLAddIn (позиционно: именованные аргументы COM теряет)
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
