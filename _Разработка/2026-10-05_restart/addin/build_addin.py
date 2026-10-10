# -*- coding: utf-8 -*-
"""Собрать надстройку «Однолинейка.xlam»: модуль SX_Addin.bas + вкладка на ленте.

py build_addin.py [выход.xlam]   (по умолчанию 3_Макросы_Excel/Однолинейка.xlam)
Нужно: Excel → Центр управления безопасностью → «Доверять доступ к объектной модели проектов VBA».
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))
import patch_106  # noqa: E402

ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
DEFAULT_OUT = os.path.join(ROOT, "3_Макросы_Excel", "Однолинейка.xlam")

RIBBON = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="sxTabOdn" label="Однолинейка">
        <group id="sxGrpDocs" label="Документы">
          <button id="sxSpec" label="Обновить спецификацию" size="large" imageMso="TableInsert" onAction="SX_Ribbon"
                  screentip="Спецификация по щиту из «Спецификация» H1 (пусто — все): аппараты, кабели, гофра, клипсы"/>
          <button id="sxTags" label="Бирки" size="large" imageMso="FileNew" onAction="SX_Ribbon"
                  screentip="Лист «Бирки»: строка на каждый провод"/>
          <button id="sxLabels" label="Наклейки" size="large" imageMso="FilePrint" onAction="SX_Ribbon"
                  screentip="Лист «Наклейки»: лента под аппаратами, клетка = 1 модуль DIN"/>
        </group>
        <group id="sxGrpChanges" label="Изменения">
          <button id="sxFix" label="Зафиксировать выдачу" size="large" imageMso="FileSave" onAction="SX_Ribbon"
                  screentip="Запомнить состояние таблицы (выдача № N)"/>
          <button id="sxChanges" label="Показать изменения" size="large" imageMso="Refresh" onAction="SX_Ribbon"
                  screentip="Что поменялось после выдачи: лист «Изменения» и подсветка"/>
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


from numbering import number_lines  # noqa: E402


def build(out=DEFAULT_OUT):
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    tmpdir = tempfile.mkdtemp()
    bas = os.path.join(tmpdir, "SX_Addin.bas")
    with open(os.path.join(HERE, "SX_Addin.bas"), encoding="utf-8") as f:
        txt = f.read().replace("\r\n", "\n").replace("\n", "\r\n")
    with open(bas, "w", encoding="cp1251", newline="") as f:
        f.write(number_lines(txt))
    raw = os.path.join(tmpdir, "raw.xlam")
    app = patch_106.excel()
    try:
        wb = app.Workbooks.Add()
        wb.VBProject.VBComponents.Import(bas)
        wb.VBProject.Name = "Odnolineyka"
        wb.SaveAs(raw, FileFormat=55)   # xlOpenXMLAddIn
        wb.Close(False)
    finally:
        patch_106.quit_excel(app)
    # лента: customUI14.xml + связь в _rels/.rels
    with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
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
    shutil.rmtree(tmpdir, ignore_errors=True)
    print("надстройка:", out)
    return out


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT)
